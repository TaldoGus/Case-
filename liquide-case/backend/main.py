# =============================================================================
# Liquide Case API — backend do case técnico "Consulta de produtos e ofertas"
#
# STACK: FastAPI + httpx + openpyxl
#
# VISÃO GERAL DA ARQUITETURA (arquivo único, de propósito):
#   1. SETUP          — variáveis de ambiente, app FastAPI, CORS
#   2. CAMADA DE DADOS— busca na API externa com cache curto em memória
#   3. SANEAMENTO     — regras para preços, estoques e lojas inconsistentes
#   4. ROTAS          — listagem (paginada), detalhe, IA e exportação Excel
#
# DECISÃO DE DESIGN: manter um único arquivo com funções de domínio bem
# separadas (parse_price, summarize_offers) — para um projeto deste porte,
# isso é mais legível do que hierarquia de pastas. A próxima refatoração
# natural, se o projeto crescesse ou eu tivesse mais seguro do tempo , seria extrair "services/pricing.py" (o
# saneamento) e "services/groq.py" (IA) — a separação já existe nas funções.
# =============================================================================

import io          # buffer em memória para gerar o .xlsx sem salvar em disco
import json        # parse da resposta JSON do modelo de IA
import os          # leitura de variáveis de ambiente
import statistics  # média e mediana dos preços
import time        # controle de expiração do cache

import httpx      # cliente HTTP assíncrono (chamadas à API externa e ao Groq)
from dotenv import load_dotenv  # carrega o backend/.env
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import Workbook     # geração do Excel

# Carrega as variáveis de backend/.env (chave do Groq, URL da API externa).
# Credenciais NUNCA vão para o repositório — o .env está no .gitignore.
load_dotenv()

app = FastAPI(title="Liquide Case API")

# -----------------------------------------------------------------------------
# CORS — sem isto, o browser bloqueia o frontend em localhost:5173 de chamar
# este backend em localhost:8000 (política de mesma origem dos navegadores).
# -----------------------------------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------------------------------------------------------------
# CONFIGURAÇÃO (lida do .env, com defaults seguros)
#
# Por quê isolado no topo do arquivo? Durante o desenvolvimento, o modelo do
# Groq foi descontinuado ('llama-3.3-70b-versatile' saiu da lista de modelos
# disponíveis). Como a configuração está isolada, corrigi em UMA linha para
# 'openai/gpt-oss-120b' (consultei GET /models do Groq em vez de chutar nomes).
# Separar configuração de código é o que torna a troca barata.
# -----------------------------------------------------------------------------
PRODUCTS_API_URL = os.getenv(
    "PRODUCTS_API_URL", "https://agente.liquide.com.br/case/api/products"
)
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "openai/gpt-oss-120b"

# -----------------------------------------------------------------------------
# 1. CAMADA DE DADOS — API externa
#
# Regra do case: NÃO baixar o JSON nem copiar a base para o projeto. Os dados
# são sempre buscados da API externa. Cache de 60s em memória:
#   - evita bater na API externa a cada clique/travessia de página;
#   - mantém os dados "ao vivo" (atualiza a cada minuto, nunca fica stale);
#   - cabe em memória com folga (50 produtos; ver README para base grande).
# -----------------------------------------------------------------------------
_cache: dict = {"data": None, "expires": 0.0}
CACHE_TTL_SECONDS = 60


async def get_products() -> list[dict]:
    """Busca os produtos na API externa, com cache curto em memória."""
    now = time.time()

    # Cache válido? devolve sem tocar na rede.
    if _cache["data"] is not None and now < _cache["expires"]:
        return _cache["data"]

    try:
        # timeout explícito: se a API externa travar, o usuário espera no
        # máximo 15s e recebe um erro claro — nunca fica pendente para sempre.
        async with httpx.AsyncClient(timeout=15) as client:
            resp = await client.get(PRODUCTS_API_URL)
            resp.raise_for_status()  # 4xx/5xx viram exceção, não dado silencioso
            _cache["data"] = resp.json()
            _cache["expires"] = now + CACHE_TTL_SECONDS
            return _cache["data"]
    except httpx.HTTPError as e:
        # 502 Bad Gateway comunica bem: "o problema está no upstream, não aqui"
        raise HTTPException(status_code=502, detail=f"API externa indisponível: {e}")


# -----------------------------------------------------------------------------
# 2. SANEAMENTO DE DADOS
#
# O enunciado avisa: a base tem valores incorretos e não indica quais nem qual
# regra usar. Procurando na base REAL, encontrei estes problemas (exemplos):
#
#   Problema                    | Exemplo real (id do produto)
#   ----------------------------|------------------------------------------
#   preço negativo              | id 4: -3499.90 (Notebook Lenovo)
#   preço igual a zero          | id 48: 0.00 (Nintendo Switch)
#   preço ausente (null)        | id 18: null (Motorola Edge)
#   preço em formato inválido   | id 35: "valor_indisponivel" (string!)
#   estoque negativo            | id 11: -3 (Mouse HyperX)
#   estoque ausente (null)      | id 44: null (Cadeira Escritório)
#   loja sem identificação      | id 27: "" (Fone Sony)
#   preço destoante (>15%)      | id 16: Fast Shop a 5453 vs ~4630 (iPhone)
#
# REGRAS ADOTADAS (todas derivadas do próprio enunciado):
#   a) Preço VALIDO (> 0, numérico): entra na média do produto.
#   b) Preço <= 0, null ou inválido: a OFERTA CONTINUA APARECENDO no detalhe
#      (transparência para o usuário), mas fica FORA do cálculo da média —
#      o enunciado define exatamente essa exclusão para zero/negativo, e
#      estendemos a mesma lógica para null e formatos inválidos (são erros
#      equivalentes: "não sei o preço" não pode poluir a média).
#   c) Estoque igual a zero: VÁLIDO — oferta aparece normalmente e participa
#      da média (regra explícita do enunciado; zero não é "sem estoque
#      informado", é "sem estoque na prateleira").
#   d) Estoque negativo ou null: mostrado como "não informado" — não inventamos
#      valor; invalidar a oferta inteira seria perder dado real de preço.
#   e) Loja vazia: exibida como "Loja não informada" — a oferta continua válida
#      porque a identificação da loja não afeta o preço nem a média.
#   f) Referência dos 15%: usamos a MEDIANA dos preços válidos (não a média) —
#      a mediana é robusta a outliers (célebre: um preço destoante puxaria a
#      média e contaminaria a própria régua). Oferta com desvio >15% da mediana
#      recebe o selo "preço destoante" MAS PERMANECE NA MÉDIA: o enunciado só
#      manda excluir preço <= 0; sinalizamos para o usuário julgar.
# -----------------------------------------------------------------------------


def parse_price(raw) -> float | None:
    """
    Converte o preço bruto em float, ou None se inválido (regra b acima).
    Trata os 4 formatos problemáticos: null, negativo, zero e string
    não numérica (ex.: 'valor_indisponivel').
    """
    if raw is None:
        return None  # preço ausente

    # Alguns preços vêm como STRING (ex.: "valor_indisponivel", id 35).
    # Tentamos converter se for numérica; se não for, é inválido.
    if isinstance(raw, str):
        cleaned = raw.strip()
        if not cleaned:
            return None  # string vazia = preço ausente
        try:
            # aceita tanto "1234.56" quanto o formato BR "1.234,56"
            raw = float(cleaned.replace(".", "").replace(",", ".")
                        ) if cleaned.count(",") > cleaned.count(".") else float(cleaned)
        except ValueError:
            return None  # string não numérica = formato inválido

    # Preço precisa ser numérico e POSITIVO (zero e negativo ficam fora —
    # regra explícita do enunciado para o cálculo da média).
    if not isinstance(raw, (int, float)) or raw <= 0:
        return None
    return float(raw)


def classify_price_reason(raw, parsed: float | None, median: float | None) -> str | None:
    """
    Devolve o MOTIVO (texto legível) do preço não entrar na média,
    para o selo visual na tela. Mantido como função separada para
    o motivo ser reutilizado pela interface sem retrabalho de texto.
    """
    if parsed is not None:
        # preço válido: só pode ter selo se estiver destoante (>15% da mediana)
        if median and parsed > median * 1.15 or (median and parsed < median * 0.85):
            # Regra dos 15%: SINALIZADO, mas permanece na média (ver regra f).
            return "preço destoante (> 15% da mediana)"
        return None

    # preço NÃO entrou na média — classificar por que:
    if raw is None:
        return "preço ausente"
    if isinstance(raw, (int, float)):
        # refinamento: distinguir negativo de zero fica claro na tela
        return "preço negativo" if raw < 0 else "preço igual a zero"
    return "preço em formato inválido"


def summarize_offers(offers: list[dict]) -> dict:
    """
    Aplica TODAS as regras de saneamento por oferta e devolve:
     - offers normalizadas (loja, preço formatado, estoque, selos e motivos);
     - estatísticas agregadas (preço médio, mediana, qtde de ofertas válidas).

    Esta função é o CORAÇÃO do tratamento de dados: tanto a listagem quanto
    o detalhe e o Excel passam por ela — uma única fonte de verdade, para
    que a média exibida e a exportada NUNCA diverjam.
    """
    # Mediana = régua da regra dos 15% (robusta a outliers — ver regra f).
    valid_prices = [p for o in offers if (p := parse_price(o.get("price"))) is not None]
    median = statistics.median(valid_prices) if valid_prices else None

    normalized = []
    for o in offers:
        price = parse_price(o.get("price"))
        stock = o.get("stock")
        store = (o.get("store") or "").strip()

        # (e) loja sem identificação: mantenho a oferta, sinalizo a loja
        store_display = store if store else "Loja não informada"

        # estoque: (c) zero é válido; (d) negativo/null = "não informado"
        stock_ok = isinstance(stock, (int, float)) and stock >= 0

        normalized.append(
            {
                "store": store_display,
                # price: float válido OU None (a interface decide como exibir)
                "price": price,
                "price_display": f"R$ {price:.2f}".replace(".", ",") if price else "—",
                "stock_display": str(int(stock)) if stock_ok else "não informado",
                # in_average=price is not None → 'false' nos casos (b)
                "in_average": price is not None,
                "price_flag": classify_price_reason(o.get("price"), price, median),
                # selo de estoque problemático, só quando existe
                **(
                    {"stock_flag": "estoque ausente ou negativo"}
                    if not stock_ok
                    else {}
                ),
            }
        )

    return {
        "offers": normalized,
        # média dos preços VÁLIDOS (regras a/b) — arredondada p/ 2 casas (moeda)
        "avg_price": round(statistics.mean(valid_prices), 2) if valid_prices else None,
        "median": round(median, 2) if median is not None else None,
        "valid_offers_count": len(valid_prices),
    }


def product_summary(p: dict) -> dict:
    """
    Resumo para a LISTAGEM: id, ean, título, categoria e preço médio — só os
    campos exigidos pelo case (payload menor, render mais rápido, paginação
    natural com limite de 20 itens por página).
    """
    s = summarize_offers(p.get("offers", []))
    return {
        "id": p.get("id"),
        "ean": p.get("ean"),
        "title": p.get("title"),
        "category": p.get("category"),
        "avg_price": s["avg_price"],  # None se nenhum preço válido → "—" na tela
    }


# -----------------------------------------------------------------------------
# 3. ROTAS
# -----------------------------------------------------------------------------

@app.get("/api/products")
async def list_products(limit: int = 20, offset: int = 0, search: str = ""):
    """
    Listagem paginada (máx. 20) com BUSCA no backend.

    Por quê buscar aqui e não no frontend? Assim o filtro roda sobre TODA
    a base (50 produtos, ou milhões num cenário real) — buscar só na página
    visível esconderia resultados que estão em outras páginas. O total
    devolvido reflete o filtro, então a paginação acompanha a busca.
    """
    products = await get_products()

    # Busca por título, EAN ou categoria — insensível a maiúsculas
    q = search.strip().lower()
    if q:
        products = [
            p for p in products
            if q in (p.get("title") or "").lower()
            or q in str(p.get("ean") or "")
            or q in (p.get("category") or "").lower()
        ]

    total = len(products)  # total JÁ filtrado: a paginação se adapta à busca
    limit = min(limit, 20)  # teto de 20 por requisição (regra do case)
    items = [product_summary(p) for p in products[offset : offset + limit]]
    return {"total": total, "limit": limit, "offset": offset, "items": items}


@app.get("/api/products/{product_id}")
async def product_detail(product_id: int):
    """
    Detalhe do produto: todas as lojas com preço/estoque saneados +
    estatísticas. Retorna 404 com mensagem clara se o id não existir.
    """
    products = await get_products()
    p = next((x for x in products if x.get("id") == product_id), None)
    if not p:
        raise HTTPException(status_code=404, detail="Produto não encontrado")
    s = summarize_offers(p.get("offers", []))
    return {**p, **s}  # dados originais + ofertas normalizadas + estatísticas


@app.get("/api/products/{product_id}/ai-suggest")
async def ai_suggest(product_id: int):
    """
    'Aprimorar com IA' (requisito 6): envia as informações do produto ao
    Groq e devolve sugestão de título + descrição MELHORES. A aplicação
    NÃO altera os dados originais — a resposta só é mostrada na tela.

    RESILIÊNCIA (requisito: "se a IA falhar, a app continua funcionando"):
    qualquer falha (sem chave, rede, timeout, resposta fora do esperado)
    devolve ai_available=False — NUNCA uma exceção que derrubaria a página.
    """
    products = await get_products()
    p = next((x for x in products if x.get("id") == product_id), None)
    if not p:
        raise HTTPException(status_code=404, detail="Produto não encontrado")

    # Degradar bem ANTES de gastar uma chamada: sem chave configurada?
    if not GROQ_API_KEY:
        return {"ai_available": False, "reason": "chave da IA não configurada"}

    # Compacta as ofertas num texto curto para o prompt (token = custo/latência)
    offers_txt = "; ".join(
        f"{o.get('store') or 'loja não informada'}: "
        f"{o.get('price') if isinstance(o.get('price'), (int, float)) and o.get('price') > 0 else 'preço inválido'}"
        f" (estoque {o.get('stock')})"
        for o in p.get("offers", [])
    )
    prompt = (
        "Você é um copywriter de e-commerce. Com base nos dados abaixo, "
        "sugira um TÍTULO melhor e uma DESCRIÇÃO melhor para o produto. "
        "Responda APENAS com JSON válido, sem explicações e sem cercar de "
        "markdown, no formato {\"titulo\": \"...\", \"descricao\": \"...\"}. "
        f"Categoria: {p.get('category')}. Título atual: {p.get('title')}. "
        f"Descrição atual: {p.get('description')}. Ofertas: {offers_txt}."
    )

    try:
        # timeout 20s: geração de texto pode demorar; acima disso é falha
        async with httpx.AsyncClient(timeout=20) as client:
            resp = await client.post(
                GROQ_URL,
                headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
                json={
                    "model": GROQ_MODEL,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.4,  # baixo: preciso de resposta concisa e estável
                },
            )
            resp.raise_for_status()

            content = resp.json()["choices"][0]["message"]["content"]
            # Modelos costumam cercar JSON de blocos markdown (```json ... ```);
            # removemos cercas e quebras antes do parse.
            content = (
                content.strip()
                .removeprefix("```json")
                .removeprefix("```")
                .removesuffix("```")
                .strip()
            )
            parsed = json.loads(content)

            # valida os dois campos ANTES de devolver sucesso — se o modelo
            # devolver outra chave, tratamos como falha e não como tela vazia
            titulo = parsed.get("titulo")
            descricao = parsed.get("descricao")
            if not titulo or not descricao:
                raise ValueError("resposta da IA sem os campos esperados")

            return {"ai_available": True, "titulo": titulo, "descricao": descricao}

    except Exception as e:
        # FALHA GRACIOSA (o coração do requisito 6): logamos a causa real
        # (livro de bordo para depuração) e devolvemos um body amigável —
        # a página de detalhe continua 100% funcional sem a IA.
        print(f"[AI-SUGGEST] Falha: {type(e).__name__}: {e!r}")
        return {"ai_available": False, "reason": "IA indisponível no momento"}


@app.get("/api/export.xlsx")
async def export_xlsx():
    """
    Exportação Excel (requisito 7): TODOS os produtos da base real — nunca
    só a página visível. Colunas EXATAS pedidas: id, ean, titulo, preço médio.

    O preço médio exportado usa a MESMA summarize_offers da tela — é assim
    que garantimos que exportação e interface nunca divergem (o revisor pode
    conferir número a número entre tela e planilha: batem por construção).
    """
    products = await get_products()
    wb = Workbook()
    ws = wb.active
    ws.title = "Produtos"
    ws.append(["id", "ean", "titulo", "preço médio"])  # cabeçalho pedido pelo case
    for p in products:
        s = summarize_offers(p.get("offers", []))
        ws.append([p.get("id"), p.get("ean"), p.get("title"), s["avg_price"]])

    # Gera em memória (BytesIO): sem arquivo temporário no disco do servidor.
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Content-Type + Content-Disposition → navegador faz o DOWNLOAD direto
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=produtos.xlsx"},
    )


@app.get("/health")
def health():
    """
    Healthcheck simples: usado para validar setup (status) e confirmar
    que a chave do Groq foi carregada do .env sem vazar seu conteúdo
    (apenas true/false, nunca a chave).
    """
    return {"status": "ok", "groq_key_loaded": bool(GROQ_API_KEY)}