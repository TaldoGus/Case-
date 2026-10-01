# Liquide — Case Técnico: consulta de produtos e ofertas

Aplicação fullstack para consultar produtos e ofertas de lojas a partir da
API externa da Liquide, com saneamento de dados inconsistentes, cálculo de
preço médio, busca/paginação, sugestão de título e descrição via IA
(Groq) e exportação da base completa em Excel.

**Stack:** Python (FastAPI, httpx, openpyxl) · JavaScript (React + Vite)

- **Backend (FastAPI):** API intermediária que consulta a API externa,
  classifica os dados inconsistentes, calcula o preço médio e expõe
  listagem, detalhe, IA e exportação. O frontend só fala com ele.
- **Frontend (React + Vite):** listagem com busca, paginação e exportação
  + tela de detalhe com todas as ofertas e o botão "Aprimorar com IA".

---

## 1. Como executar

### Backend

```bash
cd liquide-case/backend
python -m venv .venv
.venv\Scripts\activate          # Windows (Mac/Linux: source .venv/bin/activate)
pip install fastapi uvicorn httpx openpyxl python-dotenv
```

Crie o arquivo `backend/.env` (**não versionado** — está no `.gitignore`):
```
GROQ_API_KEY=sua_chave_do_groq
PRODUCTS_API_URL=https://agente.liquide.com.br/case/api/products
```

```bash
uvicorn main:app --reload --port 8000
```

- API: `http://localhost:8000` · Docs interativas: `http://localhost:8000/docs`
- Healthcheck (e verificação de que a chave carregou): `http://localhost:8000/health`

### Frontend

```bash
cd liquide-case/frontend
npm install
npm run dev
```

Interface em `http://localhost:5173`.

> A IA é opcional: sem `GROQ_API_KEY` a aplicação sobe normalmente e
> apenas o botão "Aprimorar com IA" reporta que está indisponível.

---

## 2. Organização da solução

```text
liquide-case/
├── backend/
│   ├── main.py     # arquivo único e comentado, em camadas concisas:
│   │               #   1. SETUP      — .env, app FastAPI, CORS
│   │               #   2. DADOS      — API externa + cache em memória (60s)
│   │               #   3. SANEAMENTO — parse_price / summarize_offers
│   │               #   4. ROTAS      — listagem, detalhe, IA, export, health
│   └── .env        # credenciais (fora do repositório)
└── frontend/
    └── src/
        ├── App.jsx # listagem + detalhe + IA + export
        └── App.css # estilos (mobile-first)
```

**Por que um `main.py` único?** Para o porte do case (5 endpoints, ~250
linhas), espalhar em pastas adicionaria indireção sem ganho de legibilidade.
A separação de responsabilidades existe nas funções de domínio
(`parse_price`, `summarize_offers` — funções puras, fáceis de testar); a
refatoração natural se o projeto crescesse seria extrair
`services/pricing.py` e `services/groq.py`, já que as fronteiras estão
definidas.

---

## 3. Suposições e decisões sobre os dados inconsistentes

A API externa entrega valores incorretos sem indicar quais nem a regra de
tratamento. Problemas reais encontrados na base: preço **negativo** (id 4:
-3499.90), preço **zero** (id 48), preço **null** (id 18), preço como
**string não numérica** ("valor_indisponivel", id 35), estoque **negativo**
(id 11) e **null** (id 44), **loja sem nome** (id 27).

**Princípio adotado: os dados não são corrigidos nem apagados — são
classificados.** Toda oferta problemática continua visível no detalhe,
com um selo explicando o problema; o status só decide se ela participa
do preço médio.

| Situação | Decisão | Motivo |
|---|---|---|
| Preço > 0 numérico | Entra na média | Regra normal |
| Preço <= 0, null ou formato inválido | **Exibida, FORA da média** | Enunciado exclui 0/negativo; estendi para null e string inválida ("não sei o preço" é o mesmo erro) |
| Estoque = 0 | **Válido**; oferta entra na média | Regra explícita do enunciado |
| Estoque null/negativo | "não informado"; não invalida a oferta | Não inventar valor; perder o preço real seria pior |
| Loja vazia | "Loja não informada"; oferta entra na média | Problema de identificação, não de preço |
| Outlier (> 15% da mediana) | **Selo "preço destoante"; permanece na média** | Enunciado só manda excluir preço <= 0; o usuário julga |

**A referência dos 15% é a mediana dos preços válidos** — robusta a
outliers: a média seria puxada pelo próprio destoante e contaminaria a
régua. Exemplo real: iPhone com 4611.46 / 4658.06 / 5453.71 → o
5453.71 é sinalizado (+17% da mediana) e os demais definem a média.

A média da listagem, do detalhe e do Excel sai da **mesma função**
(`summarize_offers`) — número a número, tela e planilha nunca divergem.

---

## 4. IA — "Aprimorar com IA" (Groq)

- Botão na tela de detalhe chama `GET /api/products/{id}/ai-suggest`
  e o backend envia o produto ao chat model da Groq, pedindo título
  e descrição melhores em JSON estruturado.
- **Nada é alterado na base:** a sugestão aparece em área separada
  ("Sugestão da IA"), abaixo dos dados originais.
- Configuração isolada no topo do `main.py`: durante o projeto o modelo
  `llama-3.3-70b-versatile` foi decomissionado; a troca para
  `openai/gpt-oss-120b` (confirmada via `GET /models` da Groq) custou
  1 linha — é o benefício de separar configuração de código.
- **Se a IA falhar** (sem chave, rede, timeout de 20 s, resposta fora
  do formato esperado): o backend registra a causa real no log e
  devolve `{"ai_available": false, "reason": ...}` — nunca uma
  exceção que derrubaria a página. O restante da aplicação segue
  100% funcional, e a tela mostra o motivo na área da IA.

---

## 5. Exportação Excel

`GET /api/export.xlsx` gera em memória (openpyxl + BytesIO, sem tocar o
disco) um `.xlsx` com **todos os produtos da base** — nunca só a página
visível — nas colunas exatas do case: `id`, `ean`, `titulo`, `preço médio`.
Produto sem preço válido sai com a célula vazia (coerente com o
"Indisponível" da interface).

---

## 6. O que melhoraria se a base crescesse muito

- A API externa não oferece paginação/filtro: hoje o backend busca a
  lista inteira e cacheia (60s). Com milhões de produtos, o dado migraria
  para um **banco indexado** (busca textual, índice por EAN, agregar no
  SQL) em vez de filtrar listas em memória.
- **Cache compartilhado** (Redis) e invalidação por produto, em vez de
  cache por processo com polling inteiro.
- **Ingestão orientada a evento** (webhook/fila) do distribuidor, para
  reduzir a latência entre mudança de estoque e exibição, e processar o
  saneamento em background.
- **Testes automatizados** das funções puras de saneamento (pytest) —
  exatamente porque são puras, cobrem todos os casos da tabela acima.