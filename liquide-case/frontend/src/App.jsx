// =============================================================================
// Liquide Case — Frontend (React + Vite)
//
// Duas telas:
//   1. LISTAGEM  — EAN, título, categoria e preço médio (máx. 20 por página)
//   2. DETALHE   — todas as lojas/ofertas + botão "Aprimorar com IA"
// Mais: busca no backend (filtra toda a base), exportação Excel e layout
// responsivo.
//
// O frontend NÃO fala com a API externa diretamente — tudo passa pelo
// nosso backend FastAPI (localhost:8000), como pede o case.
// =============================================================================

import { useEffect, useState } from "react";
import "./App.css";

const API = "http://localhost:8000";

// Formata número como moeda BR (ex.: 2411.01 -> "R$ 2.411,01")
const brl = (v) =>
  v == null ? "—" : v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

function App() {
  // ---- estado da listagem -------------------------------------------------
  const [products, setProducts] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(0); // 20 itens por página -> offset = page * 20
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  // ---- estado do detalhe (null = tela de listagem) ------------------------
  const [detail, setDetail] = useState(null);

  // ---- estado da IA --------------------------------------------------------
  const [ai, setAi] = useState(null); // { loading | titulo/descricao | reason }

  const PER_PAGE = 20; // regra do case: no máximo 20 produtos ao mesmo tempo

  // Busca a listagem quando página OU busca mudam (debounce de 300ms
  // para não disparar uma requisição por tecla digitada).
  useEffect(() => {
    if (detail) return; // não recarrega a listagem enquanto o detalhe está aberto
    const timer = setTimeout(() => {
      setLoading(true);
      setError("");
      // A BUSCA vai para o BACKEND: filtra TODA a base, não só esta página.
      // O total devolvido já vem filtrado, então a paginação se adapta.
      fetch(
        `${API}/api/products?limit=${PER_PAGE}&offset=${page * PER_PAGE}` +
          `&search=${encodeURIComponent(search)}`
      )
        .then((r) => {
          if (!r.ok) throw new Error("Backend indisponível");
          return r.json();
        })
        .then((data) => {
          setProducts(data.items);
          setTotal(data.total);
        })
        .catch(() =>
          setError(
            "Não foi possível carregar os produtos. Verifique se o backend está rodando em localhost:8000."
          )
        )
        .finally(() => setLoading(false));
    }, 300);
    return () => clearTimeout(timer); // cancela se o usuário digitar de novo
  }, [page, detail, search]);

  // Nova busca: reseta para a primeira página NO MOMENTO do input.
  // (Antes isso era um useEffect separado, o que causava dupla requisição:
  // uma com a página antiga e outra com a nova. Assim dispara só uma.)
  const onSearchChange = (e) => {
    setSearch(e.target.value);
    setPage(0);
  };

  // Abre o detalhe de um produto
  const openDetail = (id) => {
    setAi(null);
    setLoading(true);
    fetch(`${API}/api/products/${id}`)
      .then((r) => r.json())
      .then((data) => setDetail(data))
      .catch(() => setError("Erro ao carregar o produto."))
      .finally(() => setLoading(false));
  };

  // "Aprimorar com IA": chama o backend; se falhar, só a área da IA é afetada
  // (o restante da tela continua 100% funcional — requisito de resiliência)
  const askAi = () => {
    setAi({ loading: true });
    fetch(`${API}/api/products/${detail.id}/ai-suggest`)
      .then((r) => r.json())
      .then((data) =>
        data.ai_available
          ? setAi({ titulo: data.titulo, descricao: data.descricao })
          : setAi({ reason: data.reason })
      )
      .catch(() => setAi({ reason: "IA indisponível no momento" }));
  };

  // Exportação: baixa o Excel com TODOS os produtos (rota do backend)
  const exportXlsx = () => {
    window.open(`${API}/api/export.xlsx`, "_blank");
  };

  const totalPages = Math.ceil(total / PER_PAGE);

  // ==========================================================================
  // TELA 2 — DETALHE DO PRODUTO
  // ==========================================================================
  if (detail) {
    return (
      <div className="container">
        <button className="back" onClick={() => setDetail(null)}>
          ← Voltar para a listagem
        </button>

        <header className="detail-header">
          <h1>{detail.title}</h1>
          <p className="meta">
            EAN {detail.ean} · {detail.category}
          </p>
          <p className="desc">{detail.description}</p>
          <div className="stats">
            <span>
              Preço médio: <strong>{brl(detail.avg_price)}</strong>
            </span>
            <span>
              Mediana: <strong>{brl(detail.median)}</strong>
            </span>
            <span>
              Ofertas válidas: <strong>{detail.valid_offers_count}</strong>
            </span>
          </div>
        </header>

        <h2>Ofertas por loja</h2>
        <div className="offers">
          {/* guarda contra ofertas undefined (defesa contra resposta inesperada) */}
          {(detail.offers || []).map((o, i) => (
            <div className={`offer ${o.in_average ? "" : "flagged"}`} key={i}>
              <div className="offer-store">{o.store}</div>
              <div className="offer-price">{o.price_display}</div>
              <div className="offer-stock">Estoque: {o.stock_display}</div>
              {/* selos de dados inconsistentes — transparência para o usuário */}
              {o.price_flag && <span className="badge warn">{o.price_flag}</span>}
              {o.stock_flag && <span className="badge">{o.stock_flag}</span>}
              {!o.in_average && !o.price_flag && (
                <span className="badge">fora da média</span>
              )}
            </div>
          ))}
        </div>

        {/* ---- IA: área SEPARADA das informações originais (regra do case) -- */}
        <section className="ai-box">
          <h2>Aprimorar com IA</h2>
          <p className="ai-hint">
            Sugestão gerada por modelo de linguagem — os dados originais acima
            não são alterados.
          </p>
          <button onClick={askAi} disabled={ai?.loading}>
            {ai?.loading ? "Gerando..." : "✨ Aprimorar com IA"}
          </button>

          {ai?.titulo && (
            <div className="ai-result">
              <h3>Título sugerido</h3>
              <p>{ai.titulo}</p>
              <h3>Descrição sugerida</h3>
              <p>{ai.descricao}</p>
            </div>
          )}
          {ai?.reason && <p className="ai-error">⚠️ {ai.reason}</p>}
        </section>
      </div>
    );
  }

  // ==========================================================================
  // TELA 1 — LISTAGEM
  // ==========================================================================
  return (
    <div className="container">
      <header className="list-header">
        <h1>Consulta de Produtos e Ofertas</h1>
        <div className="toolbar">
          <input
            type="search"
            placeholder="Buscar por título, EAN ou categoria..."
            value={search}
            onChange={onSearchChange}
          />
          {/* ↓ tipográfico: renderiza em qualquer fonte (emoji ↓⬇ dependem do sistema) */}
          <button onClick={exportXlsx}>↓ Exportar Excel (todos)</button>
        </div>
      </header>

      {error && <p className="error">{error}</p>}
      {loading && <p className="loading">Carregando...</p>}

      {!loading && !error && (
        <>
          <div className="grid">
            {/* sem filtro client-side: a busca já veio filtrada do backend */}
            {products.map((p) => (
              <button className="card" key={p.id} onClick={() => openDetail(p.id)}>
                <span className="card-cat">{p.category}</span>
                <h3>{p.title}</h3>
                <p className="card-ean">EAN {p.ean}</p>
                <p className="card-price">
                  {p.avg_price == null ? "Sem preço válido" : brl(p.avg_price)}
                  <small> preço médio</small>
                </p>
              </button>
            ))}
          </div>

          {products.length === 0 && (
            <p className="loading">Nenhum produto encontrado para "{search}".</p>
          )}

          {/* paginação: setas para 1ª/última e anterior/próxima */}
            <div className="pager-btns">
              <button disabled={page === 0} onClick={() => setPage(0)}>« Primeira</button>
              <button disabled={page === 0} onClick={() => setPage(page - 1)}>‹ Anterior</button>
              <span>
                Página {page + 1} de {totalPages || 1} · {total} produto{total === 1 ? "" : "s"}
              </span>
              <button disabled={page + 1 >= totalPages} onClick={() => setPage(page + 1)}>Próxima ›</button>
              <button
                disabled={page + 1 >= totalPages}
                onClick={() => setPage(totalPages - 1)}
              >Última »</button>
          </div>
        </>
      )}
    </div>
  );
}

export default App;