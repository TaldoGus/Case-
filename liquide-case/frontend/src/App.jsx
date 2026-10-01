
import { useEffect, useState } from "react";
import "./App.css";

const API = "http://localhost:8000";
const PER_PAGE = 20;

// Converte preços numéricos ou strings para moeda brasileira.
const brl = (value) => {
  if (value == null || value === "") return "—";

  let number = value;

  if (typeof number === "string") {
    number = number.trim().replace(/[R$\s]/g, "");

    // Trata formatos como 1.234,56 e 1234.56
    if (number.includes(",")) {
      number = number.replace(/\./g, "").replace(",", ".");
    }
  }

  number = Number(number);

  return Number.isFinite(number)
    ? number.toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL",
      })
    : "—";
};

// Retorna o primeiro campo preenchido disponível.
const getField = (obj, fields, fallback = "") => {
  for (const field of fields) {
    const value = obj?.[field];
    if (value !== null && value !== undefined && value !== "") {
      return value;
    }
  }
  return fallback;
};

// Normaliza os campos do produto para exibição.
const normalizeProduct = (product) => ({
  ...product,
  id: getField(product, ["id", "product_id", "_id"]),
  title: getField(
    product,
    ["title", "titulo", "name", "product_name", "nome"],
    "Produto sem título"
  ),
  ean: getField(product, ["ean", "EAN", "gtin", "barcode"], "—"),
  category: getField(
    product,
    ["category", "categoria"],
    "Sem categoria"
  ),
  avg_price: getField(
    product,
    ["avg_price", "average_price", "preco_medio", "preco_médio"],
    null
  ),
});

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);

  if (!response.ok) {
    throw new Error(`Erro HTTP ${response.status}`);
  }

  return response.json();
}

function App() {
  // Listagem
  const [products, setProducts] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(0);
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  // Detalhes
  const [detail, setDetail] = useState(null);
  const [detailLoading, setDetailLoading] = useState(false);

  // IA
  const [ai, setAi] = useState(null);

  // Busca produtos com debounce e cancela requisições antigas.
  useEffect(() => {
    if (detail) return;

    const controller = new AbortController();

    const timer = setTimeout(async () => {
      setLoading(true);
      setError("");

      try {
        const params = new URLSearchParams({
          limit: String(PER_PAGE),
          offset: String(page * PER_PAGE),
          search: search.trim(),
        });

        const data = await fetchJson(
          `${API}/api/products?${params.toString()}`,
          { signal: controller.signal }
        );

        const items = Array.isArray(data.items)
          ? data.items
          : [];

        setProducts(items.map(normalizeProduct));
        setTotal(Number(data.total) || 0);
      } catch (err) {
        if (err.name !== "AbortError") {
          setError(
            "Não foi possível carregar os produtos. Verifique se o backend está ativo."
          );
          setProducts([]);
          setTotal(0);
        }
      } finally {
        if (!controller.signal.aborted) {
          setLoading(false);
        }
      }
    }, 300);

    return () => {
      clearTimeout(timer);
      controller.abort();
    };
  }, [page, detail, search]);

  const onSearchChange = (e) => {
    setSearch(e.target.value);
    setPage(0);
  };

  // Abre o detalhe e trata erros da API.
  const openDetail = async (id) => {
    if (!id) {
      setError("Identificador do produto não encontrado.");
      return;
    }

    setAi(null);
    setError("");
    setDetailLoading(true);

    try {
      const data = await fetchJson(
        `${API}/api/products/${encodeURIComponent(id)}`
      );
      setDetail(data);
    } catch {
      setError("Erro ao carregar os detalhes do produto.");
    } finally {
      setDetailLoading(false);
    }
  };

  const askAi = async () => {
    if (!detail?.id) return;

    setAi({ loading: true });

    try {
      const data = await fetchJson(
        `${API}/api/products/${encodeURIComponent(detail.id)}/ai-suggest`
      );

      if (data.ai_available) {
        setAi({
          titulo: data.titulo,
          descricao: data.descricao,
        });
      } else {
        setAi({
          reason: data.reason || "Sugestão de IA indisponível.",
        });
      }
    } catch {
      setAi({ reason: "IA indisponível no momento." });
    }
  };

  const exportXlsx = () => {
    window.open(`${API}/api/export.xlsx`, "_blank", "noopener,noreferrer");
  };

  const totalPages = Math.ceil(total / PER_PAGE);

  const goBack = () => {
    setDetail(null);
    setAi(null);
    setError("");
  };

  // Tela de detalhes
  if (detail) {
    const product = normalizeProduct(detail);
    const offers = Array.isArray(detail.offers)
      ? detail.offers
      : [];

    return (
      <main className="container">
        <button className="back" onClick={goBack}>
          <span aria-hidden="true">←</span> Voltar para a listagem
        </button>

        <header className="detail-header">
          <span className="eyebrow">DETALHES DO PRODUTO</span>
          <h1>{product.title}</h1>
          <p className="meta">
            EAN: {product.ean} <span>·</span> {product.category}
          </p>
          {detail.description && (
            <p className="desc">{detail.description}</p>
          )}

          <div className="stats">
            <div className="stat">
              <span>Preço médio</span>
              <strong>{brl(getField(detail, [
                "avg_price", "average_price", "preco_medio"
              ], null))}</strong>
            </div>
            <div className="stat">
              <span>Mediana</span>
              <strong>{brl(detail.median)}</strong>
            </div>
            <div className="stat">
              <span>Ofertas válidas</span>
              <strong>{detail.valid_offers_count ?? 0}</strong>
            </div>
          </div>
        </header>

        <section className="section">
          <div className="section-heading">
            <div>
              <span className="eyebrow">COMPARAÇÃO</span>
              <h2>Ofertas por loja</h2>
            </div>
            <span className="offer-count">
              {offers.length} {offers.length === 1 ? "oferta" : "ofertas"}
            </span>
          </div>

          {offers.length > 0 ? (
            <div className="offers">
              {offers.map((offer, i) => (
                <article
                  className={`offer ${offer.in_average === false ? "flagged" : ""}`}
                  key={offer.id ?? `${offer.store}-${i}`}
                >
                  <div className="offer-top">
                    <span className="offer-store">
                      {offer.store || "Loja não identificada"}
                    </span>
                    {offer.in_average === false && (
                      <span className="badge">Fora da média</span>
                    )}
                  </div>
                  <div className="offer-price">
                    {offer.price_display ?? brl(offer.price)}
                  </div>
                  <div className="offer-stock">
                    Estoque: {offer.stock_display ?? offer.stock ?? "—"}
                  </div>
                  {offer.price_flag && (
                    <span className="badge warn">{offer.price_flag}</span>
                  )}
                  {offer.stock_flag && (
                    <span className="badge">{offer.stock_flag}</span>
                  )}
                </article>
              ))}
            </div>
          ) : (
            <p className="empty">Nenhuma oferta disponível para este produto.</p>
          )}
        </section>

        <section className="ai-box">
          <div className="ai-heading">
            <div className="ai-icon">✦</div>
            <div>
              <h2>Aprimorar com IA</h2>
              <p className="ai-hint">
                Gere sugestões de título e descrição sem alterar os dados originais.
              </p>
            </div>
          </div>

          <button
            className="ai-button"
            onClick={askAi}
            disabled={ai?.loading}
          >
            {ai?.loading ? "Gerando sugestões..." : "✦ Aprimorar com IA"}
          </button>

          {ai?.titulo && (
            <div className="ai-result">
              <h3>Título sugerido</h3>
              <p>{ai.titulo}</p>
              <h3>Descrição sugerida</h3>
              <p>{ai.descricao || "Sem descrição sugerida."}</p>
            </div>
          )}

          {ai?.reason && <p className="ai-error">⚠ {ai.reason}</p>}
        </section>
      </main>
    );
  }

  // Tela de listagem
  return (
    <main className="container">
      <header className="list-header">
        <div className="brand">
          <div className="brand-mark">L</div>
          <span>liquide<span className="brand-dot">.</span></span>
        </div>

        <div className="page-intro">
          <span className="eyebrow">CATÁLOGO</span>
          <h1>Consulta de Produtos e Ofertas</h1>
          <p>Encontre produtos e compare preços de diferentes lojas.</p>
        </div>

        <div className="toolbar">
          <div className="search-wrap">
            <span className="search-icon" aria-hidden="true">⌕</span>
            <input
              type="search"
              placeholder="Buscar por título, EAN ou categoria..."
              value={search}
              onChange={onSearchChange}
              aria-label="Buscar produtos"
            />
            {search && (
              <button
                className="clear-search"
                onClick={() => onSearchChange({ target: { value: "" } })}
                aria-label="Limpar busca"
              >
                ×
              </button>
            )}
          </div>
          <button className="export-button" onClick={exportXlsx}>
            <span aria-hidden="true">↓</span> Exportar Excel
          </button>
        </div>
      </header>

      {error && <p className="error" role="alert">{error}</p>}

      {detailLoading && <p className="loading">Carregando detalhes...</p>}

      {loading && <p className="loading">Carregando produtos...</p>}

      {!loading && !error && (
        <>
          <div className="results-heading">
            <span>
              {search ? "Resultados da busca" : "Produtos disponíveis"}
            </span>
            <span className="results-total">
              {total} {total === 1 ? "produto" : "produtos"}
            </span>
          </div>

          {products.length > 0 ? (
            <div className="grid">
              {products.map((p, index) => (
                <button
                  className="card"
                  key={p.id ?? `${p.ean}-${index}`}
                  onClick={() => openDetail(p.id)}
                  disabled={!p.id}
                >
                  <div className="card-top">
                    <span className="card-cat">{p.category}</span>
                    <span className="card-arrow" aria-hidden="true">↗</span>
                  </div>
                  <h3>{p.title}</h3>
                  <p className="card-ean">EAN {p.ean}</p>
                  <div className="card-footer">
                    <div>
                      <span className="price-label">Preço médio</span>
                      <p className="card-price">
                        {p.avg_price == null
                          ? "Sem preço válido"
                          : brl(p.avg_price)}
                      </p>
                    </div>
                    <span className="view-detail">Ver detalhes</span>
                  </div>
                </button>
              ))}
            </div>
          ) : (
            <div className="empty-state">
              <div className="empty-icon">⌕</div>
              <h2>Nenhum produto encontrado</h2>
              <p>
                {search
                  ? `Não encontramos resultados para "${search}".`
                  : "Ainda não existem produtos cadastrados."}
              </p>
              {search && (
                <button className="back" onClick={() => setSearch("")}>
                  Limpar busca
                </button>
              )}
            </div>
          )}

          {totalPages > 1 && (
            <nav className="pager" aria-label="Paginação">
              <button
                disabled={page === 0}
                onClick={() => setPage(0)}
                aria-label="Primeira página"
              >«</button>
              <button
                disabled={page === 0}
                onClick={() => setPage((current) => current - 1)}
              >‹ Anterior</button>
              <span>
                Página <strong>{page + 1}</strong> de {totalPages}
              </span>
              <button
                disabled={page + 1 >= totalPages}
                onClick={() => setPage((current) => current + 1)}
              >Próxima ›</button>
              <button
                disabled={page + 1 >= totalPages}
                onClick={() => setPage(totalPages - 1)}
                aria-label="Última página"
              >»</button>
            </nav>
          )}
        </>
      )}

      <footer className="app-footer">
        <span>Liquide · Consulta de produtos</span>
        <span>Dados fornecidos pelo catálogo</span>
      </footer>
    </main>
  );
}

export default App;