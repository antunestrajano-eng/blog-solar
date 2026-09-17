// ============================================================
//  CONFIANCE ENERGY — Blog Solar
//  Google Apps Script — Publicador para GitHub Pages
//  v2.1 — SEO: links rastreáveis, páginas de categoria,
//         datas ISO 8601, links para o site principal sem www
// ============================================================
//
//  SEGURANÇA: o token do GitHub NÃO fica neste arquivo.
//  No editor do Apps Script: Configurações do projeto →
//  Propriedades do script → adicionar GITHUB_TOKEN = ghp_...
// ============================================================

const CONFIG = {
  GITHUB_USER:   'antunestrajano-eng',
  GITHUB_REPO:   'blog-solar',
  GITHUB_BRANCH: 'main',
  JSON_PATH:     'dados.json',
  SITE_URL:      'https://blog.confianceenergy.com',
  SITE_PRINCIPAL: 'https://confianceenergy.com',
  OG_IMAGE_PADRAO: 'https://raw.githubusercontent.com/antunestrajano-eng/blog-solar/main/og-image.jpg',
  LOGO_URL:      'https://i.imgur.com/jbGhPDi.jpg',
  WHATSAPP:      'https://wa.me/5584991465777',
  GA_ID:         'G-VBCJ34XQ3Q',
};

const CATEGORIAS = {
  noticias:     { label: 'Notícia',     plural: 'Notícias',     cor: '#1A3F8F', bg: '#EEF3FC',
                  titulo: 'Notícias sobre energia solar',
                  descricao: 'Reajustes de tarifa, decisões da ANEEL, bandeiras tarifárias e tudo que impacta a conta de luz de quem tem (ou quer ter) energia solar em Natal/RN e no Nordeste.' },
  informativos: { label: 'Informativo', plural: 'Informativos', cor: '#A06000', bg: '#FFF6E6',
                  titulo: 'Informativos sobre energia solar',
                  descricao: 'Guias e explicações sobre inversores, string box, taxa mínima, Lei 14.300, vistoria da Cosern e outros temas para entender sua proposta de energia solar.' },
  dicas:        { label: 'Dica',        plural: 'Dicas',        cor: '#1A6B3A', bg: '#E8F7EE',
                  titulo: 'Dicas de energia solar',
                  descricao: 'Dicas práticas de manutenção, economia, financiamento e cuidados com seu sistema de energia solar.' },
};

function getToken() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('GITHUB_TOKEN não configurado em Propriedades do script.');
  return token;
}

// ============================================================
//  PUBLICAR — lê a planilha, envia dados.json e gera HTMLs
// ============================================================

function publicarBlog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Posts');
  if (!sheet) { Logger.log('❌ Aba "Posts" não encontrada.'); return; }

  // 1. Faz upload das imagens novas para o GitHub
  processarImagens(sheet);

  // 2. Lê os posts (já com URLs do GitHub nas imagens)
  const dados = lerPosts(sheet);
  const json  = JSON.stringify(dados, null, 2);

  // 3. Envia dados.json
  const okJson = enviarArquivo(CONFIG.JSON_PATH, json);
  Logger.log(okJson ? '✅ dados.json enviado' : '❌ Erro ao enviar dados.json');

  // 4. Gera e envia sitemap.xml
  const okSitemap = gerarESendSitemap(dados.posts);
  Logger.log(okSitemap ? '✅ sitemap.xml enviado' : '❌ Erro ao enviar sitemap.xml');

  // 5. Páginas de categoria (/p/noticias/, /p/informativos/, /p/dicas/)
  //    Regeradas a cada publicação: são os hubs com links <a> para
  //    todos os posts — é por aqui que o Google descobre e distribui
  //    autoridade entre os artigos.
  Object.keys(CATEGORIAS).forEach(function(cat) {
    const ok = enviarArquivo('p/' + cat + '/index.html', gerarHtmlCategoria(cat, dados.posts));
    Logger.log(ok ? '✅ Categoria: p/' + cat + '/' : '❌ Erro na categoria: ' + cat);
    Utilities.sleep(300);
  });

  // 6. Gera páginas HTML estáticas para posts novos (coluna I)
  const linhas = sheet.getDataRange().getValues();
  let criados = 0, ignorados = 0;

  for (let i = 1; i < linhas.length; i++) {
    const row = linhas[i];
    if (!row[0] || row[6] !== 'Publicado') continue;

    const previewJaCriado = String(row[8] || '').trim() === 'Sim';
    if (previewJaCriado) { ignorados++; continue; }

    const post = dados.posts.find(p => p.id === i);
    if (!post) continue;

    const html = gerarHtmlCompleto(post, dados.posts);
    const path = `p/${post.categoria}/${post.slug}.html`;
    const ok   = enviarArquivo(path, html);

    if (ok) {
      sheet.getRange(i + 1, 9).setValue('Sim');
      criados++;
      Logger.log('✅ Página criada: ' + path);
    } else {
      Logger.log('❌ Erro ao criar página: ' + path);
    }

    Utilities.sleep(300);
  }

  Logger.log(`✅ Concluído — páginas criadas: ${criados}, já existentes: ${ignorados}`);
}

// ============================================================
//  CSS COMPARTILHADO (post e categoria)
// ============================================================

function cssBase() {
  return `
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --blue: #1A3F8F; --blue-light: #2454B8; --blue-faint: #EEF3FC;
  --gold: #F5A800; --gold-light: #FFD166;
  --text: #0F1F45; --text-muted: #5A6480;
  --border: #E2E8F4; --white: #ffffff; --bg: #F7F9FD;
  --radius: 14px; --shadow: 0 2px 16px rgba(26,63,143,0.08);
}
body { font-family: 'Sora', sans-serif; background: var(--bg); color: var(--text); }

/* HEADER */
header { background: var(--white); border-bottom: 1px solid var(--border); position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 12px rgba(26,63,143,0.07); }
.header-inner { max-width: 1200px; margin: 0 auto; padding: 0 32px; height: 72px; display: flex; align-items: center; justify-content: space-between; }
.logo { display: flex; align-items: center; gap: 12px; text-decoration: none; }
.logo img { height: 44px; width: auto; object-fit: contain; }
.logo-label { font-size: 12px; font-weight: 600; color: var(--text-muted); letter-spacing: 1px; text-transform: uppercase; }
.header-cta { background: var(--blue); color: var(--white); padding: 10px 22px; border-radius: 8px; font-size: 13px; font-weight: 600; text-decoration: none; }
.header-cta:hover { background: var(--blue-light); }
.header-back { font-size: 12px; font-weight: 600; color: var(--text-muted); text-decoration: none; border: 1px solid var(--border); padding: 7px 14px; border-radius: 7px; }
.header-back:hover { color: var(--blue); border-color: var(--blue); }

/* BREADCRUMB */
.breadcrumb { font-size: 12px; color: var(--text-muted); margin-bottom: 24px; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.breadcrumb a { color: var(--blue); text-decoration: none; }
.breadcrumb a:hover { text-decoration: underline; }

/* FOOTER */
footer { background: var(--blue); color: rgba(255,255,255,0.65); text-align: center; padding: 28px 32px; font-size: 13px; }
footer strong { color: var(--gold-light); }
footer a { color: var(--gold-light); text-decoration: none; }
footer a:hover { text-decoration: underline; }
`;
}

function htmlHead(opts) {
  return `<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${escapeHtmlAttr(opts.title)}</title>
<meta name="description" content="${escapeHtmlAttr(opts.description)}">
<meta name="robots" content="index, follow">
<link rel="canonical" href="${opts.url}">

<!-- Open Graph / WhatsApp / Instagram -->
<meta property="og:type"        content="${opts.ogType || 'website'}">
<meta property="og:site_name"   content="Confiance Energy — Blog Solar">
<meta property="og:title"       content="${escapeHtmlAttr(opts.ogTitle || opts.title)}">
<meta property="og:description" content="${escapeHtmlAttr(opts.description)}">
<meta property="og:image"       content="${escapeHtmlAttr(opts.image)}">
<meta property="og:image:secure_url" content="${escapeHtmlAttr(opts.image)}">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta property="og:url"         content="${opts.url}">
<meta property="og:locale"      content="pt_BR">

<!-- Twitter Card -->
<meta name="twitter:card"        content="summary_large_image">
<meta name="twitter:title"       content="${escapeHtmlAttr(opts.ogTitle || opts.title)}">
<meta name="twitter:description" content="${escapeHtmlAttr(opts.description)}">
<meta name="twitter:image"       content="${escapeHtmlAttr(opts.image)}">

${opts.schema ? '<!-- Schema.org -->\n<script type="application/ld+json">' + opts.schema + '</script>\n' : ''}
<!-- Google Analytics -->
<script async src="https://www.googletagmanager.com/gtag/js?id=${CONFIG.GA_ID}"></script>
<script>window.dataLayer=window.dataLayer||[];function gtag(){dataLayer.push(arguments);}gtag('js',new Date());gtag('config','${CONFIG.GA_ID}');</script>

<link rel="icon" type="image/png" href="/favicon.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;500;600;700&family=Lora:ital,wght@0,400;0,600;1,400&display=swap" rel="stylesheet">`;
}

function htmlHeader(voltarUrl) {
  const base = CONFIG.SITE_URL;
  return `<header>
  <div class="header-inner">
    <a class="logo" href="${base}/">
      <img src="${CONFIG.LOGO_URL}" alt="Confiance Energy">
      <span class="logo-label">Blog Solar</span>
    </a>
    <div style="display:flex;align-items:center;gap:10px;">
      <a href="${voltarUrl}" class="header-back">← Voltar ao blog</a>
      <a href="${CONFIG.WHATSAPP}" target="_blank" rel="noopener" class="header-cta">Solicitar Orçamento</a>
    </div>
  </div>
</header>`;
}

// Rodapé com link para o site principal (âncora com palavra-chave e
// URL canônica sem www — a versão www redireciona 301).
function htmlFooter() {
  return `<footer>
  <strong>Confiance Energy</strong> — Blog Solar &nbsp;·&nbsp; Natal, RN &nbsp;·&nbsp; © ${new Date().getFullYear()}<br>
  <a href="${CONFIG.SITE_PRINCIPAL}/">Energia solar em Natal/RN — Confiance Energy</a>
</footer>`;
}

// ============================================================
//  GERA HTML COMPLETO E INDEXÁVEL POR POST
// ============================================================

function gerarHtmlCompleto(post, todosPosts) {
  const base       = CONFIG.SITE_URL;
  const urlPost    = `${base}/p/${post.categoria}/${post.slug}.html`;
  const urlCat     = `${base}/p/${post.categoria}/`;
  const imagem     = post.imagem || CONFIG.OG_IMAGE_PADRAO;
  const cat        = CATEGORIAS[post.categoria] || CATEGORIAS.noticias;
  const dataIsoPost = dataIso(post.data);

  const palavras   = post.conteudo.replace(/<[^>]+>/g, '').split(/\s+/).length;
  const minutos    = Math.max(1, Math.round(palavras / 200));

  const relacionados = todosPosts
    .filter(p => p.categoria === post.categoria && p.slug !== post.slug)
    .slice(0, 3);

  const relacionadosHtml = relacionados.length > 0 ? `
    <section class="related">
      <h2 class="related-title">Leia também</h2>
      <div class="related-grid">
        ${relacionados.map(r => `
        <a href="${base}/p/${r.categoria}/${r.slug}.html" class="related-card">
          ${r.imagem ? `<img src="${r.imagem}" alt="${escapeHtmlAttr(r.titulo)}" loading="lazy">` : `<div class="related-img-placeholder"></div>`}
          <div class="related-body">
            <span class="related-cat">${cat.label}</span>
            <p class="related-title-text">${escapeHtmlAttr(r.titulo)}</p>
            <span class="related-date">📅 ${r.data}</span>
          </div>
        </a>`).join('')}
      </div>
      <p class="related-more"><a href="${urlCat}">Ver todas as ${cat.plural.toLowerCase()} →</a></p>
    </section>` : '';

  const schema = JSON.stringify({
    "@context": "https://schema.org",
    "@type": "Article",
    "headline": post.titulo,
    "description": post.resumo,
    "image": imagem,
    "author": { "@type": "Person", "name": post.autor },
    "publisher": {
      "@type": "Organization",
      "name": "Confiance Energy",
      "url": CONFIG.SITE_PRINCIPAL,
      "logo": { "@type": "ImageObject", "url": CONFIG.LOGO_URL }
    },
    "datePublished": dataIsoPost,
    "dateModified": dataIsoPost,
    "mainEntityOfPage": { "@type": "WebPage", "@id": urlPost }
  });

  const whatsappUrl = `https://wa.me/?text=${encodeURIComponent(post.titulo + ' ' + urlPost)}`;

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
${htmlHead({
  title: post.titulo + ' — Confiance Energy Blog Solar',
  ogTitle: post.titulo,
  description: post.resumo,
  url: urlPost,
  image: imagem,
  ogType: 'article',
  schema: schema,
})}

<style>${cssBase()}
/* ARTIGO */
.article-wrap { max-width: 820px; margin: 0 auto; padding: 48px 32px 80px; }
.article-cat { display: inline-block; font-size: 11px; font-weight: 700; letter-spacing: 1.4px; text-transform: uppercase; padding: 5px 12px; border-radius: 4px; margin-bottom: 16px; background: ${cat.bg}; color: ${cat.cor}; }
.article-title { font-family: 'Lora', serif; font-size: 38px; font-weight: 600; line-height: 1.25; color: var(--text); margin-bottom: 20px; }
.article-meta { font-size: 13px; color: var(--text-muted); display: flex; align-items: center; gap: 20px; margin-bottom: 28px; flex-wrap: wrap; padding-bottom: 20px; border-bottom: 1px solid var(--border); }
.article-img { width: 100%; max-height: 480px; object-fit: cover; border-radius: var(--radius); margin-bottom: 40px; border: 1px solid var(--border); }
.article-content { font-family: 'Lora', serif; font-size: 18px; line-height: 1.85; color: #2A3550; text-align: justify; margin-bottom: 40px; }
.article-content p { margin-bottom: 24px; }
.article-content h2 { font-family: 'Sora', sans-serif; font-size: 20px; font-weight: 600; color: var(--blue); margin: 40px 0 12px; padding-bottom: 8px; border-bottom: 2px solid var(--gold); display: inline-block; }
.article-content h3 { font-family: 'Sora', sans-serif; font-size: 17px; font-weight: 600; color: var(--text); margin: 28px 0 10px; }

/* COMPARTILHAR */
.share-section { border-top: 1px solid var(--border); border-bottom: 1px solid var(--border); padding: 24px 0; margin-bottom: 40px; }
.share-label { font-size: 12px; font-weight: 600; color: var(--text-muted); letter-spacing: 0.5px; text-transform: uppercase; margin-bottom: 14px; }
.share-btns { display: flex; gap: 10px; flex-wrap: wrap; }
.share-btn { display: inline-flex; align-items: center; gap: 8px; padding: 10px 18px; border-radius: 8px; font-size: 13px; font-weight: 600; text-decoration: none; font-family: 'Sora', sans-serif; transition: opacity 0.2s; }
.share-btn:hover { opacity: 0.85; }
.share-btn.whatsapp { background: #25D366; color: #fff; }
.share-btn.copy { background: var(--blue-faint); color: var(--blue); border: none; cursor: pointer; }

/* RELACIONADOS */
.related { margin-top: 48px; }
.related-title { font-family: 'Sora', sans-serif; font-size: 18px; font-weight: 700; color: var(--text); margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid var(--gold); display: inline-block; }
.related-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-top: 20px; }
.related-card { background: var(--white); border-radius: 10px; border: 1px solid var(--border); overflow: hidden; text-decoration: none; color: var(--text); box-shadow: var(--shadow); transition: transform 0.2s, box-shadow 0.2s; }
.related-card:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(26,63,143,0.14); }
.related-card img { width: 100%; height: 130px; object-fit: cover; display: block; }
.related-img-placeholder { width: 100%; height: 130px; background: linear-gradient(135deg, var(--blue-faint), #dce6f8); }
.related-body { padding: 14px; }
.related-cat { font-size: 10px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: var(--blue); }
.related-title-text { font-family: 'Lora', serif; font-size: 14px; font-weight: 600; line-height: 1.4; margin: 6px 0; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
.related-date { font-size: 11px; color: var(--text-muted); }
.related-more { margin-top: 20px; font-size: 14px; font-weight: 600; }
.related-more a { color: var(--blue); text-decoration: none; }
.related-more a:hover { text-decoration: underline; }

/* RESPONSIVO */
@media (max-width: 900px) { .related-grid { grid-template-columns: repeat(2, 1fr); } }
@media (max-width: 600px) {
  .header-inner { padding: 0 16px; }
  .article-wrap { padding: 28px 16px 60px; }
  .article-title { font-size: 26px; }
  .article-content { font-size: 16px; }
  .related-grid { grid-template-columns: 1fr; }
  .header-back { display: none; }
}
</style>
</head>
<body>

${htmlHeader(urlCat)}

<!-- ARTIGO -->
<article class="article-wrap" itemscope itemtype="https://schema.org/Article">

  <nav class="breadcrumb" aria-label="Navegação">
    <a href="${base}/">Blog</a>
    <span>›</span>
    <a href="${urlCat}">${cat.plural}</a>
    <span>›</span>
    <span>${escapeHtmlAttr(post.titulo)}</span>
  </nav>

  <span class="article-cat">${cat.label}</span>
  <h1 class="article-title" itemprop="headline">${escapeHtmlAttr(post.titulo)}</h1>

  <div class="article-meta">
    <span>📅 <time itemprop="datePublished" datetime="${dataIsoPost}">${post.data}</time></span>
    <span>✍️ <span itemprop="author">${escapeHtmlAttr(post.autor)}</span></span>
    <span>⏱️ ${minutos} min de leitura</span>
  </div>

  ${post.imagem ? `<img src="${escapeHtmlAttr(post.imagem)}" alt="${escapeHtmlAttr(post.titulo)}" class="article-img" itemprop="image">` : ''}

  <div class="article-content" itemprop="articleBody">
    ${post.conteudo}
  </div>

  <!-- Compartilhar -->
  <div class="share-section">
    <p class="share-label">Compartilhar</p>
    <div class="share-btns">
      <a href="${whatsappUrl}" target="_blank" rel="noopener" class="share-btn whatsapp">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
        WhatsApp
      </a>
      <button class="share-btn copy" onclick="navigator.clipboard.writeText('${urlPost}').then(()=>{this.textContent='✅ Copiado!'})">
        🔗 Copiar link
      </button>
    </div>
  </div>

  ${relacionadosHtml}

</article>

${htmlFooter()}

</body>
</html>`;
}

// ============================================================
//  PÁGINA DE CATEGORIA — hub com links <a> para todos os posts
// ============================================================

function gerarHtmlCategoria(categoria, todosPosts) {
  const base  = CONFIG.SITE_URL;
  const cat   = CATEGORIAS[categoria];
  const url   = `${base}/p/${categoria}/`;
  const posts = todosPosts.filter(p => p.categoria === categoria);

  const cardsHtml = posts.map(p => `
      <a href="${base}/p/${p.categoria}/${p.slug}.html" class="card">
        ${p.imagem ? `<img src="${escapeHtmlAttr(p.imagem)}" alt="${escapeHtmlAttr(p.titulo)}" class="card-img" loading="lazy" width="600" height="200">` : `<div class="card-img-placeholder"></div>`}
        <div class="card-body">
          <span class="card-category">${cat.label}</span>
          <h2 class="card-title">${escapeHtmlAttr(p.titulo)}</h2>
          <p class="card-excerpt">${escapeHtmlAttr(p.resumo)}</p>
          <div class="card-footer">
            <time class="card-date" datetime="${dataIso(p.data)}">📅 ${p.data}</time>
            <span class="card-read">Ler mais →</span>
          </div>
        </div>
      </a>`).join('');

  const outrasCats = Object.keys(CATEGORIAS)
    .filter(c => c !== categoria)
    .map(c => `<a href="${base}/p/${c}/">${CATEGORIAS[c].plural}</a>`)
    .join(' · ');

  const schema = JSON.stringify({
    "@context": "https://schema.org",
    "@type": "CollectionPage",
    "name": cat.titulo + ' — Blog Solar Confiance Energy',
    "description": cat.descricao,
    "url": url,
    "isPartOf": { "@type": "WebSite", "name": "Confiance Energy — Blog Solar", "url": base + '/' },
    "publisher": { "@type": "Organization", "name": "Confiance Energy", "url": CONFIG.SITE_PRINCIPAL },
    "hasPart": posts.map(p => ({ "@type": "Article", "headline": p.titulo, "url": `${base}/p/${p.categoria}/${p.slug}.html`, "datePublished": dataIso(p.data) }))
  });

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
${htmlHead({
  title: cat.titulo + ' — Blog Solar Confiance Energy',
  description: cat.descricao,
  url: url,
  image: CONFIG.OG_IMAGE_PADRAO,
  schema: schema,
})}

<style>${cssBase()}
main { max-width: 1200px; margin: 0 auto; padding: 48px 32px 80px; }
.cat-header { max-width: 820px; margin-bottom: 36px; }
.cat-title { font-family: 'Lora', serif; font-size: 34px; font-weight: 600; line-height: 1.25; color: var(--text); margin-bottom: 12px; }
.cat-title span { color: var(--gold); }
.cat-desc { font-size: 15px; color: var(--text-muted); line-height: 1.7; }
.cat-meta { font-size: 13px; color: var(--text-muted); margin-top: 12px; }
.cat-meta a { color: var(--blue); text-decoration: none; }
.cat-meta a:hover { text-decoration: underline; }
.cards-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 28px; }
.card { display: block; text-decoration: none; color: inherit; background: var(--white); border-radius: var(--radius); overflow: hidden; border: 1px solid var(--border); box-shadow: var(--shadow); transition: transform 0.25s, box-shadow 0.25s; }
.card:hover { transform: translateY(-4px); box-shadow: 0 8px 32px rgba(26,63,143,0.16); }
.card-img { width: 100%; height: 200px; object-fit: cover; display: block; }
.card-img-placeholder { width: 100%; height: 200px; background: linear-gradient(135deg, var(--blue-faint), #dce6f8); }
.card-body { padding: 22px 22px 20px; }
.card-category { display: inline-block; font-size: 10px; font-weight: 700; letter-spacing: 1.2px; text-transform: uppercase; padding: 4px 10px; border-radius: 4px; margin-bottom: 12px; background: ${cat.bg}; color: ${cat.cor}; }
.card-title { font-family: 'Lora', serif; font-size: 17px; font-weight: 600; line-height: 1.4; color: var(--text); margin-bottom: 10px; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
.card-excerpt { font-size: 13px; color: var(--text-muted); line-height: 1.65; display: -webkit-box; -webkit-line-clamp: 3; -webkit-box-orient: vertical; overflow: hidden; margin-bottom: 18px; }
.card-footer { display: flex; align-items: center; justify-content: space-between; padding-top: 16px; border-top: 1px solid var(--border); }
.card-date { font-size: 12px; color: var(--text-muted); }
.card-read { font-size: 12px; font-weight: 600; color: var(--blue); }
.empty { text-align: center; padding: 80px 20px; color: var(--text-muted); }
@media (max-width: 900px) { .cards-grid { grid-template-columns: repeat(2, 1fr); } .cat-title { font-size: 28px; } }
@media (max-width: 600px) {
  .header-inner { padding: 0 16px; }
  main { padding: 24px 16px 60px; }
  .cards-grid { grid-template-columns: 1fr; }
  .cat-title { font-size: 24px; }
  .header-back { display: none; }
}
</style>
</head>
<body>

${htmlHeader(base + '/')}

<main>
  <nav class="breadcrumb" aria-label="Navegação">
    <a href="${base}/">Blog</a>
    <span>›</span>
    <span>${cat.plural}</span>
  </nav>

  <div class="cat-header">
    <h1 class="cat-title">${escapeHtmlAttr(cat.titulo)} <span>— Blog Solar</span></h1>
    <p class="cat-desc">${escapeHtmlAttr(cat.descricao)}</p>
    <p class="cat-meta">${posts.length} publicaç${posts.length === 1 ? 'ão' : 'ões'} &nbsp;·&nbsp; Outras categorias: ${outrasCats}</p>
  </div>

  ${posts.length > 0 ? `<div class="cards-grid">${cardsHtml}
  </div>` : `<p class="empty">Nenhum conteúdo publicado ainda.</p>`}
</main>

${htmlFooter()}

</body>
</html>`;
}

function escapeHtmlAttr(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ============================================================
//  PROCESSA IMAGENS — baixa e envia pro GitHub se necessário
// ============================================================

function processarImagens(sheet) {
  const linhas = sheet.getDataRange().getValues();
  for (let i = 1; i < linhas.length; i++) {
    const row = linhas[i];
    if (!row[0] || row[6] !== 'Publicado') continue;
    const urlOriginal = String(row[4] || '').trim();
    const imagemNoGit = String(row[9] || '').trim() === 'Sim';
    if (!urlOriginal || imagemNoGit) continue;
    if (urlOriginal.includes('raw.githubusercontent.com')) {
      sheet.getRange(i + 1, 10).setValue('Sim');
      continue;
    }
    const titulo  = String(row[0] || '').trim();
    const slug    = gerarSlug(titulo) + '--' + i;
    const ext     = detectarExtensao(urlOriginal);
    const gitPath = `img/${slug}${ext}`;
    const gitUrl  = `https://raw.githubusercontent.com/${CONFIG.GITHUB_USER}/${CONFIG.GITHUB_REPO}/main/${gitPath}`;
    Logger.log('⬆️ Enviando imagem: ' + gitPath);
    try {
      const resp = UrlFetchApp.fetch(urlOriginal, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) { Logger.log('❌ Não foi possível baixar: ' + urlOriginal); continue; }
      const ok = enviarArquivoBinario(gitPath, Utilities.base64Encode(resp.getContent()));
      if (ok) {
        sheet.getRange(i + 1, 5).setValue(gitUrl);
        sheet.getRange(i + 1, 10).setValue('Sim');
        Logger.log('✅ Imagem enviada: ' + gitUrl);
      } else {
        Logger.log('❌ Erro ao enviar imagem: ' + gitPath);
      }
    } catch(e) { Logger.log('❌ Erro ao processar imagem: ' + e.message); }
    Utilities.sleep(500);
  }
}

function detectarExtensao(url) {
  const match = url.match(/\.(jpg|jpeg|png|webp|gif)(\?|$)/i);
  if (match) return '.' + match[1].toLowerCase().replace('jpeg', 'jpg');
  return '.jpg';
}

function enviarArquivoBinario(path, base64) {
  const token = getToken();
  const url   = `https://api.github.com/repos/${CONFIG.GITHUB_USER}/${CONFIG.GITHUB_REPO}/contents/${path}`;
  let sha = null;
  try {
    const get = UrlFetchApp.fetch(url, { method: 'GET', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (get.getResponseCode() === 200) sha = JSON.parse(get.getContentText()).sha;
  } catch(e) {}
  const payload = { message: '🖼️ Imagem: ' + path, content: base64, branch: CONFIG.GITHUB_BRANCH };
  if (sha) payload.sha = sha;
  try {
    const resp = UrlFetchApp.fetch(url, { method: 'PUT', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, payload: JSON.stringify(payload), muteHttpExceptions: true });
    const code = resp.getResponseCode();
    return code === 200 || code === 201;
  } catch(e) { Logger.log('Erro binário em ' + path + ': ' + e.message); return false; }
}

// ============================================================
//  LÊ OS POSTS DA PLANILHA
// ============================================================

function lerPosts(sheet) {
  const linhas = sheet.getDataRange().getValues();
  const posts  = [];
  for (let i = 1; i < linhas.length; i++) {
    const row = linhas[i];
    if (!row[0] || row[6] !== 'Publicado') continue;
    const titulo = String(row[0] || '').trim();
    const id     = i;
    const slug   = gerarSlug(titulo) + '--' + id;
    posts.push({
      id, slug, titulo,
      resumo:    String(row[1] || '').trim(),
      conteudo:  formatarConteudo(String(row[2] || '').trim()),
      categoria: String(row[3] || 'noticias').trim().toLowerCase(),
      imagem:    String(row[4] || '').trim(),
      autor:     String(row[5] || 'Redação Confiance').trim(),
      status:    String(row[6] || '').trim(),
      data:      formatarData(row[7]),
    });
  }
  posts.reverse();
  return { posts, atualizadoEm: new Date().toISOString() };
}

// ============================================================
//  REGENERAR TODAS AS PÁGINAS (use quando atualizar o layout)
//  Zera a coluna I de todos os posts e roda publicarBlog()
// ============================================================

function regenerarTodasAsPaginas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Posts');
  if (!sheet) { Logger.log('❌ Aba "Posts" não encontrada.'); return; }
  const linhas = sheet.getDataRange().getValues();
  for (let i = 1; i < linhas.length; i++) {
    if (!linhas[i][0]) continue;
    sheet.getRange(i + 1, 9).setValue(''); // Zera coluna I (HTML criado)
  }
  Logger.log('🔄 Coluna I zerada. Rodando publicarBlog()...');
  publicarBlog();
}

// ============================================================
//  ENVIA UM ARQUIVO TEXTO PARA O GITHUB VIA API
// ============================================================

function enviarArquivo(path, conteudo) {
  const token = getToken();
  const url   = `https://api.github.com/repos/${CONFIG.GITHUB_USER}/${CONFIG.GITHUB_REPO}/contents/${path}`;
  let sha = null;
  try {
    const get = UrlFetchApp.fetch(url, { method: 'GET', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (get.getResponseCode() === 200) sha = JSON.parse(get.getContentText()).sha;
  } catch(e) {}
  const payload = {
    message: '🌞 Blog atualizado em ' + new Date().toLocaleString('pt-BR'),
    content: Utilities.base64Encode(conteudo, Utilities.Charset.UTF_8),
    branch:  CONFIG.GITHUB_BRANCH,
  };
  if (sha) payload.sha = sha;
  try {
    const resp = UrlFetchApp.fetch(url, { method: 'PUT', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, payload: JSON.stringify(payload), muteHttpExceptions: true });
    const code = resp.getResponseCode();
    return code === 200 || code === 201;
  } catch(e) { Logger.log('Erro em ' + path + ': ' + e.message); return false; }
}

// ============================================================
//  GERA E ENVIA O SITEMAP.XML
// ============================================================

function gerarESendSitemap(posts) {
  const base = CONFIG.SITE_URL;
  const hoje = new Date().toISOString().split('T')[0];
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n\n';
  xml += `  <url>\n    <loc>${base}/</loc>\n    <lastmod>${hoje}</lastmod>\n    <changefreq>daily</changefreq>\n    <priority>1.0</priority>\n  </url>\n\n`;
  Object.keys(CATEGORIAS).forEach(function(cat) {
    const ultimo = posts.find(p => p.categoria === cat);
    xml += `  <url>\n    <loc>${base}/p/${cat}/</loc>\n    <lastmod>${ultimo ? dataIso(ultimo.data) : hoje}</lastmod>\n    <changefreq>weekly</changefreq>\n    <priority>0.8</priority>\n  </url>\n`;
  });
  xml += '\n';
  posts.forEach(function(post) {
    xml += `  <url>\n    <loc>${base}/p/${post.categoria}/${post.slug}.html</loc>\n    <lastmod>${dataIso(post.data)}</lastmod>\n    <changefreq>monthly</changefreq>\n    <priority>0.9</priority>\n  </url>\n`;
  });
  xml += '\n</urlset>';
  return enviarArquivo('sitemap.xml', xml);
}

// ============================================================
//  HELPERS
// ============================================================

const MESES = { 'jan':'01','fev':'02','mar':'03','abr':'04','mai':'05','jun':'06','jul':'07','ago':'08','set':'09','out':'10','nov':'11','dez':'12' };

// "17 de set. de 2026" → "2026-09-17" (ISO 8601, exigido pelo Google
// em datePublished/dateModified e em <time datetime>)
function dataIso(dataStr) {
  try {
    const partes = String(dataStr).replace(/ de /g, ' ').split(' ');
    const dia = partes[0].padStart(2, '0');
    const mes = MESES[partes[1].substring(0, 3).toLowerCase()] || '01';
    const ano = partes[2] || new Date().getFullYear().toString();
    return `${ano}-${mes}-${dia}`;
  } catch(e) { return new Date().toISOString().split('T')[0]; }
}

function gerarSlug(titulo) {
  return titulo.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9\s-]/g, '').trim().replace(/\s+/g, '-').substring(0, 80);
}

function formatarConteudo(texto) {
  if (!texto) return '';
  // O site principal responde em https://confianceenergy.com (a versão
  // www redireciona 301). Normaliza para evitar o salto de redirect.
  texto = texto.replace(/https?:\/\/www\.confianceenergy\.com/gi, CONFIG.SITE_PRINCIPAL);
  if (texto.includes('<p>') || texto.includes('<h2>')) return texto;
  return texto.split(/\n\s*\n/).map(p => '<p>' + p.replace(/\n/g, '<br>').trim() + '</p>').filter(p => p !== '<p></p>').join('');
}

function formatarData(valor) {
  if (!valor) return '';
  try {
    const d = new Date(valor);
    return d.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch(e) { return String(valor); }
}

// ============================================================
//  MENU NA PLANILHA
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('☀️ Blog Confiance')
    .addItem('🚀 Publicar blog agora', 'publicarBlog')
    .addItem('🔄 Regenerar todas as páginas', 'regenerarTodasAsPaginas')
    .addToUi();
}
