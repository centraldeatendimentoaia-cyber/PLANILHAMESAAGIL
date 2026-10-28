/* Sai aê · origem do visitante (UTM) e links de WhatsApp.
   Guarda utm_source, utm_medium e utm_campaign da URL (primeiro e último acesso, por 90 dias)
   para saber qual post ou anúncio trouxe cada cliente. Usado em todas as páginas do site. */
(function () {
  var WHATS = '5561999531848';
  var MSG = 'Quero testar o Sai aê';
  var KEY = 'sx_utm';
  var DIAS = 90;
  var CAMPOS = ['utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term'];

  var ler = function () { try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { return null; } };
  var gravar = function (v) { try { localStorage.setItem(KEY, JSON.stringify(v)); } catch (e) {} };

  var q = new URLSearchParams(location.search);
  var agora = {};
  CAMPOS.forEach(function (k) { var v = q.get(k); if (v) agora[k] = v.slice(0, 100); });
  if (q.get('fbclid')) { agora.fbclid = '1'; if (!agora.utm_source) agora.utm_source = 'facebook'; if (!agora.utm_medium) agora.utm_medium = 'paid'; }
  if (q.get('gclid')) { agora.gclid = '1'; if (!agora.utm_source) agora.utm_source = 'google'; if (!agora.utm_medium) agora.utm_medium = 'cpc'; }

  var salvo = ler();
  if (salvo && salvo.t && Date.now() - salvo.t > DIAS * 864e5) salvo = null;
  if (Object.keys(agora).length) {
    agora.pagina = location.pathname;
    salvo = { t: Date.now(), primeiro: (salvo && salvo.primeiro) || agora, ultimo: agora };
    gravar(salvo);
  }

  // Origem atual: o último acesso com UTM; sem UTM, o site.
  function origem() {
    var u = (salvo && salvo.ultimo) || {};
    return { source: u.utm_source || '', medium: u.utm_medium || '', campaign: u.utm_campaign || '', content: u.utm_content || '' };
  }

  // Tipo de origem: 'anuncio', 'instagram' ou 'site'.
  function tipo() {
    var o = origem(), m = o.medium.toLowerCase(), s = o.source.toLowerCase();
    if (/^(cpc|ppc|paid|paid_social|paidsocial|ads?|anuncio)$/.test(m) || (salvo && salvo.ultimo && (salvo.ultimo.fbclid || salvo.ultimo.gclid))) return 'anuncio';
    if (/instagram|^ig$/.test(s)) return 'instagram';
    return 'site';
  }

  // Link do WhatsApp com a mensagem pronta, diferente por origem.
  function waUrl(texto) {
    var t = texto || MSG;
    var k = tipo();
    if (!texto) t += k === 'anuncio' ? ' (vim pelo anúncio)' : k === 'instagram' ? ' (vim pelo Instagram)' : '';
    return 'https://wa.me/' + WHATS + '?text=' + encodeURIComponent(t);
  }

  // Acrescenta as UTMs num link (ex.: cadastro do app), para o app gravar a origem no cadastro.
  function comUtm(url) {
    var u = (salvo && salvo.ultimo) || null;
    if (!u || !url || url.charAt(0) === '#') return url;
    try {
      var x = new URL(url, location.href);
      CAMPOS.forEach(function (k) { if (u[k] && !x.searchParams.has(k)) x.searchParams.set(k, u[k]); });
      return x.toString();
    } catch (e) { return url; }
  }

  // Evento nos dois medidores (só funcionam depois do aceite de cookies).
  function evento(ga, meta, dados) {
    try { if (window.gtag && ga) gtag('event', ga, dados || {}); } catch (e) {}
    try { if (window.fbq && meta) fbq(meta.custom ? 'trackCustom' : 'track', meta.name, dados || {}); } catch (e) {}
  }

  window.SX = { waUrl: waUrl, comUtm: comUtm, origem: origem, tipo: tipo, evento: evento, WHATS: WHATS, MSG: MSG };

  /* ==========================================================
     EDITE AQUI · números de prova, depoimentos e vídeo
     ========================================================== */
  // Números reais de uso. Atualize os valores e as datas (AAAA-MM-DD); a página mostra o período sozinha.
  var NUMEROS = { pedidos: '3.600+', valor: 'R$ 133 mil', inicio: '2026-08-12', fim: '2026-10-23' };

  // Depoimentos em vídeo (3 a 5). A seção só aparece quando houver pelo menos 1.
  // video: link do YouTube ou arquivo .mp4 (ex.: 'img/depoimento-maria.mp4'). capa: imagem opcional.
  // Exemplo: { nome: 'Maria', banca: 'Pastel da Maria', bairro: 'Feira de Taguatinga', frase: 'A fila andou.', video: 'https://youtu.be/XXXX', capa: '' }
  var DEPOIMENTOS = [];

  // Vídeo de demonstração (30 a 45 s): pedido, cozinha e senha na TV. Vazio = espaço escondido.
  var VIDEO_DEMO = { video: '', capa: '' };
  /* ========================================================== */

  var esc = function (t) { return String(t == null ? '' : t).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var dataBR = function (iso) { var p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; };
  function player(v, capa, titulo) {
    var yt = String(v).match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?v=|shorts\/|embed\/))([\w-]{6,})/);
    if (yt) return '<iframe src="https://www.youtube-nocookie.com/embed/' + yt[1] + '" title="' + esc(titulo) + '" loading="lazy" allow="encrypted-media; picture-in-picture" allowfullscreen></iframe>';
    return '<video src="' + esc(v) + '"' + (capa ? ' poster="' + esc(capa) + '"' : '') + ' controls playsinline preload="metadata" aria-label="' + esc(titulo) + '"></video>';
  }
  function conteudo() {
    var s = document.createElement('style');
    s.textContent = '.sx-vid{position:relative;width:100%;aspect-ratio:9/16;border-radius:20px;overflow:hidden;background:#18171C}' +
      '.sx-vid iframe,.sx-vid video{position:absolute;inset:0;width:100%;height:100%;border:0;object-fit:cover}' +
      '.sx-vid--h{aspect-ratio:16/9}' +
      '.sx-deps{display:grid;gap:16px;grid-template-columns:repeat(auto-fill,minmax(220px,1fr))}' +
      '.sx-dep{display:flex;flex-direction:column;gap:10px;margin:0}' +
      '.sx-dep blockquote{margin:0;font-family:Outfit,system-ui,sans-serif;font-weight:700;font-size:18px;line-height:1.3}' +
      '.sx-dep figcaption{font-size:14px;color:#5E5C66}.sx-dep figcaption b{color:#18171C}';
    document.head.appendChild(s);

    // números
    var semanas = Math.max(1, Math.round((Date.parse(NUMEROS.fim) - Date.parse(NUMEROS.inicio)) / 6048e5));
    var put = function (sel, txt) { var els = document.querySelectorAll(sel); for (var i = 0; i < els.length; i++) els[i].textContent = txt; };
    put('[data-num="pedidos"]', NUMEROS.pedidos);
    put('[data-num="valor"]', NUMEROS.valor);
    put('[data-num="semanas"]', semanas + ' semanas');
    put('[data-num="periodo"]', 'Dados reais de uso, de ' + dataBR(NUMEROS.inicio).slice(0, 5) + ' a ' + dataBR(NUMEROS.fim) + '.');

    // depoimentos
    var boxes = document.querySelectorAll('[data-depoimentos]');
    for (var i = 0; i < boxes.length; i++) {
      if (!DEPOIMENTOS.length) continue;
      var lista = boxes[i].querySelector('[data-dep-lista]');
      lista.className = 'sx-deps';
      lista.innerHTML = DEPOIMENTOS.map(function (d) {
        return '<figure class="sx-dep"><div class="sx-vid">' + player(d.video, d.capa, 'Depoimento de ' + d.nome) + '</div>' +
          (d.frase ? '<blockquote>"' + esc(d.frase) + '"</blockquote>' : '') +
          '<figcaption><b>' + esc(d.nome) + '</b>' + (d.banca ? ', ' + esc(d.banca) : '') + (d.bairro ? ' · ' + esc(d.bairro) : '') + '</figcaption></figure>';
      }).join('');
      boxes[i].hidden = false;
    }

    // vídeo de demonstração
    var slots = document.querySelectorAll('[data-video-demo]');
    for (var j = 0; j < slots.length; j++) {
      if (!VIDEO_DEMO.video) continue;
      slots[j].querySelector('[data-vid]').innerHTML = '<div class="sx-vid sx-vid--h">' + player(VIDEO_DEMO.video, VIDEO_DEMO.capa, 'Demonstração do Sai aê') + '</div>';
      slots[j].hidden = false;
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', conteudo); else conteudo();

  // Todo link marcado com data-wa abre o WhatsApp com a mensagem da origem e conta como Lead.
  function pintar() {
    var links = document.querySelectorAll('a[data-wa]');
    for (var i = 0; i < links.length; i++) {
      var a = links[i];
      a.href = waUrl(a.getAttribute('data-wa') || '');
      a.target = '_blank'; a.rel = 'noopener';
    }
  }
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[href*="wa.me/"]');
    if (a) evento('generate_lead', { name: 'Lead' }, { metodo: 'whatsapp', origem: tipo() });
  }, true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', pintar); else pintar();
})();
