/* Aviso de cookies (LGPD): medição só depois do "Aceitar". Negado por padrão em gtag('consent','default'). */
(function () {
  var KEY = 'sx_consent';
  var read = function () { try { return localStorage.getItem(KEY); } catch (e) { return null; } };
  var save = function (v) { try { localStorage.setItem(KEY, v); } catch (e) {} };
  var PIXEL_ID = ''; // Pixel do Meta: cole o ID aqui; vazio = pixel desligado
  var loadPixel = function () {
    if (!PIXEL_ID || window.fbq) return;
    !function (f, b, e, v, n, t, s) { if (f.fbq) return; n = f.fbq = function () { n.callMethod ? n.callMethod.apply(n, arguments) : n.queue.push(arguments); };
      if (!f._fbq) f._fbq = n; n.push = n; n.loaded = !0; n.version = '2.0'; n.queue = []; t = b.createElement(e); t.async = !0;
      t.src = v; s = b.getElementsByTagName(e)[0]; s.parentNode.insertBefore(t, s); }(window, document, 'script', 'https://connect.facebook.net/en_US/fbevents.js');
    fbq('init', PIXEL_ID); fbq('track', 'PageView');
  };
  var grant = function () { if (window.gtag) gtag('consent', 'update', { analytics_storage: 'granted', ad_storage: 'granted', ad_user_data: 'granted', ad_personalization: 'granted' }); loadPixel(); };

  var saved = read();
  if (saved === 'granted') { grant(); return; }
  if (saved === 'denied') return;

  var css = document.createElement('style');
  css.textContent =
    '.sx-ck{position:fixed;left:16px;right:16px;bottom:calc(16px + env(safe-area-inset-bottom,0px));z-index:90;max-width:560px;margin:0 auto;' +
    'display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between;padding:14px 16px;border-radius:16px;' +
    'background:#1c1c1c;color:#fff;font:500 14px/1.4 Figtree,system-ui,sans-serif;box-shadow:0 8px 28px rgba(0,0,0,.28)}' +
    '.sx-ck p{margin:0;flex:1 1 240px}.sx-ck a{color:#FFC21A}' +
    '.sx-ck button{font:700 14px Figtree,system-ui,sans-serif;min-height:44px;padding:0 18px;border-radius:999px;border:0;cursor:pointer}' +
    '.sx-ck .ok{background:#FFC21A;color:#1c1c1c}.sx-ck .no{background:transparent;color:#fff;box-shadow:inset 0 0 0 1.5px #fff}';
  document.head.appendChild(css);

  var box = document.createElement('div');
  box.className = 'sx-ck';
  box.setAttribute('role', 'region');
  box.setAttribute('aria-label', 'Aviso de cookies');
  box.innerHTML =
    '<p>Usamos cookies para medir como o site é usado e melhorar a página. Detalhes na ' +
    '<a href="https://app.saiae.com.br/privacidade" target="_blank" rel="noopener">Política de Privacidade</a>.</p>' +
    '<span style="display:flex;gap:8px"><button type="button" class="no">Recusar</button><button type="button" class="ok">Aceitar</button></span>';
  document.body.appendChild(box);

  box.querySelector('.ok').addEventListener('click', function () { save('granted'); grant(); box.remove(); });
  box.querySelector('.no').addEventListener('click', function () { save('denied'); box.remove(); });
})();
