"""Gera a pasta site/ pronta para subir na Vercel a partir dos arquivos da raiz.
Uso: python3 tools/build_site.py"""
import pathlib, re, shutil

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / 'site'
OUT.mkdir(exist_ok=True)

# Endereço público do site. Troque para 'https://saiae.com.br' quando o domínio estiver ativo.
SITE_URL = 'https://saiae.com.br'
TITLE = 'Sai aê · Sistema de pedidos para feira, food truck e lanchonete'
DESC = ('Teste 7 dias grátis. O pedido sai do caixa direto pra cozinha e a senha aparece na TV. '
        'Feito pra feira, food truck e lanchonete, por preço de barraca.')
HEAD = f'''<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>{TITLE}</title>
<link rel="canonical" href="{SITE_URL}/">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="description" content="{DESC}">
<meta name="theme-color" content="#FFC21A">
<meta property="og:type" content="website">
<meta property="og:locale" content="pt_BR">
<meta property="og:title" content="Sai aê · Sem papel, sem grito, sem pedido esquecido">
<meta property="og:description" content="{DESC}">
<meta property="og:url" content="{SITE_URL}/">
<meta property="og:site_name" content="Sai aê">
<meta property="og:image" content="{SITE_URL}/og-image.png">
<meta property="og:image:secure_url" content="{SITE_URL}/og-image.png">
<meta property="og:image:type" content="image/png">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta property="og:image:alt" content="Sai aê: sem papel, sem grito, sem pedido esquecido. Teste 7 dias grátis.">
<meta name="twitter:card" content="summary_large_image">
<link rel="icon" type="image/svg+xml" href="/favicon.svg">
<link rel="icon" type="image/png" sizes="512x512" href="/icon-512.png">
<link rel="apple-touch-icon" href="/apple-touch-icon.png">
'''

# ---------- index.html (página de vendas) ----------
lp = (ROOT / 'saiae-lp.html').read_text()
cut = lp.index('<div id="saiae-lp">')
head_part, body_part = lp[:cut], lp[cut:]
head_part = re.sub(r'<title>.*?</title>\n?', '', head_part, count=1)
body_part = body_part.replace("trialUrl: 'raio-x.html',", "trialUrl: '/raio-x',")
(OUT / 'index.html').write_text(HEAD + head_part.strip() + '\n</head>\n<body>\n' + body_part.strip() + '\n</body>\n</html>\n')

# ---------- raio-x.html (teste) ----------
rx = (ROOT / 'raio-x.html').read_text()
rx = rx.replace('<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">',
                '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'
                '<meta name="description" content="Teste grátis de 2 minutos: descubra quanto o seu negócio deixa na mesa e baixe a planilha de controle de vendas.">\n'
                '<meta name="theme-color" content="#FFC21A">\n<link rel="icon" type="image/svg+xml" href="/favicon.svg">\n'
                '<link rel="apple-touch-icon" href="/apple-touch-icon.png">\n'
                '<meta property="og:type" content="website">\n<meta property="og:locale" content="pt_BR">\n'
                '<meta property="og:title" content="Raio-X grátis do seu negócio · Sai aê">\n'
                '<meta property="og:description" content="Em 2 minutos, descubra quanto você deixa na mesa todo mês e ganhe a planilha de controle de vendas.">\n'
                f'<meta property="og:url" content="{SITE_URL}/raio-x">\n'
                f'<meta property="og:image" content="{SITE_URL}/og-image.png">\n'
                '<meta property="og:image:width" content="1200">\n<meta property="og:image:height" content="630">\n'
                '<meta name="twitter:card" content="summary_large_image">', 1)
rx = rx.replace('<div class="brand"><svg', '<a class="brand" href="/" aria-label="Sai aê, voltar ao site"><svg', 1)
rx = rx.replace('<use href="#logo-cheio"/></svg></div>', '<use href="#logo-cheio"/></svg></a>', 1)
rx = rx.replace('var PLANILHA_URL = "planilha-controle-vendas-sai-ae.xlsx";', 'var PLANILHA_URL = "/planilha-controle-vendas-sai-ae.xlsx";')
(OUT / 'raio-x.html').write_text(rx)

# ---------- imagens de compartilhamento e ícones ----------
shutil.copy(ROOT / 'assets' / 'og-image.png', OUT / 'og-image.png')
shutil.copy(ROOT / 'assets' / 'icon-180.png', OUT / 'apple-touch-icon.png')
shutil.copy(ROOT / 'assets' / 'icon-512.png', OUT / 'icon-512.png')

# ---------- planilha limpa ----------
shutil.copy(ROOT / 'planilha-controle-vendas-sai-ae.xlsx', OUT / 'planilha-controle-vendas-sai-ae.xlsx')

# ---------- favicon (ícone do app, tirado dos símbolos da logo) ----------
balao = re.search(r'<path id="I-balao" d="([^"]+)"', lp).group(1)
chapeu = re.search(r'<path id="I-chapeu" d="([^"]+)"', lp).group(1)
(OUT / 'favicon.svg').write_text(
    f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 196 196"><rect width="196" height="196" rx="40" fill="#FFC21A"/>'
    f'<path d="{balao}" fill="#18171C"/><path d="{chapeu}" fill="#FFC21A"/></svg>\n')

# ---------- vercel.json ----------
(OUT / 'vercel.json').write_text('''{
  "cleanUrls": true,
  "trailingSlash": false,
  "headers": [
    {
      "source": "/(.*)\\\\.xlsx",
      "headers": [
        { "key": "Content-Disposition", "value": "attachment; filename=\\"Controle de Vendas - Sai ae.xlsx\\"" },
        { "key": "Content-Type", "value": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
      ]
    }
  ]
}
''')
# vercel.json da raiz: faz a Vercel publicar a pasta site/ mesmo sem configurar Root Directory
import json
(ROOT / 'vercel.json').write_text(json.dumps({'framework': None, 'outputDirectory': 'site', **json.loads((OUT / 'vercel.json').read_text())}, ensure_ascii=False, indent=2) + '\n')
print('site/ gerado:', sorted(p.name for p in OUT.iterdir()))
