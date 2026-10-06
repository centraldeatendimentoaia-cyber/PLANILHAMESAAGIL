"""Gera a página de vendas do Sai aê CRM (crm/sai-ae-crm.html) a partir de tools/crm-src.html."""
import base64, html, io, json, re
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parent.parent
BASE = 'https://saiae.com.br'          # onde ficam fotos, favicon e a imagem do link (og-crm.png)
WHATSAPP = '5561999531848'

MSGS = {
    'header':       'Oi! Vim pelo site do Sai aê CRM e quero saber mais.',
    'hero':         'Oi! Quero conhecer o Sai aê CRM pra minha lanchonete.',
    'essencial':    'Oi! Quero assinar o Sai aê CRM no plano Essencial (R$ 49/mês).',
    'pro':          'Oi! Quero assinar o Sai aê CRM no plano Pro (R$ 99/mês).',
    'escala':       'Oi! Quero assinar o Sai aê CRM no plano Escala (R$ 199/mês).',
    'duvida_plano': 'Oi! Estou em dúvida sobre qual plano do Sai aê CRM escolher.',
    'combo':        'Oi! Quero usar o Sai aê no caixa junto com o CRM.',
    'faq':          'Oi! Tenho uma dúvida sobre o Sai aê CRM.',
    'final':        'Oi! Quero começar a usar o Sai aê CRM.',
    'rodape':       'Oi! Vim pelo site do Sai aê CRM.',
}

FAQ = [
    ('Preciso trocar de número?',
     'Não. Você conecta o WhatsApp que já usa lendo um QR code, igual ao WhatsApp Web.'),
    ('Funciona no celular?',
     'Funciona no navegador do celular, do tablet ou do computador. Não precisa instalar nada.'),
    ('O que é "disparo"?',
     'É cada mensagem enviada pela campanha. Mandou uma promoção pra 200 clientes? São 200 disparos do seu limite do mês.'),
    ('Como a IA ajuda?',
     'Você diz o que quer divulgar e a IA escreve o texto da mensagem. Aí é só ajustar e enviar.'),
    ('Meu número pode ser bloqueado?',
     'O WhatsApp bloqueia quem manda mensagem pra quem não pediu. Use o CRM com quem já é seu cliente e aceitou receber suas novidades, e evite listas compradas. Assim você cuida do seu número e do seu cliente.'),
    ('Posso mudar de plano depois?',
     'Pode. Começou no Essencial e o movimento cresceu? É só chamar a gente no WhatsApp.'),
    ('Como eu assino?',
     'Clique no botão do plano que você quer. Ele abre uma conversa no nosso WhatsApp e a gente faz tudo com você por lá.'),
]

def embutir(m):
    """Troca {{IMG:arquivo}} por data URI webp (a página funciona em qualquer domínio)."""
    from PIL import Image
    nome = m.group(1)
    pasta = 'app' if nome.startswith('app-') else 'fotos'
    im = Image.open(ROOT / 'assets' / pasta / nome).convert('RGB')
    largura = 420 if pasta == 'app' else 720
    if im.width > largura:
        im = im.resize((largura, round(im.height * largura / im.width)), Image.LANCZOS)
    buf = io.BytesIO(); im.save(buf, 'WEBP', quality=72, method=6)
    return 'data:image/webp;base64,' + base64.b64encode(buf.getvalue()).decode()

def wa(key):
    return f'https://wa.me/{WHATSAPP}?text={quote(MSGS[key])}'

src = (ROOT / 'tools' / 'crm-src.html').read_text(encoding='utf-8')
lp = (ROOT / 'saiae-lp.html').read_text(encoding='utf-8')
sprite = re.search(r'<svg class="sprite".*?</defs>\s*</svg>', lp, re.S).group(0)

faq_html = '\n'.join(
    f'      <details class="rv"><summary>{html.escape(q)}<i class="ms">add</i></summary><p>{html.escape(a)}</p></details>'
    for q, a in FAQ)
faq_ld = json.dumps({'@context': 'https://schema.org', '@type': 'FAQPage', 'mainEntity': [
    {'@type': 'Question', 'name': q, 'acceptedAnswer': {'@type': 'Answer', 'text': a}} for q, a in FAQ]},
    ensure_ascii=False)

out = src.replace('{{SPRITE}}', sprite).replace('{{FAQ_HTML}}', faq_html).replace('{{FAQ_JSONLD}}', faq_ld)
out = re.sub(r'\{\{WA:(\w+)\}\}', lambda m: html.escape(wa(m.group(1))), out)
out = re.sub(r'\{\{IMG:([\w.-]+)\}\}', embutir, out)
out = out.replace('{{BASE}}', BASE)
assert '{{' not in out, re.findall(r'\{\{.*?\}\}', out)
dest = ROOT / 'crm' / 'sai-ae-crm.html'
dest.parent.mkdir(exist_ok=True)
dest.write_text(out, encoding='utf-8')
print(dest, len(out))
