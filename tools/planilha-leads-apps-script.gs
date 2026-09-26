/**
 * Sai aê · Raio-X do Negócio → Google Planilhas
 * Cole este código em: Planilha > Extensões > Apps Script (substitua tudo que estiver lá).
 * Depois: Implantar > Nova implantação > Tipo "App da Web"
 *   - Executar como: Eu
 *   - Quem pode acessar: Qualquer pessoa
 * Copie o endereço gerado (termina em /exec) e passe para o Claude colocar no site.
 */
const TOKEN = 'saiae-raiox-2026';   // tem que ser igual ao PLANILHA_LEADS_TOKEN do raio-x.html
const ABA = 'Contatos';
const CABECALHO = ['Data e hora', 'Nome', 'WhatsApp', 'Tipo de negócio', 'Pedidos por dia', 'Ticket médio',
  'Como anota hoje', 'Confusões por semana', 'Perde cliente na fila', 'Potencial/mês (R$)',
  'Faturamento estimado/mês (R$)', 'Origem', 'Campanha'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return resposta(false);

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA) || planilha.insertSheet(ABA);
    const trava = LockService.getScriptLock();
    trava.waitLock(10000);
    if (aba.getLastRow() === 0) {
      aba.appendRow(CABECALHO);
      aba.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold').setBackground('#FFC21A');
      aba.setFrozenRows(1);
    }
    aba.appendRow([
      new Date(), limpo(d.nome), limpo(d.whatsapp), limpo(d.tipo), limpo(d.pedidos), limpo(d.ticket),
      limpo(d.anota), limpo(d.erros), limpo(d.fila), Number(d.potencial) || 0,
      Number(d.faturamento) || 0, limpo(d.origem), limpo(d.campanha)
    ]);
    trava.releaseLock();
    return resposta(true);
  } catch (err) {
    return resposta(false);
  }
}

// Evita que um texto vire fórmula na planilha e limita o tamanho.
function limpo(v) {
  let t = String(v == null ? '' : v).slice(0, 200);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function resposta(ok) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok })).setMimeType(ContentService.MimeType.JSON);
}

// Abrir o endereço /exec no navegador mostra esta mensagem: serve para conferir se a implantação está pública.
function doGet() {
  return ContentService.createTextOutput('Sai aê · planilha de contatos no ar. Envios chegam por POST do site.');
}

// Para testar direto no editor: escolha "testarGravacao" no menu ao lado de "Executar" e clique em Executar.
// Deve aparecer uma linha "TESTE" na aba Contatos. Depois é só apagar a linha.
function testarGravacao() {
  const r = doPost({ postData: { contents: JSON.stringify({
    token: TOKEN, nome: 'TESTE', whatsapp: '(61) 90000-0000', tipo: 'Barraca de feira',
    pedidos: 'até 20 pedidos', ticket: 'até R$ 15', anota: 'no papel', erros: 'quase nunca',
    fila: 'não, nunca', potencial: 0, faturamento: 0, origem: 'teste no editor', campanha: ''
  }) } });
  Logger.log(r.getContent());
}
