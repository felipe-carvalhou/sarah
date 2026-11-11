/**
 * Convida&u — Confirmações de presença do convite
 * 15 anos da Sarah Alves — 09/01/2027, 19h
 *
 * Como usar (uma vez por convite):
 *  1. Cole este arquivo inteiro no Apps Script e salve. Pode ser um projeto avulso (script.google.com)
 *     ou aberto pela planilha (Extensões → Apps Script) — funciona dos dois jeitos.
 *  2. No topo, escolha a função "configurar" e clique em Executar (autorize quando o Google pedir).
 *     Num projeto avulso, ela CRIA a planilha sozinha — o link aparece no "Registro de execução".
 *  3. (opcional) Rode "testarConfirmacao" e veja a linha de teste na planilha — depois apague a linha.
 *  4. Implantar → Nova implantação → App da Web
 *       Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 *  5. Copie o "URL do app da Web" (termina em /exec) e cole em CONFIG.planilha no index.html
 *     E TAMBÉM em CONFIG.planilha no painel-ryex4ikzq4.html.
 *
 * Painel da família: painel-ryex4ikzq4.html (a senha já vai dentro do painel; é a CHAVE_PAINEL abaixo).
 *
 * LIMITE POR LINK: cada link do convite termina em ?c=CÓDIGO. O convite repassa esse código para cá como 'cod'
 * (nunca 'c' ou 'sid': são nomes reservados do Google). A planilha só aceita até o limite do código.
 * Link sem código (ou com código errado) = 1 pessoa.
 *
 * Se você mudar este código depois: Implantar → Gerenciar implantações → lápis → Versão: Nova versão.
 * (Use sempre "Gerenciar implantações": assim o link /exec continua o mesmo e não precisa mexer no convite.)
 * (Sem isso, o link continua rodando o código antigo.)
 */

const ABA = 'Confirmações';
const CHAVE_PAINEL = 'sa15-dg0xhxjce';    // senha do painel da família — use uma diferente para cada cliente
const CABECALHO = ['Data/hora', 'Nome', 'Resposta', 'Pessoas', 'Recado', 'ID', 'Convite'];

// Código do link → máximo de pessoas (o mesmo convite, 5 links)
const CODIGOS = {
  '8m4uz2': 1,
  '6sd2yv': 2,
  'tq69s7': 3,
  'hnw67j': 4,
  'oou3k8': 5
};

/** Rode uma vez: cria a aba de confirmações e a aba de resumo com os totais. */
function configurar() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {                               // projeto avulso: usa (ou cria) a planilha deste convite
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('PLANILHA_ID');
    ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('Confirmações • 15 anos da Sarah Alves');
    props.setProperty('PLANILHA_ID', ss.getId());
  }
  abaConfirmacoes_(ss);

  const r = ss.getSheetByName('Resumo') || ss.insertSheet('Resumo', 0);
  r.clear();
  r.getRange('A1').setValue('Resumo das confirmações').setFontWeight('bold').setFontSize(14);
  r.getRange('A3:A5').setValues([['Pessoas confirmadas (total)'], ['Respostas "vou"'], ['Respostas "não vou"']]);
  // Planilha em português (e outros idiomas) separa as partes da fórmula com ";" em vez de ","
  const sep = /^(pt|es|fr|de|it|nl|ru|tr|id)/i.test(ss.getSpreadsheetLocale()) ? ';' : ',';
  r.getRange('B3').setFormula(`=SUMIF('Confirmações'!C:C${sep}"Vai*"${sep}'Confirmações'!D:D)`);
  r.getRange('B4').setFormula(`=COUNTIF('Confirmações'!C:C${sep}"Vai*")`);
  r.getRange('B5').setFormula(`=COUNTIF('Confirmações'!C:C${sep}"Não vai*")`);
  r.getRange('B3:B5').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  r.setColumnWidth(1, 240);
  r.setColumnWidth(2, 90);
  r.getRange('A7').setValue('A lista completa fica na aba "Confirmações". Se alguém mudar a resposta, a linha dele é atualizada.')
    .setFontColor('#888888');
  Logger.log('Planilha das confirmações: ' + ss.getUrl());
}

/** A planilha do convite: a que abriu o script ou a que o "configurar" criou. */
function planilha_() {
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) return ativa;
  const id = PropertiesService.getScriptProperties().getProperty('PLANILHA_ID');
  if (!id) throw new Error('Rode a função "configurar" uma vez antes de usar.');
  return SpreadsheetApp.openById(id);
}

/** Recebe confirmação por POST (jeito antigo — continua aceitando). */
function doPost(e) {
  return registrar_((e && e.parameter) || {});
}

/** Grava (ou atualiza) uma confirmação na planilha. */
function registrar_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);                    // evita que duas respostas ao mesmo tempo se atropelem
  try {
    const limpa = (t, n) => {              // corta o tamanho e impede fórmulas maliciosas na planilha
      const s = String(t || '').trim().slice(0, n);
      return /^[=+\-@]/.test(s) ? "'" + s : s;
    };
    const nome = limpa(p.nome, 120);
    if (!nome) return resposta_({ ok: false, erro: 'sem nome' });

    const vai = p.vai === 'sim';
    // o código do link chega como 'cod' (o Google reserva 'c' e 'sid' — com eles o pedido nem chega aqui)
    const limite = CODIGOS[String(p.cod || '').toLowerCase()] || 1;   // confere o limite aqui, não só no convite
    const pessoas = vai ? Math.min(limite, Math.max(1, parseInt(p.pessoas, 10) || 1)) : 0;
    const id = limpa(p.id, 40) || Utilities.getUuid();
    const linha = [new Date(), nome, vai ? 'Vai ✅' : 'Não vai ❌', pessoas, limpa(p.recado, 500), id, 'até ' + limite];

    const sh = abaConfirmacoes_(planilha_());
    const ultima = sh.getLastRow();
    let alvo = ultima + 1;                 // linha nova…
    if (ultima > 1) {                      // …ou atualiza a mesma pessoa, se ela mudou a resposta
      const ids = sh.getRange(2, 6, ultima - 1, 1).getValues().map((r) => String(r[0]));
      const i = ids.indexOf(id);
      if (i >= 0) alvo = i + 2;
    }
    sh.getRange(alvo, 1, 1, linha.length).setValues([linha]);
    return resposta_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/**
 * ?acao=confirmar&...: grava a confirmação enviada pelo convite (é o caminho que o convite usa).
 * Sem senha: só mostra que está funcionando.
 * Com ?chave=SENHA certa: devolve a lista de confirmações para o painel da família.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.acao === 'confirmar') return registrar_(p);
  if (!p.chave) return resposta_({ ok: true, msg: 'Receptor de confirmações ativo' });
  if (p.chave !== CHAVE_PAINEL) return resposta_({ ok: false, erro: 'chave' });

  const sh = planilha_().getSheetByName(ABA);
  const n = sh ? sh.getLastRow() - 1 : 0;
  const lista = n > 0
    ? sh.getRange(2, 1, n, 5).getValues()
        .filter((r) => String(r[1]).trim())
        .map((r) => ({
          data: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
          nome: String(r[1]),
          vai: String(r[2]).indexOf('Vai') === 0,
          pessoas: Number(r[3]) || 0,
          recado: String(r[4] || '')
        }))
    : [];
  return resposta_({ ok: true, lista });
}

function abaConfirmacoes_(ss) {
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]).setFontWeight('bold').setBackground('#fde3ea');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 220);
    sh.setColumnWidth(3, 110);
    sh.setColumnWidth(5, 320);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.hideColumns(6);                     // ID técnico, a cliente não precisa ver
  }
  return sh;
}

function resposta_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Teste pelo editor: escolha "testarConfirmacao" e clique em Executar. Grava uma linha "Teste editor" na planilha. */
function testarConfirmacao() {
  const r = registrar_({ nome: 'Teste editor', vai: 'sim', pessoas: '9', cod: 'tq69s7', id: 'teste-editor' });
  Logger.log(r.getContent() + '  (pediu 9 pessoas num link de até 3: tem que gravar 3)');
}
