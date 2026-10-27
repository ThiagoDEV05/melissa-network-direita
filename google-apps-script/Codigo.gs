/**
 * NETWORKING DA DIREITA: recebe os cadastros do site e grava nesta planilha.
 *
 * 1. Cole este código em Extensões → Apps Script (na planilha) e salve.
 * 2. Rode a função setup uma vez (cria a aba CADASTROS formatada).
 * 3. Implantar → Nova implantação → App da Web
 *    Executar como: Eu | Quem pode acessar: Qualquer pessoa
 * 4. Copie a URL (termina em /exec) e cole em site/config.js
 */

const CONFIG = {
  ABA: 'CADASTROS',
  // Opcional: e-mail que recebe aviso a cada novo cadastro. Deixe '' para não avisar.
  EMAIL_AVISO: ''
};

const COLUNAS = [
  ['ID', 90], ['Data do cadastro', 130], ['Status', 100], ['Nome completo', 200], ['Nome da empresa', 170],
  ['Profissão', 150], ['Categoria profissional', 200], ['Especialidade', 160], ['Cidade e estado', 140],
  ['Instagram', 140], ['Link do Instagram', 240], ['WhatsApp comercial', 140], ['Descrição dos serviços', 340],
  ['Interesse em parcerias', 110], ['Autorizou divulgação', 100], ['Observações da administração', 220]
];
const CABECALHO = COLUNAS.map(c => c[0]);
const CATEGORIAS = [
  'Marketing e Comunicação', 'Advocacia e Jurídico', 'Saúde e Bem-estar', 'Beleza e Estética',
  'Contabilidade e Finanças', 'Tecnologia e Informática', 'Vendas e Comércio',
  'Empreendedorismo e Consultoria', 'Imóveis e Construção', 'Educação e Cursos',
  'Alimentação e Gastronomia', 'Prestadores de Serviços', 'Outros Profissionais'
];
const PARCERIAS = ['Sim', 'A combinar', 'Não'];
const STATUS = ['Novo', 'Aprovado', 'Recusado', 'Inativo'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.ABA) || ss.insertSheet(CONFIG.ABA, 0);
  sh.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#0A6B3B').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 40);
  COLUNAS.forEach((c, i) => sh.setColumnWidth(i + 1, c[1]));

  const linhas = sh.getMaxRows() - 1;
  sh.getRange(2, 3, linhas).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build());
  sh.getRange(2, 2, linhas).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2, 12, linhas).setNumberFormat('@');

  const tudo = sh.getRange(2, 1, linhas, CABECALHO.length);
  const regra = (st, cor) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="' + st + '"').setBackground(cor).setRanges([tudo]).build();
  sh.setConditionalFormatRules([
    regra('Novo', '#FFF5D2'), regra('Aprovado', '#E4EFE7'), regra('Recusado', '#F6E1E1'), regra('Inativo', '#EEEEEE')
  ]);
  if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), CABECALHO.length).createFilter();

  const vazia = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (vazia && vazia.getLastRow() === 0) ss.deleteSheet(vazia);
  ss.toast('Aba CADASTROS pronta. Agora implante como App da Web.', 'Networking da Direita', 10);
}

// Abrir a URL /exec no navegador mostra esta mensagem: serve para testar se está no ar.
function doGet() {
  return json_({ ok: true, mensagem: 'Formulário do Networking da Direita ativo.' });
}

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true }); // campo invisível: só robôs preenchem
    return json_({ ok: true, id: salvar_(d) });
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  }
}

function salvar_(d) {
  const obrig = [['nome', 'nome'], ['profissao', 'profissão'], ['categoria', 'categoria'], ['cidade', 'cidade'],
    ['instagram', 'Instagram'], ['descricao', 'descrição'], ['parcerias', 'interesse em parcerias']];
  const faltando = obrig.filter(o => !String(d[o[0]] || '').trim()).map(o => o[1]);
  if (faltando.length) throw new Error('Preencha: ' + faltando.join(', ') + '.');
  if (d.autorizacao !== true) throw new Error('É preciso autorizar a divulgação dos dados.');

  const usuario = String(d.instagram).trim().toLowerCase()
    .replace(/^https?:\/\/(www\.)?instagram\.com\//, '').replace(/^@/, '').split(/[/?#]/)[0]
    .replace(/[^a-z0-9._]/g, '').slice(0, 30);
  if (!usuario) throw new Error('Instagram inválido.');
  const whats = String(d.whatsapp || '').replace(/\D/g, '').slice(0, 13);
  const link = 'https://www.instagram.com/' + usuario + '/';

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CONFIG.ABA) || (setup(), ss.getSheetByName(CONFIG.ABA));
    const id = 'NDD-' + ('000' + proximoNumero_(sh)).slice(-4);
    const v = {
      'ID': id, 'Data do cadastro': new Date(), 'Status': 'Novo',
      'Nome completo': texto_(d.nome, 80), 'Nome da empresa': texto_(d.empresa, 80),
      'Profissão': texto_(d.profissao, 60),
      'Categoria profissional': CATEGORIAS.indexOf(d.categoria) >= 0 ? d.categoria : 'Outros Profissionais',
      'Especialidade': texto_(d.especialidade, 60), 'Cidade e estado': texto_(d.cidade, 60),
      'Instagram': "'@" + usuario, 'Link do Instagram': '',
      'WhatsApp comercial': whats ? "'" + whats : '',
      'Descrição dos serviços': texto_(d.descricao, 280),
      'Interesse em parcerias': PARCERIAS.indexOf(d.parcerias) >= 0 ? d.parcerias : 'A combinar',
      'Autorizou divulgação': 'Sim', 'Observações da administração': ''
    };
    sh.appendRow(CABECALHO.map(h => v[h]));
    sh.getRange(sh.getLastRow(), CABECALHO.indexOf('Link do Instagram') + 1)
      .setRichTextValue(SpreadsheetApp.newRichTextValue().setText(link).setLinkUrl(link).build());

    if (CONFIG.EMAIL_AVISO) {
      try {
        MailApp.sendEmail(CONFIG.EMAIL_AVISO, 'Novo cadastro: ' + v['Nome completo'],
          v['Nome completo'] + ' (' + v['Profissão'] + ', ' + v['Cidade e estado'] + ') se cadastrou.\n' + ss.getUrl());
      } catch (e) { /* sem cota de e-mail: segue sem aviso */ }
    }
    return id;
  } finally {
    lock.releaseLock();
  }
}

function proximoNumero_(sh) {
  let maior = 0;
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      const n = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      if (!isNaN(n) && n > maior) maior = n;
    });
  }
  return maior + 1;
}

// Corta o tamanho e impede que um texto enviado vire fórmula na planilha.
function texto_(v, max) {
  const s = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
