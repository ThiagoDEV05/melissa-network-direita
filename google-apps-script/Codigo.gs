/**
 * NETWORKING DA DIREITA: recebe os cadastros do site, grava nesta planilha
 * e entrega ao site a lista pública de profissionais.
 *
 * 1. Cole este código em Extensões → Apps Script (na planilha) e salve.
 * 2. Rode a função setup uma vez (cria a aba CADASTROS formatada).
 * 3. Implantar → Nova implantação → App da Web
 *    Executar como: Eu | Quem pode acessar: Qualquer pessoa
 * 4. Copie a URL (termina em /exec) e cole em site/config.js
 *
 * Ao atualizar este código: Implantar → Gerenciar implantações → lápis →
 * Versão: Nova versão → Implantar (a URL continua a mesma).
 */

const CONFIG = {
  ABA: 'CADASTROS',
  // Opcional: e-mail que recebe aviso a cada novo cadastro. Deixe '' para não avisar.
  EMAIL_AVISO: '',
  // false: o cadastro aparece no site na hora. Para tirar do ar, mude o Status para Recusado ou Inativo.
  // true: só aparece no site depois que o Status for mudado para Aprovado.
  EXIGIR_APROVACAO: false,
  // Segundos que a lista fica guardada antes de ler a planilha de novo.
  CACHE_SEGUNDOS: 30
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

/* ---------- API usada pelo site ---------- */

// Lista pública de profissionais (só dados comerciais autorizados).
function doGet() {
  return json_({ ok: true, mensagem: 'Formulário do Networking da Direita ativo.', profissionais: listar_() });
}

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true }); // campo invisível: só robôs preenchem
    const r = salvar_(d);
    return json_({ ok: true, id: r.id, publicado: !!r.perfil, perfil: r.perfil });
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  }
}

// Qualquer edição na planilha (ex.: mudar Status) atualiza a lista do site em segundos.
function onEdit() {
  CacheService.getScriptCache().remove('lista');
}

/* ---------- Leitura ---------- */

function listar_() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('lista');
  if (guardado) return JSON.parse(guardado);
  const lista = linhas_().map(publico_).filter(Boolean);
  try { cache.put('lista', JSON.stringify(lista), CONFIG.CACHE_SEGUNDOS); } catch (e) { /* lista grande demais para o cache */ }
  return lista;
}

function linhas_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const cab = v.shift().map(h => String(h).trim());
  return v.map(l => { const o = {}; cab.forEach((h, i) => { o[h] = l[i]; }); return o; });
}

// Converte uma linha da planilha no cartão público, ou null se não deve aparecer no site.
function publico_(r) {
  const s = k => String(r[k] == null ? '' : r[k]).trim();
  const st = s('Status');
  if (st === 'Recusado' || st === 'Inativo') return null;
  if (CONFIG.EXIGIR_APROVACAO && st !== 'Aprovado') return null;
  if (s('Autorizou divulgação') !== 'Sim' || !s('Nome completo')) return null;
  return {
    id: s('ID'), nome: s('Nome completo'), empresa: s('Nome da empresa'),
    profissao: s('Profissão'), especialidade: s('Especialidade'), categoria: s('Categoria profissional'),
    cidade: s('Cidade e estado'), instagram: usuario_(s('Instagram')),
    whatsapp: s('WhatsApp comercial').replace(/\D/g, ''), descricao: s('Descrição dos serviços')
  };
}

/* ---------- Gravação ---------- */

function salvar_(d) {
  const obrig = [['nome', 'nome'], ['profissao', 'profissão'], ['categoria', 'categoria'], ['cidade', 'cidade'],
    ['instagram', 'Instagram'], ['descricao', 'descrição'], ['parcerias', 'interesse em parcerias']];
  const faltando = obrig.filter(o => !String(d[o[0]] || '').trim()).map(o => o[1]);
  if (faltando.length) throw new Error('Preencha: ' + faltando.join(', ') + '.');
  if (d.autorizacao !== true) throw new Error('É preciso autorizar a divulgação dos dados.');

  const usuario = usuario_(d.instagram);
  if (!usuario) throw new Error('Instagram inválido.');
  const whats = String(d.whatsapp || '').replace(/\D/g, '').slice(0, 13);
  const link = 'https://www.instagram.com/' + usuario + '/';

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CONFIG.ABA) || (setup(), ss.getSheetByName(CONFIG.ABA));

    // Evita cadastro repetido (clique duplo ou envio de novo pela mesma pessoa).
    const repetido = linhas_().some(r => usuario_(r['Instagram']) === usuario && String(r['Status']).trim() !== 'Recusado');
    if (repetido) throw new Error('Este Instagram já está cadastrado no guia.');

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
    CacheService.getScriptCache().remove('lista');

    if (CONFIG.EMAIL_AVISO) {
      try {
        MailApp.sendEmail(CONFIG.EMAIL_AVISO, 'Novo cadastro: ' + v['Nome completo'],
          v['Nome completo'] + ' (' + v['Profissão'] + ', ' + v['Cidade e estado'] + ') se cadastrou.\n' + ss.getUrl());
      } catch (e) { /* sem cota de e-mail: segue sem aviso */ }
    }

    const linha = {};
    Object.keys(v).forEach(k => { linha[k] = typeof v[k] === 'string' ? v[k].replace(/^'/, '') : v[k]; });
    linha['WhatsApp comercial'] = whats;
    return { id: id, perfil: publico_(linha) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Utilitários ---------- */

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

function usuario_(v) {
  return String(v || '').trim().toLowerCase()
    .replace(/^https?:\/\/(www\.)?instagram\.com\//, '').replace(/^'?@/, '').split(/[/?#]/)[0]
    .replace(/[^a-z0-9._]/g, '').slice(0, 30);
}

// Corta o tamanho e impede que um texto enviado vire fórmula na planilha.
function texto_(v, max) {
  const s = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
