/**
 * Threshold Lab — recebe as respostas do formulário (GitHub Pages) e grava na planilha.
 * Este arquivo NÃO vai para o GitHub Pages; ele é colado no Apps Script da planilha.
 *
 * Como instalar (uma vez, ~3 min):
 * 1. Crie uma planilha Google chamada "Threshold Lab — respostas".
 * 2. Extensões → Apps Script. Apague o que vier e cole este arquivo inteiro. Salve (Ctrl+S).
 * 3. Escolha a função `configurar` e clique em Executar. Autorize (é o seu próprio script).
 *    Ela cria as abas "Respostas" e "Painel".
 * 4. Implantar → Nova implantação → tipo "App da Web".
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 *    Implantar → copie a URL que termina em /exec.
 * 5. Cole essa URL em ENDPOINT, no index.html (ou mande para o Claude publicar).
 *
 * Mudou este script? Implantar → Gerenciar implantações → editar → Nova versão.
 * Sem isso a URL continua rodando a versão antiga.
 */

const PONTOS = ['Horto Florestal', 'Beira-Mar Norte', 'Parque de Coqueiros', 'UFSC', 'Outro'];
const DIAS = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'];
const HORARIOS = ['Antes das 6h', '6h–7h', '7h–8h', '18h–19h', '19h–20h'];
const FAIXAS = ['L1', 'L2', 'L3', 'L4', 'Base'];

// [chave enviada pela página, título da coluna]
const COLUNAS = [
  ['_data', 'Data e hora'], ['origem', 'Origem'], ['faixa', 'Faixa'], ['limiar', 'Limiar (min/km)'],
  ['modo', 'Como informou'], ['fonte', 'Origem do número'], ['metros12', 'Teste 12 min (m)'],
  ['prova_dist', 'Prova · distância'], ['prova_tempo', 'Prova · tempo'], ['prova_mes', 'Prova · mês'],
  ['qualidade', 'Treinos de qualidade/mês'], ['pontos', 'Pontos'], ['dias', 'Dias'], ['horarios', 'Horários'],
  ['compromisso', 'Compromisso'], ['traz', 'Traz alguém'], ['nome', 'Nome'], ['whatsapp', 'WhatsApp'],
  ['instagram', 'Instagram'], ['objetivo', 'Objetivo 2027'], ['pagaria', 'Pagaria plano individual'],
  ['consentimento', 'Consentimento LGPD']
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // envios simultâneos não se sobrescrevem
  try {
    const p = (e && e.parameter) || {};
    if (p.site) return ok_();                 // campo-isca preenchido = robô
    if (p.consentimento !== 'Sim' || !p.nome || !p.whatsapp) return ok_();
    const sh = SpreadsheetApp.getActive().getSheetByName('Respostas');
    const linha = COLUNAS.map(([k]) => k === '_data' ? new Date() : limpa_(p[k]));
    sh.appendRow(linha);
    return ok_();
  } finally {
    lock.releaseLock();
  }
}

// GET só para conferir no navegador que a implantação está no ar.
function doGet() {
  return ContentService.createTextOutput('Threshold Lab: endpoint no ar.');
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ok: true})).setMimeType(ContentService.MimeType.JSON);
}

// Corta tamanho e neutraliza fórmula injetada (texto começando com = + - @).
function limpa_(v) {
  v = String(v == null ? '' : v).slice(0, 300);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function configurar() {
  const ss = SpreadsheetApp.getActive();

  // ---- Respostas ----
  let sh = ss.getSheetByName('Respostas') || ss.insertSheet('Respostas');
  const titulos = COLUNAS.map(c => c[1]);
  sh.getRange(1, 1, 1, titulos.length).setValues([titulos]).setFontWeight('bold').setBackground('#15181B').setFontColor('#F6F6F4');
  sh.setFrozenRows(1);
  sh.getRange('B:V').setNumberFormat('@');   // "4:35" fica texto, não vira hora
  sh.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && padrao.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(padrao);

  // ---- Painel ----
  let pn = ss.getSheetByName('Painel') || ss.insertSheet('Painel');
  pn.clear();
  const R = 'Respostas!';
  const col = t => R + '$' + letra_(titulos.indexOf(t) + 1) + ':$' + letra_(titulos.indexOf(t) + 1);
  const cFaixa = col('Faixa');

  const blocos = [
    ['Resumo', ['Respostas', 'Pode contar comigo', 'Talvez'], (rotulo, f) =>
      rotulo === 'Respostas' ? `=COUNTIF(${cFaixa},"${f}")`
        : `=COUNTIFS(${cFaixa},"${f}",${col('Compromisso')},"${rotulo}")`],
    ['Pontos', PONTOS, (v, f) => `=COUNTIFS(${cFaixa},"${f}",${col('Pontos')},"*${v}*")`],
    ['Dias', DIAS, (v, f) => `=COUNTIFS(${cFaixa},"${f}",${col('Dias')},"*${v}*")`],
    ['Horários', HORARIOS, (v, f) => `=COUNTIFS(${cFaixa},"${f}",${col('Horários')},"*${v}*")`]
  ];

  let r = 1;
  pn.getRange(r, 1).setValue('Threshold Lab · painel por faixa (atualiza sozinho)').setFontWeight('bold');
  r += 2;
  blocos.forEach(([nome, itens, formula]) => {
    pn.getRange(r, 1, 1, FAIXAS.length + 2).setValues([[nome].concat(FAIXAS, ['Total'])])
      .setFontWeight('bold').setBackground('#ECEDEA');
    itens.forEach((it, i) => {
      const linha = [it].concat(FAIXAS.map(f => formula(it, f)));
      linha.push(`=SUM(B${r + 1 + i}:${letra_(FAIXAS.length + 1)}${r + 1 + i})`);
      pn.getRange(r + 1 + i, 1, 1, linha.length).setValues([linha]);
    });
    r += itens.length + 2;
  });
  pn.getRange(r, 1).setValue('Leitura: precisa haver 6+ pessoas da mesma faixa no mesmo ponto, dia e horário. ' +
    'O cruzamento fino (ponto × dia × horário) se faz com filtro na aba Respostas.').setFontColor('#5E646A');
  pn.setColumnWidth(1, 200);
  pn.setFrozenRows(0);
  ss.setActiveSheet(pn);
}

function letra_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
