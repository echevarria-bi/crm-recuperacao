var XLSX = require('C:\\Users\\mar\\AppData\\Local\\Temp\\opencode\\node_modules\\xlsx');
var fs = require('fs');

var BASE_DIR = 'C:\\Users\\mar\\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\\Área de Trabalho\\';
var BASE = BASE_DIR + '_bases\\base_8026_2026.xlsx';
var PAINEL = BASE_DIR + 'dashboards\\crm-recuperacao\\Painel CRM - Junho.xlsx';
var OUT_DIR = BASE_DIR + 'dashboards\\crm-recuperacao\\';

var ACTIVE = ['1596', '1464', '1211', '1429', '9886', '1624'];
var NOMES = { '1596': 'Ariane', '1464': 'Camila', '1211': 'Cristielen', '1429': 'Natália', '9886': 'Tatiana', '1624': 'Tatiana' };
var NOMES_REVERSE = { 'Ariane': '1596', 'Camila': '1464', 'Cristielen': '1211', 'Natália': '1429', 'Natalia': '1429', 'Tatiana': '9886' };
var MES_NOMES = { 1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro' };
var MES_NUM = { 'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8, 'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12 };

function serialToDate(s) {
  if (!s || s < 60) return null;
  return new Date(Math.round((s - 25569) * 86400000));
}

// ============================================================
// 1. Ler base_8026 e indexar faturamento por CODCLI+Mês
// ============================================================
console.log('Lendo base_8026_2026.xlsx...');
var wb = XLSX.readFile(BASE);
var ws = wb.Sheets['Plan1'];
var raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
console.log('Total linhas: ' + raw.length);

var now = new Date();
var currentYear = now.getUTCFullYear();
var currentMonth = now.getUTCMonth() + 1;

// Index: fatIndex[codcli][mesNum] = total fat
var fatIndex = {};
var sellerMonth = {};
var sellerTotal = {};
var globalClientes = {};
var monthClientes = {};

ACTIVE.forEach(function (c) {
  sellerTotal[c] = { fat: 0, cli: {} };
  for (var m = 1; m <= 12; m++) {
    if (!sellerMonth[m]) sellerMonth[m] = {};
    sellerMonth[m][c] = { fat: 0, cli: {} };
  }
});

console.log('\n[1/4] Indexando faturamento por CODCLI+Mês...');
for (var ri = 1; ri < raw.length; ri++) {
  var r = raw[ri];
  if (!r || !r[27]) continue;
  var codusur = String(r[27]).trim();
  var dt = serialToDate(parseFloat(r[2]));
  if (!dt || dt.getUTCFullYear() !== currentYear) continue;
  var mes = dt.getUTCMonth() + 1;
  if (mes < 1 || mes > currentMonth) continue;
  var fat = parseFloat(r[34]) || (parseFloat(r[9]) || 0) + (parseFloat(r[10]) || 0);
  var codcli = String(r[11]).trim();
  var nome = String(r[12] || '').trim();

  // Index for faturamento lookup
  if (!fatIndex[codcli]) fatIndex[codcli] = {};
  if (!fatIndex[codcli][mes]) fatIndex[codcli][mes] = { fat: 0, nome: nome };
  fatIndex[codcli][mes].fat += fat;

  // Seller monthly data
  if (ACTIVE.indexOf(codusur) >= 0) {
    sellerMonth[mes][codusur].fat += fat;
    sellerTotal[codusur].fat += fat;
    if (fat >= 1) {
      sellerMonth[mes][codusur].cli[codcli] = 1;
      sellerTotal[codusur].cli[codcli] = 1;
      globalClientes[codcli] = { nome: nome, codcli: codcli };
      if (!monthClientes[mes]) monthClientes[mes] = {};
      monthClientes[mes][codcli] = { nome: nome, codcli: codcli, fat: 0, prof: NOMES[codusur] || codusur };
    }
    if (fat >= 1 && monthClientes[mes] && monthClientes[mes][codcli]) {
      monthClientes[mes][codcli].fat += fat;
    }
  }
}

// ============================================================
// 2. Ler Painel CRM - Junho.xlsx → aba FATURAMENTO RECUPERACAO
// ============================================================
console.log('[2/4] Lendo FATURAMENTO RECUPERACAO do Painel CRM...');
var wbPainel = XLSX.readFile(PAINEL);
var wsFat = wbPainel.Sheets['FATURAMENTO RECUPERACAO'];
var rawFat = XLSX.utils.sheet_to_json(wsFat, { header: 1, defval: '' });
console.log('Linhas na aba: ' + rawFat.length);

var MONTHS = ['Abril', 'Maio', 'Junho', 'Julho'];
var monthNums = [4, 5, 6, 7];

// Build faturamentoClientes from the sheet, refreshed from base_8026
var faturamentoClientesArr = [];
var atualizados = 0, naoEncontrados = 0;

for (var fi = 1; fi < rawFat.length; fi++) {
  var row = rawFat[fi];
  if (!row || !row[0]) continue;
  var prof = String(row[0]).trim();
  var codcli = String(row[1]).trim();
  var mesNome = String(row[4] || '').trim();
  var mesNum = MES_NUM[mesNome];

  if (!mesNum || MONTHS.indexOf(mesNome) < 0) continue;

  // Look up fresh faturamento from base_8026
  var fatAtual = 0;
  var nomeCliente = '';
  if (fatIndex[codcli] && fatIndex[codcli][mesNum]) {
    fatAtual = fatIndex[codcli][mesNum].fat;
    nomeCliente = fatIndex[codcli][mesNum].nome;
    atualizados++;
  } else {
    // Keep original value from sheet
    fatAtual = parseFloat(row[3]) || 0;
    nomeCliente = String(row[2] || '').trim();
    naoEncontrados++;
  }

  if (fatAtual > 0) {
    faturamentoClientesArr.push({
      cliente: parseInt(codcli) || 0,
      nome: nomeCliente,
      valor: Math.round(fatAtual * 100) / 100,
      mes: mesNome,
      prof: prof
    });
  }
}

faturamentoClientesArr.sort(function (a, b) { return b.valor - a.valor; });

console.log('  Clientes atualizados da base: ' + atualizados);
console.log('  Mantidos da planilha (não encontrados): ' + naoEncontrados);
console.log('  Total faturamentoClientes: ' + faturamentoClientesArr.length);

console.log('[3/4] Construindo data.json...');
var MONTHS = ['Abril', 'Maio', 'Junho', 'Julho'];
var monthNums = [4, 5, 6, 7];

var indicadores = [
  { label: 'Faturamento Total' },
  { label: 'Positivação (Clientes Únicos)' },
  { label: 'Ticket Médio' }
];
MONTHS.forEach(function (mk, mi) {
  var mNum = monthNums[mi];
  var totalFat = 0, totalCli = 0;
  ACTIVE.forEach(function (c) {
    var sm = sellerMonth[mNum] && sellerMonth[mNum][c];
    if (sm) { totalFat += sm.fat; totalCli += Object.keys(sm.cli).length; }
  });
  indicadores[0][mk] = Math.round(totalFat * 100) / 100;
  indicadores[1][mk] = totalCli;
  indicadores[2][mk] = totalCli > 0 ? Math.round(totalFat / totalCli * 100) / 100 : 0;
});

var valorPorProf = [];
ACTIVE.forEach(function (c) {
  var entry = { profissional: NOMES[c] || c };
  var total = 0;
  MONTHS.forEach(function (mk, mi) {
    var mNum = monthNums[mi];
    var sm = sellerMonth[mNum] && sellerMonth[mNum][c];
    var fat = sm ? Math.round(sm.fat * 100) / 100 : 0;
    if (fat > 0) entry[mk] = fat;
    total += fat;
  });
  entry.Total = Math.round(total * 100) / 100;
  valorPorProf.push(entry);
});

var recuperadosPorProf = [];
ACTIVE.forEach(function (c) {
  var entry = { profissional: NOMES[c] || c };
  MONTHS.forEach(function (mk, mi) {
    var mNum = monthNums[mi];
    var sm = sellerMonth[mNum] && sellerMonth[mNum][c];
    var cliCount = sm ? Object.keys(sm.cli).length : 0;
    entry[mk + ' META'] = cliCount;
    entry[mk + ' REAL'] = cliCount;
  });
  recuperadosPorProf.push(entry);
});

function buildPainelMonth(mNum) {
  var profs = [];
  var totalMeta = 0, totalReal = 0, totalFat = 0;
  ACTIVE.forEach(function (c) {
    var sm = sellerMonth[mNum] && sellerMonth[mNum][c];
    var fat = sm ? Math.round(sm.fat * 100) / 100 : 0;
    var cli = sm ? Object.keys(sm.cli).length : 0;
    profs.push({ profissional: NOMES[c] || c, meta: cli, realizado: cli, pct: 100, dias: [], faturamento: fat });
    totalMeta += cli; totalReal += cli; totalFat += fat;
  });
  return {
    recuperacao: { profissionais: profs, total: { profissional: 'Total', meta: totalMeta, realizado: totalReal, pct: 100, dias: [] }, faturamentoGeral: totalFat, dayNums: [] },
    metaContatos: { profissionais: [], total: null },
    motivos: [], inatividade: [], motivosDiarios: [], rcaExterno: null
  };
}

var faturamentoTotal = ACTIVE.reduce(function (s, c) { return s + sellerTotal[c].fat; }, 0);

var output = {
  data: {
    evolucaoMeses: { indicadores: indicadores, valorPorProf: valorPorProf, recuperadosPorProf: recuperadosPorProf },
    painelGeral: buildPainelMonth(6),
    painelAbril: buildPainelMonth(4),
    painelMaio: buildPainelMonth(5),
    painelJulho: buildPainelMonth(7),
    profissionais: {},
    faturamentoTotal: Math.round(faturamentoTotal * 100) / 100,
    faturamentoClientes: faturamentoClientesArr
  },
  fileName: 'base_8026_2026.xlsx',
  updatedAt: new Date().toISOString()
};

fs.writeFileSync(OUT_DIR + 'data.json', JSON.stringify(output, null, 2));
console.log('data.json salvo');

console.log('\n[4/4] Resumo:');
MONTHS.forEach(function (mk, mi) {
  var mNum = monthNums[mi];
  var totalFat = 0, totalCli = 0;
  ACTIVE.forEach(function (c) {
    var sm = sellerMonth[mNum] && sellerMonth[mNum][c];
    if (sm) { totalFat += sm.fat; totalCli += Object.keys(sm.cli).length; }
  });
  console.log('  ' + mk + ': R$ ' + totalFat.toFixed(2) + ' | ' + totalCli + ' clientes');
});

// Faturamento from sheet (recuperados)
var fatSheet = {};
faturamentoClientesArr.forEach(function (c) {
  if (!fatSheet[c.mes]) fatSheet[c.mes] = 0;
  fatSheet[c.mes] += c.valor;
});
console.log('\n  Faturamento Recuperados (da planilha):');
MONTHS.forEach(function (mk) {
  console.log('    ' + mk + ': R$ ' + (fatSheet[mk] || 0).toFixed(2));
});

console.log('\n  Total base: R$ ' + faturamentoTotal.toFixed(2));
console.log('  Clientes únicos: ' + Object.keys(globalClientes).length);
