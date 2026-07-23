var XLSX = require('C:\\Users\\mar\\AppData\\Local\\Temp\\opencode\\node_modules\\xlsx');
var fs = require('fs');

var BASE_DIR = 'C:\\Users\\mar\\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\\Área de Trabalho\\';
var BASE = BASE_DIR + '_bases\\base_8026_2026.xlsx';
var OUT_DIR = BASE_DIR + 'dashboards\\crm-recuperacao\\';

var ACTIVE = ['1596', '1464', '1211', '1429', '9886', '1624'];
var NOMES = { '1596': 'Ariane', '1464': 'Camila', '1211': 'Cristielen', '1429': 'Natália', '9886': 'Tatiana', '1624': 'Tatiana' };
var MES_NOMES = { 1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro' };

function serialToDate(s) {
  if (!s || s < 60) return null;
  return new Date(Math.round((s - 25569) * 86400000));
}

console.log('Lendo base_8026_2026.xlsx...');
var wb = XLSX.readFile(BASE);
var ws = wb.Sheets['Plan1'];
var raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
console.log('Total linhas: ' + raw.length);

var now = new Date();
var currentYear = now.getUTCFullYear();
var currentMonth = now.getUTCMonth() + 1;

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

console.log('\n[1/3] Extraindo dados por vendedor/mês...');
for (var ri = 1; ri < raw.length; ri++) {
  var r = raw[ri];
  if (!r || !r[27]) continue;
  var codusur = String(r[27]).trim();
  if (ACTIVE.indexOf(codusur) < 0) continue;
  var dt = serialToDate(parseFloat(r[2]));
  if (!dt || dt.getUTCFullYear() !== currentYear) continue;
  var mes = dt.getUTCMonth() + 1;
  if (mes < 1 || mes > currentMonth) continue;
  var fat = parseFloat(r[34]) || (parseFloat(r[9]) || 0) + (parseFloat(r[10]) || 0);
  var codcli = String(r[11]).trim();
  var nome = String(r[12] || '').trim();
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

console.log('[2/3] Construindo data.json...');
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

var faturamentoClientesArr = [];
Object.keys(monthClientes).forEach(function (m) {
  Object.keys(monthClientes[m]).forEach(function (c) {
    var cl = monthClientes[m][c];
    var mname = MES_NOMES[parseInt(m)] || '';
    if (MONTHS.indexOf(mname) >= 0 && cl.fat > 0) {
      faturamentoClientesArr.push({ cliente: parseInt(c) || 0, nome: cl.nome, valor: Math.round(cl.fat * 100) / 100, mes: mname, prof: cl.prof });
    }
  });
});
faturamentoClientesArr.sort(function (a, b) { return b.valor - a.valor; });

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

console.log('\n[3/3] Resumo:');
MONTHS.forEach(function (mk, mi) {
  var mNum = monthNums[mi];
  var totalFat = 0, totalCli = 0;
  ACTIVE.forEach(function (c) {
    var sm = sellerMonth[mNum] && sellerMonth[mNum][c];
    if (sm) { totalFat += sm.fat; totalCli += Object.keys(sm.cli).length; }
  });
  console.log('  ' + mk + ': R$ ' + totalFat.toFixed(2) + ' | ' + totalCli + ' clientes');
});
console.log('\n  Total: R$ ' + faturamentoTotal.toFixed(2));
console.log('  Clientes únicos: ' + Object.keys(globalClientes).length);
