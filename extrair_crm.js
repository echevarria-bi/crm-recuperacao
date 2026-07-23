var XLSX = require('C:\\Users\\mar\\AppData\\Local\\Temp\\opencode\\node_modules\\xlsx');
var fs = require('fs');

var BASE_DIR = 'C:\\Users\\mar\\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\\Área de Trabalho\\';
var BASE = BASE_DIR + '_bases\\base_8026_2026.xlsx';
var PAINEL = BASE_DIR + 'dashboards\\crm-recuperacao\\Painel CRM - Junho.xlsx';
var OUT_DIR = BASE_DIR + 'dashboards\\crm-recuperacao\\';

var ACTIVE = ['1596', '1464', '1211', '1429', '9886', '1624', '1571'];
var NOMES = { '1596': 'Ariane', '1464': 'Camila', '1211': 'Cristielen', '1429': 'Natália', '9886': 'Tatiana', '1624': 'Tatiana', '1571': 'Anna' };
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

// Build faturamentoClientes from the sheet (nova estrutura: A=TELEVENDEDORA, B=CODCLI, C=MES)
// Acumula faturamento do mes de referencia ate Julho
var faturamentoClientesArr = [];
var encontrados = 0, semFat = 0;

for (var fi = 1; fi < rawFat.length; fi++) {
  var row = rawFat[fi];
  if (!row || (!row[0] && row[0] !== 0)) continue;
  var prof = String(row[0]).trim();
  var codcli = String(row[1]).trim();
  var mesNome = String(row[2] || '').trim();
  var mesNum = MES_NUM[mesNome];

  if (!mesNum || mesNum < 4 || mesNum > 7) continue;

  // Acumula faturamento do mes de referencia ate Julho
  var somaFat = 0;
  var nomeCliente = '';
  for (var m = mesNum; m <= 7; m++) {
    if (fatIndex[codcli] && fatIndex[codcli][m]) {
      somaFat += fatIndex[codcli][m].fat;
      if (!nomeCliente) nomeCliente = fatIndex[codcli][m].nome;
    }
  }

  if (somaFat > 0) {
    faturamentoClientesArr.push({
      cliente: parseInt(codcli) || 0,
      nome: nomeCliente,
      valor: Math.round(somaFat * 100) / 100,
      mes: mesNome,
      prof: prof
    });
    encontrados++;
  } else {
    semFat++;
  }
}

faturamentoClientesArr.sort(function (a, b) { return b.valor - a.valor; });

console.log('  Clientes com faturamento: ' + encontrados);
console.log('  Sem faturamento na base: ' + semFat);
console.log('  Total faturamentoClientes: ' + faturamentoClientesArr.length);

// ============================================================
// 3. Consolidar dados da aba FATURAMENTO RECUPERACAO por mes
// ============================================================
console.log('[3/4] Construindo data.json a partir da aba FATURAMENTO RECUPERACAO...');
var MONTHS = ['Abril', 'Maio', 'Junho', 'Julho'];
var monthNums = [4, 5, 6, 7];

// Acumular faturamento e clientes por profissional + mes de referencia
// profFat[prof][mesRef] = soma acumulada (mesRef ate Julho)
// profCli[prof][mesRef] = clientes unicos com fat>0
var profFat = {}, profCli = {};
var totalFatPorMes = { 4: 0, 5: 0, 6: 0, 7: 0 };
var totalCliPorMes = { 4: 0, 5: 0, 6: 0, 7: 0 };

faturamentoClientesArr.forEach(function (c) {
  var mesNum = MES_NUM[c.mes];
  if (!mesNum) return;
  if (!profFat[c.prof]) profFat[c.prof] = {};
  if (!profCli[c.prof]) profCli[c.prof] = {};
  if (!profFat[c.prof][mesNum]) profFat[c.prof][mesNum] = 0;
  if (!profCli[c.prof][mesNum]) profCli[c.prof][mesNum] = 0;
  profFat[c.prof][mesNum] += c.valor;
  profCli[c.prof][mesNum]++;
  totalFatPorMes[mesNum] += c.valor;
  totalCliPorMes[mesNum]++;
});

// Arredondar
for (var p in profFat) {
  for (var m in profFat[p]) {
    profFat[p][m] = Math.round(profFat[p][m] * 100) / 100;
  }
}
for (var m in totalFatPorMes) totalFatPorMes[m] = Math.round(totalFatPorMes[m] * 100) / 100;

// Indicadores baseados na aba FATURAMENTO RECUPERACAO
var indicadores = [
  { label: 'Faturamento Total' },
  { label: 'Positivação (Clientes Únicos)' },
  { label: 'Ticket Médio' },
  { label: 'Contatos Realizados' },
  { label: 'Meta de Contatos' },
  { label: 'Meta de Clientes Recuperados' },
  { label: 'Realizado de Clientes Recuperados' },
  { label: 'Valor Recuperado' }
];
MONTHS.forEach(function (mk, mi) {
  var mNum = monthNums[mi];
  var tf = totalFatPorMes[mNum] || 0;
  var tc = totalCliPorMes[mNum] || 0;
  indicadores[0][mk] = tf;
  indicadores[1][mk] = tc;
  indicadores[2][mk] = tc > 0 ? Math.round(tf / tc * 100) / 100 : 0;
  indicadores[3][mk] = tc;
  indicadores[4][mk] = tc;
  indicadores[5][mk] = tc;
  indicadores[6][mk] = tc;
  indicadores[7][mk] = tf;
});

// Valor por profissional (acumulado a partir do mes de referencia)
var valorPorProf = [];
var nomesUsados = {};
faturamentoClientesArr.forEach(function (c) { nomesUsados[c.prof] = 1; });
Object.keys(nomesUsados).sort().forEach(function (prof) {
  var entry = { profissional: prof };
  var total = 0;
  MONTHS.forEach(function (mk, mi) {
    var mNum = monthNums[mi];
    var val = (profFat[prof] && profFat[prof][mNum]) || 0;
    if (val > 0) entry[mk] = val;
    total += val;
  });
  entry.Total = Math.round(total * 100) / 100;
  valorPorProf.push(entry);
});

// Recuperados por profissional
var recuperadosPorProf = [];
Object.keys(nomesUsados).sort().forEach(function (prof) {
  var entry = { profissional: prof };
  MONTHS.forEach(function (mk, mi) {
    var mNum = monthNums[mi];
    var cli = (profCli[prof] && profCli[prof][mNum]) || 0;
    entry[mk + ' META'] = cli;
    entry[mk + ' REAL'] = cli;
  });
  recuperadosPorProf.push(entry);
});

// Build painel por mes (acumulado a partir do mes de referencia)
function buildPainelMonth(mesRef) {
  var profs = [];
  var totalMeta = 0, totalReal = 0, totalFat = 0;
  Object.keys(nomesUsados).sort().forEach(function (prof) {
    var fat = (profFat[prof] && profFat[prof][mesRef]) || 0;
    var cli = (profCli[prof] && profCli[prof][mesRef]) || 0;
    fat = Math.round(fat * 100) / 100;
    profs.push({ profissional: prof, meta: cli, realizado: cli, pct: cli > 0 ? 100 : 0, dias: [], faturamento: fat });
    totalMeta += cli; totalReal += cli; totalFat += fat;
  });
  return {
    recuperacao: { profissionais: profs, total: { profissional: 'Total', meta: totalMeta, realizado: totalReal, pct: totalMeta > 0 ? 100 : 0, dias: [] }, faturamentoGeral: Math.round(totalFat * 100) / 100, dayNums: [] },
    metaContatos: { profissionais: [], total: null },
    motivos: [], inatividade: [], motivosDiarios: [], rcaExterno: null
  };
}

var faturamentoTotal = Math.round(faturamentoClientesArr.reduce(function (s, c) { return s + c.valor; }, 0) * 100) / 100;

var output = {
  data: {
    evolucaoMeses: { indicadores: indicadores, valorPorProf: valorPorProf, recuperadosPorProf: recuperadosPorProf },
    painelGeral: buildPainelMonth(6),
    painelAbril: buildPainelMonth(4),
    painelMaio: buildPainelMonth(5),
    painelJulho: buildPainelMonth(7),
    profissionais: {},
    faturamentoTotal: faturamentoTotal,
    faturamentoClientes: faturamentoClientesArr
  },
  fileName: 'base_8026_2026.xlsx',
  updatedAt: new Date().toISOString()
};

fs.writeFileSync(OUT_DIR + 'data.json', JSON.stringify(output, null, 2));
console.log('data.json salvo');

console.log('\n[4/4] Resumo (aba FATURAMENTO RECUPERACAO):');
MONTHS.forEach(function (mk, mi) {
  var mNum = monthNums[mi];
  console.log('  ' + mk + ': R$ ' + (totalFatPorMes[mNum] || 0).toFixed(2) + ' | ' + (totalCliPorMes[mNum] || 0) + ' clientes');
});
console.log('\n  Total acumulado: R$ ' + faturamentoTotal.toFixed(2));
console.log('  Total clientes: ' + faturamentoClientesArr.length);
