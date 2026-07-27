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
var purchaseIndex = {}; // purchaseIndex[codcli] = [{data, produto, qtde, punit, valor, frete, numnota, fornecedor, codfilial, codusur, mes}]
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

  // Detailed purchase index (all clients, all sellers)
  if (!purchaseIndex[codcli]) purchaseIndex[codcli] = [];
  var dataFmt = dt.getUTCFullYear() + '-' + String(dt.getUTCMonth()+1).padStart(2,'0') + '-' + String(dt.getUTCDate()).padStart(2,'0');
  purchaseIndex[codcli].push({
    data: dataFmt,
    dia: dt.getUTCDate(),
    mes: mes,
    produto: String(r[5] || '').trim(),
    qtde: parseFloat(r[7]) || 0,
    punit: parseFloat(r[8]) || 0,
    valor: Math.round(fat * 100) / 100,
    frete: parseFloat(r[10]) || 0,
    numnota: String(r[3] || '').trim(),
    fornecedor: String(r[26] || '').trim(),
    codfilial: String(r[0] || '').trim(),
    codusur: codusur
  });

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
// 2. Ler Painel CRM - Junho.xlsx → aba FATURAMENTO RECUPERACAO + Painel Geral
// ============================================================
console.log('[2/4] Lendo FATURAMENTO RECUPERACAO e Painel Geral...');
var wbPainel = XLSX.readFile(PAINEL);
var wsFat = wbPainel.Sheets['FATURAMENTO RECUPERACAO'];
var rawFat = XLSX.utils.sheet_to_json(wsFat, { header: 1, defval: '' });
console.log('Linhas na aba: ' + rawFat.length);

var MONTHS = ['Abril', 'Maio', 'Junho', 'Julho'];
var monthNums = [4, 5, 6, 7];

// Ler Meta/Realizado de cada aba Painel Geral (B9=Meta, C9=Realizado)
// E parsing completo da matriz META RECUPERACAO (profissionais, dias, meta/real por dia)
var painelSheets = {
  'Abril': 'Painel Geral - Abril',
  'Maio': 'Painel Geral - Maio ',
  'Junho': 'Painel Geral - Junho',
  'Julho': 'Painel Geral - Julho'
};
var metaRealPorMes = {};
var painelMatrizes = {};

MONTHS.forEach(function (mk) {
  var sn = painelSheets[mk];
  var ws = wbPainel.Sheets[sn];
  if (!ws) { console.log('  Aba "' + sn + '" não encontrada'); metaRealPorMes[mk] = { meta: 0, realizado: 0 }; return; }
  var data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Parse da matriz META RECUPERACAO
  var recRow = -1;
  for (var r = 0; r < data.length; r++) {
    var c = String((data[r] && data[r][0]) || '').trim().toUpperCase();
    if (c === 'META RECUPERACAO' || c === 'META RECUPERAÇÃO') { recRow = r; break; }
  }

  // Meta/Realizado total — find first "Total" row after recRow
  var totalIdx = -1;
  if (recRow >= 0) {
    for (var ti = recRow + 1; ti < data.length; ti++) {
      if (data[ti] && String(data[ti][0]).trim() === 'Total') { totalIdx = ti; break; }
    }
  }
  var row9 = totalIdx >= 0 ? data[totalIdx] : [];
  var meta = parseInt(row9[1]) || 0;
  var realizado = parseInt(row9[2]) || 0;
  metaRealPorMes[mk] = { meta: meta, realizado: realizado };
  var matriz = { dayNums: [], profissionais: [] };
  if (recRow >= 0) {
    // Day numbers from the same row as "META RECUPERACAO" (recRow), cols 4,6,8,...
    var dayRow = data[recRow] || [];
    for (var dc = 4; dc < dayRow.length; dc += 2) {
      var dv = dayRow[dc];
      if (dv !== '' && dv !== undefined && dv !== null) matriz.dayNums.push(parseInt(dv) || matriz.dayNums.length + 1);
    }
    // Profissionais from rows 4+ (recRow+2 = header, recRow+3..)
    var profStart = recRow + 2;
    for (var ri = profStart; ri < data.length; ri++) {
      var row = data[ri];
      if (!row || !row[0]) continue;
      var name = String(row[0]).trim();
      if (name === 'Total' || name === '' || name.indexOf('Conversão') >= 0 || name.indexOf('conversão') >= 0 || name.indexOf('Excedente') >= 0 || name.indexOf('GAP') >= 0) break;
      var prof = { profissional: name, meta: parseInt(row[1]) || 0, realizado: parseInt(row[2]) || 0, pct: 0, dias: [], faturamento: 0 };
      prof.pct = prof.meta > 0 ? Math.round(prof.realizado / prof.meta * 1000) / 10 : 0;
      // Dias: pares Meta/Real a partir da coluna 4
      for (var di = 0; di < matriz.dayNums.length; di++) {
        var metaCol = 4 + di * 2;
        var realCol = 5 + di * 2;
        var dMeta = parseInt(row[metaCol]) || 0;
        var dReal = parseInt(row[realCol]) || 0;
        prof.dias.push({ meta: dMeta, realizado: dReal });
      }
      matriz.profissionais.push(prof);
    }
  }
  painelMatrizes[mk] = matriz;
  console.log('  ' + mk + ': Meta=' + meta + ' Realizado=' + realizado + ' | ' + matriz.profissionais.length + ' profissionais, ' + matriz.dayNums.length + ' dias');
});

console.log('Linhas na aba: ' + rawFat.length);

// Pre-deduplicate: for same client (regardless of prof), keep only earliest recovery month
var earliestRecovery = {}; // key: codcli → earliest mesNum
for (var fi = 1; fi < rawFat.length; fi++) {
  var row = rawFat[fi];
  if (!row || (!row[0] && row[0] !== 0)) continue;
  var codcli = String(row[1]).trim();
  var mesNome = String(row[2] || '').trim();
  var mesNum = MES_NUM[mesNome];
  if (!mesNum || mesNum < 4 || mesNum > 7) continue;
  if (!earliestRecovery[codcli] || mesNum < earliestRecovery[codcli]) {
    earliestRecovery[codcli] = mesNum;
  }
}

// Build faturamentoClientes from the sheet (nova estrutura: A=TELEVENDEDORA, B=CODCLI, C=MES)
// 1 linha por cliente por mes com faturamento real daquele mes
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

  // Skip if this is not the earliest recovery for this client (regardless of prof)
  if (earliestRecovery[codcli] !== mesNum) continue;

  // 1 entrada por mes do mesRef ate Julho
  var nomeCliente = '';
  for (var m = mesNum; m <= 7; m++) {
    var val = 0;
    if (fatIndex[codcli] && fatIndex[codcli][m]) {
      val = Math.round(fatIndex[codcli][m].fat * 100) / 100;
      if (!nomeCliente) nomeCliente = fatIndex[codcli][m].nome;
    }
    faturamentoClientesArr.push({
      cliente: parseInt(codcli) || 0,
      nome: nomeCliente,
      valor: val,
      mes: MES_NOMES[m],
      prof: prof
    });
    if (val > 0) encontrados++;
    else semFat++;
  }
}

faturamentoClientesArr.sort(function (a, b) { return b.valor - a.valor; });

// Build purchase details per client: unique purchase days AFTER recovery month
// and full purchase details for click-to-expand
var clientesDetalhes = {};
var seenClients = {}; // track unique client+prof+mesRef combos
for (var fi2 = 1; fi2 < rawFat.length; fi2++) {
  var row2 = rawFat[fi2];
  if (!row2 || (!row2[0] && row2[0] !== 0)) continue;
  var prof2 = String(row2[0]).trim();
  var codcli2 = String(row2[1]).trim();
  var mesNome2 = String(row2[2] || '').trim();
  var mesRef = MES_NUM[mesNome2];
  if (!mesRef || mesRef < 4 || mesRef > 7) continue;

  var key = codcli2 + '_' + prof2 + '_' + mesRef;
  if (seenClients[key]) continue;
  seenClients[key] = 1;

  var purchases = purchaseIndex[codcli2] || [];
  // Use earliest recovery month if client appears multiple times
  var mesRefFinal = mesRef;
  if (clientesDetalhes[codcli2] && clientesDetalhes[codcli2].mesRefNum < mesRefFinal) {
    mesRefFinal = clientesDetalhes[codcli2].mesRefNum;
  }
  var afterRecovery = purchases.filter(function (p) { return p.mes >= mesRefFinal; });
  var uniqueDays = {};
  afterRecovery.forEach(function (p) {
    var dk = p.mes + '_' + p.dia;
    if (!uniqueDays[dk]) uniqueDays[dk] = { data: p.data, valor: 0, itens: 0 };
    uniqueDays[dk].valor += p.valor;
    uniqueDays[dk].itens++;
  });
  var uniqueDayCount = Object.keys(uniqueDays).length;

  clientesDetalhes[codcli2] = {
    comprasAposRecuperacao: uniqueDayCount,
    recuperadoEm: MES_NOMES[mesRefFinal],
    mesRefNum: mesRefFinal,
    detalhes: afterRecovery.map(function (p) {
      return { data: p.data, mes: MES_NOMES[p.mes], produto: p.produto, qtde: p.qtde, punit: p.punit, valor: p.valor, numnota: p.numnota, fornecedor: p.fornecedor };
    }),
    diasCompra: Object.keys(uniqueDays).map(function (dk) { return uniqueDays[dk]; }).sort(function (a,b) { return a.data > b.data ? 1 : -1; })
  };
}

console.log('  Linhas com faturamento > 0: ' + encontrados);
console.log('  Linhas com faturamento = 0: ' + semFat);
console.log('  Total linhas faturamentoClientes: ' + faturamentoClientesArr.length);

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

// Indicadores baseados na aba FATURAMENTO RECUPERACAO + Painel Geral
var indicadores = [
  { label: 'Clientes Ativos' },
  { label: 'Ticket Médio' },
  { label: 'Meta de Contatos' },
  { label: 'Realizado de Clientes Recuperados' },
  { label: 'Valor Recuperado' }
];
MONTHS.forEach(function (mk, mi) {
  var mNum = monthNums[mi];
  var tf = totalFatPorMes[mNum] || 0;
  var tc = totalCliPorMes[mNum] || 0;
  var pr = metaRealPorMes[mk] || { meta: 0, realizado: 0 };
  indicadores[0][mk] = tc;
  indicadores[1][mk] = tc > 0 ? Math.round(tf / tc * 100) / 100 : 0;
  indicadores[2][mk] = pr.meta;
  indicadores[3][mk] = pr.realizado;
  indicadores[4][mk] = tf;
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

// Recuperados por profissional — usar dados reais do Painel Geral (matriz)
var recuperadosPorProf = {};
MONTHS.forEach(function (mk) {
  var mat = painelMatrizes[mk] || { profissionais: [] };
  mat.profissionais.forEach(function (mp) {
    var baseName = mp.profissional.split(' - ')[0].trim();
    if (!recuperadosPorProf[baseName]) recuperadosPorProf[baseName] = { profissional: baseName };
    recuperadosPorProf[baseName][mk + ' META'] = mp.meta;
    recuperadosPorProf[baseName][mk + ' REAL'] = mp.realizado;
  });
});
var recuperadosPorProfArr = Object.keys(recuperadosPorProf).sort().map(function (k) { return recuperadosPorProf[k]; });

// Build painel por mes usando dados da matriz META RECUPERACAO + faturamento da base_8026
function buildPainelMonth(mesNome) {
  var mNum = MES_NUM[mesNome];
  var pr = metaRealPorMes[mesNome] || { meta: 0, realizado: 0 };
  var mat = painelMatrizes[mesNome] || { dayNums: [], profissionais: [] };

  // Enriquecer profissionais da matriz com faturamento da base_8026
  var profs = [];
  mat.profissionais.forEach(function (mp) {
    var fat = 0;
    // Buscar faturamento por nome, com normalização de acentos
    // Prioridade: 1) fullName completo, 2) fullName sem " - ", 3) baseName
    var baseName = mp.profissional.split(' - ')[0].trim();
    var fullName = mp.profissional.trim();
    var baseNameNorm = baseName.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    var fullNameNorm = fullName.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    var shortNameNorm = fullName.replace(/\s*-\s*/g, ' ').trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    var found = false;
    // 1) Try exact full name
    for (var pk in profFat) {
      if (pk.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase() === fullNameNorm) {
        fat = profFat[pk][mNum] || 0; found = true; break;
      }
    }
    // 2) Try short name (remove " - " separator, e.g. "Tatiana - LITORAL" → "Tatiana LITORAL")
    if (!found) {
      for (var pk in profFat) {
        if (pk.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase() === shortNameNorm) {
          fat = profFat[pk][mNum] || 0; found = true; break;
        }
      }
    }
    // 3) Try base name
    if (!found) {
      if (profFat[baseName]) {
        fat = profFat[baseName][mNum] || 0;
      } else {
        for (var pk in profFat) {
          if (pk.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase() === baseNameNorm) {
            fat = profFat[pk][mNum] || 0; break;
          }
        }
      }
    }
    profs.push({
      profissional: mp.profissional,
      meta: mp.meta,
      realizado: mp.realizado,
      pct: mp.pct,
      dias: mp.dias,
      faturamento: Math.round(fat * 100) / 100
    });
  });

  var pct = pr.meta > 0 ? Math.round(pr.realizado / pr.meta * 1000) / 10 : 0;
  return {
    recuperacao: {
      profissionais: profs,
      total: { profissional: 'Total', meta: pr.meta, realizado: pr.realizado, pct: pct, dias: [] },
      faturamentoGeral: Math.round(profs.reduce(function (s, p) { return s + p.faturamento; }, 0) * 100) / 100,
      dayNums: mat.dayNums
    },
    metaContatos: { profissionais: [], total: null },
    motivos: [], inatividade: [], motivosDiarios: [], rcaExterno: null
  };
}

// ============================================================
// 5. Ler aba RCA EXTERNO — coluna A (CODCLI) + matriz H-J
// ============================================================
console.log('\n[5/5] Lendo aba RCA EXTERNO...');
var wbPainel = XLSX.readFile(PAINEL);
var wsRcaExt = wbPainel.Sheets['RCA EXTERNO'];
var rcaExternoCodclis = [];
var rcaExternoMatriz = [];
if (wsRcaExt) {
  var rawRcaExt = XLSX.utils.sheet_to_json(wsRcaExt, { header: 1, defval: '' });
  for (var ri = 0; ri < rawRcaExt.length; ri++) {
    var rr = rawRcaExt[ri];
    // Coluna A: CODCLI
    if (ri >= 1 && rr[0] !== '' && rr[0] !== null && rr[0] !== undefined) {
      rcaExternoCodclis.push(String(rr[0]).trim());
    }
    // Colunas H-J: matriz
    if (rr[7] !== '' && rr[7] !== null && rr[7] !== undefined) {
      rcaExternoMatriz.push({ mes: String(rr[7]).trim(), clientes: parseFloat(rr[8]) || 0, faturamento: parseFloat(rr[9]) || 0 });
    }
  }
} else {
  console.log('  Aba RCA EXTERNO não encontrada!');
}
console.log('  CODCLI externos: ' + rcaExternoCodclis.length);
console.log('  Matriz: ' + rcaExternoMatriz.length + ' linhas');
rcaExternoMatriz.forEach(function(m) { console.log('    ' + m.mes + ': ' + m.clientes + ' clientes | R$ ' + (m.faturamento || 0).toFixed(2)); });

var faturamentoTotal = Math.round(faturamentoClientesArr.reduce(function (s, c) { return s + c.valor; }, 0) * 100) / 100;

var output = {
  data: {
    evolucaoMeses: { indicadores: indicadores, valorPorProf: valorPorProf, recuperadosPorProf: recuperadosPorProfArr },
    painelGeral: buildPainelMonth('Junho'),
    painelAbril: buildPainelMonth('Abril'),
    painelMaio: buildPainelMonth('Maio'),
    painelJulho: buildPainelMonth('Julho'),
    profissionais: {},
    faturamentoTotal: faturamentoTotal,
    faturamentoClientes: faturamentoClientesArr,
    clientesDetalhes: clientesDetalhes,
    rcaExterno: { codclis: rcaExternoCodclis, matriz: rcaExternoMatriz }
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
