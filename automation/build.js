// Pipeline: CSVs extraídos do Excel (data/raw/) -> shipments -> orders/sku_rows/summary -> diff contra a
// última rodada -> HTML final (fluxo-exportacao.html, na raiz do projeto), usando automation/template.html.
//
// Chamado pelo automation/atualizar_fluxo_exportacao.ps1 depois que o PowerShell já extraiu as abas do
// Excel para CSV. Pode também ser rodado manualmente: `node automation/build.js`.

const fs = require('fs');
const path = require('path');
const os = require('os');

const ROOT = path.join(__dirname, '..');
const RAW_DIR = path.join(ROOT, 'data', 'raw');
const DATA_DIR = path.join(ROOT, 'data');
const PREV_DIR = path.join(ROOT, 'data', 'previous');
const TEMPLATE_PATH = path.join(__dirname, 'template.html');
const OUTPUT_HTML = path.join(ROOT, 'fluxo-exportacao.html');
// O botão "Validar" do painel não tem como gravar direto no disco (é um HTML estático, sem
// servidor) — ele baixa um marcador pra essa pasta, e essa rodada de build (a próxima 8:10/15:30)
// é quem de fato aplica a validação, lendo os marcadores daqui.
const DOWNLOADS_DIR = path.join(os.homedir(), 'Downloads');

for (const dir of [DATA_DIR, PREV_DIR]) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

// ---------- CSV parsing ----------
function parseCSV(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else { inQuotes = false; }
      } else {
        field += c;
      }
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\r') { /* skip */ }
      else if (c === '\n') { row.push(field); field = ''; rows.push(row); row = []; }
      else field += c;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows;
}

const MONTHS_PT = { jan: 1, fev: 2, mar: 3, abr: 4, mai: 5, jun: 6, jul: 7, ago: 8, set: 9, out: 10, nov: 11, dez: 12 };

function excelSerialToDate(serial) {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(epoch.getTime() + serial * 86400000);
}

function parseDateValue(v, defaultYear) {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).trim();
  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = parseFloat(s);
    if (n > 20000 && n < 60000) return excelSerialToDate(n);
    return null;
  }
  // "^...pattern" (sem $ no final) de propósito: células como "17/03/2026 (FATURAR NF DIA
  // 13/03)" ou "29/04/2026 - NF 22/04/2026" trazem anotação extra depois da data real - o match
  // ancorado só no início pega a data e ignora o resto, em vez de falhar o regex inteiro.
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  m = s.match(/^(\d{1,2})-([a-zç]{3})\.?-(\d{2})/i);
  if (m) {
    const mon = MONTHS_PT[m[2].toLowerCase()];
    if (mon) return new Date(Date.UTC(2000 + parseInt(m[3], 10), mon - 1, +m[1]));
  }
  // "28/abr", "18/mai": dia/mês abreviado sem ano - usa o ano da própria aba (ex.: 2026).
  m = s.match(/^(\d{1,2})\/([a-zç]{3,4})\.?$/i);
  if (m && defaultYear) {
    const mon = MONTHS_PT[m[2].toLowerCase()];
    if (mon) return new Date(Date.UTC(defaultYear, mon - 1, +m[1]));
  }
  return null;
}

function toNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// column maps: index is 0-based into row array
//
// As 3 abas (2024/2025/2026) têm o MESMO layout de colunas 0-14 (cliente, tipo, pedido, invoice,
// data, volume, codigo, produto, pais, estufagem, [vazia], agente, transportadora, etiquetagem/obs,
// formato de data) - conferido direto nos CSVs exportados em 01/09/2026. Os schemas de 2024/2025
// estavam com índices errados (herdados de uma versão antiga da planilha com colunas em outra
// ordem) e nunca foram atualizados: "volume" lia a coluna de data, "produto" lia a coluna de
// volume, etc. - por isso o volume acumulado de 2024/2025 aparecia quase zerado no painel.
const SCHEMAS = {
  2026: { cliente: 0, tipo: 1, pedido: 2, invoice: 3, dataCarreg: 4, volume: 5, codigo: 6, produto: 7, pais: 8, estufagem: 9, agenteCarga: 11, transportadora: 12, etiquetagem: 13, formatoData: 14 },
  2025: { cliente: 0, tipo: 1, pedido: 2, invoice: 3, dataCarreg: 4, volume: 5, codigo: 6, produto: 7, pais: 8, estufagem: 9, agenteCarga: 11, transportadora: 12, etiquetagem: 13, formatoData: 14, valorInvoice: 23 },
  2024: { cliente: 0, tipo: 1, pedido: 2, invoice: 3, dataCarreg: 4, volume: 5, codigo: 6, produto: 7, pais: 8, estufagem: 9, agenteCarga: 11, transportadora: 12, obs: 13, formatoData: 14, valorInvoice: 17 },
};

const BAD_CODIGO = new Set(['#N/A', '#N/D', '#REF!', '#VALOR!', '#NOME?', '#NULO!', '#DIV/0!', '-', '']);

function loadCodigoColumn(filename) {
  if (!filename) return null;
  const p = path.join(RAW_DIR, filename);
  if (!fs.existsSync(p)) return null;
  const lines = fs.readFileSync(p, 'utf8').split(/\r?\n/).filter(Boolean);
  return lines.map(line => {
    const idx = line.indexOf(',');
    let val = line.slice(idx + 1);
    if (val.startsWith('"') && val.endsWith('"')) val = val.slice(1, -1).replace(/""/g, '"');
    val = val.trim();
    return BAD_CODIGO.has(val) ? '' : val;
  });
}

function extractYear(year, filename, codigoFilename) {
  const text = fs.readFileSync(path.join(RAW_DIR, filename), 'utf8');
  const rows = parseCSV(text);
  const schema = SCHEMAS[year];
  const codigoColumn = loadCodigoColumn(codigoFilename);
  let headerIdx = rows.findIndex(r => r.some(cell => cell.trim() === 'Cliente'));
  const dataRows = rows.slice(headerIdx + 1);

  const shipments = [];
  let current = null;

  for (let rowIdx = 0; rowIdx < dataRows.length; rowIdx++) {
    const r = dataRows[rowIdx];
    const codigo = codigoColumn ? (codigoColumn[headerIdx + 1 + rowIdx] || '') : '';
    if (r.every(c => (c || '').trim() === '')) continue;
    const clienteRaw = (r[schema.cliente] || '').trim();
    if (/^total(is)?$|^sub ?total$/i.test(clienteRaw)) { current = null; continue; }
    if (clienteRaw !== '') {
      current = {
        cliente: clienteRaw,
        tipo: schema.tipo !== undefined ? (r[schema.tipo] || '').trim() : '',
        pedido: (r[schema.pedido] || '').trim(),
        invoice: (r[schema.invoice] || '').trim(),
        pais: (r[schema.pais] || '').trim(),
        estufagem: (r[schema.estufagem] || '').trim(),
        transportadora: (r[schema.transportadora] || '').trim(),
        agenteCarga: schema.agenteCarga !== undefined ? (r[schema.agenteCarga] || '').trim() : '',
        etiquetagem: schema.etiquetagem !== undefined ? (r[schema.etiquetagem] || '').trim() : '',
        formatoData: schema.formatoData !== undefined ? (r[schema.formatoData] || '').trim() : '',
        obs: schema.obs !== undefined ? (r[schema.obs] || '').trim() : '',
        dataCarreg: parseDateValue(r[schema.dataCarreg], year),
        valorInvoice: schema.valorInvoice !== undefined ? toNumber(r[schema.valorInvoice]) : null,
        items: [],
      };
      shipments.push(current);
    }
    if (!current) continue;
    const vol = toNumber(r[schema.volume]);
    const produto = (r[schema.produto] || '').trim();
    if (vol !== null || produto !== '') {
      current.items.push({ volume: vol || 0, produto, codigo: codigo || '' });
    }
  }

  return shipments.map(s => ({
    ...s,
    year,
    totalVolume: s.items.reduce((a, b) => a + (b.volume || 0), 0),
  }));
}

// ---------- country name normalization ----------
function norm(s) {
  return (s || '').trim().toUpperCase().replace(/\s+/g, ' ');
}
const COUNTRY_FIX = {
  'AUSTRALIA': 'AUSTRÁLIA', 'CANADA': 'CANADÁ', 'ITALIA': 'ITÁLIA', 'SUIÇA': 'SUÍÇA', 'SUICA': 'SUÍÇA',
  'AUSTRIA': 'ÁUSTRIA', 'ESTADOS UNIDOS': 'EUA', 'USA': 'EUA', 'UK': 'REINO UNIDO',
  // Variantes sem acento/abreviadas mais comuns - cobertura ampliada (17/08/2026) pra qualquer
  // país novo já cair certo na bandeira/mapa sem precisar de ajuste manual depois.
  'SUECIA': 'SUÉCIA', 'FINLANDIA': 'FINLÂNDIA', 'DINAMARCA': 'DINAMARCA', 'ISLANDIA': 'ISLÂNDIA',
  'RUSSIA': 'RÚSSIA', 'POLONIA': 'POLÔNIA', 'ROMENIA': 'ROMÊNIA', 'UCRANIA': 'UCRÂNIA',
  'SERVIA': 'SÉRVIA', 'CROACIA': 'CROÁCIA', 'ESLOVENIA': 'ESLOVÊNIA', 'ESLOVAQUIA': 'ESLOVÁQUIA',
  'REPUBLICA TCHECA': 'REPÚBLICA TCHECA', 'TCHEQUIA': 'REPÚBLICA TCHECA', 'LITUANIA': 'LITUÂNIA',
  'LETONIA': 'LETÔNIA', 'ESTONIA': 'ESTÔNIA', 'ARMENIA': 'ARMÊNIA', 'AZERBAIJAO': 'AZERBAIJÃO',
  'ARABIA SAUDITA': 'ARÁBIA SAUDITA', 'LIBANO': 'LÍBANO', 'IRAQ': 'IRAQUE', 'IRA': 'IRÃ',
  'COREIA DO SUL': 'COREIA DO SUL', 'COREIA': 'COREIA DO SUL', 'TAILANDIA': 'TAILÂNDIA',
  'VIETNA': 'VIETNÃ', 'VIETNAM': 'VIETNÃ', 'INDONESIA': 'INDONÉSIA', 'INDIA': 'ÍNDIA',
  'PAQUISTAO': 'PAQUISTÃO', 'COLOMBIA': 'COLÔMBIA', 'BOLIVIA': 'BOLÍVIA', 'NICARAGUA': 'NICARÁGUA',
  'COSTA DO MARFIM': 'COSTA DO MARFIM', 'QUENIA': 'QUÊNIA', 'ETIOPIA': 'ETIÓPIA',
  'TUNISIA': 'TUNÍSIA', 'ARGELIA': 'ARGÉLIA', 'NIGERIA': 'NIGÉRIA',
  'ESCOSIA': 'ESCÓCIA', 'ESCOCIA': 'ESCÓCIA', 'SCOTLAND': 'ESCÓCIA',
};
function normCountry(s) {
  const v = norm(s);
  return COUNTRY_FIX[v] || v;
}

// ---------- shipments -> orders / sku rows ----------
function buildOrders(shipments) {
  return shipments.map((s, i) => ({
    id: i,
    cliente: (s.cliente || '').trim(),
    tipo: (s.tipo || '').trim(),
    pedido: (s.pedido || '').trim(),
    invoice: (s.invoice || '').trim(),
    pais: s.pais ? normCountry(s.pais) : '',
    estufagem: (s.estufagem || '').trim(),
    transportadora: (s.transportadora || '').trim(),
    agenteCarga: (s.agenteCarga || '').trim(),
    etiquetagem: (s.etiquetagem || '').trim(),
    formatoData: (s.formatoData || '').trim(),
    obs: (s.obs || '').trim(),
    data: s.dataCarreg,
    year: s.year,
    volume: Math.round(s.totalVolume),
    items: s.items
      .map(it => ({ produto: (it.produto || '').trim(), codigo: (it.codigo || '').trim(), volume: Math.round(it.volume || 0) }))
      .filter(it => it.produto || it.volume),
  }));
}

function buildSkuRows(orders) {
  const rows = [];
  orders.forEach(o => {
    o.items.forEach((it, ii) => {
      rows.push({
        id: o.id + '-' + ii, data: o.data, year: o.year, cliente: o.cliente, pedido: o.pedido,
        invoice: o.invoice, pais: o.pais, tipo: o.tipo, transportadora: o.transportadora,
        codigo: it.codigo, produto: it.produto, volume: it.volume,
      });
    });
  });
  return rows;
}

// ---------- summary aggregation (mirrors the dashboard's own client-side logic for the 'DATA' global) ----------
function topN(map, n) {
  return Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, n);
}
function topNWithOthers(map, n, label) {
  const entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
  const top = entries.slice(0, n);
  const rest = entries.slice(n).reduce((a, [, v]) => a + v, 0);
  const total = entries.reduce((a, [, v]) => a + v, 0);
  const result = top.map(([k, v]) => ({ [label]: k, volume: v, share: total ? v / total : 0 }));
  if (rest > 0) result.push({ [label]: 'OUTROS', volume: rest, share: total ? rest / total : 0, isOther: true });
  return { items: result, total, distinctCount: entries.length };
}

function buildSummary(skuRows) {
  const byYear = {}, byMonth = {}, byMonthCount = {}, byCountry = {}, byClient = {}, byProduct = {}, byTipo = {};
  const monthsPresent = new Set();

  // Need shipment-level totals for byYear/monthlySeries "shipments" counts; derive from skuRows via order id prefix.
  const orderIdsByMonth = {};
  const orderIdsTotal = new Set();

  skuRows.forEach(r => {
    byYear[r.year] = (byYear[r.year] || 0) + r.volume;
    const oi = r.id.split('-')[0];
    orderIdsTotal.add(r.year + ':' + oi);

    if (r.data) {
      const d = new Date(r.data);
      const key = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
      byMonth[key] = (byMonth[key] || 0) + r.volume;
      if (!orderIdsByMonth[key]) orderIdsByMonth[key] = new Set();
      orderIdsByMonth[key].add(oi);
      monthsPresent.add(key);
    }

    const country = r.pais ? normCountry(r.pais) : '';
    if (country) byCountry[country] = (byCountry[country] || 0) + r.volume;

    const client = norm(r.cliente);
    if (client) byClient[client] = (byClient[client] || 0) + r.volume;

    let tipo = norm(r.tipo);
    if (tipo === 'EXPORTAÇÃO DIRETA') tipo = 'DIRETA';
    if (tipo) byTipo[tipo] = (byTipo[tipo] || 0) + r.volume;

    const p = norm(r.produto).replace(/\s*-?\s*(CX|FD|SC|BX)\s*\d+.*$/, '').trim();
    if (p) byProduct[p] = (byProduct[p] || 0) + r.volume;
  });

  Object.keys(orderIdsByMonth).forEach(k => { byMonthCount[k] = orderIdsByMonth[k].size; });
  const months = Array.from(monthsPresent).sort();

  const countriesAgg = topNWithOthers(byCountry, 8, 'pais');
  const clientsAgg = topNWithOthers(byClient, 8, 'cliente');
  const productsAgg = topNWithOthers(byProduct, 10, 'produto');
  const tipoTotal = Object.values(byTipo).reduce((a, b) => a + b, 0);

  return {
    totals: { shipments: orderIdsTotal.size, volume: skuRows.reduce((a, r) => a + r.volume, 0) },
    byYear,
    monthlySeries: months.map(m => ({ month: m, volume: byMonth[m], shipments: byMonthCount[m] })),
    topCountries: countriesAgg.items,
    topClients: clientsAgg.items,
    topProducts: productsAgg.items,
    byTipo: topN(byTipo, 10).map(([k, v]) => ({ tipo: k, volume: v, share: tipoTotal ? v / tipoTotal : 0 })),
    countryCount: countriesAgg.distinctCount,
    clientCount: clientsAgg.distinctCount,
  };
}

// Normaliza data (objeto Date OU string ISO OU null) pro mesmo formato antes de comparar -
// necessário porque currentOrders acabou de ser montado (Date de verdade) enquanto
// previousOrders veio de JSON.parse (string ISO); comparar os dois formatos diretamente
// (ex.: via String()) nunca dá igual mesmo quando o dia é o mesmo.
function dateKey(d) {
  if (!d) return '';
  const dt = (d instanceof Date) ? d : new Date(d);
  return isNaN(dt.getTime()) ? '' : dt.toISOString();
}

// ---------- diff against the previous run ----------
function diffOrders(currentOrders, previousOrders) {
  const newInvoices = [];
  const changedInvoices = [];
  const itemChanges = [];
  const dateChanges = [];
  if (!previousOrders || !previousOrders.length) {
    return { generatedAt: new Date().toISOString(), newInvoices, changedInvoices, itemChanges, dateChanges };
  }
  const prevByInvoice = {};
  previousOrders.forEach(o => { if (o.invoice) prevByInvoice[o.invoice] = o; });

  currentOrders.forEach(o => {
    if (!o.invoice) return;
    const prev = prevByInvoice[o.invoice];
    let orderChanged = false;

    if (!prev) {
      // Pedido inteiro novo desde a ultima rodada (nao so um item dentro de um pedido ja
      // existente) - destaque proprio em verde no dashboard, sem marcar item por item como
      // "novo" (o pedido inteiro ja fica destacado, marcar cada linha seria redundante).
      if (o.items.length) newInvoices.push(o.invoice);
      return;
    } else {
      // Data de carregamento/navio mudou - o vendedor pode adiar/anticipar o embarque sem
      // avisar, o mesmo tipo de "mudança silenciosa" que os itens já cobrem. o.data aqui ainda
      // é um objeto Date de verdade (orders acabou de ser montado nesta rodada), enquanto
      // prev.data já é string ISO (veio de JSON.parse da base salva) - comparar via
      // String(Date) contra a string ISO SEMPRE dá diferente mesmo pro mesmo dia (formatos
      // distintos), o que gerou um falso positivo em quase todos os pedidos. dateKey() leva os
      // dois pro mesmo formato antes de comparar.
      var oldKey = dateKey(prev.data);
      var newKey = dateKey(o.data);
      if (oldKey !== newKey) {
        dateChanges.push({ invoice: o.invoice, oldDate: oldKey || null, newDate: newKey || null });
        orderChanged = true;
      }

      // Pedidos podem ter o mesmo produto em mais de uma linha (lotes/datas diferentes) -
      // por isso o pool guarda um MULTISET de volumes por produto+codigo, nao um mapa simples,
      // senao a segunda linha do mesmo produto "roubava" a comparacao da primeira e virava
      // falso positivo de alteracao a cada rodada, mesmo sem nada ter mudado de verdade.
      const pool = {};
      prev.items.forEach(it => {
        const key = it.produto + '|' + it.codigo;
        (pool[key] || (pool[key] = [])).push(it.volume);
      });
      o.items.forEach(it => {
        const key = it.produto + '|' + it.codigo;
        const candidates = pool[key];
        if (!candidates || !candidates.length) {
          itemChanges.push({ invoice: o.invoice, produto: it.produto, type: 'novo', oldVolume: 0, newVolume: it.volume });
          orderChanged = true;
          return;
        }
        const exactIdx = candidates.findIndex(v => Math.round(v) === Math.round(it.volume));
        if (exactIdx !== -1) {
          candidates.splice(exactIdx, 1);
        } else {
          const oldVolume = candidates.shift();
          itemChanges.push({ invoice: o.invoice, produto: it.produto, type: 'quantidade', oldVolume: Math.round(oldVolume), newVolume: Math.round(it.volume) });
          orderChanged = true;
        }
      });
    }
    if (orderChanged) changedInvoices.push(o.invoice);
  });

  return { generatedAt: new Date().toISOString(), newInvoices, changedInvoices, itemChanges, dateChanges };
}

// ---------- validação manual (botão "Validar" do painel) ----------
// O clique no botão baixa um arquivo "validar-<invoice>.json" pra pasta Downloads (única forma
// de um HTML estático "avisar" o computador sem servidor). Aqui a gente lê esses marcadores,
// aceita o estado ATUAL de cada invoice validado como novo baseline (fazendo o destaque sumir
// pra sempre, até uma mudança nova de verdade) e apaga o marcador pra não reprocessar.
function processValidationMarkers(baseline, currentOrders) {
  if (!fs.existsSync(DOWNLOADS_DIR)) return baseline;

  const byInvoice = {};
  baseline.forEach(o => { if (o.invoice) byInvoice[o.invoice] = o; });
  const currentByInvoice = {};
  currentOrders.forEach(o => { if (o.invoice) currentByInvoice[o.invoice] = o; });

  const files = fs.readdirSync(DOWNLOADS_DIR).filter(f => /^validar-.*\.json$/i.test(f));
  const processedInvoices = [];
  files.forEach(f => {
    const full = path.join(DOWNLOADS_DIR, f);
    try {
      const marker = JSON.parse(fs.readFileSync(full, 'utf8'));
      const current = marker.invoice && currentByInvoice[marker.invoice];
      if (current) { byInvoice[marker.invoice] = current; processedInvoices.push(marker.invoice); }
    } catch (e) {
      // marcador corrompido/ilegível - ignora o conteúdo, mas remove mesmo assim (linha abaixo)
    }
    fs.unlinkSync(full);
  });
  if (processedInvoices.length) {
    console.log('Validacao aplicada (destaque removido): ' + processedInvoices.join(', '));
  }

  return Object.values(byInvoice);
}

// ---------- main ----------
function main() {
  const all2024 = extractYear(2024, 'Programação_Exportações_2024.csv', null);
  const all2025 = extractYear(2025, 'Programação_Exportações_2025.csv', null);
  const all2026 = extractYear(2026, 'Programação_Exportações_2026.csv', 'codigo_2026.csv');
  const shipments = [...all2024, ...all2025, ...all2026];

  const orders = buildOrders(shipments);
  const skuRows = buildSkuRows(orders);
  const summary = buildSummary(skuRows);

  const prevOrdersPath = path.join(PREV_DIR, 'orders.json');
  let baseline = fs.existsSync(prevOrdersPath) ? JSON.parse(fs.readFileSync(prevOrdersPath, 'utf8')) : null;
  if (baseline) baseline = processValidationMarkers(baseline, orders);

  const changes = diffOrders(orders, baseline);

  fs.writeFileSync(path.join(DATA_DIR, 'orders.json'), JSON.stringify(orders), 'utf8');
  fs.writeFileSync(path.join(DATA_DIR, 'sku_rows.json'), JSON.stringify(skuRows), 'utf8');
  fs.writeFileSync(path.join(DATA_DIR, 'summary.json'), JSON.stringify(summary, null, 2), 'utf8');
  fs.writeFileSync(path.join(DATA_DIR, 'changes.json'), JSON.stringify(changes, null, 2), 'utf8');

  // A base só avança para invoices validados (processValidationMarkers, acima) - nunca é
  // substituída pelo estado atual inteiro. Se fosse, todo destaque sumiria sozinho na próxima
  // rodada agendada, que é exatamente o problema que a validação manual existe pra resolver.
  // Só na primeiríssima rodada (sem base ainda) é que ela nasce igual ao estado atual, porque
  // não há nada pra comparar/validar ainda.
  fs.writeFileSync(prevOrdersPath, JSON.stringify(baseline || orders), 'utf8');

  // Render final HTML from the template.
  let html = fs.readFileSync(TEMPLATE_PATH, 'utf8');
  html = html.replace('__DATA_JSON__', () => JSON.stringify(summary));
  html = html.replace('__SKU_JSON__', () => JSON.stringify(skuRows));
  html = html.replace('__ORDERS_JSON__', () => JSON.stringify(orders));
  html = html.replace('__CHANGES_JSON__', () => JSON.stringify(changes));
  fs.writeFileSync(OUTPUT_HTML, html, 'utf8');
  // Copia idêntica como index.html: é o nome que o Cloudflare Pages (e qualquer host estático)
  // serve automaticamente na raiz do site - sem ela, abrir o link publicado não mostraria nada.
  fs.writeFileSync(path.join(ROOT, 'index.html'), html, 'utf8');

  console.log('OK: %d pedidos, %d linhas de SKU, %d pedidos novos, %d pedidos alterados, %d itens novos/alterados',
    orders.length, skuRows.length, changes.newInvoices.length, changes.changedInvoices.length, changes.itemChanges.length);
  console.log('Painel gerado em: ' + OUTPUT_HTML);
}

main();
