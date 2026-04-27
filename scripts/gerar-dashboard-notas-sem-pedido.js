#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const https = require('https');
const { execFileSync } = require('child_process');

const SPREADSHEET_ID = '1g_YvcsZN-jXDSrUevgPYMmg8daxSu2NC8XIWAnjHHBI';
const SPREADSHEET_TITLE = 'Notas fiscais sem pedido';
const GOG = 'C:\\Users\\gabriel.abel\\AppData\\Roaming\\npm\\gog.exe';
const OUTPUT = path.resolve(__dirname, '..', 'index.html');
const TEAM_SHEET = 'EQUIPE STEIN';
const DATA_SHEETS = [
 'ITA',
 'STEIN',
 'STEIN E BERTEMES',
 'HASA 13',
 'COSMOPOLITAN',
 'STEIN LITORAL',
 'VERTIKAL',
 'STEIN ALAMEDA',
 'ELS PARTICIPAÇÕES',
 'COSTA ESMERALDA'
];

function gogJson(args) {
 const finalArgs = [...args];
 if (!finalArgs.includes('--account')) finalArgs.push('--account', 'suporte.ti@cstein.com.br');
 const raw = execFileSync(GOG, finalArgs, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
 return JSON.parse(raw);
}

function norm(v) { return String(v ?? '').trim(); }
function slug(v) { return norm(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); }
function parseMoney(v) {
 const s = norm(v).replace(/R\$/gi, '').replace(/\./g, '').replace(',', '.').replace(/\s+/g, '');
 const n = Number.parseFloat(s);
 return Number.isFinite(n) ? n : 0;
}
function parseDateBR(v) {
 const s = norm(v);
 const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
 if (!m) return null;
 return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}
function daysDiff(a, b) { return Math.floor((a - b) / 86400000); }
function fmtMoney(v) { return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(v || 0); }
function fmtDate(d) { return d ? new Intl.NumberFormat('pt-BR').format(d) : '-'; }
function esc(s) { return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function buildMapFromRows(values) {
 const header = values[0] || [];
 return (values.slice(1) || []).map(row => {
  const obj = {};
  for (let i = 0; i < header.length; i++) obj[norm(header[i])] = norm(row[i]);
  return obj;
 });
}

function googleApiJson(url) {
 return new Promise((resolve, reject) => {
  let token;
  try {
   token = JSON.parse(fs.readFileSync('C:\\Users\\gabriel.abel\\.openclaw\\workspace\\google-token.json', 'utf8')).access_token;
  } catch {
   token = JSON.parse(fs.readFileSync('C:\\Users\\gabriel.abel\\.openclaw\\google\\token.json', 'utf8')).access_token;
  }
  https.get(url, { headers: { Authorization: 'Bearer ' + token } }, res => {
   let d = '';
   res.on('data', c => d += c);
   res.on('end', () => {
    if (res.statusCode !== 200) return reject(new Error('Google API ' + res.statusCode + ': ' + d));
    try { resolve(JSON.parse(d)); } catch (e) { reject(e); }
   });
  }).on('error', reject);
 });
}

function isRed(c) {
 const r = Number(c?.red ?? 1), g = Number(c?.green ?? 1), b = Number(c?.blue ?? 1);
 return r > 0.85 && g < 0.93 && b < 0.93 && r > g && r > b;
}

async function loadColorFlags() {
 const ranges = DATA_SHEETS.map(s => 'ranges=' + encodeURIComponent(s + '!A1:Z5000')).join('&');
 const fields = encodeURIComponent('sheets(properties/title,data(rowData(values(userEnteredValue,effectiveFormat/backgroundColor))))');
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}?includeGridData=true&${ranges}&fields=${fields}`;
 const json = await googleApiJson(url);
 const out = new Map();
 for (const sheet of json.sheets || []) {
  const title = sheet.properties?.title || '';
  const rows = sheet.data?.[0]?.rowData || [];
  const head = (rows[0]?.values || []).map(v => String(v.userEnteredValue?.stringValue || '').trim());
  const idxSol = head.findIndex(h => /SOLICITA/i.test(h));
  const idxPed = head.findIndex(h => /PEDIDO|MEDI/i.test(h));
  for (let r = 1; r < rows.length; r++) {
   const vals = rows[r].values || [];
   out.set(title + '::' + (r + 1), {
    erroSolicitacao: idxSol >= 0 ? isRed(vals[idxSol]?.effectiveFormat?.backgroundColor) : false,
    erroPedidoMedicao: idxPed >= 0 ? isRed(vals[idxPed]?.effectiveFormat?.backgroundColor) : false
   });
  }
 }
 return out;
}

function loadContacts() {
 const team = gogJson(['sheets', 'get', SPREADSHEET_ID, `${TEAM_SHEET}!A1:Z300`, '--json']);
 const values = team.values || [];
 const contactsByObra = new Map();
 const emailsByEngineer = new Map();

 for (const row of values.slice(1)) {
  const engenheiro = norm(row[0]);
  const fone = norm(row[1]);
  const obra = norm(row[2]);
  const engenheiroEmailNome = norm(row[8]);
  const email = norm(row[9]);

  if (obra) contactsByObra.set(slug(obra), { obra, engenheiro, fone });
  if (engenheiroEmailNome && email) {
   emailsByEngineer.set(slug(engenheiroEmailNome), email);
   const primeiroNome = slug(engenheiroEmailNome).split(' ')[0];
   if (primeiroNome) emailsByEngineer.set(primeiroNome, email);
  }
 }

 return { contactsByObra, emailsByEngineer };
}

function pick(row, keys) {
 for (const key of keys) if (norm(row[key])) return norm(row[key]);
 return '';
}

function normalizeObra(sheet, obra) {
 const s = norm(obra);
 if (!s) return sheet;
 if (sheet === 'HASA 13' && slug(s) === 'hasa 13') return 'RESERVA';
 return s;
}

function classifyStatus(rec) {
 const tipo = norm(rec.tipoLancamento).toUpperCase();
 if (rec.lancamentoNf) return 'Lançado no Mega';
 if (tipo === 'LIVRE') return 'Pode lançar';
 if (rec.solicitacao && rec.pedidoMedicao) return 'Pode lançar';
 if (rec.solicitacao && !rec.pedidoMedicao) return 'Com solicitação, sem pedido/medição';
 return 'Sem solicitação';
}

function isValidOperationalObra(v, empresa = '') {
 const s = norm(v);
 if (!s) return false;
 const x = slug(s);
 const empresaSlug = slug(empresa);
 const bloqueios = ['sem obra', 'sem info', 'verificando', 'escritorio', 'stein', 'ita', 'vertikal', 'costa esmeralda', 'els participacoes', 'hasa 13'];
 if (bloqueios.includes(x)) return false;
 if (empresaSlug && x === empresaSlug) return false;
 return true;
}

function displayObra(v, empresa) {
 const s = norm(v);
 if (!s) return empresa;
 return s;
}

function loadLegacyDocs(sheet, rows, contactsByObra, emailsByEngineer, today) {
 const docs = [];
 for (const row of rows) {
  const notaFiscal = pick(row, ['Nota Fiscal']);
  const fornecedor = pick(row, ['Razão Social']);
  const dataEmissaoRaw = pick(row, ['Data de Emissão']);
  if (!notaFiscal && !fornecedor && !dataEmissaoRaw) continue;
  const obra = normalizeObra(sheet, pick(row, ['OBRA']) || sheet);
  const contact = contactsByObra.get(slug(obra)) || {};
  const engineerName = pick(row, ['ENGENHEIRO']) || contact.engenheiro || '';
  const engineerEmail = emailsByEngineer.get(slug(engineerName)) || emailsByEngineer.get(slug(engineerName).split(' ')[0]) || '';
  const dataEmissao = parseDateBR(dataEmissaoRaw);
  const daysLate = dataEmissao ? Math.max(0, daysDiff(today, dataEmissao) - 2) : 0;
  const solicitacao = pick(row, ['NÚMERO SOLICITAÇÃO']);
  const rec = {
   aba: sheet,
   tipoNota: pick(row, ['Status']),
   dataEmissao,
   dataEmissaoRaw,
   notaFiscal,
   cnpjCpf: pick(row, ['CNPJ / CPF']),
   razaoSocial: fornecedor,
   valor: parseMoney(pick(row, ['Valor Total'])),
   obra: displayObra(obra, sheet),
   tipoLancamento: '',
   solicitacao,
   pedidoMedicao: '',
   lancamentoNf: '',
   observacao: '',
   engenheiro: engineerName,
   contatoFone: contact.fone || '',
   contatoEmail: engineerEmail || '',
   daysLate,
   erroSolicitacao: false,
   erroPedidoMedicao: false
  };
  rec.statusOperacional = solicitacao ? 'Com solicitação, sem pedido/medição' : 'Sem solicitação';
  rec.dataEmissaoIso = dataEmissao ? new Date(dataEmissao.getTime() - dataEmissao.getTimezoneOffset()*60000).toISOString().slice(0,10) : '';
  rec.dataEmissaoBr = dataEmissao ? fmtDate(dataEmissao) : '';
  docs.push(rec);
 }
 return docs;
}

async function loadDocs(contactData) {
 const docs = [];
 const today = new Date();
 const colorFlags = await loadColorFlags();
 for (const sheet of DATA_SHEETS) {
  const json = gogJson(['sheets', 'get', SPREADSHEET_ID, `${sheet}!A1:Z5000`, '--json']);
  const rows = buildMapFromRows(json.values || []);
  const headerKeys = Object.keys(rows[0] || {});
  const isLegacy = headerKeys.includes('NÚMERO SOLICITAÇÃO') && !headerKeys.includes('LANÇAMENTO NF');
  if (isLegacy) {
   const legacyDocs = loadLegacyDocs(sheet, rows, contactData.contactsByObra, contactData.emailsByEngineer, today);
   legacyDocs.forEach((rec, idx) => {
    const f = colorFlags.get(sheet + '::' + (idx + 2));
    if (f) {
     rec.erroSolicitacao = !!f.erroSolicitacao;
     rec.erroPedidoMedicao = !!f.erroPedidoMedicao;
    }
   });
   docs.push(...legacyDocs);
   continue;
  }
  rows.forEach((row, idx) => {
   const notaFiscal = pick(row, ['NOTA FISCAL']);
   const fornecedor = pick(row, ['RAZÃO SOCIAL']);
   const dataEmissaoRaw = pick(row, ['DATA DE EMISSÃO']);
   if (!notaFiscal && !fornecedor && !dataEmissaoRaw) return;
   const obra = normalizeObra(sheet, pick(row, ['OBRA']) || sheet);
   const contact = contactData.contactsByObra.get(slug(obra)) || {};
   const dataEmissao = parseDateBR(dataEmissaoRaw);
   const daysLate = dataEmissao ? Math.max(0, daysDiff(today, dataEmissao) - 2) : 0;
   const engineerName = pick(row, ['ENGENHEIRO']) || contact.engenheiro || '';
   const engineerEmail = contactData.emailsByEngineer.get(slug(engineerName)) || contactData.emailsByEngineer.get(slug(engineerName).split(' ')[0]) || '';
   const rec = {
    aba: sheet,
    tipoNota: pick(row, ['TIPO DE NOTA']),
    dataEmissao,
    dataEmissaoRaw,
    notaFiscal,
    cnpjCpf: pick(row, ['CNPJ / CPF']),
    razaoSocial: fornecedor,
    valor: parseMoney(pick(row, ['VALOR TOTAL'])),
    obra: displayObra(obra, sheet),
    tipoLancamento: pick(row, ['TIPO LANÇAMENTO']),
    solicitacao: pick(row, ['SOLICITAÇÃO']),
    pedidoMedicao: pick(row, ['PEDIDO/MEDIÇÃO']),
    lancamentoNf: pick(row, ['LANÇAMENTO NF']),
    observacao: pick(row, ['OBSERVAÇÃO']),
    engenheiro: engineerName,
    contatoFone: contact.fone || '',
    contatoEmail: engineerEmail || '',
    daysLate,
    erroSolicitacao: false,
    erroPedidoMedicao: false
   };
   rec.statusOperacional = classifyStatus(rec);
   rec.dataEmissaoIso = dataEmissao ? new Date(dataEmissao.getTime() - dataEmissao.getTimezoneOffset()*60000).toISOString().slice(0,10) : '';
   rec.dataEmissaoBr = dataEmissao ? fmtDate(dataEmissao) : '';
   const f = colorFlags.get(sheet + '::' + (idx + 2));
   if (f) {
    rec.erroSolicitacao = !!f.erroSolicitacao;
    rec.erroPedidoMedicao = !!f.erroPedidoMedicao;
   }
   docs.push(rec);
  });
 }
 return docs;
}

function buildHtml(data) {
 return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Dashboard Notas fiscais sem pedido</title>
<style>
:root{--bg:#06101d;--bg-soft:#0a1628;--panel:#0d1b31;--panel-2:#11213d;--line:rgba(255,255,255,.08);--line-soft:rgba(148,163,184,.12);--text:#eef4ff;--muted:#94a3b8;--muted-2:#c7d2e1;--blue:#60a5fa;--cyan:#22d3ee;--violet:#8b5cf6;--green:#22c55e;--amber:#f59e0b;--red:#ef4444;--shadow:0 20px 50px rgba(0,0,0,.28)}
*{box-sizing:border-box} html{scroll-behavior:smooth} body{margin:0;font-family:Inter,Segoe UI,Arial,sans-serif;color:var(--text);background:radial-gradient(circle at top left,rgba(34,211,238,.16) 0,transparent 24%),radial-gradient(circle at top right,rgba(139,92,246,.18) 0,transparent 28%),linear-gradient(180deg,#07111f 0%,#050c16 100%)}
body:before{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(255,255,255,.018) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.018) 1px,transparent 1px);background-size:24px 24px;pointer-events:none;mask-image:linear-gradient(180deg,rgba(0,0,0,.55),transparent 92%)}
.container{position:relative;max-width:1440px;margin:0 auto;padding:30px}.hero{display:flex;justify-content:space-between;gap:20px;align-items:stretch;flex-wrap:wrap;margin-bottom:22px}.hero-main,.hero-side,.card{position:relative;background:linear-gradient(180deg,rgba(13,27,49,.92),rgba(7,14,27,.96));border:1px solid var(--line);box-shadow:var(--shadow);backdrop-filter:blur(14px);border-radius:28px}.hero-main{flex:1 1 850px;padding:28px;overflow:hidden}.hero-side{flex:0 0 290px;padding:24px;display:flex;flex-direction:column;justify-content:space-between}.hero h1{margin:0 0 10px;font-size:40px;line-height:1.05;letter-spacing:-.03em}.hero p{margin:0;color:var(--muted-2);max-width:860px;font-size:15px;line-height:1.6}.hero-meta{margin-top:18px;display:flex;gap:12px;flex-wrap:wrap}.hero-pill,.filter-note{padding:10px 14px;border-radius:999px;background:rgba(255,255,255,.05);border:1px solid var(--line-soft);font-size:13px;color:#dbe6f5}.stamp-title{font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}.stamp-value{font-size:24px;font-weight:800;line-height:1.15}.stamp-sub{margin-top:10px;color:var(--muted-2);font-size:13px;line-height:1.5}
.filter-bar{display:grid;grid-template-columns:minmax(220px,280px) minmax(180px,1fr) minmax(180px,1fr) auto;gap:12px;align-items:start;margin-bottom:12px}.filter-group label{display:block;font-size:11px;color:var(--muted);margin-bottom:6px;text-transform:uppercase;letter-spacing:.08em}.filter-group input{width:100%;padding:10px 12px;border-radius:12px;background:#0a1628;color:var(--text);border:1px solid var(--line-soft);outline:none;min-height:42px}.btn{padding:11px 16px;border-radius:12px;border:1px solid rgba(96,165,250,.24);background:linear-gradient(90deg,rgba(34,211,238,.18),rgba(96,165,250,.18));color:var(--text);font-weight:700;cursor:pointer;min-height:42px}.company-list{display:flex;flex-direction:column;gap:4px}.company-item{position:relative;display:flex;align-items:center;gap:10px;padding:7px 12px;border-radius:8px;cursor:pointer;font-size:11.5px;font-weight:600;letter-spacing:.04em;color:rgba(219,230,245,.7);border:1px solid transparent;transition:background .15s,border-color .15s,color .15s;user-select:none}.company-item::before{content:'';flex:0 0 7px;height:7px;border-radius:50%;background:rgba(148,163,184,.35);transition:background .15s,box-shadow .15s}.company-item:hover{background:rgba(56,189,248,.06);border-color:rgba(56,189,248,.15);color:rgba(219,230,245,.9)}.company-item input[type=checkbox]{position:absolute;opacity:0;width:0;height:0;pointer-events:none}.company-item:has(input:checked){background:linear-gradient(90deg,rgba(6,182,212,.15),rgba(56,189,248,.08));border-color:rgba(6,182,212,.35);color:#e2f0ff}.company-item:has(input:checked)::before{background:rgba(6,182,212,.9);box-shadow:0 0 6px rgba(6,182,212,.6)}
.grid{display:grid;gap:16px}.kpis{grid-template-columns:repeat(auto-fit,minmax(240px,1fr));margin-bottom:18px}.card{padding:20px;overflow:hidden}.card:before{content:'';position:absolute;inset:0;background:linear-gradient(135deg,rgba(255,255,255,.04),transparent 30%,transparent 70%,rgba(255,255,255,.02));pointer-events:none}.card.kpi:after{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--cyan),var(--blue),var(--violet))}.label{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:12px;font-weight:700}.metric{font-size:clamp(32px,3vw,46px);font-weight:900;letter-spacing:-.04em;line-height:1.02;word-break:break-word;overflow-wrap:anywhere}.hint{margin-top:8px;font-size:13px;color:var(--muted-2);line-height:1.5}.submetric{margin-top:12px;font-size:12px;color:var(--muted)}
.row,.row2{display:grid;gap:16px;margin-bottom:16px}.row{grid-template-columns:1.25fr .95fr}.row2{grid-template-columns:.9fr 1.3fr}.section-title{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:16px}.section-title h2{margin:0;font-size:18px;letter-spacing:-.02em}.section-title span{font-size:12px;color:var(--muted)}.bars{display:grid;gap:14px}.bar-row{display:grid;grid-template-columns:170px 1fr 46px;gap:12px;align-items:center}.bar-label{font-size:13px;color:#e2e8f0;font-weight:600}.bar{height:14px;background:linear-gradient(180deg,#0d1930,#152647);border:1px solid rgba(255,255,255,.04);border-radius:999px;overflow:hidden}.fill{height:100%;border-radius:999px;background:linear-gradient(90deg,var(--cyan) 0%,var(--blue) 55%,var(--violet) 100%)}.split{display:grid;grid-template-columns:1fr 1fr;gap:12px}.mini,.callout{padding:16px;border-radius:18px;background:linear-gradient(180deg,#132440,#0f1d34);border:1px solid var(--line-soft)}.chips{display:flex;flex-wrap:wrap;gap:10px}.chip{padding:11px 13px;border-radius:14px;background:linear-gradient(180deg,#12213b,#0e1b30);border:1px solid var(--line-soft);font-size:13px;color:#d7e3f4;min-height:44px;display:flex;align-items:center}.table-wrap{overflow:auto;border-radius:18px;border:1px solid rgba(255,255,255,.04)}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:12px 10px;border-bottom:1px solid rgba(148,163,184,.10);text-align:left;vertical-align:top}th{color:var(--muted);font-weight:700;font-size:12px;text-transform:uppercase;letter-spacing:.06em;background:rgba(255,255,255,.02)}tbody tr:hover{background:rgba(255,255,255,.025)}.empty{padding:22px;color:var(--muted-2)}th.result-col,td.result-col{background:rgba(34,197,94,.10)!important;color:#dcfce7;font-weight:800}th.result-col{box-shadow:inset 0 -2px 0 rgba(34,197,94,.35)}
@media(max-width:1050px){.row,.row2,.filter-bar{grid-template-columns:1fr}.split{grid-template-columns:1fr}.hero h1{font-size:32px}.hero-side{flex:1 1 100%}}
</style>
</head>
<body>
<div class="container">
 <div class="hero">
 <div class="hero-main">
 <h1>Dashboard executivo, Notas fiscais sem pedido</h1>
 <p>Versão funcional com filtros por empresa e data para leitura da esteira operacional de notas fiscais sem pedido, com foco em solicitação, pedido/medição, lançamento no Mega e gargalos reais da base.</p>
 <div class="hero-meta">
 <div class="hero-pill">Filtro por empresa</div>
 <div class="hero-pill">Filtro por período</div>
 <div class="hero-pill">Leitura somente consulta</div>
 </div>
 </div>
 <div class="hero-side">
 <div><div class="stamp-title">Última atualização</div><div class="stamp-value">${esc(data.generatedAt)}</div></div>
 <div class="stamp-sub">Fonte principal: <b>${esc(data.spreadsheetTitle)}</b><br>Total lido na base: <b>${data.docs.length}</b> notas.</div>
 </div>
 </div>

 <div class="card" style="margin-bottom:18px">
 <div class="section-title"><h2>Filtros da dashboard</h2><span>Use para olhar empresas e intervalo de datas</span></div>
 <div class="filter-bar" style="grid-template-columns:1.2fr 1fr 1fr auto;align-items:start">
 <div class="filter-group"><label>Empresas</label><div id="empresaCheckboxes" class="company-list" style="padding:2px 0;max-height:150px;overflow:auto"></div></div>
 <div class="filter-group"><label for="dataInicialFiltro">Data inicial</label><input id="dataInicialFiltro" type="text" inputmode="numeric" placeholder="dd-mm-yyyy" maxlength="10" autocomplete="off" style="font-variant-numeric:tabular-nums;letter-spacing:.04em" /></div>
 <div class="filter-group"><label for="dataFinalFiltro">Data final</label><input id="dataFinalFiltro" type="text" inputmode="numeric" placeholder="dd-mm-yyyy" maxlength="10" autocomplete="off" style="font-variant-numeric:tabular-nums;letter-spacing:.04em" /></div>
 <div><button id="limparFiltros" class="btn">Limpar filtros</button></div>
 </div>
 <div class="filter-note" id="resumoFiltros" style="margin-top:4px;padding:9px 12px;font-size:12px">Mostrando visão consolidada. O período só é aplicado quando data inicial e data final forem preenchidas.</div>
 </div>

 <div class="grid kpis">
 <div class="card kpi"><div class="label">Total de notas</div><div class="metric" id="kpiTotal">-</div><div class="hint">Volume total no recorte atual.</div><div class="submetric">Visão consolidada da base filtrada.</div></div>
 <div class="card kpi"><div class="label">Sem solicitação</div><div class="metric" id="kpiSemSolicitacao">-</div><div class="hint">Notas sem avanço inicial no fluxo.</div><div class="submetric">Gargalo primário da esteira.</div></div>
 <div class="card kpi" style="border-color:rgba(239,68,68,.35);box-shadow:0 20px 50px rgba(239,68,68,.10)"><div class="label" style="color:#fca5a5">Erro na solicitação</div><div class="metric" id="kpiErroSolicitacao">-</div><div class="hint">Qualquer vermelho na coluna Solicitação.</div><div class="submetric">Sinalização visual de erro na etapa.</div></div>
 <div class="card kpi"><div class="label">Com solicitação, sem pedido/medição</div><div class="metric" id="kpiSemPedido">-</div><div class="hint">Já entraram no fluxo, mas ainda sem pedido/medição.</div><div class="submetric">Backlog intermediário da operação.</div></div>
 <div class="card kpi" style="border-color:rgba(239,68,68,.35);box-shadow:0 20px 50px rgba(239,68,68,.10)"><div class="label" style="color:#fca5a5">Erro no pedido/medição</div><div class="metric" id="kpiErroPedido">-</div><div class="hint">Qualquer vermelho na coluna Pedido/Medição.</div><div class="submetric">Sinalização visual de erro na etapa.</div></div>
 <div class="card kpi"><div class="label">Pode lançar</div><div class="metric" id="kpiPodeLancar">-</div><div class="hint">Solicitação e pedido/medição preenchidos, sem lançamento NF ou fluxo LIVRE pronto.</div><div class="submetric">Prontos para lançamento no Mega.</div></div>
 <div class="card kpi"><div class="label">Lançado no Mega</div><div class="metric" id="kpiLancado">-</div><div class="hint">Documentos com a coluna Lançamento NF preenchida.</div><div class="submetric">Etapa concluída dentro do ERP.</div></div>
 <div class="card kpi"><div class="label">Valor total monitorado</div><div class="metric" id="kpiValor">-</div><div class="hint">Soma financeira do recorte atual.</div><div class="submetric">Mede exposição do universo filtrado.</div></div>
 </div>

 <div class="row">
 <div class="card"><div class="section-title"><h2>Volume por obra</h2><span>Concentração de notas no recorte</span></div><div class="bars" id="volumePorObra"></div></div>
 <div class="card">
 <div class="section-title"><h2>Gargalos prioritários</h2><span>Onde a operação está travando</span></div>
 <div class="split">
 <div class="mini"><div class="label">Gargalo principal</div><div class="metric" style="font-size:28px" id="gargaloPrincipal">-</div><div class="hint">Etapa com maior volume no recorte</div></div>
 <div class="mini"><div class="label">Empresa crítica</div><div class="metric" style="font-size:28px" id="empresaCritica">-</div><div class="hint">Maior concentração por aba</div></div>
 </div>
 <div style="height:12px"></div>
 <div class="split">
 <div class="mini"><div class="label">Obra com maior atraso</div><div class="metric" style="font-size:28px" id="obraCritica">-</div><div class="hint">Maior volume de pendências antigas no recorte</div></div>
 <div class="mini"><div class="label">Maior valor parado</div><div class="metric" style="font-size:28px" id="valorParado">-</div><div class="hint">Soma das etapas ainda não lançadas no Mega</div></div>
 </div>
 <div style="height:12px"></div>
 <div class="chips"><span class="chip"><b>Empresa</b>: baseado na aba de origem</span><span class="chip"><b>Data</b>: filtro por emissão</span></div>
 </div>
 </div>

 <div class="row2">
 <div class="card"><div class="section-title"><h2>Fluxo de entrada por dia</h2><span>Distribuição dentro do recorte atual</span></div><div class="bars" id="fluxoEntrada"></div></div>
 <div class="card"><div class="section-title"><h2>Ranking de obras críticas</h2><span>Obras com mais notas no recorte</span></div><div class="table-wrap"><table><thead><tr><th>Obra</th><th>Pendentes</th><th>Sem solicitação</th><th>Erro solicitação</th><th>Sem pedido</th><th>Erro pedido</th><th>Pode lançar</th><th class="result-col">Lançado</th></tr></thead><tbody id="rankingBody"></tbody></table></div></div>
 </div>

 <div class="row">
 <div class="card"><div class="section-title"><h2>Matriz de contato</h2><span>Responsáveis do recorte selecionado</span></div><div class="table-wrap"><table><thead><tr><th>Obra</th><th>Engenheiro</th><th>Telefone</th><th>E-mail</th></tr></thead><tbody id="contatoBody"></tbody></table></div></div>
 <div class="card"><div class="section-title"><h2>Notas prontas para lançar</h2><span>Prioridade operacional</span></div><div class="table-wrap"><table><thead><tr><th>Empresa</th><th>Obra</th><th>Nota</th><th>Fornecedor</th><th>Pedido/Medição</th><th>Valor</th></tr></thead><tbody id="podeLancarBody"></tbody></table></div></div>
 </div>
</div>
<script>
const RAW_DOCS = ${JSON.stringify(data.docs)};
const fmtMoney = v => new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(v||0);
const toDateInput = s => { if(!s) return ''; const p=s.split('/'); return p.length===3 ? [p[2], p[1].padStart(2,'0'), p[0].padStart(2,'0')].join('-') : ''; };
const fromIsoToBr = iso => {
 if (!iso) return '';
 const parts = String(iso).split('-');
 if (parts.length !== 3 || parts[0].length !== 4 || parts[1].length !== 2 || parts[2].length !== 2) return '';
 return parts[2] + '/' + parts[1] + '/' + parts[0];
};
function toIso(v){ if(!v||v.length!==10) return ''; const p=v.split('-'); return (p.length===3&&p[0].length===2&&p[1].length===2&&p[2].length===4) ? p[2]+'-'+p[1]+'-'+p[0] : ''; }
function normalizeDateInput(el){ const digits = String(el.value || '').replace(/\D/g,'').slice(0,8); if (!digits) { el.value = ''; return; } if (digits.length <= 2) { el.value = digits; return; } if (digits.length <= 4) { el.value = digits.slice(0,2) + '-' + digits.slice(2); return; } el.value = digits.slice(0,2) + '-' + digits.slice(2,4) + '-' + digits.slice(4,8); }
const byCount = (arr, keyFn) => { const m = new Map(); arr.forEach(r => { const k = keyFn(r); m.set(k,(m.get(k)||0)+1); }); return Array.from(m.entries()).map(([key,value])=>({key,value})); };
const empresaBox = document.getElementById('empresaCheckboxes'); const dataInicialFiltro = document.getElementById('dataInicialFiltro'); const dataFinalFiltro = document.getElementById('dataFinalFiltro');
const setHtml = (id, html) => document.getElementById(id).innerHTML = html;
const setText = (id, text) => document.getElementById(id).textContent = text;

function uniqueEmpresas(){ const empresas = Array.from(new Set(RAW_DOCS.map(x => x.aba).filter(Boolean))); const prioridade = ['STEIN','ITA']; return empresas.sort((a,b)=>{ const ia = prioridade.indexOf(a); const ib = prioridade.indexOf(b); if (ia !== -1 || ib !== -1) { if (ia === -1) return 1; if (ib === -1) return -1; return ia - ib; } return a.localeCompare(b,'pt-BR'); }); }
function isValidOperationalObra(v, empresa){ const s = String(v || '').trim(); if(!s) return false; const x = s.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); const empresaSlug = String(empresa||'').trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); const bloqueios = ['sem obra','sem info','verificando','escritorio','stein','ita','vertikal','costa esmeralda','els participacoes','hasa 13','stein e bertemes','stein litoral','stein alameda','cosmopolitan']; if(bloqueios.includes(x)) return false; if(empresaSlug && x === empresaSlug) return false; return true; }
function selectedEmpresas(){ return Array.from(document.querySelectorAll('input[name="empresaFiltro"]:checked')).map(x => x.value); }
function getPeriodo(){ const dtIniRaw=dataInicialFiltro.value; const dtFimRaw=dataFinalFiltro.value; let dtIni=toIso(dtIniRaw); let dtFim=toIso(dtFimRaw); const parcial=(!!dtIniRaw&&!dtFimRaw)||(!dtIniRaw&&!!dtFimRaw); if(parcial) return {usar:false,parcial:true,dtIni:dtIniRaw,dtFim:dtFimRaw}; if(!(dtIni&&dtFim)) return {usar:false,parcial:false,dtIni:'',dtFim:''}; if(dtIni>dtFim){const tmp=dtIni;dtIni=dtFim;dtFim=tmp;} return {usar:true,parcial:false,dtIni,dtFim}; }
function filteredDocs(){ const empresas = selectedEmpresas(); const periodo = getPeriodo(); return RAW_DOCS.filter(r => { const d = r.dataEmissaoIso || ''; return (empresas.length===0 || empresas.includes(r.aba)) && (!periodo.usar || (d && d >= periodo.dtIni && d <= periodo.dtFim)); }); }
function emptyRow(colspan, txt){ return '<tr><td colspan="' + colspan + '" class="empty">' + txt + '</td></tr>'; }
function renderBars(targetId, items, emptyText){ if(!items.length){ setHtml(targetId, '<div class="empty">' + emptyText + '</div>'); return; } const max = Math.max(...items.map(x=>x.value),1); setHtml(targetId, items.map(x => '<div class="bar-row"><div class="bar-label">' + x.key + '</div><div class="bar"><div class="fill" style="width:' + ((x.value/max)*100) + '%"></div></div><div>' + x.value + '</div></div>').join('')); }

function render(){
 const periodo = getPeriodo();
 const docs = filteredDocs();
 const erroSolicitacao = docs.filter(r => r.erroSolicitacao);
 const erroPedido = docs.filter(r => r.erroPedidoMedicao);
 const semSolicitacao = docs.filter(r => r.statusOperacional === 'Sem solicitação');
 const semPedido = docs.filter(r => r.statusOperacional === 'Com solicitação, sem pedido/medição');
 const podeLancar = docs.filter(r => r.statusOperacional === 'Pode lançar');
 const lancado = docs.filter(r => r.statusOperacional === 'Lançado no Mega');

 setText('kpiTotal', docs.length);
 setText('kpiSemSolicitacao', semSolicitacao.length);
 setText('kpiErroSolicitacao', erroSolicitacao.length);
 setText('kpiSemPedido', semPedido.length);
 setText('kpiErroPedido', erroPedido.length);
 setText('kpiPodeLancar', podeLancar.length);
 setText('kpiLancado', lancado.length);
 setText('kpiValor', fmtMoney(docs.reduce((s,r)=>s+(r.valor||0),0)));

 const statusRank = [
 {key:'Sem solicitação', value:semSolicitacao.length},
 {key:'Com solicitação, sem pedido/medição', value:semPedido.length},
 {key:'Pode lançar', value:podeLancar.length},
 {key:'Lançado no Mega', value:lancado.length}
 ].sort((a,b)=>b.value-a.value);
 setText('gargaloPrincipal', statusRank[0]?.key || '-');

 const empresaRank = byCount(docs, r => r.aba).sort((a,b)=>b.value-a.value);
 setText('empresaCritica', empresaRank[0]?.key || '-');

 const docsComObraValida = docs.filter(r => isValidOperationalObra((r.obra && r.obra.trim()) ? r.obra : '', r.aba || ''));
 const obrasRank = byCount(docsComObraValida, r => r.obra).sort((a,b)=>b.value-a.value);
 const atrasoPorObra = new Map();
 docsComObraValida.filter(r => r.statusOperacional !== 'Lançado no Mega').forEach(r => {
 const key = r.obra;
 const item = atrasoPorObra.get(key) || { obra:key, atrasoTotal:0, qtd:0 };
 item.atrasoTotal += Math.max(0, Number(r.daysLate || 0));
 item.qtd += 1;
 atrasoPorObra.set(key, item);
 });
 const obraMaiorAtraso = Array.from(atrasoPorObra.values()).sort((a,b)=>b.atrasoTotal-a.atrasoTotal || b.qtd-a.qtd)[0];
 setText('obraCritica', obraMaiorAtraso?.obra || '-');
 const valorParado = docs.filter(r => r.statusOperacional !== 'Lançado no Mega').reduce((s,r)=>s+(r.valor||0),0);
 setText('valorParado', fmtMoney(valorParado));

 const obras = obrasRank.slice(0,8);
 renderBars('volumePorObra', obras, 'Nenhuma nota encontrada para o recorte atual.');

 const fluxo = byCount(docs.filter(r => r.dataEmissaoIso), r => r.dataEmissaoIso).sort((a,b)=>a.key.localeCompare(b.key)).slice(-10).map(x => ({ key: fromIsoToBr(x.key), value: x.value }));
 renderBars('fluxoEntrada', fluxo, 'Sem documentos com data válida nesse recorte.');

 const rankingMap = new Map();
 docsComObraValida.forEach(r => {
 const key = r.obra;
 const item = rankingMap.get(key) || {obra:key,totalAtivo:0,semSolicitacao:0,erroSolicitacao:0,semPedido:0,erroPedido:0,podeLancar:0,lancado:0};
 if(r.statusOperacional !== 'Lançado no Mega') item.totalAtivo++;
 if(r.statusOperacional === 'Sem solicitação') item.semSolicitacao++;
 if(r.erroSolicitacao) item.erroSolicitacao++;
 if(r.statusOperacional === 'Com solicitação, sem pedido/medição') item.semPedido++;
 if(r.erroPedidoMedicao) item.erroPedido++;
 if(r.statusOperacional === 'Pode lançar') item.podeLancar++;
 if(r.statusOperacional === 'Lançado no Mega') item.lancado++;
 rankingMap.set(key,item);
 });
 const ranking = Array.from(rankingMap.values()).sort((a,b)=>b.totalAtivo-a.totalAtivo || b.lancado-a.lancado).slice(0,10);
 setHtml('rankingBody', ranking.length ? ranking.map(r => '<tr><td>' + r.obra + '</td><td>' + r.totalAtivo + '</td><td>' + r.semSolicitacao + '</td><td>' + r.erroSolicitacao + '</td><td>' + r.semPedido + '</td><td>' + r.erroPedido + '</td><td>' + r.podeLancar + '</td><td class="result-col">' + r.lancado + '</td></tr>').join('') : emptyRow(8,'Sem dados para esse filtro.'));

 const contatoMap = new Map();
 docs.forEach(r => {
 const key = (r.obra && r.obra.trim()) ? r.obra : '';
 if (!isValidOperationalObra(key, r.aba || '')) return;
 if (!contatoMap.has(key)) contatoMap.set(key, {obra:key,engenheiro:r.engenheiro||'-',fone:r.contatoFone||'-',email:r.contatoEmail||'-'});
 });
 const contatos = Array.from(contatoMap.values()).filter(c => c.engenheiro !== '-' || c.fone !== '-' || c.email !== '-').slice(0,12);
 setHtml('contatoBody', contatos.length ? contatos.map(c => '<tr><td>' + c.obra + '</td><td>' + c.engenheiro + '</td><td>' + c.fone + '</td><td>' + c.email + '</td></tr>').join('') : emptyRow(4,'Sem contatos para esse filtro.'));

 const podeLancarVisivel = podeLancar.filter(r => isValidOperationalObra(r.obra || '', r.aba || ''));
 setHtml('podeLancarBody', podeLancarVisivel.length ? podeLancarVisivel.slice(0,15).map(r => '<tr><td>' + r.aba + '</td><td>' + (r.obra || '-') + '</td><td>' + r.notaFiscal + '</td><td>' + r.razaoSocial + '</td><td>' + (r.pedidoMedicao || '-') + '</td><td>' + fmtMoney(r.valor) + '</td></tr>').join('') : emptyRow(6,'Nenhuma nota pronta para lançar nesse recorte.'));

 const empresas = selectedEmpresas(); const empresaTxt = empresas.length ? empresas.join(', ') : 'todas as empresas';
 const dtIniMissing = !dataInicialFiltro.value && !!dataFinalFiltro.value;
 const dtFimMissing = !!dataInicialFiltro.value && !dataFinalFiltro.value;
 dataInicialFiltro.style.borderColor = dtIniMissing ? 'rgba(245,158,11,.85)' : '';
 dataInicialFiltro.style.boxShadow = dtIniMissing ? '0 0 0 3px rgba(245,158,11,.18)' : '';
 dataFinalFiltro.style.borderColor = dtFimMissing ? 'rgba(245,158,11,.85)' : '';
 dataFinalFiltro.style.boxShadow = dtFimMissing ? '0 0 0 3px rgba(245,158,11,.18)' : '';
 let dataTxt = 'sem período aplicado';
 if (periodo.parcial) dataTxt = '⚠️ preencha a data ' + (dtIniMissing ? 'inicial' : 'final') + ' para filtrar por período';
 else if (periodo.usar) dataTxt = periodo.dtIni + ' até ' + periodo.dtFim;
 setText('resumoFiltros', 'Mostrando ' + docs.length + ' notas no recorte de ' + empresaTxt + (periodo.parcial ? '. ' + dataTxt : ' e período ' + dataTxt) + '.');
}

empresaBox.innerHTML = ['<label class="company-item"><input type="checkbox" id="selecionarTodasEmpresas"><span>Todas as empresas</span></label>'].concat(uniqueEmpresas().map(e => '<label class="company-item"><input type="checkbox" name="empresaFiltro" value="' + e + '"><span>' + e + '</span></label>')).join('');
document.getElementById('selecionarTodasEmpresas').addEventListener('change', function(){ document.querySelectorAll('input[name="empresaFiltro"]').forEach(x => x.checked = this.checked); render(); });
document.querySelectorAll('input[name="empresaFiltro"]').forEach(el => el.addEventListener('change', () => { const all = Array.from(document.querySelectorAll('input[name="empresaFiltro"]')); document.getElementById('selecionarTodasEmpresas').checked = all.length > 0 && all.every(x => x.checked); render(); }));
dataInicialFiltro.addEventListener('input', render);
dataFinalFiltro.addEventListener('input', render);
dataInicialFiltro.addEventListener('blur', function(){ normalizeDateInput(this); render(); });
dataFinalFiltro.addEventListener('blur', function(){ normalizeDateInput(this); render(); });
document.getElementById('limparFiltros').addEventListener('click', ()=>{ document.querySelectorAll('input[name="empresaFiltro"]').forEach(x => x.checked=false); document.getElementById('selecionarTodasEmpresas').checked = false; dataInicialFiltro.value=''; dataFinalFiltro.value=''; render(); });
render();
</script>
</body>
</html>`;
}

async function main() {
 const contactData = loadContacts();
 const docs = await loadDocs(contactData);
 const html = buildHtml({
  spreadsheetTitle: SPREADSHEET_TITLE,
  generatedAt: new Date().toLocaleString('pt-BR'),
  docs
 });

 fs.writeFileSync(OUTPUT, html, 'utf8');
 console.log(`Dashboard gerada em: ${OUTPUT}`);
 console.log(`Total de notas: ${docs.length}`);
 console.log(`Sem solicitação: ${docs.filter(r => r.statusOperacional === 'Sem solicitação').length}`);
 console.log(`Erro na solicitação: ${docs.filter(r => r.erroSolicitacao).length}`);
 console.log(`Com solicitação, sem pedido/medição: ${docs.filter(r => r.statusOperacional === 'Com solicitação, sem pedido/medição').length}`);
 console.log(`Erro no pedido/medição: ${docs.filter(r => r.erroPedidoMedicao).length}`);
 console.log(`Pode lançar: ${docs.filter(r => r.statusOperacional === 'Pode lançar').length}`);
 console.log(`Lançado no Mega: ${docs.filter(r => r.statusOperacional === 'Lançado no Mega').length}`);
}

main().catch(err => {
 console.error(err);
 process.exit(1);
});
