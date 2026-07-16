#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const html = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
function embedded(pattern, name) {
 const match = html.match(pattern);
 if (!match) throw new Error(`Bloco ${name} não encontrado no HTML`);
 return JSON.parse(match[1]);
}
const docs = embedded(/const RAW_DOCS = (\[.*?\]);\nconst DATA_QUALITY/s, 'RAW_DOCS');
const quality = embedded(/const DATA_QUALITY = (\{.*?\});/, 'DATA_QUALITY');
const statuses = ['Sem solicitação','Com solicitação, sem pedido/medição','Liberado para lançar','Lançado no Mega','Fora do fluxo'];
const assert = (ok, message) => { if (!ok) throw new Error(message); };
assert(docs.length === quality.sourceRows - quality.duplicateRowsRemoved, 'Total único não reconcilia com fonte menos duplicidades');
assert(docs.reduce((s,r)=>s+(r.ocorrenciasOrigem||1),0) === quality.sourceRows, 'Ocorrências de origem não reconciliam');
assert(statuses.reduce((s,status)=>s+docs.filter(r=>r.statusOperacional===status).length,0) === docs.length, 'Categorias operacionais não somam o total');
assert(quality.sheetStats.reduce((s,x)=>s+x.included,0) === quality.sourceRows, 'Totais por aba não somam a fonte');
assert(quality.sheetStats.find(x=>x.sheet==='COSMOPOLITAN')?.included === 6, 'COSMOPOLITAN não foi lida integralmente');
assert(quality.sheetStats.find(x=>x.sheet==='STEIN LITORAL')?.included === 18, 'STEIN LITORAL não foi lida integralmente');
assert(docs.filter(r=>r.aba==='STEIN LITORAL').every(r=>r.notaFiscal && r.razaoSocial && r.dataEmissaoIso && r.valor>0), 'STEIN LITORAL ainda contém registros corrompidos');
assert(docs.filter(r=>r.aba==='STEIN E BERTEMES').every(r=>r.dataEmissaoIso), 'STEIN E BERTEMES ainda contém data ausente');
assert(docs.every(r=>Number.isFinite(r.valor) && r.valor>=0), 'Há valor inválido');
assert(docs.filter(r=>!r.dataEmissaoIso).length === quality.missingDates, 'Contagem de datas ausentes diverge');
assert(!html.includes("replace(/D/g,''"), 'Regex de normalização de data foi corrompida no HTML');
console.log(JSON.stringify({
 sourceRows:quality.sourceRows,
 uniqueDocs:docs.length,
 duplicateRowsRemoved:quality.duplicateRowsRemoved,
 missingDates:quality.missingDates,
 totalValueUnique:Number(docs.reduce((s,r)=>s+r.valor,0).toFixed(2)),
 totalValueSource:Number(docs.reduce((s,r)=>s+r.valor*(r.ocorrenciasOrigem||1),0).toFixed(2)),
 redSolicitacao:docs.reduce((s,r)=>s+(r.errosSolicitacaoOrigem||0),0),
 redPedidoMedicao:docs.reduce((s,r)=>s+(r.errosPedidoOrigem||0),0),
 statuses:Object.fromEntries(statuses.map(status=>[status,docs.filter(r=>r.statusOperacional===status).length])),
 sheets:quality.sheetStats
},null,2));
