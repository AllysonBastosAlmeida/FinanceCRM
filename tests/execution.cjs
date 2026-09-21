const { readFileSync } = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const { test } = require('node:test');

// Load the actual spreadsheet rules without starting Microsoft authentication.
const source = readFileSync('src/FinanceCRM.jsx', 'utf8');
const context = vm.createContext({});
vm.runInContext(
  source.slice(source.indexOf('const ID_KEYS'), source.indexOf('// ======== CLIENTE MICROSOFT')) +
  source.slice(source.indexOf('function toDate'), source.indexOf('export default function FinanceCRM')),
  context
);

test('marker identifies approved work without a payment date', () => {
  assert.equal(context.isInExecution({ orcamento_status: 'Aprovado', Data_Criação: '  CRIAR   NF  ' }), true);
  assert.equal(context.isInExecution({ orcamento_status: 'Aprovado', data_criacao: 'Criar NF', status: 'Não pago' }), true);
});
test('scheduled and paid rows cannot also be in execution', () => {
  for (const extra of [
    { data_pagamento_q: 46300 },
    { data_pagamento: '20/10/2026' },
    { 'Data de Pagamento': '2026-10-20' },
    { status: 'Pago' },
    { status_raw: 'Recebido' },
  ]) assert.equal(context.isInExecution({ orcamento_status: 'Aprovado', data_criacao: 'Criar NF', ...extra }), false);
});
test('dates, empty cells and unrelated instructions are not execution markers', () => {
  for (const value of ['', null, '20/09/2026', 46300, 'Não criar NF']) {
    assert.equal(context.isInExecution({ orcamento_status: 'Aprovado', data_criacao: value }), false);
  }
});

test('execution requires an explicitly approved budget', () => {
  for (const status of ['Reprovado', 'Em analise', '', undefined]) {
    assert.equal(context.isInExecution({ orcamento_status: status, data_criacao: 'Criar NF' }), false);
  }
});

test('Status is authoritative and deleting Approval preserves all financial fields', async () => {
  const headers = ['ID', 'PO', 'Data', 'Cliente', 'Assunto', 'Extra', 'Valor', 'Status', 'Outro',
    'Aprovação', 'K', 'L', 'M', 'NF', 'Data_Criação', 'Prazo_Dias', 'Data de Pagamento', 'Situação'];
  const row = [1, 'PO-1', '01/09/2026', 'Cliente teste', 'Instalação', '', 1200, 'Reprovado', '',
    'Aprovado', '', '', '', 'NF-1', '05/09/2026', 30, '05/10/2026', 'Pendente'];
  const importRows = async (values) => {
    const ctx = vm.createContext({ console, window: {},
      getGraphClient: async () => ({ api: () => ({ get: async () => ({ values }) }) }),
    });
    vm.runInContext(
      source.slice(source.indexOf('const ID_KEYS'), source.indexOf('// ======== CLIENTE MICROSOFT')) +
      source.slice(source.indexOf('const SHAREPOINT_SITE_ID'), source.indexOf('export default function FinanceCRM')),
      ctx
    );
    return JSON.parse(JSON.stringify(await ctx.loadExcelAsRows()));
  };
  const [before] = await importRows([headers, row]);
  const [after] = await importRows([headers.filter((_, i) => i !== 9), row.filter((_, i) => i !== 9)]);
  assert.equal(before.orcamento_status, 'Reprovado');
  assert.equal(before.status, 'Pendente');
  for (const key of ['orcamento_status', 'status', 'nf', 'data_criacao', 'data_pagamento', 'data_pagamento_q', 'prazo_dias', 'valor']) {
    assert.equal(after[key], before[key], key);
  }
  const approved = [...row]; approved[7] = 'Aprovado'; approved[9] = 'Reprovado';
  approved[14] = 'Criar NF'; approved[16] = '';
  const [work] = await importRows([headers, approved]);
  assert.equal(work.orcamento_status, 'Aprovado');
  assert.equal(context.isInExecution(work), true);
  approved[7] = '';
  const [unknown] = await importRows([headers, approved]);
  assert.equal(unknown.orcamento_status, 'Nao informado');
  assert.equal(context.isInExecution(unknown), false);
});
test('header lookup supports accents, underscores and filled aliases', () => {
  const row = { Data_Criação: 'Criar NF', pagamento: '', 'Data de Pagamento': '20/10/2026' };
  assert.equal(context.pickFilled(row, ['data criacao']), 'Criar NF');
  assert.equal(context.pickFilled(row, ['pagamento', 'data_de_pagamento']), '20/10/2026');
});
test('execution remains visible across payment periods while scheduled rows obey dates', () => {
  const inProgress = { orcamento_status: 'Aprovado', data_criacao: 'Criar NF' };
  const scheduled = { data_pagamento: '2026-10-20' };
  Object.assign(context, {
    rows: [inProgress, scheduled], q: '', cliente: 'Todos', status: 'Todos', activePage: 'overview',
    summarizeClients: () => ({}), parseValor: Number,
    orcamentoStatusFilter: 'Todos', executionFilter: 'Todos', ano: '2025', mes: 'Todos', quickRange: 'Todos',
    useMemo: fn => fn(), matchesCommonFilters: () => true,
    matchesDateFilters: () => false, getRowDate: () => null,
    getEffectivePaymentDate: row => context.toDate(row.data_pagamento),
  });
  vm.runInContext(source.slice(source.indexOf('  const filtered = useMemo'), source.indexOf('  const pendingInvoiceReceivable')), context);
  assert.equal(vm.runInContext('filtered.length', context), 1);
  assert.equal(vm.runInContext('pendingInvoiceRows.length', context), 1);
  assert.equal(vm.runInContext('paymentFiltered.length', context), 0);
});
