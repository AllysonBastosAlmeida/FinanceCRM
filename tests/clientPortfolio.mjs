import { test } from 'node:test';
import assert from 'node:assert/strict';
import { summarizeClients } from '../src/clientPortfolio.mjs';
const rules = { valueOf: Number, isPaid: r => r.paid, isInExecution: r => r.execution, paymentDate: r => r.date };
test('portfolio reconciles approved amounts without counting paid work twice', () => {
  const result = summarizeClients([
    { cliente: 'A', valor: 100, paid: true, date: true },
    { cliente: 'A', valor: 200, execution: true },
    { cliente: 'B', valor: 400, date: true },
    { cliente: 'B', valor: 50 },
  ], rules);
  assert.equal(result.total, 750);
  assert.equal(result.paid, 100);
  assert.equal(result.execution, 200);
  assert.equal(result.scheduled, 400);
  assert.equal(result.unscheduled, 50);
  assert.equal(result.clients[0].name, 'B');
  assert.equal(result.clients[1].count, 2);
  assert.equal(result.total, result.paid + result.execution + result.scheduled + result.unscheduled);
});
test('empty filters produce zero totals and no clients', () => {
  const result = summarizeClients([], rules);
  assert.equal(result.total, 0);
  assert.equal(result.clients.length, 0);
});
