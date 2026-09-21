export function summarizeClients(rows, { valueOf, isPaid, isInExecution, paymentDate }) {
  const clients = new Map();
  const summary = { total: 0, paid: 0, execution: 0, scheduled: 0, unscheduled: 0, count: rows.length };
  for (const row of rows) {
    const name = String(row.cliente || '').trim() || 'Cliente não informado';
    if (!clients.has(name)) clients.set(name, { name, total: 0, paid: 0, count: 0 });
    const client = clients.get(name);
    const value = valueOf(row.valor);
    summary.total += value;
    client.total += value;
    client.count++;
    if (isPaid(row)) { summary.paid += value; client.paid += value; }
    else if (isInExecution(row)) summary.execution += value;
    else if (paymentDate(row)) summary.scheduled += value;
    else summary.unscheduled += value;
  }
  return { ...summary, clients: [...clients.values()].sort((a, b) => b.total - a.total) };
}
