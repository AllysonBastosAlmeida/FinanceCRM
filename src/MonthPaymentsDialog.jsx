import { useEffect, useRef } from "react";

export default function MonthPaymentsDialog({ month, rows, getDate, formatMoney, getStatus, onClose }) {
  const dialogRef = useRef(null);
  useEffect(() => {
    const dialog = dialogRef.current;
    const previousFocus = document.activeElement;
    const overflow = document.body.style.overflow;
    dialog.showModal();
    document.body.style.overflow = "hidden";
    return () => {
      dialog.close();
      document.body.style.overflow = overflow;
      previousFocus?.focus();
    };
  }, []);

  const payments = rows.map(row => ({ row, date: getDate(row) }))
    .filter(({ date }) => date && `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}` === month.key)
    .sort((a, b) => a.date - b.date);

  return (
    <dialog ref={dialogRef} className="month-payments-dialog" aria-labelledby="month-payments-title"
      onCancel={onClose} onClick={event => { if (event.target === event.currentTarget) onClose(); }}>
      <div className="month-payments-content">
        <div className="client-section-header">
          <div>
            <div className="eyebrow">Pagamentos do mês</div>
            <h2 id="month-payments-title">{month.monthLabel}</h2>
            <div className="kpi-sub">{payments.length} registro(s) · Conforme os filtros selecionados</div>
          </div>
          <button type="button" className="theme-btn" onClick={onClose} autoFocus>Fechar</button>
        </div>
        <div className="month-payments-totals">
          <div><span>Total</span><strong>{formatMoney(month.total)}</strong></div>
          <div><span>Pago</span><strong>{formatMoney(month.pago)}</strong></div>
          <div><span>Em aberto</span><strong>{formatMoney(month.aberto)}</strong></div>
        </div>
        <div className="table-wrapper">
          <table className="client-table">
            <thead><tr><th>Cliente</th><th>Serviço</th><th>PO / NF</th><th>Pagamento</th><th>Valor</th><th>Status</th></tr></thead>
            <tbody>
              {payments.map(({ row, date }, index) => (
                <tr key={`${row.id || row.po || "row"}-${index}`}>
                  <td>{row.cliente || "—"}</td><td>{row.assunto || row.servico || "—"}</td>
                  <td>{row.po || "—"} / {row.nf || "—"}</td>
                  <td className="client-money">{date.toLocaleDateString("pt-BR")}</td>
                  <td className="client-money">{formatMoney(row.valor)}</td><td>{getStatus(row)}</td>
                </tr>
              ))}
              {!payments.length && <tr><td colSpan={6} className="client-empty">Nenhum pagamento neste mês com os filtros selecionados.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </dialog>
  );
}
