import { useEffect, useMemo, useState } from "react";
import logo from "./assets/logo.png";
import MonthPaymentsDialog from "./MonthPaymentsDialog";
import { summarizeClients } from "./clientPortfolio.mjs";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
  PieChart,
  Pie,
  Cell,
  Legend,
  AreaChart,
  Area,
} from "recharts";
import {
  InteractionRequiredAuthError,
  PublicClientApplication,
} from "@azure/msal-browser";
import { msalConfig } from "./authconfig";
import { Client } from "@microsoft/microsoft-graph-client";

// ======== FORMATADORES ========
const parseValor = (value) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return 0;
    const hasComma = trimmed.includes(",");
    let normalized = trimmed.replace(/[^\d.,-]/g, "");

    if (hasComma) {
      normalized = normalized.replace(/\./g, "").replace(",", ".");
    } else {
      const parts = normalized.split(".");
      if (parts.length > 2) {
        const last = parts.pop();
        normalized = `${parts.join("")}.${last}`;
      }
    }

    const num = Number(normalized);
    return Number.isFinite(num) ? num : 0;
  }

  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
};

const BRL = (n) =>
  parseValor(n).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });

const parseDate = (s) => {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
};

// ======== MSAL CONFIGURAÃ‡ÃƒO ========
const msalInstance = new PublicClientApplication(msalConfig);
const graphScopes = ["User.Read", "Files.Read", "Files.Read.All"];
const isMobileDevice =
  typeof navigator !== "undefined" &&
  /Mobi|Android|iPhone|iPad|Mobile/.test(navigator.userAgent || "");
const msalInitPromise = msalInstance.initialize();
let loginPromise = null;
const ID_KEYS = ["id"];
const PO_KEYS = ["po"];
const CLIENTE_KEYS = ["cliente", "empresa", "razao social"];
const ASSUNTO_KEYS = ["assunto", "descricao", "servico", "servico principal"];
const NF_KEYS = ["nf", "nota fiscal", "numero nf", "numero da nf"];
const VALOR_KEYS = ["valor", "valor total", "valor nf", "valor_orcamento"];
const PRAZO_KEYS = ["prazo dias", "prazo_dias", "qtd de dias", "qtd_dias"];
const ORCAMENTO_KEYS = ["status"];
const PAGAMENTO_STATUS_KEYS = ["situacao", "status pagamento", "status financeiro", "situacao pagamento"];
const DATA_BASE_KEYS = ["data base", "data_base", "data", "data referencia", "data_referencia"];
const PAG_KEYS = ["data de pagamento", "data_pagamento", "pagamento", "data pagamento"];
const EMI_KEYS = [
  "data criacao",
  "data_criacao",
  "data de emissao",
  "data de vencimento",
  "emissao",
  "vencimento",
  "data",
];
const SERV_KEYS = ["assunto", "descricao", "servico"];

const normalizeToken = (value) =>
  String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();

const normalizeHeader = (value) =>
  normalizeToken(value)
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");

const hasValue = (value) => value != null && String(value).trim() !== "";
const isCriarNfMarker = (value) => normalizeToken(value).replace(/\s+/g, " ") === "criar nf";

const toTitle = (value) =>
  String(value || "")
    .toLowerCase()
    .replace(/\b\w/g, (l) => l.toUpperCase());

function normalizePaymentStatus(value) {
  const text = normalizeToken(value);
  if (!text) return "Indefinido";
  if (text.includes("nao pago") || text.includes("nao recebido")) return "Pendente";
  if (text.includes("pago") || text.includes("recebido")) return "Pago";
  if (text.includes("atras")) return "Atrasado";
  if (
    text.includes("pend") ||
    text.includes("aberto") ||
    text.includes("nao pago") ||
    text.includes("a receber")
  ) {
    return "Pendente";
  }
  return toTitle(value);
}

function normalizeBudgetStatus(value) {
  const text = normalizeToken(value);
  if (!text) return "Nao informado";
  if (text.includes("reprov") || text.includes("negad") || text.includes("recus")) {
    return "Reprovado";
  }
  if (text.includes("aprov")) return "Aprovado";
  if (
    text.includes("aguard") ||
    text.includes("analise") ||
    text.includes("avali") ||
    text.includes("pendente")
  ) {
    return "Em analise";
  }
  return toTitle(value);
}

function statusBadgeClass(value) {
  const text = normalizeToken(value).replace(/\s+/g, "_");
  if (text.includes("pago")) return "pago";
  if (text.includes("pend")) return "pendente";
  if (text.includes("atras")) return "atrasado";
  if (text.includes("execu")) return "em_analise";
  if (text.includes("reprov")) return "reprovado";
  if (text.includes("aprov")) return "aprovado";
  if (text.includes("analise")) return "em_analise";
  return "";
}

// ======== CLIENTE MICROSOFT GRAPH ========
async function acquireToken(scopes) {
  await msalInitPromise;

  // Processa retornos de loginRedirect (especialmente em mobile)
  const redirectResult = await msalInstance.handleRedirectPromise();
  if (redirectResult?.account) {
    msalInstance.setActiveAccount(redirectResult.account);
    return { accessToken: redirectResult.accessToken, account: redirectResult.account };
  }

  const doLogin = async () => {
    if (loginPromise) return loginPromise;

    const p = (async () => {
      if (isMobileDevice) {
        await msalInstance.loginRedirect({
          scopes,
          prompt: "select_account",
        });
        return new Promise(() => {}); // fluxo continua no redirect
      }
      const loginResp = await msalInstance.loginPopup({
        scopes,
        prompt: "select_account",
      });
      msalInstance.setActiveAccount(loginResp.account);
      return loginResp;
    })();

    loginPromise = p;
    try {
      return await p;
    } finally {
      loginPromise = null;
    }
  };

  let account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) {
    const loginResp = await doLogin();
    account = loginResp?.account;
    return { accessToken: loginResp?.accessToken, account };
  }

  msalInstance.setActiveAccount(account);

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({ scopes, account });
    return { accessToken: tokenResponse.accessToken, account };
  } catch (err) {
    const needsInteraction =
      err instanceof InteractionRequiredAuthError ||
      err.errorCode === "login_required" ||
      err.errorCode === "consent_required" ||
      err.errorCode === "block_iframe_reload" ||
      err.errorCode === "no_tokens_found";

    if (err.errorCode === "interaction_in_progress" && loginPromise) {
      const loginResp = await loginPromise.catch(() => null);
      return { accessToken: loginResp?.accessToken, account: loginResp?.account };
    }

    if (needsInteraction) {
      const loginResp = await doLogin();
      return { accessToken: loginResp?.accessToken, account: loginResp?.account };
    }
    throw err;
  }
}

async function getGraphClient() {
  const { accessToken } = await acquireToken(graphScopes);

  return Client.init({
    authProvider: (done) => done(null, accessToken),
  });
}

// ======== PERFIL DO USUÃRIO MICROSOFT ========
async function getUserProfile(graphClient) {
  const profile = await graphClient.api("/me").get();
  console.log(
    "ðŸ‘¤ UsuÃ¡rio logado:",
    profile.displayName,
    profile.mail || profile.userPrincipalName
  );
  return {
    name: profile.displayName,
    email: profile.mail || profile.userPrincipalName,
  };
}

async function getUserPhoto(accessToken, setPhoto) {
  try {
    if (localStorage.getItem("userNoPhoto") === "1") return;
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/photo/$value",
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );
    if (!response.ok) {
      localStorage.setItem("userNoPhoto", "1");
      return;
    }
    const blob = await response.blob();
    const imageUrl = URL.createObjectURL(blob);
    setPhoto(imageUrl);
    localStorage.removeItem("userNoPhoto");
  } catch (error) {
    console.warn("Foto de perfil nÃ£o encontrada:", error?.message || error);
  }
}

const SHAREPOINT_SITE_ID = "3f31c6b1-8a58-4f1e-b65e-7e6455584500";
const SHAREPOINT_DRIVE_ID =
  "b!scYxP1iKHk-2Xn5kVVhFAGdG9X8BFZBGvt5w-aBi12Mo0YszQX9hSakmug3Ij2Qf";
const SHAREPOINT_FILE_ID = "01FWWAKIQQT5LZBQ5UGNBILG4UQ6VQGUJ3";
const SHAREPOINT_WORKSHEET = "Processos_Or\u00E7amentos";

async function loadExcelAsRows() {
  const client = await getGraphClient();

  try {
    const used = await client
      .api(
        `/sites/${SHAREPOINT_SITE_ID}/drives/${SHAREPOINT_DRIVE_ID}/items/${SHAREPOINT_FILE_ID}/workbook/worksheets('${SHAREPOINT_WORKSHEET}')/usedRange`
      )
      .get();

    const values = used.values || [];
    if (!values.length) return [];

    const headerCounters = new Map();
    const headers = (values[0] || []).map((header, index) => {
      const base = normalizeHeader(header) || `col_${index + 1}`;
      const seen = headerCounters.get(base) || 0;
      headerCounters.set(base, seen + 1);
      return seen ? `${base}_${seen + 1}` : base;
    });

    const rows = values
      .slice(1)
      .map((row) => {
        const raw = Object.fromEntries(headers.map((h, i) => [h, row[i]]));

        // Cabeçalhos permanecem válidos após excluir ou reordenar colunas.
        // Status (H) é orçamento; Situação é pagamento. Aprovação (J) é ignorada.
        const statusRaw = pickFilled(raw, PAGAMENTO_STATUS_KEYS);
        const orcamentoRaw = pickFilled(raw, ORCAMENTO_KEYS);
        const dataBaseRaw = pickFilled(raw, DATA_BASE_KEYS);
        const dataCriacaoRaw = pickFilled(raw, EMI_KEYS);
        const dataPagamentoRaw = pickFilled(raw, PAG_KEYS);
        const prazoRaw = pickFilled(raw, PRAZO_KEYS);
        const assuntoRaw = pickFilled(raw, ASSUNTO_KEYS);
        const valorRaw = pickFilled(raw, VALOR_KEYS);
        const nfRaw = pickFilled(raw, NF_KEYS);

        const canonical = {
          ...raw,
          id: pickFilled(raw, ID_KEYS),
          po: pickFilled(raw, PO_KEYS),
          cliente: pickFilled(raw, CLIENTE_KEYS),
          assunto: assuntoRaw,
          servico: pickFilled(raw, SERV_KEYS) ?? assuntoRaw,
          descricao: assuntoRaw,
          nf: nfRaw,
          "data base": dataBaseRaw,
          data_base: dataBaseRaw,
          "data criacao": dataCriacaoRaw,
          data_criacao: dataCriacaoRaw,
          "prazo dias": prazoRaw,
          prazo_dias: prazoRaw,
          "data de pagamento": dataPagamentoRaw,
          data_pagamento: dataPagamentoRaw,
          data_pagamento_q: dataPagamentoRaw,
          data: dataBaseRaw || dataPagamentoRaw || dataCriacaoRaw,
          valor: valorRaw,
          status_raw: statusRaw,
          status: normalizePaymentStatus(statusRaw),
          situacao: normalizePaymentStatus(statusRaw),
          orcamento_status_raw: orcamentoRaw,
          orcamento_status: normalizeBudgetStatus(orcamentoRaw),
          status_orcamento: normalizeBudgetStatus(orcamentoRaw),
        };

        return Object.fromEntries(
          Object.entries(canonical).map(([k, v]) => [k, typeof v === "string" ? v.trim() : v])
        );
      })
      .filter((r) =>
        [
          r.id,
          r.po,
          r.cliente,
          r.assunto,
          r.nf,
          r.valor,
          r.status_raw,
          r.orcamento_status_raw,
        ].some(hasValue)
      );

    console.log(`âœ… NFs carregada: ${rows.length} linhas`);
    window._rowsHeaders = headers;
    window._rowsDebug = rows;

    return rows;
  } catch (err) {
    console.error("âŒ Erro ao carregar NFs:", err);
    return [];
  }
}

// Converte serial do Excel ou string "dd/mm/yyyy" para Date
function toDate(val) {
  if (val == null || val === "") return null;

  if (typeof val === "number") {
    const excelBaseTime = Date.UTC(1900, 0, 1) - 2 * 86400000;
    const d = new Date(excelBaseTime + val * 86400000 + 43200000);
    return d;
  }

  if (typeof val === "string") {
    const s = val.trim();
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
      const [d, m, y] = s.split("/").map(Number);
      return new Date(y, m - 1, d);
    }

    const d2 = new Date(s);
    return isNaN(d2) ? null : d2;
  }

  const d3 = new Date(val);
  return isNaN(d3) ? null : d3;
}

// Reutiliza o índice de cabeçalhos das linhas imutáveis em todos os indicadores.
const headerIndexes = new WeakMap();
function getHeaderIndex(obj) {
  if (!headerIndexes.has(obj)) {
    headerIndexes.set(obj, Object.fromEntries(Object.keys(obj).map(key => [normalizeHeader(key), key])));
  }
  return headerIndexes.get(obj);
}
function pick(obj, keys) {
  const index = getHeaderIndex(obj);
  for (const key of keys) {
    const field = index[normalizeHeader(key)];
    if (field !== undefined) return obj[field];
  }
  return null;
}
function pickFilled(obj, keys) {
  const index = getHeaderIndex(obj);
  for (const key of keys) {
    const field = index[normalizeHeader(key)];
    if (field !== undefined && hasValue(obj[field])) return obj[field];
  }
  return null;
}
// Status confirma a aprovação; Data_Criação identifica o serviço em execução.
function isInExecution(row) {
  return normalizeBudgetStatus(row.orcamento_status_raw || row.orcamento_status) === "Aprovado" &&
    isCriarNfMarker(row.data_criacao ?? pickFilled(row, EMI_KEYS)) &&
    !toDate(row.data_pagamento_q) && !toDate(pickFilled(row, PAG_KEYS)) &&
    normalizePaymentStatus(row.status_raw || row.status) !== "Pago";
}

export default function FinanceCRM() {

  // MENU MOBILE
  const [showMobileMenu, setShowMobileMenu] = useState(false);
  const [activePage, setActivePage] = useState("overview");
  const navItems = [
    { id: "overview", label: "Visao Geral", desc: "KPIs e tendencia" },
    { id: "pendencias", label: "Pendencias", desc: "Atrasos e risco" },
    { id: "historico", label: "Historico", desc: "Tabela completa" },
    { id: "clientes", label: "Clientes", desc: "Ranking e carteira" },
  ];

  // ======== AUTH ========
  const [user, setUser] = useState(null);
  const [loadingAuth, setLoadingAuth] = useState(true);
  const [userPhoto, setUserPhoto] = useState(null);

  // ======== DATA ========
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [errMsg, setErrMsg] = useState("");

  useEffect(() => {
    async function fetchProfile() {
      try {
        setLoadingAuth(true);

        const storedUserName = localStorage.getItem("userName");
        const storedUserPhoto = localStorage.getItem("userPhoto");

        if (storedUserName) {
          setUser({ name: storedUserName });
        }

        if (storedUserPhoto) {
          setUserPhoto(storedUserPhoto);
        }

        const { accessToken: userToken, account } = await acquireToken(["User.Read"]);
        const client = await getGraphClient();
        const userInfo = await getUserProfile(client);
        setUser(userInfo);

        if (userToken) {
          await getUserPhoto(userToken, setUserPhoto);
        }

        localStorage.setItem("userName", userInfo.name);
        localStorage.setItem("userPhoto", userPhoto || "");
      } catch (err) {
        console.error("Erro ao autenticar ou carregar dados:", err);
        setErrMsg("Falha na autenticaÃ§Ã£o com Microsoft.");
      } finally {
        setLoadingAuth(false);
      }
    }
    fetchProfile();
  }, []);

  useEffect(() => {
    if (loadingAuth || !user) return;

    async function fetchData() {
      setLoading(true);
      setErrMsg("");
      const data = await loadExcelAsRows();
      setRows(data);
      if (!data.length) {
        setErrMsg("Nenhum registro encontrado na planilha atual ou sem permissao de leitura.");
      }
      setLoading(false);
    }

    fetchData();
  }, [user, loadingAuth]);

  // ======== FILTERS ========
  const [q, setQ] = useState("");
  const clientes = useMemo(
    () => ["Todos", ...Array.from(new Set(rows.map((r) => r.cliente).filter(hasValue))).sort()],
    [rows]
  );
  const [cliente, setCliente] = useState("Todos");
  const [status, setStatus] = useState("Todos");
  const [orcamentoStatusFilter, setOrcamentoStatusFilter] = useState("Todos");
  const [executionFilter, setExecutionFilter] = useState("Todos");

  const getRowBaseDate = (r) => toDate(pick(r, DATA_BASE_KEYS));
  const getRowPaymentDate = (r) => toDate(pick(r, PAG_KEYS));
  const getRowPaymentDateQ = (r) => toDate(r.data_pagamento_q);
  const getEffectivePaymentDate = (r) => getRowPaymentDateQ(r) || getRowPaymentDate(r);
  const getRowEmissionDate = (r) => toDate(pick(r, EMI_KEYS));
  const getRowDate = (r) => getRowBaseDate(r) || getRowPaymentDate(r) || getRowEmissionDate(r);

  const statusOptions = useMemo(() => {
    const opts = Array.from(
      new Set(rows.map((r) => normalizePaymentStatus(r.status_raw || r.status)).filter(hasValue))
    ).sort();
    return ["Todos", ...opts];
  }, [rows]);

  const orcamentoOptions = useMemo(() => {
    const opts = Array.from(
      new Set(
        rows
          .map((r) => normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status))
          .filter(hasValue)
      )
    ).sort();
    return ["Todos", ...opts];
  }, [rows]);

  const currentYear = new Date().getFullYear();
  const anos = useMemo(() => {
    const set = new Set(rows.map((r) => getRowDate(r)?.getFullYear()).filter(Boolean));
    set.add(currentYear);
    return ["Todos", ...Array.from(set).sort((a, b) => b - a)];
  }, [rows, currentYear]);
  const [ano, setAno] = useState("Todos");

  const meses = [
    "Todos",
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
  ];

  const [mes, setMes] = useState("Todos");
  const [monthlyStatusMode, setMonthlyStatusMode] = useState("Aprovado");

  const columnMap = useMemo(
    () => [
      { key: "id", label: "ID", style: { width: "50px" } },
      { key: "po", label: "PO" },
      { key: "cliente", label: "Cliente" },
      { key: "assunto", label: "Servico principal" },
      { key: "orcamento_status", label: "Orcamento", type: "status" },
      { key: "etapa", label: "Etapa", type: "status" },
      { key: "nf", label: "NF" },
      { key: "valor", label: "Valor", type: "currency" },
      { key: "data criacao", label: "Emissao", type: "date" },
      { key: "data de pagamento", label: "Pagamento", type: "date" },
      { key: "status", label: "Status financeiro", type: "status" },
    ],
    []
  );

  // ======== FILTROS RAPIDOS ========
  const [quickRange, setQuickRange] = useState("Todos");

  const matchesCommonFilters = (r) => {
    const statusNormalizado = normalizePaymentStatus(r.status_raw || r.status);
    const orcamentoNormalizado = normalizeBudgetStatus(
      r.orcamento_status_raw || r.orcamento_status
    );

    const matchTxt =
      q.trim() === "" ||
      `${r.cliente || ""} ${r.servico || ""} ${r.assunto || ""} ${r.nf || ""} ${r.po || ""}`
        .toLowerCase()
        .includes(q.toLowerCase());

    const matchCli = cliente === "Todos" || r.cliente === cliente;
    const matchSt = status === "Todos" || statusNormalizado === status;
    const matchOrc =
      activePage === "clientes"
        ? orcamentoNormalizado === "Aprovado"
        : orcamentoStatusFilter === "Todos" || orcamentoNormalizado === orcamentoStatusFilter;

    const matchExecution = executionFilter === "Todos" ||
      (executionFilter === "Em execução" ? isInExecution(r) : !!getEffectivePaymentDate(r));
    return matchTxt && matchCli && matchSt && matchOrc && matchExecution;
  };

  const matchesDateFilters = (d) => {
    let matchPeriodo = true;

    if (quickRange !== "Todos" && d) {
      const today = new Date();
      if (quickRange === "30d") {
        const start = new Date();
        start.setDate(today.getDate() - 30);
        matchPeriodo = d >= start && d <= today;
      } else if (quickRange === "90d") {
        const start = new Date();
        start.setDate(today.getDate() - 90);
        matchPeriodo = d >= start && d <= today;
      } else if (quickRange === "YTD") {
        const start = new Date(today.getFullYear(), 0, 1);
        matchPeriodo = d >= start && d <= today;
      }
    }

    const matchAno =
      ano === "Todos" || (d && d.getFullYear().toString() === ano.toString());

    const matchMesFiltro =
      mes === "Todos" || (d && meses[d.getMonth() + 1] === mes);

    return matchPeriodo && matchAno && matchMesFiltro;
  };

  const filtered = useMemo(() => {
    return rows.filter((r) => {
      const d = getRowDate(r);
      return matchesCommonFilters(r) && (isInExecution(r) || matchesDateFilters(d));
    });
  }, [rows, q, cliente, status, orcamentoStatusFilter, executionFilter, ano, mes, quickRange, activePage]);

  const paymentFiltered = useMemo(() => {
    return rows.filter((r) => {
      const d = getEffectivePaymentDate(r);
      return !!d && matchesCommonFilters(r) && matchesDateFilters(d);
    });
  }, [rows, q, cliente, status, orcamentoStatusFilter, executionFilter, ano, mes, quickRange, activePage]);

  const pendingInvoiceRows = useMemo(() => filtered.filter(isInExecution), [filtered]);

  const clientPortfolio = useMemo(() => summarizeClients(filtered, {
    valueOf: parseValor,
    isPaid: row => normalizePaymentStatus(row.status_raw || row.status) === "Pago",
    isInExecution,
    paymentDate: getEffectivePaymentDate,
  }), [filtered]);

  const pendingInvoiceReceivable = useMemo(
    () => ({
      total: pendingInvoiceRows.reduce((acc, r) => acc + parseValor(r.valor), 0),
      count: pendingInvoiceRows.length,
    }),
    [pendingInvoiceRows]
  );

  // ======== KPIs ========
  const total = filtered.reduce((a, b) => a + parseValor(b.valor), 0);
  const totalPago = filtered
    .filter((r) => normalizePaymentStatus(r.status_raw || r.status) === "Pago")
    .reduce((a, b) => a + parseValor(b.valor), 0);
  const totalPend = total - totalPago;
  const pctPago = total > 0 ? (totalPago / total) * 100 : 0;

  const orcamentoResumo = useMemo(() => {
    const resumo = {
      aprovados: 0,
      reprovados: 0,
      emAnalise: 0,
      naoInformado: 0,
      valorAprovado: 0,
      valorReprovado: 0,
    };

    filtered.forEach((r) => {
      const statusOrc = normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status);
      const valor = parseValor(r.valor);
      if (statusOrc === "Aprovado") {
        resumo.aprovados += 1;
        resumo.valorAprovado += valor;
      } else if (statusOrc === "Reprovado") {
        resumo.reprovados += 1;
        resumo.valorReprovado += valor;
      } else if (statusOrc === "Em analise") {
        resumo.emAnalise += 1;
      } else {
        resumo.naoInformado += 1;
      }
    });

    const decididos = resumo.aprovados + resumo.reprovados;
    return {
      ...resumo,
      total: filtered.length,
      decididos,
      taxaAprovacao: decididos ? (resumo.aprovados / decididos) * 100 : 0,
    };
  }, [filtered]);

  const notasResumo = useMemo(() => {
    const resumo = {
      emitidas: 0,
      semNota: 0,
      pagas: 0,
      pendentes: 0,
    };

    filtered.forEach((r) => {
      const temNota = hasValue(r.nf);
      if (!temNota) {
        resumo.semNota += 1;
        return;
      }

      resumo.emitidas += 1;
      const st = normalizePaymentStatus(r.status_raw || r.status);
      if (st === "Pago") resumo.pagas += 1;
      else resumo.pendentes += 1;
    });

    return resumo;
  }, [filtered]);

  // ======== ALERTAS DE ATRASO ========
  const atrasados = useMemo(() => {
    const hoje = new Date();

    return filtered
      .map((r) => {
        const dPag = getRowPaymentDateQ(r);
        const dEmi = getRowEmissionDate(r);
        const diffDays = dPag ? Math.floor((hoje - dPag) / 86400000) : null;
        const stNorm = normalizePaymentStatus(r.status_raw || r.status);

        return {
          ...r,
          servico: r.assunto || r.servico || "-",
          __dPag: dPag,
          __dEmi: dEmi,
          __diff: diffDays,
          __statusNorm: stNorm,
        };
      })
      .filter(
        (r) =>
          r.__dPag &&
          r.__diff != null &&
          r.__diff > 0 &&
          (r.__statusNorm === "Pendente" || r.__statusNorm === "Atrasado")
      );
  }, [filtered]);

  // ======== ALERTA: PRÃ“XIMOS PAGAMENTOS (MÃŠS ATUAL) ========
  const proximosPagamentos = useMemo(() => {
    const hoje = new Date();
    const mesAtual = hoje.getMonth();
    const anoAtual = hoje.getFullYear();

    return filtered
      .map((r) => {
        const d = getRowPaymentDateQ(r);
        return {
          ...r,
          __data: d,
        };
      })
      .filter((r) => {
        if (!r.__data) return false;
        const stNorm = normalizePaymentStatus(r.status_raw || r.status);
        if (!(stNorm === "Pendente" || stNorm === "Atrasado")) return false;

        const mes = r.__data.getMonth();
        const ano = r.__data.getFullYear();

        const futuro = r.__data >= hoje;
        const mesmoMes = mes === mesAtual && ano === anoAtual;

        return futuro && mesmoMes;
      })
      .sort((a, b) => a.__data - b.__data);
  }, [filtered]);

  // ======== INSIGHTS EXECUTIVOS ========

  const monthlyTrend = useMemo(() => {
    const base = new Date();
    const months = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(base.getFullYear(), base.getMonth() - i, 1);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const label = d.toLocaleDateString("pt-BR", { month: "short" }).replace(".", "");
      months.push({ key, label, total: 0, pago: 0 });
    }

    filtered.forEach((r) => {
      const d = getRowDate(r) || toDate(pick(r, EMI_KEYS));
      if (!d) return;
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const bucket = months.find((m) => m.key === key);
      if (!bucket) return;
      const valor = parseValor(r.valor);
      bucket.total += valor;
      if (normalizePaymentStatus(r.status_raw || r.status) === "Pago") {
        bucket.pago += valor;
      }
    });

    return months;
  }, [filtered]);

  const trendLast = monthlyTrend[monthlyTrend.length - 1] || { pago: 0, total: 0 };
  const trendPrev = monthlyTrend[monthlyTrend.length - 2] || { pago: 0, total: 0 };
  const momPaid =
    trendPrev.pago === 0 ? null : ((trendLast.pago - trendPrev.pago) / trendPrev.pago) * 100;

  const agingBuckets = useMemo(() => {
    const buckets = {
      futuro: 0,
      dias15: 0,
      dias30: 0,
      dias60: 0,
      dias60p: 0,
    };
    const hoje = new Date();

    filtered.forEach((r) => {
      const st = normalizePaymentStatus(r.status_raw || r.status);
      if (st === "Pago") return;
      const d = getEffectivePaymentDate(r);
      if (!d) return;
      const diff = Math.floor((hoje - d) / 86400000);
      const valor = parseValor(r.valor);

      if (diff <= 0) buckets.futuro += valor;
      else if (diff <= 15) buckets.dias15 += valor;
      else if (diff <= 30) buckets.dias30 += valor;
      else if (diff <= 60) buckets.dias60 += valor;
      else buckets.dias60p += valor;
    });

    const totalPendBucket =
      buckets.futuro + buckets.dias15 + buckets.dias30 + buckets.dias60 + buckets.dias60p;

    return {
      chart: [
        { name: "Futuro", valor: buckets.futuro },
        { name: "0-15d", valor: buckets.dias15 },
        { name: "16-30d", valor: buckets.dias30 },
        { name: "31-60d", valor: buckets.dias60 },
        { name: "60d+", valor: buckets.dias60p },
      ],
      totalPendBucket,
    };
  }, [filtered]);

  const avgAging = useMemo(() => {
    const hoje = new Date();
    let sum = 0;
    let count = 0;
    filtered.forEach((r) => {
      const st = normalizePaymentStatus(r.status_raw || r.status);
      if (st === "Pago") return;
      const d = getEffectivePaymentDate(r);
      if (!d) return;
      const diff = Math.floor((hoje - d) / 86400000);
      if (isFinite(diff) && diff >= 0) {
        sum += diff;
        count += 1;
      }
    });
    return count === 0 ? 0 : sum / count;
  }, [filtered]);

  const proj30d = useMemo(() => {
    const hoje = new Date();
    const limite = new Date();
    limite.setDate(hoje.getDate() + 30);
    return filtered
      .filter((r) => {
        const st = normalizePaymentStatus(r.status_raw || r.status);
        if (st === "Pago") return false;
        const d = getEffectivePaymentDate(r);
        if (!d) return false;
        return d >= hoje && d <= limite;
      })
      .reduce((acc, r) => acc + parseValor(r.valor), 0);
  }, [filtered]);

  const paymentCalendar = useMemo(() => {
    const hoje = new Date();
    const ancora = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    const mesesCalendario = Array.from({ length: 13 }, (_, index) => {
      const offset = index - 6;
      const data = new Date(ancora.getFullYear(), ancora.getMonth() + offset, 1);
      const key = `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, "0")}`;

      return {
        key,
        offset,
        label: data
          .toLocaleDateString("pt-BR", { month: "short", year: "2-digit" })
          .replace(".", ""),
        monthLabel: data.toLocaleDateString("pt-BR", { month: "long", year: "numeric" }),
        total: 0,
        pago: 0,
        pendente: 0,
        atrasado: 0,
        aberto: 0,
        count: 0,
      };
    });

    const monthMap = new Map(mesesCalendario.map((bucket) => [bucket.key, bucket]));

    paymentFiltered.forEach((r) => {
      const dataPagamento = getEffectivePaymentDate(r);
      if (!dataPagamento) return;

      const key = `${dataPagamento.getFullYear()}-${String(dataPagamento.getMonth() + 1).padStart(
        2,
        "0"
      )}`;
      const bucket = monthMap.get(key);
      if (!bucket) return;

      const valor = parseValor(r.valor);
      const statusNormalizado = normalizePaymentStatus(r.status_raw || r.status);

      bucket.total += valor;
      bucket.count += 1;

      if (statusNormalizado === "Pago") {
        bucket.pago += valor;
        return;
      }

      if (statusNormalizado === "Atrasado") {
        bucket.atrasado += valor;
      } else {
        bucket.pendente += valor;
      }

      bucket.aberto += valor;
    });

    return mesesCalendario;
  }, [paymentFiltered]);

  const currentPaymentMonth =
    paymentCalendar.find((bucket) => bucket.offset === 0) || {
      total: 0,
      pago: 0,
      aberto: 0,
      count: 0,
      monthLabel: "-",
    };

  const nextPaymentMonth =
    paymentCalendar.find((bucket) => bucket.offset === 1) || {
      total: 0,
      pago: 0,
      aberto: 0,
      count: 0,
      monthLabel: "-",
    };

  const paymentQuarterProjection = useMemo(
    () =>
      paymentCalendar
        .filter((bucket) => bucket.offset >= 0 && bucket.offset <= 2)
        .reduce(
          (acc, bucket) => ({
            total: acc.total + bucket.total,
            pago: acc.pago + bucket.pago,
            aberto: acc.aberto + bucket.aberto,
            count: acc.count + bucket.count,
          }),
          { total: 0, pago: 0, aberto: 0, count: 0 }
        ),
    [paymentCalendar]
  );

  const paymentHistory = useMemo(
    () =>
      paymentCalendar
        .filter((bucket) => bucket.offset < 0)
        .reduce(
          (acc, bucket) => ({
            total: acc.total + bucket.total,
            pago: acc.pago + bucket.pago,
            aberto: acc.aberto + bucket.aberto,
            count: acc.count + bucket.count,
          }),
          { total: 0, pago: 0, aberto: 0, count: 0 }
        ),
    [paymentCalendar]
  );

  const overdueCarryover = useMemo(
    () =>
      paymentCalendar
        .filter((bucket) => bucket.offset < 0)
        .reduce((acc, bucket) => acc + bucket.aberto, 0),
    [paymentCalendar]
  );

  const paymentRowsCount = useMemo(
    () => paymentCalendar.reduce((acc, bucket) => acc + bucket.count, 0),
    [paymentCalendar]
  );

  const paymentFutureMonths = useMemo(
    () => paymentCalendar.filter((bucket) => bucket.offset >= 0),
    [paymentCalendar]
  );

  const summaryYear = ano === "Todos" ? new Date().getFullYear() : Number(ano);
  const [selectedPaymentMonth, setSelectedPaymentMonth] = useState(null);
  const paymentSnapshotMonths = useMemo(() => {
    const today = new Date();
    const months = Array.from({ length: 12 }, (_, month) => ({
      key: `${summaryYear}-${String(month + 1).padStart(2, "0")}`,
      monthLabel: new Date(summaryYear, month, 1).toLocaleDateString("pt-BR", {
        month: "long", year: "numeric",
      }),
      offset: (summaryYear - today.getFullYear()) * 12 + month - today.getMonth(),
      total: 0, pago: 0, aberto: 0, count: 0,
    }));
    paymentFiltered.forEach((row) => {
      const date = getEffectivePaymentDate(row);
      if (!date || date.getFullYear() !== summaryYear) return;
      const bucket = months[date.getMonth()];
      const value = parseValor(row.valor);
      bucket.total += value;
      bucket.count += 1;
      if (normalizePaymentStatus(row.status_raw || row.status) === "Pago") {
        bucket.pago += value;
      } else {
        bucket.aberto += value;
      }
    });
    return months;
  }, [paymentFiltered, summaryYear]);

  const paymentCardDetails = useMemo(() => {
    const hoje = new Date();
    const ancora = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    const withPaymentMeta = paymentFiltered
      .map((r) => ({
        ...r,
        __dataPagamento: getEffectivePaymentDate(r),
        __statusNorm: normalizePaymentStatus(r.status_raw || r.status),
      }))
      .filter((r) => r.__dataPagamento)
      .sort((a, b) => a.__dataPagamento - b.__dataPagamento);

    const getMonthOffset = (date) =>
      (date.getFullYear() - ancora.getFullYear()) * 12 + (date.getMonth() - ancora.getMonth());

    return {
      currentMonth: {
        title: "Cai neste mes",
        subtitle: "Registros com pagamento previsto no mes atual",
        rows: withPaymentMeta.filter((r) => getMonthOffset(r.__dataPagamento) === 0),
      },
      nextMonth: {
        title: "Proximo mes",
        subtitle: "Registros com pagamento previsto para o proximo mes",
        rows: withPaymentMeta.filter((r) => getMonthOffset(r.__dataPagamento) === 1),
      },
      pendingInvoice: {
        title: "Serviços em andamento",
        subtitle: "Serviços aprovados em execução, sem data de pagamento. Emitir NF ao concluir.",
        dateLabel: "Data Criacao",
        rows: pendingInvoiceRows.map((r) => ({
          ...r,
          __dataPagamento: null,
          __statusNorm: "Em execução",
          __dateText: r.data_criacao || pick(r, EMI_KEYS) || "-",
        })),
      },
    };
  }, [paymentFiltered, pendingInvoiceRows]);

  // ======== CHARTS ========
  const COLORS = [
    "#3b82f6",
    "#22c55e",
    "#f59e0b",
    "#ef4444",
    "#a855f7",
    "#06b6d4",
    "#f97316",
    "#84cc16",
  ];

  const byCliente = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      if (!hasValue(r.cliente)) return;
      m.set(r.cliente, (m.get(r.cliente) || 0) + parseValor(r.valor));
    });
    return Array.from(m, ([cliente, valor]) => ({ cliente, valor }))
      .sort((a, b) => b.valor - a.valor)
      .slice(0, 12);
  }, [filtered]);

  const byStatus = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const st = normalizePaymentStatus(r.status_raw || r.status);
      m.set(st, (m.get(st) || 0) + parseValor(r.valor));
    });
    return Array.from(m, ([status, valor]) => ({ name: status, value: valor }));
  }, [filtered]);

  const byOrcamento = useMemo(() => {
    const m = new Map();
    filtered.forEach((r) => {
      const st = normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status);
      const current = m.get(st) || { name: st, count: 0, value: 0 };
      current.count += 1;
      current.value += parseValor(r.valor);
      m.set(st, current);
    });
    return Array.from(m.values()).sort((a, b) => b.count - a.count);
  }, [filtered]);

  const monthlyOrcamentoTrend = useMemo(() => {
    const base = new Date();
    const months = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(base.getFullYear(), base.getMonth() - i, 1);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const label = d.toLocaleDateString("pt-BR", { month: "short" }).replace(".", "");
      months.push({ key, label, aprovados: 0, reprovados: 0, emAnalise: 0 });
    }

    filtered.forEach((r) => {
      const d = getRowEmissionDate(r) || getRowDate(r);
      if (!d) return;
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const bucket = months.find((m) => m.key === key);
      if (!bucket) return;
      const st = normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status);
      if (st === "Aprovado") bucket.aprovados += 1;
      else if (st === "Reprovado") bucket.reprovados += 1;
      else if (st === "Em analise") bucket.emAnalise += 1;
    });

    return months;
  }, [filtered]);

  const monthlyExecutive = useMemo(() => {
    const months = Array.from({ length: 12 }, (_, monthIndex) => {
      const d = new Date(currentYear, monthIndex, 1);
      return {
        monthIndex,
        monthLabel: d.toLocaleDateString("pt-BR", { month: "short" }).replace(".", ""),
        totalOrcado: 0,
        totalAprovado: 0,
        totalReprovado: 0,
        totalEmAnalise: 0,
        qtdAprovado: 0,
        qtdReprovado: 0,
        qtdEmAnalise: 0,
        nfsEmitidas: 0,
      };
    });

    filtered.forEach((r) => {
      const d = getRowDate(r);
      if (!d) return;
      if (ano !== "Todos" && d.getFullYear() !== Number(ano)) return;

      const bucket = months[d.getMonth()];
      const valor = parseValor(r.valor);
      const statusOrc = normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status);

      bucket.totalOrcado += valor;
      if (statusOrc === "Aprovado") {
        bucket.totalAprovado += valor;
        bucket.qtdAprovado += 1;
      } else if (statusOrc === "Reprovado") {
        bucket.totalReprovado += valor;
        bucket.qtdReprovado += 1;
      } else if (statusOrc === "Em analise") {
        bucket.totalEmAnalise += valor;
        bucket.qtdEmAnalise += 1;
      }
      if (hasValue(r.nf)) bucket.nfsEmitidas += 1;
    });

    return months;
  }, [filtered, currentYear, ano]);

  const monthlyExecutivePie = useMemo(
    () =>
      monthlyExecutive
        .filter((m) => m.totalOrcado > 0)
        .map((m) => ({ name: m.monthLabel, value: m.totalOrcado })),
    [monthlyExecutive]
  );

  const monthlyBarsConfig = useMemo(() => {
    if (monthlyStatusMode === "Todos") {
      return [
        { dataKey: "totalAprovado", name: "Aprovado", fill: "#22c55e" },
        { dataKey: "totalReprovado", name: "Reprovado", fill: "#ef4444" },
        { dataKey: "totalEmAnalise", name: "Em analise", fill: "#f59e0b" },
      ];
    }

    if (monthlyStatusMode === "Reprovado") {
      return [{ dataKey: "totalReprovado", name: "Reprovado", fill: "#ef4444" }];
    }

    if (monthlyStatusMode === "Em analise") {
      return [{ dataKey: "totalEmAnalise", name: "Em analise", fill: "#f59e0b" }];
    }

    return [{ dataKey: "totalAprovado", name: "Aprovado", fill: "#22c55e" }];
  }, [monthlyStatusMode]);

  const portfolioShare = useMemo(() => {
    const totalValor = filtered.reduce((a, b) => a + parseValor(b.valor), 0);
    const ranked = byCliente.map((c) => ({
      ...c,
      share: totalValor > 0 ? (c.valor / totalValor) * 100 : 0,
    }));
    const top5 = ranked.slice(0, 5);
    const top5Share = top5.reduce((a, b) => a + b.share, 0);
    return { ranked, top5Share, totalValor };
  }, [filtered, byCliente]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    window._dashboardDebug = {
      rowsTotal: rows.length,
      rowsFiltrados: filtered.length,
      rowsPagamentoFiltrados: paymentFiltered.length,
      rowsSemNfParaReceber: pendingInvoiceRows.length,
      total,
      totalPago,
      totalPend,
      pctPago,
      totalSemNfParaReceber: pendingInvoiceReceivable.total,
      orcamentoResumo,
      notasResumo,
      topCliente: byCliente[0] || null,
      statusFinanceiro: byStatus,
      statusOrcamento: byOrcamento,
      tendenciaMensalFinanceira: monthlyTrend,
      tendenciaMensalOrcamento: monthlyOrcamentoTrend,
      executivoMensal: monthlyExecutive,
      calendarioRecebimentos: paymentCalendar,
    };
  }, [
    rows.length,
    filtered.length,
    paymentFiltered.length,
    pendingInvoiceRows.length,
    pendingInvoiceReceivable.total,
    total,
    totalPago,
    totalPend,
    pctPago,
    orcamentoResumo,
    notasResumo,
    byCliente,
    byStatus,
    byOrcamento,
    monthlyTrend,
    monthlyOrcamentoTrend,
    monthlyExecutive,
    paymentCalendar,
  ]);

  const rowsAprovados = useMemo(
    () =>
      filtered.filter(
        (r) => normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status) === "Aprovado"
      ),
    [filtered]
  );

  const rowsReprovados = useMemo(
    () =>
      filtered.filter(
        (r) => normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status) === "Reprovado"
      ),
    [filtered]
  );

  const rowsComNf = useMemo(() => filtered.filter((r) => hasValue(r.nf)), [filtered]);

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function openCardRowsDebug(title, rowsList) {
    const win = window.open("", "_blank", "width=1200,height=800");
    if (!win) return;

    const totalValor = rowsList.reduce((acc, r) => acc + parseValor(r.valor), 0);
    const rowsHtml = rowsList
      .map((r, idx) => {
        const dtCriacao = toDate(r["data criacao"])?.toLocaleDateString("pt-BR") || "-";
        const dtPag = toDate(r["data de pagamento"])?.toLocaleDateString("pt-BR") || "-";
        const stOrc = normalizeBudgetStatus(r.orcamento_status_raw || r.orcamento_status);
        return `
          <tr>
            <td>${idx + 1}</td>
            <td>${escapeHtml(r.id)}</td>
            <td>${escapeHtml(r.po)}</td>
            <td>${escapeHtml(r.cliente)}</td>
            <td>${escapeHtml(r.assunto)}</td>
            <td>${escapeHtml(r.orcamento_status_raw || "")}</td>
            <td>${escapeHtml(stOrc)}</td>
            <td>${escapeHtml(r.nf)}</td>
            <td>${escapeHtml(r.valor)}</td>
            <td>${escapeHtml(BRL(r.valor))}</td>
            <td>${escapeHtml(dtCriacao)}</td>
            <td>${escapeHtml(dtPag)}</td>
          </tr>
        `;
      })
      .join("");

    win.document.write(`
      <html>
        <head>
          <title>Debug Card - ${escapeHtml(title)}</title>
          <style>
            body { font-family: Segoe UI, Arial, sans-serif; margin: 16px; color: #111827; }
            h1 { margin: 0 0 6px; font-size: 22px; }
            .meta { margin-bottom: 12px; color: #4b5563; font-size: 14px; }
            table { width: 100%; border-collapse: collapse; font-size: 13px; }
            th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; }
            th { background: #f3f4f6; position: sticky; top: 0; }
            tbody tr:nth-child(even) { background: #f9fafb; }
          </style>
        </head>
        <body>
          <h1>${escapeHtml(title)}</h1>
          <div class="meta">Registros: <b>${rowsList.length}</b> | Soma do valor: <b>${escapeHtml(
      BRL(totalValor)
    )}</b></div>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>ID</th>
                <th>PO</th>
                <th>Cliente</th>
                <th>Assunto</th>
                <th>Aprovacao (raw)</th>
                <th>Aprovacao (normalizada)</th>
                <th>NF</th>
                <th>Valor (raw)</th>
                <th>Valor (R$)</th>
                <th>Data Criacao</th>
                <th>Data Pagamento</th>
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </body>
      </html>
    `);
    win.document.close();
    win.focus();
  }

  // ======== EXPORT CSV ========
  function exportCSV() {
    const cols = [
      "id",
      "po",
      "cliente",
      "assunto",
      "orcamento_status",
      "nf",
      "valor",
      "data_criacao",
      "data_pagamento",
      "status",
    ];
    const header = cols.join(";");
    const lines = filtered.map((r) =>
      [
        r.id || "",
        r.po || "",
        (r.cliente || "").replace(/;/g, ","),
        (r.assunto || "").replace(/;/g, ","),
        (r.orcamento_status || "").replace(/;/g, ","),
        (r.nf || "").replace(/;/g, ","),
        String(r.valor || 0).replace(".", ","),
        toDate(r["data criacao"])?.toLocaleDateString("pt-BR") || (isCriarNfMarker(r.data_criacao) ? "Criar NF" : ""),
        toDate(r["data de pagamento"])?.toLocaleDateString("pt-BR") || "",
        r.status || "",
      ].join(";")
    );
    const csv = [header, ...lines].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `financas-filtrado-${new Date()
      .toISOString()
      .slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function exportExecutivePDF() {
    const win = window.open("", "_blank", "width=980,height=1200");
    if (!win) return;
    const totalFmt = BRL(total);
    const pagoFmt = BRL(totalPago);
    const pendFmt = BRL(totalPend);
    const projFmt = BRL(proj30d);
    const riscoFmt = BRL(agingBuckets.totalPendBucket || 0);
    const pctPagoFmt = `${pctPago.toFixed(1)}%`;
    const taxaAprovFmt = `${orcamentoResumo.taxaAprovacao.toFixed(1)}%`;
    const resumoOrcFmt = `${orcamentoResumo.aprovados} aprovados / ${orcamentoResumo.reprovados} reprovados`;
    const resumoNfFmt = `${notasResumo.emitidas} emitidas | ${notasResumo.pendentes} pendentes`;
    const ticketPend =
      atrasados.length === 0 ? BRL(0) : BRL(totalPend / Math.max(atrasados.length, 1));
    const statusTotal = byStatus.reduce((a, b) => a + b.value, 0);
    const statusPalette = ["#22c55e", "#f59e0b", "#ef4444", "#3b82f6", "#6366f1", "#14b8a6"];
    const statusSegments = byStatus
      .map((s, idx) => {
        const pct = statusTotal > 0 ? (s.value / statusTotal) * 100 : 0;
        return `<span class="status-seg" style="width:${pct.toFixed(1)}%; background:${
          statusPalette[idx % statusPalette.length]
        };"></span>`;
      })
      .join("");
    const statusLegend = byStatus
      .map((s, idx) => {
        const pct = statusTotal > 0 ? (s.value / statusTotal) * 100 : 0;
        return `<div class="legend-item"><span class="legend-swatch" style="background:${
          statusPalette[idx % statusPalette.length]
        };"></span><span>${s.name}</span><span class="legend-value">${BRL(
          s.value
        )}</span><span class="legend-pct">${pct.toFixed(1)}%</span></div>`;
      })
      .join("");
    const maxClienteValor = byCliente.reduce((max, c) => Math.max(max, c.valor), 0);
    const topClienteRows = byCliente
      .slice(0, 12)
      .map((c) => {
        const pct = maxClienteValor ? (c.valor / maxClienteValor) * 100 : 0;
        return `<div class="mini-row"><div class="mini-label">${c.cliente}</div><div class="mini-track"><div class="mini-fill" style="width:${pct.toFixed(
          1
        )}%;"></div></div><div class="mini-value">${BRL(c.valor)}</div></div>`;
      })
      .join("");
    const rankingRows = portfolioShare.ranked
      .slice(0, 10)
      .map(
        (c, idx) =>
          `<tr><td>${idx + 1}</td><td>${c.cliente}</td><td>${BRL(
            c.valor
          )}</td><td>${c.share.toFixed(1)}%</td></tr>`
      )
      .join("");
    const bucketColors = ["#22c55e", "#f59e0b", "#f97316", "#ef4444", "#b91c1c"];
    const bucketRows = agingBuckets.chart
      .map((bucket, idx) => {
        const pct = agingBuckets.totalPendBucket
          ? (bucket.valor / agingBuckets.totalPendBucket) * 100
          : 0;
        return `
          <div class="bar-row">
            <div class="bar-label">${bucket.name}</div>
            <div class="bar-track">
              <div class="bar-fill" style="width:${pct.toFixed(1)}%; background:${
          bucketColors[idx % bucketColors.length]
        }"></div>
            </div>
            <div class="bar-value">${BRL(bucket.valor)}</div>
          </div>
        `;
      })
      .join("");
    const maxTrend = monthlyTrend.reduce((max, m) => Math.max(max, m.total), 0);
    const trendBars = monthlyTrend
      .map((m) => {
        const totalPct = maxTrend ? (m.total / maxTrend) * 100 : 0;
        const pagoPct = maxTrend ? (m.pago / maxTrend) * 100 : 0;
        return `<div class="trend-row"><div class="trend-label">${
          m.label
        }</div><div class="trend-bars"><div class="trend-bar"><span style="width:${totalPct.toFixed(
          1
        )}%; background:#93c5fd;"></span></div><div class="trend-bar"><span style="width:${pagoPct.toFixed(
          1
        )}%; background:#34d399;"></span></div></div><div class="trend-value">${BRL(
          m.total
        )}</div></div>`;
      })
      .join("");

    win.document.write(`
      <html>
        <head>
          <title>Relatorio Executivo - FinanceCRM</title>
          <style>
            @page { size: A4; margin: 16mm; }
            * { box-sizing: border-box; }
            body {
              font-family: "Segoe UI", Arial, sans-serif;
              padding: 0;
              margin: 0;
              color: #0f172a;
              background: #f8fafc;
              -webkit-print-color-adjust: exact;
              print-color-adjust: exact;
            }
            .page {
              padding: 24px;
            }
            .header {
              display: flex;
              align-items: center;
              justify-content: space-between;
              margin-bottom: 14px;
            }
            .brand {
              font-weight: 800;
              font-size: 22px;
            }
            .sub {
              color: #64748b;
              font-size: 12px;
            }
            .divider {
              height: 1px;
              background: #e2e8f0;
              margin: 14px 0;
            }
            .section-title {
              font-size: 12px;
              text-transform: uppercase;
              letter-spacing: 0.12em;
              color: #64748b;
              font-weight: 800;
              margin: 16px 0 8px;
            }
            .grid {
              display: grid;
              grid-template-columns: repeat(3, minmax(0, 1fr));
              gap: 12px;
            }
            .kpi-card {
              border-radius: 12px;
              padding: 12px 14px;
              border: 1px solid #e2e8f0;
              background: #ffffff;
              position: relative;
              overflow: hidden;
              box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
              page-break-inside: avoid;
            }
            .kpi-card::before {
              content: "";
              position: absolute;
              left: 0;
              right: 0;
              top: 0;
              height: 4px;
              background: var(--accent, #3b82f6);
            }
            .kpi-label {
              font-size: 11px;
              text-transform: uppercase;
              letter-spacing: 0.1em;
              color: #64748b;
              font-weight: 700;
            }
            .kpi-value {
              font-size: 18px;
              font-weight: 800;
              margin-top: 6px;
            }
            .kpi-sub {
              font-size: 12px;
              color: #64748b;
              margin-top: 4px;
            }
            .highlight {
              display: grid;
              grid-template-columns: repeat(2, minmax(0, 1fr));
              gap: 12px;
            }
            .alert-card {
              border-radius: 12px;
              padding: 12px 14px;
              color: #0f172a;
              background: #fff1f2;
              border: 1px solid #fecdd3;
              display: flex;
              align-items: center;
              justify-content: space-between;
              font-weight: 700;
            }
            .alert-card.warning {
              background: #fffbeb;
              border-color: #fde68a;
            }
            .alert-dot {
              width: 10px;
              height: 10px;
              border-radius: 50%;
              background: #ef4444;
              margin-right: 8px;
              display: inline-block;
            }
            .progress {
              height: 8px;
              background: #e2e8f0;
              border-radius: 999px;
              overflow: hidden;
              margin-top: 6px;
            }
            .progress > span {
              display: block;
              height: 100%;
              width: ${pctPago.toFixed(1)}%;
              background: linear-gradient(90deg, #3b82f6, #22c55e);
            }
            .bar-row {
              display: grid;
              grid-template-columns: 70px 1fr 90px;
              align-items: center;
              gap: 10px;
              margin-bottom: 6px;
              font-size: 12px;
            }
            .bar-track {
              height: 8px;
              border-radius: 999px;
              background: #e2e8f0;
              overflow: hidden;
            }
            .bar-fill {
              height: 100%;
              border-radius: 999px;
            }
            .bar-value {
              text-align: right;
              font-weight: 700;
              color: #0f172a;
            }
            .two-col {
              display: grid;
              grid-template-columns: repeat(2, minmax(0, 1fr));
              gap: 14px;
            }
            .card {
              border-radius: 12px;
              padding: 12px 14px;
              border: 1px solid #e2e8f0;
              background: #ffffff;
              box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
              page-break-inside: avoid;
            }
            .card-title {
              font-size: 13px;
              font-weight: 800;
              margin-bottom: 8px;
              color: #0f172a;
            }
            .legend-inline {
              display: flex;
              align-items: center;
              gap: 12px;
              font-size: 11px;
              color: #64748b;
              margin-bottom: 8px;
            }
            .legend-dot {
              width: 8px;
              height: 8px;
              border-radius: 50%;
              display: inline-block;
              margin-right: 6px;
            }
            .trend-list {
              display: grid;
              gap: 6px;
            }
            .trend-row {
              display: grid;
              grid-template-columns: 40px 1fr 90px;
              gap: 8px;
              align-items: center;
              font-size: 12px;
            }
            .trend-bars {
              display: flex;
              flex-direction: column;
              gap: 4px;
            }
            .trend-bar {
              height: 6px;
              border-radius: 999px;
              background: #e2e8f0;
              overflow: hidden;
            }
            .trend-bar span {
              display: block;
              height: 100%;
              border-radius: 999px;
            }
            .trend-value {
              text-align: right;
              color: #0f172a;
              font-weight: 700;
            }
            .mini-list {
              display: grid;
              gap: 6px;
            }
            .mini-row {
              display: grid;
              grid-template-columns: 120px 1fr 90px;
              gap: 8px;
              align-items: center;
              font-size: 12px;
            }
            .mini-track {
              height: 6px;
              border-radius: 999px;
              background: #e2e8f0;
              overflow: hidden;
            }
            .mini-fill {
              height: 100%;
              border-radius: 999px;
              background: linear-gradient(90deg, #60a5fa, #2563eb);
            }
            .mini-value {
              text-align: right;
              font-weight: 700;
            }
            .stacked-bar {
              display: flex;
              height: 10px;
              border-radius: 999px;
              overflow: hidden;
              background: #e2e8f0;
            }
            .status-seg {
              height: 100%;
              display: block;
            }
            .legend {
              display: grid;
              gap: 6px;
              margin-top: 10px;
            }
            .legend-item {
              display: grid;
              grid-template-columns: 10px 1fr auto auto;
              gap: 6px;
              align-items: center;
              font-size: 11px;
              color: #475569;
            }
            .legend-swatch {
              width: 10px;
              height: 10px;
              border-radius: 50%;
            }
            .legend-value {
              font-weight: 700;
              color: #0f172a;
            }
            .legend-pct {
              color: #64748b;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-top: 6px;
            }
            th, td {
              border-bottom: 1px solid #e2e8f0;
              padding: 8px 6px;
              text-align: left;
              font-size: 12px;
            }
            th {
              background: #f1f5f9;
              text-transform: uppercase;
              font-size: 11px;
              letter-spacing: 0.08em;
            }
          </style>
        </head>
        <body>
          <div class="page">
            <div class="header">
              <div>
                <div class="sub">Relatorio Executivo</div>
                <div class="brand">FinanceCRM</div>
                <div class="sub">Gerado em ${new Date().toLocaleString(
                  "pt-BR"
                )}</div>
              </div>
              <div class="sub">
                Registros filtrados: <strong>${filtered.length}</strong>
              </div>
            </div>

            <div class="divider"></div>

            <div class="section-title">Dashboard Financeiro</div>
            <div class="grid">
              <div class="kpi-card" style="--accent:#22c55e; background:#ecfdf3;">
                <div class="kpi-label">Recebido</div>
                <div class="kpi-value">${pagoFmt}</div>
                <div class="kpi-sub">Ultimo mes: ${BRL(trendLast.pago || 0)}</div>
              </div>
              <div class="kpi-card" style="--accent:#ef4444; background:#fff1f2;">
                <div class="kpi-label">Pendente</div>
                <div class="kpi-value">${pendFmt}</div>
                <div class="kpi-sub">Ticket medio: ${ticketPend}</div>
              </div>
              <div class="kpi-card" style="--accent:#3b82f6; background:#eff6ff;">
                <div class="kpi-label">Total Geral</div>
                <div class="kpi-value">${totalFmt}</div>
                <div class="kpi-sub">Carteira em risco: ${riscoFmt}</div>
              </div>
              <div class="kpi-card" style="--accent:#14b8a6; background:#ecfeff;">
                <div class="kpi-label">% Pago</div>
                <div class="kpi-value">${pctPagoFmt}</div>
                <div class="progress"><span></span></div>
              </div>
              <div class="kpi-card" style="--accent:#f59e0b; background:#fffbeb;">
                <div class="kpi-label">Projecao 30d</div>
                <div class="kpi-value">${projFmt}</div>
                <div class="kpi-sub">Aging medio: ${avgAging.toFixed(0)} dias</div>
              </div>
              <div class="kpi-card" style="--accent:#6366f1; background:#eef2ff;">
                <div class="kpi-label">Aprovacao Orcamento</div>
                <div class="kpi-value">${taxaAprovFmt}</div>
                <div class="kpi-sub">${resumoOrcFmt}</div>
              </div>
              <div class="kpi-card" style="--accent:#6366f1; background:#eef2ff;">
                <div class="kpi-label">Notas Fiscais</div>
                <div class="kpi-value">${notasResumo.emitidas}</div>
                <div class="kpi-sub">${resumoNfFmt}</div>
              </div>
            </div>

            <div class="section-title">Alertas Prioritarios</div>
            <div class="highlight">
              <div class="alert-card">
                <span><span class="alert-dot"></span>Pagamentos em atraso</span>
                <span>${atrasados.length} pendencias</span>
              </div>
              <div class="alert-card warning">
                <span>Pagamentos proximos</span>
                <span>${proximosPagamentos.length} neste mes</span>
              </div>
            </div>

            <div class="section-title">Aging e Risco</div>
            <div class="card">
              ${bucketRows || "<div class='sub'>Sem pendencias.</div>"}
            </div>

            <div class="section-title">Dashboards Visao Geral</div>
            <div class="two-col">
              <div class="card">
                <div class="card-title">Tendencia 12 meses</div>
                <div class="legend-inline">
                  <span><span class="legend-dot" style="background:#93c5fd;"></span>Total</span>
                  <span><span class="legend-dot" style="background:#34d399;"></span>Pago</span>
                </div>
                <div class="trend-list">
                  ${trendBars}
                </div>
              </div>
              <div class="card">
                <div class="card-title">Top 12 por Cliente</div>
                <div class="mini-list">
                  ${topClienteRows || "<div class='sub'>Sem dados.</div>"}
                </div>
              </div>
            </div>

            <div class="section-title">Status e Ranking</div>
            <div class="two-col">
              <div class="card">
                <div class="card-title">Status da Carteira</div>
                <div class="stacked-bar">
                  ${statusSegments || ""}
                </div>
                <div class="legend">
                  ${statusLegend || "<div class='sub'>Sem dados.</div>"}
                </div>
              </div>
              <div class="card">
                <div class="card-title">Ranking de Clientes</div>
                <table>
                  <thead><tr><th>#</th><th>Cliente</th><th>Valor</th><th>%</th></tr></thead>
                  <tbody>
                    ${rankingRows || "<tr><td colspan='4'>Sem dados.</td></tr>"}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </body>
      </html>
    `);
    win.document.close();
    win.focus();
    win.print();
  }




  // ======== TABELA / ALERTAS STATES ========
  const [limit, setLimit] = useState(5);
  const [showAtrasados, setShowAtrasados] = useState(false);
  const [showProximos, setShowProximos] = useState(false);
  const [activePaymentDetail, setActivePaymentDetail] = useState(null);

  const openPaymentDetail = (detailKey) => {
    setActivePaymentDetail((current) => {
      const next = current === detailKey ? null : detailKey;
      if (next && typeof window !== "undefined") {
        window.setTimeout(() => {
          document
            .getElementById("payment-card-details")
            ?.scrollIntoView({ behavior: "smooth", block: "start" });
        }, 80);
      }
      return next;
    });
  };

  const handleLogout = async () => {
    try {
      await msalInstance.logoutPopup();
      localStorage.clear();
      window.location.reload();
    } catch (err) {
      console.error("Erro ao sair:", err);
    }
  };

  const PageHero = () => (
    <div className="card hero">
      <div>
        <div className="eyebrow">{activePage === "clientes" ? "Carteira aprovada" : "Recebimentos"}</div>
        <div className="page-title executive-title" style={{ margin: 0 }}>
          {activePage === "clientes" ? "Clientes e serviços" : "Painel Financeiro"}
        </div>
        <p className="page-subtitle">
          {activePage === "clientes"
            ? "Valores dos serviços aprovados, recebimentos e trabalhos em andamento por cliente."
            : "Valores previstos para o mes atual, proximos meses e historico recente por data de pagamento."}
        </p>
      </div>
      <div className="hero-actions">
        <div className="tag">{activePage === "clientes" ? `${filtered.length} serviços aprovados` : `Registros com pagamento: ${paymentRowsCount}`}</div>
        <div className="year-focus">
          <span className="year-focus-label">Ano</span>
          <select
            className="year-focus-select"
            value={ano}
            onChange={(e) => setAno(e.target.value)}
          >
            {anos.map((a, index) => (
              <option key={index} value={a}>
                {a}
              </option>
            ))}
          </select>
        </div>
      </div>
    </div>
  );

  const [isMobile, setIsMobile] = useState(() => {
    if (typeof window === "undefined") return false;
    return window.innerWidth <= 768;
  });
  const [showFilters, setShowFilters] = useState(true);

  useEffect(() => {
    const handleResize = () => {
      const mobile = window.innerWidth <= 768;
      setIsMobile(mobile);
      if (mobile) {
        setShowFilters(false);
      } else {
        setShowFilters(true);
      }
    };
    handleResize();
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    if (activePage !== "overview") {
      setActivePaymentDetail(null);
    }
  }, [activePage]);

  const FiltersCard = ({ sidebar = false }) => (
    <div className={`card filters-compact${sidebar ? " sidebar-filters-card" : ""}`}>
      <div className={`filters-inline${sidebar ? " sidebar-filters-inline" : ""}`}>
        <div className="filter-group inline wide">
          <label className="filter-label">Buscar</label>
          <input
            type="text"
            placeholder="Cliente ou servico"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            className="input search-input"
          />
        </div>

        <div className="filter-group inline">
          <label className="filter-label">Cliente</label>
          <select className="select" value={cliente} onChange={(e) => setCliente(e.target.value)}>
            {clientes.map((c, index) => (
              <option key={index} value={c}>
                {c}
              </option>
            ))}
          </select>
        </div>

        <div className="filter-group inline">
          <label className="filter-label">Status</label>
          <select className="select" value={status} onChange={(e) => setStatus(e.target.value)}>
            {statusOptions.map((s, index) => (
              <option key={index} value={s}>
                {s}
              </option>
            ))}
          </select>
        </div>

        <div className="filter-group inline">
          <label className="filter-label">Orcamento</label>
          <select
            className="select"
            value={activePage === "clientes" ? "Aprovado" : orcamentoStatusFilter}
            disabled={activePage === "clientes"}
            onChange={(e) => setOrcamentoStatusFilter(e.target.value)}
          >
            {orcamentoOptions.map((s, index) => (
              <option key={index} value={s}>
                {s}
              </option>
            ))}
          </select>
        </div>

        <div className="filter-group inline">
          <label className="filter-label" htmlFor={sidebar ? "etapa-sidebar" : "etapa-inline"}>Etapa do serviço</label>
          <select id={sidebar ? "etapa-sidebar" : "etapa-inline"} className="select" value={executionFilter}
            onChange={(e) => {
              setExecutionFilter(e.target.value);
              if (e.target.value === "Em execução") {
                setStatus("Todos");
                setOrcamentoStatusFilter("Todos");
                setActivePage("overview");
                setActivePaymentDetail("pendingInvoice");
              }
            }}>
            <option value="Todos">Todas as etapas</option>
            <option value="Em execução">Aprovados em execução · Criar NF</option>
            <option value="Com pagamento">Com data de pagamento</option>
          </select>
          <div className="kpi-sub">Em execução: sem mês de recebimento. Ano, mês e período não limitam esses serviços.</div>
        </div>

        <div className="filter-group inline small">
          <label className="filter-label">Mes</label>
          <select className="select" value={mes} onChange={(e) => setMes(e.target.value)}>
            {meses.map((m, index) => (
              <option key={index} value={m}>
                {m === "Todos" ? "Todos" : m}
              </option>
            ))}
          </select>
        </div>

        <div className="filter-quick inline">
          <span className="filter-label">Periodo:</span>
          {[
            { k: "30d", label: "30d" },
            { k: "90d", label: "90d" },
            { k: "YTD", label: "Ano" },
            { k: "Todos", label: "Todos" },
          ].map((b) => (
            <button
              key={b.k}
              className={`chip ${quickRange === b.k ? "active" : ""}`}
              onClick={() => setQuickRange(b.k)}
            >
              {b.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );

  const goToPendencias = () => {
    openAtrasadosDetails();
  };

  const openAtrasadosDetails = () => {
    setActivePage("pendencias");
    setShowAtrasados(true);
    setShowProximos(false);
    if (typeof window !== "undefined") {
      window.setTimeout(() => {
        document
          .getElementById("pendencias-vencidas-detalhes")
          ?.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 80);
    }
  };

  const AlertsSummary = ({ showUpcoming = true, compact = false }) => (
    <div className="alerts-row">
      {atrasados.length > 0 && (
        <div
          className={`card alert-card${compact ? " compact" : ""}`}
          aria-expanded={compact ? undefined : showAtrasados}
          role={compact ? undefined : "button"}
          tabIndex={compact ? undefined : "0"}
          onClick={
            compact
              ? undefined
              : () => setShowAtrasados(!showAtrasados)
          }
          onKeyDown={
            compact
              ? undefined
              : (e) => {
                  if (e.key === "Enter" || e.key === " ") {
                    setShowAtrasados(!showAtrasados);
                    e.preventDefault();
                  }
                }
          }
          style={{
            background: "linear-gradient(135deg, rgba(255, 68, 68, 0.18), rgba(120, 15, 35, 0.2))",
            border: "1px solid rgba(255, 68, 68, 0.4)",
            color: "#fff",
            cursor: compact ? "default" : "pointer",
            transition: "all 0.3s ease",
            maxHeight: compact ? "78px" : "110px",
            overflow: compact ? "visible" : "hidden",
          }}
        >
          <div className={`alert-bar${compact ? " compact" : ""}`}>
            {compact ? (
              <>
                <span className="alert-dot" aria-hidden="true" />
                <div className="alert-text">
                  <span className="alert-title alert-title-sm">Pendente Pagamento</span>
                  <span className="alert-meta">
                    {atrasados.length} pendencia(s)
                  </span>
                </div>
                <button
                  className="alert-action"
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    goToPendencias();
                  }}
                >
                  Ir para pendencias
                </button>
              </>
            ) : (
              <>
                <span className="alert-title alert-title-sm">Pendente Pagamento</span>
                <span className={`alert-toggle ${showAtrasados ? "open" : ""}`}>
                  {showAtrasados ? "Ocultar" : "Ver detalhes"}
                </span>
              </>
            )}
          </div>
        </div>
      )}

      {showUpcoming && proximosPagamentos.length > 0 && (
        <div
          className="card alert-card"
          role="button"
          tabIndex="0"
          onClick={() => setShowProximos((p) => !p)}
          onKeyDown={(e) => {
            if (e.key === "Enter" || e.key === " ") {
              setShowProximos((p) => !p);
              e.preventDefault();
            }
          }}
          style={{
            background: "linear-gradient(135deg, rgba(255, 200, 0, 0.18), rgba(120, 100, 20, 0.2))",
            border: "1px solid rgba(255, 200, 0, 0.35)",
            color: "#fff",
            cursor: "pointer",
            transition: "all 0.3s ease",
          }}
        >
          <div className="alert-bar">
            <span className="alert-title alert-title-sm">Proximos Pagamentos</span>
            <span className="alert-toggle">
              {showProximos ? "Ocultar" : "Ver detalhes"}
            </span>
          </div>
        </div>
      )}
    </div>
  );

  const AlertsDetails = () => (
    <>
      {showAtrasados && atrasados.length > 0 && (
        <div
          id="pendencias-vencidas-detalhes"
          className="card"
          style={{ background: "rgba(255,255,255,0.05)" }}
        >
          <div
            style={{
              fontWeight: 600,
              color: "#3b82f6",
              marginBottom: "8px",
            }}
          >
            Pendencias Vencidas (detalhes)
          </div>

          <div className="table-wrapper">
            <table
              style={{
                width: "100%",
                fontSize: "0.9rem",
                color: "#ddd",
              }}
            >
              <thead>
                <tr>
                  <th>PO</th>
                  <th>Cliente</th>
                  <th>Servico</th>
                  <th>Valor</th>
                  <th>Data Emissao</th>
                  <th>Data Pagamento</th>
                  <th>Dias em atraso</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {atrasados.map((r, i) => (
                  <tr key={i}>
                    <td>{r.po || "-"}</td>
                    <td>{r.cliente || "-"}</td>
                    <td>{r.servico || "-"}</td>
                    <td>{BRL(r.valor)}</td>
                    <td>
                      {r.__dEmi ? r.__dEmi.toLocaleDateString("pt-BR") : "-"}
                    </td>
                    <td>
                      {r.__dPag ? r.__dPag.toLocaleDateString("pt-BR") : "-"}
                    </td>
                    <td
                      style={{
                        color: r.__diff > 30 ? "#ef4444" : "#facc15",
                      }}
                    >
                      {r.__diff != null ? `${r.__diff} dias` : "N/A"}
                    </td>
                    <td>
                      <span
                        className={`badge ${statusBadgeClass(r.__statusNorm || r.status || "")}`}
                      >
                        {r.__statusNorm || r.status || "-"}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {showProximos && proximosPagamentos.length > 0 && (
        <div className="card" style={{ background: "rgba(255,255,255,0.05)" }}>
          <div
            style={{
              fontWeight: 600,
              color: "#facc15",
              marginBottom: "8px",
            }}
          >
            Proximos Pagamentos (detalhes)
          </div>

          <div className="table-wrapper">
            <table style={{ width: "100%", fontSize: "0.9rem", color: "#ddd" }}>
              <thead>
                <tr>
                  <th>PO</th>
                  <th>Cliente</th>
                  <th>Servico</th>
                  <th>Valor</th>
                  <th>Data Pagamento</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {proximosPagamentos.map((r, i) => (
                  <tr key={i}>
                    <td>{r.po || "-"}</td>
                    <td>{r.cliente || "-"}</td>
                    <td>{r.assunto || "-"}</td>
                    <td>{BRL(r.valor)}</td>
                    <td>{r.__data ? r.__data.toLocaleDateString("pt-BR") : "-"}</td>
                    <td>
                      <span
                        className={`badge ${statusBadgeClass(r.status || "")}`}
                      >
                        {r.status || "-"}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </>
  );

  const KpiGrid = () => (
    <div className="grid kpi-grid">
      <div
        className="card kpi-card info"
        role="button"
        tabIndex={0}
        title="Abrir debug do Total orcado"
        onClick={() => openCardRowsDebug("Card: Total orcado", filtered)}
        onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && openCardRowsDebug("Card: Total orcado", filtered)}
      >
        <div className="kpi-title">Total orcado</div>
        <div className="kpi-value">{BRL(total)}</div>
        <div className="kpi-sub">Registros: {filtered.length}</div>
      </div>
      <div
        className="card kpi-card success"
        role="button"
        tabIndex={0}
        title="Abrir debug do Total aprovado"
        onClick={() => openCardRowsDebug("Card: Total aprovado", rowsAprovados)}
        onKeyDown={(e) =>
          (e.key === "Enter" || e.key === " ") &&
          openCardRowsDebug("Card: Total aprovado", rowsAprovados)
        }
      >
        <div className="kpi-title">Total aprovado</div>
        <div className="kpi-value">{BRL(orcamentoResumo.valorAprovado)}</div>
        <div className="kpi-sub">Aprovados: {orcamentoResumo.aprovados}</div>
      </div>
      <div
        className="card kpi-card warning"
        role="button"
        tabIndex={0}
        title="Abrir debug do Total reprovado"
        onClick={() => openCardRowsDebug("Card: Total reprovado", rowsReprovados)}
        onKeyDown={(e) =>
          (e.key === "Enter" || e.key === " ") &&
          openCardRowsDebug("Card: Total reprovado", rowsReprovados)
        }
      >
        <div className="kpi-title">Total reprovado</div>
        <div className="kpi-value">{BRL(orcamentoResumo.valorReprovado)}</div>
        <div className="kpi-sub">Reprovados: {orcamentoResumo.reprovados}</div>
      </div>
      <div className="card kpi-card accent">
        <div className="kpi-title">Taxa aprovacao</div>
        <div className="kpi-value">{orcamentoResumo.taxaAprovacao.toFixed(1)}%</div>
        <div className="kpi-sub">
          Aprovados: {orcamentoResumo.aprovados} | Reprovados: {orcamentoResumo.reprovados}
        </div>
      </div>
      <div
        className="card kpi-card muted"
        role="button"
        tabIndex={0}
        title="Abrir debug de Notas fiscais"
        onClick={() => openCardRowsDebug("Card: Notas fiscais (NF preenchida)", rowsComNf)}
        onKeyDown={(e) =>
          (e.key === "Enter" || e.key === " ") &&
          openCardRowsDebug("Card: Notas fiscais (NF preenchida)", rowsComNf)
        }
      >
        <div className="kpi-title">Notas fiscais</div>
        <div className="kpi-value">{notasResumo.emitidas}</div>
        <div className="kpi-sub">
          Pendentes: {notasResumo.pendentes} | Sem NF: {notasResumo.semNota}
        </div>
      </div>
    </div>
  );

  const PaymentKpiGrid = () => (
    <div className="grid kpi-grid">
      <div
        className={`card kpi-card info${activePaymentDetail === "currentMonth" ? " selected" : ""}`}
        role="button"
        tabIndex={0}
        onClick={() => openPaymentDetail("currentMonth")}
        onKeyDown={(e) =>
          (e.key === "Enter" || e.key === " ") && openPaymentDetail("currentMonth")
        }
      >
        <div className="kpi-title">Cai neste mes</div>
        <div className="kpi-value">{BRL(currentPaymentMonth.total)}</div>
        <div className="kpi-sub">
          Recebido: {BRL(currentPaymentMonth.pago)} | Em aberto: {BRL(currentPaymentMonth.aberto)}
        </div>
      </div>
      <div
        className={`card kpi-card success${activePaymentDetail === "nextMonth" ? " selected" : ""}`}
        role="button"
        tabIndex={0}
        onClick={() => openPaymentDetail("nextMonth")}
        onKeyDown={(e) =>
          (e.key === "Enter" || e.key === " ") && openPaymentDetail("nextMonth")
        }
      >
        <div className="kpi-title">Proximo mes</div>
        <div className="kpi-value">{BRL(nextPaymentMonth.total)}</div>
        <div className="kpi-sub">
          {nextPaymentMonth.monthLabel} | {nextPaymentMonth.count} registro(s)
        </div>
      </div>
      <div
        className={`card kpi-card warning${activePaymentDetail === "pendingInvoice" ? " selected" : ""}`}
        role="button"
        tabIndex={0}
        onClick={() => openPaymentDetail("pendingInvoice")}
        onKeyDown={(e) =>
          (e.key === "Enter" || e.key === " ") && openPaymentDetail("pendingInvoice")
        }
      >
        <div className="kpi-title">Serviços em andamento</div>
        <div className="kpi-value">{BRL(pendingInvoiceReceivable.total)}</div>
        <div className="kpi-sub">{pendingInvoiceReceivable.count} serviço(s) · Sem data de pagamento</div>
        <div className="kpi-sub">Emitir NF após a conclusão · Ver serviços</div>
      </div>
    </div>
  );

  const PaymentDetailsCard = () => {
    const detail = activePaymentDetail ? paymentCardDetails[activePaymentDetail] : null;

    if (!detail) return null;

    return (
      <div id="payment-card-details" className="card" style={{ background: "rgba(255,255,255,0.05)" }}>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: "10px",
            marginBottom: "10px",
            flexWrap: "wrap",
          }}
        >
          <div>
            <div style={{ fontWeight: 700, color: "#93c5fd" }}>{detail.title}</div>
            <div className="kpi-sub" style={{ marginTop: 2 }}>
              {detail.subtitle} | {detail.rows.length} registro(s) | Total:{" "}
              {BRL(detail.rows.reduce((acc, r) => acc + parseValor(r.valor), 0))}
            </div>
          </div>
          <button className="chip" type="button" onClick={() => setActivePaymentDetail(null)}>
            Fechar
          </button>
        </div>

        <div className="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>PO</th>
                <th>Cliente</th>
                <th>Servico</th>
                <th>NF</th>
                <th>Valor</th>
                <th>{detail.dateLabel || "Data Pagamento"}</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {detail.rows.length ? (
                detail.rows.map((row, index) => (
                  <tr key={`${row.id || row.po || row.cliente || "row"}-${index}`}>
                    <td>{row.po || "-"}</td>
                    <td>{row.cliente || "-"}</td>
                    <td>{row.assunto || row.servico || "-"}</td>
                    <td>{row.nf || "-"}</td>
                    <td>{BRL(row.valor)}</td>
                    <td>{row.__dateText || row.__dataPagamento?.toLocaleDateString("pt-BR") || "-"}</td>
                    <td>
                      <span className={`badge ${statusBadgeClass(row.__statusNorm || row.status || "")}`}>
                        {row.__statusNorm || row.status || "-"}
                      </span>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={7}>Nenhum registro encontrado para este card com os filtros atuais.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const PendenciasKpis = () => {
    const pendenciasRows = atrasados;
    const valorPendenteTotal = pendenciasRows.reduce((acc, r) => acc + parseValor(r.valor), 0);
    const qtdPendencias = pendenciasRows.length;
    const ticketPend = qtdPendencias === 0 ? 0 : valorPendenteTotal / qtdPendencias;

    const valorAtrasado = valorPendenteTotal;
    const atrasosCriticos = atrasados.filter((r) => (r.__diff || 0) > 30);
    const valorCritico = atrasosCriticos.reduce((acc, r) => acc + parseValor(r.valor), 0);

    const pendenciasAbertas = filtered.filter((r) => {
      const dataQ = getRowPaymentDateQ(r);
      if (!dataQ) return false;
      const stNorm = normalizePaymentStatus(r.status_raw || r.status);
      return stNorm === "Pendente" || stNorm === "Atrasado";
    });
    const valorAbertoTotal = pendenciasAbertas.reduce((acc, r) => acc + parseValor(r.valor), 0);

    const hoje = new Date();
    const proximos30Rows = pendenciasAbertas.filter((r) => {
      const dataQ = getRowPaymentDateQ(r);
      if (!dataQ) return false;
      const dias = Math.ceil((dataQ - hoje) / 86400000);
      return dias >= 0 && dias <= 30;
    });
    const valorProximos30 = proximos30Rows.reduce((acc, r) => acc + parseValor(r.valor), 0);

    const taxaInadimplencia =
      valorAbertoTotal === 0 ? 0 : (valorAtrasado / valorAbertoTotal) * 100;

    return (
      <div className="grid kpi-grid">
        <div
          className="card kpi-card warning"
          role="button"
          tabIndex={0}
          title="Abrir lista de pendencias vencidas"
          onClick={openAtrasadosDetails}
          onKeyDown={(e) =>
            (e.key === "Enter" || e.key === " ") && openAtrasadosDetails()
          }
        >
          <div className="kpi-title">Total pendente</div>
          <div className="kpi-value">{BRL(valorPendenteTotal)}</div>
          <div className="kpi-sub">Em aberto: {qtdPendencias} registro(s)</div>
        </div>
        <div className="card kpi-card info">
          <div className="kpi-title">Pendencias abertas</div>
          <div className="kpi-value">{pendenciasAbertas.length}</div>
          <div className="kpi-sub">Atrasadas: {atrasados.length}</div>
        </div>
        <div className="card kpi-card muted">
          <div className="kpi-title">Vencendo em 30 dias</div>
          <div className="kpi-value">{BRL(valorProximos30)}</div>
          <div className="kpi-sub">Registros: {proximos30Rows.length}</div>
        </div>
        <div className="card kpi-card accent">
          <div className="kpi-title">Atraso critico +30d</div>
          <div className="kpi-value">{BRL(valorCritico)}</div>
          <div className="kpi-sub">Registros: {atrasosCriticos.length}</div>
        </div>
        <div className="card kpi-card muted">
          <div className="kpi-title">Ticket medio pendente</div>
          <div className="kpi-value">{BRL(ticketPend)}</div>
          <div className="kpi-sub">Aging medio: {avgAging ? `${avgAging.toFixed(0)} dias` : "N/A"}</div>
        </div>
        <div className="card kpi-card info">
          <div className="kpi-title">Indice inadimplencia</div>
          <div className="kpi-value">{taxaInadimplencia.toFixed(1)}%</div>
          <div className="kpi-sub">
            Atrasado: {BRL(valorAtrasado)} / Aberto: {BRL(valorAbertoTotal)}
          </div>
        </div>
      </div>
    );
  };

  const TrendChartCard = () => (
    <div className="card">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Tendencia 12 meses
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        Pago M/M: {momPaid == null ? "N/A" : `${momPaid >= 0 ? "+" : ""}${momPaid.toFixed(1)}%`}
      </div>
      <div style={{ height: 320 }}>
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart data={monthlyTrend} margin={{ top: 10, right: 12, left: 0, bottom: 0 }}>
            <defs>
              <linearGradient id="gradPago" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#10b981" stopOpacity={0.7} />
                <stop offset="100%" stopColor="#10b981" stopOpacity={0.1} />
              </linearGradient>
              <linearGradient id="gradTotal" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#3b82f6" stopOpacity={0.6} />
                <stop offset="100%" stopColor="#3b82f6" stopOpacity={0.1} />
              </linearGradient>
            </defs>
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis
              dataKey="label"
              tick={{ fill: "#b0b8c1", fontSize: 11, fontWeight: 600 }}
              interval={1}
            />
            <YAxis
              tick={{ fill: "var(--muted)", fontSize: 12, fontWeight: 600 }}
              tickFormatter={(value) => BRL(value).replace("R$", "")}
            />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
              formatter={(value, name) => [BRL(value), name]}
            />
            <Area
              type="monotone"
              dataKey="total"
              name="Total"
              stroke="#3b82f6"
              fill="url(#gradTotal)"
              strokeWidth={2}
              dot={false}
            />
            <Area
              type="monotone"
              dataKey="pago"
              name="Pago"
              stroke="#10b981"
              fill="url(#gradPago)"
              strokeWidth={2}
              dot={false}
            />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const AgingChartCard = () => (
    <div className="card">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Aging / Risco
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        Pendencias: {BRL(agingBuckets.totalPendBucket)}
      </div>
      <div style={{ height: 320 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={agingBuckets.chart} margin={{ top: 10, right: 12, left: 0, bottom: 0 }}>
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis
              dataKey="name"
              tick={{ fill: "#b0b8c1", fontWeight: 600, fontSize: 12 }}
            />
            <YAxis
              tick={{ fill: "var(--muted)", fontSize: 12, fontWeight: 600 }}
              tickFormatter={(value) => BRL(value).replace("R$", "")}
            />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
              formatter={(value) => BRL(value)}
            />
            <Bar dataKey="valor" radius={[8, 8, 0, 0]}>
              {agingBuckets.chart.map((entry, i) => (
                <Cell key={i} fill={COLORS[i % COLORS.length]} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const TopClientsChartCard = () => (
    <div className="card">
      <div className="kpi-title" style={{ marginBottom: 8 }}>
        Top 12 por Cliente
      </div>
      <div style={{ height: 300 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            data={byCliente}
            margin={{ top: 10, right: 16, left: 0, bottom: 0 }}
          >
            <defs>
              <linearGradient id="barGrad" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#60a5fa" />
                <stop offset="100%" stopColor="#3b82f6" />
              </linearGradient>
            </defs>
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis
              dataKey="cliente"
              tick={{
                fill: "#b0b8c1",
                fontSize: 11,
                fontWeight: 600,
              }}
              interval={0}
              angle={-20}
              height={80}
              tickMargin={10}
              dy={20}
            />
            <YAxis
              tick={{
                fill: "var(--muted)",
                fontSize: 12,
                fontWeight: 600,
              }}
              domain={[0, (dataMax) => Math.ceil(dataMax * 1.1)]}
              tickFormatter={(value) => value.toLocaleString("pt-BR")}
            />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
              itemStyle={{ color: "#fff", fontWeight: 500 }}
              labelStyle={{
                color: "#00aaff",
                fontWeight: 600,
              }}
              formatter={(value) => BRL(value)}
            />
            <Bar
              dataKey="valor"
              radius={[8, 8, 0, 0]}
              cursor="pointer"
              onClick={(data) =>
                setCliente(cliente === data.cliente ? "Todos" : data.cliente)
              }
            >
              {byCliente.map((entry, i) => (
                <Cell
                  key={i}
                  fill="url(#barGrad)"
                  stroke={cliente === entry.cliente ? "#93c5fd" : "none"}
                  strokeWidth={cliente === entry.cliente ? 2 : 0}
                />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const StatusPieCard = () => (
    <div className="card">
      <div className="kpi-title" style={{ marginBottom: 8 }}>
        Por Status
      </div>
      <div style={{ height: 300 }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={byStatus}
              dataKey="value"
              nameKey="name"
              outerRadius={110}
              innerRadius={55}
              stroke="none"
            >
              {byStatus.map((e, i) => (
                <Cell key={i} fill={COLORS[i % COLORS.length]} />
              ))}
            </Pie>
            <Tooltip
              contentStyle={{
                background: "rgba(15,15,18,0.98)",
                border: "1px solid rgba(255,255,255,0.1)",
                color: "#fff",
                borderRadius: 8,
                padding: "8px 12px",
                boxShadow: "0 4px 12px rgba(0,0,0,0.4)",
              }}
              itemStyle={{ color: "#fff" }}
              labelStyle={{ color: "#ccc" }}
              formatter={(value, name) => [BRL(value), name]}
            />
            <Legend />
          </PieChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const OrcamentoPieCard = () => (
    <div className="card">
      <div className="kpi-title" style={{ marginBottom: 8 }}>
        Orcamentos por decisao
      </div>
      <div style={{ height: 300 }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={byOrcamento}
              dataKey="count"
              nameKey="name"
              outerRadius={110}
              innerRadius={55}
              stroke="none"
            >
              {byOrcamento.map((e, i) => (
                <Cell key={i} fill={COLORS[i % COLORS.length]} />
              ))}
            </Pie>
            <Tooltip
              contentStyle={{
                background: "rgba(15,15,18,0.98)",
                border: "1px solid rgba(255,255,255,0.1)",
                color: "#fff",
                borderRadius: 8,
                padding: "8px 12px",
                boxShadow: "0 4px 12px rgba(0,0,0,0.4)",
              }}
              itemStyle={{ color: "#fff" }}
              labelStyle={{ color: "#ccc" }}
              formatter={(value, name, payload) => [
                `${value} registros | ${BRL(payload?.payload?.value || 0)}`,
                name,
              ]}
            />
            <Legend />
          </PieChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const OrcamentoTrendCard = () => (
    <div className="card">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Aprovacoes x reprovacoes (12 meses)
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        Aprovados: {orcamentoResumo.aprovados} | Reprovados: {orcamentoResumo.reprovados}
      </div>
      <div style={{ height: 320 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            data={monthlyOrcamentoTrend}
            margin={{ top: 10, right: 12, left: 0, bottom: 0 }}
          >
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis dataKey="label" tick={{ fill: "#b0b8c1", fontSize: 11, fontWeight: 600 }} />
            <YAxis tick={{ fill: "var(--muted)", fontSize: 12, fontWeight: 600 }} />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
            />
            <Legend />
            <Bar dataKey="aprovados" fill="#22c55e" radius={[8, 8, 0, 0]} />
            <Bar dataKey="reprovados" fill="#ef4444" radius={[8, 8, 0, 0]} />
            <Bar dataKey="emAnalise" fill="#f59e0b" radius={[8, 8, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const MonthlyBarExecutiveCard = () => (
    <div className="card chart-card chart-wide">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Orcamentos por mes ({ano})
      </div>
      <div className="chart-toolbar">
        {["Aprovado", "Reprovado", "Em analise", "Todos"].map((opt) => (
          <button
            key={opt}
            type="button"
            className={`chart-filter-chip ${monthlyStatusMode === opt ? "active" : ""}`}
            onClick={() => setMonthlyStatusMode(opt)}
          >
            {opt}
          </button>
        ))}
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        {monthlyStatusMode === "Todos"
          ? "Comparativo mensal por status de orcamento"
          : `Serie mensal: ${monthlyStatusMode}`}
      </div>
      <div style={{ height: 290 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={monthlyExecutive} margin={{ top: 10, right: 12, left: 0, bottom: 0 }}>
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis dataKey="monthLabel" tick={{ fill: "#b0b8c1", fontSize: 11, fontWeight: 600 }} />
            <YAxis
              tick={{ fill: "var(--muted)", fontSize: 12, fontWeight: 600 }}
              tickFormatter={(value) => BRL(value).replace("R$", "")}
            />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
              formatter={(value) => BRL(value)}
            />
            <Legend />
            {monthlyBarsConfig.map((bar) => (
              <Bar
                key={bar.dataKey}
                dataKey={bar.dataKey}
                name={bar.name}
                fill={bar.fill}
                radius={[8, 8, 0, 0]}
              />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const MonthlyPieExecutiveCard = () => (
    <div className="card chart-card">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Participacao mensal do orcado ({ano})
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        Distribuicao por mes no ano selecionado
      </div>
      <div style={{ height: 320 }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={monthlyExecutivePie}
              dataKey="value"
              nameKey="name"
              outerRadius={110}
              innerRadius={55}
              stroke="none"
            >
              {monthlyExecutivePie.map((entry, i) => (
                <Cell key={i} fill={COLORS[i % COLORS.length]} />
              ))}
            </Pie>
            <Tooltip
              contentStyle={{
                background: "rgba(15,15,18,0.98)",
                border: "1px solid rgba(255,255,255,0.1)",
                color: "#fff",
                borderRadius: 8,
                padding: "8px 12px",
                boxShadow: "0 4px 12px rgba(0,0,0,0.4)",
              }}
              itemStyle={{ color: "#fff" }}
              labelStyle={{ color: "#ccc" }}
              formatter={(value) => BRL(value)}
            />
            <Legend />
          </PieChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const MonthlyDecisionCountCard = () => (
    <div className="card chart-card">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Volume mensal de decisoes ({ano})
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        Quantidade de aprovados, reprovados e em analise
      </div>
      <div style={{ height: 320 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={monthlyExecutive} margin={{ top: 10, right: 12, left: 0, bottom: 0 }}>
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis dataKey="monthLabel" tick={{ fill: "#b0b8c1", fontSize: 11, fontWeight: 600 }} />
            <YAxis allowDecimals={false} tick={{ fill: "var(--muted)", fontSize: 12, fontWeight: 600 }} />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
              formatter={(value, name) => [`${value} orcamento(s)`, name]}
            />
            <Legend />
            <Bar dataKey="qtdAprovado" name="Aprovados" fill="#22c55e" radius={[8, 8, 0, 0]} />
            <Bar dataKey="qtdReprovado" name="Reprovados" fill="#ef4444" radius={[8, 8, 0, 0]} />
            <Bar dataKey="qtdEmAnalise" name="Em analise" fill="#f59e0b" radius={[8, 8, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const PaymentCalendarChartCard = () => (
    <div className="card chart-card chart-wide">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Calendario de recebimento
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        Seis meses anteriores, mes atual e proximos seis meses por data de pagamento. Itens
        sem data de pagamento ficam fora do calendario.
      </div>
      <div style={{ height: 320 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={paymentCalendar} margin={{ top: 10, right: 12, left: 0, bottom: 0 }}>
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis dataKey="label" tick={{ fill: "#b0b8c1", fontSize: 11, fontWeight: 600 }} />
            <YAxis width={100}
              tick={{ fill: "var(--muted)", fontSize: 12, fontWeight: 600 }}
              tickFormatter={(value) => BRL(value).replace("R$", "")}
            />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
              formatter={(value) => BRL(value)}
            />
            <Legend />
            <Bar dataKey="pago" name="Pago" stackId="recebimentos" fill="#22c55e" radius={[8, 8, 0, 0]} />
            <Bar dataKey="pendente" name="Pendente" stackId="recebimentos" fill="#3b82f6" radius={[8, 8, 0, 0]} />
            <Bar dataKey="atrasado" name="Atrasado" stackId="recebimentos" fill="#f59e0b" radius={[8, 8, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const PaymentProjectionChartCard = () => (
    <div className="card chart-card">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Caixa mensal previsto
      </div>
      <div className="kpi-sub" style={{ marginBottom: 8 }}>
        Total do mes vs saldo em aberto do mes atual em diante
      </div>
      <div style={{ height: 320 }}>
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart data={paymentFutureMonths} margin={{ top: 10, right: 12, left: 0, bottom: 0 }}>
            <defs>
              <linearGradient id="gradReceberTotal" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#38bdf8" stopOpacity={0.7} />
                <stop offset="100%" stopColor="#38bdf8" stopOpacity={0.1} />
              </linearGradient>
              <linearGradient id="gradReceberAberto" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#f59e0b" stopOpacity={0.65} />
                <stop offset="100%" stopColor="#f59e0b" stopOpacity={0.08} />
              </linearGradient>
            </defs>
            <CartesianGrid stroke="var(--border)" vertical={false} />
            <XAxis dataKey="label" tick={{ fill: "#b0b8c1", fontSize: 11, fontWeight: 600 }} />
            <YAxis
              tick={{ fill: "var(--muted)", fontSize: 12, fontWeight: 600 }}
              tickFormatter={(value) => BRL(value).replace("R$", "")}
            />
            <Tooltip
              contentStyle={{
                background: "rgba(25,25,30,0.95)",
                border: "1px solid var(--border)",
                borderRadius: 10,
                color: "#ffffff",
                boxShadow: "0 2px 8px rgba(0,0,0,0.6)",
              }}
              formatter={(value, name) => [BRL(value), name]}
            />
            <Legend />
            <Area
              type="monotone"
              dataKey="total"
              name="Previsto no mes"
              stroke="#38bdf8"
              fill="url(#gradReceberTotal)"
              strokeWidth={2}
              dot={false}
            />
            <Area
              type="monotone"
              dataKey="aberto"
              name="Em aberto"
              stroke="#f59e0b"
              fill="url(#gradReceberAberto)"
              strokeWidth={2}
              dot={false}
            />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );

  const PaymentSnapshotCard = () => (
    <div className="card chart-card">
      <div className="kpi-title" style={{ marginBottom: 4 }}>
        Resumo por mes
      </div>
      <div className="kpi-sub" style={{ marginBottom: 10 }}>
        Janeiro a dezembro de {summaryYear} · Valores conforme os filtros selecionados
      </div>
      <div className="month-summary-list" tabIndex={0} role="region" aria-label={`Resumo mensal de ${summaryYear}`}>
        {paymentSnapshotMonths.map((bucket) => (
          <button type="button"
            key={bucket.key}
            className={`month-summary-row${bucket.offset === 0 ? " current" : ""}`}
            onClick={() => setSelectedPaymentMonth(bucket.key)}
            aria-haspopup="dialog"
            aria-label={`Ver pagamentos de ${bucket.monthLabel}`}
          >
            <div className="month-summary-head">
              <span className="month-summary-label">{bucket.monthLabel}</span>
              <span className="month-summary-total">{BRL(bucket.total)}</span>
            </div>
            <div className="month-summary-meta">
              <span>Pago: {BRL(bucket.pago)}</span>
              <span>Aberto: {BRL(bucket.aberto)}</span>
              <span>{bucket.count} registro(s)</span>
            </div>
          </button>
        ))}
      </div>
    </div>
  );

  const OverviewCharts = () => (
    <>
      <div className="grid charts-grid">
        <PaymentCalendarChartCard />
        <PaymentProjectionChartCard />
        <PaymentSnapshotCard />
      </div>
    </>
  );

  const RiskCharts = () => (
    <div className="grid charts-grid">
      <AgingChartCard />
      <StatusPieCard />
    </div>
  );

  const TableCard = () => (
    <div className="card">
      <div className="table-wrapper">
        <table>
          <thead>
            <tr>
              {columnMap.map((col, index) => (
                <th key={index} style={col.style}>
                  {col.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.slice(0, limit).map((row, index) => (
              <tr key={index}>
                {columnMap.map((col, i) => {
                  const rawValue = col.key === "etapa"
                    ? (isInExecution(row) ? "Em execução" : getEffectivePaymentDate(row) ? "Com pagamento previsto" : "Não definida")
                    : row[col.key] ?? row[col.key.replace(/\s/g, "_")];
                  let formatted = rawValue;

                  if (col.type === "currency") {
                    formatted = BRL(rawValue);
                  } else if (col.type === "date") {
                    const dateObj = toDate(rawValue);
                    formatted = dateObj ? dateObj.toLocaleDateString("pt-BR") : isCriarNfMarker(rawValue) ? "Criar NF" : "-";
                  }

                  if (col.type === "status") {
                    const statusClass = statusBadgeClass(formatted || "");

                    return (
                      <td key={i}>
                        <span className={`badge ${statusClass}`}>
                          {formatted || "-"}
                        </span>
                      </td>
                    );
                  }

                  return <td key={i}>{formatted || "-"}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {filtered.length > 5 && (
        <div style={{ textAlign: "center", marginTop: "16px" }}>
          {limit < filtered.length ? (
            <button className="theme-btn" onClick={() => setLimit(limit + 10)}>
              Listar mais
            </button>
          ) : (
            <button className="theme-btn" onClick={() => setLimit(5)}>
              Mostrar menos
            </button>
          )}
        </div>
      )}
    </div>
  );

  const ClientesKpis = () => (
    <div className="grid kpi-grid client-kpis">
      {[
        ["Total aprovado", clientPortfolio.total, clientPortfolio.count + " serviços • " + clientPortfolio.clients.length + " clientes", "info"],
        ["Recebido", clientPortfolio.paid, "Pagamentos confirmados", "success"],
        ["A receber com data", clientPortfolio.scheduled, "Recebimentos agendados em aberto", "accent"],
        ["Serviços em andamento", clientPortfolio.execution, "Criar NF • sem data de pagamento", "warning"],
      ].map(([title, value, subtitle, tone]) => (
        <div className={"card kpi-card " + tone} key={title}>
          <div className="kpi-title">{title}</div>
          <div className="kpi-value">{BRL(value)}</div>
          <div className="kpi-sub">{subtitle}</div>
        </div>
      ))}
    </div>
  );

  const ClientesTable = () => (
    <div className="card client-ranking">
      <div className="client-section-header">
        <div>
          <div className="kpi-title">Ranking de clientes</div>
          <div className="kpi-sub">Ordenado pelo valor aprovado • Apenas serviços aprovados no Status da planilha</div>
        </div>
        <span className="tag">{clientPortfolio.clients.length} clientes</span>
      </div>
      {clientPortfolio.unscheduled !== 0 && (
        <div className="client-note">Outros aprovados sem data de pagamento: <strong>{BRL(clientPortfolio.unscheduled)}</strong>. Incluídos no total aprovado.</div>
      )}
      <div className="table-wrapper">
        <table className="client-table">
          <thead><tr><th>Cliente / participação</th><th>Serviços</th><th>Total aprovado</th><th>Recebido</th><th>Saldo a receber</th><th></th></tr></thead>
          <tbody>
            {clientPortfolio.clients.map((client, index) => {
              const share = clientPortfolio.total > 0 ? client.total / clientPortfolio.total * 100 : 0;
              return (
                <tr key={client.name}>
                  <td>
                    <div className="client-name"><span className="client-rank">{index + 1}</span><strong>{client.name}</strong></div>
                    <div className="client-share"><div className="client-track"><span style={{ width: Math.max(0, Math.min(100, share)) + "%" }} /></div><span>{share.toFixed(1)}%</span></div>
                  </td>
                  <td>{client.count}</td><td className="client-money">{BRL(client.total)}</td>
                  <td className="client-money received">{BRL(client.paid)}</td><td className="client-money">{BRL(client.total - client.paid)}</td>
                  <td><button className="chip" onClick={() => { setCliente(client.name === "Cliente não informado" ? "Todos" : client.name); }}>Filtrar</button></td>
                </tr>
              );
            })}
            {!clientPortfolio.clients.length && <tr><td colSpan={6} className="client-empty">Nenhum serviço aprovado encontrado com os filtros atuais.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );

  const renderPage = () => {
    switch (activePage) {
      case "pendencias":
        return (
          <>
            <AlertsSummary />
            <AlertsDetails />
            <PendenciasKpis />
            <RiskCharts />
          </>
        );
      case "historico":
        return (
          <>
            <TableCard />
          </>
        );
      case "clientes":
        return (
          <>
            <ClientesKpis />
            <ClientesTable />
          </>
        );
      default:
        return (
          <>
            <AlertsSummary showUpcoming={false} compact />
            <PaymentKpiGrid />
            <PaymentDetailsCard />
            <OverviewCharts />
            <TableCard />
          </>
        );
    }
  };

  return (
    <div>
      {selectedPaymentMonth && paymentSnapshotMonths.some(month => month.key === selectedPaymentMonth) && (
        <MonthPaymentsDialog
          month={paymentSnapshotMonths.find(month => month.key === selectedPaymentMonth)}
          rows={paymentFiltered} getDate={getEffectivePaymentDate} formatMoney={BRL}
          getStatus={row => normalizePaymentStatus(row.status_raw || row.status)}
          onClose={() => setSelectedPaymentMonth(null)}
        />
      )}
      <div className="header">
        <div className="container header-inner">
          <div className="header-left">
            <div className="logo-wrap">
              <img src={logo} alt="Clever Connection Logo" className="logo" />
            </div>
            <div className="brand-stack">
              <span className="header-title-main">Clever Connection</span>
              <span className="header-subtitle">Finance Command Center</span>
            </div>
          </div>

          <div className="header-spacer" />

          <div className="header-right">
            {user && (
              <div className="user-info-desktop">
                {userPhoto ? (
                  <img src={userPhoto} alt="Foto do usuario" className="user-avatar" />
                ) : (
                  <div className="user-avatar anonymous">CC</div>
                )}

                <span className="user-name">{user.name}</span>
              </div>
            )}

            <button
              className="theme-btn mobile-menu-btn"
              onClick={() => setShowMobileMenu(true)}
              aria-expanded={showMobileMenu}
              title="Menu lateral"
            >
              Menu
            </button>
          </div>
        </div>
      </div>

      {showMobileMenu && (
        <div className="mobile-menu-overlay" onClick={() => setShowMobileMenu(false)}>
          <div className="mobile-menu-popup" onClick={(e) => e.stopPropagation()}>
            <div className="menu-header">
              <h3>Menu</h3>
              <button
                className="close-btn"
                onClick={() => setShowMobileMenu(false)}
                aria-label="Fechar Menu"
              >
                &times;
              </button>
            </div>

            <div className="mobile-nav">
              {navItems.map((item) => (
                <button
                  key={item.id}
                  className={`menu-item ${activePage === item.id ? "active" : ""}`}
                  onClick={() => {
                    setActivePage(item.id);
                    setShowMobileMenu(false);
                  }}
                >
                  <div className="menu-title">{item.label}</div>
                  <div className="menu-desc">{item.desc}</div>
                </button>
              ))}
            </div>

            <hr />

            <button
              className="menu-item"
              onClick={() => {
                exportCSV();
                setShowMobileMenu(false);
              }}
            >
              Exportar CSV
            </button>

            <button
              className="menu-item"
              onClick={() => {
                exportExecutivePDF();
                setShowMobileMenu(false);
              }}
            >
              PDF Executivo
            </button>

            {user && (
              <>
                <div className="user-info-mobile">
                  {userPhoto ? (
                    <img src={userPhoto} alt="Foto" />
                  ) : (
                    <div className="user-icon">CC</div>
                  )}
                  <span>
                    Logado como: <b>{user.name}</b>
                  </span>
                </div>
                <button className="menu-item danger" onClick={handleLogout}>
                  Sair da Conta
                </button>
              </>
            )}
          </div>
        </div>
      )}

      <div className="app-shell">
        <aside className="sidebar">
          <div className="sidebar-heading">Paginas</div>
          <div className="nav-list">
            {navItems.map((item) => (
              <button
                key={item.id}
                className={`nav-item ${activePage === item.id ? "active" : ""}`}
                onClick={() => setActivePage(item.id)}
              >
                <div className="nav-title">{item.label}</div>
                <div className="nav-desc">{item.desc}</div>
              </button>
            ))}
          </div>

          <div className="sidebar-footer">
            <button className="sidebar-btn" onClick={exportCSV}>
              Exportar CSV
            </button>
            <button className="sidebar-btn" onClick={exportExecutivePDF}>
              PDF Executivo
            </button>
            {user && (
              <button className="sidebar-btn danger" onClick={handleLogout}>
                Sair
              </button>
            )}
          </div>
        </aside>

        <main className="main-content">
          <div className="content-container">
            <PageHero />
            {isMobile && (
              <button
                className="theme-btn filters-toggle"
                onClick={() => setShowFilters((v) => !v)}
                aria-expanded={showFilters}
              >
                {showFilters ? "Ocultar filtros" : "Mostrar filtros"}
              </button>
            )}
            {isMobile && showFilters && <FiltersCard />}
            {loading ? (
              <div className="card">
                <div className="kpi-title">Carregando dados da planilha...</div>
                <div className="kpi-sub">Aguarde alguns segundos.</div>
              </div>
            ) : errMsg ? (
              <div className="card" style={{ borderColor: "rgba(239, 68, 68, 0.4)" }}>
                <div className="kpi-title" style={{ color: "#f87171" }}>
                  Nao foi possivel carregar os dados
                </div>
                <div className="kpi-sub">{errMsg}</div>
              </div>
            ) : (
              renderPage()
            )}
            <div className="footer">Clever Connection {new Date().getFullYear()}</div>
          </div>
        </main>

        {!isMobile && (
          <aside className="filters-column">
            <div className="sidebar-heading">Filtros</div>
            <FiltersCard sidebar />
          </aside>
        )}
      </div>
    </div>
  );
}
