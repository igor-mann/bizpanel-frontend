import { useState, useCallback, useMemo, useRef } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } from "recharts";
import * as XLSX from "xlsx";
import { fetchMetrics, login, register } from "./api";

/* ─────────────── THEME ─────────────── */
const C = {
  bg: "#0B1120", surface: "#111827", surfaceAlt: "#0F172A",
  border: "#1E293B", borderLight: "#334155",
  blue: "#3B82F6", blueDim: "#1E40AF",
  green: "#10B981", greenDim: "#065F46",
  red: "#EF4444", redDim: "#7F1D1D",
  amber: "#F59E0B", amberDim: "#78350F",
  purple: "#8B5CF6", cyan: "#06B6D4", pink: "#EC4899", orange: "#F97316",
  text: "#F1F5F9", textDim: "#94A3B8", textMuted: "#64748B",
};
const PALETTE = [C.blue, C.green, C.amber, C.red, C.purple, C.pink, C.cyan, C.orange, "#14B8A6", "#A855F7"];

/* ─────────────── API DATA PROCESSING ─────────────── */
const METRIC_NAMES = {
  revenue: "Выручка",
  profit: "Прибыль",
  receivables: "Дебиторская задолженность",
  payables: "Кредиторская задолженность",
  cash: "Денежные средства",
  tax_debt: "Задолженность по налогам",
  tax_liability: "Задолженность по налогам",
};

const METRIC_COLORS = {
  revenue: C.green,
  receivables: C.blue,
  payables: C.red,
  cash: C.cyan,
  tax_debt: C.amber,
};

function processApiData(rows) {
  const latest = {};
  rows.forEach(r => {
    if (!latest[r.metric_key] || r.collected_at > latest[r.metric_key].collected_at) {
      latest[r.metric_key] = r;
    }
  });
  return Object.values(latest);
}

function ApiDashboard({ metrics, onBack }) {
  const period = metrics[0]?.period_end || "";

  return (
    <div style={{ background: C.bg, minHeight: "100vh", fontFamily: "'Inter',sans-serif", color: C.text }}>
      <div style={{ borderBottom: `1px solid ${C.border}`, padding: "14px 28px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div onClick={onBack} style={{ width: 34, height: 34, borderRadius: 10, background: C.surface, border: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", fontSize: 14 }}>←</div>
          <div style={{ width: 34, height: 34, borderRadius: 10, background: `linear-gradient(135deg, ${C.blue}, ${C.green})`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 14, color: "#fff" }}>B</div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 800 }}>BizPanel</div>
            <div style={{ fontSize: 11, color: C.textMuted }}>Данные из 1С:Бухгалтерия</div>
          </div>
        </div>
        <div style={{ fontSize: 11, color: C.textDim, padding: "5px 12px", background: C.surface, borderRadius: 7, border: `1px solid ${C.border}` }}>
          📅 {period}
        </div>
      </div>

      <div style={{ padding: "24px 28px", maxWidth: 1200, margin: "0 auto" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 16, marginBottom: 32 }}>
          {metrics.map((m, i) => {
            const name = METRIC_NAMES[m.metric_key] || m.metric_key;
            const color = METRIC_COLORS[m.metric_key] || C.blue;
            const value = m.metric_value || 0;
            const formatted = value >= 1e12 ? `${(value/1e12).toFixed(1)} трлн`
              : value >= 1e9 ? `${(value/1e9).toFixed(1)} млрд`
              : value >= 1e6 ? `${(value/1e6).toFixed(1)} млн`
              : `${value.toFixed(0)}`;

            return (
              <div key={i} style={{ background: C.surface, border: `1px solid ${color}40`, borderRadius: 16, padding: "24px" }}>
                <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>{name}</div>
                <div style={{ fontSize: 28, fontWeight: 800, color, fontFamily: "'JetBrains Mono',monospace" }}>{formatted}</div>
                <div style={{ fontSize: 11, color: C.textDim, marginTop: 6 }}>сум</div>
              </div>
            );
          })}
        </div>

        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 24 }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 16 }}>Сравнение метрик</div>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={metrics.map(m => ({ name: METRIC_NAMES[m.metric_key] || m.metric_key, value: m.metric_value / 1e9 }))}>
              <XAxis dataKey="name" tick={{ fontSize: 10, fill: C.textDim }} />
              <YAxis tick={{ fontSize: 10, fill: C.textDim }} />
              <Tooltip content={({ active, payload }) => {
                if (!active || !payload?.[0]) return null;
                return (
                  <div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "8px 12px" }}>
                    <div style={{ fontSize: 12, color: C.text, fontWeight: 700 }}>{payload[0].payload.name}</div>
                    <div style={{ fontSize: 12, color: C.blue, fontFamily: "'JetBrains Mono',monospace" }}>{payload[0].value.toFixed(1)} млрд сум</div>
                  </div>
                );
              }} />
              <Bar dataKey="value" radius={[6,6,0,0]}>
                {metrics.map((m, i) => <Cell key={i} fill={METRIC_COLORS[m.metric_key] || C.blue} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ─────────────── UTILS ─────────────── */
const fmt = (v) => {
  const a = Math.abs(v);
  if (a >= 1e6) return `${(v/1e6).toFixed(1)} трлн`;
  if (a >= 1e3) return `${(v/1e3).toFixed(1)} млрд`;
  if (a >= 1) return `${v.toFixed(1)} млн`;
  return `${(v*1e3).toFixed(0)} тыс`;
};
const fmtN = (v) => new Intl.NumberFormat("ru-RU",{maximumFractionDigits:1}).format(v);

/* ─────────────── DATA PROCESSING ─────────────── */
function processJsonData(raw) {
  const all = Array.isArray(raw) ? raw : [raw].flat();
  let totalDebit = 0, totalCredit = 0;
  const byOrg = {}, byAccount = {}, byCounterparty = {};

  all.forEach(r => {
    const d = r.debit_balance_mln || 0;
    const c = r.credit_balance_mln || 0;
    totalDebit += d; totalCredit += c;
    const org = (r.organization || "").trim();
    if (!byOrg[org]) byOrg[org] = { debit: 0, credit: 0 };
    byOrg[org].debit += d; byOrg[org].credit += c;
    const acc = r.account_code || "—";
    const accName = r.account_name || "";
    if (!byAccount[acc]) byAccount[acc] = { name: accName, debit: 0, credit: 0 };
    byAccount[acc].debit += d; byAccount[acc].credit += c;
    const cp = r.subconto1 || "Без контрагента";
    if (!byCounterparty[cp]) byCounterparty[cp] = { debit: 0, credit: 0 };
    byCounterparty[cp].debit += d; byCounterparty[cp].credit += c;
  });

  const reportGroup = all[0]?.report_group || "Анализ данных";
  const balanceDate = all[0]?.balance_date || "—";
  const companyName = all[0]?.organization || "";
  const isReceivable = totalDebit > totalCredit;

  return {
    reportGroup, balanceDate, companyName, totalDebit, totalCredit, recordCount: all.length,
    isReceivable,
    mainValue: isReceivable ? totalDebit : totalCredit,
    orgs: Object.entries(byOrg)
      .map(([name, v]) => ({ name, debit: v.debit, credit: v.credit, value: isReceivable ? v.debit : v.credit }))
      .filter(x => x.value > 0).sort((a,b) => b.value - a.value),
    accounts: Object.entries(byAccount)
      .map(([code, v]) => ({ code, name: v.name, debit: v.debit, credit: v.credit, value: Math.max(v.debit, v.credit) }))
      .filter(x => x.value > 0).sort((a,b) => b.value - a.value),
    debitAccounts: Object.entries(byAccount)
      .map(([code, v]) => ({ code, name: v.name, value: v.debit }))
      .filter(x => x.value > 0).sort((a,b) => b.value - a.value),
    creditAccounts: Object.entries(byAccount)
      .map(([code, v]) => ({ code, name: v.name, value: v.credit }))
      .filter(x => x.value > 0).sort((a,b) => b.value - a.value),
    counterparties: Object.entries(byCounterparty)
      .map(([name, v]) => ({ name, debit: v.debit, credit: v.credit, value: isReceivable ? v.debit : v.credit }))
      .filter(x => x.value > 0).sort((a,b) => b.value - a.value),
  };
}

/* Standard Uzbek chart of accounts names */
const ACCOUNT_NAMES = {
  "0100":"Основные средства","0120":"Здания и сооружения","0130":"Машины и оборудование",
  "0140":"Мебель","0150":"Транспортные средства","0160":"Инструменты","0190":"Прочие ОС",
  "0200":"Амортизация ОС","0300":"Нематериальные активы","0400":"Инвестиции в недвижимость",
  "0500":"Оборудование к установке","0600":"Капитальные вложения","0700":"Долгосрочные инвестиции",
  "0800":"Долгосрочная дебиторка","0900":"Отложенные налоговые активы",
  "1000":"Материалы","1010":"Сырьё и материалы","1040":"Топливо","1060":"Запчасти",
  "1080":"Инвентарь","1090":"Прочие материалы",
  "2000":"Незавершённое производство","2300":"Полуфабрикаты","2500":"Брак",
  "2700":"Готовая продукция","2800":"Товары","2900":"Расходы будущих периодов",
  "2910":"Расходы будущих периодов",
  "4010":"Счета к получению (покупатели)","4020":"Внутригрупповая дебиторка",
  "4110":"Авансы выданные персоналу","4200":"Задолженность персонала",
  "4300":"Авансы выданные","4310":"Авансы поставщикам",
  "4400":"Авансовые платежи по налогам","4410":"Авансы по налогам","4510":"Авансы по ИНПС",
  "4700":"Задолженность учредителей","4800":"Прочая дебиторка","4810":"Аренда",
  "4820":"Краткосрочная аренда","4860":"Претензии","4890":"Прочие дебиторы",
  "5010":"Касса (сум)","5020":"Касса (валюта)","5110":"Расчётный счёт",
  "5200":"Валютные счета","5500":"Депозиты","5520":"Краткосрочные инвестиции",
  "5610":"Выбывающие активы",
  "6010":"Счета к оплате (поставщики)","6020":"Внутригрупповая кредиторка",
  "6110":"Авансы от покупателей","6300":"Авансы полученные","6310":"Авансы от покупателей",
  "6400":"Задолженность по платежам в бюджет","6410":"Налог на прибыль","6420":"НДС",
  "6500":"Страхование","6510":"ИНПС","6520":"Накопительный пенсионный",
  "6600":"Начисленные обязательства","6700":"Задолженность учредителям","6710":"Дивиденды",
  "6800":"Доходы будущих периодов","6900":"Прочая кредиторка","6950":"Кредиторка по лизингу",
  "6990":"Прочие обязательства",
  "7000":"Долгосрочные кредиты","7100":"Облигации","7200":"Долгосрочные обязательства",
  "7300":"Долгосрочные обязательства","7810":"Обязательства по лизингу",
  "7900":"Доходы будущих периодов (долгосрочные)",
  "8300":"Уставный капитал","8400":"Резервный капитал","8500":"Добавленный капитал",
  "8600":"Нераспределённая прибыль","8700":"Целевые поступления","8800":"Резервы",
  "9000":"Себестоимость реализации","9010":"Себестоимость продукции",
  "9100":"Доход от реализации","9110":"Выручка",
  "9200":"Расходы периода","9210":"Расходы на реализацию","9220":"Административные расходы",
  "9300":"Прочие доходы","9310":"Выбытие ОС","9320":"Безвозмездно полученное",
  "9400":"Прочие расходы","9430":"Убытки","9500":"Финансовые доходы",
  "9510":"Доходы по процентам","9600":"Финансовые расходы","9610":"Расходы по процентам",
  "9700":"Чрезвычайные доходы","9800":"Чрезвычайные расходы",
  "9900":"Конечный финансовый результат","9910":"Налог на прибыль (расход)",
};

function getAccountName(code) {
  if (ACCOUNT_NAMES[code]) return ACCOUNT_NAMES[code];
  const parent = code.split(".")[0];
  if (ACCOUNT_NAMES[parent]) return ACCOUNT_NAMES[parent];
  const group = parent.slice(0, 2) + "00";
  if (ACCOUNT_NAMES[group]) return ACCOUNT_NAMES[group];
  return "Счёт " + code;
}

function processExcelData(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const COLS = ["A","B","C","D","E","F","G","H","I","J"];

  /* Read with column-letter keys — works for all 1C ОСВ variants */
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: "A", defval: null });
  if (!rawRows.length) return null;

  /* ── Step 1: grab company and period from first 3 rows ── */
  let companyName = "", periodName = "";
  for (let i = 0; i < Math.min(rawRows.length, 4); i++) {
    const a = String(rawRows[i].A || "").trim();
    if (!a) continue;
    if (!companyName) { companyName = a; continue; }
    if (!periodName && a.toLowerCase().includes("ведомость")) { periodName = a; }
  }

  /* ── Step 2: find the header row that contains "Дебет" in any column ── */
  let headerRowIdx = -1;
  let endDebitCol = "G", endCreditCol = "I"; // defaults for full OSV
  let startDebitCol = "B", startCreditCol = "D";
  let isSingleAccount = false;

  for (let i = 0; i < Math.min(rawRows.length, 12); i++) {
    const r = rawRows[i];
    for (const col of COLS) {
      const v = String(r[col] || "").trim();
      if (v === "Дебет" || v === "дебет") {
        headerRowIdx = i;
        /* Determine column layout by scanning the row */
        const debitCols = COLS.filter(c => String(r[c] || "").trim() === "Дебет");
        const creditCols = COLS.filter(c => String(r[c] || "").trim() === "Кредит");
        startDebitCol = debitCols[0] || "B";
        startCreditCol = creditCols[0] || "D";
        endDebitCol = debitCols[debitCols.length - 1] || "G";
        endCreditCol = creditCols[creditCols.length - 1] || "I";
        /* Detect single-account OSV by checking row above for "Контрагенты" */
        const prevRow = i > 0 ? rawRows[i - 1] : null;
        if (prevRow) {
          const aVal = String(prevRow.A || "").trim();
          if (aVal === "Контрагенты" || aVal === "Субконто 1" || aVal.toLowerCase().includes("контраг")) {
            isSingleAccount = true;
          }
        }
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }

  if (headerRowIdx < 0) return null; /* Not a 1C OSV */

  const dataRows = rawRows.slice(headerRowIdx + 1);
  const mapped = [];

  /* ── Strategy A: Single-account OSV with subcontos (hierarchical) ── */
  const isAccountCode = (s) => /^\d{3,4}(\.\d+)?$/.test(String(s || "").trim());
  const isContract = (s) => {
    const t = String(s || "").trim();
    return t.startsWith("№") || t.startsWith("Договор") || t.startsWith("Акт")
      || t.startsWith("Перевод") || /от \d{2}\.\d{2}/.test(t);
  };

  /* Check if data is hierarchical (has counterparty rows without account codes) */
  const nonAccountRows = dataRows.filter(r => {
    const a = String(r.A || "").trim();
    return a && a !== "Итого" && !isAccountCode(a);
  });
  const counterpartyRows = nonAccountRows.filter(r => !isContract(r.A));

  if (counterpartyRows.length > 0) {
    /* Hierarchical: parse counterparty rows only */
    for (const r of dataRows) {
      const name = String(r.A || "").trim();
      if (!name || name === "Итого") continue;
      if (isAccountCode(name)) continue; /* skip account total row */
      if (isContract(name)) continue;   /* skip contract detail rows */

      const endD = typeof r[endDebitCol] === "number" ? r[endDebitCol] : 0;
      const endC = typeof r[endCreditCol] === "number" ? r[endCreditCol] : 0;
      const startD = typeof r[startDebitCol] === "number" ? r[startDebitCol] : 0;
      const startC = typeof r[startCreditCol] === "number" ? r[startCreditCol] : 0;
      if (endD === 0 && endC === 0 && startD === 0 && startC === 0) continue;

      mapped.push({
        account_code: "subconto",
        account_name: name,
        subconto1: name,
        organization: companyName || "Организация",
        debit_balance_mln: endD / 1e6,
        credit_balance_mln: endC / 1e6,
        report_group: periodName || "ОСВ",
        balance_date: new Date().toISOString().split("T")[0],
      });
    }
    if (mapped.length > 0) return mapped;
  }

  /* ── Strategy B: Full ОСВ (flat, one row per account) ── */
  for (const r of dataRows) {
    const code = String(r.A || "").trim();
    if (!code || code === "Итого") continue;

    const endD = typeof r[endDebitCol] === "number" ? r[endDebitCol] : 0;
    const endC = typeof r[endCreditCol] === "number" ? r[endCreditCol] : 0;
    const startD = typeof r[startDebitCol] === "number" ? r[startDebitCol] : 0;
    const startC = typeof r[startCreditCol] === "number" ? r[startCreditCol] : 0;
    if (endD === 0 && endC === 0 && startD === 0 && startC === 0) continue;

    mapped.push({
      account_code: code,
      account_name: getAccountName(code),
      subconto1: getAccountName(code),
      organization: companyName || "Организация",
      debit_balance_mln: endD / 1e6,
      credit_balance_mln: endC / 1e6,
      report_group: periodName || "ОСВ",
      balance_date: new Date().toISOString().split("T")[0],
    });
  }
  return mapped.length > 0 ? mapped : null;
}

/* ─────────────── COMPONENTS ─────────────── */
const inputStyle = {
  padding: "12px 14px", borderRadius: 10, border: `1px solid ${C.border}`,
  background: C.surfaceAlt, color: C.text, fontSize: 13, outline: "none",
};

const validatePassword = (pwd) => {
  if (pwd.length < 8) return "Минимум 8 символов";
  if (pwd.length > 64) return "Максимум 64 символа";
  return null;
};

function AuthScreen({ mode, onLogin, onRegister, onSwitchMode, error, loading }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [fullName, setFullName] = useState("");
  const [pwdError, setPwdError] = useState("");
  const [showPassword, setShowPassword] = useState(false);

  const handleSubmit = (e) => {
    e.preventDefault();
    const err = validatePassword(password);
    if (err) { setPwdError(err); return; }
    setPwdError("");
    if (mode === "login") onLogin(email, password);
    else onRegister(email, password, fullName);
  };

  return (
    <div style={{ background: C.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", padding: 24 }}>
      <div style={{ maxWidth: 380, width: "100%" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 32, justifyContent: "center" }}>
          <div style={{ width: 48, height: 48, borderRadius: 14, background: `linear-gradient(135deg, ${C.blue}, ${C.green})`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 22, color: "#fff" }}>B</div>
          <div>
            <div style={{ fontSize: 24, fontWeight: 800, color: C.text, letterSpacing: -1 }}>BizPanel</div>
            <div style={{ fontSize: 12, color: C.textMuted }}>{mode === "login" ? "Вход в аккаунт" : "Регистрация"}</div>
          </div>
        </div>

        <form onSubmit={handleSubmit} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 28, display: "flex", flexDirection: "column", gap: 14 }}>
          {mode === "register" && (
            <input type="text" placeholder="Имя" value={fullName} onChange={(e) => setFullName(e.target.value)} required style={inputStyle} />
          )}
          <input type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} required style={inputStyle} />
          <div style={{ position: "relative" }}>
            <input
              type={showPassword ? "text" : "password"}
              placeholder="Пароль"
              value={password}
              onChange={(e) => { setPassword(e.target.value); setPwdError(""); }}
              required
              style={{ ...inputStyle, width: "100%", paddingRight: 42, boxSizing: "border-box" }}
            />
            <span
              onClick={() => setShowPassword(!showPassword)}
              style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", cursor: "pointer", fontSize: 16, color: C.textMuted, userSelect: "none" }}
            >
              {showPassword ? "🙈" : "👁️"}
            </span>
          </div>
          {mode === "register" && (
            <div style={{ fontSize: 11, color: C.textMuted, marginTop: -8 }}>
              От 8 до 64 символов
            </div>
          )}
          {pwdError && <div style={{ color: C.red, fontSize: 12 }}>{pwdError}</div>}

          {error && <div style={{ color: C.red, fontSize: 12 }}>{error}</div>}

          <button type="submit" disabled={loading} style={{
            padding: "12px", borderRadius: 10, background: C.blue, color: "#fff", border: "none",
            fontSize: 14, fontWeight: 700, cursor: loading ? "default" : "pointer", opacity: loading ? 0.6 : 1,
          }}>
            {loading ? "Загрузка..." : mode === "login" ? "Войти" : "Зарегистрироваться"}
          </button>
        </form>

        <div style={{ textAlign: "center", marginTop: 16, fontSize: 13, color: C.textMuted }}>
          {mode === "login" ? "Нет аккаунта?" : "Уже есть аккаунт?"}{" "}
          <span onClick={onSwitchMode} style={{ color: C.blue, cursor: "pointer", fontWeight: 600 }}>
            {mode === "login" ? "Зарегистрироваться" : "Войти"}
          </span>
        </div>
      </div>
    </div>
  );
}

function UploadScreen({ onData, onLogout, onLogin }) {
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const fileInputRef = useRef(null);

  const handleLoadFromServer = useCallback(async () => {
    setError(""); setLoading(true);
    try {
      const rows = await fetchMetrics("test-org-1");
      if (!rows.length) { setError("Нет данных на сервере"); setLoading(false); return; }
      onData(rows, "server", true);
    } catch (e) { setError("Ошибка: " + e.message); }
    setLoading(false);
  }, [onData]);

  const handleFile = useCallback(async (file) => {
    console.log("handleFile called", file?.name, file?.size);
    setError(""); setLoading(true);
    try {
      if (file.name.endsWith(".json")) {
        const text = await file.text();
        const json = JSON.parse(text);
        const data = Array.isArray(json) ? json : [json];
        onData(processJsonData(data), file.name);
      } else if (file.name.match(/\.(xlsx|xls|csv)$/i)) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const mapped = processExcelData(wb);
        console.log("mapped", mapped?.length, mapped?.[0]);
        if (!mapped || !mapped.length) { setError("Не удалось распознать структуру ОСВ. Убедитесь что в файле есть колонки Дебет/Кредит."); setLoading(false); return; }
        onData(processJsonData(mapped), file.name);
      } else {
        setError("Поддерживаются форматы: .json, .xlsx, .xls, .csv");
      }
    } catch (e) { setError("Ошибка чтения файла: " + e.message); }
    setLoading(false);
  }, [onData]);

  const onDrop = useCallback((e) => { e.preventDefault(); setDragging(false); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }, [handleFile]);

  return (
    <div style={{ background: C.bg, minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", padding: 24, position: "relative" }}>
      {onLogout ? (
        <button onClick={onLogout} style={{ position: "absolute", top: 20, right: 24, padding: "8px 16px", borderRadius: 8, background: C.surface, color: C.textDim, border: `1px solid ${C.border}`, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
          Выйти
        </button>
      ) : (
        <button onClick={onLogin} style={{ position: "absolute", top: 20, right: 24, padding: "8px 16px", borderRadius: 8, background: C.blue, color: "#fff", border: "none", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
          Войти
        </button>
      )}
      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 40 }}>
        <div style={{ width: 48, height: 48, borderRadius: 14, background: `linear-gradient(135deg, ${C.blue}, ${C.green})`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 22, color: "#fff" }}>B</div>
        <div>
          <div style={{ fontSize: 28, fontWeight: 800, color: C.text, letterSpacing: -1 }}>BizPanel</div>
          <div style={{ fontSize: 13, color: C.textMuted }}>Аналитика 1С</div>
        </div>
      </div>

      <div style={{ maxWidth: 520, width: "100%", textAlign: "center", marginBottom: 32 }}>
        <div style={{ fontSize: 14, color: C.textDim, lineHeight: 1.7 }}>
          Загрузите из 1С: <span style={{color: C.text, fontWeight: 600}}>ОСВ в Excel</span> или <span style={{color: C.text, fontWeight: 600}}>JSON</span>, и мгновенно получите аналитический дашборд
        </div>
      </div>

      {/* Steps */}
      <div style={{ display: "flex", gap: 16, marginBottom: 36, flexWrap: "wrap", justifyContent: "center" }}>
        {[
          { n: "1", t: "Откройте 1С", d: "Бухгалтерия → Отчёты → ОСВ" },
          { n: "2", t: "Сформируйте ОСВ", d: "По нужному счёту или всем" },
          { n: "3", t: "Выгрузите в Excel", d: "Файл → Сохранить как .xlsx" },
          { n: "4", t: "Загрузите сюда ↓", d: "Перетащите или выберите" },
        ].map((s, i) => (
          <div key={i} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "14px 18px", width: 135, textAlign: "center" }}>
            <div style={{ width: 28, height: 28, borderRadius: 8, background: `${C.blue}20`, color: C.blue, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 8px", fontWeight: 800, fontSize: 13 }}>{s.n}</div>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.text, marginBottom: 4 }}>{s.t}</div>
            <div style={{ fontSize: 10, color: C.textMuted }}>{s.d}</div>
          </div>
        ))}
      </div>

      {/* Drop zone */}
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        style={{
          maxWidth: 520, width: "100%", border: `2px dashed ${dragging ? C.blue : C.borderLight}`,
          borderRadius: 20, padding: "48px 32px", textAlign: "center",
          background: dragging ? `${C.blue}08` : "transparent", transition: "all 0.2s",
          cursor: "pointer",
        }}
        onClick={() => fileInputRef.current.click()}
      >
        <input ref={fileInputRef} type="file" accept=".json,.xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
        {loading ? (
          <div style={{ color: C.blue, fontSize: 15 }}>Обработка файла...</div>
        ) : (
          <>
            <div style={{ fontSize: 40, marginBottom: 12 }}>📊</div>
            <div style={{ fontSize: 15, fontWeight: 600, color: C.text, marginBottom: 6 }}>Перетащите файл сюда</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 16 }}>или нажмите для выбора</div>
            <div style={{ display: "inline-flex", gap: 8 }}>
              {[".xlsx", ".xls", ".csv", ".json"].map(f => (
                <span key={f} style={{ fontSize: 11, color: C.textDim, background: C.surface, padding: "3px 10px", borderRadius: 6, border: `1px solid ${C.border}` }}>{f}</span>
              ))}
            </div>
          </>
        )}
      </div>

      {onLogout && (
        <button
          onClick={handleLoadFromServer}
          style={{
            marginTop: 16, padding: "12px 32px", borderRadius: 12,
            background: C.blue, color: "#fff", border: "none",
            fontSize: 14, fontWeight: 700, cursor: "pointer"
          }}
        >
          Загрузить из сервера
        </button>
      )}
      {error && <div style={{ color: C.red, fontSize: 13, marginTop: 16, maxWidth: 520, textAlign: "center" }}>{error}</div>}

      <div style={{ marginTop: 48, display: "flex", alignItems: "center", gap: 8, padding: "10px 20px", background: `${C.green}15`, border: `1px solid ${C.green}40`, borderRadius: 10 }}>
        <span style={{ fontSize: 16 }}>🔒</span>
        <span style={{ fontSize: 12, color: C.green, fontWeight: 600 }}>
          Ваши данные обрабатываются локально в браузере и никогда не отправляются на сервер
        </span>
      </div>
    </div>
  );
}

/* ─── Dashboard ─── */
function RankTable({ data, title, color, limit = 10 }) {
  const max = Math.max(...data.slice(0, limit).map(d => d.value));
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 24 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 14 }}>{title}</div>
      {data.slice(0, limit).map((item, i) => (
        <div key={i} style={{ position: "relative", marginBottom: 2 }}>
          <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: `${(item.value/max)*100}%`, background: `${color}10`, borderRadius: 6 }} />
          <div style={{ position: "relative", display: "flex", alignItems: "center", padding: "7px 12px", gap: 8 }}>
            <span style={{ fontSize: 10, color: C.textMuted, fontFamily: "'JetBrains Mono',monospace", width: 16, flexShrink: 0 }}>{i+1}</span>
            <span title={item.name || item.code} style={{ fontSize: 12, color: C.text, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", textAlign: "left" }}>{item.name || item.code}</span>
            <span style={{ fontSize: 12, fontWeight: 700, color, fontFamily: "'JetBrains Mono',monospace", flexShrink: 0, marginLeft: 12 }}>{fmt(item.value)}</span>
          </div>
        </div>
      ))}
    </div>
  );
}

function HBar({ data, color, title, limit = 10 }) {
  const d = data.slice(0, limit);
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 24 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 14 }}>{title}</div>
      <ResponsiveContainer width="100%" height={Math.max(d.length * 32, 120)}>
        <BarChart data={d} layout="vertical" margin={{ left: 0, right: 16, top: 0, bottom: 0 }}>
          <XAxis type="number" hide />
          <YAxis type="category" dataKey="name" width={140}
            tick={{ fontSize: 10, fill: C.textDim }}
            tickFormatter={v => v.length > 22 ? v.slice(0,20)+"…" : v}
            axisLine={false} tickLine={false} />
          <Tooltip content={({ active, payload }) => {
            if (!active || !payload?.[0]) return null;
            return (<div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "6px 12px" }}>
              <div style={{ fontSize: 11, color: C.textDim }}>{payload[0].payload.name}</div>
              <div style={{ fontSize: 13, fontWeight: 700, color, fontFamily: "'JetBrains Mono',monospace" }}>{fmtN(payload[0].value)} млн</div>
            </div>);
          }} />
          <Bar dataKey="value" radius={[0,6,6,0]} maxBarSize={22}>
            {d.map((_, i) => <Cell key={i} fill={`${color}${i===0?"cc":"55"}`} />)}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function DonutChart({ data, title }) {
  const d = data.slice(0, 7);
  const isSubconto = d[0]?.code === "subconto";
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 24 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 14 }}>{title}</div>
      <div style={{ display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ width: 170, height: 170, flexShrink: 0 }}>
          <ResponsiveContainer width="100%" height="100%">
            <PieChart>
              <Pie data={d} dataKey="value" cx="50%" cy="50%" innerRadius={50} outerRadius={75} paddingAngle={2} strokeWidth={0}>
                {d.map((_, i) => <Cell key={i} fill={PALETTE[i]} />)}
              </Pie>
              <Tooltip content={({ active, payload }) => {
                if (!active || !payload?.[0]) return null;
                const item = payload[0].payload;
                return (<div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "6px 12px", maxWidth: 260 }}>
                  <div style={{ fontSize: 11, color: C.textDim }}>{isSubconto ? item.name : `${item.code} — ${item.name}`}</div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: C.text, fontFamily: "'JetBrains Mono',monospace" }}>{fmt(item.value)}</div>
                </div>);
              }} />
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 0 }}>
          {d.map((item, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 8, height: 8, borderRadius: 2, background: PALETTE[i], flexShrink: 0 }} />
              {!isSubconto && <span style={{ fontSize: 11, color: C.textMuted, width: 38, flexShrink: 0, fontFamily: "'JetBrains Mono',monospace" }}>{item.code}</span>}
              <span style={{ fontSize: 11, color: C.text, flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={item.name}>{item.name}</span>
              <span style={{ fontSize: 11, fontWeight: 700, color: C.text, fontFamily: "'JetBrains Mono',monospace", flexShrink: 0, marginLeft: 4 }}>{fmt(item.value)}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function InsightsPanel({ data }) {
  const insights = [];
  const top = data.counterparties[0];
  const topPct = ((top?.value || 0) / data.mainValue * 100).toFixed(0);
  if (top && topPct > 30) insights.push({ icon: "⚠️", color: C.amber, text: `${top.name.slice(0,30)} — ${topPct}% от общей суммы. Высокая концентрация риска.` });
  if (data.counterparties.length > 0) {
    const top5sum = data.counterparties.slice(0,5).reduce((s,c) => s+c.value, 0);
    const top5pct = (top5sum / data.mainValue * 100).toFixed(0);
    insights.push({ icon: "📊", color: C.blue, text: `Топ-5 контрагентов = ${top5pct}% от общей суммы (${fmt(top5sum)} сум)` });
  }
  if (data.orgs.length > 1) {
    const maxOrg = data.orgs[0];
    const maxPct = (maxOrg.value / data.mainValue * 100).toFixed(0);
    insights.push({ icon: "🏢", color: C.purple, text: `${maxOrg.name.slice(0,25)} — ${maxPct}% (${fmt(maxOrg.value)} сум)` });
  }
  const zeroCount = data.counterparties.filter(c => c.value < 0.001 && c.value > 0).length;
  if (zeroCount > 5) insights.push({ icon: "🔍", color: C.cyan, text: `${zeroCount} контрагентов с копеечными остатками. Рекомендуется сверка.` });

  if (!insights.length) return null;
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 24 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 14 }}>💡 На что обратить внимание</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {insights.map((ins, i) => (
          <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "10px 14px", background: `${ins.color}08`, borderRadius: 10, border: `1px solid ${ins.color}20` }}>
            <span style={{ fontSize: 16, flexShrink: 0 }}>{ins.icon}</span>
            <span style={{ fontSize: 13, color: C.text, lineHeight: 1.5 }}>{ins.text}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function Dashboard({ data, fileName, onBack, token, setShowAuth, setAuthMode }) {

  return (
    <div style={{ background: C.bg, minHeight: "100vh", fontFamily: "'Inter',sans-serif", color: C.text }}>
      {/* Header */}
      <div style={{ borderBottom: `1px solid ${C.border}`, padding: "14px 28px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div onClick={onBack} style={{ width: 34, height: 34, borderRadius: 10, background: C.surface, border: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", fontSize: 14 }}>←</div>
          <div style={{ width: 34, height: 34, borderRadius: 10, background: `linear-gradient(135deg, ${C.blue}, ${C.green})`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 14, color: "#fff" }}>B</div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ fontSize: 15, fontWeight: 800, letterSpacing: -0.5 }}>BizPanel</div>
            <div style={{ fontSize: 13, color: C.textMuted }}>/</div>
            <div style={{ fontSize: 14, fontWeight: 600, color: C.text }}>{data.companyName || data.reportGroup}</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <div style={{ fontSize: 11, color: C.textDim, padding: "5px 12px", background: C.surface, borderRadius: 7, border: `1px solid ${C.border}` }}>📁 {fileName}</div>
          {data.reportGroup && <div style={{ fontSize: 11, color: C.textDim, padding: "5px 12px", background: C.surface, borderRadius: 7, border: `1px solid ${C.border}` }}>📋 {data.reportGroup.slice(0, 60)}</div>}
          <div style={{ fontSize: 11, color: C.textDim, padding: "5px 12px", background: C.surface, borderRadius: 7, border: `1px solid ${C.border}` }}>{data.recordCount} счетов</div>
        </div>
      </div>

      <div style={{ padding: "24px 28px", maxWidth: 1400, margin: "0 auto" }}>
        {/* Main cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginBottom: 24 }}>
          <div style={{ background: C.surface, border: `1px solid ${C.blue}40`, borderRadius: 14, padding: "20px 24px" }}>
            <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>Итого Дебет</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: C.blue, fontFamily: "'JetBrains Mono',monospace" }}>{fmt(data.totalDebit)}</div>
            <div style={{ fontSize: 11, color: C.textDim, marginTop: 4 }}>{fmtN(data.totalDebit)} млн сум</div>
          </div>
          <div style={{ background: C.surface, border: `1px solid ${C.red}40`, borderRadius: 14, padding: "20px 24px" }}>
            <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>Итого Кредит</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: C.red, fontFamily: "'JetBrains Mono',monospace" }}>{fmt(data.totalCredit)}</div>
            <div style={{ fontSize: 11, color: C.textDim, marginTop: 4 }}>{fmtN(data.totalCredit)} млн сум</div>
          </div>
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "20px 24px" }}>
            <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>Счетов</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: C.text, fontFamily: "'JetBrains Mono',monospace" }}>{data.accounts.length}</div>
            <div style={{ fontSize: 11, color: C.textDim, marginTop: 4 }}>с ненулевым сальдо</div>
          </div>
          <div style={{ background: C.surface, border: `1px solid ${data.totalDebit !== data.totalCredit ? C.amber + "40" : C.green + "40"}`, borderRadius: 14, padding: "20px 24px" }}>
            <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>Разница Дт−Кт</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: Math.abs(data.totalDebit - data.totalCredit) < 1 ? C.green : C.amber, fontFamily: "'JetBrains Mono',monospace" }}>
              {Math.abs(data.totalDebit - data.totalCredit) < 1 ? "✓ Баланс" : fmt(data.totalDebit - data.totalCredit)}
            </div>
            <div style={{ fontSize: 11, color: C.textDim, marginTop: 4 }}>{Math.abs(data.totalDebit - data.totalCredit) < 1 ? "Дебет = Кредит" : "Есть расхождение"}</div>
          </div>
        </div>

        {/* Insights */}
        <div style={{ marginBottom: 20 }}>
          <InsightsPanel data={data} />
        </div>

        {/* Charts */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(380px, 1fr))", gap: 18, marginBottom: 18 }}>
          {data.debitAccounts && data.debitAccounts[0]?.code === "subconto"
            ? <RankTable data={data.debitAccounts} title="🏆 Топ должников (сальдо конец)" color={C.blue} limit={15} />
            : <RankTable data={data.debitAccounts || data.accounts} title="📥 Дебетовые счета (Активы)" color={C.blue} />
          }
          {data.creditAccounts && data.creditAccounts.length > 0
            ? <RankTable data={data.creditAccounts} title="📤 Кредитовые счета (Пассивы)" color={C.red} />
            : data.debitAccounts && data.debitAccounts.length > 0 && <DonutChart data={data.debitAccounts.slice(0,8)} title="Структура дебиторки" />
          }
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(380px, 1fr))", gap: 18, marginBottom: 18 }}>
          {data.debitAccounts && data.debitAccounts[0]?.code !== "subconto" && data.debitAccounts.length > 1 && <DonutChart data={data.debitAccounts.slice(0,8)} title="Структура активов" />}
          {data.creditAccounts && data.creditAccounts[0]?.code !== "subconto" && data.creditAccounts.length > 1 && <DonutChart data={data.creditAccounts.slice(0,8)} title="Структура пассивов" />}
        </div>

        <div style={{ marginTop: 28, paddingTop: 14, borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", fontSize: 11, color: C.textMuted, flexWrap: "wrap", gap: 8 }}>
          <span style={{ display: "flex", alignItems: "center", gap: 6, color: C.green, fontWeight: 600 }}>🔒 Ваши данные обрабатываются локально в браузере и никогда не отправляются на сервер</span>
          <span>Загружено: {fileName} • {data.recordCount} строк</span>
        </div>
      </div>
      {!token && (
        <div style={{
          position: "fixed", bottom: 0, left: 0, right: 0,
          background: `linear-gradient(135deg, ${C.blue}, #1D4ED8)`,
          padding: "16px 24px",
          display: "flex", alignItems: "center", justifyContent: "space-between",
          flexWrap: "wrap", gap: 12, zIndex: 100,
          boxShadow: "0 -4px 24px rgba(0,0,0,0.3)"
        }}>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#fff" }}>
              Вы просматриваете в режиме гостя
            </div>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,0.8)", marginTop: 2 }}>
              Зарегистрируйтесь бесплатно: сохраняйте дашборды, подключайте 1С, получайте автообновление
            </div>
          </div>
          <button
            onClick={() => { setShowAuth(true); setAuthMode("register"); }}
            style={{
              padding: "10px 24px", borderRadius: 10,
              background: "#fff", color: C.blue,
              border: "none", fontSize: 13, fontWeight: 700, cursor: "pointer",
              whiteSpace: "nowrap"
            }}
          >
            Зарегистрироваться бесплатно →
          </button>
        </div>
      )}
    </div>
  );
}

function AppLayout({ user, onLogout, children }) {
  const [activeMenu, setActiveMenu] = useState("dashboards");

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.bg, fontFamily: "'Inter',sans-serif" }}>
      {/* Sidebar */}
      <div style={{ width: 240, background: C.surface, borderRight: `1px solid ${C.border}`, display: "flex", flexDirection: "column", padding: "20px 0", flexShrink: 0 }}>
        {/* Logo */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "0 20px 24px" }}>
          <div style={{ width: 36, height: 36, borderRadius: 10, background: `linear-gradient(135deg, ${C.blue}, ${C.green})`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 16, color: "#fff" }}>B</div>
          <div style={{ fontSize: 16, fontWeight: 800, color: C.text }}>BizPanel</div>
        </div>

        {/* Nav */}
        <div style={{ flex: 1, padding: "0 12px" }}>
          {[
            { id: "dashboards", icon: "📊", label: "Дашборды" },
            { id: "settings", icon: "⚙️", label: "Настройки" },
          ].map(item => (
            <div key={item.id} onClick={() => setActiveMenu(item.id)} style={{
              display: "flex", alignItems: "center", gap: 10, padding: "10px 12px",
              borderRadius: 8, cursor: "pointer", marginBottom: 4,
              background: activeMenu === item.id ? `${C.blue}20` : "transparent",
              color: activeMenu === item.id ? C.blue : C.textDim,
              fontWeight: activeMenu === item.id ? 600 : 400, fontSize: 14,
            }}>
              <span>{item.icon}</span>{item.label}
            </div>
          ))}
        </div>

        {/* User */}
        <div style={{ padding: "16px 20px", borderTop: `1px solid ${C.border}` }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginBottom: 4 }}>{user?.full_name || "Пользователь"}</div>
          <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 12 }}>{user?.email}</div>
          <button onClick={onLogout} style={{ width: "100%", padding: "8px", borderRadius: 8, background: "transparent", color: C.textMuted, border: `1px solid ${C.border}`, fontSize: 12, cursor: "pointer" }}>
            Выйти
          </button>
        </div>
      </div>

      {/* Content */}
      <div style={{ flex: 1, overflow: "auto" }}>
        {/* Header */}
        <div style={{ borderBottom: `1px solid ${C.border}`, padding: "16px 28px", background: C.surface }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text }}>
            {activeMenu === "dashboards" ? "Мои дашборды" : "Настройки"}
          </div>
        </div>
        {/* Page content */}
        <div style={{ padding: 28 }}>{children}</div>
      </div>
    </div>
  );
}

function DashboardsPage({ onLoadFile, onLoadFromServer }) {
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div style={{ fontSize: 14, color: C.textMuted }}>Загрузите ОСВ из 1С для анализа</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onLoadFromServer} style={{ padding: "10px 20px", borderRadius: 10, background: C.green, color: "#fff", border: "none", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            ⚡ Загрузить из 1С
          </button>
          <button onClick={onLoadFile} style={{ padding: "10px 20px", borderRadius: 10, background: C.blue, color: "#fff", border: "none", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            + Загрузить ОСВ
          </button>
        </div>
      </div>
      <div style={{ textAlign: "center", padding: "80px 0", color: C.textMuted, fontSize: 14 }}>
        📊 Нет загруженных дашбордов
      </div>
    </div>
  );
}

/* ─────────────── APP ─────────────── */
export default function BizPanelApp() {
  const [result, setResult] = useState(null);
  const [fileName, setFileName] = useState("");
  const [token, setToken] = useState(localStorage.getItem("bp_token"));
  const [authMode, setAuthMode] = useState("login"); // login | register
  const [authError, setAuthError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [showAuth, setShowAuth] = useState(false);
  const [user, setUser] = useState(null);
  const hiddenFileRef = useRef(null);

  const handleLogin = async (email, password) => {
    setAuthLoading(true); setAuthError("");
    try {
      const data = await login(email, password);
      localStorage.setItem("bp_token", data.access_token);
      localStorage.setItem("bp_user_id", data.user.id);
      setToken(data.access_token);
      setUser(data.user);
    } catch (e) { setAuthError(e.message); }
    setAuthLoading(false);
  };

  const handleRegister = async (email, password, fullName) => {
    setAuthLoading(true); setAuthError("");
    try {
      const data = await register(email, password, fullName);
      localStorage.setItem("bp_token", data.access_token);
      localStorage.setItem("bp_user_id", data.user.id);
      setToken(data.access_token);
      setUser(data.user);
    } catch (e) { setAuthError(e.message); }
    setAuthLoading(false);
  };

  const handleLogout = () => {
    localStorage.removeItem("bp_token");
    localStorage.removeItem("bp_user_id");
    setToken(null);
    setResult(null);
  };

  const handleData = useCallback((data, name, isApi = false) => {
    console.log("handleData called", name, isApi, !!data);
    setResult(isApi ? { __isApi: true, metrics: processApiData(data) } : data);
    setFileName(name);
  }, []);

  const handleLoadFromServer = useCallback(async () => {
    const orgId = localStorage.getItem("bp_user_id");
    if (!orgId) { alert("Войдите в аккаунт"); return; }
    try {
      const rows = await fetchMetrics(orgId);
      if (!rows.length) { alert("Нет данных на сервере"); return; }
      handleData(rows, "1С сервер", true);
    } catch (e) { alert("Ошибка: " + e.message); }
  }, [handleData]);

  const handleFileUpload = useCallback(async (file) => {
    try {
      if (file.name.endsWith(".json")) {
        const text = await file.text();
        const json = JSON.parse(text);
        const data = Array.isArray(json) ? json : [json];
        handleData(processJsonData(data), file.name);
      } else if (file.name.match(/\.(xlsx|xls|csv)$/i)) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const mapped = processExcelData(wb);
        if (!mapped || !mapped.length) return;
        handleData(processJsonData(mapped), file.name);
      }
    } catch (e) { console.error("Ошибка чтения файла:", e.message); }
  }, [handleData]);

  if (!token && showAuth) return <AuthScreen
    mode={authMode}
    onLogin={handleLogin}
    onRegister={handleRegister}
    onSwitchMode={() => { setAuthMode(authMode === "login" ? "register" : "login"); setAuthError(""); }}
    error={authError}
    loading={authLoading}
  />;
  if (result && result.__isApi) return <ApiDashboard metrics={result.metrics} onBack={() => { setResult(null); setFileName(""); }} />;
  if (result) return <Dashboard data={result} fileName={fileName} onBack={() => { setResult(null); setFileName(""); }} token={token} setShowAuth={setShowAuth} setAuthMode={setAuthMode} />;
  if (!token) return <UploadScreen onData={handleData} onLogout={null} onLogin={() => { setAuthMode("login"); setShowAuth(true); }} />;
  return <AppLayout user={user} onLogout={handleLogout}>
    <DashboardsPage onLoadFile={() => hiddenFileRef.current.click()} onLoadFromServer={handleLoadFromServer} />
    <input ref={hiddenFileRef} type="file" accept=".json,.xlsx,.xls,.csv" style={{ display: "none" }}
      onChange={(e) => e.target.files[0] && handleFileUpload(e.target.files[0])} />
  </AppLayout>;
}
