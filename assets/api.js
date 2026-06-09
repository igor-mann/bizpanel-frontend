const API_URL = "https://bizpanel-api.onrender.com";

export async function fetchMetrics(orgId = "test-org-1") {
  const token = localStorage.getItem("token");
  const res = await fetch(`${API_URL}/api/v1/metrics?org_id=${orgId}`, {
    headers: token ? { "Authorization": `Bearer ${token}` } : {},
  });
  if (!res.ok) throw new Error("Ошибка загрузки данных");
  return res.json();
}

export async function pushMetrics(orgId, metrics) {
  const res = await fetch(`${API_URL}/api/v1/agent/push`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ org_id: orgId, metrics }),
  });
  if (!res.ok) throw new Error("Ошибка отправки данных");
  return res.json();
}

export async function register(email, password, fullName) {
  let res;
  try {
    res = await fetch(`${API_URL}/api/v1/auth/register`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password, full_name: fullName }),
    });
  } catch {
    throw new Error("Сервер недоступен. Подождите 30 секунд и попробуйте снова.");
  }
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.detail || "Ошибка регистрации");
  }
  return res.json();
}

export async function login(email, password) {
  let res;
  try {
    res = await fetch(`${API_URL}/api/v1/auth/login`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });
  } catch {
    throw new Error("Сервер недоступен. Подождите 30 секунд и попробуйте снова.");
  }
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.detail || "Неверный email или пароль");
  }
  return res.json();
}