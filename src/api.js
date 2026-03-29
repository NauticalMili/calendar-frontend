import axios from "axios";

const TOKEN_KEY = "elh_jwt";

export function getToken() {
  return localStorage.getItem(TOKEN_KEY);
}

export function setToken(t) {
  if (t) localStorage.setItem(TOKEN_KEY, t);
  else localStorage.removeItem(TOKEN_KEY);
}

const client = axios.create({
  baseURL: "/api",
});

client.interceptors.request.use((config) => {
  const t = getToken();
  if (t) {
    config.headers.Authorization = `Bearer ${t}`;
  }
  return config;
});

client.interceptors.response.use(
  (res) => res,
  (err) => {
    if (err.response?.status === 403) {
      const msg =
        typeof err.response?.data === "string"
          ? err.response.data
          : err.response?.data?.message || "Permission denied";
      window.dispatchEvent(new CustomEvent("elh-forbidden", { detail: { message: msg } }));
    }
    return Promise.reject(err);
  }
);

export async function login(username, password) {
  const { data } = await client.post("/auth/login", { username, password });
  return data;
}

export async function fetchEntries(geo) {
  const { data } = await client.get("/calendar-entries", { params: geo ? { geo } : {} });
  return data;
}

export async function createEntry(body) {
  const { data } = await client.post("/calendar-entries", body);
  return data;
}

export async function updateEntry(id, body) {
  const { data } = await client.put(`/calendar-entries/${id}`, body);
  return data;
}

export async function cancelEntry(id) {
  const { data } = await client.patch(`/calendar-entries/${id}/cancel`);
  return data;
}

export async function fetchAudit(entryId) {
  const { data } = await client.get(`/calendar-entries/${entryId}/audit`);
  return data;
}

export async function fetchNotifications() {
  const { data } = await client.get("/notifications");
  return data;
}

export async function downloadExcel(geo) {
  const params = new URLSearchParams();
  if (geo) params.set("geo", geo);
  const q = params.toString();
  const url = `/api/export/calendar.xlsx${q ? `?${q}` : ""}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${getToken()}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(text || "Export failed");
  }
  const blob = await res.blob();
  const a = document.createElement("a");
  const href = URL.createObjectURL(blob);
  a.href = href;
  a.download = "elh-calendar.xlsx";
  a.click();
  URL.revokeObjectURL(href);
}

export async function fetchAdminUsers() {
  const { data } = await client.get("/admin/users");
  return data;
}

export async function createAdminUser(body) {
  const { data } = await client.post("/admin/users", body);
  return data;
}
