import { useCallback, useEffect, useMemo, useState } from "react";
import {
  cancelEntry,
  createEntry,
  fetchAudit,
  fetchEntries,
  fetchNotifications,
  fetchAdminUsers,
  createAdminUser,
  getToken,
  login,
  setToken,
  updateEntry,
} from "./api";
import { THEME_STORAGE_KEY } from "./theme.js";
import * as XLSX from "xlsx";
import "./App.css";

const ROLES = {
  CENTRAL_ADMIN: "CENTRAL_ADMIN",
  USER_EDITOR: "USER_EDITOR",
  USER_VIEWER: "USER_VIEWER",
};

const SORT_LABELS = {
  date: "Date",
  title: "Class name",
  location: "Market",
};

const MARKET_OPTIONS = ["Americas", "APAC", "GIC", "JAPAN", "R-EMEA", "UKI"];
const ML_CIC_OPTIONS = ["ML", "CIC"];
const EPH_TEAM_ROLE_OPTIONS = ["Team Member", "Leader", "LC"];
const MODE_OF_DELIVERY_OPTIONS = ["F2F", "Hybrid", "LVC", "Virtual"];
const QUARTER_OPTIONS = ["Q1", "Q2", "Q3", "Q4"];
const CLASS_SIZE_OPTIONS = [
  "0-10",
  "10-20",
  "20-30",
  "30-40",
  "40-50",
  "50-60",
  "60-70",
  "70-80",
  "80-90",
  "90-100",
  ">100",
];
const CLASS_TITLE_OPTIONS = [
  "Associate Induction",
  "Stay Ahead Event",
  "EAI Induction",
  "EAI Insight Event",
  "Others- Lead with Impact (EMEA only)",
  "Others- Advanced Client Presentation (EMEA only)",
];
const EDITOR_OPTIONS = [
  { value: "editor_a", label: "editor_a (Location_A)" },
  { value: "editor_b", label: "editor_b (Location_B)" },
];

function safeParseJson(s) {
  if (typeof s !== "string") return null;
  const trimmed = s.trim();
  if (!trimmed) return null;
  if (!trimmed.startsWith("{")) return null;
  try {
    return JSON.parse(trimmed);
  } catch {
    return null;
  }
}

function defaultSheetDetails() {
  return {
    mlCic: "ML",
    teamRole: "Team Member",
    personInCharge: EDITOR_OPTIONS[0]?.value || "",
    associateInduction: "",
    classSize: "0-10",
    modeOfDelivery: "F2F",
    cityName: "",
    quarter: "Q1",
    localLkSupported: false,
    localLkNames: "",
    localBusinessSupported: false,
    localBusinessNames: "",
    needFssSupport: false,
    fssSupportNames: "",
    notesText: "",
  };
}

function detailsFromEntry(entry) {
  const fallback = defaultSheetDetails();
  const parsed = safeParseJson(entry?.notes);
  if (!parsed) {
    const d = { ...fallback, notesText: typeof entry?.notes === "string" ? entry.notes : "" };
    if (typeof entry?.cohort === "string" && /^Q[1-4]$/.test(entry.cohort)) d.quarter = entry.cohort;
    return d;
  }
  const d = { ...fallback, ...parsed };
  if (typeof d.quarter !== "string" && typeof entry?.cohort === "string" && /^Q[1-4]$/.test(entry.cohort)) {
    d.quarter = entry.cohort;
  }
  return d;
}

function encodeDetails(details) {
  const compact = {
    mlCic: details.mlCic,
    teamRole: details.teamRole,
    personInCharge: details.personInCharge,
    associateInduction: details.associateInduction,
    classSize: details.classSize,
    modeOfDelivery: details.modeOfDelivery,
    cityName: details.cityName,
    quarter: details.quarter,
    localLkSupported: !!details.localLkSupported,
    localLkNames: (details.localLkNames || "").trim(),
    localBusinessSupported: !!details.localBusinessSupported,
    localBusinessNames: (details.localBusinessNames || "").trim(),
    needFssSupport: !!details.needFssSupport,
    fssSupportNames: (details.fssSupportNames || "").trim(),
    notesText: (details.notesText || "").trim(),
  };
  return JSON.stringify(compact);
}

function buildSummaryTitle(details) {
  const parts = [details.quarter, details.cityName, details.modeOfDelivery].filter(Boolean);
  return parts.length ? parts.join(" · ") : "ELH Class";
}

function commaNamesToArray(s) {
  if (!s || typeof s !== "string") return [];
  return s
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);
}

function arrayToCommaNames(arr) {
  return (arr || []).map((x) => String(x).trim()).filter(Boolean).join(", ");
}

function useSession() {
  const [session, setSession] = useState(() => {
    const t = getToken();
    const raw = localStorage.getItem("elh_session");
    if (t && raw) {
      try {
        return JSON.parse(raw);
      } catch {
        return null;
      }
    }
    return null;
  });

  const save = useCallback((payload) => {
    setToken(payload.token);
    const loc = payload.location ?? payload.geo ?? "";
    const s = {
      username: payload.username,
      role: payload.role,
      geo: loc,
      location: loc,
    };
    localStorage.setItem("elh_session", JSON.stringify(s));
    setSession(s);
  }, []);

  const logout = useCallback(() => {
    setToken(null);
    localStorage.removeItem("elh_session");
    setSession(null);
  }, []);

  return { session, save, logout };
}

function useTheme() {
  const [theme, setTheme] = useState(() => document.body.getAttribute("data-theme") || "dark");

  useEffect(() => {
    document.body.setAttribute("data-theme", theme);
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  }, [theme]);

  const toggleTheme = useCallback(() => {
    setTheme((t) => (t === "dark" ? "light" : "dark"));
  }, []);

  return { theme, toggleTheme };
}

function cmpDate(a, b) {
  if (a < b) return -1;
  if (a > b) return 1;
  return 0;
}

function sortRows(rows, sortKey, sortDir) {
  const dir = sortDir === "desc" ? -1 : 1;
  const list = [...rows];
  list.sort((a, b) => {
    let c = 0;
    if (sortKey === "date") c = cmpDate(a.startDate, b.startDate);
    else if (sortKey === "title") c = String(a.title).localeCompare(String(b.title));
    else if (sortKey === "location") c = String(a.geo).localeCompare(String(b.geo));
    return c * dir;
  });
  return list;
}

function ThemeToggleButton({ theme, onToggle }) {
  const isDark = theme === "dark";
  return (
    <button
      type="button"
      className="theme-toggle"
      onClick={onToggle}
      aria-label={isDark ? "Switch to light mode" : "Switch to dark mode"}
      title={isDark ? "Light mode" : "Dark mode"}
    >
      {isDark ? "☀" : "☾"}
    </button>
  );
}

export default function App() {
  const { session, save, logout } = useSession();
  const { theme, toggleTheme } = useTheme();

  const [username, setUsername] = useState("admin");
  const [password, setPassword] = useState("demo123");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const [entries, setEntries] = useState([]);
  const [geoFilter, setGeoFilter] = useState("");
  const [notifications, setNotifications] = useState([]);
  const [auditFor, setAuditFor] = useState(null);
  const [auditRows, setAuditRows] = useState([]);

  const [adminUsers, setAdminUsers] = useState([]);
  const [adminUsersBusy, setAdminUsersBusy] = useState(false);
  const [adminUserModalOpen, setAdminUserModalOpen] = useState(false);
  const [adminNewUser, setAdminNewUser] = useState({
    username: "",
    password: "",
    role: ROLES.USER_EDITOR,
    location: "",
  });
  const [localLkNewName, setLocalLkNewName] = useState("");
  const [localBusinessNewName, setLocalBusinessNewName] = useState("");
  const [fssNewName, setFssNewName] = useState("");

  const addLocalLkName = () => {
    const v = localLkNewName.trim();
    if (!v) return;
    const existing = commaNamesToArray(form.localLkNames);
    if (!existing.includes(v)) existing.push(v);
    setForm((f) => ({ ...f, localLkNames: arrayToCommaNames(existing) }));
    setLocalLkNewName("");
  };

  const removeLocalLkName = (name) => {
    const existing = commaNamesToArray(form.localLkNames).filter((x) => x !== name);
    setForm((f) => ({ ...f, localLkNames: arrayToCommaNames(existing) }));
  };

  const addLocalBusinessName = () => {
    const v = localBusinessNewName.trim();
    if (!v) return;
    const existing = commaNamesToArray(form.localBusinessNames);
    if (!existing.includes(v)) existing.push(v);
    setForm((f) => ({ ...f, localBusinessNames: arrayToCommaNames(existing) }));
    setLocalBusinessNewName("");
  };

  const removeLocalBusinessName = (name) => {
    const existing = commaNamesToArray(form.localBusinessNames).filter((x) => x !== name);
    setForm((f) => ({ ...f, localBusinessNames: arrayToCommaNames(existing) }));
  };

  const addFssSupportName = () => {
    const v = fssNewName.trim();
    if (!v) return;
    const existing = commaNamesToArray(form.fssSupportNames);
    if (!existing.includes(v)) existing.push(v);
    setForm((f) => ({ ...f, fssSupportNames: arrayToCommaNames(existing) }));
    setFssNewName("");
  };

  const removeFssSupportName = (name) => {
    const existing = commaNamesToArray(form.fssSupportNames).filter((x) => x !== name);
    setForm((f) => ({ ...f, fssSupportNames: arrayToCommaNames(existing) }));
  };

  const [mainTab, setMainTab] = useState("view");

  const [sortKey, setSortKey] = useState("title");
  const [sortDir, setSortDir] = useState("asc");
  const [filterDateFrom, setFilterDateFrom] = useState("");
  const [filterDateTo, setFilterDateTo] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [filterLocationClient, setFilterLocationClient] = useState("");

  const [sheetView, setSheetView] = useState("compact"); // compact | detailed

  const [form, setForm] = useState({
    startDate: "",
    endDate: "",
    geo: MARKET_OPTIONS[0],
    mlCic: "ML",
    teamRole: "Team Member",
    personInCharge: EDITOR_OPTIONS[0]?.value || "",
    associateInduction: "",
    classSize: "0-10",
    modeOfDelivery: "F2F",
    cityName: "",
    quarter: "Q1",
    localLkSupported: false,
    localLkNames: "",
    localBusinessSupported: false,
    localBusinessNames: "",
    needFssSupport: false,
    fssSupportNames: "",
    notesText: "",
  });
  const [editingId, setEditingId] = useState(null);

  const canWrite =
    session && (session.role === ROLES.USER_EDITOR || session.role === ROLES.CENTRAL_ADMIN);
  const isViewer = session && session.role === ROLES.USER_VIEWER;
  const isAdmin = session && session.role === ROLES.CENTRAL_ADMIN;
  const isEditor = session && session.role === ROLES.USER_EDITOR;
  const showNotifications = isAdmin;
  const showGeoFilter = isAdmin;
  const showLocationFilterView = isAdmin || isEditor;

  const userLocation = session?.location ?? session?.geo ?? "";

  const marketsForDropdown = useMemo(() => {
    const loc = userLocation?.trim();
    if (loc && !MARKET_OPTIONS.includes(loc)) return [loc, ...MARKET_OPTIONS];
    return MARKET_OPTIONS;
  }, [userLocation]);

  const editorOptionsForMarket = useMemo(() => {
    if (isEditor) {
      return [{ value: session.username, label: `${session.username}` }];
    }
    if (!form.geo) return [];
    if (!isAdmin) return [];
    const market = form.geo;
    return (adminUsers || [])
      .filter((u) => u.role === ROLES.USER_EDITOR && (u.location || "") === market)
      .map((u) => ({ value: u.username, label: u.username }));
  }, [adminUsers, form.geo, isAdmin, isEditor, session]);

  useEffect(() => {
    if (!isAdmin) return;
    setAdminUsersBusy(true);
    fetchAdminUsers()
      .then((u) => setAdminUsers(u || []))
      .catch((e) => setError(e.response?.data?.message || e.message || "Failed to load users"))
      .finally(() => setAdminUsersBusy(false));
  }, [isAdmin]);

  useEffect(() => {
    const onForbidden = (e) => {
      setError(e.detail?.message || "Permission denied");
    };
    window.addEventListener("elh-forbidden", onForbidden);
    return () => window.removeEventListener("elh-forbidden", onForbidden);
  }, []);

  const load = useCallback(async () => {
    if (!session) return;
    setBusy(true);
    setError("");
    try {
      const data = await fetchEntries(showGeoFilter && geoFilter ? geoFilter : undefined);
      setEntries(data);
      if (showNotifications) {
        const n = await fetchNotifications();
        setNotifications(n);
      }
    } catch (e) {
      setError(e.response?.data?.message || e.message || "Request failed");
      if (e.response?.status === 401) logout();
    } finally {
      setBusy(false);
    }
  }, [session, showGeoFilter, geoFilter, logout, showNotifications]);

  useEffect(() => {
    load();
  }, [load]);

  useEffect(() => {
    if (session?.role === ROLES.USER_EDITOR && userLocation) {
      setForm((f) => ({ ...f, geo: userLocation, personInCharge: session.username }));
      setFilterLocationClient(userLocation);
    }
  }, [session, userLocation]);

  // Clear city name when mode doesn't support it (Virtual or LVC)
  useEffect(() => {
    if (form.modeOfDelivery === "Virtual" || form.modeOfDelivery === "LVC") {
      if (form.cityName) {
        setForm((f) => ({ ...f, cityName: "" }));
      }
    }
  }, [form.modeOfDelivery]);

  useEffect(() => {
    if (!isAdmin) return;
    if (!editorOptionsForMarket || editorOptionsForMarket.length === 0) return;
    const current = form.personInCharge;
    if (!current || !editorOptionsForMarket.some((o) => o.value === current)) {
      setForm((f) => ({ ...f, personInCharge: editorOptionsForMarket[0].value }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAdmin, editorOptionsForMarket]);

  useEffect(() => {
    if (isViewer && mainTab === "add") {
      setMainTab("view");
    }
  }, [isViewer, mainTab]);

  const uniqueLocations = useMemo(() => {
    const s = new Set(entries.map((r) => r.geo).filter(Boolean));
    return [...s].sort();
  }, [entries]);

  const displayedEntries = useMemo(() => {
    let list = entries;
    if (filterDateFrom) {
      list = list.filter((r) => r.startDate >= filterDateFrom);
    }
    if (filterDateTo) {
      list = list.filter((r) => r.startDate <= filterDateTo);
    }
    if (filterStatus) {
      list = list.filter((r) => r.status === filterStatus);
    }
    if (showLocationFilterView && filterLocationClient) {
      list = list.filter((r) => r.geo === filterLocationClient);
    }
    return sortRows(list, sortKey, sortDir);
  }, [
    entries,
    filterDateFrom,
    filterDateTo,
    filterStatus,
    filterLocationClient,
    showLocationFilterView,
    sortKey,
    sortDir,
  ]);

  const displayedEntriesWithDetails = useMemo(() => {
    return displayedEntries.map((e) => ({
      entry: e,
      details: detailsFromEntry(e),
    }));
  }, [displayedEntries]);

  const onLogin = async (e) => {
    e.preventDefault();
    setBusy(true);
    setError("");
    try {
      const data = await login(username, password);
      save(data);
      setMainTab("view");
    } catch (err) {
      setError(err.response?.data?.message || "Login failed");
    } finally {
      setBusy(false);
    }
  };

  const onSaveEntry = async (e) => {
    e.preventDefault();
    if (!canWrite) return;
    setBusy(true);
    setError("");

    if (form.startDate && form.endDate && form.endDate <= form.startDate) {
      setError("End date must be after start date");
      setBusy(false);
      return;
    }

    const details = {
      mlCic: form.mlCic,
      teamRole: form.teamRole,
      personInCharge: form.personInCharge,
      associateInduction: form.associateInduction,
      classSize: form.classSize,
      modeOfDelivery: form.modeOfDelivery,
      cityName: form.cityName,
      quarter: form.quarter,
      localLkSupported: form.localLkSupported,
      localLkNames: form.localLkNames,
      localBusinessSupported: form.localBusinessSupported,
      localBusinessNames: form.localBusinessNames,
      needFssSupport: form.needFssSupport,
      fssSupportNames: form.fssSupportNames,
      notesText: form.notesText,
    };

    const body = {
      geo: form.geo,
      title: buildSummaryTitle(details),
      startDate: form.startDate,
      endDate: form.endDate,
      cohort: form.quarter,
      notes: encodeDetails(details),
      status: "SCHEDULED",
    };
    try {
      if (editingId) {
        await updateEntry(editingId, body);
      } else {
        await createEntry(body);
      }
      setForm((f) => ({
        ...f,
        startDate: "",
        endDate: "",
        associateInduction: "",
        cityName: "",
        localLkSupported: false,
        localLkNames: "",
        localBusinessSupported: false,
        localBusinessNames: "",
        needFssSupport: false,
        fssSupportNames: "",
        notesText: "",
      }));
      setEditingId(null);
      setMainTab("view");
      await load();
    } catch (err) {
      setError(err.response?.data?.message || err.message || "Save failed");
    } finally {
      setBusy(false);
    }
  };

  const onEdit = (row) => {
    setEditingId(row.id);
    const d = detailsFromEntry(row);
    setForm({
      geo: row.geo,
      startDate: row.startDate,
      endDate: row.endDate,
      mlCic: d.mlCic,
      teamRole: d.teamRole,
      personInCharge: d.personInCharge || (isEditor ? session.username : ""),
      associateInduction: d.associateInduction,
      classSize: d.classSize,
      modeOfDelivery: d.modeOfDelivery,
      cityName: d.cityName,
      quarter: d.quarter,
      localLkSupported: d.localLkSupported,
      localLkNames: d.localLkNames,
      localBusinessSupported: d.localBusinessSupported,
      localBusinessNames: d.localBusinessNames,
      needFssSupport: d.needFssSupport,
      fssSupportNames: d.fssSupportNames,
      notesText: d.notesText,
    });
    setMainTab("add");
  };

  const onCancelClass = async (id) => {
    if (!canWrite) return;
    if (!window.confirm("Mark this class as Cancelled? (Entries cannot be deleted.)")) return;
    setBusy(true);
    try {
      await cancelEntry(id);
      await load();
    } catch (err) {
      setError(err.response?.data?.message || err.message);
    } finally {
      setBusy(false);
    }
  };

  const onRevertCancellation = async (entry) => {
    if (!canWrite) return;
    if (!window.confirm("Reinstate this class?")) return;
    setBusy(true);
    setError("");
    try {
      // Update back to SCHEDULED (backend forbids only setting CANCELLED via update).
      await updateEntry(entry.id, {
        geo: entry.geo,
        title: entry.title,
        startDate: entry.startDate,
        endDate: entry.endDate,
        cohort: entry.cohort || null,
        notes: entry.notes || null,
        status: "SCHEDULED",
      });
      await load();
    } catch (err) {
      setError(err.response?.data?.message || err.message || "Revert failed");
    } finally {
      setBusy(false);
    }
  };

  const onExport = async () => {
    setBusy(true);
    setError("");
    try {
      const headers = [
        "Market",
        "ML/CIC",
        "EPH Program Team Member/Leader/LC",
        "Person in-charge",
        "Class Title",
        "Class Size",
        "Mode of Delivery",
        "City Name",
        "Quarter",
        "Start date",
        "End date",
        "Local L&K facilitators/L&K Members (names)",
        "Local Business facilitators (names)",
        "Need Facilitator support from FSS",
        "FSS Support names",
        "Notes",
      ];

      const aoa = [headers];
      for (const { entry, details } of displayedEntriesWithDetails) {
        const localLk = details.localLkSupported ? (details.localLkNames || "Yes") : "No";
        const localBiz = details.localBusinessSupported ? (details.localBusinessNames || "Yes") : "No";
        const needFss = details.needFssSupport ? "Yes" : "No";
        const fssNames = details.needFssSupport ? details.fssSupportNames || "" : "";

        aoa.push([
          entry.geo || "",
          details.mlCic || "",
          details.teamRole || "",
          details.personInCharge || "",
          details.associateInduction || "",
          details.classSize || "",
          details.modeOfDelivery || "",
          details.cityName || "",
          details.quarter || "",
          entry.startDate || "",
          entry.endDate || "",
          localLk,
          localBiz,
          needFss,
          fssNames,
          details.notesText || "",
        ]);
      }

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "ELH Calendar");
      const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });

      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "elh-calendar-detailed.xlsx";
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      setError(e?.message || "Export failed");
    } finally {
      setBusy(false);
    }
  };

  const openAudit = async (id) => {
    setAuditFor(id);
    setBusy(true);
    try {
      const rows = await fetchAudit(id);
      setAuditRows(rows);
    } catch (err) {
      setError(err.response?.data?.message || err.message);
    } finally {
      setBusy(false);
    }
  };

  const onCreateAdminUser = async (e) => {
    e.preventDefault();
    setAdminUsersBusy(true);
    setError("");
    try {
      await createAdminUser({
        username: adminNewUser.username.trim(),
        password: adminNewUser.password,
        role: adminNewUser.role,
        location: adminNewUser.role === ROLES.CENTRAL_ADMIN ? null : adminNewUser.location,
      });
      const u = await fetchAdminUsers();
      setAdminUsers(u || []);
      setAdminUserModalOpen(false);
    } catch (err) {
      setError(err.response?.data?.message || err.message || "Failed to add user");
    } finally {
      setAdminUsersBusy(false);
    }
  };

  const demoHint = useMemo(
    () => "Demo: editor_a / editor_c / editor_b / viewer_b / admin — password: demo123",
    []
  );

  const openAddTab = () => {
    setMainTab("add");
    if (!editingId) {
      setForm((f) => ({
        ...f,
        geo: session.role === ROLES.CENTRAL_ADMIN ? f.geo : userLocation || f.geo,
      }));
    }
  };

  if (!session) {
    return (
      <div className="app-layout">
        <header className="app-navbar app-navbar--simple">
          <div className="navbar-left">
            <span className="navbar-brand">ELH Calendar</span>
          </div>
          <div className="navbar-right">
            <ThemeToggleButton theme={theme} onToggle={toggleTheme} />
          </div>
        </header>
        <main className="app-main">
          <div className="shell">
            <header className="hero">
              <p className="eyebrow">Internal · ELH</p>
              <h1>Calendar management</h1>
              <p className="lede">
                Sign in to schedule classes by location. Central admins see all sites; editors one site; viewers are
                read-only. Audit trail and export available after login.
              </p>
              <p className="hint">{demoHint}</p>
            </header>
            <form className="card login" onSubmit={onLogin}>
              <h2>Sign in</h2>
              {error && <div className="banner error login-banner-error">{error}</div>}
              <label>
                Username
                <input value={username} onChange={(e) => setUsername(e.target.value)} autoComplete="username" />
              </label>
              <label>
                Password
                <input
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  autoComplete="current-password"
                />
              </label>
              <button type="submit" disabled={busy} className="primary">
                {busy ? "Signing in…" : "Continue"}
              </button>
            </form>
          </div>
        </main>
      </div>
    );
  }

  return (
    <div className="app-layout">
      <header className="app-navbar">
        <div className="navbar-left">
          <span className="navbar-brand">ELH Calendar</span>
          {canWrite && (
            <nav className="navbar-tabs" role="tablist" aria-label="Main sections">
              <button
                type="button"
                role="tab"
                id="tab-add"
                aria-selected={mainTab === "add"}
                className={`navbar-tab ${mainTab === "add" ? "active" : ""}`}
                onClick={openAddTab}
              >
                Add Class
              </button>
              <button
                type="button"
                role="tab"
                id="tab-view"
                aria-selected={mainTab === "view"}
                className={`navbar-tab ${mainTab === "view" ? "active" : ""}`}
                onClick={() => setMainTab("view")}
              >
                View Classes
              </button>
            </nav>
          )}
        </div>
        <div className="navbar-right">
          <span className="navbar-user">
            <strong>{session.username}</strong>
            <br />
            {session.role}
            {userLocation ? ` · ${userLocation}` : " · All locations"}
          </span>
          <ThemeToggleButton theme={theme} onToggle={toggleTheme} />
          <button type="button" className="btn-secondary" onClick={onExport} disabled={busy}>
            Export Excel
          </button>
          {isAdmin && (
            <button
              type="button"
              className="btn-secondary"
              onClick={() => {
                setAdminNewUser({
                  username: "",
                  password: "",
                  role: ROLES.USER_EDITOR,
                  location: userLocation || MARKET_OPTIONS[0],
                });
                setAdminUserModalOpen(true);
              }}
              disabled={adminUsersBusy}
              title="Create another user"
            >
              Add user
            </button>
          )}
          <button type="button" className="btn-secondary" onClick={load} disabled={busy}>
            Refresh
          </button>
          <button type="button" className="btn-secondary" onClick={logout}>
            Log out
          </button>
        </div>
      </header>

      {error && <div className="banner error">{error}</div>}

      <main className="app-main">
        <div className="shell wide">
          <div className="grid">
            {mainTab === "view" && (
              <section className="card">
                <div className="section-head">
                  <h3>Classes</h3>
                  {showGeoFilter && (
                    <label className="inline">
                      Market (server filter)
                      <input
                        placeholder="e.g. Americas"
                        value={geoFilter}
                        onChange={(e) => setGeoFilter(e.target.value)}
                      />
                    </label>
                  )}
                </div>

                <div className="filters-bar">
                  <button
                    type="button"
                    className="sort-class-btn"
                    onClick={() => setSortDir(sortDir === "asc" ? "desc" : "asc")}
                    title="Toggle sort by class name"
                  >
                    Class name
                    <span aria-hidden style={{ marginLeft: 6 }}>
                      {sortDir === "asc" ? "↑" : "↓"}
                    </span>
                  </button>
                  <span className="sheet-view-toggle" role="group" aria-label="Sheet view">
                    <button
                      type="button"
                      className={`sheet-view-btn ${sheetView === "compact" ? "active" : ""}`}
                      onClick={() => setSheetView("compact")}
                    >
                      Compact
                    </button>
                    <button
                      type="button"
                      className={`sheet-view-btn ${sheetView === "detailed" ? "active" : ""}`}
                      onClick={() => setSheetView("detailed")}
                    >
                      Detailed
                    </button>
                  </span>
                  <label className="inline small">
                    From
                    <input
                      type="date"
                      value={filterDateFrom}
                      onChange={(e) => setFilterDateFrom(e.target.value)}
                    />
                  </label>
                  <label className="inline small">
                    To
                    <input type="date" value={filterDateTo} onChange={(e) => setFilterDateTo(e.target.value)} />
                  </label>
                  <label className="inline small">
                    Status
                    <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                      <option value="">Any</option>
                      <option value="SCHEDULED">SCHEDULED</option>
                      <option value="CANCELLED">CANCELLED</option>
                    </select>
                  </label>
                  {showLocationFilterView && (
                    <label className="inline small">
                      Market
                      <select
                        value={filterLocationClient}
                        onChange={(e) => setFilterLocationClient(e.target.value)}
                        disabled={!isAdmin}
                      >
                        <option value="">Worldwide</option>
                        {uniqueLocations.map((loc) => (
                          <option key={loc} value={loc}>
                            {loc}
                          </option>
                        ))}
                      </select>
                    </label>
                  )}
                </div>

                <div className="table-wrap">
                  <table className="data-table">
                    <thead>
                      {sheetView === "compact" ? (
                        <tr>
                          <th>
                            Market {sortKey === "location" ? (sortDir === "asc" ? "↑" : "↓") : ""}
                          </th>
                          <th>ML/CIC</th>
                          <th>EPH Team</th>
                          <th>In-charge</th>
                          <th>Quarter</th>
                          <th>City</th>
                          <th>Mode</th>
                          <th>Class Size</th>
                          <th>
                            Start {sortKey === "date" ? (sortDir === "asc" ? "↑" : "↓") : ""}
                          </th>
                          <th>End</th>
                          <th>Status</th>
                          <th>Actions</th>
                        </tr>
                      ) : (
                        <tr>
                          <th>Market</th>
                          <th>ML/CIC</th>
                          <th>EPH Team</th>
                          <th>In-charge</th>
                          <th>Class Title</th>
                          <th>Class Size</th>
                          <th>Mode</th>
                          <th>City</th>
                          <th>Quarter</th>
                          <th>
                            Start {sortKey === "date" ? (sortDir === "asc" ? "↑" : "↓") : ""}
                          </th>
                          <th>End</th>
                          <th>Local L&K</th>
                          <th>Local Business</th>
                          <th>FSS support</th>
                          <th>FSS names</th>
                          <th>Notes</th>
                          <th>Status</th>
                          <th>Actions</th>
                        </tr>
                      )}
                    </thead>
                    <tbody>
                      {displayedEntriesWithDetails.map(({ entry, details }) => (
                        <tr key={entry.id}>
                          {sheetView === "compact" ? (
                            <>
                              <td>{entry.geo}</td>
                              <td>{details.mlCic}</td>
                              <td>{details.teamRole}</td>
                              <td>{details.personInCharge || "—"}</td>
                              <td>{details.quarter}</td>
                              <td>{details.cityName || "—"}</td>
                              <td>{details.modeOfDelivery}</td>
                              <td>{details.classSize}</td>
                              <td>{entry.startDate}</td>
                              <td>{entry.endDate}</td>
                              <td>
                                <span className={entry.status === "CANCELLED" ? "tag bad" : "tag ok"}>{entry.status}</span>
                              </td>
                              <td className="row-actions">
                                <button type="button" className="link" onClick={() => openAudit(entry.id)}>
                                  Audit
                                </button>
                                {canWrite && (isAdmin || details.personInCharge === session.username) && entry.status !== "CANCELLED" && (
                                  <>
                                    <button type="button" className="link" onClick={() => onEdit(entry)}>
                                      Edit
                                    </button>
                                    <button type="button" className="link danger" onClick={() => onCancelClass(entry.id)}>
                                      Cancel
                                    </button>
                                  </>
                                )}
                                {canWrite && (isAdmin || details.personInCharge === session.username) && entry.status === "CANCELLED" && (
                                  <button type="button" className="link" onClick={() => onRevertCancellation(entry)}>
                                    Undo cancel
                                  </button>
                                )}
                                {isViewer && entry.status !== "CANCELLED" && (
                                  <>
                                    <button type="button" className="link" disabled title="View-only access">
                                      Edit
                                    </button>
                                    <button type="button" className="link danger" disabled title="View-only access">
                                      Cancel
                                    </button>
                                  </>
                                )}
                              </td>
                            </>
                          ) : (
                            <>
                              <td>{entry.geo}</td>
                              <td>{details.mlCic}</td>
                              <td>{details.teamRole}</td>
                              <td>{details.personInCharge || "—"}</td>
                              <td>{details.associateInduction || "—"}</td>
                              <td>{details.classSize}</td>
                              <td>{details.modeOfDelivery}</td>
                              <td>{details.cityName || "—"}</td>
                              <td>{details.quarter}</td>
                              <td>{entry.startDate}</td>
                              <td>{entry.endDate}</td>
                              <td>{details.localLkSupported ? details.localLkNames || "Yes" : "No"}</td>
                              <td>{details.localBusinessSupported ? details.localBusinessNames || "Yes" : "No"}</td>
                              <td>{details.needFssSupport ? "Yes" : "No"}</td>
                              <td>{details.needFssSupport ? details.fssSupportNames || "—" : "—"}</td>
                              <td>{details.notesText || "—"}</td>
                              <td>
                                <span className={entry.status === "CANCELLED" ? "tag bad" : "tag ok"}>{entry.status}</span>
                              </td>
                              <td className="row-actions">
                                <button type="button" className="link" onClick={() => openAudit(entry.id)}>
                                  Audit
                                </button>
                                {canWrite && (isAdmin || details.personInCharge === session.username) && entry.status !== "CANCELLED" && (
                                  <>
                                    <button type="button" className="link" onClick={() => onEdit(entry)}>
                                      Edit
                                    </button>
                                    <button type="button" className="link danger" onClick={() => onCancelClass(entry.id)}>
                                      Cancel
                                    </button>
                                  </>
                                )}
                                {canWrite && (isAdmin || details.personInCharge === session.username) && entry.status === "CANCELLED" && (
                                  <button type="button" className="link" onClick={() => onRevertCancellation(entry)}>
                                    Undo cancel
                                  </button>
                                )}
                                {isViewer && entry.status !== "CANCELLED" && (
                                  <>
                                    <button type="button" className="link" disabled title="View-only access">
                                      Edit
                                    </button>
                                    <button type="button" className="link danger" disabled title="View-only access">
                                      Cancel
                                    </button>
                                  </>
                                )}
                              </td>
                            </>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {displayedEntries.length === 0 && (
                    <p className="muted table-empty-msg">No rows match filters.</p>
                  )}
                </div>
              </section>
            )}

            {mainTab === "add" && canWrite && (
              <section className="card form-card grid-main-full">
                <h3>{editingId ? `Edit entry #${editingId}` : "Add class"}</h3>
                <form onSubmit={onSaveEntry} className="form">
                  <label>
                    Market
                    <select
                      value={form.geo}
                      onChange={(e) => setForm({ ...form, geo: e.target.value })}
                      disabled={session.role === ROLES.USER_EDITOR}
                      required
                    >
                      {marketsForDropdown.map((m) => (
                        <option key={m} value={m}>
                          {m}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    ML/CIC
                    <div className="toggle-group">
                      {ML_CIC_OPTIONS.map((opt) => (
                        <button
                          type="button"
                          key={opt}
                          className={`toggle-option ${form.mlCic === opt ? "active" : ""}`}
                          onClick={() => setForm((f) => ({ ...f, mlCic: opt }))}
                        >
                          {opt}
                        </button>
                      ))}
                    </div>
                  </label>
                  <div className="two">
                    <label>
                      EPH Program Team Member/Leader/LC
                      <select value={form.teamRole} onChange={(e) => setForm({ ...form, teamRole: e.target.value })}>
                        {EPH_TEAM_ROLE_OPTIONS.map((opt) => (
                          <option key={opt} value={opt}>
                            {opt}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      Person in-charge (editor)
                      <select
                        value={form.personInCharge}
                        onChange={(e) => setForm({ ...form, personInCharge: e.target.value })}
                        disabled={isEditor}
                      >
                        {editorOptionsForMarket.length === 0 ? (
                          <option value="">No editors found</option>
                        ) : (
                          editorOptionsForMarket.map((u) => (
                            <option key={u.value} value={u.value}>
                              {u.label}
                            </option>
                          ))
                        )}
                      </select>
                    </label>
                  </div>
                  <label>
                    Class Title
                    <select
                      value={form.associateInduction}
                      onChange={(e) => setForm({ ...form, associateInduction: e.target.value })}
                    >
                      <option value="">Select a class title</option>
                      {CLASS_TITLE_OPTIONS.map((title) => (
                        <option key={title} value={title}>
                          {title}
                        </option>
                      ))}
                    </select>
                  </label>
                  <div className="two">
                    <label>
                      Class Size
                      <select value={form.classSize} onChange={(e) => setForm({ ...form, classSize: e.target.value })}>
                        {CLASS_SIZE_OPTIONS.map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      Mode of Delivery
                      <select
                        value={form.modeOfDelivery}
                        onChange={(e) => setForm({ ...form, modeOfDelivery: e.target.value })}
                      >
                        {MODE_OF_DELIVERY_OPTIONS.map((m) => (
                          <option key={m} value={m}>
                            {m}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                  <div className="two">
                    <label>
                      City Name
                      <input
                        value={form.cityName}
                        onChange={(e) => setForm({ ...form, cityName: e.target.value })}
                        disabled={!(form.modeOfDelivery === "F2F" || form.modeOfDelivery === "Hybrid")}
                        title={!(form.modeOfDelivery === "F2F" || form.modeOfDelivery === "Hybrid") ? "City name not applicable for Virtual/LVC" : ""}
                      />
                    </label>
                    <label>
                      Quarter
                      <select value={form.quarter} onChange={(e) => setForm({ ...form, quarter: e.target.value })}>
                        {QUARTER_OPTIONS.map((q) => (
                          <option key={q} value={q}>
                            {q}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                  <div className="two">
                    <label>
                      Start date
                      <input
                        type="date"
                        value={form.startDate}
                        onChange={(e) => setForm({ ...form, startDate: e.target.value })}
                        required
                      />
                    </label>
                    <label>
                      End date
                      <input
                        type="date"
                        value={form.endDate}
                        onChange={(e) => setForm({ ...form, endDate: e.target.value })}
                        required
                      />
                    </label>
                  </div>
                  <div className="checkbox-block">
                    <label className="checkbox-row">
                      <input
                        type="checkbox"
                        checked={form.localLkSupported}
                        onChange={(e) => setForm((f) => ({ ...f, localLkSupported: e.target.checked }))}
                      />
                      Do you have any local L&amp;K facilitators/L&amp;K Members to support the class?
                    </label>
                    {form.localLkSupported && (
                      <>
                        <div className="name-add-row">
                          <input
                            value={localLkNewName}
                            onChange={(e) => setLocalLkNewName(e.target.value)}
                            placeholder="Add a name"
                          />
                          <button type="button" className="btn-secondary" onClick={addLocalLkName}>
                            Add
                          </button>
                        </div>
                        <div className="muted small" style={{ marginTop: 6 }}>
                          {commaNamesToArray(form.localLkNames).length} added
                        </div>
                        <div className="chip-list">
                          {commaNamesToArray(form.localLkNames).map((n) => (
                            <button key={n} type="button" className="chip" onClick={() => removeLocalLkName(n)} title="Remove">
                              {n} ×
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                  <div className="checkbox-block">
                    <label className="checkbox-row">
                      <input
                        type="checkbox"
                        checked={form.localBusinessSupported}
                        onChange={(e) => setForm((f) => ({ ...f, localBusinessSupported: e.target.checked }))}
                      />
                      Do you have any local Business facilitators to support the class?
                    </label>
                    {form.localBusinessSupported && (
                      <>
                        <div className="name-add-row">
                          <input
                            value={localBusinessNewName}
                            onChange={(e) => setLocalBusinessNewName(e.target.value)}
                            placeholder="Add a name"
                          />
                          <button type="button" className="btn-secondary" onClick={addLocalBusinessName}>
                            Add
                          </button>
                        </div>
                        <div className="muted small" style={{ marginTop: 6 }}>
                          {commaNamesToArray(form.localBusinessNames).length} added
                        </div>
                        <div className="chip-list">
                          {commaNamesToArray(form.localBusinessNames).map((n) => (
                            <button key={n} type="button" className="chip" onClick={() => removeLocalBusinessName(n)} title="Remove">
                              {n} ×
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                  <div className="checkbox-block">
                    <label className="checkbox-row">
                      <input
                        type="checkbox"
                        checked={form.needFssSupport}
                        onChange={(e) => setForm((f) => ({ ...f, needFssSupport: e.target.checked }))}
                      />
                      Do you need Facilitator support from FSS?
                    </label>
                    {form.needFssSupport && (
                      <>
                        <div className="name-add-row">
                          <input value={fssNewName} onChange={(e) => setFssNewName(e.target.value)} placeholder="Add a name" />
                          <button type="button" className="btn-secondary" onClick={addFssSupportName}>
                            Add
                          </button>
                        </div>
                        <div className="muted small" style={{ marginTop: 6 }}>
                          {commaNamesToArray(form.fssSupportNames).length} added
                        </div>
                        <div className="chip-list">
                          {commaNamesToArray(form.fssSupportNames).map((n) => (
                            <button key={n} type="button" className="chip" onClick={() => removeFssSupportName(n)} title="Remove">
                              {n} ×
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                  <label>
                    Notes
                    <textarea
                      rows={3}
                      value={form.notesText}
                      onChange={(e) => setForm({ ...form, notesText: e.target.value })}
                    />
                  </label>
                  <div className="form-actions">
                    <button type="submit" className="primary" disabled={busy}>
                      {editingId ? "Save changes" : "Create entry"}
                    </button>
                    {editingId && (
                      <button
                        type="button"
                        className="ghost"
                        onClick={() => {
                          setEditingId(null);
                          setMainTab("view");
                        }}
                      >
                        Cancel edit
                      </button>
                    )}
                  </div>
                </form>
              </section>
            )}

            {showNotifications && mainTab === "view" && canWrite && (
              <section className="card">
                <h3>Admin notifications</h3>
                <p className="muted small">In-app feed when classes are created or updated.</p>
                <ul className="notify">
                  {notifications.map((n) => (
                    <li key={n.id}>
                      <span className="muted small">{new Date(n.createdAt).toLocaleString()}</span>
                      <div>{n.message}</div>
                    </li>
                  ))}
                </ul>
                {notifications.length === 0 && <p className="muted">No notifications yet.</p>}
              </section>
            )}
          </div>
        </div>
      </main>

      {auditFor && (
        <div className="modal-backdrop" role="presentation" onClick={() => setAuditFor(null)}>
          <div className="modal card" role="dialog" onClick={(e) => e.stopPropagation()}>
            <h3>Audit · entry #{auditFor}</h3>
            <ul className="audit">
              {auditRows.map((a) => (
                <li key={a.id}>
                  <div className="muted small">
                    {new Date(a.at).toLocaleString()} · {a.actor} · {a.action}
                  </div>
                  <div>{a.details}</div>
                </li>
              ))}
            </ul>
            <button type="button" className="primary" onClick={() => setAuditFor(null)}>
              Close
            </button>
          </div>
        </div>
      )}

      {adminUserModalOpen && (
        <div className="modal-backdrop" role="presentation" onClick={() => setAdminUserModalOpen(false)}>
          <div className="modal card" role="dialog" onClick={(e) => e.stopPropagation()}>
            <h3>Add user</h3>
            <form onSubmit={onCreateAdminUser} className="form">
              <label>
                Username
                <input
                  value={adminNewUser.username}
                  onChange={(e) => setAdminNewUser((u) => ({ ...u, username: e.target.value }))}
                  required
                  autoComplete="username"
                />
              </label>
              <label>
                Password
                <input
                  type="password"
                  value={adminNewUser.password}
                  onChange={(e) => setAdminNewUser((u) => ({ ...u, password: e.target.value }))}
                  required
                  autoComplete="new-password"
                />
              </label>
              <label>
                Role
                <select
                  value={adminNewUser.role}
                  onChange={(e) => setAdminNewUser((u) => ({ ...u, role: e.target.value }))}
                >
                  <option value={ROLES.USER_EDITOR}>USER_EDITOR</option>
                  <option value={ROLES.USER_VIEWER}>USER_VIEWER</option>
                  <option value={ROLES.CENTRAL_ADMIN}>CENTRAL_ADMIN</option>
                </select>
              </label>
              {adminNewUser.role !== ROLES.CENTRAL_ADMIN && (
                <label>
                  Location (market)
                  <select
                    value={adminNewUser.location}
                    onChange={(e) => setAdminNewUser((u) => ({ ...u, location: e.target.value }))}
                    required
                  >
                    {MARKET_OPTIONS.map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </label>
              )}
              <div className="form-actions">
                <button type="submit" className="primary" disabled={adminUsersBusy}>
                  {adminUsersBusy ? "Saving…" : "Create user"}
                </button>
                <button type="button" className="ghost" onClick={() => setAdminUserModalOpen(false)} disabled={adminUsersBusy}>
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
