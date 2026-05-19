import { useState, useEffect, useCallback, createContext, useContext, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area
} from "recharts";

// ─── THEME ────────────────────────────────────────────────────────────────────
const ThemeCtx = createContext({ dark: false, toggle: () => {} });
const useTheme = () => useContext(ThemeCtx);

// Primary brand: #004b23 (deep forest green) + white
const GREEN      = "#004b23";
const GREEN_MID  = "#005c2b";
const GREEN_LIGHT= "#006e34";
const GREEN_PALE = "#e8f5ee";

const T = {
  // ── Light: white content area, green sidebar ─────────────
  light: {
    bg:          "#F2F7F4",          // very pale green-white page bg
    surface:     "#FFFFFF",          // pure white cards
    surfaceAlt:  "#EAF3ED",          // pale green tint for alt rows
    surfaceHover:"#D6EBDE",          // hover state
    sidebar:     "#004b23",          // ← THE PANE colour
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.55)",
    sidebarBorder:"rgba(255,255,255,0.12)",
    sidebarActive:"rgba(255,255,255,0.15)",
    border:      "rgba(0,75,35,0.14)",
    borderStrong:"rgba(0,75,35,0.38)",
    text:        "#002A14",          // very dark green text
    textMuted:   "#3A6B4E",          // muted green
    textFaint:   "#A8C8B4",
    gold:        "#004b23",          // primary brand green
    goldLight:   "#006e34",
    goldDark:    "#002e16",
    success:     "#0A7A45",
    danger:      "#C0392B",
    info:        "#1565C0",
    warn:        "#C87A0A",
    navy:        "#002A14",
    btnFrom:     "#004b23",
    btnTo:       "#006e34",
    btnText:     "#FFFFFF",
  },
  // ── Dark: deep green + white text ─────────────────────────
  dark: {
    bg:          "#011a0d",
    surface:     "#012a14",
    surfaceAlt:  "#013d1e",
    surfaceHover:"#01531f",
    sidebar:     "#004b23",          // same green pane in dark
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.50)",
    sidebarBorder:"rgba(255,255,255,0.10)",
    sidebarActive:"rgba(255,255,255,0.14)",
    border:      "rgba(0,200,80,0.15)",
    borderStrong:"rgba(0,200,80,0.38)",
    text:        "#E8F5EE",
    textMuted:   "#7AB898",
    textFaint:   "#1A4028",
    gold:        "#4CD98A",          // bright green accent in dark
    goldLight:   "#80EDB0",
    goldDark:    "#00913A",
    success:     "#2ECC71",
    danger:      "#E74C3C",
    info:        "#4A9EDB",
    warn:        "#F39C12",
    navy:        "#011a0d",
    btnFrom:     "#004b23",
    btnTo:       "#00913A",
    btnText:     "#FFFFFF",
  },
};

// ─── RESPONSIVE HOOK ─────────────────────────────────────────────────────────
const useIsMobile = () => {
  const [mobile, setMobile] = useState(() => window.innerWidth < 768);
  useEffect(() => {
    const handler = () => setMobile(window.innerWidth < 768);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler);
  }, []);
  return mobile;
};

// Inject global responsive CSS once
const injectGlobalCSS = () => {
  if (document.getElementById("uct-responsive")) return;
  const style = document.createElement("style");
  style.id = "uct-responsive";
  style.textContent = `
    * { box-sizing: border-box; }
    body { margin: 0; overflow-x: hidden; }
    ::-webkit-scrollbar { width: 5px; height: 5px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #004b2344; border-radius: 4px; }
    .uct-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
    .uct-grid-3 { display: grid; grid-template-columns: repeat(3,1fr); gap: 14px; }
    .uct-grid-kpi { display: grid; grid-template-columns: repeat(auto-fit,minmax(145px,1fr)); gap:14px; }
    .uct-chart-row { display: grid; grid-template-columns: 1.5fr 1fr; gap: 20px; }
    .uct-form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .uct-table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    .uct-hide-mobile { display: flex; }
    .uct-show-mobile { display: none; }
    @media (max-width: 767px) {
      .uct-grid-2   { grid-template-columns: 1fr !important; }
      .uct-grid-3   { grid-template-columns: 1fr 1fr !important; }
      .uct-chart-row{ grid-template-columns: 1fr !important; }
      .uct-form-grid{ grid-template-columns: 1fr !important; }
      .uct-grid-kpi { grid-template-columns: repeat(2,1fr) !important; }
      .uct-hide-mobile { display: none !important; }
      .uct-show-mobile { display: flex !important; }
      .uct-full-col   { grid-column: 1 !important; }
      .uct-page-pad   { padding: 14px !important; }
      .uct-topbar-title { display: none; }
    }
  `;
  document.head.appendChild(style);
};
const WORKBOOK_KEY = "uct_ss_workbook";
const SHEETS = {
  RECORDS:  "Attendance Records",
  TEACHERS: "Teachers",
  CLASSES:  "Classes",
  CHURCH:   "Church Attendance",
  PROGRAMS: "Church Programs",
  META:     "Meta",
};

const RECORD_COLS = [
  "id","date","day_of_week","service_type","class_name","teacher_name","assistant_teacher",
  "submitted_by","time_started","time_ended","total_beginning","total_closing",
  "male_present","female_present","first_timers","visitors","absent_members",
  "bibles_beginning","bibles_closing","members_without_bibles","topic",
  "bible_references","memory_verse","key_notes","assignment","teachers_present",
  "teacher_names","challenges","prayer_requests","announcements","status","created_at"
];

const TEACHER_COLS = ["id","name","phone","email","class_name","role","is_active","joined_date","notes"];

const CHURCH_COLS = [
  "id","date","day_of_week","program",
  "male_beginning","female_beginning","total_beginning",
  "male_closing","female_closing","total_closing",
  "first_timers","visitors","notes","submitted_by","created_at"
];

const PROGRAM_COLS = ["id","name","is_active","sort_order"];

const USER_COLS = ["id","name","email","password","role","is_active","permissions","created_at"];

// All pages / features that can be individually toggled per user
const ALL_PERMISSIONS = [
  { key:"dashboard",  label:"Dashboard",            group:"Navigation" },
  { key:"attendance", label:"SS Attendance Records",group:"Navigation" },
  { key:"church",     label:"Church Attendance",    group:"Navigation" },
  { key:"analytics",  label:"Analytics & Trends",   group:"Navigation" },
  { key:"teachers",   label:"Teacher Management",   group:"Navigation" },
  { key:"classes",    label:"Class Management",     group:"Navigation" },
  { key:"programs",   label:"Programs Management",  group:"Navigation" },
  { key:"export",     label:"Export Center",        group:"Navigation" },
  { key:"submit",     label:"Submit SS Report",     group:"Navigation" },
  { key:"approve_records",  label:"Approve / Reject Records",   group:"Actions" },
  { key:"delete_records",   label:"Delete Records",              group:"Actions" },
  { key:"view_analytics",   label:"View Analytics Data",         group:"Actions" },
  { key:"view_church_data", label:"View Church Attendance",      group:"Actions" },
  { key:"view_export",      label:"Download / Export Data",      group:"Actions" },
  { key:"view_all_classes", label:"See All Classes' Records",    group:"Actions" },
  { key:"manage_teachers",  label:"Add / Edit / Delete Teachers",group:"Actions" },
];

// Default full access for admin
const ADMIN_PERMS = ALL_PERMISSIONS.map(p=>p.key);
// Default limited access for teacher
const TEACHER_PERMS_DEFAULT = ["submit","attendance","church","view_church_data"];

// ── Built-in roles (always available) ────────────────────────
const BUILT_IN_ROLES = [
  { id:"admin",          name:"Admin",          color:"#004b23", permissions: ADMIN_PERMS,           isBuiltIn: true,  description:"Full access to everything" },
  { id:"teacher",        name:"Teacher",        color:"#1565C0", permissions: TEACHER_PERMS_DEFAULT, isBuiltIn: true,  description:"Submit reports and view own records" },
  { id:"superintendent", name:"Superintendent", color:"#9B59B6", permissions: [...TEACHER_PERMS_DEFAULT,"analytics","view_analytics","dashboard","export","view_export","view_all_classes"], isBuiltIn: true, description:"Teacher + analytics and exports" },
  { id:"assistant",      name:"Asst. Teacher",  color:"#E67E22", permissions: ["submit","church","view_church_data"], isBuiltIn: true, description:"Submit SS reports and church records only" },
];

const ROLES_KEY = "uct_custom_roles_v1";

// Load custom roles from localStorage (synced separately)
const loadCustomRoles = () => {
  try {
    const raw = localStorage.getItem(ROLES_KEY);
    if (raw) { const p = JSON.parse(raw); if (Array.isArray(p)) return p; }
  } catch {}
  return [];
};
const saveCustomRoles = (roles) => {
  try { localStorage.setItem(ROLES_KEY, JSON.stringify(roles)); } catch {}
};

// All roles = built-in + custom
const getAllRoles = () => [...BUILT_IN_ROLES, ...loadCustomRoles()];

const USERS_KEY = "uct_ss_users";

// ─── BRANDING CONFIG ──────────────────────────────────────────────────────────
// Branding is loaded from localStorage (set via the Logo & Name page).
// On first run, defaults to "Unique Christian Tabernacle".
// To hard-code defaults, edit the fallback strings below.
const _savedBrand = (() => { try { return JSON.parse(localStorage.getItem("uct_branding")||"{}"); } catch { return {}; } })();
const CHURCH_NAME  = _savedBrand.name  || "Unique Christian Tabernacle";
const CHURCH_SHORT = _savedBrand.short || "UCT";
const CHURCH_SUB   = _savedBrand.sub   || "Sunday School Management System";
const ACTIVE_LOGO  = _savedBrand.logo  || null;

// Logo component — shows image if provided, cross icon otherwise
const ChurchLogo = ({ size = 36, border = true }) => {
  const style = {
    width: size, height: size, borderRadius: "50%", flexShrink: 0,
    display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden",
    background: ACTIVE_LOGO ? "transparent" : "rgba(255,255,255,0.18)",
    border: border ? "2px solid rgba(255,255,255,0.35)" : "none",
  };
  if (ACTIVE_LOGO) {
    return (
      <div style={style}>
        <img src={ACTIVE_LOGO} alt={CHURCH_SHORT}
          style={{ width: "100%", height: "100%", objectFit: "cover", borderRadius: "50%" }} />
      </div>
    );
  }
  return (
    <div style={style}>
      <Icon name="cross" size={Math.round(size * 0.47)} color="#FFFFFF" />
    </div>
  );
};

// Settings page for admin to change name / upload logo
const BrandingSettings = ({ onClose }) => {
  const { t, inp, lbl, btnGold, btnOutline, card } = useThemeStyles();
  const [name,  setName]  = useState(CHURCH_NAME);
  const [short, setShort] = useState(CHURCH_SHORT);
  const [sub,   setSub]   = useState(CHURCH_SUB);
  const [logo,  setLogo]  = useState(ACTIVE_LOGO || "");
  const [preview, setPreview] = useState(ACTIVE_LOGO || "");
  const fileRef = useRef();

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => { setLogo(ev.target.result); setPreview(ev.target.result); };
    reader.readAsDataURL(file);
  };

  const handleSave = () => {
    // Persist to localStorage so it survives refresh
    try {
      localStorage.setItem("uct_branding", JSON.stringify({ name, short, sub, logo }));
      // Reload so the constants pick up the new values
      window.location.reload();
    } catch { alert("Saved! Reload the page to see changes."); }
  };

  return (
    <div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:20 }}>
        <div style={{ display:"flex", flexDirection:"column", gap:5, gridColumn:"1 / -1" }}>
          <label style={lbl}>Church / Organisation Name *</label>
          <input style={inp} value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. Grace Bible Church" />
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={lbl}>Short / Abbreviation</label>
          <input style={inp} value={short} onChange={e=>setShort(e.target.value)} placeholder="e.g. GBC" maxLength={8} />
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={lbl}>Subtitle / System Name</label>
          <input style={inp} value={sub} onChange={e=>setSub(e.target.value)} placeholder="e.g. Attendance Management" />
        </div>
      </div>

      {/* Logo upload */}
      <div style={{ marginBottom:20 }}>
        <label style={{ ...lbl, display:"block", marginBottom:8 }}>Church Logo</label>
        <div style={{ display:"flex", alignItems:"center", gap:16 }}>
          {/* Preview circle */}
          <div style={{ width:72, height:72, borderRadius:"50%", background:"#004b23",
            display:"flex", alignItems:"center", justifyContent:"center",
            border:"3px solid rgba(0,75,35,0.3)", overflow:"hidden", flexShrink:0 }}>
            {preview
              ? <img src={preview} alt="logo" style={{ width:"100%", height:"100%", objectFit:"cover", borderRadius:"50%" }} />
              : <Icon name="cross" size={30} color="#FFFFFF" />}
          </div>
          <div style={{ flex:1 }}>
            <div style={{ marginBottom:8 }}>
              <label style={{ ...lbl, marginBottom:5, display:"block" }}>Logo URL (https://…)</label>
              <input style={{ ...inp }} value={logo} onChange={e=>{ setLogo(e.target.value); setPreview(e.target.value); }}
                placeholder="https://yourchurch.org/logo.png" />
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>— or —</span>
              <input ref={fileRef} type="file" accept="image/*" style={{ display:"none" }} onChange={handleFile} />
              <button style={{ ...btnOutline, padding:"7px 14px", fontSize:12 }} onClick={()=>fileRef.current.click()}>
                Upload Image File
              </button>
              {preview && (
                <button onClick={()=>{ setLogo(""); setPreview(""); }} style={{ fontSize:11, color:t.danger, background:"transparent", border:`1px solid ${t.danger}33`, borderRadius:6, padding:"6px 10px", cursor:"pointer", fontFamily:"'Trebuchet MS',sans-serif" }}>
                  Remove
                </button>
              )}
            </div>
            <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:6 }}>
              Recommended: square image, PNG or JPG, at least 128×128px
            </div>
          </div>
        </div>
      </div>

      <div style={{ padding:"11px 14px", background:t.surfaceAlt, borderRadius:9, border:`1px solid ${t.border}`, marginBottom:20 }}>
        <div style={{ fontSize:12, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", fontWeight:700, marginBottom:4 }}>Preview</div>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:40, height:40, borderRadius:"50%", background:"#004b23",
            display:"flex", alignItems:"center", justifyContent:"center", overflow:"hidden", border:"2px solid rgba(0,75,35,0.3)" }}>
            {preview ? <img src={preview} alt="p" style={{ width:"100%", height:"100%", objectFit:"cover", borderRadius:"50%" }} />
                     : <Icon name="cross" size={18} color="#FFFFFF" />}
          </div>
          <div>
            <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>{name || "Church Name"}</div>
            <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>{sub || "Subtitle"}</div>
          </div>
        </div>
      </div>

      <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
        <button style={{ ...btnOutline, padding:"10px 20px" }} onClick={onClose}>Cancel</button>
        <button style={{ ...btnGold, padding:"10px 24px" }} onClick={handleSave}>Save & Reload</button>
      </div>
    </div>
  );
};

// Encode workbook to base64 and save to localStorage
const saveWorkbook = (wb) => {
  const raw = XLSX.write(wb, { type: "base64", bookType: "xlsx" });
  try { localStorage.setItem(WORKBOOK_KEY, raw); } catch {}
};

// Load workbook from localStorage or create fresh
const loadWorkbook = () => {
  try {
    const raw = localStorage.getItem(WORKBOOK_KEY);
    if (raw) return XLSX.read(raw, { type: "base64" });
  } catch {}
  return null;
};

// ── SHARED USER STORAGE via Supabase ─────────────────────────
// Users are stored in Supabase so ALL devices see the same data.
// Set your Supabase URL and anon key below after creating a free project.
// Until configured, falls back to localStorage (single-device only).

const SUPABASE_URL   = "https://phnqknoritzmgcjzsgic.supabase.co";
const SUPABASE_ANON  = "sb_publishable_mDTovCRDkK_7WUyM3amo_w_mKwyUFjD";

const SUPABASE_READY = Boolean(SUPABASE_URL && SUPABASE_ANON);

// ── Supabase REST helpers (no SDK needed) ─────────────────────
const sbFetch = async (path, opts = {}) => {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...opts,
    headers: {
      "Content-Type":  "application/json",
      "apikey":        SUPABASE_ANON,
      "Authorization": `Bearer ${SUPABASE_ANON}`,
      "Prefer":        "return=representation",
      ...(opts.headers || {}),
    },
  });
  if (!res.ok) { const e = await res.text(); throw new Error(e); }
  const txt = await res.text();
  return txt ? JSON.parse(txt) : [];
};

// ── Fallback localStorage (used until Supabase is configured) ─
const USERS_VERSION    = "v4";
const USERS_KEY_VERSIONED = `${USERS_KEY}_${USERS_VERSION}`;

const DEFAULT_USERS = [
  { id:"U001", name:"System Admin",      email:"admin@uct.org",    password:"admin123",  role:"admin",   is_active:"YES", permissions: JSON.stringify(ADMIN_PERMS),           created_at:"2024-01-01" },
  { id:"U002", name:"Bro. Emmanuel Adu", email:"emmanuel@uct.org", password:"teacher1",  role:"teacher", is_active:"YES", permissions: JSON.stringify(TEACHER_PERMS_DEFAULT), created_at:"2024-01-01" },
  { id:"U003", name:"Sis. Grace Mensah", email:"grace@uct.org",    password:"teacher2",  role:"teacher", is_active:"YES", permissions: JSON.stringify(TEACHER_PERMS_DEFAULT), created_at:"2024-01-01" },
  { id:"U004", name:"Bro. Samuel Ofori", email:"samuel@uct.org",   password:"teacher3",  role:"teacher", is_active:"YES", permissions: JSON.stringify(TEACHER_PERMS_DEFAULT), created_at:"2024-01-01" },
  { id:"U005", name:"Superintendent",    email:"super@uct.org",    password:"super123",  role:"teacher", is_active:"YES", permissions: JSON.stringify([...TEACHER_PERMS_DEFAULT,"analytics","view_analytics","dashboard","export","view_export"]), created_at:"2024-01-01" },
];

const loadUsersLocal = () => {
  try {
    const raw = localStorage.getItem(USERS_KEY_VERSIONED);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    }
  } catch {}
  const seed = JSON.parse(JSON.stringify(DEFAULT_USERS));
  localStorage.setItem(USERS_KEY_VERSIONED, JSON.stringify(seed));
  return seed;
};

const saveUsersLocal = (users) => {
  localStorage.setItem(USERS_KEY_VERSIONED, JSON.stringify(users));
};

// loadUsers — async when Supabase ready, sync fallback otherwise
const loadUsers = () => loadUsersLocal();  // sync wrapper for login

const parsePerms = (p) => { try { return JSON.parse(p || "[]"); } catch { return []; } };

// ── useUsers hook — handles both Supabase + localStorage ──────
const useUsers = () => {
  const [appUsers, setAppUsers] = useState([]);
  const [usersReady, setUsersReady] = useState(false);

  // Load on mount
  useEffect(() => {
    const fetchUsers = async () => {
      if (SUPABASE_READY) {
        try {
          const rows = await sbFetch("uct_users?select=*&order=created_at.asc");
          if (rows.length > 0) {
            setAppUsers(rows);
            // Mirror to localStorage as cache
            saveUsersLocal(rows);
            setUsersReady(true);
            return;
          }
          // Empty table — seed defaults
          for (const u of DEFAULT_USERS) {
            await sbFetch("uct_users", { method:"POST", body: JSON.stringify(u) });
          }
          setAppUsers([...DEFAULT_USERS]);
          saveUsersLocal([...DEFAULT_USERS]);
        } catch (err) {
          console.warn("Supabase unavailable, falling back to localStorage:", err.message);
          setAppUsers(loadUsersLocal());
        }
      } else {
        setAppUsers(loadUsersLocal());
      }
      setUsersReady(true);
    };
    fetchUsers();
  }, []);

  const saveAll = useCallback(async (users) => {
    setAppUsers([...users]);
    saveUsersLocal(users);
    // No Supabase sync needed here — individual operations handle it
  }, []);

  const addUser = useCallback(async (data) => {
    if (!data.name?.trim())     return { error: "Name is required." };
    if (!data.email?.trim())    return { error: "Email is required." };
    if (!data.password?.trim()) return { error: "Password is required." };

    const current = SUPABASE_READY ? appUsers : loadUsersLocal();
    const exists  = current.find(u => u.email.toLowerCase() === data.email.toLowerCase().trim());
    if (exists) return { error: "An account with this email already exists." };

    const newU = {
      id:          `U${Date.now()}`,
      name:        data.name.trim(),
      email:       data.email.toLowerCase().trim(),
      password:    data.password.trim(),
      role:        data.role || "teacher",
      is_active:   "YES",
      permissions: JSON.stringify(Array.isArray(data.permissions) ? data.permissions : TEACHER_PERMS_DEFAULT),
      created_at:  new Date().toISOString().slice(0,10),
    };

    if (SUPABASE_READY) {
      try { await sbFetch("uct_users", { method:"POST", body: JSON.stringify(newU) }); }
      catch(e) { return { error: `Save failed: ${e.message}` }; }
    }

    const updated = [...current, newU];
    saveUsersLocal(updated);
    setAppUsers([...updated]);
    return { success: true, user: newU };
  }, [appUsers]);

  const updateUser = useCallback(async (id, updates) => {
    const current = SUPABASE_READY ? appUsers : loadUsersLocal();
    const updated = current.map(u => {
      if (u.id !== id) return u;
      const merged = { ...u };
      if (updates.name)                    merged.name      = updates.name.trim();
      if (updates.email)                   merged.email     = updates.email.toLowerCase().trim();
      if (updates.role)                    merged.role      = updates.role;
      if (updates.is_active !== undefined) merged.is_active = updates.is_active;
      if (updates.password?.trim())        merged.password  = updates.password.trim();
      if (updates.permissions !== undefined) {
        merged.permissions = Array.isArray(updates.permissions)
          ? JSON.stringify(updates.permissions)
          : updates.permissions;
      }
      return merged;
    });

    if (SUPABASE_READY) {
      const changed = updated.find(u => u.id === id);
      try {
        await sbFetch(`uct_users?id=eq.${id}`, {
          method: "PATCH",
          body:   JSON.stringify(changed),
        });
      } catch(e) { alert(`Update failed: ${e.message}`); return; }
    }

    saveUsersLocal(updated);
    setAppUsers([...updated]);
  }, [appUsers]);

  const deleteUser = useCallback(async (id) => {
    if (SUPABASE_READY) {
      try { await sbFetch(`uct_users?id=eq.${id}`, { method:"DELETE" }); }
      catch(e) { alert(`Delete failed: ${e.message}`); return; }
    }
    const updated = loadUsersLocal().filter(u => u.id !== id);
    saveUsersLocal(updated);
    setAppUsers([...updated]);
  }, []);

  const toggleUserActive = useCallback(async (id) => {
    const current = SUPABASE_READY ? appUsers : loadUsersLocal();
    const user    = current.find(u => u.id === id);
    if (user) await updateUser(id, { is_active: user.is_active === "YES" ? "NO" : "YES" });
  }, [appUsers, updateUser]);

  return { appUsers, usersReady, addUser, updateUser, deleteUser, toggleUserActive };
};

// Parse a sheet into array of objects
const sheetToRows = (wb, sheetName) => {
  if (!wb.Sheets[sheetName]) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
};

// Write array of objects into a sheet
const rowsToSheet = (wb, sheetName, rows, cols) => {
  const ws = XLSX.utils.json_to_sheet(rows, { header: cols });
  ws["!cols"] = cols.map(c => ({ wch: Math.max(c.length + 2, 14) }));
  wb.Sheets[sheetName] = ws;
  if (!wb.SheetNames.includes(sheetName)) wb.SheetNames.push(sheetName);
};

const initWorkbook = () => {
  const wb = XLSX.utils.book_new();
  const defaultClasses = [
    "Children (5–9)","Teen (10–15)","New Convert",
    "Prophet Class","Church Age","C.O.D","Joint Sunday School"
  ];
  const defaultTeachers = [
    { id:"T001", name:"Bro. Emmanuel Adu",    phone:"", email:"", class_name:"Children (5–9)",  role:"Teacher",       is_active:"YES", joined_date:"2023-01-01", notes:"" },
    { id:"T002", name:"Sis. Grace Mensah",    phone:"", email:"", class_name:"Teen (10–15)",    role:"Teacher",       is_active:"YES", joined_date:"2023-01-01", notes:"" },
    { id:"T003", name:"Bro. Kweku Asante",    phone:"", email:"", class_name:"New Convert",     role:"Teacher",       is_active:"YES", joined_date:"2023-01-01", notes:"" },
    { id:"T004", name:"Sis. Abena Osei",      phone:"", email:"", class_name:"Prophet Class",   role:"Teacher",       is_active:"YES", joined_date:"2023-01-01", notes:"" },
    { id:"T005", name:"Bro. Samuel Ofori",    phone:"", email:"", class_name:"Church Age",      role:"Teacher",       is_active:"YES", joined_date:"2023-01-01", notes:"" },
    { id:"T006", name:"Sis. Priscilla Boateng",phone:"",email:"", class_name:"C.O.D",           role:"Asst. Teacher", is_active:"YES", joined_date:"2023-01-01", notes:"" },
    { id:"T007", name:"Bro. Daniel Tetteh",   phone:"", email:"", class_name:"Church Age",      role:"Teacher",       is_active:"YES", joined_date:"2023-01-01", notes:"" },
  ];
  const defaultPrograms = [
    { id:"P001", name:"Sunday Church Service",       is_active:"YES", sort_order:1 },
    { id:"P002", name:"Wednesday Prayer Meeting",    is_active:"YES", sort_order:2 },
    { id:"P003", name:"Friday Bible Study",          is_active:"YES", sort_order:3 },
    { id:"P004", name:"Sunday Evening Service",      is_active:"YES", sort_order:4 },
    { id:"P005", name:"Special Revival Service",     is_active:"YES", sort_order:5 },
    { id:"P006", name:"Wedding Occasion",            is_active:"YES", sort_order:6 },
    { id:"P007", name:"Funeral / Memorial Service",  is_active:"YES", sort_order:7 },
    { id:"P008", name:"Harvest / Thanksgiving",      is_active:"YES", sort_order:8 },
    { id:"P009", name:"Dedication Service",          is_active:"YES", sort_order:9 },
    { id:"P010", name:"Outreach / Evangelism",       is_active:"YES", sort_order:10 },
    { id:"P011", name:"Youth Service",               is_active:"YES", sort_order:11 },
    { id:"P012", name:"Men's Fellowship",            is_active:"YES", sort_order:12 },
    { id:"P013", name:"Women's Fellowship",          is_active:"YES", sort_order:13 },
    { id:"P014", name:"Choir / Praise Night",        is_active:"YES", sort_order:14 },
    { id:"P015", name:"Communion Service",           is_active:"YES", sort_order:15 },
  ];
  rowsToSheet(wb, SHEETS.RECORDS,  [], RECORD_COLS);
  rowsToSheet(wb, SHEETS.TEACHERS, defaultTeachers, TEACHER_COLS);
  rowsToSheet(wb, SHEETS.CLASSES,  defaultClasses.map((n,i)=>({ id:i+1, name:n, is_active:"YES" })), ["id","name","is_active"]);
  rowsToSheet(wb, SHEETS.CHURCH,   [], CHURCH_COLS);
  rowsToSheet(wb, SHEETS.PROGRAMS, defaultPrograms, PROGRAM_COLS);
  rowsToSheet(wb, SHEETS.META,     [{ key:"created", value: new Date().toISOString() }], ["key","value"]);
  saveWorkbook(wb);
  return wb;
};


// ─── SUPABASE DATA HOOK ───────────────────────────────────────────────────────
const useSupabaseDB = () => {
  const [records,    setRecords]    = useState([]);
  const [teachers,   setTeachers]   = useState([]);
  const [classes,    setClasses]    = useState([]);
  const [churchRecs, setChurchRecs] = useState([]);
  const [programs,   setPrograms]   = useState([]);
  const [loading,    setLoading]    = useState(true);

  // ── Generic fetch ────────────────────────────────────────
  const sbGet = useCallback(async (table, order = "created_at.asc") => {
    try {
      const rows = await sbFetch(`${table}?select=*&order=${order}`);
      return Array.isArray(rows) ? rows : [];
    } catch (e) {
      console.error(`sbGet ${table}:`, e.message);
      return [];
    }
  }, []);

  // ── Load all data ─────────────────────────────────────────
  const loadAll = useCallback(async () => {
    setLoading(true);
    const [r, t, c, ch, p] = await Promise.all([
      sbGet("uct_records",  "created_at.desc"),
      sbGet("uct_teachers", "name.asc"),
      sbGet("uct_classes",  "id.asc"),
      sbGet("uct_church",   "created_at.desc"),
      sbGet("uct_programs", "sort_order.asc"),
    ]);
    setRecords(r);
    setTeachers(t);
    setClasses(c.length ? c : [
      {id:"C1",name:"Children (5–9)",is_active:"YES"},
      {id:"C2",name:"Teen (10–15)",is_active:"YES"},
      {id:"C3",name:"New Convert",is_active:"YES"},
      {id:"C4",name:"Prophet Class",is_active:"YES"},
      {id:"C5",name:"Church Age",is_active:"YES"},
      {id:"C6",name:"C.O.D",is_active:"YES"},
      {id:"C7",name:"Joint Sunday School",is_active:"YES"},
    ]);
    setChurchRecs(ch);
    setPrograms(p.length ? p : [
      {id:"P001",name:"Sunday Church Service",is_active:"YES",sort_order:"1"},
      {id:"P002",name:"Wednesday Prayer Meeting",is_active:"YES",sort_order:"2"},
      {id:"P003",name:"Friday Bible Study",is_active:"YES",sort_order:"3"},
      {id:"P004",name:"Sunday Evening Service",is_active:"YES",sort_order:"4"},
      {id:"P005",name:"Special Revival Service",is_active:"YES",sort_order:"5"},
      {id:"P006",name:"Wedding Occasion",is_active:"YES",sort_order:"6"},
      {id:"P007",name:"Funeral / Memorial Service",is_active:"YES",sort_order:"7"},
      {id:"P008",name:"Harvest / Thanksgiving",is_active:"YES",sort_order:"8"},
      {id:"P009",name:"Dedication Service",is_active:"YES",sort_order:"9"},
      {id:"P010",name:"Outreach / Evangelism",is_active:"YES",sort_order:"10"},
      {id:"P011",name:"Youth Service",is_active:"YES",sort_order:"11"},
      {id:"P012",name:"Men's Fellowship",is_active:"YES",sort_order:"12"},
      {id:"P013",name:"Women's Fellowship",is_active:"YES",sort_order:"13"},
      {id:"P014",name:"Choir / Praise Night",is_active:"YES",sort_order:"14"},
      {id:"P015",name:"Communion Service",is_active:"YES",sort_order:"15"},
    ]);
    setLoading(false);
  }, [sbGet]);

  useEffect(() => { loadAll(); }, [loadAll]);

  // ── SS Records ───────────────────────────────────────────
  const addRecord = useCallback(async (rec) => {
    const newRec = { ...rec, id:`R${Date.now()}`, status:"pending", created_at:new Date().toISOString() };
    try {
      await sbFetch("uct_records", { method:"POST", body:JSON.stringify(newRec) });
      setRecords(r => [newRec, ...r]);
    } catch(e) { alert("Save failed: " + e.message); }
    return newRec;
  }, []);

  const updateRecord = useCallback(async (id, updates) => {
    try {
      await sbFetch(`uct_records?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(updates) });
      setRecords(r => r.map(x => x.id===id ? {...x,...updates} : x));
    } catch(e) { alert("Update failed: " + e.message); }
  }, []);

  const deleteRecord = useCallback(async (id) => {
    try {
      await sbFetch(`uct_records?id=eq.${id}`, { method:"DELETE" });
      setRecords(r => r.filter(x => x.id!==id));
    } catch(e) { alert("Delete failed: " + e.message); }
  }, []);

  const approveRecord = useCallback((id) => updateRecord(id, { status:"approved" }), [updateRecord]);

  // ── Teachers ─────────────────────────────────────────────
  const addTeacher = useCallback(async (data) => {
    const newT = { ...data, id:`T${Date.now()}`, is_active:"YES",
      joined_date: data.joined_date || new Date().toISOString().slice(0,10),
      created_at: new Date().toISOString() };
    try {
      await sbFetch("uct_teachers", { method:"POST", body:JSON.stringify(newT) });
      setTeachers(t => [...t, newT].sort((a,b)=>a.name.localeCompare(b.name)));
    } catch(e) { alert("Save failed: " + e.message); }
  }, []);

  const updateTeacher = useCallback(async (id, updates) => {
    try {
      await sbFetch(`uct_teachers?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(updates) });
      setTeachers(t => t.map(x => x.id===id ? {...x,...updates} : x));
    } catch(e) { alert("Update failed: " + e.message); }
  }, []);

  const deleteTeacher = useCallback(async (id) => {
    try {
      await sbFetch(`uct_teachers?id=eq.${id}`, { method:"DELETE" });
      setTeachers(t => t.filter(x => x.id!==id));
    } catch(e) { alert("Delete failed: " + e.message); }
  }, []);

  const toggleTeacherActive = useCallback(async (id) => {
    const teacher = teachers.find(t => t.id===id);
    if (teacher) await updateTeacher(id, { is_active: teacher.is_active==="YES" ? "NO" : "YES" });
  }, [teachers, updateTeacher]);

  // ── Church Attendance ─────────────────────────────────────
  const addChurchRec = useCallback(async (data) => {
    const mb = Number(data.male_beginning)||0, fb = Number(data.female_beginning)||0;
    const mc = Number(data.male_closing)||0,   fc = Number(data.female_closing)||0;
    const rec = { ...data, id:`C${Date.now()}`,
      total_beginning: String(mb+fb), total_closing: String(mc+fc),
      created_at: new Date().toISOString() };
    try {
      await sbFetch("uct_church", { method:"POST", body:JSON.stringify(rec) });
      setChurchRecs(c => [rec, ...c]);
    } catch(e) { alert("Save failed: " + e.message); }
    return rec;
  }, []);

  const updateChurchRec = useCallback(async (id, updates) => {
    const merged = { ...updates };
    merged.total_beginning = String((Number(merged.male_beginning)||0)+(Number(merged.female_beginning)||0));
    merged.total_closing   = String((Number(merged.male_closing)||0)+(Number(merged.female_closing)||0));
    try {
      await sbFetch(`uct_church?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(merged) });
      setChurchRecs(c => c.map(x => x.id===id ? {...x,...merged} : x));
    } catch(e) { alert("Update failed: " + e.message); }
  }, []);

  const deleteChurchRec = useCallback(async (id) => {
    try {
      await sbFetch(`uct_church?id=eq.${id}`, { method:"DELETE" });
      setChurchRecs(c => c.filter(x => x.id!==id));
    } catch(e) { alert("Delete failed: " + e.message); }
  }, []);

  // ── Programs ──────────────────────────────────────────────
  const addProgram = useCallback(async (name) => {
    const newP = { id:`P${Date.now()}`, name:name.trim(), is_active:"YES",
      sort_order: String(programs.length+1) };
    try {
      await sbFetch("uct_programs", { method:"POST", body:JSON.stringify(newP) });
      setPrograms(p => [...p, newP]);
    } catch(e) { alert("Save failed: " + e.message); }
  }, [programs]);

  const updateProgram = useCallback(async (id, updates) => {
    try {
      await sbFetch(`uct_programs?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(updates) });
      setPrograms(p => p.map(x => x.id===id ? {...x,...updates} : x));
    } catch(e) { alert("Update failed: " + e.message); }
  }, []);

  const deleteProgram = useCallback(async (id) => {
    try {
      await sbFetch(`uct_programs?id=eq.${id}`, { method:"DELETE" });
      setPrograms(p => p.filter(x => x.id!==id));
    } catch(e) { alert("Delete failed: " + e.message); }
  }, []);

  const toggleProgramActive = useCallback(async (id) => {
    const prog = programs.find(p => p.id===id);
    if (prog) await updateProgram(id, { is_active: prog.is_active==="YES" ? "NO" : "YES" });
  }, [programs, updateProgram]);

  // ── Export to Excel (download only, not storage) ──────────
  const downloadWorkbook = useCallback(() => {
    const wb = XLSX.utils.book_new();
    const toSheet = (rows) => XLSX.utils.json_to_sheet(rows.length ? rows : [{}]);
    XLSX.utils.book_append_sheet(wb, toSheet(records),    "SS Attendance");
    XLSX.utils.book_append_sheet(wb, toSheet(churchRecs), "Church Attendance");
    XLSX.utils.book_append_sheet(wb, toSheet(teachers),   "Teachers");
    XLSX.utils.book_append_sheet(wb, toSheet(classes),    "Classes");
    XLSX.utils.book_append_sheet(wb, toSheet(programs),   "Programs");
    XLSX.writeFile(wb, `UCT_Report_${new Date().toISOString().slice(0,10)}.xlsx`);
  }, [records, churchRecs, teachers, classes, programs]);

  return {
    records, teachers, classes, churchRecs, programs, loading, loadAll,
    addRecord, updateRecord, deleteRecord, approveRecord,
    addTeacher, updateTeacher, deleteTeacher, toggleTeacherActive,
    addChurchRec, updateChurchRec, deleteChurchRec,
    addProgram, updateProgram, deleteProgram, toggleProgramActive,
    downloadWorkbook,
  };
};




// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const CLASS_COLORS = ["#004b23","#1565C0","#2ECC71","#9B59B6","#E67E22","#E74C3C","#1ABC9C"];
const fmtDate = (d) => { try { return new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}); } catch { return d||"—"; } };
const statusInfo = (s) => ({ approved:{color:"#0E8A5F",label:"Approved"}, pending:{color:"#C87A0A",label:"Pending"}, rejected:{color:"#E74C3C",label:"Rejected"} }[s]||{color:"#7A8AAD",label:s||"—"});
const uid = () => Math.random().toString(36).slice(2,9);

// ─── ICONS ────────────────────────────────────────────────────────────────────
const PATHS = {
  dashboard:"M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h-3m-6 0a1 1 0 001-1v-4a1 1 0 011-1h2a1 1 0 011 1v4a1 1 0 001 1m-6 0h6",
  attendance:"M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4",
  analytics:"M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z",
  teachers:"M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z",
  classes:"M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4",
  export:"M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z",
  submit:"M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z",
  logout:"M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1",
  check:"M5 13l4 4L19 7", trash:"M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16",
  edit:"M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z",
  cross:"M12 2v20M2 12h20", plus:"M12 4v16m8-8H4",
  menu:"M4 6h16M4 12h16M4 18h16", close:"M6 18L18 6M6 6l12 12",
  eye:"M15 12a3 3 0 11-6 0 3 3 0 016 0z M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z",
  sun:"M12 3v1m0 16v1m9-9h-1M4 12H3m15.364-6.364l-.707.707M6.343 17.657l-.707.707m12.728 0l-.707-.707M6.343 6.343l-.707-.707M16 12a4 4 0 11-8 0 4 4 0 018 0z",
  moon:"M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z",
  bible:"M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253",
  upload:"M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12",
  pause:"M10 9v6m4-6v6m7-3a9 9 0 11-18 0 9 9 0 0118 0z",
  play:"M14.752 11.168l-3.197-2.132A1 1 0 0010 9.87v4.263a1 1 0 001.555.832l3.197-2.132a1 1 0 000-1.664z M21 12a9 9 0 11-18 0 9 9 0 0118 0z",
  info:"M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z",
};
const Icon = ({ name, size=18, color="currentColor" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
    {(PATHS[name]||"").split("M").filter(Boolean).map((d,i)=><path key={i} d={"M"+d}/>)}
  </svg>
);

// ─── SHARED UI ────────────────────────────────────────────────────────────────
const Badge = ({ label, color }) => (
  <span style={{ display:"inline-flex", alignItems:"center", padding:"3px 10px", borderRadius:20,
    fontSize:11, fontWeight:700, fontFamily:"'Trebuchet MS',sans-serif",
    background:color+"1A", color, border:`1px solid ${color}33` }}>{label}</span>
);

const Avatar = ({ name, size=40, color="#004b23" }) => {
  const initials = name?.split(" ").map(n=>n[0]).join("").slice(0,2).toUpperCase()||"??";
  return (
    <div style={{ width:size, height:size, borderRadius:"50%", flexShrink:0,
      background:`linear-gradient(135deg,${color}88,${color})`,
      display:"flex", alignItems:"center", justifyContent:"center",
      color:"#FFFFFF", fontWeight:700, fontFamily:"sans-serif", fontSize:size*0.36 }}>
      {initials}
    </div>
  );
};

const useThemeStyles = () => {
  const { dark } = useTheme();
  const t = T[dark?"dark":"light"];
  const btnGrad = `linear-gradient(135deg,${t.btnFrom},${t.btnTo})`;
  return {
    t,
    card:   { background:t.surface,    border:`1px solid ${t.border}`,       borderRadius:14, padding:20 },
    cardAlt:{ background:t.surfaceAlt, border:`1px solid ${t.border}`,       borderRadius:14, padding:20 },
    inp:    { background:dark?"rgba(255,255,255,0.06)":t.surfaceAlt, border:`1px solid ${t.border}`, borderRadius:9, padding:"10px 14px", color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, width:"100%", boxSizing:"border-box", outline:"none" },
    sel:    { background:dark?"#102654":t.surfaceAlt, border:`1px solid ${t.border}`, borderRadius:9, padding:"10px 14px", color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, outline:"none" },
    lbl:    { fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1 },
    th:     { padding:"10px 14px", textAlign:"left", borderBottom:`1px solid ${t.border}`, color:t.gold, fontSize:10, textTransform:"uppercase", letterSpacing:1.2, fontFamily:"'Trebuchet MS',sans-serif", whiteSpace:"nowrap" },
    td:     { padding:"11px 14px", borderBottom:`1px solid ${t.border}22`, color:t.text, verticalAlign:"middle", fontFamily:"'Trebuchet MS',sans-serif", fontSize:13 },
    btnGold:   { padding:"10px 20px", borderRadius:9, border:"none", cursor:"pointer", background:btnGrad, color:t.btnText, fontWeight:700, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13 },
    btnOutline:{ padding:"10px 18px", borderRadius:9, border:`1px solid ${t.gold}`, cursor:"pointer", background:"transparent", color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13 },
    btnGhost:  { padding:"8px 12px", borderRadius:8, border:`1px solid ${t.border}`, cursor:"pointer", background:"transparent" },
    tooltip:   { contentStyle:{ background:t.surface, border:`1px solid ${t.gold}33`, borderRadius:10, color:t.text, fontSize:12, fontFamily:"'Trebuchet MS',sans-serif" }, cursor:{ stroke:t.gold, strokeWidth:1 } },
  };
};

// ─── TOAST ────────────────────────────────────────────────────────────────────
const ToastContainer = ({ toasts, dismiss }) => (
  <div style={{ position:"fixed", top:20, right:20, zIndex:9999, display:"flex", flexDirection:"column", gap:10 }}>
    {toasts.map(({ id, msg, type }) => {
      const { dark } = useTheme();
      const t = T[dark?"dark":"light"];
      const colors = { success:t.success, danger:t.danger, info:t.gold, warn:t.warn };
      const c = colors[type]||t.gold;
      return (
        <div key={id} style={{ background:t.surface, border:`1px solid ${c}66`, borderLeft:`3px solid ${c}`,
          borderRadius:10, padding:"13px 18px", minWidth:290, boxShadow:`0 8px 32px rgba(0,0,0,0.25)`,
          display:"flex", justifyContent:"space-between", alignItems:"center", gap:12,
          fontFamily:"'Trebuchet MS',sans-serif" }}>
          <div>
            <div style={{ color:c, fontSize:12, fontWeight:700, marginBottom:2 }}>
              {type==="success"?"✓ Saved":type==="danger"?"⚠ Error":type==="warn"?"⚠ Warning":"ℹ Notice"}
            </div>
            <div style={{ color:t.text, fontSize:13 }}>{msg}</div>
          </div>
          <span onClick={()=>dismiss(id)} style={{ cursor:"pointer", color:t.textMuted, fontSize:18 }}>×</span>
        </div>
      );
    })}
  </div>
);

// ─── MODAL ────────────────────────────────────────────────────────────────────
const Modal = ({ title, onClose, children, width=560 }) => {
  const { dark } = useTheme();
  const t = T[dark?"dark":"light"];
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.7)", zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:t.surface, border:`1px solid ${t.border}`, borderRadius:16, width:"100%", maxWidth:width,
        maxHeight:"90vh", overflow:"auto", boxShadow:`0 24px 80px rgba(0,0,0,0.5)` }}>
        <div style={{ padding:"20px 24px", borderBottom:`1px solid ${t.border}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div style={{ fontSize:17, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif" }}>{title}</div>
          <button onClick={onClose} style={{ background:"transparent", border:"none", cursor:"pointer", color:t.textMuted, padding:4 }}>
            <Icon name="close" size={20} color={t.textMuted} />
          </button>
        </div>
        <div style={{ padding:24 }}>{children}</div>
      </div>
    </div>
  );
};

// ─── KPI CARD ─────────────────────────────────────────────────────────────────
const KpiCard = ({ label, value, sub, icon, color }) => {
  const { t } = useThemeStyles();
  const c = color||t.gold;
  return (
    <div style={{ background:t.surfaceAlt, border:`1px solid ${t.border}`, borderRadius:14,
      padding:"18px 20px", borderTop:`2px solid ${c}44` }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
        <div>
          <div style={{ fontSize:10, textTransform:"uppercase", letterSpacing:1.4, color:t.textMuted,
            fontFamily:"'Trebuchet MS',sans-serif", marginBottom:8 }}>{label}</div>
          <div style={{ fontSize:28, fontWeight:700, color:c, fontFamily:"'Georgia',serif", lineHeight:1 }}>{value}</div>
          {sub && <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:5 }}>{sub}</div>}
        </div>
        <div style={{ width:42, height:42, borderRadius:12, background:c+"18",
          display:"flex", alignItems:"center", justifyContent:"center" }}>
          <Icon name={icon} size={21} color={c} />
        </div>
      </div>
    </div>
  );
};

// ─── TEACHER FORM (Add / Edit) ────────────────────────────────────────────────
const TeacherForm = ({ initial, classes, onSave, onClose }) => {
  const { t, inp, sel, lbl, btnGold, btnOutline } = useThemeStyles();
  const blank = { name:"", phone:"", email:"", class_name:"", role:"Teacher", is_active:"YES", joined_date:new Date().toISOString().slice(0,10), notes:"" };
  const [form, setForm] = useState(initial ? { ...blank, ...initial } : blank);

  const handleChange = useCallback((e) => {
    const { name, value } = e.target;
    setForm(f => ({ ...f, [name]: value }));
  }, []);

  return (
    <div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
        <div style={{ display:"flex", flexDirection:"column", gap:5, gridColumn:"1 / -1" }}>
          <label style={lbl}>Full Name *</label>
          <input name="name" style={inp} value={form.name} onChange={handleChange} />
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={lbl}>Phone</label>
          <input name="phone" style={inp} value={form.phone} onChange={handleChange} />
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={lbl}>Email</label>
          <input name="email" style={inp} type="email" value={form.email} onChange={handleChange} />
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={lbl}>Assigned Class</label>
          <select name="class_name" style={{ ...sel, width:"100%" }} value={form.class_name} onChange={handleChange}>
            <option value="">Select…</option>
            {classes.map(c=><option key={c.name}>{c.name}</option>)}
          </select>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={lbl}>Role</label>
          <select name="role" style={{ ...sel, width:"100%" }} value={form.role} onChange={handleChange}>
            {["Teacher","Asst. Teacher","Superintendent","Observer"].map(o=><option key={o}>{o}</option>)}
          </select>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={lbl}>Date Joined</label>
          <input name="joined_date" style={inp} type="date" value={form.joined_date} onChange={handleChange} />
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5, gridColumn:"1 / -1" }}>
          <label style={lbl}>Notes</label>
          <textarea name="notes" style={{ ...inp, minHeight:60, resize:"vertical" }} value={form.notes} onChange={handleChange} />
        </div>
      </div>
      <div style={{ display:"flex", gap:10, marginTop:22, justifyContent:"flex-end" }}>
        <button style={btnOutline} onClick={onClose}>Cancel</button>
        <button style={btnGold} onClick={()=>{ if(!form.name.trim()){ alert("Name is required."); return; } onSave(form); }}>
          {initial ? "Save Changes" : "Add Teacher"}
        </button>
      </div>
    </div>
  );
};

// ─── TEACHERS PAGE ────────────────────────────────────────────────────────────
const TeachersPage = ({ db }) => {
  const { t, card, btnGold, btnGhost, th, td } = useThemeStyles();
  const { teachers, classes, addTeacher, updateTeacher, deleteTeacher, toggleTeacherActive } = db;
  const [modal, setModal]   = useState(null); // null | "add" | teacher-obj
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("all");

  const filtered = teachers.filter(t2 => {
    const matchSearch = t2.name?.toLowerCase().includes(search.toLowerCase()) ||
                        t2.class_name?.toLowerCase().includes(search.toLowerCase());
    const matchStatus = filterStatus === "all" || (filterStatus === "active" ? t2.is_active==="YES" : t2.is_active!=="YES");
    return matchSearch && matchStatus;
  });

  const active   = teachers.filter(t2=>t2.is_active==="YES").length;
  const inactive = teachers.length - active;

  const handleSave = (form) => {
    if (modal === "add") { addTeacher(form); }
    else { updateTeacher(modal.id, form); }
    setModal(null);
  };

  const inp = { background:"rgba(255,255,255,0.04)", border:`1px solid ${t.border}`, borderRadius:9,
    padding:"9px 14px", color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, outline:"none" };
  const sel = { ...inp, background:t.surfaceAlt };

  return (
    <div>
      {modal && (
        <Modal title={modal==="add"?"Add New Teacher":`Edit — ${modal.name}`} onClose={()=>setModal(null)}>
          <TeacherForm initial={modal==="add"?null:modal} classes={classes} onSave={handleSave} onClose={()=>setModal(null)} />
        </Modal>
      )}

      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:24 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Teacher Management</div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            {teachers.length} registered · {active} active · {inactive} inactive
          </div>
        </div>
        <button style={btnGold} onClick={()=>setModal("add")}>
          <span style={{ display:"flex", alignItems:"center", gap:7 }}><Icon name="plus" size={16} color="#0B1628" /> Add Teacher</span>
        </button>
      </div>

      {/* Summary cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:14, marginBottom:24 }}>
        <KpiCard label="Total Teachers" value={teachers.length} icon="teachers" color={t.gold} />
        <KpiCard label="Active"   value={active}   sub="Currently teaching" icon="play"  color={t.success} />
        <KpiCard label="Inactive" value={inactive} sub="Not currently active" icon="pause" color={t.textMuted} />
        <KpiCard label="Classes Covered" value={[...new Set(teachers.map(x=>x.class_name).filter(Boolean))].length} icon="classes" color={t.info} />
      </div>

      {/* Filters */}
      <div style={{ ...card, marginBottom:16, display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
        <input style={{ ...inp, maxWidth:240 }} placeholder="Search by name or class…" value={search} onChange={e=>setSearch(e.target.value)} />
        <select style={{ ...sel }} value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
          <option value="all">All Status</option>
          <option value="active">Active Only</option>
          <option value="inactive">Inactive Only</option>
        </select>
        <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginLeft:"auto" }}>{filtered.length} shown</span>
      </div>

      {/* Table */}
      <div style={{ ...card, padding:0, overflow:"hidden" }}>
        <table style={{ width:"100%", borderCollapse:"collapse" }}>
          <thead>
            <tr style={{ background:t.surfaceAlt }}>
              {["Teacher","Class","Role","Contact","Joined","Status","Actions"].map(h=>(
                <th key={h} style={th}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((teacher, i) => {
              const active = teacher.is_active === "YES";
              const clrIdx = classes.findIndex(c=>c.name===teacher.class_name);
              const color  = CLASS_COLORS[clrIdx >= 0 ? clrIdx % CLASS_COLORS.length : i % CLASS_COLORS.length];
              return (
                <tr key={teacher.id} style={{ transition:"background 0.15s" }}
                  onMouseEnter={e=>e.currentTarget.style.background=t.surfaceHover}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  {/* Teacher name + avatar */}
                  <td style={td}>
                    <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                      <Avatar name={teacher.name} size={36} color={active?color:t.textFaint} />
                      <div>
                        <div style={{ fontWeight:600, color:active?t.text:t.textMuted }}>{teacher.name}</div>
                        {teacher.notes && <div style={{ fontSize:11, color:t.textMuted, marginTop:1 }}>{teacher.notes.slice(0,40)}{teacher.notes.length>40?"…":""}</div>}
                      </div>
                    </div>
                  </td>
                  <td style={td}>
                    {teacher.class_name
                      ? <Badge label={teacher.class_name} color={color} />
                      : <span style={{ color:t.textMuted }}>—</span>}
                  </td>
                  <td style={{ ...td, color:t.textMuted }}>{teacher.role||"Teacher"}</td>
                  <td style={{ ...td, color:t.textMuted }}>
                    {teacher.phone && <div style={{ fontSize:12 }}>📞 {teacher.phone}</div>}
                    {teacher.email && <div style={{ fontSize:12 }}>✉ {teacher.email}</div>}
                    {!teacher.phone && !teacher.email && "—"}
                  </td>
                  <td style={{ ...td, color:t.textMuted, fontSize:12 }}>{teacher.joined_date ? fmtDate(teacher.joined_date) : "—"}</td>
                  <td style={td}>
                    <Badge label={active?"Active":"Inactive"} color={active?t.success:t.textMuted} />
                  </td>
                  <td style={td}>
                    <div style={{ display:"flex", gap:5 }}>
                      {/* Edit */}
                      <button style={btnGhost} title="Edit" onClick={()=>setModal(teacher)}>
                        <Icon name="edit" size={14} color={t.gold} />
                      </button>
                      {/* Toggle Active/Inactive */}
                      <button style={btnGhost} title={active?"Set Inactive":"Set Active"} onClick={()=>toggleTeacherActive(teacher.id)}>
                        <Icon name={active?"pause":"play"} size={14} color={active?t.warn:t.success} />
                      </button>
                      {/* Delete */}
                      <button style={btnGhost} title="Delete" onClick={()=>{
                        if(window.confirm(`Permanently delete ${teacher.name}?`)) deleteTeacher(teacher.id);
                      }}>
                        <Icon name="trash" size={14} color={t.danger} />
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {!filtered.length && (
              <tr><td colSpan={7} style={{ ...td, textAlign:"center", padding:48, color:t.textMuted }}>
                No teachers found. {search && "Try clearing the search."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Class distribution */}
      <div style={{ ...card, marginTop:20 }}>
        <div style={{ fontSize:14, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:16 }}>Teachers per Class</div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:10 }}>
          {classes.map((cls, i) => {
            const count = teachers.filter(x=>x.class_name===cls.name&&x.is_active==="YES").length;
            const color = CLASS_COLORS[i % CLASS_COLORS.length];
            return (
              <div key={cls.name} style={{ padding:"8px 16px", borderRadius:24, border:`1px solid ${color}44`,
                background:color+"11", display:"flex", gap:8, alignItems:"center" }}>
                <span style={{ fontSize:12, color, fontFamily:"'Trebuchet MS',sans-serif", fontWeight:600 }}>{cls.name}</span>
                <span style={{ fontSize:13, color, fontWeight:700, fontFamily:"'Georgia',serif" }}>{count}</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

// ─── SHARED ANALYTICS HELPERS ────────────────────────────────────────────────
const getWeekKey = (dateStr) => {
  const d = new Date(dateStr);
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const wn = Math.ceil(((d - jan1) / 86400000 + jan1.getDay() + 1) / 7);
  return `${d.getFullYear()}-W${String(wn).padStart(2,"0")}`;
};
const shortDate = (d) => { try { return new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short"}); } catch { return d; } };
const monthLabel = (ym) => { try { return new Date(ym+"-01").toLocaleDateString("en-GB",{month:"short",year:"2-digit"}); } catch { return ym; } };

// Build SS-vs-Church comparison data keyed by date
const buildComparison = (ssRecords, churchRecs, dateRange) => {
  // Collect all unique dates that appear in either dataset
  const dateSet = new Set();
  ssRecords.forEach(r => r.date && dateSet.add(r.date));
  churchRecs.forEach(r => r.date && dateSet.add(r.date));

  const rows = [...dateSet].sort().filter(d => {
    if (dateRange.from && d < dateRange.from) return false;
    if (dateRange.to   && d > dateRange.to)   return false;
    return true;
  }).map(date => {
    const ssDay = ssRecords.filter(r => r.date === date);
    const chDay = churchRecs.filter(r => r.date === date);
    const ssTotal = ssDay.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
    const chTotal = chDay.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
    const ssBegin = ssDay.reduce((s,r)=>s+(Number(r.total_beginning)||0),0);
    const chBegin = chDay.reduce((s,r)=>s+(Number(r.total_beginning)||0),0);
    return {
      date, label: shortDate(date),
      ssBegin, ssClose: ssTotal,
      chBegin, chClose: chTotal,
      retention: chTotal > 0 ? Math.round(ssTotal / chTotal * 100) : 0,
      diff: chTotal - ssTotal,
    };
  }).filter(r => r.ssClose > 0 || r.chClose > 0);
  return rows;
};

// ─── FILTER BAR ───────────────────────────────────────────────────────────────
const FilterBar = ({ filters, setFilters, months, programs, classes, showClass=true, showProgram=false, showMonth=true, showDateRange=true }) => {
  const { t, sel, inp } = useThemeStyles();
  const s = { ...sel, fontSize:12, padding:"7px 11px" };
  const i = { ...inp, fontSize:12, padding:"7px 11px" };
  return (
    <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
      {showDateRange && <>
        <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
          <span style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:0.8 }}>From</span>
          <input style={i} type="date" value={filters.from||""} onChange={e=>setFilters(f=>({...f,from:e.target.value}))} />
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
          <span style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:0.8 }}>To</span>
          <input style={i} type="date" value={filters.to||""} onChange={e=>setFilters(f=>({...f,to:e.target.value}))} />
        </div>
      </>}
      {showMonth && months?.length > 0 && (
        <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
          <span style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:0.8 }}>Month</span>
          <select style={s} value={filters.month||""} onChange={e=>setFilters(f=>({...f,month:e.target.value}))}>
            <option value="">All Months</option>
            {months.map(m=><option key={m} value={m}>{monthLabel(m)}</option>)}
          </select>
        </div>
      )}
      {showClass && classes?.length > 0 && (
        <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
          <span style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:0.8 }}>Class</span>
          <select style={s} value={filters.cls||""} onChange={e=>setFilters(f=>({...f,cls:e.target.value}))}>
            <option value="">All Classes</option>
            {classes.map(c=><option key={c.name} value={c.name}>{c.name}</option>)}
          </select>
        </div>
      )}
      {showProgram && programs?.length > 0 && (
        <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
          <span style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:0.8 }}>Program</span>
          <select style={s} value={filters.program||""} onChange={e=>setFilters(f=>({...f,program:e.target.value}))}>
            <option value="">All Programs</option>
            {programs.map(p=><option key={p.id} value={p.name}>{p.name}</option>)}
          </select>
        </div>
      )}
      {(filters.from||filters.to||filters.month||filters.cls||filters.program) && (
        <button onClick={()=>setFilters({})} style={{ padding:"7px 12px", borderRadius:7, border:`1px solid ${t.danger}44`, background:t.danger+"11", color:t.danger, fontFamily:"'Trebuchet MS',sans-serif", fontSize:11, cursor:"pointer", marginTop:14 }}>
          ✕ Clear
        </button>
      )}
    </div>
  );
};

// ─── INSIGHT PILL ─────────────────────────────────────────────────────────────
const Insight = ({ text, color, icon="info" }) => {
  const { t } = useThemeStyles();
  return (
    <div style={{ display:"flex", alignItems:"flex-start", gap:10, padding:"11px 14px", borderRadius:10, background:color+"0D", border:`1px solid ${color}33` }}>
      <Icon name={icon} size={15} color={color} />
      <span style={{ fontSize:12, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", lineHeight:1.6 }}>{text}</span>
    </div>
  );
};

// ─── DASHBOARD PAGE ───────────────────────────────────────────────────────────
const DashboardPage = ({ db }) => {
  const { t, card, tooltip, sel } = useThemeStyles();
  const { records, teachers, classes, churchRecs, programs } = db;
  const active = teachers.filter(x=>x.is_active==="YES").length;

  const [filters, setFilters] = useState({});

  // All unique months across both datasets
  const allMonths = [...new Set([
    ...records.map(r=>(r.date||"").slice(0,7)),
    ...churchRecs.map(r=>(r.date||"").slice(0,7))
  ].filter(Boolean))].sort();

  // Apply date/month filters
  const applyRange = (arr) => arr.filter(r => {
    if (!r.date) return false;
    if (filters.from && r.date < filters.from) return false;
    if (filters.to   && r.date > filters.to)   return false;
    if (filters.month && !r.date.startsWith(filters.month)) return false;
    return true;
  });

  const fRec  = applyRange(records);
  const fChurch = applyRange(churchRecs);

  // KPIs
  const ssTotal   = fRec.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const chTotal   = fChurch.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const bibles    = fRec.reduce((s,r)=>s+(Number(r.bibles_closing)||0),0);
  const bibleRate = ssTotal > 0 ? Math.round(bibles/ssTotal*100) : 0;
  const totalFT   = fChurch.reduce((s,r)=>s+(Number(r.first_timers)||0),0) + fRec.reduce((s,r)=>s+(Number(r.first_timers)||0),0);
  const retentionRate = chTotal > 0 ? Math.round(ssTotal/chTotal*100) : 0;
  const pending = records.filter(r=>r.status==="pending").length;

  // SS vs Church comparison by date
  const compData = buildComparison(fRec, fChurch, {});

  // Weekly trend (SS + Church combined)
  const weekMap = {};
  fRec.forEach(r => {
    if (!r.date) return;
    const wk = getWeekKey(r.date);
    if (!weekMap[wk]) weekMap[wk] = { wk, label: r.date.slice(5), ss:0, church:0, bibles:0 };
    weekMap[wk].ss     += Number(r.total_closing)||0;
    weekMap[wk].bibles += Number(r.bibles_closing)||0;
  });
  fChurch.forEach(r => {
    if (!r.date) return;
    const wk = getWeekKey(r.date);
    if (!weekMap[wk]) weekMap[wk] = { wk, label: r.date.slice(5), ss:0, church:0, bibles:0 };
    weekMap[wk].church += Number(r.total_closing)||0;
  });
  const weeklyTrend = Object.values(weekMap).sort((a,b)=>a.wk>b.wk?1:-1).slice(-10);

  // Class pie
  const classTotals = classes.map((cls,i)=>({
    name: cls.name.split(" ")[0],
    value: fRec.filter(r=>r.class_name===cls.name).reduce((s,r)=>s+(Number(r.total_closing)||0),0),
    color: CLASS_COLORS[i%CLASS_COLORS.length],
  })).filter(x=>x.value>0);

  // Generate storytelling insight
  const topClass = [...classes].sort((a,b)=>
    fRec.filter(r=>r.class_name===b.name).reduce((s,r)=>s+(Number(r.total_closing)||0),0) -
    fRec.filter(r=>r.class_name===a.name).reduce((s,r)=>s+(Number(r.total_closing)||0),0)
  )[0];
  const retentionStory = retentionRate > 0
    ? retentionRate >= 80 ? `✝ Strong conversion: ${retentionRate}% of Sunday service attendees are enrolled in Sunday School.`
    : retentionRate >= 50 ? `📈 Moderate engagement: ${retentionRate}% of Sunday service attendees attend Sunday School. There is room to grow.`
    : `⚠ Low SS engagement: Only ${retentionRate}% of Sunday service attendees are in Sunday School. Consider outreach.`
    : null;

  const ChartCard = ({ title, sub, children, span=1 }) => (
    <div style={{ ...card, gridColumn:span===2?"1 / -1":"auto" }}>
      <div style={{ marginBottom:14 }}>
        <div style={{ fontSize:14, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif" }}>{title}</div>
        {sub && <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:2 }}>{sub}</div>}
      </div>
      {children}
    </div>
  );

  return (
    <div>
      {/* Header + Filters */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:16, flexWrap:"wrap", gap:12 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Overview Dashboard</div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            {records.length} SS records · {churchRecs.length} church services
            {(filters.from||filters.to||filters.month) ? " · Filtered" : " · All time"}
          </div>
        </div>
      </div>

      {/* Filter Bar */}
      <div style={{ ...card, marginBottom:20, padding:"14px 18px" }}>
        <div style={{ fontSize:11, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:10 }}>
          📅 Dashboard Filters
        </div>
        <FilterBar filters={filters} setFilters={setFilters} months={allMonths}
          programs={programs} classes={classes} showClass={false} showProgram={false} />
      </div>

      {/* Storytelling insight */}
      {retentionStory && (
        <div style={{ marginBottom:20 }}>
          <Insight text={retentionStory} color={retentionRate >= 80 ? t.success : retentionRate >= 50 ? t.warn : t.danger} icon="info" />
        </div>
      )}

      {/* KPIs */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))", gap:14, marginBottom:24 }}>
        <KpiCard label="SS Total (Closing)"     value={ssTotal}    sub="Sunday School" icon="attendance" color={t.gold} />
        <KpiCard label="Church Total (Closing)" value={chTotal}    sub="Main service"  icon="cross"      color={t.info} />
        <KpiCard label="SS Retention Rate"      value={`${retentionRate}%`} sub="SS ÷ Church closing" icon="analytics" color={retentionRate>=80?t.success:retentionRate>=50?t.warn:t.danger} />
        <KpiCard label="Bibles Brought"         value={bibles}     sub={`${bibleRate}% of SS`} icon="bible"      color="#9B59B6" />
        <KpiCard label="First Timers"           value={totalFT}    sub="SS + Church"   icon="plus"       color="#E67E22" />
        <KpiCard label="Active Teachers"        value={active}     sub={`of ${teachers.length}`} icon="teachers" color={t.success} />
        <KpiCard label="Pending Reviews"        value={pending}    sub="Needs action"  icon="info"       color={pending>0?t.warn:t.textMuted} />
      </div>

      {/* Row 1: SS vs Church side-by-side comparison by date */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20, marginBottom:20 }}>
        <ChartCard title="SS vs Church Closing — by Date"
          sub="Compares Sunday School closing count to same-day church service closing count">
          {compData.length > 0 ? (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={compData} barCategoryGap="25%">
                <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                <YAxis stroke={t.textMuted} fontSize={11} />
                <Tooltip {...tooltip} />
                <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                <Bar dataKey="ssClose"  name="SS Closing"     fill={t.gold}    radius={[3,3,0,0]} />
                <Bar dataKey="chClose"  name="Church Closing" fill={t.info}    radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          ) : <div style={{ height:220, display:"flex", alignItems:"center", justifyContent:"center", color:t.textMuted, fontSize:13, fontFamily:"'Trebuchet MS',sans-serif" }}>Record both SS and Church attendance on the same date to see comparison</div>}
        </ChartCard>

        <ChartCard title="SS vs Church Beginning — by Date"
          sub="Compares how many people arrived at Sunday School vs the church service">
          {compData.length > 0 ? (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={compData} barCategoryGap="25%">
                <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                <YAxis stroke={t.textMuted} fontSize={11} />
                <Tooltip {...tooltip} />
                <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                <Bar dataKey="ssBegin"  name="SS Beginning"     fill={t.goldDark||"#9A7A2C"} radius={[3,3,0,0]} />
                <Bar dataKey="chBegin"  name="Church Beginning" fill="#4A9EDB88"   radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          ) : <div style={{ height:220, display:"flex", alignItems:"center", justifyContent:"center", color:t.textMuted, fontSize:13, fontFamily:"'Trebuchet MS',sans-serif" }}>No overlapping dates found yet</div>}
        </ChartCard>
      </div>

      {/* Row 2: Weekly combined trend + Class pie */}
      <div style={{ display:"grid", gridTemplateColumns:"1.5fr 1fr", gap:20, marginBottom:20 }}>
        <ChartCard title="Weekly Trend — SS & Church Attendance"
          sub="How Sunday School and main service attendance move together week by week">
          <ResponsiveContainer width="100%" height={210}>
            <AreaChart data={weeklyTrend.length ? weeklyTrend : [{label:"No data",ss:0,church:0}]}>
              <defs>
                <linearGradient id="gSS"  x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%"  stopColor={t.gold} stopOpacity={0.3}/><stop offset="95%" stopColor={t.gold} stopOpacity={0}/>
                </linearGradient>
                <linearGradient id="gCh"  x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%"  stopColor={t.info} stopOpacity={0.2}/><stop offset="95%" stopColor={t.info} stopOpacity={0}/>
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
              <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
              <YAxis stroke={t.textMuted} fontSize={11} />
              <Tooltip {...tooltip} />
              <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
              <Area type="monotone" dataKey="ss"     stroke={t.gold} fill="url(#gSS)" strokeWidth={2.5} name="Sunday School" />
              <Area type="monotone" dataKey="church" stroke={t.info} fill="url(#gCh)" strokeWidth={2.5} name="Church Service" />
            </AreaChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Sunday School by Class" sub="Closing attendance share">
          {classTotals.length > 0 ? (
            <>
              <ResponsiveContainer width="100%" height={165}>
                <PieChart>
                  <Pie data={classTotals} cx="50%" cy="50%" innerRadius={44} outerRadius={76} dataKey="value" paddingAngle={3}>
                    {classTotals.map((e,i)=><Cell key={i} fill={e.color}/>)}
                  </Pie>
                  <Tooltip {...tooltip} />
                </PieChart>
              </ResponsiveContainer>
              <div style={{ display:"flex", flexWrap:"wrap", gap:7, marginTop:8 }}>
                {classTotals.map(d=>(
                  <div key={d.name} style={{ display:"flex", alignItems:"center", gap:4 }}>
                    <div style={{ width:7, height:7, borderRadius:2, background:d.color }} />
                    <span style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>{d.name} ({d.value})</span>
                  </div>
                ))}
              </div>
            </>
          ) : <div style={{ height:165, display:"flex", alignItems:"center", justifyContent:"center", color:t.textMuted, fontSize:13, fontFamily:"'Trebuchet MS',sans-serif" }}>No SS records yet</div>}
        </ChartCard>
      </div>

      {/* Row 3: Retention % over time */}
      {compData.length > 1 && (
        <ChartCard title="SS Retention Rate Over Time" sub="What % of church service attendees were in Sunday School on the same day" span={2}>
          <ResponsiveContainer width="100%" height={180}>
            <LineChart data={compData}>
              <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
              <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
              <YAxis stroke={t.textMuted} fontSize={11} unit="%" domain={[0,110]} />
              <Tooltip {...tooltip} formatter={(v)=>`${v}%`} />
              <Line type="monotone" dataKey="retention" stroke={t.success} strokeWidth={2.5} dot={{ fill:t.success, r:4 }} name="Retention %" />
            </LineChart>
          </ResponsiveContainer>
        </ChartCard>
      )}

      {/* Recent Activity */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20, marginTop:20 }}>
        <div style={{ ...card, padding:0, overflow:"hidden" }}>
          <div style={{ padding:"13px 18px", borderBottom:`1px solid ${t.border}`, display:"flex", justifyContent:"space-between" }}>
            <span style={{ fontSize:13, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif" }}>Recent SS Reports</span>
            <Badge label={`${fRec.length} total`} color={t.gold} />
          </div>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr style={{ background:t.surfaceAlt }}>
              {["Date","Class","Present","Status"].map(h=>(
                <th key={h} style={{ padding:"8px 12px", textAlign:"left", borderBottom:`1px solid ${t.border}`, color:t.gold, fontSize:10, textTransform:"uppercase", letterSpacing:1, fontFamily:"'Trebuchet MS',sans-serif" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {[...fRec].reverse().slice(0,6).map(r=>{
                const si=statusInfo(r.status);
                return <tr key={r.id}>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14`, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>{fmtDate(r.date)}</td>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14`, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>{r.class_name}</td>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14`, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12, fontWeight:700 }}>{r.total_closing||0}</td>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14` }}><Badge label={si.label} color={si.color}/></td>
                </tr>;
              })}
              {!fRec.length && <tr><td colSpan={4} style={{ padding:"28px 12px", textAlign:"center", color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>No SS records</td></tr>}
            </tbody>
          </table>
        </div>
        <div style={{ ...card, padding:0, overflow:"hidden" }}>
          <div style={{ padding:"13px 18px", borderBottom:`1px solid ${t.border}`, display:"flex", justifyContent:"space-between" }}>
            <span style={{ fontSize:13, fontWeight:700, color:t.info, fontFamily:"'Georgia',serif" }}>Recent Church Services</span>
            <Badge label={`${fChurch.length} total`} color={t.info} />
          </div>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr style={{ background:t.surfaceAlt }}>
              {["Date","Program","Closing","1st Timers"].map(h=>(
                <th key={h} style={{ padding:"8px 12px", textAlign:"left", borderBottom:`1px solid ${t.border}`, color:t.info, fontSize:10, textTransform:"uppercase", letterSpacing:1, fontFamily:"'Trebuchet MS',sans-serif" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {[...fChurch].reverse().slice(0,6).map(r=>(
                <tr key={r.id}>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14`, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>{fmtDate(r.date)}</td>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14`, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>{r.program}</td>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14`, color:t.info, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12, fontWeight:700 }}>{r.total_closing||0}</td>
                  <td style={{ padding:"9px 12px", borderBottom:`1px solid ${t.border}14`, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>{r.first_timers||0}</td>
                </tr>
              ))}
              {!fChurch.length && <tr><td colSpan={4} style={{ padding:"28px 12px", textAlign:"center", color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>No church records</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─── SUBMIT PAGE ──────────────────────────────────────────────────────────────
const SubmitPage = ({ db, user, onSuccess }) => {
  const { t, inp, sel, lbl, btnGold, btnOutline, card } = useThemeStyles();
  const { teachers, classes, addRecord } = db;
  const activeTeachers = teachers.filter(x => x.is_active === "YES");

  const makeBlank = () => ({
    date: new Date().toISOString().slice(0,10), day_of_week:"Sunday",
    service_type:"Regular Sunday School", class_name:"", teacher_name: user?.name||"",
    assistant_teacher:"", submitted_by: user?.name||"", time_started:"09:00", time_ended:"10:30",
    total_beginning:"", total_closing:"", male_present:"", female_present:"",
    first_timers:"", visitors:"", absent_members:"",
    bibles_beginning:"", bibles_closing:"", members_without_bibles:"",
    topic:"", bible_references:"", memory_verse:"", key_notes:"", assignment:"",
    teachers_present:"", teacher_names:"", challenges:"", prayer_requests:"", announcements:""
  });

  const [form, setForm]     = useState(makeBlank);
  const [loading, setLoading] = useState(false);
  const [done, setDone]     = useState(false);

  // KEY FIX: single stable handler using input's `name` attribute
  const handleChange = useCallback((e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  }, []);

  const handleSubmit = () => {
    if (!form.class_name)   { alert("Please select a class."); return; }
    if (!form.topic.trim()) { alert("Please enter the lesson topic."); return; }
    setLoading(true);
    setTimeout(() => {
      addRecord(form);
      setLoading(false);
      setDone(true);
      setTimeout(() => { setDone(false); setForm(makeBlank()); onSuccess && onSuccess(); }, 2500);
    }, 400);
  };

  const g2  = { display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 };
  const fw  = (span=1) => ({ display:"flex", flexDirection:"column", gap:5, gridColumn: span===2?"1 / -1":"auto" });
  const sec = { fontSize:11, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif",
    textTransform:"uppercase", letterSpacing:1.8, padding:"18px 0 8px",
    borderTop:`1px solid ${t.border}`, marginTop:16 };

  if (done) return (
    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:400, gap:16 }}>
      <div style={{ width:80, height:80, borderRadius:"50%", background:t.success+"18", display:"flex", alignItems:"center", justifyContent:"center" }}>
        <Icon name="check" size={40} color={t.success} />
      </div>
      <div style={{ fontSize:24, color:t.gold, fontFamily:"'Georgia',serif" }}>Report Submitted!</div>
      <div style={{ color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:14 }}>Saved to your Excel workbook.</div>
    </div>
  );

  return (
    <div>
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Submit Attendance Report</div>
      <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:24 }}>Saved directly to your local Excel workbook — no internet required.</div>
      <div style={card}>

        <div style={sec}>Basic Information</div>
        <div style={g2}>
          <div style={fw()}><label style={lbl}>Date</label><input name="date" style={inp} type="date" value={form.date} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Day</label><select name="day_of_week" style={{...sel,width:"100%"}} value={form.day_of_week} onChange={handleChange}>{["Sunday","Saturday","Wednesday"].map(o=><option key={o}>{o}</option>)}</select></div>
          <div style={fw()}><label style={lbl}>Service Type</label><select name="service_type" style={{...sel,width:"100%"}} value={form.service_type} onChange={handleChange}><option>Regular Sunday School</option><option>Joint Sunday School</option><option>Special Service</option></select></div>
          <div style={fw()}><label style={lbl}>Class Name *</label><select name="class_name" style={{...sel,width:"100%"}} value={form.class_name} onChange={handleChange}><option value="">Select…</option>{classes.map(c=><option key={c.name} value={c.name}>{c.name}</option>)}</select></div>
          <div style={fw()}><label style={lbl}>Teacher Name</label><select name="teacher_name" style={{...sel,width:"100%"}} value={form.teacher_name} onChange={handleChange}><option value="">Select…</option>{activeTeachers.map(x=><option key={x.name} value={x.name}>{x.name}</option>)}</select></div>
          <div style={fw()}><label style={lbl}>Assistant Teacher</label><select name="assistant_teacher" style={{...sel,width:"100%"}} value={form.assistant_teacher} onChange={handleChange}><option value="">None</option>{activeTeachers.map(x=><option key={x.name} value={x.name}>{x.name}</option>)}</select></div>
          <div style={fw()}><label style={lbl}>Submitted By</label><input name="submitted_by" style={inp} value={form.submitted_by} onChange={handleChange} /></div>
          <div />
          <div style={fw()}><label style={lbl}>Time Started</label><input name="time_started" style={inp} type="time" value={form.time_started} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Time Ended</label><input name="time_ended" style={inp} type="time" value={form.time_ended} onChange={handleChange} /></div>
        </div>

        <div style={sec}>Attendance Information</div>
        <div style={g2}>
          <div style={fw()}><label style={lbl}>Total Present (Beginning)</label><input name="total_beginning" style={inp} type="number" value={form.total_beginning} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Total Present (Closing)</label><input name="total_closing" style={inp} type="number" value={form.total_closing} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Male Present</label><input name="male_present" style={inp} type="number" value={form.male_present} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Female Present</label><input name="female_present" style={inp} type="number" value={form.female_present} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>First Timers</label><input name="first_timers" style={inp} type="number" value={form.first_timers} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Visitors</label><input name="visitors" style={inp} type="number" value={form.visitors} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Absent Members</label><input name="absent_members" style={inp} type="number" value={form.absent_members} onChange={handleChange} /></div>
        </div>

        <div style={sec}>Bible Records</div>
        <div style={g2}>
          <div style={fw()}><label style={lbl}>Bibles at Beginning</label><input name="bibles_beginning" style={inp} type="number" value={form.bibles_beginning} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Bibles at Closing</label><input name="bibles_closing" style={inp} type="number" value={form.bibles_closing} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Members Without Bibles</label><input name="members_without_bibles" style={inp} type="number" value={form.members_without_bibles} onChange={handleChange} /></div>
        </div>

        <div style={sec}>Lesson Information</div>
        <div style={g2}>
          <div style={fw(2)}><label style={lbl}>Topic Discussed *</label><input name="topic" style={inp} value={form.topic} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Bible References</label><input name="bible_references" style={inp} value={form.bible_references} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Memory Verse</label><input name="memory_verse" style={inp} value={form.memory_verse} onChange={handleChange} /></div>
          <div style={fw(2)}><label style={lbl}>Key Notes / Summary</label><textarea name="key_notes" style={{...inp,minHeight:65,resize:"vertical"}} value={form.key_notes} onChange={handleChange} /></div>
          <div style={fw(2)}><label style={lbl}>Assignment / Homework</label><textarea name="assignment" style={{...inp,minHeight:65,resize:"vertical"}} value={form.assignment} onChange={handleChange} /></div>
        </div>

        <div style={sec}>Teacher Attendance</div>
        <div style={g2}>
          <div style={fw()}><label style={lbl}>Number of Teachers Present</label><input name="teachers_present" style={inp} type="number" value={form.teachers_present} onChange={handleChange} /></div>
          <div style={fw()}><label style={lbl}>Names of Teachers Present</label><input name="teacher_names" style={inp} value={form.teacher_names} onChange={handleChange} /></div>
        </div>

        <div style={sec}>Additional Notes</div>
        <div style={g2}>
          <div style={fw(2)}><label style={lbl}>Challenges</label><textarea name="challenges" style={{...inp,minHeight:65,resize:"vertical"}} value={form.challenges} onChange={handleChange} /></div>
          <div style={fw(2)}><label style={lbl}>Prayer Requests</label><textarea name="prayer_requests" style={{...inp,minHeight:65,resize:"vertical"}} value={form.prayer_requests} onChange={handleChange} /></div>
          <div style={fw(2)}><label style={lbl}>Announcements</label><textarea name="announcements" style={{...inp,minHeight:65,resize:"vertical"}} value={form.announcements} onChange={handleChange} /></div>
        </div>

        <div style={{ marginTop:24, display:"flex", gap:12 }}>
          <button onClick={handleSubmit} disabled={loading} style={{ ...btnGold, padding:"13px 36px", fontSize:14 }}>
            {loading ? "Saving…" : "Submit Report"}
          </button>
          <button onClick={()=>setForm(makeBlank())} style={{ ...btnOutline, padding:"13px 24px" }}>Clear Form</button>
        </div>
      </div>
    </div>
  );
};


// ─── ATTENDANCE PAGE ──────────────────────────────────────────────────────────
const AttendancePage = ({ db, user }) => {
  const { t, card, btnGhost, th, td } = useThemeStyles();
  const { records, classes, approveRecord, deleteRecord } = db;
  const [detail, setDetail] = useState(null);
  const [filter, setFilter] = useState({ cls:"", status:"", search:"" });

  const filtered = records.filter(r =>
    (!filter.cls    || r.class_name===filter.cls) &&
    (!filter.status || r.status===filter.status) &&
    (!filter.search || r.topic?.toLowerCase().includes(filter.search.toLowerCase()) ||
                       r.teacher_name?.toLowerCase().includes(filter.search.toLowerCase()))
  );

  const sel = { background:t.surfaceAlt, border:`1px solid ${t.border}`, borderRadius:9, padding:"8px 12px", color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, outline:"none" };
  const inp = { ...sel, background:"rgba(255,255,255,0.04)" };

  if (detail) {
    const r = detail;
    const Row = ({l,v}) => (
      <div style={{ display:"flex", justifyContent:"space-between", padding:"8px 0", borderBottom:`1px solid ${t.border}18` }}>
        <span style={{ color:t.textMuted, fontSize:13, fontFamily:"'Trebuchet MS',sans-serif" }}>{l}</span>
        <span style={{ color:t.text, fontSize:13, fontFamily:"'Trebuchet MS',sans-serif", textAlign:"right", maxWidth:"55%" }}>{v||"—"}</span>
      </div>
    );
    return (
      <div>
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:24 }}>
          <button onClick={()=>setDetail(null)} style={{ padding:"9px 18px", borderRadius:9, border:`1px solid ${t.border}`, background:"transparent", color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", cursor:"pointer" }}>← Back</button>
          <div style={{ fontSize:20, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif" }}>Record Detail</div>
          <Badge label={statusInfo(r.status).label} color={statusInfo(r.status).color} />
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:18 }}>
          {[
            ["Basic Info", [["Date",fmtDate(r.date)],["Day",r.day_of_week],["Service",r.service_type],["Class",r.class_name],["Teacher",r.teacher_name],["Assistant",r.assistant_teacher],["Time",`${r.time_started||""} – ${r.time_ended||""}`]]],
            ["Attendance", [["At Beginning",r.total_beginning],["At Closing",r.total_closing],["Male",r.male_present],["Female",r.female_present],["First Timers",r.first_timers],["Visitors",r.visitors],["Absent",r.absent_members]]],
            ["Bible Records", [["Bibles (Begin)",r.bibles_beginning],["Bibles (Closing)",r.bibles_closing],["Without Bibles",r.members_without_bibles]]],
            ["Lesson", [["Topic",r.topic],["Bible Refs",r.bible_references],["Memory Verse",r.memory_verse],["Key Notes",r.key_notes],["Assignment",r.assignment]]],
          ].map(([sec,rows])=>(
            <div key={sec} style={card}>
              <div style={{ fontSize:11, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.5, marginBottom:10 }}>{sec}</div>
              {rows.map(([l,v])=><Row key={l} l={l} v={v}/>)}
            </div>
          ))}
          <div style={{ ...card, gridColumn:"1 / -1" }}>
            <div style={{ fontSize:11, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.5, marginBottom:10 }}>Additional Notes</div>
            <Row l="Challenges"       v={r.challenges} />
            <Row l="Prayer Requests"  v={r.prayer_requests} />
            <Row l="Announcements"    v={r.announcements} />
            <Row l="Teachers Present" v={`${r.teachers_present||""} — ${r.teacher_names||""}`} />
          </div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Attendance Records</div>
      <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:20 }}>{records.length} records in workbook</div>
      <div style={{ ...card, marginBottom:16, display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
        <input style={{ ...inp, maxWidth:220 }} placeholder="Search topic or teacher…" value={filter.search} onChange={e=>setFilter(f=>({...f,search:e.target.value}))} />
        <select style={sel} value={filter.cls} onChange={e=>setFilter(f=>({...f,cls:e.target.value}))}>
          <option value="">All Classes</option>
          {db.classes.map(c=><option key={c.name}>{c.name}</option>)}
        </select>
        <select style={sel} value={filter.status} onChange={e=>setFilter(f=>({...f,status:e.target.value}))}>
          <option value="">All Status</option>
          <option value="approved">Approved</option>
          <option value="pending">Pending</option>
        </select>
        <button onClick={()=>setFilter({cls:"",status:"",search:""})} style={{ padding:"8px 14px", borderRadius:9, border:`1px solid ${t.border}`, background:"transparent", color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", cursor:"pointer", fontSize:13 }}>Clear</button>
        <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginLeft:"auto" }}>{filtered.length} shown</span>
      </div>
      <div style={{ ...card, padding:0, overflow:"hidden" }}>
        <table style={{ width:"100%", borderCollapse:"collapse" }}>
          <thead><tr style={{ background:t.surfaceAlt }}>
            {["Date","Class","Teacher","Present","Bibles","First T.","Topic","Status","Actions"].map(h=><th key={h} style={th}>{h}</th>)}
          </tr></thead>
          <tbody>
            {[...filtered].reverse().map(r=>{
              const si = statusInfo(r.status);
              return (
                <tr key={r.id} onMouseEnter={e=>e.currentTarget.style.background=t.surfaceHover} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={td}>{fmtDate(r.date)}</td>
                  <td style={td}>{r.class_name}</td>
                  <td style={{ ...td, color:t.textMuted }}>{r.teacher_name}</td>
                  <td style={td}>{r.total_closing}</td>
                  <td style={td}>{r.bibles_closing}</td>
                  <td style={td}>{r.first_timers}</td>
                  <td style={{ ...td, maxWidth:140, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", color:t.textMuted }}>{r.topic}</td>
                  <td style={td}><Badge label={si.label} color={si.color}/></td>
                  <td style={td}>
                    <div style={{ display:"flex", gap:5 }}>
                      <button style={btnGhost} title="View" onClick={()=>setDetail(r)}><Icon name="eye" size={13} color={t.textMuted}/></button>
                      {user?.role==="admin" && r.status!=="approved" && <button style={btnGhost} title="Approve" onClick={()=>approveRecord(r.id)}><Icon name="check" size={13} color={t.success}/></button>}
                      {user?.role==="admin" && <button style={btnGhost} title="Delete" onClick={()=>{ if(window.confirm("Delete record?")) deleteRecord(r.id); }}><Icon name="trash" size={13} color={t.danger}/></button>}
                    </div>
                  </td>
                </tr>
              );
            })}
            {!filtered.length && <tr><td colSpan={9} style={{ ...td, textAlign:"center", padding:48, color:t.textMuted }}>No records found</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
};

// ─── ANALYTICS PAGE ───────────────────────────────────────────────────────────
const AnalyticsPage = ({ db }) => {
  const { t, card, tooltip } = useThemeStyles();
  const { records, classes, churchRecs, programs } = db;

  const [filters, setFilters] = useState({});
  const [tab, setTab] = useState("comparison"); // comparison | ss | church | gender

  const allMonths = [...new Set([
    ...records.map(r=>(r.date||"").slice(0,7)),
    ...churchRecs.map(r=>(r.date||"").slice(0,7))
  ].filter(Boolean))].sort();

  const applyRange = (arr) => arr.filter(r => {
    if (!r.date) return false;
    if (filters.from  && r.date < filters.from)  return false;
    if (filters.to    && r.date > filters.to)     return false;
    if (filters.month && !r.date.startsWith(filters.month)) return false;
    if (filters.cls   && r.class_name !== filters.cls) return false;
    if (filters.program && r.program  !== filters.program) return false;
    return true;
  });

  const fRec    = applyRange(records);
  const fChurch = applyRange(churchRecs);

  // ── Comparison data ──────────────────────────────────────
  const compData = buildComparison(fRec, fChurch, {});

  // ── Monthly grouped: SS + Church ─────────────────────────
  const monthMap = {};
  fRec.forEach(r => {
    const m = (r.date||"").slice(0,7); if(!m) return;
    if (!monthMap[m]) monthMap[m] = { m, label:monthLabel(m), ssClose:0, ssBegin:0, chClose:0, chBegin:0, bibles:0, ft:0 };
    monthMap[m].ssClose += Number(r.total_closing)||0;
    monthMap[m].ssBegin += Number(r.total_beginning)||0;
    monthMap[m].bibles  += Number(r.bibles_closing)||0;
    monthMap[m].ft      += Number(r.first_timers)||0;
  });
  fChurch.forEach(r => {
    const m = (r.date||"").slice(0,7); if(!m) return;
    if (!monthMap[m]) monthMap[m] = { m, label:monthLabel(m), ssClose:0, ssBegin:0, chClose:0, chBegin:0, bibles:0, ft:0 };
    monthMap[m].chClose += Number(r.total_closing)||0;
    monthMap[m].chBegin += Number(r.total_beginning)||0;
    monthMap[m].ft      += Number(r.first_timers)||0;
  });
  const monthlyData = Object.values(monthMap).sort((a,b)=>a.m>b.m?1:-1);

  // ── SS per-class monthly ──────────────────────────────────
  const classMonthMap = {};
  fRec.forEach(r => {
    const m = (r.date||"").slice(0,7); if(!m||!r.class_name) return;
    if (!classMonthMap[m]) classMonthMap[m] = { label: monthLabel(m) };
    classMonthMap[m][r.class_name] = (classMonthMap[m][r.class_name]||0) + (Number(r.total_closing)||0);
  });
  const classMonthlyData = Object.values(classMonthMap).sort((a,b)=>a.label>b.label?1:-1);

  // ── Bible participation ───────────────────────────────────
  const bibleClass = classes.map((cls,i) => {
    const recs = fRec.filter(r=>r.class_name===cls.name && Number(r.total_closing)>0);
    const rate = recs.length ? Math.round(recs.reduce((s,r)=>s+((Number(r.bibles_closing)||0)/Number(r.total_closing)*100),0)/recs.length) : 0;
    return { name:cls.name.split(" ")[0], rate, fill:CLASS_COLORS[i%CLASS_COLORS.length] };
  }).filter(x=>x.rate>0);

  // ── Gender breakdown ──────────────────────────────────────
  const genderMonthSS = {};
  fRec.forEach(r => {
    const m=(r.date||"").slice(0,7); if(!m) return;
    if (!genderMonthSS[m]) genderMonthSS[m]={label:monthLabel(m),male:0,female:0};
    genderMonthSS[m].male   += Number(r.male_present)||0;
    genderMonthSS[m].female += Number(r.female_present)||0;
  });
  const genderSSData = Object.values(genderMonthSS).sort((a,b)=>a.label>b.label?1:-1);

  const genderMonthCh = {};
  fChurch.forEach(r => {
    const m=(r.date||"").slice(0,7); if(!m) return;
    if (!genderMonthCh[m]) genderMonthCh[m]={label:monthLabel(m),male:0,female:0};
    genderMonthCh[m].male   += Number(r.male_closing)||0;
    genderMonthCh[m].female += Number(r.female_closing)||0;
  });
  const genderChData = Object.values(genderMonthCh).sort((a,b)=>a.label>b.label?1:-1);

  // ── Church by program ─────────────────────────────────────
  const progMap = {};
  fChurch.forEach(r => {
    const p = r.program||"Unknown";
    if (!progMap[p]) progMap[p] = { name:p, total:0, sessions:0 };
    progMap[p].total    += Number(r.total_closing)||0;
    progMap[p].sessions += 1;
  });
  const progData = Object.values(progMap).sort((a,b)=>b.total-a.total);

  // ── Retention trend ───────────────────────────────────────
  const retentionMonthly = monthlyData.map(m => ({
    label: m.label,
    rate: m.chClose > 0 ? Math.round(m.ssClose/m.chClose*100) : null,
  })).filter(m => m.rate !== null);

  // ── Aggregate KPIs ────────────────────────────────────────
  const ssTotal  = fRec.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const chTotal  = fChurch.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const avgRetention = retentionMonthly.length ? Math.round(retentionMonthly.reduce((s,m)=>s+m.rate,0)/retentionMonthly.length) : 0;
  const bibleTotal = fRec.reduce((s,r)=>s+(Number(r.bibles_closing)||0),0);
  const bibleRate  = ssTotal > 0 ? Math.round(bibleTotal/ssTotal*100) : 0;
  const ftTotal    = fRec.reduce((s,r)=>s+(Number(r.first_timers)||0),0) + fChurch.reduce((s,r)=>s+(Number(r.first_timers)||0),0);

  // ── Insight generator ─────────────────────────────────────
  const insights = [];
  if (avgRetention > 0) {
    if (avgRetention >= 80) insights.push({ text:`Excellent! An average of ${avgRetention}% of Sunday service attendees participate in Sunday School.`, color:t.success });
    else if (avgRetention >= 60) insights.push({ text:`Good engagement: ${avgRetention}% average SS retention from Sunday service. Aim for 80%+.`, color:t.warn });
    else insights.push({ text:`Only ${avgRetention}% of Sunday attendees join Sunday School on average. This is a key growth area.`, color:t.danger });
  }
  if (bibleRate >= 75) insights.push({ text:`Strong Bible culture: ${bibleRate}% of SS members bring their Bibles.`, color:t.success });
  else if (bibleRate > 0) insights.push({ text:`Bible participation at ${bibleRate}%. Encourage members to bring Bibles each week.`, color:t.warn });
  const topMonth = monthlyData.sort((a,b)=>b.chClose-a.chClose)[0];
  if (topMonth?.chClose > 0) insights.push({ text:`Best attended service: ${topMonth.label} with ${topMonth.chClose} at church closing.`, color:t.info });

  const ChartCard = ({ title, sub, children, span=1 }) => (
    <div style={{ ...card, gridColumn:span===2?"1 / -1":"auto" }}>
      <div style={{ marginBottom:14 }}>
        <div style={{ fontSize:14, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif" }}>{title}</div>
        {sub && <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:2 }}>{sub}</div>}
      </div>
      {children}
    </div>
  );

  const noData = (h=200) => (
    <div style={{ height:h, display:"flex", alignItems:"center", justifyContent:"center", color:t.textMuted, fontSize:13, fontFamily:"'Trebuchet MS',sans-serif", flexDirection:"column", gap:8 }}>
      <Icon name="info" size={24} color={t.textFaint} />
      Add more records to see this chart
    </div>
  );

  const tabStyle = (id) => ({
    padding:"8px 16px", borderRadius:8, border:"none", cursor:"pointer",
    fontFamily:"'Trebuchet MS',sans-serif", fontSize:12, fontWeight:700,
    background: tab===id ? t.gold : "transparent",
    color: tab===id ? "#0B1628" : t.textMuted,
    borderBottom: tab===id ? "none" : `1px solid ${t.border}`,
  });

  return (
    <div>
      {/* Header */}
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Analytics & Trends</div>
      <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:20 }}>
        Deep insights across Sunday School and Church attendance data
      </div>

      {/* Filter bar */}
      <div style={{ ...card, marginBottom:20, padding:"14px 18px" }}>
        <div style={{ fontSize:11, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:10 }}>
          🔍 Analytics Filters
        </div>
        <FilterBar filters={filters} setFilters={setFilters} months={allMonths}
          programs={programs} classes={classes} showClass={tab==="ss"} showProgram={tab==="church"} />
      </div>

      {/* KPI Summary row */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:14, marginBottom:24 }}>
        <KpiCard label="SS Total"        value={ssTotal}           sub="Closing count"   icon="attendance" color={t.gold} />
        <KpiCard label="Church Total"    value={chTotal}           sub="Closing count"   icon="cross"      color={t.info} />
        <KpiCard label="Avg Retention"   value={`${avgRetention}%`} sub="SS÷Church avg"  icon="analytics"  color={avgRetention>=80?t.success:avgRetention>=60?t.warn:t.danger} />
        <KpiCard label="Bible Rate"      value={`${bibleRate}%`}   sub="SS participation" icon="bible"    color="#9B59B6" />
        <KpiCard label="First Timers"    value={ftTotal}           sub="Combined"        icon="plus"       color="#E67E22" />
        <KpiCard label="Months of Data"  value={allMonths.length}  sub="On record"       icon="settings"   color={t.success} />
      </div>

      {/* Story insights */}
      {insights.length > 0 && (
        <div style={{ display:"flex", flexDirection:"column", gap:8, marginBottom:24 }}>
          {insights.map((ins,i) => <Insight key={i} text={ins.text} color={ins.color} icon="info" />)}
        </div>
      )}

      {/* Tab navigation */}
      <div style={{ display:"flex", gap:6, marginBottom:20, borderBottom:`1px solid ${t.border}`, paddingBottom:0 }}>
        {[
          {id:"comparison", label:"SS vs Church"},
          {id:"ss",         label:"Sunday School"},
          {id:"church",     label:"Church Services"},
          {id:"gender",     label:"Gender Breakdown"},
        ].map(tb => (
          <button key={tb.id} style={tabStyle(tb.id)} onClick={()=>setTab(tb.id)}>{tb.label}</button>
        ))}
      </div>

      {/* ── TAB: SS vs Church Comparison ── */}
      {tab === "comparison" && (
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
          <ChartCard title="SS vs Church Closing — by Date"
            sub="Each bar pair shows how many people closed Sunday School vs. the same-day church service" span={2}>
            {compData.length > 0 ? (
              <ResponsiveContainer width="100%" height={240}>
                <BarChart data={compData} barCategoryGap="20%">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                  <Bar dataKey="ssClose"  name="SS Closing"     fill={t.gold}  radius={[3,3,0,0]} />
                  <Bar dataKey="chClose"  name="Church Closing" fill={t.info}  radius={[3,3,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : noData(240)}
          </ChartCard>

          <ChartCard title="SS vs Church Beginning — by Date"
            sub="How many arrived at Sunday School vs. the church service at opening" span={2}>
            {compData.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={compData} barCategoryGap="20%">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                  <Bar dataKey="ssBegin"  name="SS Beginning"     fill="#9A7A2C" radius={[3,3,0,0]} />
                  <Bar dataKey="chBegin"  name="Church Beginning" fill="#4A9EDB88" radius={[3,3,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>

          <ChartCard title="Monthly SS & Church Closing Trend"
            sub="Month-by-month view of attendance growth for both gatherings">
            {monthlyData.length > 0 ? (
              <ResponsiveContainer width="100%" height={210}>
                <LineChart data={monthlyData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                  <Line type="monotone" dataKey="ssClose"  stroke={t.gold}    strokeWidth={2.5} dot={{ fill:t.gold, r:3 }} name="SS Closing" />
                  <Line type="monotone" dataKey="chClose"  stroke={t.info}    strokeWidth={2.5} dot={{ fill:t.info, r:3 }} name="Church Closing" />
                  <Line type="monotone" dataKey="ssBegin"  stroke={t.gold}    strokeWidth={1.5} dot={false} strokeDasharray="4 2" name="SS Beginning" />
                  <Line type="monotone" dataKey="chBegin"  stroke={t.info}    strokeWidth={1.5} dot={false} strokeDasharray="4 2" name="Church Beginning" />
                </LineChart>
              </ResponsiveContainer>
            ) : noData(210)}
          </ChartCard>

          <ChartCard title="SS Retention Rate — Monthly"
            sub="What % of people in Sunday Church Service also attended Sunday School that same day">
            {retentionMonthly.length > 0 ? (
              <ResponsiveContainer width="100%" height={210}>
                <AreaChart data={retentionMonthly}>
                  <defs>
                    <linearGradient id="gRet" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%"  stopColor={t.success} stopOpacity={0.3}/>
                      <stop offset="95%" stopColor={t.success} stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} unit="%" domain={[0,110]} />
                  <Tooltip {...tooltip} formatter={v=>`${v}%`} />
                  <Area type="monotone" dataKey="rate" stroke={t.success} fill="url(#gRet)" strokeWidth={2.5} name="Retention %" />
                </AreaChart>
              </ResponsiveContainer>
            ) : noData(210)}
          </ChartCard>
        </div>
      )}

      {/* ── TAB: Sunday School deep dive ── */}
      {tab === "ss" && (
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
          <ChartCard title="Monthly SS Attendance" sub="Total Sunday School closing count per month" span={2}>
            {monthlyData.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={monthlyData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Bar dataKey="ssClose" fill={t.gold} radius={[4,4,0,0]} name="SS Closing">
                    {monthlyData.map((_,i)=><Cell key={i} fill={CLASS_COLORS[i%CLASS_COLORS.length]}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>

          <ChartCard title="Classes by Month" sub="How each class contributes to Sunday School attendance over time">
            {classMonthlyData.length > 0 && classes.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={classMonthlyData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Legend wrapperStyle={{ fontSize:10, color:t.textMuted }} />
                  {classes.map((cls,i)=>(
                    <Bar key={cls.name} dataKey={cls.name} stackId="a" fill={CLASS_COLORS[i%CLASS_COLORS.length]} name={cls.name.split(" ")[0]} radius={i===classes.length-1?[3,3,0,0]:[0,0,0,0]} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>

          <ChartCard title="Bible Participation by Class" sub="Average % of members who brought a Bible (closing)">
            {bibleClass.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={bibleClass} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis type="number" stroke={t.textMuted} fontSize={10} unit="%" domain={[0,100]} />
                  <YAxis type="category" dataKey="name" stroke={t.textMuted} fontSize={10} width={70} />
                  <Tooltip {...tooltip} formatter={v=>`${v}%`} />
                  <Bar dataKey="rate" name="Bible Rate" radius={[0,4,4,0]}>
                    {bibleClass.map((e,i)=><Cell key={i} fill={e.fill}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>

          <ChartCard title="SS Beginning vs Closing" sub="How many started vs ended each month — drop-off indicator">
            {monthlyData.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={monthlyData} barCategoryGap="30%">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                  <Bar dataKey="ssBegin" name="Beginning" fill={t.gold+"88"} radius={[3,3,0,0]} />
                  <Bar dataKey="ssClose" name="Closing"   fill={t.gold}     radius={[3,3,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>
        </div>
      )}

      {/* ── TAB: Church Services deep dive ── */}
      {tab === "church" && (
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
          <ChartCard title="Church Attendance by Month" sub="Total closing attendance across all church services" span={2}>
            {monthlyData.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <AreaChart data={monthlyData}>
                  <defs>
                    <linearGradient id="gCM" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%"  stopColor={t.info} stopOpacity={0.3}/><stop offset="95%" stopColor={t.info} stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Area type="monotone" dataKey="chClose" stroke={t.info} fill="url(#gCM)" strokeWidth={2.5} name="Church Closing" />
                  <Area type="monotone" dataKey="chBegin" stroke={t.info} fill="transparent" strokeWidth={1.5} strokeDasharray="4 2" name="Church Beginning" />
                </AreaChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>

          <ChartCard title="Attendance by Program Type" sub="Which church programs draw the most people">
            {progData.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={progData} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis type="number" stroke={t.textMuted} fontSize={10} />
                  <YAxis type="category" dataKey="name" stroke={t.textMuted} fontSize={9} width={120} />
                  <Tooltip {...tooltip} />
                  <Bar dataKey="total" name="Total Attendance" radius={[0,4,4,0]}>
                    {progData.map((_,i)=><Cell key={i} fill={CLASS_COLORS[i%CLASS_COLORS.length]}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>

          <ChartCard title="Sessions per Program" sub="How frequently each program type is held">
            {progData.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={progData} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis type="number" stroke={t.textMuted} fontSize={10} />
                  <YAxis type="category" dataKey="name" stroke={t.textMuted} fontSize={9} width={120} />
                  <Tooltip {...tooltip} />
                  <Bar dataKey="sessions" name="Sessions" radius={[0,4,4,0]}>
                    {progData.map((_,i)=><Cell key={i} fill={CLASS_COLORS[(i+2)%CLASS_COLORS.length]}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>

          <ChartCard title="First Timers — Monthly" sub="Combined first timers from church and SS records">
            {monthlyData.length > 0 ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={monthlyData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Bar dataKey="ft" name="First Timers" fill="#E67E22" radius={[4,4,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : noData(220)}
          </ChartCard>
        </div>
      )}

      {/* ── TAB: Gender breakdown ── */}
      {tab === "gender" && (
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
          <ChartCard title="Sunday School — Male vs Female by Month"
            sub="Monthly gender breakdown of SS closing attendance">
            {genderSSData.length > 0 ? (
              <ResponsiveContainer width="100%" height={230}>
                <BarChart data={genderSSData} barCategoryGap="28%">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                  <Bar dataKey="male"   name="Male"   fill={t.info}    radius={[3,3,0,0]} />
                  <Bar dataKey="female" name="Female" fill="#E67E22"   radius={[3,3,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : noData(230)}
          </ChartCard>

          <ChartCard title="Church Service — Male vs Female by Month"
            sub="Monthly gender breakdown of church service closing count">
            {genderChData.length > 0 ? (
              <ResponsiveContainer width="100%" height={230}>
                <BarChart data={genderChData} barCategoryGap="28%">
                  <CartesianGrid strokeDasharray="3 3" stroke={t.textFaint} />
                  <XAxis dataKey="label" stroke={t.textMuted} fontSize={10} />
                  <YAxis stroke={t.textMuted} fontSize={11} />
                  <Tooltip {...tooltip} />
                  <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                  <Bar dataKey="male"   name="Male"   fill={t.info}    radius={[3,3,0,0]} />
                  <Bar dataKey="female" name="Female" fill="#E67E22"   radius={[3,3,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : noData(230)}
          </ChartCard>

          {/* Gender pie SS */}
          {(() => {
            const totalM = fRec.reduce((s,r)=>s+(Number(r.male_present)||0),0);
            const totalF = fRec.reduce((s,r)=>s+(Number(r.female_present)||0),0);
            const totalMC = fChurch.reduce((s,r)=>s+(Number(r.male_closing)||0),0);
            const totalFC = fChurch.reduce((s,r)=>s+(Number(r.female_closing)||0),0);
            return <>
              <ChartCard title="SS Gender Split" sub="Overall male vs female in Sunday School">
                {totalM+totalF > 0 ? (
                  <ResponsiveContainer width="100%" height={220}>
                    <PieChart>
                      <Pie data={[{name:"Male",value:totalM},{name:"Female",value:totalF}]}
                        cx="50%" cy="50%" innerRadius={50} outerRadius={85} dataKey="value" paddingAngle={4}>
                        <Cell fill={t.info}/><Cell fill="#E67E22"/>
                      </Pie>
                      <Tooltip {...tooltip} />
                      <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                    </PieChart>
                  </ResponsiveContainer>
                ) : noData(220)}
              </ChartCard>

              <ChartCard title="Church Gender Split" sub="Overall male vs female at church service">
                {totalMC+totalFC > 0 ? (
                  <ResponsiveContainer width="100%" height={220}>
                    <PieChart>
                      <Pie data={[{name:"Male",value:totalMC},{name:"Female",value:totalFC}]}
                        cx="50%" cy="50%" innerRadius={50} outerRadius={85} dataKey="value" paddingAngle={4}>
                        <Cell fill={t.info}/><Cell fill="#E67E22"/>
                      </Pie>
                      <Tooltip {...tooltip} />
                      <Legend wrapperStyle={{ fontSize:11, color:t.textMuted }} />
                    </PieChart>
                  </ResponsiveContainer>
                ) : noData(220)}
              </ChartCard>
            </>;
          })()}
        </div>
      )}
    </div>
  );
};

// ─── EXPORT PAGE ──────────────────────────────────────────────────────────────
const ExportPage = ({ db }) => {
  const { t, card, btnGold } = useThemeStyles();
  const { records, teachers, classes, churchRecs, programs, downloadWorkbook } = db;
  const fileRef = useRef();

  const exportCSV = () => {
    const headers = ["Date","Class","Teacher","Present","Bibles","First Timers","Visitors","Topic","Status"];
    const rows = records.map(r=>[r.date,r.class_name,r.teacher_name,r.total_closing,r.bibles_closing,r.first_timers,r.visitors,`"${(r.topic||"").replace(/"/g,"")}"`,r.status]);
    const csv = [headers,...rows].map(r=>r.join(",")).join("\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
    a.download = "attendance.csv"; a.click();
  };

  const exportChurchCSV = () => {
    const headers = ["Date","Day","Program","Male Begin","Female Begin","Total Begin","Male Close","Female Close","Total Close","First Timers","Visitors"];
    const rows = churchRecs.map(r=>[r.date,r.day_of_week,r.program,r.male_beginning,r.female_beginning,r.total_beginning,r.male_closing,r.female_closing,r.total_closing,r.first_timers,r.visitors]);
    const csv = [headers,...rows].map(r=>r.join(",")).join("\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
    a.download = "church_attendance.csv"; a.click();
  };

  const items = [
    { title:"Full Excel Workbook (.xlsx)", sub:`All data: SS records, Church, Teachers, Classes · ${records.length + churchRecs.length} total rows`, color:t.success, icon:"export", action:downloadWorkbook },
    { title:"SS Attendance CSV",           sub:`${records.length} Sunday School records`,      color:t.info,    icon:"export", action:exportCSV },
    { title:"Church Attendance CSV",       sub:`${churchRecs.length} church service records`,  color:"#9B59B6", icon:"export", action:exportChurchCSV },
  ];

  return (
    <div>
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Export Center</div>
      <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:24 }}>
        Download reports from your Supabase database
      </div>

      {/* Info banner */}
      <div style={{ ...card, marginBottom:24, display:"flex", gap:14, alignItems:"center", borderLeft:`3px solid ${t.gold}` }}>
        <Icon name="info" size={22} color={t.gold} />
        <div>
          <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>Data is stored in Supabase</div>
          <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:2 }}>
            All records sync across devices automatically. Export anytime to get a local Excel or CSV copy.
          </div>
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))", gap:16, marginBottom:24 }}>
        {items.map(e=>(
          <div key={e.title} onClick={e.action} style={{ ...card, cursor:"pointer", borderTop:`3px solid ${e.color}44` }}
            onMouseEnter={el=>el.currentTarget.style.borderTopColor=e.color}
            onMouseLeave={el=>el.currentTarget.style.borderTopColor=e.color+"44"}>
            <div style={{ width:48, height:48, background:e.color+"18", borderRadius:12, display:"flex", alignItems:"center", justifyContent:"center", marginBottom:14 }}>
              <Icon name={e.icon} size={22} color={e.color} />
            </div>
            <div style={{ fontSize:15, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>{e.title}</div>
            <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:14 }}>{e.sub}</div>
            <Badge label="Click to Download" color={e.color} />
          </div>
        ))}
      </div>

      {/* Live data summary */}
      <div style={card}>
        <div style={{ fontSize:14, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:14 }}>Database Summary</div>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
          {[
            { label:"SS Records",      count:records.length,    color:t.gold },
            { label:"Church Records",  count:churchRecs.length, color:"#E67E22" },
            { label:"Teachers",        count:teachers.length,   color:t.info },
            { label:"Classes",         count:classes.length,    color:t.success },
            { label:"Programs",        count:programs.length,   color:"#9B59B6" },
          ].map(s=>(
            <div key={s.label} style={{ padding:"10px 16px", borderRadius:10, border:`1px solid ${s.color}33`, background:s.color+"0D" }}>
              <div style={{ fontSize:13, fontWeight:700, color:s.color, fontFamily:"'Trebuchet MS',sans-serif" }}>{s.label}</div>
              <div style={{ fontSize:22, color:s.color, fontWeight:700, fontFamily:"'Georgia',serif" }}>{s.count}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─── CLASSES PAGE ─────────────────────────────────────────────────────────────
const ClassesPage = ({ db }) => {
  const { t, card } = useThemeStyles();
  const { classes, records } = db;
  return (
    <div>
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:20 }}>Class Management</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(210px,1fr))", gap:16 }}>
        {classes.map((cls,i)=>{
          const c = CLASS_COLORS[i%CLASS_COLORS.length];
          const recs = records.filter(r=>r.class_name===cls.name);
          const avgAtt = recs.length ? Math.round(recs.reduce((s,r)=>s+(Number(r.total_closing)||0),0)/recs.length) : 0;
          return (
            <div key={cls.name} style={{ ...card, borderTop:`3px solid ${c}` }}>
              <div style={{ fontSize:15, fontWeight:700, color:c, fontFamily:"'Georgia',serif", marginBottom:4 }}>{cls.name}</div>
              <div style={{ display:"flex", gap:14, marginTop:10 }}>
                <div>
                  <div style={{ fontSize:10, textTransform:"uppercase", letterSpacing:1, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:3 }}>Records</div>
                  <div style={{ fontSize:18, color:c, fontWeight:700, fontFamily:"'Georgia',serif" }}>{recs.length}</div>
                </div>
                <div>
                  <div style={{ fontSize:10, textTransform:"uppercase", letterSpacing:1, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:3 }}>Avg Att.</div>
                  <div style={{ fontSize:18, color:t.text, fontWeight:700, fontFamily:"'Georgia',serif" }}>{avgAtt}</div>
                </div>
                <div style={{ marginLeft:"auto", alignSelf:"flex-end" }}>
                  <span style={{ fontSize:10, padding:"3px 8px", borderRadius:20, background:t.success+"18", color:t.success, fontFamily:"'Trebuchet MS',sans-serif", fontWeight:700 }}>Active</span>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

// ─── CHURCH ATTENDANCE PAGE ───────────────────────────────────────────────────
const ChurchAttendancePage = ({ db, user }) => {
  const { t, card, btnGold, btnOutline, btnGhost, inp, sel, lbl, th, td } = useThemeStyles();
  const { churchRecs, programs, addChurchRec, updateChurchRec, deleteChurchRec } = db;

  const activePrograms = programs.filter(p => p.is_active === "YES")
    .sort((a,b) => Number(a.sort_order) - Number(b.sort_order));

  const DAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];

  const blank = {
    date: new Date().toISOString().slice(0,10),
    day_of_week: DAYS[new Date().getDay()],
    program: "",
    male_beginning: "", female_beginning: "",
    male_closing:   "", female_closing:   "",
    first_timers: "", visitors: "",
    notes: "", submitted_by: user?.name || "",
  };

  const [form, setForm]       = useState(blank);
  const [editId, setEditId]   = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [filter, setFilter]   = useState({ program:"", month:"" });
  const [loading, setLoading] = useState(false);

  // Stable handler — prevents focus loss on each keystroke
  const handleChange = useCallback((e) => {
    const { name, value } = e.target;
    setForm(f => ({ ...f, [name]: value }));
  }, []);

  // Auto-calc totals for display
  const totalBegin  = (Number(form.male_beginning)||0) + (Number(form.female_beginning)||0);
  const totalClose  = (Number(form.male_closing)||0)   + (Number(form.female_closing)||0);

  const handleSave = () => {
    if (!form.program) { alert("Please select a program."); return; }
    if (!form.date)    { alert("Please select a date."); return; }
    setLoading(true);
    setTimeout(() => {
      if (editId) { updateChurchRec(editId, form); setEditId(null); }
      else        { addChurchRec(form); }
      setForm(blank); setShowForm(false); setLoading(false);
    }, 300);
  };

  const handleEdit = (rec) => {
    setForm({ ...blank, ...rec });
    setEditId(rec.id);
    setShowForm(true);
    window.scrollTo({ top:0, behavior:"smooth" });
  };

  const handleCancel = () => { setForm(blank); setEditId(null); setShowForm(false); };

  const filtered = churchRecs.filter(r =>
    (!filter.program || r.program === filter.program) &&
    (!filter.month   || (r.date||"").slice(0,7) === filter.month)
  );

  // KPI totals
  const totalAtt   = churchRecs.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const totalMale  = churchRecs.reduce((s,r)=>s+(Number(r.male_closing)||0),0);
  const totalFemale= churchRecs.reduce((s,r)=>s+(Number(r.female_closing)||0),0);
  const totalFT    = churchRecs.reduce((s,r)=>s+(Number(r.first_timers)||0),0);

  // Unique months for filter
  const months = [...new Set(churchRecs.map(r=>(r.date||"").slice(0,7)).filter(Boolean))].sort().reverse();

  const numBox = useCallback((label, field, color) => (
    <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
      <label style={{ ...lbl, color }}>{label}</label>
      <input name={field} style={{ ...inp, borderColor: color+"55", color }} type="number" min="0"
        value={form[field]} onChange={handleChange} />
    </div>
  ), [form, handleChange, inp, lbl]);

  return (
    <div>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:24 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>
            ✝ Church Attendance
          </div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            Main church service attendance records — {churchRecs.length} entries
          </div>
        </div>
        {!showForm && (
          <button style={btnGold} onClick={()=>setShowForm(true)}>
            <span style={{ display:"flex", alignItems:"center", gap:7 }}>
              <Icon name="plus" size={16} color="#0B1628" /> Record Attendance
            </span>
          </button>
        )}
      </div>

      {/* KPI Row */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:14, marginBottom:24 }}>
        <KpiCard label="Total Attendance" value={totalAtt}    sub="All services" icon="attendance" color={t.gold} />
        <KpiCard label="Total Male"       value={totalMale}   sub="Closing count" icon="users"      color={t.info} />
        <KpiCard label="Total Female"     value={totalFemale} sub="Closing count" icon="users"      color="#E67E22" />
        <KpiCard label="First Timers"     value={totalFT}     sub="All time"      icon="plus"       color={t.success} />
        <KpiCard label="Services"         value={churchRecs.length} sub="Recorded" icon="classes"  color="#9B59B6" />
      </div>

      {/* Entry Form */}
      {showForm && (
        <div style={{ ...card, marginBottom:24, borderTop:`3px solid ${t.gold}` }}>
          <div style={{ fontSize:15, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:20 }}>
            {editId ? "✎ Edit Attendance Record" : "➕ New Attendance Record"}
          </div>

          {/* Row 1: Date, Day, Program */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 2fr", gap:14, marginBottom:16 }}>
            <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
              <label style={lbl}>Date *</label>
              <input style={inp} type="date" value={form.date} name="date" onChange={handleChange} />
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
              <label style={lbl}>Day</label>
              <select style={{ ...sel, width:"100%" }} value={form.day_of_week} name="day_of_week" onChange={handleChange}>
                {DAYS.map(d=><option key={d}>{d}</option>)}
              </select>
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
              <label style={lbl}>Church Program *</label>
              <select style={{ ...sel, width:"100%" }} value={form.program} name="program" onChange={handleChange}>
                <option value="">— Select Program —</option>
                {activePrograms.map(p=><option key={p.id} value={p.name}>{p.name}</option>)}
              </select>
            </div>
          </div>

          {/* Attendance grid: Beginning | Closing side by side */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
            {/* BEGINNING */}
            <div style={{ background:t.surfaceAlt, borderRadius:12, padding:16, border:`1px solid ${t.info}33` }}>
              <div style={{ fontSize:11, fontWeight:700, color:t.info, fontFamily:"'Trebuchet MS',sans-serif",
                textTransform:"uppercase", letterSpacing:1.5, marginBottom:14 }}>⏵ Beginning Count</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:12 }}>
                {numBox("Male", "male_beginning", t.info)}
                {numBox("Female", "female_beginning", "#E67E22")}
              </div>
              <div style={{ background:t.surface, borderRadius:8, padding:"10px 14px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>Total Beginning</span>
                <span style={{ fontSize:22, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif" }}>{totalBegin}</span>
              </div>
            </div>

            {/* CLOSING */}
            <div style={{ background:t.surfaceAlt, borderRadius:12, padding:16, border:`1px solid ${t.success}33` }}>
              <div style={{ fontSize:11, fontWeight:700, color:t.success, fontFamily:"'Trebuchet MS',sans-serif",
                textTransform:"uppercase", letterSpacing:1.5, marginBottom:14 }}>⏹ Closing Count</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:12 }}>
                {numBox("Male", "male_closing", t.info)}
                {numBox("Female", "female_closing", "#E67E22")}
              </div>
              <div style={{ background:t.surface, borderRadius:8, padding:"10px 14px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>Total Closing</span>
                <span style={{ fontSize:22, fontWeight:700, color:t.success, fontFamily:"'Georgia',serif" }}>{totalClose}</span>
              </div>
            </div>
          </div>

          {/* Extra fields */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14, marginBottom:16 }}>
            <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
              <label style={lbl}>First Timers</label>
              <input style={inp} type="number" min="0" value={form.first_timers} name="first_timers" onChange={handleChange} />
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
              <label style={lbl}>Visitors</label>
              <input style={inp} type="number" min="0" value={form.visitors} name="visitors" onChange={handleChange} />
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
              <label style={lbl}>Submitted By</label>
              <input style={inp} type="text" value={form.submitted_by} name="submitted_by" onChange={handleChange} />
            </div>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:5, marginBottom:20 }}>
            <label style={lbl}>Notes / Remarks</label>
            <textarea style={{ ...inp, minHeight:55, resize:"vertical" }} value={form.notes} name="notes" onChange={handleChange} />
          </div>

          <div style={{ display:"flex", gap:10 }}>
            <button style={{ ...btnGold, padding:"12px 32px" }} onClick={handleSave} disabled={loading}>
              {loading ? "Saving…" : editId ? "Save Changes" : "Save Record"}
            </button>
            <button style={{ ...btnOutline, padding:"12px 24px" }} onClick={handleCancel}>Cancel</button>
          </div>
        </div>
      )}

      {/* Filters */}
      <div style={{ ...card, marginBottom:16, display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
        <select style={{ ...sel }} value={filter.program} onChange={e=>setFilter(f=>({...f,program:e.target.value}))}>
          <option value="">All Programs</option>
          {activePrograms.map(p=><option key={p.id} value={p.name}>{p.name}</option>)}
        </select>
        <select style={{ ...sel }} value={filter.month} onChange={e=>setFilter(f=>({...f,month:e.target.value}))}>
          <option value="">All Months</option>
          {months.map(m=><option key={m} value={m}>{new Date(m+"-01").toLocaleDateString("en-GB",{month:"long",year:"numeric"})}</option>)}
        </select>
        <button onClick={()=>setFilter({program:"",month:""})} style={{ padding:"8px 14px", borderRadius:9, border:`1px solid ${t.border}`, background:"transparent", color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", cursor:"pointer", fontSize:13 }}>Clear</button>
        <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginLeft:"auto" }}>{filtered.length} records</span>
      </div>

      {/* Records Table */}
      <div style={{ ...card, padding:0, overflow:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:820 }}>
          <thead>
            <tr style={{ background:t.surfaceAlt }}>
              {["Date","Day","Program","Begin (M/F)","Begin Total","Closing (M/F)","Closing Total","1st Timers","Visitors","Actions"].map(h=>(
                <th key={h} style={th}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {[...filtered].reverse().map(r => (
              <tr key={r.id}
                onMouseEnter={e=>e.currentTarget.style.background=t.surfaceHover}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                <td style={td}>{fmtDate(r.date)}</td>
                <td style={{ ...td, color:t.textMuted }}>{r.day_of_week}</td>
                <td style={td}>
                  <span style={{ fontWeight:600, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif" }}>{r.program}</span>
                </td>
                {/* Beginning Male/Female */}
                <td style={td}>
                  <span style={{ color:t.info }}>{r.male_beginning||0}M</span>
                  <span style={{ color:t.textFaint, margin:"0 4px" }}>·</span>
                  <span style={{ color:"#E67E22" }}>{r.female_beginning||0}F</span>
                </td>
                <td style={td}>
                  <span style={{ fontWeight:700, color:t.text, fontFamily:"'Georgia',serif" }}>{r.total_beginning||0}</span>
                </td>
                {/* Closing Male/Female */}
                <td style={td}>
                  <span style={{ color:t.info }}>{r.male_closing||0}M</span>
                  <span style={{ color:t.textFaint, margin:"0 4px" }}>·</span>
                  <span style={{ color:"#E67E22" }}>{r.female_closing||0}F</span>
                </td>
                <td style={td}>
                  <span style={{ fontWeight:700, color:t.success, fontFamily:"'Georgia',serif" }}>{r.total_closing||0}</span>
                </td>
                <td style={{ ...td, color:t.textMuted }}>{r.first_timers||0}</td>
                <td style={{ ...td, color:t.textMuted }}>{r.visitors||0}</td>
                <td style={td}>
                  <div style={{ display:"flex", gap:5 }}>
                    <button style={btnGhost} title="Edit" onClick={()=>handleEdit(r)}>
                      <Icon name="edit" size={13} color={t.gold} />
                    </button>
                    <button style={btnGhost} title="Delete" onClick={()=>{
                      if(window.confirm(`Delete this ${r.program} record?`)) deleteChurchRec(r.id);
                    }}>
                      <Icon name="trash" size={13} color={t.danger} />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
            {!filtered.length && (
              <tr><td colSpan={10} style={{ ...td, textAlign:"center", padding:52, color:t.textMuted }}>
                No church attendance records yet. Click "Record Attendance" to add one.
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

// ─── PROGRAMS MANAGEMENT PAGE ─────────────────────────────────────────────────
const ProgramsPage = ({ db }) => {
  const { t, card, btnGold, btnOutline, btnGhost, inp, lbl } = useThemeStyles();
  const { programs, addProgram, updateProgram, deleteProgram, toggleProgramActive } = db;

  const [newName, setNewName]     = useState("");
  const [editId, setEditId]       = useState(null);
  const [editName, setEditName]   = useState("");

  const sorted = [...programs].sort((a,b)=>Number(a.sort_order)-Number(b.sort_order));
  const active = programs.filter(p=>p.is_active==="YES").length;

  const handleAdd = () => {
    if (!newName.trim()) return;
    if (programs.find(p=>p.name.toLowerCase()===newName.trim().toLowerCase())) {
      alert("A program with that name already exists."); return;
    }
    addProgram(newName.trim());
    setNewName("");
  };

  const handleSaveEdit = (id) => {
    if (!editName.trim()) return;
    updateProgram(id, { name: editName.trim() });
    setEditId(null); setEditName("");
  };

  return (
    <div>
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>
        Church Program Types
      </div>
      <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:24 }}>
        Manage the dropdown list of programs used in church attendance records.
        {" "}{active} active of {programs.length} total.
      </div>

      {/* Add new */}
      <div style={{ ...card, marginBottom:24, display:"flex", gap:12, alignItems:"flex-end", flexWrap:"wrap" }}>
        <div style={{ flex:1, minWidth:240 }}>
          <label style={{ ...lbl, display:"block", marginBottom:6 }}>New Program Name</label>
          <input style={inp} placeholder="e.g. Anniversary Sunday Service"
            value={newName} onChange={e=>setNewName(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&handleAdd()} />
        </div>
        <button style={{ ...btnGold, padding:"10px 24px" }} onClick={handleAdd}>
          <span style={{ display:"flex", alignItems:"center", gap:7 }}>
            <Icon name="plus" size={15} color="#0B1628" /> Add Program
          </span>
        </button>
      </div>

      {/* List */}
      <div style={{ ...card, padding:0, overflow:"hidden" }}>
        <div style={{ padding:"14px 20px", borderBottom:`1px solid ${t.border}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <span style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>All Programs</span>
          <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>{programs.length} total</span>
        </div>
        {sorted.map((prog, i) => {
          const isActive = prog.is_active === "YES";
          const isEditing = editId === prog.id;
          return (
            <div key={prog.id}
              style={{ display:"flex", alignItems:"center", gap:14, padding:"13px 20px",
                borderBottom:`1px solid ${t.border}18`,
                background: isEditing ? t.surfaceAlt : "transparent",
                transition:"background 0.15s" }}
              onMouseEnter={e=>{ if(!isEditing) e.currentTarget.style.background=t.surfaceHover; }}
              onMouseLeave={e=>{ if(!isEditing) e.currentTarget.style.background="transparent"; }}>

              {/* Drag handle / number */}
              <span style={{ fontSize:12, color:t.textFaint, fontFamily:"'Georgia',serif", minWidth:24, textAlign:"right" }}>{i+1}</span>

              {/* Name or edit input */}
              {isEditing ? (
                <input autoFocus style={{ ...inp, flex:1, padding:"7px 12px" }}
                  value={editName} onChange={e=>setEditName(e.target.value)}
                  onKeyDown={e=>{ if(e.key==="Enter") handleSaveEdit(prog.id); if(e.key==="Escape"){ setEditId(null); setEditName(""); } }} />
              ) : (
                <span style={{ flex:1, fontSize:14, color:isActive?t.text:t.textMuted,
                  fontFamily:"'Trebuchet MS',sans-serif", fontWeight:isActive?600:400,
                  textDecoration:isActive?"none":"line-through" }}>
                  {prog.name}
                </span>
              )}

              {/* Status badge */}
              <Badge label={isActive?"Active":"Inactive"} color={isActive?t.success:t.textMuted} />

              {/* Actions */}
              <div style={{ display:"flex", gap:5 }}>
                {isEditing ? (
                  <>
                    <button style={{ ...btnGold, padding:"6px 14px", fontSize:12 }} onClick={()=>handleSaveEdit(prog.id)}>Save</button>
                    <button style={{ ...btnOutline, padding:"6px 12px", fontSize:12 }} onClick={()=>{ setEditId(null); setEditName(""); }}>Cancel</button>
                  </>
                ) : (
                  <>
                    <button style={btnGhost} title="Edit name" onClick={()=>{ setEditId(prog.id); setEditName(prog.name); }}>
                      <Icon name="edit" size={13} color={t.gold} />
                    </button>
                    <button style={btnGhost} title={isActive?"Deactivate":"Activate"} onClick={()=>toggleProgramActive(prog.id)}>
                      <Icon name={isActive?"pause":"play"} size={13} color={isActive?t.warn:t.success} />
                    </button>
                    <button style={btnGhost} title="Delete" onClick={()=>{
                      if(window.confirm(`Delete "${prog.name}"? This cannot be undone.`)) deleteProgram(prog.id);
                    }}>
                      <Icon name="trash" size={13} color={t.danger} />
                    </button>
                  </>
                )}
              </div>
            </div>
          );
        })}
        {!programs.length && (
          <div style={{ padding:40, textAlign:"center", color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13 }}>
            No programs yet. Add one above.
          </div>
        )}
      </div>
    </div>
  );
};

// ─── USERS & PERMISSIONS PAGE ─────────────────────────────────────────────────
const UsersPage = ({ users: userHook }) => {
  const { t, card, btnGold, btnOutline, btnGhost, inp, sel, lbl, th, td } = useThemeStyles();
  const { appUsers, addUser, updateUser, deleteUser, toggleUserActive } = userHook;

  const [modal, setModal]   = useState(null); // null | "add" | user-obj
  const [editPerms, setEditPerms] = useState(null); // user being permissions-edited
  const [search, setSearch] = useState("");

  const filtered = appUsers.filter(u =>
    u.name?.toLowerCase().includes(search.toLowerCase()) ||
    u.email?.toLowerCase().includes(search.toLowerCase())
  );

  // ── Add/Edit User Form ──────────────────────────────────────
  const UserForm = ({ initial, onSave, onClose }) => {
    const blank = { name:"", email:"", password:"", role:"teacher", is_active:"YES",
      permissions: initial ? parsePerms(initial.permissions) : [...TEACHER_PERMS_DEFAULT] };
    const [form, setForm] = useState(initial
      ? { ...blank, ...initial, password: "", permissions: parsePerms(initial.permissions) }  // password always starts blank on edit
      : blank);
    const [err, setErr]   = useState("");
    const [showPw, setShowPw] = useState(false);

    // Stable handler — no focus loss
    const handleChange = useCallback((e) => {
      const { name, value } = e.target;
      setForm(f => ({ ...f, [name]: value }));
    }, []);

    const handleRoleChange = (e) => {
      const roleId = e.target.value;
      const allRoles = getAllRoles();
      const found = allRoles.find(r => r.id === roleId);
      setForm(f => ({
        ...f,
        role: roleId,
        permissions: found ? [...found.permissions] : [...TEACHER_PERMS_DEFAULT]
      }));
    };

    const togglePerm = (key) => {
      setForm(f => ({
        ...f,
        permissions: f.permissions.includes(key)
          ? f.permissions.filter(p=>p!==key)
          : [...f.permissions, key]
      }));
    };

    const handleSave = () => {
      if (!form.name.trim())  { setErr("Name is required."); return; }
      if (!form.email.trim()) { setErr("Email is required."); return; }
      if (!initial && !form.password.trim()) { setErr("Password is required for new users."); return; }
      onSave(form);
    };

    const groups = [...new Set(ALL_PERMISSIONS.map(p=>p.group))];
    const eyeBtn = { background:"transparent", border:"none", cursor:"pointer", padding:"0 8px", position:"absolute", right:8, top:"50%", transform:"translateY(-50%)" };

    return (
      <div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:20 }}>
          <div style={{ display:"flex", flexDirection:"column", gap:5, gridColumn:"1 / -1" }}>
            <label style={lbl}>Full Name *</label>
            <input name="name" style={inp} value={form.name} onChange={handleChange} placeholder="e.g. Bro. John Mensah" />
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
            <label style={lbl}>Email Address *</label>
            <input name="email" style={inp} type="email" value={form.email} onChange={handleChange} placeholder="email@uct.org" />
          </div>
          {/* Password with show/hide toggle */}
          <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
            <label style={lbl}>{initial ? "New Password (leave blank to keep)" : "Password *"}</label>
            <div style={{ position:"relative" }}>
              <input name="password" style={{ ...inp, paddingRight:40 }}
                type={showPw ? "text" : "password"}
                value={form.password} onChange={handleChange} placeholder="••••••••" />
              <button type="button" style={eyeBtn} onClick={()=>setShowPw(v=>!v)} title={showPw?"Hide":"Show"}>
                <Icon name="eye" size={16} color={t.textMuted} />
              </button>
            </div>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
            <label style={lbl}>Role</label>
            <select style={{ ...sel, width:"100%" }} value={form.role} onChange={handleRoleChange}>
              {getAllRoles().map(r => (
                <option key={r.id} value={r.id}>{r.name}{r.isBuiltIn ? "" : " (Custom)"}</option>
              ))}
            </select>
            <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
              Selecting a role auto-fills permissions below
            </div>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
            <label style={lbl}>Status</label>
            <select name="is_active" style={{ ...sel, width:"100%" }} value={form.is_active} onChange={handleChange}>
              <option value="YES">Active</option>
              <option value="NO">Inactive</option>
            </select>
          </div>
        </div>

        {/* Permission matrix */}
        <div style={{ background:t.surfaceAlt, borderRadius:12, padding:16, border:`1px solid ${t.border}`, marginBottom:20 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
            <div style={{ fontSize:12, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2 }}>
              Page & Feature Access
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <button onClick={()=>setForm(f=>({...f,permissions:[...ADMIN_PERMS]}))} style={{ fontSize:11, padding:"4px 10px", borderRadius:6, border:`1px solid ${t.success}44`, background:t.success+"11", color:t.success, cursor:"pointer", fontFamily:"'Trebuchet MS',sans-serif" }}>Grant All</button>
              <button onClick={()=>setForm(f=>({...f,permissions:[]}))} style={{ fontSize:11, padding:"4px 10px", borderRadius:6, border:`1px solid ${t.danger}44`, background:t.danger+"11", color:t.danger, cursor:"pointer", fontFamily:"'Trebuchet MS',sans-serif" }}>Revoke All</button>
            </div>
          </div>
          {groups.map(group => (
            <div key={group} style={{ marginBottom:14 }}>
              <div style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:8, borderBottom:`1px solid ${t.border}`, paddingBottom:4 }}>{group}</div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))", gap:8 }}>
                {ALL_PERMISSIONS.filter(p=>p.group===group).map(perm => {
                  const on = form.permissions.includes(perm.key);
                  return (
                    <label key={perm.key} onClick={()=>togglePerm(perm.key)}
                      style={{ display:"flex", alignItems:"center", gap:9, padding:"8px 12px", borderRadius:8, cursor:"pointer",
                        background: on ? t.gold+"14" : t.surface,
                        border: `1px solid ${on ? t.gold+"55" : t.border}`,
                        transition:"all 0.15s" }}>
                      <div style={{ width:16, height:16, borderRadius:4, border:`2px solid ${on?t.gold:t.textFaint}`,
                        background: on ? t.gold : "transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                        {on && <Icon name="check" size={10} color="#FFFFFF" />}
                      </div>
                      <span style={{ fontSize:12, color:on?t.text:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>{perm.label}</span>
                    </label>
                  );
                })}
              </div>
            </div>
          ))}
        </div>

        {err && <div style={{ color:t.danger, fontSize:12, marginBottom:12, padding:"8px 12px", background:t.danger+"11", borderRadius:7 }}>{err}</div>}

        {/* Password hint for edit mode */}
        {initial && (
          <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
            padding:"8px 12px", background:t.surfaceAlt, borderRadius:7, marginBottom:12,
            border:`1px solid ${t.border}` }}>
            💡 Leave the password field blank to keep the existing password unchanged.
          </div>
        )}

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
          <button style={{ ...btnOutline, padding:"10px 20px" }} onClick={onClose}>Cancel</button>
          {initial ? (
            <button style={{ ...btnGold, padding:"10px 28px", background:"linear-gradient(135deg,#0A5A9C,#1A7DC8)" }} onClick={handleSave}>
              ✓ Update User
            </button>
          ) : (
            <button style={{ ...btnGold, padding:"10px 28px" }} onClick={handleSave}>
              + Save New User
            </button>
          )}
        </div>
      </div>
    );
  };

  const handleSave = (form) => {
    if (modal === "add") {
      const result = addUser(form);
      if (result?.error) { alert(result.error); return; }
      alert(`✓ User created!\n\nLogin details:\nEmail: ${form.email.toLowerCase().trim()}\nPassword: ${form.password.trim()}`);
    } else {
      updateUser(modal.id, form);
      const pwMsg = form.password?.trim()
        ? `\nPassword changed to: ${form.password.trim()}`
        : "\nPassword: unchanged";
      alert(`✓ User updated!\n\nEmail: ${form.email.toLowerCase().trim()}${pwMsg}`);
    }
    setModal(null);
  };

  const activeCount = appUsers.filter(u=>u.is_active==="YES").length;
  const adminCount  = appUsers.filter(u=>u.role==="admin").length;
  const [showCredentials, setShowCredentials] = useState(false);
  const storedUsers = loadUsers();

  return (
    <div>
      {modal && (
        <Modal title={modal==="add"?"Create New User":`Edit User — ${modal.name}`} onClose={()=>setModal(null)} width={680}>
          <UserForm initial={modal==="add"?null:modal} onSave={handleSave} onClose={()=>setModal(null)} />
        </Modal>
      )}

      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:24 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Users & Access Control</div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            Manage accounts and assign page/feature permissions per user
          </div>
        </div>
        <button style={btnGold} onClick={()=>setModal("add")}>
          <span style={{ display:"flex", alignItems:"center", gap:7 }}><Icon name="plus" size={16} color="#0B1628" />Add User</span>
        </button>
      </div>

      {/* KPIs */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:14, marginBottom:24 }}>
        <KpiCard label="Total Users"   value={appUsers.length} sub="All accounts"     icon="users"    color={t.gold} />
        <KpiCard label="Active"        value={activeCount}     sub="Can log in"        icon="play"     color={t.success} />
        <KpiCard label="Inactive"      value={appUsers.length-activeCount} sub="Locked out" icon="pause" color={t.textMuted} />
        <KpiCard label="Admins"        value={adminCount}      sub="Full access"       icon="settings" color={t.info} />
      </div>

      {/* Login Credentials Panel */}
      <div style={{ ...card, marginBottom:20, borderLeft:`3px solid ${t.gold}` }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div>
            <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>
              Login Credentials (Admin View)
            </div>
            <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:2 }}>
              These are the exact emails and passwords stored — what users must type to log in
            </div>
          </div>
          <button onClick={()=>setShowCredentials(v=>!v)}
            style={{ padding:"7px 14px", borderRadius:8, border:`1px solid ${t.border}`,
              background:showCredentials?t.gold:"transparent",
              color:showCredentials?"#0B1628":t.textMuted,
              fontFamily:"'Trebuchet MS',sans-serif", fontSize:12, cursor:"pointer" }}>
            {showCredentials ? "Hide" : "Show Credentials"}
          </button>
        </div>
        {showCredentials && (
          <div style={{ marginTop:14, overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontFamily:"'Trebuchet MS',sans-serif", fontSize:13 }}>
              <thead>
                <tr style={{ background:t.surfaceAlt }}>
                  {["Name","Email (type exactly)","Password (type exactly)","Role","Status"].map(h=>(
                    <th key={h} style={{ padding:"9px 14px", textAlign:"left", borderBottom:`1px solid ${t.border}`,
                      color:t.gold, fontSize:10, textTransform:"uppercase", letterSpacing:1 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {storedUsers.map(u => (
                  <tr key={u.id}>
                    <td style={{ padding:"10px 14px", borderBottom:`1px solid ${t.border}18`, color:t.text }}>{u.name}</td>
                    <td style={{ padding:"10px 14px", borderBottom:`1px solid ${t.border}18` }}>
                      <code style={{ background:t.surfaceAlt, padding:"3px 8px", borderRadius:5, fontSize:12, color:t.info }}>{u.email}</code>
                    </td>
                    <td style={{ padding:"10px 14px", borderBottom:`1px solid ${t.border}18` }}>
                      <code style={{ background:t.surfaceAlt, padding:"3px 8px", borderRadius:5, fontSize:12, color:t.success }}>{u.password}</code>
                    </td>
                    <td style={{ padding:"10px 14px", borderBottom:`1px solid ${t.border}18` }}>
                      <Badge label={getAllRoles().find(r=>r.id===u.role)?.name || u.role} color={u.role==="admin"?t.gold:t.info} />
                    </td>
                    <td style={{ padding:"10px 14px", borderBottom:`1px solid ${t.border}18` }}>
                      <Badge label={u.is_active==="YES"?"Active":"Inactive"} color={u.is_active==="YES"?t.success:t.textMuted} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Search */}
      <div style={{ ...card, marginBottom:16, display:"flex", gap:12, alignItems:"center" }}>
        <input style={{ ...inp, maxWidth:260 }} placeholder="Search by name or email…" value={search} onChange={e=>setSearch(e.target.value)} />
        <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginLeft:"auto" }}>{filtered.length} users</span>
      </div>

      {/* Users Table */}
      <div style={{ ...card, padding:0, overflow:"hidden" }}>
        <table style={{ width:"100%", borderCollapse:"collapse" }}>
          <thead>
            <tr style={{ background:t.surfaceAlt }}>
              {["User","Email","Role","Permissions","Status","Actions"].map(h=><th key={h} style={th}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {filtered.map((u) => {
              const perms = parsePerms(u.permissions);
              const isActive = u.is_active === "YES";
              const isAdmin  = u.role === "admin";
              return (
                <tr key={u.id}
                  onMouseEnter={e=>e.currentTarget.style.background=t.surfaceHover}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={td}>
                    <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                      <Avatar name={u.name} size={36} color={isAdmin?t.gold:isActive?t.info:t.textFaint} />
                      <div>
                        <div style={{ fontWeight:600, color:isActive?t.text:t.textMuted, fontSize:13, fontFamily:"'Trebuchet MS',sans-serif" }}>{u.name}</div>
                        <div style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>ID: {u.id}</div>
                      </div>
                    </div>
                  </td>
                  <td style={{ ...td, color:t.textMuted, fontSize:12 }}>{u.email}</td>
                  <td style={td}>{(() => {
                    const allRoles = getAllRoles();
                    const found = allRoles.find(r => r.id === u.role);
                    const label = found?.name || u.role;
                    const color = found?.color || (u.role==="admin" ? t.gold : t.info);
                    return <Badge label={label} color={color} />;
                  })()}</td>
                  <td style={td}>
                    {isAdmin ? (
                      <span style={{ fontSize:11, color:t.success, fontFamily:"'Trebuchet MS',sans-serif" }}>Full Access ({perms.length})</span>
                    ) : (
                      <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                        {perms.slice(0,4).map(pk => {
                          const pm = ALL_PERMISSIONS.find(x=>x.key===pk);
                          return pm ? <span key={pk} style={{ fontSize:10, padding:"2px 7px", borderRadius:12, background:t.info+"14", color:t.info, fontFamily:"'Trebuchet MS',sans-serif" }}>{pm.label}</span> : null;
                        })}
                        {perms.length > 4 && <span style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>+{perms.length-4} more</span>}
                        {perms.length === 0 && <span style={{ fontSize:11, color:t.danger, fontFamily:"'Trebuchet MS',sans-serif" }}>No access</span>}
                      </div>
                    )}
                  </td>
                  <td style={td}><Badge label={isActive?"Active":"Inactive"} color={isActive?t.success:t.textMuted}/></td>
                  <td style={td}>
                    <div style={{ display:"flex", gap:5 }}>
                      <button style={btnGhost} title="Edit user & permissions" onClick={()=>setModal(u)}><Icon name="edit" size={13} color={t.gold}/></button>
                      <button style={btnGhost} title={isActive?"Deactivate (lock out)":"Activate"} onClick={()=>toggleUserActive(u.id)}>
                        <Icon name={isActive?"pause":"play"} size={13} color={isActive?t.warn:t.success}/>
                      </button>
                      <button style={btnGhost} title="Delete permanently" onClick={()=>{
                        if(u.role==="admin" && appUsers.filter(x=>x.role==="admin"&&x.is_active==="YES").length<=1){
                          alert("Cannot delete the last active admin account."); return;
                        }
                        if(window.confirm(`Permanently delete ${u.name}?`)) deleteUser(u.id);
                      }}><Icon name="trash" size={13} color={t.danger}/></button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {!filtered.length && (
              <tr><td colSpan={6} style={{ ...td, textAlign:"center", padding:48, color:t.textMuted }}>No users found</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Permission legend */}
      <div style={{ ...card, marginTop:20 }}>
        <div style={{ fontSize:13, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:14 }}>Permission Reference</div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
          {[...new Set(ALL_PERMISSIONS.map(p=>p.group))].map(grp => (
            <div key={grp}>
              <div style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:8 }}>{grp}</div>
              {ALL_PERMISSIONS.filter(p=>p.group===grp).map(p=>(
                <div key={p.key} style={{ display:"flex", gap:8, padding:"4px 0", borderBottom:`1px solid ${t.border}12` }}>
                  <code style={{ fontSize:10, color:t.info, fontFamily:"monospace", minWidth:140 }}>{p.key}</code>
                  <span style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>{p.label}</span>
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─── ROLES PAGE ───────────────────────────────────────────────────────────────
const RolesPage = () => {
  const { t, card, btnGold, btnOutline, btnGhost, inp, lbl } = useThemeStyles();
  const [roles, setRoles]         = useState(getAllRoles);
  const [modal, setModal]         = useState(null); // null | "add" | role-obj
  const [expandedId, setExpandedId] = useState(null);

  const customRoles = roles.filter(r => !r.isBuiltIn);
  const builtInRoles = roles.filter(r => r.isBuiltIn);

  // ── Role Form ──────────────────────────────────────────────
  const RoleForm = ({ initial, onSave, onClose }) => {
    const ROLE_COLORS = ["#004b23","#1565C0","#9B59B6","#E67E22","#E74C3C","#1ABC9C","#C87A0A","#2ECC71"];
    const blank = { name:"", description:"", color:"#1565C0", permissions:[] };
    const [form, setForm] = useState(initial ? { ...blank, ...initial } : blank);
    const [err, setErr]   = useState("");

    const handleChange = useCallback((e) => {
      const { name, value } = e.target;
      setForm(f => ({ ...f, [name]: value }));
    }, []);

    const togglePerm = (key) => {
      setForm(f => ({
        ...f,
        permissions: f.permissions.includes(key)
          ? f.permissions.filter(p => p !== key)
          : [...f.permissions, key]
      }));
    };

    const handleSave = () => {
      if (!form.name.trim()) { setErr("Role name is required."); return; }
      const allR = getAllRoles();
      const dup = allR.find(r => r.name.toLowerCase() === form.name.trim().toLowerCase() && r.id !== initial?.id);
      if (dup) { setErr("A role with this name already exists."); return; }
      onSave({ ...form, name: form.name.trim(), description: form.description.trim() });
    };

    const groups = [...new Set(ALL_PERMISSIONS.map(p => p.group))];

    return (
      <div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:20 }}>
          <div style={{ display:"flex", flexDirection:"column", gap:5, gridColumn:"1 / -1" }}>
            <label style={lbl}>Role Name *</label>
            <input name="name" style={inp} value={form.name} onChange={handleChange} placeholder="e.g. Class Monitor, Secretary" />
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:5, gridColumn:"1 / -1" }}>
            <label style={lbl}>Description</label>
            <input name="description" style={inp} value={form.description} onChange={handleChange} placeholder="What does this role do?" />
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
            <label style={lbl}>Colour</label>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginTop:4 }}>
              {ROLE_COLORS.map(c => (
                <div key={c} onClick={()=>setForm(f=>({...f,color:c}))}
                  style={{ width:28, height:28, borderRadius:"50%", background:c, cursor:"pointer",
                    border: form.color===c ? `3px solid ${t.text}` : "3px solid transparent",
                    transition:"transform 0.1s", transform: form.color===c ? "scale(1.2)" : "scale(1)" }} />
              ))}
            </div>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
            <label style={lbl}>Preview</label>
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <span style={{ display:"inline-flex", alignItems:"center", padding:"5px 14px", borderRadius:20,
                fontSize:13, fontWeight:700, fontFamily:"'Trebuchet MS',sans-serif",
                background:form.color+"22", color:form.color, border:`1px solid ${form.color}44` }}>
                {form.name || "Role Name"}
              </span>
            </div>
          </div>
        </div>

        {/* Permission matrix */}
        <div style={{ background:t.surfaceAlt, borderRadius:12, padding:16, border:`1px solid ${t.border}`, marginBottom:16 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
            <div style={{ fontSize:12, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2 }}>
              Permissions for this Role
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <button onClick={()=>setForm(f=>({...f,permissions:[...ADMIN_PERMS]}))}
                style={{ fontSize:11, padding:"4px 10px", borderRadius:6, border:`1px solid ${t.success}44`, background:t.success+"11", color:t.success, cursor:"pointer", fontFamily:"'Trebuchet MS',sans-serif" }}>
                Select All
              </button>
              <button onClick={()=>setForm(f=>({...f,permissions:[]}))}
                style={{ fontSize:11, padding:"4px 10px", borderRadius:6, border:`1px solid ${t.danger}44`, background:t.danger+"11", color:t.danger, cursor:"pointer", fontFamily:"'Trebuchet MS',sans-serif" }}>
                Clear All
              </button>
            </div>
          </div>
          {groups.map(group => (
            <div key={group} style={{ marginBottom:12 }}>
              <div style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
                textTransform:"uppercase", letterSpacing:1.2, marginBottom:8,
                borderBottom:`1px solid ${t.border}`, paddingBottom:4 }}>{group}</div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))", gap:7 }}>
                {ALL_PERMISSIONS.filter(p => p.group===group).map(perm => {
                  const on = form.permissions.includes(perm.key);
                  return (
                    <label key={perm.key} onClick={()=>togglePerm(perm.key)}
                      style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 10px",
                        borderRadius:8, cursor:"pointer",
                        background: on ? form.color+"18" : t.surface,
                        border: `1px solid ${on ? form.color+"66" : t.border}`,
                        transition:"all 0.15s" }}>
                      <div style={{ width:15, height:15, borderRadius:4,
                        border:`2px solid ${on ? form.color : t.textFaint}`,
                        background: on ? form.color : "transparent",
                        display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                        {on && <Icon name="check" size={9} color="#FFFFFF" />}
                      </div>
                      <span style={{ fontSize:11, color:on?t.text:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>{perm.label}</span>
                    </label>
                  );
                })}
              </div>
            </div>
          ))}
        </div>

        <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
          padding:"8px 12px", background:t.surfaceAlt, borderRadius:7, marginBottom:16 }}>
          💡 {form.permissions.length} permission{form.permissions.length!==1?"s":""} selected
        </div>

        {err && <div style={{ color:t.danger, fontSize:12, marginBottom:12,
          padding:"8px 12px", background:t.danger+"11", borderRadius:7 }}>{err}</div>}

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
          <button style={{ ...btnOutline, padding:"10px 20px" }} onClick={onClose}>Cancel</button>
          {initial ? (
            <button style={{ ...btnGold, padding:"10px 28px", background:"linear-gradient(135deg,#0A5A9C,#1A7DC8)" }} onClick={handleSave}>
              ✓ Update Role
            </button>
          ) : (
            <button style={{ ...btnGold, padding:"10px 28px" }} onClick={handleSave}>
              + Create Role
            </button>
          )}
        </div>
      </div>
    );
  };

  const handleSave = (form) => {
    let updated;
    if (modal === "add") {
      const newRole = { ...form, id:`ROLE_${Date.now()}`, isBuiltIn: false };
      const customs = [...loadCustomRoles(), newRole];
      saveCustomRoles(customs);
      updated = [...BUILT_IN_ROLES, ...customs];
    } else {
      const customs = loadCustomRoles().map(r => r.id===modal.id ? { ...r, ...form } : r);
      saveCustomRoles(customs);
      updated = [...BUILT_IN_ROLES, ...customs];
    }
    setRoles(updated);
    setModal(null);
  };

  const handleDelete = (id) => {
    if (!window.confirm("Delete this role? Users assigned to it will keep their current permissions.")) return;
    const customs = loadCustomRoles().filter(r => r.id !== id);
    saveCustomRoles(customs);
    setRoles([...BUILT_IN_ROLES, ...customs]);
  };

  const RoleCard = ({ role }) => {
    const expanded = expandedId === role.id;
    return (
      <div style={{ ...card, borderLeft:`4px solid ${role.color}`, padding:0, overflow:"hidden" }}>
        <div style={{ padding:"16px 20px", display:"flex", alignItems:"center", gap:14, cursor:"pointer" }}
          onClick={()=>setExpandedId(expanded ? null : role.id)}>
          {/* Color dot */}
          <div style={{ width:36, height:36, borderRadius:"50%", background:role.color+"22",
            border:`2px solid ${role.color}55`, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
            <Icon name={role.isBuiltIn?"settings":"users"} size={16} color={role.color} />
          </div>
          <div style={{ flex:1 }}>
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <span style={{ fontSize:15, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>{role.name}</span>
              {role.isBuiltIn && <Badge label="Built-in" color={t.textMuted} />}
              <Badge label={`${role.permissions.length} permissions`} color={role.color} />
            </div>
            {role.description && (
              <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:3 }}>{role.description}</div>
            )}
          </div>
          {!role.isBuiltIn && (
            <div style={{ display:"flex", gap:6 }} onClick={e=>e.stopPropagation()}>
              <button style={btnGhost} title="Edit" onClick={()=>setModal(role)}>
                <Icon name="edit" size={14} color={t.gold} />
              </button>
              <button style={btnGhost} title="Delete" onClick={()=>handleDelete(role.id)}>
                <Icon name="trash" size={14} color={t.danger} />
              </button>
            </div>
          )}
          <div style={{ color:t.textMuted, fontSize:18, marginLeft:4 }}>{expanded?"▲":"▼"}</div>
        </div>

        {/* Expanded permissions */}
        {expanded && (
          <div style={{ padding:"0 20px 16px", borderTop:`1px solid ${t.border}` }}>
            <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
              textTransform:"uppercase", letterSpacing:1, marginBottom:10, marginTop:12 }}>
              Permissions Included
            </div>
            <div style={{ display:"flex", flexWrap:"wrap", gap:7 }}>
              {role.permissions.length > 0
                ? role.permissions.map(pk => {
                    const pm = ALL_PERMISSIONS.find(x => x.key===pk);
                    return pm ? (
                      <span key={pk} style={{ fontSize:11, padding:"3px 10px", borderRadius:20,
                        background:role.color+"14", color:role.color,
                        border:`1px solid ${role.color}33`, fontFamily:"'Trebuchet MS',sans-serif" }}>
                        {pm.label}
                      </span>
                    ) : null;
                  })
                : <span style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>No permissions assigned</span>
              }
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div>
      {modal && (
        <Modal
          title={modal==="add" ? "Create New Role" : `Edit Role — ${modal.name}`}
          onClose={()=>setModal(null)}
          width={680}>
          <RoleForm
            initial={modal==="add" ? null : modal}
            onSave={handleSave}
            onClose={()=>setModal(null)} />
        </Modal>
      )}

      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:24 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>
            Roles Management
          </div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            Define custom roles with specific permissions. Assign roles to users in Users & Access.
          </div>
        </div>
        <button style={btnGold} onClick={()=>setModal("add")}>
          <span style={{ display:"flex", alignItems:"center", gap:7 }}>
            <Icon name="plus" size={16} color="#FFFFFF" /> Create Role
          </span>
        </button>
      </div>

      {/* KPIs */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:14, marginBottom:24 }}>
        <KpiCard label="Total Roles"    value={roles.length}       sub="All roles"        icon="users"    color={t.gold} />
        <KpiCard label="Built-in"       value={builtInRoles.length} sub="Cannot be deleted" icon="settings" color={t.info} />
        <KpiCard label="Custom"         value={customRoles.length} sub="Created by admin" icon="edit"     color={t.success} />
        <KpiCard label="Permissions"    value={ALL_PERMISSIONS.length} sub="Available"    icon="check"    color="#9B59B6" />
      </div>

      {/* Info banner */}
      <div style={{ ...card, marginBottom:24, display:"flex", gap:12, alignItems:"flex-start", borderLeft:`3px solid ${t.info}` }}>
        <Icon name="info" size={18} color={t.info} />
        <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", lineHeight:1.7 }}>
          <strong style={{ color:t.text }}>How roles work:</strong> Create a role here with the permissions you want,
          then go to <strong style={{ color:t.text }}>Users & Access</strong> to assign it to a user.
          Built-in roles (Admin, Teacher, Superintendent, Asst. Teacher) cannot be deleted but custom roles can be edited freely.
        </div>
      </div>

      {/* Built-in roles */}
      <div style={{ fontSize:12, fontWeight:700, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
        textTransform:"uppercase", letterSpacing:1.2, marginBottom:12 }}>Built-in Roles</div>
      <div style={{ display:"flex", flexDirection:"column", gap:10, marginBottom:28 }}>
        {builtInRoles.map(r => <RoleCard key={r.id} role={r} />)}
      </div>

      {/* Custom roles */}
      <div style={{ fontSize:12, fontWeight:700, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
        textTransform:"uppercase", letterSpacing:1.2, marginBottom:12 }}>Custom Roles</div>
      {customRoles.length > 0 ? (
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          {customRoles.map(r => <RoleCard key={r.id} role={r} />)}
        </div>
      ) : (
        <div style={{ ...card, textAlign:"center", padding:40, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
          No custom roles yet. Click <strong style={{ color:t.gold }}>"Create Role"</strong> to add one — e.g. Secretary, Class Monitor, Treasurer.
        </div>
      )}
    </div>
  );
};

// ─── SIDEBAR ──────────────────────────────────────────────────────────────────
const Sidebar = ({ page, setPage, user, onLogout, collapsed, setCollapsed }) => {
  const { dark, toggle } = useTheme();
  const t = T[dark?"dark":"light"];
  const perms = user?.permissions || [];
  const can = (key) => user?.role==="admin" || perms.includes(key);

  const allAdminNav = [
    {id:"dashboard",  label:"Dashboard",      icon:"dashboard",  perm:"dashboard"},
    {id:"attendance", label:"SS Records",     icon:"attendance", perm:"attendance"},
    {id:"church",     label:"Church Attend.", icon:"cross",      perm:"church"},
    {id:"analytics",  label:"Analytics",      icon:"analytics",  perm:"analytics"},
    {id:"teachers",   label:"Teachers",       icon:"teachers",   perm:"teachers"},
    {id:"classes",    label:"Classes",        icon:"classes",    perm:"classes"},
    {id:"programs",   label:"Programs",       icon:"settings",   perm:"programs"},
    {id:"users",      label:"Users & Access", icon:"users",      perm:"__admin_only__"},
    {id:"roles",      label:"Roles",          icon:"edit",       perm:"__admin_only__"},
    {id:"branding",   label:"Logo & Name",    icon:"bible",      perm:"__admin_only__"},
    {id:"export",     label:"Export",         icon:"export",     perm:"export"},
  ];
  const allTeacherNav = [
    {id:"submit",     label:"Submit Report",     icon:"submit",     perm:"submit"},
    {id:"church",     label:"Church Attendance", icon:"cross",      perm:"church"},
    {id:"attendance", label:"My Records",        icon:"attendance", perm:"attendance"},
    {id:"analytics",  label:"Analytics",         icon:"analytics",  perm:"analytics"},
    {id:"export",     label:"Export",            icon:"export",     perm:"export"},
  ];

  const isAdmin = user?.role === "admin";
  const rawNav  = isAdmin ? allAdminNav : allTeacherNav;
  // Filter nav items by permission; admins always see everything
  const nav = rawNav.filter(n => {
    if (n.perm === "__admin_only__") return isAdmin;
    return isAdmin || perms.includes(n.perm);
  });

  const navItem = (active) => ({
    display:"flex", alignItems:"center", gap:10, padding:"10px 18px", cursor:"pointer",
    borderLeft: active ? `3px solid rgba(255,255,255,0.9)` : "3px solid transparent",
    background: active ? t.sidebarActive : "transparent",
    color: active ? "#FFFFFF" : t.sidebarMuted,
    fontSize:13, fontFamily:"'Trebuchet MS',sans-serif", transition:"all 0.15s", whiteSpace:"nowrap",
  });
  return (
    <div style={{ width:collapsed?60:230, background:t.sidebar, borderRight:`1px solid ${t.sidebarBorder}`,
      display:"flex", flexDirection:"column", flexShrink:0, transition:"width 0.22s", overflow:"hidden", minHeight:"100vh" }}>
      <div style={{ padding:collapsed?"18px 12px":"22px 18px 18px", borderBottom:`1px solid ${t.sidebarBorder}` }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <ChurchLogo size={collapsed?34:36} />
          {!collapsed && (
            <div>
              <div style={{ fontSize:9.5, fontWeight:700, color:"#FFFFFF", fontFamily:"'Trebuchet MS',sans-serif", letterSpacing:1.2, lineHeight:1.4 }}>{CHURCH_NAME.toUpperCase()}</div>
              <div style={{ fontSize:8.5, color:"rgba(255,255,255,0.55)", fontFamily:"'Trebuchet MS',sans-serif", letterSpacing:0.5 }}>SSM SYSTEM</div>
            </div>
          )}
        </div>
      </div>
      <nav style={{ flex:1, paddingTop:8 }}>
        {nav.map(n => (
          <div key={n.id} style={navItem(page===n.id)} onClick={()=>setPage(n.id)} title={collapsed?n.label:""}
            onMouseEnter={e=>{ if(page!==n.id) e.currentTarget.style.background="rgba(255,255,255,0.08)"; }}
            onMouseLeave={e=>{ if(page!==n.id) e.currentTarget.style.background="transparent"; }}>
            <Icon name={n.icon} size={17} color={page===n.id?"#FFFFFF":t.sidebarMuted} />
            {!collapsed && <span>{n.label}</span>}
          </div>
        ))}
      </nav>
      <div style={{ padding:collapsed?"12px 0":"14px 18px", borderTop:`1px solid ${t.sidebarBorder}` }}>
        {!collapsed && (
          <div style={{ marginBottom:10 }}>
            <div style={{ fontSize:12, color:"#FFFFFF", fontFamily:"'Trebuchet MS',sans-serif", fontWeight:600 }}>{user?.name}</div>
            <div style={{ fontSize:10, color:"rgba(255,255,255,0.55)", fontFamily:"'Trebuchet MS',sans-serif", textTransform:"capitalize" }}>
              {getAllRoles().find(r=>r.id===user?.role)?.name || user?.role}{!isAdmin && ` · ${perms.length} permissions`}
            </div>
          </div>
        )}
        <div style={{ display:"flex", gap:6, justifyContent:collapsed?"center":"flex-start" }}>
          <button title="Toggle theme" onClick={toggle} style={{ padding:"7px 8px", borderRadius:7, border:"1px solid rgba(255,255,255,0.2)", background:"transparent", cursor:"pointer" }}>
            <Icon name={dark?"sun":"moon"} size={14} color="rgba(255,255,255,0.7)" />
          </button>
          <button title="Logout" onClick={onLogout} style={{ padding:"7px 8px", borderRadius:7, border:"1px solid rgba(255,255,255,0.2)", background:"transparent", cursor:"pointer" }}>
            <Icon name="logout" size={14} color="rgba(255,255,255,0.7)" />
          </button>
          <button title="Toggle sidebar" onClick={()=>setCollapsed(c=>!c)} style={{ padding:"7px 8px", borderRadius:7, border:"1px solid rgba(255,255,255,0.2)", background:"transparent", cursor:"pointer" }}>
            <Icon name={collapsed?"menu":"close"} size={14} color="rgba(255,255,255,0.7)" />
          </button>
        </div>
      </div>
    </div>
  );
};

// ─── LOGIN PAGE ───────────────────────────────────────────────────────────────
const LoginPage = ({ onLogin }) => {
  const { dark, toggle } = useTheme();
  const t = T[dark?"dark":"light"];
  const [email, setEmail]       = useState("");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw]     = useState(false);
  const [err, setErr]           = useState("");
  const [loading, setLoading]   = useState(false);

  const handleLogin = () => {
    if (!email || !password) { setErr("Please fill in all fields."); return; }
    setLoading(true); setErr("");
    setTimeout(() => {
      const users = loadUsersLocal();
      const emailLower = email.toLowerCase().trim();
      const pwTrim     = password.trim();
      const found = users.find(u =>
        u.email.toLowerCase().trim() === emailLower &&
        u.password.trim() === pwTrim &&
        u.is_active === "YES"
      );
      setLoading(false);
      if (!found) {
        setErr("Invalid email or password, or account is inactive.");
        return;
      }
      const perms = parsePerms(found.permissions);
      onLogin({ id:found.id, name:found.name, email:found.email, role:found.role, permissions:perms });
    }, 300);
  };

  const inp = {
    background: dark?"rgba(255,255,255,0.05)":t.surfaceAlt,
    border:`1px solid ${t.border}`, borderRadius:10, padding:"12px 16px",
    color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:14,
    width:"100%", boxSizing:"border-box", outline:"none"
  };

  return (
    <div style={{ minHeight:"100vh", background:t.bg, display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'Trebuchet MS',sans-serif", transition:"background 0.3s" }}>
      <button onClick={toggle} style={{ position:"fixed", top:20, right:20, background:t.surface, border:`1px solid ${t.border}`, borderRadius:9, padding:"8px 14px", cursor:"pointer", color:t.text, display:"flex", alignItems:"center", gap:7, fontSize:13 }}>
        <Icon name={dark?"sun":"moon"} size={15} color={t.gold} />{dark?"Light":"Dark"}
      </button>

      <div style={{ width:420, padding:44, background:t.surface, borderRadius:22, border:`1px solid ${t.border}`, boxShadow:`0 24px 80px rgba(0,0,0,0.15)` }}>
        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:36 }}>
          <div style={{ width:80, height:80, borderRadius:"50%", background:`linear-gradient(135deg,#004b23,#006e34)`,
            display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 18px",
            boxShadow:`0 8px 28px #004b2355`, overflow:"hidden", border:"3px solid rgba(255,255,255,0.5)" }}>
            {ACTIVE_LOGO
              ? <img src={ACTIVE_LOGO} alt={CHURCH_SHORT} style={{ width:"100%", height:"100%", objectFit:"cover" }} />
              : <Icon name="cross" size={32} color="#FFFFFF" />}
          </div>
          <div style={{ fontSize:20, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", letterSpacing:0.3 }}>{CHURCH_NAME}</div>
          <div style={{ fontSize:11, color:t.textMuted, marginTop:5, letterSpacing:0.8, textTransform:"uppercase" }}>{CHURCH_SUB}</div>
        </div>

        {/* Fields */}
        <div style={{ display:"flex", flexDirection:"column", gap:14, marginBottom:20 }}>
          <div>
            <div style={{ fontSize:10, color:t.textMuted, textTransform:"uppercase", letterSpacing:1.2, marginBottom:7 }}>Email Address</div>
            <input style={inp} type="email" value={email}
              onChange={e=>setEmail(e.target.value)}
              placeholder="your@email.com"
              onKeyDown={e=>e.key==="Enter"&&handleLogin()} />
          </div>
          <div>
            <div style={{ fontSize:10, color:t.textMuted, textTransform:"uppercase", letterSpacing:1.2, marginBottom:7 }}>Password</div>
            <div style={{ position:"relative" }}>
              <input style={{ ...inp, paddingRight:44 }}
                type={showPw ? "text" : "password"}
                value={password}
                onChange={e=>setPassword(e.target.value)}
                placeholder="••••••••"
                onKeyDown={e=>e.key==="Enter"&&handleLogin()} />
              <button type="button" onClick={()=>setShowPw(v=>!v)}
                style={{ position:"absolute", right:12, top:"50%", transform:"translateY(-50%)",
                  background:"transparent", border:"none", cursor:"pointer", padding:4 }}
                title={showPw?"Hide password":"Show password"}>
                <Icon name="eye" size={17} color={t.textMuted} />
              </button>
            </div>
          </div>
        </div>

        {err && (
          <div style={{ color:t.danger, fontSize:12, marginBottom:14, padding:"9px 12px",
            background:t.danger+"12", borderRadius:7, border:`1px solid ${t.danger}33` }}>
            {err}
          </div>
        )}

        <button onClick={handleLogin} disabled={loading}
          style={{ width:"100%", padding:14, borderRadius:10, border:"none", cursor:"pointer",
            background:`linear-gradient(135deg,#004b23,#006e34)`,
            color:"#FFFFFF", fontWeight:700, fontSize:15,
            fontFamily:"'Trebuchet MS',sans-serif",
            boxShadow:`0 4px 16px #004b2333`, opacity:loading?0.7:1 }}>
          {loading ? "Signing in…" : "Sign In"}
        </button>
      </div>
    </div>
  );
};

// ─── BRANDING PAGE (full page wrapper) ───────────────────────────────────────
const BrandingPage = () => {
  const { t, card } = useThemeStyles();
  return (
    <div>
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>
        Logo & Church Name
      </div>
      <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:24 }}>
        Customise how your church name and logo appear across the entire app
      </div>
      <div style={{ maxWidth:640 }}>
        <div style={card}>
          <BrandingSettings onClose={()=>{}} />
        </div>
      </div>
    </div>
  );
};

// ─── MOBILE BOTTOM NAV ────────────────────────────────────────────────────────
const BottomNav = ({ page, setPage, user, onMenuOpen }) => {
  const perms   = user?.permissions || [];
  const isAdmin = user?.role === "admin";
  const quickTabs = (isAdmin ? [
    { id:"dashboard",  icon:"dashboard",  label:"Home" },
    { id:"attendance", icon:"attendance", label:"Records" },
    { id:"church",     icon:"cross",      label:"Church" },
    { id:"analytics",  icon:"analytics",  label:"Stats" },
  ] : [
    { id:"submit",     icon:"submit",     label:"Submit" },
    { id:"church",     icon:"cross",      label:"Church" },
    { id:"attendance", icon:"attendance", label:"Records" },
    { id:"analytics",  icon:"analytics",  label:"Stats" },
  ].filter(n => isAdmin || perms.includes(n.id))).slice(0,4);

  return (
    <div style={{ position:"fixed", bottom:0, left:0, right:0, zIndex:200,
      background:"#004b23", borderTop:"1px solid rgba(255,255,255,0.15)",
      display:"flex", alignItems:"center",
      paddingBottom:"env(safe-area-inset-bottom,0px)",
      boxShadow:"0 -3px 16px rgba(0,75,35,0.35)" }}>
      {quickTabs.map(tab => {
        const active = page === tab.id;
        return (
          <button key={tab.id} onClick={()=>setPage(tab.id)} style={{
            flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
            padding:"10px 2px 8px", border:"none", background:"transparent", cursor:"pointer",
            borderTop: active ? "2px solid #FFFFFF" : "2px solid transparent",
          }}>
            <Icon name={tab.icon} size={21} color={active?"#FFFFFF":"rgba(255,255,255,0.45)"} />
            <span style={{ fontSize:9, color:active?"#FFFFFF":"rgba(255,255,255,0.45)",
              fontFamily:"'Trebuchet MS',sans-serif", marginTop:3, fontWeight:active?700:400 }}>{tab.label}</span>
          </button>
        );
      })}
      <button onClick={onMenuOpen} style={{
        flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
        padding:"10px 2px 8px", border:"none", borderTop:"2px solid transparent",
        background:"transparent", cursor:"pointer" }}>
        <Icon name="menu" size={21} color="rgba(255,255,255,0.5)" />
        <span style={{ fontSize:9, color:"rgba(255,255,255,0.5)", fontFamily:"'Trebuchet MS',sans-serif", marginTop:3 }}>More</span>
      </button>
    </div>
  );
};

// ─── MOBILE DRAWER ────────────────────────────────────────────────────────────
const MobileDrawer = ({ open, onClose, page, setPage, user, onLogout, db }) => {
  const { dark, toggle } = useTheme();
  const t = T[dark?"dark":"light"];
  const perms   = user?.permissions || [];
  const isAdmin = user?.role === "admin";
  const can = (key) => isAdmin || perms.includes(key);

  const allNav = isAdmin ? [
    {id:"dashboard",  label:"Dashboard",      icon:"dashboard"},
    {id:"attendance", label:"SS Records",     icon:"attendance"},
    {id:"church",     label:"Church Attend.", icon:"cross"},
    {id:"analytics",  label:"Analytics",      icon:"analytics"},
    {id:"teachers",   label:"Teachers",       icon:"teachers"},
    {id:"classes",    label:"Classes",        icon:"classes"},
    {id:"programs",   label:"Programs",       icon:"settings"},
    {id:"users",      label:"Users & Access", icon:"users"},
    {id:"roles",      label:"Roles",          icon:"edit"},
    {id:"branding",   label:"Logo & Name",    icon:"bible"},
    {id:"export",     label:"Export",         icon:"export"},
  ] : [
    {id:"submit",     label:"Submit Report",  icon:"submit",    perm:"submit"},
    {id:"church",     label:"Church Attend.", icon:"cross",     perm:"church"},
    {id:"attendance", label:"My Records",     icon:"attendance",perm:"attendance"},
    {id:"analytics",  label:"Analytics",      icon:"analytics", perm:"analytics"},
    {id:"export",     label:"Export",         icon:"export",    perm:"export"},
  ].filter(n => can(n.perm||n.id));

  if (!open) return null;
  return (
    <>
      <div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", zIndex:300 }} />
      <div style={{ position:"fixed", left:0, top:0, bottom:0, width:280, zIndex:400,
        background:"#004b23", display:"flex", flexDirection:"column",
        boxShadow:"6px 0 32px rgba(0,0,0,0.4)" }}>
        {/* Header */}
        <div style={{ padding:"20px 18px 16px", borderBottom:"1px solid rgba(255,255,255,0.12)",
          display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <ChurchLogo size={36} />
            <div>
              <div style={{ fontSize:10, fontWeight:700, color:"#FFFFFF", fontFamily:"'Trebuchet MS',sans-serif", letterSpacing:1.2 }}>{CHURCH_NAME.toUpperCase()}</div>
              <div style={{ fontSize:8.5, color:"rgba(255,255,255,0.5)", fontFamily:"'Trebuchet MS',sans-serif" }}>SSM SYSTEM</div>
            </div>
          </div>
          <button onClick={onClose} style={{ background:"transparent", border:"1px solid rgba(255,255,255,0.2)", borderRadius:7, padding:"6px 8px", cursor:"pointer" }}>
            <Icon name="close" size={16} color="rgba(255,255,255,0.8)" />
          </button>
        </div>
        {/* User */}
        <div style={{ padding:"13px 18px", borderBottom:"1px solid rgba(255,255,255,0.10)", display:"flex", alignItems:"center", gap:10 }}>
          <Avatar name={user?.name} size={36} color="rgba(255,255,255,0.22)" />
          <div>
            <div style={{ fontSize:13, fontWeight:600, color:"#FFFFFF", fontFamily:"'Trebuchet MS',sans-serif" }}>{user?.name}</div>
            <div style={{ fontSize:10, color:"rgba(255,255,255,0.5)", fontFamily:"'Trebuchet MS',sans-serif", textTransform:"capitalize" }}>
              {getAllRoles().find(r=>r.id===user?.role)?.name || user?.role}
            </div>
          </div>
        </div>
        {/* Nav */}
        <nav style={{ flex:1, overflowY:"auto", paddingTop:6 }}>
          {allNav.map(n => {
            const active = page === n.id;
            return (
              <div key={n.id} onClick={()=>{ setPage(n.id); onClose(); }}
                style={{ display:"flex", alignItems:"center", gap:12, padding:"12px 18px", cursor:"pointer",
                  background: active ? "rgba(255,255,255,0.14)" : "transparent",
                  borderLeft: active ? "3px solid #FFFFFF" : "3px solid transparent",
                  color: active ? "#FFFFFF" : "rgba(255,255,255,0.6)",
                  fontFamily:"'Trebuchet MS',sans-serif", fontSize:14 }}>
                <Icon name={n.icon} size={18} color={active?"#FFFFFF":"rgba(255,255,255,0.5)"} />
                {n.label}
              </div>
            );
          })}
        </nav>
        {/* Footer */}
        <div style={{ padding:"13px 18px", borderTop:"1px solid rgba(255,255,255,0.12)", display:"flex", gap:8 }}>
          <button onClick={toggle} style={{ flex:1, padding:"9px 0", borderRadius:8, border:"1px solid rgba(255,255,255,0.2)", background:"transparent", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:6, color:"rgba(255,255,255,0.75)", fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>
            <Icon name={dark?"sun":"moon"} size={14} color="rgba(255,255,255,0.75)" />{dark?"Light":"Dark"}
          </button>
          {can("export") && (
            <button onClick={()=>{ db.downloadWorkbook(); onClose(); }} style={{ flex:1, padding:"9px 0", borderRadius:8, border:"1px solid rgba(255,255,255,0.2)", background:"transparent", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:6, color:"rgba(255,255,255,0.75)", fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>
              <Icon name="export" size={14} color="rgba(255,255,255,0.75)" />Save
            </button>
          )}
          <button onClick={()=>{ onLogout(); onClose(); }} style={{ flex:1, padding:"9px 0", borderRadius:8, border:"1px solid rgba(255,255,255,0.2)", background:"transparent", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:6, color:"rgba(255,255,255,0.75)", fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>
            <Icon name="logout" size={14} color="rgba(255,255,255,0.75)" />Out
          </button>
        </div>
      </div>
    </>
  );
};

// ─── ROOT APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [dark, setDark]             = useState(false);
  const [user, setUser]             = useState(null);
  const [page, setPage]             = useState("dashboard");
  const [collapsed, setCollapsed]   = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [toasts, setToasts]         = useState([]);
  const db     = useSupabaseDB();
  const users  = useUsers();
  const mobile = useIsMobile();

  useEffect(() => { injectGlobalCSS(); }, []);
  useEffect(() => { if (mobile) setCollapsed(true); }, [mobile]);

  const toast = useCallback((msg, type="info") => {
    const id = uid();
    setToasts(ts=>[...ts,{id,msg,type}]);
    setTimeout(()=>setToasts(ts=>ts.filter(x=>x.id!==id)), 5000);
  }, []);

  const handleLogin = (u) => {
    const perms = u.permissions || [];
    const landingOrder = ["dashboard","submit","church","attendance","analytics"];
    const landing = u.role==="admin" ? "dashboard" : (landingOrder.find(p=>perms.includes(p)) || "submit");
    setUser(u); setPage(landing);
    toast(`Welcome, ${u.name}!`, "success");
  };
  const handleLogout = () => { setUser(null); setPage("dashboard"); setDrawerOpen(false); };

  const can = useCallback((key) => {
    if (!user) return false;
    if (user.role === "admin") return true;
    return (user.permissions||[]).includes(key);
  }, [user]);

  const AccessDenied = () => {
    const t = T[dark?"dark":"light"];
    return (
      <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:300, gap:16, padding:20 }}>
        <div style={{ width:64, height:64, borderRadius:"50%", background:t.danger+"14", display:"flex", alignItems:"center", justifyContent:"center" }}>
          <Icon name="close" size={28} color={t.danger} />
        </div>
        <div style={{ fontSize:18, color:t.danger, fontFamily:"'Georgia',serif", textAlign:"center" }}>Access Restricted</div>
        <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textAlign:"center", maxWidth:300 }}>
          You don't have permission to view this page. Contact your administrator.
        </div>
      </div>
    );
  };

  const guard = (permKey, component) => can(permKey) ? component : <AccessDenied />;

  const pages = {
    dashboard: guard("dashboard",  <DashboardPage db={db} />),
    attendance:guard("attendance", <AttendancePage db={db} user={user} />),
    church:    guard("church",     <ChurchAttendancePage db={db} user={user} />),
    analytics: guard("analytics",  <AnalyticsPage db={db} />),
    teachers:  guard("teachers",   <TeachersPage db={db} />),
    classes:   guard("classes",    <ClassesPage db={db} />),
    programs:  guard("programs",   <ProgramsPage db={db} />),
    users:     user?.role==="admin" ? <UsersPage users={users} /> : <AccessDenied />,
    roles:     user?.role==="admin" ? <RolesPage /> : <AccessDenied />,
    branding:  user?.role==="admin" ? <BrandingPage /> : <AccessDenied />,
    export:    guard("export",     <ExportPage db={db} />),
    submit:    guard("submit",     <SubmitPage db={db} user={user} onSuccess={()=>toast("Report saved!","success")} />),
  };

  const t = T[dark?"dark":"light"];

  if (!user) return (
    <ThemeCtx.Provider value={{ dark, toggle:()=>setDark(d=>!d) }}>
      <LoginPage onLogin={handleLogin} />
    </ThemeCtx.Provider>
  );

  // Show loading screen while Supabase fetches data
  if (db.loading) {
    const t = T[dark?"dark":"light"];
    return (
      <ThemeCtx.Provider value={{ dark, toggle:()=>setDark(d=>!d) }}>
        <div style={{ minHeight:"100vh", background:t.bg, display:"flex", flexDirection:"column",
          alignItems:"center", justifyContent:"center", gap:20, fontFamily:"'Trebuchet MS',sans-serif" }}>
          <div style={{ width:56, height:56, borderRadius:"50%", background:"#004b23",
            display:"flex", alignItems:"center", justifyContent:"center" }}>
            <Icon name="cross" size={26} color="#FFFFFF" />
          </div>
          <div style={{ fontSize:16, color:t.gold, fontWeight:700 }}>Loading data…</div>
          <div style={{ fontSize:13, color:t.textMuted }}>Connecting to Supabase database</div>
          <div style={{ width:200, height:4, background:t.border, borderRadius:4, overflow:"hidden" }}>
            <div style={{ height:"100%", background:"#004b23", borderRadius:4,
              animation:"loadbar 1.4s ease-in-out infinite",
              width:"60%" }} />
          </div>
          <style>{`@keyframes loadbar{0%{transform:translateX(-100%)}100%{transform:translateX(240%)}}`}</style>
        </div>
      </ThemeCtx.Provider>
    );
  }

  return (
    <ThemeCtx.Provider value={{ dark, toggle:()=>setDark(d=>!d) }}>
      <div style={{ fontFamily:"'Trebuchet MS',sans-serif", background:t.bg, minHeight:"100vh",
        display:"flex", color:t.text, transition:"background 0.25s", position:"relative" }}>

        <ToastContainer toasts={toasts} dismiss={id=>setToasts(ts=>ts.filter(x=>x.id!==id))} />

        {/* Mobile slide-out drawer */}
        {mobile && (
          <MobileDrawer open={drawerOpen} onClose={()=>setDrawerOpen(false)}
            page={page} setPage={setPage} user={user} onLogout={handleLogout} db={db} />
        )}

        {/* Desktop sidebar */}
        {!mobile && (
          <Sidebar page={page} setPage={setPage} user={user} onLogout={handleLogout}
            collapsed={collapsed} setCollapsed={setCollapsed} />
        )}

        {/* Main area */}
        <div style={{ flex:1, display:"flex", flexDirection:"column",
          overflowX:"hidden", minHeight:"100vh",
          paddingBottom: mobile ? 68 : 0 }}>

          {/* Topbar — green band */}
          <div style={{ background:"#004b23", padding: mobile ? "11px 14px" : "11px 24px",
            display:"flex", alignItems:"center", justifyContent:"space-between",
            flexShrink:0, position:"sticky", top:0, zIndex:100,
            boxShadow:"0 2px 12px rgba(0,75,35,0.3)" }}>

            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              {mobile && (
                <button onClick={()=>setDrawerOpen(true)}
                  style={{ background:"transparent", border:"1px solid rgba(255,255,255,0.25)", borderRadius:7, padding:"6px 8px", cursor:"pointer", marginRight:2 }}>
                  <Icon name="menu" size={18} color="#FFFFFF" />
                </button>
              )}
              <ChurchLogo size={28} border={false} />
              <div>
                <div style={{ fontFamily:"'Georgia',serif", color:"#FFFFFF", fontSize: mobile?12:14, fontStyle:"italic", lineHeight:1.2 }}>
                  {mobile ? CHURCH_SHORT : CHURCH_NAME}
                </div>
                {!mobile && <div style={{ fontSize:10, color:"rgba(255,255,255,0.6)", fontFamily:"'Trebuchet MS',sans-serif" }}>{CHURCH_SUB}</div>}
              </div>
            </div>

            <div style={{ display:"flex", alignItems:"center", gap: mobile?7:10 }}>
              {!mobile && (
                <div style={{ padding:"3px 9px", borderRadius:20, background:"rgba(255,255,255,0.15)",
                  border:"1px solid rgba(255,255,255,0.28)", display:"flex", alignItems:"center", gap:5 }}>
                  <span style={{ fontSize:10, color:"#FFFFFF", fontFamily:"'Trebuchet MS',sans-serif", fontWeight:700 }}>
                    {user.role==="admin" ? "ADMIN" : (getAllRoles().find(r=>r.id===user.role)?.name || user.role).toUpperCase()}
                  </span>
                </div>
              )}
              <button onClick={()=>setDark(d=>!d)}
                style={{ padding:"6px 8px", borderRadius:7, border:"1px solid rgba(255,255,255,0.22)", background:"transparent", cursor:"pointer" }}>
                <Icon name={dark?"sun":"moon"} size={14} color="rgba(255,255,255,0.85)" />
              </button>
              {!mobile && can("export") && (
                <button onClick={db.downloadWorkbook}
                  style={{ padding:"6px 11px", borderRadius:7, border:"1px solid rgba(255,255,255,0.22)", background:"transparent", cursor:"pointer", display:"flex", alignItems:"center", gap:5, color:"rgba(255,255,255,0.85)", fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>
                  <Icon name="export" size={13} color="rgba(255,255,255,0.85)" />Save
                </button>
              )}
              <Avatar name={user.name} size={mobile?28:32} color="rgba(255,255,255,0.22)" />
            </div>
          </div>

          {/* Page content */}
          <div style={{ padding: mobile ? "14px 14px 20px" : 28, flex:1 }}>
            {pages[page] || pages.dashboard}
          </div>
        </div>

        {/* Mobile bottom navigation */}
        {mobile && (
          <BottomNav page={page} setPage={setPage} user={user} onMenuOpen={()=>setDrawerOpen(true)} />
        )}
      </div>
    </ThemeCtx.Provider>
  );
}
