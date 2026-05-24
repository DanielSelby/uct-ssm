import { useState, useEffect, useCallback, createContext, useContext, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area
} from "recharts";

// ─── THEME ────────────────────────────────────────────────────────────────────
// 3-mode cycle: light → navy → dark
const ThemeCtx = createContext({ mode: "light", toggle: () => {} });
const useTheme = () => useContext(ThemeCtx);

// Primary brand: #004b23 (deep forest green) + white
const GREEN      = "#004b23";
const GREEN_MID  = "#005c2b";
const GREEN_LIGHT= "#006e34";
const GREEN_PALE = "#e8f5ee";

// Active / selected highlight across all themes
const ACTIVE_COLOR = "#ff6700";

const T = {
  // ── Light: white content area, green sidebar ─────────────
  light: {
    bg:          "#F2F7F4",
    surface:     "#FFFFFF",
    surfaceAlt:  "#EAF3ED",
    surfaceHover:"#D6EBDE",
    sidebar:     "#004b23",
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.55)",
    sidebarBorder:"rgba(255,255,255,0.12)",
    sidebarActive:"rgba(255,102,0,0.18)",   // orange tint
    border:      "rgba(0,75,35,0.14)",
    borderStrong:"rgba(0,75,35,0.38)",
    text:        "#002A14",
    textMuted:   "#3A6B4E",
    textFaint:   "#A8C8B4",
    gold:        "#004b23",
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
    topbar:      "#004b23",
  },
  // ── Navy: midnight blue sidebar, cream/sage page bg ───────
  navy: {
    bg:          "#eef4ed",
    surface:     "#FFFFFF",
    surfaceAlt:  "#e4ede3",
    surfaceHover:"#d5e8d4",
    sidebar:     "#0b2545",
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.50)",
    sidebarBorder:"rgba(255,255,255,0.10)",
    sidebarActive:"rgba(255,102,0,0.20)",
    border:      "rgba(11,37,69,0.14)",
    borderStrong:"rgba(11,37,69,0.38)",
    text:        "#0b2545",
    textMuted:   "#3a5a80",
    textFaint:   "#a8c0d8",
    gold:        "#0b2545",
    goldLight:   "#1a4080",
    goldDark:    "#07172e",
    success:     "#0A7A45",
    danger:      "#C0392B",
    info:        "#1565C0",
    warn:        "#C87A0A",
    navy:        "#0b2545",
    btnFrom:     "#0b2545",
    btnTo:       "#1a4080",
    btnText:     "#FFFFFF",
    topbar:      "#0b2545",
  },
  // ── Dark: deep green + white text ─────────────────────────
  dark: {
    bg:          "#011a0d",
    surface:     "#012a14",
    surfaceAlt:  "#013d1e",
    surfaceHover:"#01531f",
    sidebar:     "#004b23",
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.50)",
    sidebarBorder:"rgba(255,255,255,0.10)",
    sidebarActive:"rgba(255,102,0,0.18)",
    border:      "rgba(0,200,80,0.15)",
    borderStrong:"rgba(0,200,80,0.38)",
    text:        "#E8F5EE",
    textMuted:   "#7AB898",
    textFaint:   "#1A4028",
    gold:        "#4CD98A",
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
    topbar:      "#004b23",
  },
  // ── Dark Blue: deep navy bg + blue sidebar ─────────────────
  darkBlue: {
    bg:          "#060d1a",
    surface:     "#0d1b2e",
    surfaceAlt:  "#112240",
    surfaceHover:"#1a3560",
    sidebar:     "#0b2545",
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.50)",
    sidebarBorder:"rgba(255,255,255,0.10)",
    sidebarActive:"rgba(255,102,0,0.20)",
    border:      "rgba(100,160,255,0.15)",
    borderStrong:"rgba(100,160,255,0.35)",
    text:        "#ccd6f6",
    textMuted:   "#7a9ac9",
    textFaint:   "#1a3560",
    gold:        "#64b5f6",
    goldLight:   "#90caf9",
    goldDark:    "#1565c0",
    success:     "#4dd0a1",
    danger:      "#ef5350",
    info:        "#82b1ff",
    warn:        "#ffb74d",
    navy:        "#060d1a",
    btnFrom:     "#0b2545",
    btnTo:       "#1a4080",
    btnText:     "#FFFFFF",
    topbar:      "#0b2545",
  },
  // ── Orange: warm amber sidebar, light cream page ───────────
  orange: {
    bg:          "#fff8f0",
    surface:     "#FFFFFF",
    surfaceAlt:  "#fff0e0",
    surfaceHover:"#ffe4c4",
    sidebar:     "#cc5803",
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.55)",
    sidebarBorder:"rgba(255,255,255,0.14)",
    sidebarActive:"rgba(255,255,255,0.18)",
    border:      "rgba(204,88,3,0.16)",
    borderStrong:"rgba(204,88,3,0.40)",
    text:        "#4a1a00",
    textMuted:   "#8a4010",
    textFaint:   "#e8b890",
    gold:        "#cc5803",
    goldLight:   "#e87020",
    goldDark:    "#8a3a00",
    success:     "#2e7d32",
    danger:      "#c0392b",
    info:        "#1565c0",
    warn:        "#cc5803",
    navy:        "#4a1a00",
    btnFrom:     "#cc5803",
    btnTo:       "#e87020",
    btnText:     "#FFFFFF",
    topbar:      "#cc5803",
  },
  // ── Dark Orange: deep brown-black bg + orange sidebar ──────
  darkOrange: {
    bg:          "#1a0a00",
    surface:     "#2a1200",
    surfaceAlt:  "#3d1c00",
    surfaceHover:"#552700",
    sidebar:     "#cc5803",
    sidebarText: "#FFFFFF",
    sidebarMuted:"rgba(255,255,255,0.50)",
    sidebarBorder:"rgba(255,255,255,0.10)",
    sidebarActive:"rgba(255,255,255,0.16)",
    border:      "rgba(255,140,60,0.18)",
    borderStrong:"rgba(255,140,60,0.40)",
    text:        "#ffe8cc",
    textMuted:   "#d4956a",
    textFaint:   "#5a2e00",
    gold:        "#ffaa55",
    goldLight:   "#ffc880",
    goldDark:    "#cc5803",
    success:     "#66bb6a",
    danger:      "#ef5350",
    info:        "#64b5f6",
    warn:        "#ffa726",
    navy:        "#1a0a00",
    btnFrom:     "#cc5803",
    btnTo:       "#e87020",
    btnText:     "#FFFFFF",
    topbar:      "#cc5803",
  },
};

// Convenience: is the mode "dark" for dark-specific branches
const isDark = (mode) => mode === "dark" || mode === "darkBlue" || mode === "darkOrange";

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
  { key:"ssreport",   label:"SS Report Viewer",     group:"Navigation" },
  { key:"lessons",    label:"Lesson Register",      group:"Navigation" },
  { key:"ai",         label:"AI Assistant",         group:"Navigation" },
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
  // Dashboard: 4-square grid (overview/home feel)
  dashboard:"M4 5a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1H5a1 1 0 01-1-1V5zm10 0a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1h-4a1 1 0 01-1-1V5zM4 15a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1H5a1 1 0 01-1-1v-4zm10 0a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1h-4a1 1 0 01-1-1v-4z",
  attendance:"M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4",
  // Analytics: rising trend line with a circle dot at peak
  analytics:"M3 17l4-8 4 4 4-6 4 4M22 12a1 1 0 11-2 0 1 1 0 012 0z",
  // SSReport: table/grid report icon
  ssreport:"M3 10h18M3 14h18M10 3v18M6 3h12a1 1 0 011 1v16a1 1 0 01-1 1H6a1 1 0 01-1-1V4a1 1 0 011-1z",
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
  book:"M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253",
  lesson:"M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01",
  upload:"M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12",
  pause:"M10 9v6m4-6v6m7-3a9 9 0 11-18 0 9 9 0 0118 0z",
  play:"M14.752 11.168l-3.197-2.132A1 1 0 0010 9.87v4.263a1 1 0 001.555.832l3.197-2.132a1 1 0 000-1.664z M21 12a9 9 0 11-18 0 9 9 0 0118 0z",
  info:"M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z",
  ai:"M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z",
  sparkle:"M5 3l1.5 4.5L11 9l-4.5 1.5L5 15l-1.5-4.5L-1 9l4.5-1.5L5 3zm12 8l1 3 3 1-3 1-1 3-1-3-3-1 3-1 1-3z",
  // Church / cross icon kept for nav
  settings:"M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z M15 12a3 3 0 11-6 0 3 3 0 016 0z",
  users:"M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197M13 7a4 4 0 11-8 0 4 4 0 018 0z",
  // Trend / chart with up-arrow for KPI cards
  trend:"M13 7h8m0 0v8m0-8l-8 8-4-4-6 6",
  // Bell for notifications
  bell:"M15 17h5l-1.405-1.405A2.032 2.032 0 0118 14.158V11a6.002 6.002 0 00-4-5.659V5a2 2 0 10-4 0v.341C7.67 6.165 6 8.388 6 11v3.159c0 .538-.214 1.055-.595 1.436L4 17h5m6 0v1a3 3 0 11-6 0v-1m6 0H9",
};

const Icon = ({ name, size=18, color="currentColor" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
    {(PATHS[name]||"").split("M").filter(Boolean).map((d,i)=><path key={i} d={"M"+d}/>)}
  </svg>
);

// ─── ANIMATED BIBLE ───────────────────────────────────────────────────────────
// Rendered inline in the sidebar header / loading screen
const AnimatedBible = ({ size = 48 }) => {
  const styleId = "uct-bible-anim";
  useEffect(() => {
    if (document.getElementById(styleId)) return;
    const s = document.createElement("style");
    s.id = styleId;
    s.textContent = `
      @keyframes biblePageLeft  { 0%,100%{transform:rotateY(0deg)}   50%{transform:rotateY(-28deg)} }
      @keyframes biblePageRight { 0%,100%{transform:rotateY(0deg)}   50%{transform:rotateY(28deg)} }
      @keyframes bibleGlow      { 0%,100%{opacity:0.55} 50%{opacity:1} }
      @keyframes bibleFloat     { 0%,100%{transform:translateY(0px)} 50%{transform:translateY(-4px)} }
      @keyframes bibleSpine     { 0%,100%{opacity:0.7} 50%{opacity:1} }
      .bible-wrap   { animation: bibleFloat 2.8s ease-in-out infinite; display:inline-block; }
      .bible-left   { transform-origin: right center; animation: biblePageLeft  2.8s ease-in-out infinite; }
      .bible-right  { transform-origin: left center;  animation: biblePageRight 2.8s ease-in-out infinite; }
      .bible-glow   { animation: bibleGlow 2.8s ease-in-out infinite; }
      .bible-spine  { animation: bibleSpine 2.8s ease-in-out infinite; }
    `;
    document.head.appendChild(s);
  }, []);

  const w = size, h = size * 0.75;
  const bookW = w * 0.44, bookH = h * 0.82;
  const spineW = w * 0.07;
  const cx = w / 2, cy = h / 2;

  return (
    <div className="bible-wrap" style={{ width:w, height:h, position:"relative", display:"inline-flex", alignItems:"center", justifyContent:"center" }}>
      <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ overflow:"visible" }}>
        <defs>
          <linearGradient id="bgl" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0%" stopColor="#d4a843" />
            <stop offset="100%" stopColor="#a8832a" />
          </linearGradient>
          <linearGradient id="bgr" x1="1" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#c49b38" />
            <stop offset="100%" stopColor="#8c6b1e" />
          </linearGradient>
          <linearGradient id="bgs" x1="0" y1="0" x2="1" y2="0">
            <stop offset="0%" stopColor="#7a5a14" />
            <stop offset="100%" stopColor="#b8942e" />
          </linearGradient>
          <filter id="bglow">
            <feGaussianBlur stdDeviation="2" result="blur"/>
            <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
          </filter>
        </defs>

        {/* Glow beneath */}
        <ellipse className="bible-glow" cx={cx} cy={h*0.93} rx={bookW*0.7} ry={h*0.06}
          fill="rgba(212,168,67,0.35)" />

        {/* LEFT page */}
        <g className="bible-left">
          <rect x={cx - spineW/2 - bookW} y={cy - bookH/2}
            width={bookW} height={bookH} rx={size*0.025}
            fill="url(#bgl)" stroke="#8c6b1e" strokeWidth={size*0.018} />
          {/* Page lines */}
          {[0.28,0.42,0.56,0.70].map((f,i) => (
            <line key={i}
              x1={cx - spineW/2 - bookW + bookW*0.12} y1={cy - bookH/2 + bookH*f}
              x2={cx - spineW/2 - bookW*0.12}         y2={cy - bookH/2 + bookH*f}
              stroke="rgba(255,255,255,0.35)" strokeWidth={size*0.018} strokeLinecap="round"/>
          ))}
          {/* Cross symbol on left */}
          <line x1={cx - spineW/2 - bookW*0.56} y1={cy - bookH*0.22}
                x2={cx - spineW/2 - bookW*0.56} y2={cy + bookH*0.14}
                stroke="rgba(255,255,255,0.55)" strokeWidth={size*0.033} strokeLinecap="round"/>
          <line x1={cx - spineW/2 - bookW*0.72} y1={cy - bookH*0.08}
                x2={cx - spineW/2 - bookW*0.40} y2={cy - bookH*0.08}
                stroke="rgba(255,255,255,0.55)" strokeWidth={size*0.033} strokeLinecap="round"/>
        </g>

        {/* RIGHT page */}
        <g className="bible-right">
          <rect x={cx + spineW/2} y={cy - bookH/2}
            width={bookW} height={bookH} rx={size*0.025}
            fill="url(#bgr)" stroke="#8c6b1e" strokeWidth={size*0.018} />
          {[0.28,0.42,0.56,0.70].map((f,i) => (
            <line key={i}
              x1={cx + spineW/2 + bookW*0.12} y1={cy - bookH/2 + bookH*f}
              x2={cx + spineW/2 + bookW*0.88} y2={cy - bookH/2 + bookH*f}
              stroke="rgba(255,255,255,0.28)" strokeWidth={size*0.018} strokeLinecap="round"/>
          ))}
        </g>

        {/* Spine */}
        <rect className="bible-spine"
          x={cx - spineW/2} y={cy - bookH/2}
          width={spineW} height={bookH}
          fill="url(#bgs)" />
      </svg>
    </div>
  );
};

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
  const { mode } = useTheme();
  const t = T[mode] || T.light;
  const dark = isDark(mode);
  const btnGrad = `linear-gradient(135deg,${t.btnFrom},${t.btnTo})`;
  return {
    t, dark,
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
      const { mode, dark } = useTheme();
      const t = T[mode]||T.light;
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
  const { mode, dark } = useTheme();
  const t = T[mode]||T.light;
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
      padding:"18px 20px", borderTop:`3px solid ${c}` }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
        <div>
          <div style={{ fontSize:14, fontWeight:800, color:t.text,
            fontFamily:"'Trebuchet MS',sans-serif", marginBottom:8, lineHeight:1.3 }}>{label}</div>
          <div style={{ fontSize:30, fontWeight:900, color:c, fontFamily:"'Georgia',serif", lineHeight:1 }}>{value}</div>
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
  const { records, teachers, classes, churchRecs, programs, loadAll } = db;
  const active = teachers.filter(x=>x.is_active==="YES").length;

  const [filters, setFilters] = useState({});
  const [lastRefresh, setLastRefresh] = useState(new Date());
  const [refreshing, setRefreshing] = useState(false);

  // Auto-refresh every 30 seconds to catch new submissions without logout
  useEffect(() => {
    const interval = setInterval(() => {
      loadAll().then ? loadAll().then(()=>setLastRefresh(new Date())) : (loadAll(), setLastRefresh(new Date()));
    }, 30000);
    return () => clearInterval(interval);
  }, [loadAll]);

  const handleManualRefresh = async () => {
    setRefreshing(true);
    await loadAll();
    setLastRefresh(new Date());
    setRefreshing(false);
  };

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

  const fRec    = applyRange(records);
  const fChurch = applyRange(churchRecs);

  // KPIs — both beginning and closing
  const ssBeginTotal = fRec.reduce((s,r)=>s+(Number(r.total_beginning)||0),0);
  const ssTotal      = fRec.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const chBeginTotal = fChurch.reduce((s,r)=>s+(Number(r.total_beginning)||0),0);
  const chTotal      = fChurch.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const biblesBegin  = fRec.reduce((s,r)=>s+(Number(r.bibles_beginning)||0),0);
  const bibles       = fRec.reduce((s,r)=>s+(Number(r.bibles_closing)||0),0);
  const bibleRate    = ssTotal > 0 ? Math.round(bibles/ssTotal*100) : 0;
  const totalFT      = fChurch.reduce((s,r)=>s+(Number(r.first_timers)||0),0) + fRec.reduce((s,r)=>s+(Number(r.first_timers)||0),0);
  const retentionRate = chTotal > 0 ? Math.round(ssTotal/chTotal*100) : 0;
  const pending       = records.filter(r=>r.status==="pending").length;

  // SS vs Church comparison by date
  const compData = buildComparison(fRec, fChurch, {});

  // Weekly trend
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

  const retentionStory = retentionRate > 0
    ? retentionRate >= 80 ? `✝ Strong: ${retentionRate}% of Sunday service attendees are in Sunday School.`
    : retentionRate >= 50 ? `📈 Moderate: ${retentionRate}% of Sunday service attendees attend Sunday School.`
    : `⚠ Only ${retentionRate}% of Sunday attendees join Sunday School. Consider outreach.`
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
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:16, flexWrap:"wrap", gap:12 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>Overview Dashboard</div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            {records.length} SS records · {churchRecs.length} church services
            {(filters.from||filters.to||filters.month) ? " · Filtered" : " · All time"}
          </div>
        </div>
        {/* Refresh button */}
        <button onClick={handleManualRefresh} disabled={refreshing}
          style={{ display:"flex", alignItems:"center", gap:7, padding:"9px 18px", borderRadius:9,
            border:`1px solid ${t.gold}44`, background:t.gold+"11", color:t.gold,
            fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, cursor:"pointer" }}>
          <span style={{ fontSize:16, lineHeight:1, display:"inline-block",
            animation: refreshing ? "spin 0.7s linear infinite" : "none" }}>↻</span>
          {refreshing ? "Refreshing…" : "Refresh"}
          <span style={{ fontSize:10, color:t.textMuted }}>
            · {lastRefresh.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}
          </span>
        </button>
        <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
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

      {/* KPIs — Row 1: SS */}
      <div style={{ fontSize:11, fontWeight:700, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:10 }}>
        Sunday School
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:12, marginBottom:16 }}>
        <KpiCard label="SS Begin (Total)"   value={ssBeginTotal} sub="At opening"        icon="attendance" color={t.info} />
        <KpiCard label="SS Closing (Total)" value={ssTotal}      sub="At close"          icon="attendance" color={t.gold} />
        <KpiCard label="Bibles at Begin"    value={biblesBegin}  sub="Brought at start"  icon="bible"      color="#9B59B6" />
        <KpiCard label="Bibles at Closing"  value={bibles}       sub={`${bibleRate}% rate`} icon="bible"   color="#7B3FBE" />
        <KpiCard label="First Timers"       value={fRec.reduce((s,r)=>s+(Number(r.first_timers)||0),0)} sub="SS only" icon="plus" color="#E67E22" />
        <KpiCard label="Pending Reviews"    value={pending}      sub="Needs approval"    icon="info"       color={pending>0?t.warn:t.textMuted} />
      </div>

      {/* KPIs — Row 2: Church */}
      <div style={{ fontSize:11, fontWeight:700, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:10 }}>
        Church Service
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:12, marginBottom:24 }}>
        <KpiCard label="Church Begin"       value={chBeginTotal} sub="At opening"         icon="cross"      color={t.info} />
        <KpiCard label="Church Closing"     value={chTotal}      sub="At close"           icon="cross"      color="#1A5DC8" />
        <KpiCard label="SS Retention"       value={`${retentionRate}%`} sub="SS÷Church"  icon="analytics"  color={retentionRate>=80?t.success:retentionRate>=50?t.warn:t.danger} />
        <KpiCard label="First Timers"       value={fChurch.reduce((s,r)=>s+(Number(r.first_timers)||0),0)} sub="Church only" icon="plus" color="#E67E22" />
        <KpiCard label="Active Teachers"    value={active}       sub={`of ${teachers.length}`} icon="teachers" color={t.success} />
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
const SubmitPage = ({ db, user, onSuccess, editRecord: editProp, onCancelEdit }) => {
  const { t, inp, sel, lbl, btnGold, btnOutline, card } = useThemeStyles();
  const { teachers, classes, addRecord, updateRecord } = db;
  const activeTeachers = teachers.filter(x => x.is_active === "YES");

  const isEditMode = !!editProp;

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

  const [form, setForm]       = useState(isEditMode ? { ...makeBlank(), ...editProp } : makeBlank);
  const [loading, setLoading] = useState(false);
  const [done, setDone]       = useState(false);

  // Sync form when editProp changes (e.g. user selects a different record)
  useEffect(() => {
    if (editProp) setForm({ ...makeBlank(), ...editProp });
    else          setForm(makeBlank());
  }, [editProp?.id]);

  const handleChange = useCallback((e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  }, []);

  const handleSubmit = () => {
    if (!form.class_name)   { alert("Please select a class."); return; }
    if (!form.topic.trim()) { alert("Please enter the lesson topic."); return; }
    setLoading(true);
    setTimeout(() => {
      if (isEditMode) {
        // Update existing — preserve id and status
        updateRecord(editProp.id, { ...form, status: "pending" });
      } else {
        addRecord(form);
      }
      setLoading(false);
      setDone(true);
      setTimeout(() => {
        setDone(false);
        setForm(makeBlank());
        onSuccess && onSuccess();
        onCancelEdit && onCancelEdit();
      }, 2200);
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
      <div style={{ fontSize:24, color:t.gold, fontFamily:"'Georgia',serif" }}>
        {isEditMode ? "Record Updated!" : "Report Submitted!"}
      </div>
      <div style={{ color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", fontSize:14 }}>
        {isEditMode ? "Changes saved. Awaiting admin approval." : "Saved to your records."}
      </div>
    </div>
  );

  return (
    <div>
      {/* Page title + edit banner */}
      <div style={{ marginBottom:20 }}>
        <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>
          {isEditMode ? "Edit Report" : "Submit Attendance Report"}
        </div>
        <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
          {isEditMode ? "You are editing a pending submission. Changes will reset it to pending for re-approval." : "Saved directly to your records."}
        </div>
      </div>

      {/* Edit mode banner */}
      {isEditMode && (
        <div style={{ marginBottom:18, padding:"12px 16px", borderRadius:10,
          background: t.warn+"18", border:`1px solid ${t.warn}44`,
          display:"flex", alignItems:"center", justifyContent:"space-between", gap:12, flexWrap:"wrap" }}>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:18 }}>✏️</span>
            <div>
              <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>
                Editing: {editProp.class_name} — {editProp.date}
              </div>
              <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
                Status will return to <strong>Pending</strong> after saving, requiring admin re-approval.
              </div>
            </div>
          </div>
          <button onClick={onCancelEdit} style={{ padding:"7px 14px", borderRadius:8,
            border:`1px solid ${t.border}`, background:"transparent", color:t.textMuted,
            fontFamily:"'Trebuchet MS',sans-serif", fontSize:12, cursor:"pointer" }}>
            ✕ Cancel Edit
          </button>
        </div>
      )}

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

        <div style={{ marginTop:24, display:"flex", gap:12, flexWrap:"wrap" }}>
          <button onClick={handleSubmit} disabled={loading} style={{ ...btnGold, padding:"13px 36px", fontSize:14 }}>
            {loading ? "Saving…" : isEditMode ? "💾 Save Changes" : "Submit Report"}
          </button>
          {isEditMode
            ? <button onClick={onCancelEdit} style={{ ...btnOutline, padding:"13px 24px" }}>Cancel</button>
            : <button onClick={()=>setForm(makeBlank())} style={{ ...btnOutline, padding:"13px 24px" }}>Clear Form</button>
          }
        </div>
      </div>
    </div>
  );
};


// ─── ATTENDANCE PAGE ──────────────────────────────────────────────────────────
const AttendancePage = ({ db, user, onEditRecord }) => {
  const { t, card, btnGhost, th, td } = useThemeStyles();
  const { records, classes, approveRecord, deleteRecord } = db;
  const [detail, setDetail] = useState(null);
  const [filter, setFilter] = useState({
    cls: "", status: "", search: "", teacher: "",
    dateFrom: "", dateTo: "", month: "",
  });
  const [sortDir, setSortDir] = useState("desc"); // "desc" newest first, "asc" oldest first

  const sel = { background:t.surfaceAlt, border:`1px solid ${t.border}`, borderRadius:9, padding:"8px 12px", color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, outline:"none" };
  const inp = { ...sel };

  // Derived filter options
  const allDates    = [...new Set(records.map(r=>r.date).filter(Boolean))].sort().reverse();
  const allMonths   = [...new Set(records.map(r=>(r.date||"").slice(0,7)).filter(Boolean))].sort().reverse();
  const allTeachers = [...new Set(records.map(r=>r.teacher_name).filter(Boolean))].sort();

  const filtered = records.filter(r => {
    if (filter.cls     && r.class_name   !== filter.cls)     return false;
    if (filter.status  && r.status       !== filter.status)  return false;
    if (filter.teacher && r.teacher_name !== filter.teacher) return false;
    if (filter.month   && !(r.date||"").startsWith(filter.month)) return false;
    if (filter.dateFrom && r.date < filter.dateFrom)         return false;
    if (filter.dateTo   && r.date > filter.dateTo)           return false;
    if (filter.search) {
      const q = filter.search.toLowerCase();
      if (!r.topic?.toLowerCase().includes(q) &&
          !r.teacher_name?.toLowerCase().includes(q) &&
          !r.class_name?.toLowerCase().includes(q))          return false;
    }
    return true;
  });

  const sorted = [...filtered].sort((a,b) => sortDir==="asc"
    ? (a.date||"").localeCompare(b.date||"")
    : (b.date||"").localeCompare(a.date||""));

  const hasActiveFilter = Object.values(filter).some(v=>v!=="");

  // Summary stats for the filtered set
  const stats = sorted.reduce((acc, r) => ({
    begin:   acc.begin   + (Number(r.total_beginning)||0),
    close:   acc.close   + (Number(r.total_closing)||0),
    bibles:  acc.bibles  + (Number(r.bibles_closing)||0),
    firstT:  acc.firstT  + (Number(r.first_timers)||0),
    visitors:acc.visitors+ (Number(r.visitors)||0),
  }), { begin:0, close:0, bibles:0, firstT:0, visitors:0 });

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
      {/* Page title */}
      <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>SS Records</div>
      <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:18 }}>
        {records.length} total records · {sorted.length} shown
      </div>

      {/* ── Filter bar ─────────────────────────────────────────────────────── */}
      <div style={{ ...card, marginBottom:14, display:"flex", flexDirection:"column", gap:12 }}>

        {/* Row 1: search + class + teacher + status */}
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
          <input style={{ ...inp, minWidth:200, flex:1 }} placeholder="🔍  Search topic, teacher or class…"
            value={filter.search} onChange={e=>setFilter(f=>({...f,search:e.target.value}))} />

          <select style={{ ...sel, minWidth:150 }} value={filter.cls} onChange={e=>setFilter(f=>({...f,cls:e.target.value}))}>
            <option value="">All Classes</option>
            {[...new Set(records.map(r=>r.class_name).filter(Boolean))].sort().map(c=><option key={c}>{c}</option>)}
          </select>

          {user?.role==="admin" && (
            <select style={{ ...sel, minWidth:150 }} value={filter.teacher} onChange={e=>setFilter(f=>({...f,teacher:e.target.value}))}>
              <option value="">All Teachers</option>
              {allTeachers.map(t2=><option key={t2}>{t2}</option>)}
            </select>
          )}

          <select style={{ ...sel }} value={filter.status} onChange={e=>setFilter(f=>({...f,status:e.target.value}))}>
            <option value="">All Status</option>
            <option value="approved">✅ Approved</option>
            <option value="pending">🕐 Pending</option>
          </select>
        </div>

        {/* Row 2: date filters */}
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
          {/* Quick month picker */}
          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
            <span style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", whiteSpace:"nowrap" }}>Month</span>
            <select style={{ ...sel }} value={filter.month} onChange={e=>setFilter(f=>({...f,month:e.target.value,dateFrom:"",dateTo:""}))}>
              <option value="">All</option>
              {allMonths.map(m=>{
                const [yr,mo]=m.split("-");
                const label=new Date(yr,mo-1).toLocaleString("default",{month:"short",year:"numeric"});
                return <option key={m} value={m}>{label}</option>;
              })}
            </select>
          </div>

          {/* Date range */}
          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
            <span style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", whiteSpace:"nowrap" }}>From</span>
            <input type="date" style={{ ...inp, width:148 }} value={filter.dateFrom}
              onChange={e=>setFilter(f=>({...f,dateFrom:e.target.value,month:""}))} />
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
            <span style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", whiteSpace:"nowrap" }}>To</span>
            <input type="date" style={{ ...inp, width:148 }} value={filter.dateTo}
              onChange={e=>setFilter(f=>({...f,dateTo:e.target.value,month:""}))} />
          </div>

          {/* Sort direction */}
          <button onClick={()=>setSortDir(d=>d==="desc"?"asc":"desc")}
            style={{ ...sel, cursor:"pointer", border:`1px solid ${t.border}`, display:"flex", alignItems:"center", gap:5, whiteSpace:"nowrap" }}>
            {sortDir==="desc" ? "↓ Newest first" : "↑ Oldest first"}
          </button>

          {/* Action buttons */}
          <button onClick={()=>db.loadAll()} style={{ padding:"8px 14px", borderRadius:9, border:`1px solid ${t.gold}44`, background:t.gold+"11", color:t.gold, fontFamily:"'Trebuchet MS',sans-serif", cursor:"pointer", fontSize:13, display:"flex", alignItems:"center", gap:5 }}>
            ↻ Refresh
          </button>

          {hasActiveFilter && (
            <button onClick={()=>setFilter({cls:"",status:"",search:"",teacher:"",dateFrom:"",dateTo:"",month:""})}
              style={{ padding:"8px 14px", borderRadius:9, border:`1px solid ${t.danger}44`, background:t.danger+"11", color:t.danger, fontFamily:"'Trebuchet MS',sans-serif", cursor:"pointer", fontSize:13 }}>
              ✕ Clear filters
            </button>
          )}
        </div>
      </div>

      {/* ── Summary stats strip ────────────────────────────────────────────── */}
      {sorted.length > 0 && (
        <div style={{ display:"flex", gap:10, marginBottom:14, flexWrap:"wrap" }}>
          {[
            { label:"Records",   value:sorted.length,  color:t.gold },
            { label:"Avg Begin", value:sorted.length ? Math.round(stats.begin/sorted.length) : 0, color:t.info },
            { label:"Avg Close", value:sorted.length ? Math.round(stats.close/sorted.length) : 0, color:t.success },
            { label:"Total Bibles", value:stats.bibles, color:"#9B59B6" },
            { label:"First Timers", value:stats.firstT, color:"#E67E22" },
            { label:"Visitors",     value:stats.visitors,color:t.info },
          ].map(s => (
            <div key={s.label} style={{ ...card, padding:"10px 16px", display:"flex", flexDirection:"column", gap:2, minWidth:100, flex:1 }}>
              <div style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1 }}>{s.label}</div>
              <div style={{ fontSize:20, fontWeight:800, color:s.color, fontFamily:"'Georgia',serif" }}>{s.value}</div>
            </div>
          ))}
        </div>
      )}

      {/* ── Table ──────────────────────────────────────────────────────────── */}
      <div style={{ ...card, padding:0, overflowX:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:900 }}>
          <thead><tr style={{ background:t.surfaceAlt }}>
            {["Date","Class","Teacher","Begin","Closing","Bible Begin","Bible Close","First T.","Topic","Status","Actions"].map(h=><th key={h} style={th}>{h}</th>)}
          </tr></thead>
          <tbody>
            {sorted.map(r=>{
              const si = statusInfo(r.status);
              return (
                <tr key={r.id} onMouseEnter={e=>e.currentTarget.style.background=t.surfaceHover} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={td}>{fmtDate(r.date)}</td>
                  <td style={td}>{r.class_name}</td>
                  <td style={{ ...td, color:t.textMuted }}>{r.teacher_name}</td>
                  <td style={td}><span style={{ fontWeight:600, color:t.info }}>{r.total_beginning||0}</span></td>
                  <td style={td}><span style={{ fontWeight:600, color:t.gold }}>{r.total_closing||0}</span></td>
                  <td style={td}><span style={{ color:t.info }}>{r.bibles_beginning||0}</span></td>
                  <td style={td}><span style={{ color:t.gold }}>{r.bibles_closing||0}</span></td>
                  <td style={td}>{r.first_timers||0}</td>
                  <td style={{ ...td, maxWidth:130, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", color:t.textMuted }}>{r.topic}</td>
                  <td style={td}><Badge label={si.label} color={si.color}/></td>
                  <td style={td}>
                    <div style={{ display:"flex", gap:5 }}>
                      <button style={btnGhost} title="View" onClick={()=>setDetail(r)}><Icon name="eye" size={13} color={t.textMuted}/></button>
                      {/* Edit — only for pending records; teacher can edit own, admin can edit any */}
                      {r.status !== "approved" && (user?.role==="admin" || r.teacher_name===user?.name || r.submitted_by===user?.name) && (
                        <button style={btnGhost} title="Edit" onClick={()=>onEditRecord && onEditRecord(r)}>
                          <Icon name="edit" size={13} color={t.warn}/>
                        </button>
                      )}
                      {user?.role==="admin" && r.status!=="approved" && <button style={btnGhost} title="Approve" onClick={()=>approveRecord(r.id)}><Icon name="check" size={13} color={t.success}/></button>}
                      {user?.role==="admin" && <button style={btnGhost} title="Delete" onClick={()=>{ if(window.confirm("Delete record?")) deleteRecord(r.id); }}><Icon name="trash" size={13} color={t.danger}/></button>}
                    </div>
                  </td>
                </tr>
              );
            })}
            {!sorted.length && <tr><td colSpan={11} style={{ ...td, textAlign:"center", padding:48, color:t.textMuted }}>
              {hasActiveFilter ? "No records match the current filters." : "No records yet."}
            </td></tr>}
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
  const ssBeginTotal = fRec.reduce((s,r)=>s+(Number(r.total_beginning)||0),0);
  const ssTotal      = fRec.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const chBeginTotal = fChurch.reduce((s,r)=>s+(Number(r.total_beginning)||0),0);
  const chTotal      = fChurch.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const avgRetention = retentionMonthly.length ? Math.round(retentionMonthly.reduce((s,m)=>s+m.rate,0)/retentionMonthly.length) : 0;
  const bibleBeginTotal = fRec.reduce((s,r)=>s+(Number(r.bibles_beginning)||0),0);
  const bibleTotal   = fRec.reduce((s,r)=>s+(Number(r.bibles_closing)||0),0);
  const bibleRate    = ssTotal > 0 ? Math.round(bibleTotal/ssTotal*100) : 0;
  const ftTotal      = fRec.reduce((s,r)=>s+(Number(r.first_timers)||0),0) + fChurch.reduce((s,r)=>s+(Number(r.first_timers)||0),0);

  // ── Insight generator ─────────────────────────────────────
  const insights = [];
  if (avgRetention > 0) {
    if (avgRetention >= 80) insights.push({ text:`Excellent! An average of ${avgRetention}% of Sunday service attendees participate in Sunday School.`, color:t.success });
    else if (avgRetention >= 60) insights.push({ text:`Good engagement: ${avgRetention}% average SS retention from Sunday service. Aim for 80%+.`, color:t.warn });
    else insights.push({ text:`Only ${avgRetention}% of Sunday attendees join Sunday School on average. This is a key growth area.`, color:t.danger });
  }
  if (bibleRate >= 75) insights.push({ text:`Strong Bible culture: ${bibleRate}% of SS members bring their Bibles.`, color:t.success });
  else if (bibleRate > 0) insights.push({ text:`Bible participation at ${bibleRate}%. Encourage members to bring Bibles each week.`, color:t.warn });
  const topMonth = [...monthlyData].sort((a,b)=>b.chClose-a.chClose)[0];
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

      {/* KPI Summary — SS */}
      <div style={{ fontSize:11, fontWeight:700, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:10 }}>Sunday School</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(135px,1fr))", gap:12, marginBottom:16 }}>
        <KpiCard label="SS Begin"        value={ssBeginTotal}      sub="At opening"       icon="attendance" color={t.info} />
        <KpiCard label="SS Closing"      value={ssTotal}           sub="At close"         icon="attendance" color={t.gold} />
        <KpiCard label="Bible Begin"     value={bibleBeginTotal}   sub="Brought at start" icon="bible"      color="#9B59B6" />
        <KpiCard label="Bible Closing"   value={bibleTotal}        sub={`${bibleRate}% rate`} icon="bible" color="#7B3FBE" />
        <KpiCard label="First Timers"    value={fRec.reduce((s,r)=>s+(Number(r.first_timers)||0),0)} sub="SS" icon="plus" color="#E67E22" />
      </div>

      {/* KPI Summary — Church */}
      <div style={{ fontSize:11, fontWeight:700, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1.2, marginBottom:10 }}>Church Service</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(135px,1fr))", gap:12, marginBottom:24 }}>
        <KpiCard label="Church Begin"    value={chBeginTotal}      sub="At opening"       icon="cross"      color={t.info} />
        <KpiCard label="Church Closing"  value={chTotal}           sub="At close"         icon="cross"      color="#1A5DC8" />
        <KpiCard label="Avg Retention"   value={`${avgRetention}%`} sub="SS÷Church avg"  icon="analytics"  color={avgRetention>=80?t.success:avgRetention>=60?t.warn:t.danger} />
        <KpiCard label="First Timers"    value={fChurch.reduce((s,r)=>s+(Number(r.first_timers)||0),0)} sub="Church" icon="plus" color="#E67E22" />
        <KpiCard label="Months of Data"  value={allMonths.length}  sub="On record"        icon="settings"   color={t.success} />
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

// ─── SS REPORT PAGE ───────────────────────────────────────────────────────────
const SSReportPage = ({ db }) => {
  const { t, card, btnGold, btnOutline, sel, inp } = useThemeStyles();
  const { records, classes, loadAll } = db;
  const [refreshing, setRefreshing] = useState(false);

  // ── Canonical class display order ─────────────────────────────────────────
  // Any class name containing one of these tokens (case-insensitive) is
  // matched; unrecognised classes fall to the bottom in submission order.
  const CLASS_ORDER = [
    "children",
    "teen",
    "new convert",
    "prophet",
    "church age",
    "c.o.d",
    "joint",
  ];

  const classRank = (name = "") => {
    const lower = name.toLowerCase();
    const idx   = CLASS_ORDER.findIndex(token => lower.includes(token));
    return idx === -1 ? CLASS_ORDER.length : idx;   // unknown → end
  };

  // Filters
  const [filterDate, setFilterDate]   = useState(new Date().toISOString().slice(0,10));
  const [filterClass, setFilterClass] = useState("");

  const handleRefresh = async () => {
    setRefreshing(true);
    await loadAll();
    setRefreshing(false);
  };

  // Get all unique service dates sorted descending
  const allDates = [...new Set(records.map(r=>r.date).filter(Boolean))].sort().reverse();

  // Current session = records on filterDate (or latest date if none selected matches)
  const currentDate = filterDate && records.some(r=>r.date===filterDate)
    ? filterDate
    : allDates[0] || "";

  // Previous session = most recent date BEFORE currentDate
  const prevDate = allDates.find(d => d < currentDate) || "";

  const currentRecs = records.filter(r => r.date===currentDate && (!filterClass || r.class_name===filterClass));
  const prevRecs    = records.filter(r => r.date===prevDate    && (!filterClass || r.class_name===filterClass));

  // Build per-class comparison rows — include any class that has EVER had a
  // record (not just classes in the `classes` table), so data submitted by
  // any teacher always appears.
  const submittedClassNames = [...new Set(records.map(r => r.class_name).filter(Boolean))];

  const allClassNames = filterClass
    ? [filterClass]
    : (() => {
        const fromClasses = classes.filter(c => c.is_active === "YES").map(c => c.name);
        // merge: start with canonical list, then add any DB classes, then any
        // submitted names not yet in either — preserving canonical rank
        const merged = [...new Set([...fromClasses, ...submittedClassNames])];
        return merged;
      })();

  const rows = allClassNames.map(clsName => {
    const cur  = currentRecs.find(r => r.class_name === clsName);
    const prev = prevRecs.find(r   => r.class_name === clsName);

    const ssBegin  = Number(cur?.total_beginning) || 0;
    const ssClose  = Number(cur?.total_closing)   || 0;
    const bibBegin = Number(cur?.bibles_beginning) || 0;
    const bibClose = Number(cur?.bibles_closing)   || 0;
    const male     = Number(cur?.male_present)     || 0;
    const female   = Number(cur?.female_present)   || 0;
    const firstT   = Number(cur?.first_timers)     || 0;
    const visitors = Number(cur?.visitors)         || 0;

    const pSSBegin  = Number(prev?.total_beginning) || 0;
    const pSSClose  = Number(prev?.total_closing)   || 0;
    const pBibBegin = Number(prev?.bibles_beginning) || 0;
    const pBibClose = Number(prev?.bibles_closing)  || 0;

    const diff = (cur, prv) => {
      if (!prv && !cur) return null;
      const d   = cur - prv;
      const pct = prv > 0 ? Math.round((d / prv) * 100) : null;
      return { d, pct, up: d >= 0 };
    };

    return {
      cls: clsName,
      ssBegin, ssClose, bibBegin, bibClose,
      male, female, firstT, visitors,
      hasCurrent: !!cur,
      dSSBegin:  diff(ssBegin,  pSSBegin),
      dSSClose:  diff(ssClose,  pSSClose),
      dBibBegin: diff(bibBegin, pBibBegin),
      dBibClose: diff(bibClose, pBibClose),
      prevDate,
    };
  })
    // Only show rows that have data (current or any historical)
    .filter(r => r.hasCurrent || records.some(r2 => r2.class_name === r.cls))
    // ── Sort by canonical class order ─────────────────────────────────────
    .sort((a, b) => classRank(a.cls) - classRank(b.cls));

  // Totals row
  const tot = (key) => rows.reduce((s,r)=>s+(r[key]||0),0);

  // Export this report to Excel
  const exportReport = () => {
    const XLSX_mod = window.XLSX || (typeof XLSX !== "undefined" ? XLSX : null);
    if (!XLSX_mod) { alert("XLSX not available"); return; }
    const data = rows.map(r => ({
      "Class":            r.cls,
      "SS Begin":         r.ssBegin,
      "SS Close":         r.ssClose,
      "Bible Begin":      r.bibBegin,
      "Bible Close":      r.bibClose,
      "Male":             r.male,
      "Female":           r.female,
      "First Timers":     r.firstT,
      "Visitors":         r.visitors,
      "Prev SS Begin":    r.dSSBegin  ? r.ssBegin  - r.dSSBegin.d  : "",
      "Prev SS Close":    r.dSSClose  ? r.ssClose  - r.dSSClose.d  : "",
      "Prev Bible Begin": r.dBibBegin ? r.bibBegin - r.dBibBegin.d : "",
      "Prev Bible Close": r.dBibClose ? r.bibClose - r.dBibClose.d : "",
      "SS Begin Diff":    r.dSSBegin?.d  ?? "",
      "SS Close Diff":    r.dSSClose?.d  ?? "",
      "Date":             currentDate,
      "Prev Date":        prevDate,
    }));
    const wb = XLSX_mod.utils.book_new();
    const ws = XLSX_mod.utils.json_to_sheet(data);
    ws["!cols"] = Object.keys(data[0]||{}).map(()=>({wch:16}));
    XLSX_mod.utils.book_append_sheet(wb, ws, "SS Report");
    XLSX_mod.writeFile(wb, `SS_Report_${currentDate}.xlsx`);
  };

  // Diff cell component
  const DiffCell = ({ diff, style={} }) => {
    if (!diff) return <td style={{ ...style, color:t.textFaint, fontSize:11 }}>—</td>;
    const color = diff.up ? t.success : t.danger;
    const arrow = diff.up ? "▲" : "▼";
    return (
      <td style={{ ...style, fontFamily:"'Trebuchet MS',sans-serif", fontSize:12 }}>
        <div style={{ display:"flex", alignItems:"center", gap:4 }}>
          <span style={{ color, fontSize:13 }}>{arrow}</span>
          <span style={{ color, fontWeight:700 }}>{Math.abs(diff.d)}</span>
          {diff.pct !== null &&
            <span style={{ color, fontSize:10 }}>/ {diff.pct > 0 ? "+" : ""}{diff.pct}%</span>}
        </div>
      </td>
    );
  };

  const thS = { padding:"9px 12px", textAlign:"left", fontSize:10, fontWeight:700,
    fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1,
    borderBottom:`2px solid ${t.border}`, whiteSpace:"nowrap" };
  const tdS = { padding:"10px 12px", fontFamily:"'Trebuchet MS',sans-serif", fontSize:13,
    borderBottom:`1px solid ${t.border}18`, verticalAlign:"middle" };

  return (
    <div>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:20, flexWrap:"wrap", gap:12 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:3 }}>
            SS Comparison Report
          </div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            Current session vs previous session · auto-updates when new reports are submitted
          </div>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={handleRefresh} disabled={refreshing}
            style={{ display:"flex", alignItems:"center", gap:6, padding:"9px 16px", borderRadius:9,
              border:`1px solid ${t.gold}44`, background:t.gold+"11", color:t.gold,
              fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, cursor:"pointer" }}>
            <span style={{ fontSize:16, animation:refreshing?"spin 0.7s linear infinite":"none", display:"inline-block" }}>↻</span>
            {refreshing ? "Refreshing…" : "Refresh"}
          </button>
          <button onClick={exportReport}
            style={{ display:"flex", alignItems:"center", gap:6, padding:"9px 16px", borderRadius:9,
              border:`1px solid ${t.success}44`, background:t.success+"11", color:t.success,
              fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, cursor:"pointer" }}>
            <Icon name="export" size={14} color={t.success} /> Export Excel
          </button>
        </div>
      </div>

      {/* Filters */}
      <div style={{ ...card, marginBottom:20, display:"flex", gap:14, flexWrap:"wrap", alignItems:"flex-end" }}>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1 }}>
            Current Date (Service)
          </label>
          <select style={{ ...sel }} value={currentDate} onChange={e=>setFilterDate(e.target.value)}>
            {allDates.map(d=><option key={d} value={d}>{fmtDate(d)}</option>)}
            {!allDates.length && <option value="">No records yet</option>}
          </select>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
          <label style={{ fontSize:10, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1 }}>
            Class Filter
          </label>
          <select style={{ ...sel }} value={filterClass} onChange={e=>setFilterClass(e.target.value)}>
            <option value="">All Classes</option>
            {classes.map(c=><option key={c.name} value={c.name}>{c.name}</option>)}
          </select>
        </div>
        <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", paddingBottom:2 }}>
          Comparing <strong style={{ color:t.gold }}>{currentDate ? fmtDate(currentDate) : "—"}</strong>
          {" "}vs previous <strong style={{ color:t.info }}>{prevDate ? fmtDate(prevDate) : "—"}</strong>
        </div>
      </div>

      {/* Summary KPIs */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))", gap:12, marginBottom:20 }}>
        <KpiCard label="SS Begin"     value={tot("ssBegin")}  sub="Current session" icon="attendance" color={t.info} />
        <KpiCard label="SS Closing"   value={tot("ssClose")}  sub="Current session" icon="attendance" color={t.gold} />
        <KpiCard label="Bible Begin"  value={tot("bibBegin")} sub="Current session" icon="bible"      color="#9B59B6" />
        <KpiCard label="Bible Close"  value={tot("bibClose")} sub="Current session" icon="bible"      color="#7B3FBE" />
        <KpiCard label="First Timers" value={tot("firstT")}   sub="Current session" icon="plus"       color="#E67E22" />
        <KpiCard label="Visitors"     value={tot("visitors")} sub="Current session" icon="eye"        color={t.success} />
      </div>

      {/* Main Report Table */}
      <div style={{ ...card, padding:0, overflowX:"auto" }}>
        {/* Table header groups */}
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:1050 }}>
          <thead>
            {/* Group header — colours follow the active theme */}
            <tr style={{ background:t.sidebar }}>
              <th style={{ ...thS, color:"#FFFFFF", borderBottom:"none", padding:"8px 12px", background:t.sidebar }}></th>
              <th colSpan={4} style={{ ...thS, color:"#FFFFFF", textAlign:"center", borderBottom:"none",
                borderRight:`1px solid rgba(255,255,255,0.2)`, background:t.sidebar }}>
                📅 Current Date — {currentDate ? fmtDate(currentDate) : "No data"}
              </th>
              <th colSpan={4} style={{ ...thS, color:"#FFD700", textAlign:"center", borderBottom:"none",
                background:t.goldDark }}>
                🕐 vs Previous — {prevDate ? fmtDate(prevDate) : "No previous"}
              </th>
              <th colSpan={2} style={{ ...thS, color:"#FFFFFF", textAlign:"center", borderBottom:"none",
                background:t.info }}>
                Other
              </th>
            </tr>
            {/* Column headers */}
            <tr style={{ background:t.surfaceAlt }}>
              <th style={{ ...thS, color:t.gold, minWidth:160 }}>Class</th>
              <th style={{ ...thS, color:t.info }}>SS Begin</th>
              <th style={{ ...thS, color:t.gold }}>SS Close</th>
              <th style={{ ...thS, color:"#9B59B6" }}>Bible Begin</th>
              <th style={{ ...thS, color:"#7B3FBE", borderRight:`2px solid ${t.border}` }}>Bible Close</th>
              <th style={{ ...thS, color:t.info }}>Prev.SS Begin</th>
              <th style={{ ...thS, color:t.gold }}>Prev.SS Close</th>
              <th style={{ ...thS, color:"#9B59B6" }}>Prev.Bible Begin</th>
              <th style={{ ...thS, color:"#7B3FBE", borderRight:`2px solid ${t.border}` }}>Prev.Bible Close</th>
              <th style={{ ...thS, color:"#E67E22" }}>Male / Female</th>
              <th style={{ ...thS, color:t.success }}>1st T / Visitors</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.cls}
                style={{ background: i%2===0 ? "transparent" : t.surfaceAlt+"66" }}
                onMouseEnter={e=>e.currentTarget.style.background=t.surfaceHover}
                onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"transparent":t.surfaceAlt+"66"}>

                {/* Class name — full name, never truncated */}
                <td style={{ ...tdS, fontWeight:700, color:t.gold, whiteSpace:"nowrap", minWidth:160 }}>
                  {r.cls}
                  {!r.hasCurrent && <span style={{ fontSize:10, color:t.textMuted, display:"block" }}>no submission</span>}
                </td>

                {/* Current */}
                <td style={{ ...tdS, color:t.info, fontWeight:700 }}>{r.hasCurrent ? r.ssBegin : "—"}</td>
                <td style={{ ...tdS, color:t.gold, fontWeight:700 }}>{r.hasCurrent ? r.ssClose : "—"}</td>
                <td style={{ ...tdS, color:"#9B59B6", fontWeight:700 }}>{r.hasCurrent ? r.bibBegin : "—"}</td>
                <td style={{ ...tdS, color:"#7B3FBE", fontWeight:700, borderRight:`2px solid ${t.border}` }}>{r.hasCurrent ? r.bibClose : "—"}</td>

                {/* Diff columns */}
                <DiffCell diff={r.dSSBegin}  style={tdS} />
                <DiffCell diff={r.dSSClose}  style={tdS} />
                <DiffCell diff={r.dBibBegin} style={tdS} />
                <DiffCell diff={r.dBibClose} style={{ ...tdS, borderRight:`2px solid ${t.border}` }} />

                {/* Male/Female */}
                <td style={{ ...tdS }}>
                  <span style={{ color:t.info }}>{r.male}M</span>
                  <span style={{ color:t.textFaint, margin:"0 4px" }}>·</span>
                  <span style={{ color:"#E67E22" }}>{r.female}F</span>
                </td>

                {/* First timers / Visitors */}
                <td style={{ ...tdS }}>
                  <span style={{ color:"#E67E22", fontWeight:700 }}>{r.firstT}</span>
                  <span style={{ color:t.textFaint, margin:"0 4px" }}>/</span>
                  <span style={{ color:t.success }}>{r.visitors}</span>
                </td>
              </tr>
            ))}

            {/* Totals row */}
            {rows.length > 1 && (() => {
              // Aggregate prev values for totals diff
              const pTotSSBegin  = prevRecs.reduce((s,r)=>s+(Number(r.total_beginning)||0),0);
              const pTotSSClose  = prevRecs.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
              const pTotBibBegin = prevRecs.reduce((s,r)=>s+(Number(r.bibles_beginning)||0),0);
              const pTotBibClose = prevRecs.reduce((s,r)=>s+(Number(r.bibles_closing)||0),0);
              const mkDiff = (cur, prv) => {
                if (!prv && !cur) return null;
                const d = cur - prv;
                const pct = prv > 0 ? Math.round((d/prv)*100) : null;
                return { d, pct, up: d >= 0 };
              };
              const tSSBegin  = tot("ssBegin");
              const tSSClose  = tot("ssClose");
              const tBibBegin = tot("bibBegin");
              const tBibClose = tot("bibClose");
              const diffs = [
                mkDiff(tSSBegin,  pTotSSBegin),
                mkDiff(tSSClose,  pTotSSClose),
                mkDiff(tBibBegin, pTotBibBegin),
                mkDiff(tBibClose, pTotBibClose),
              ];
              const TotDiffCell = ({ diff, extra={} }) => {
                if (!diff) return <td style={{ ...tdS, color:t.textFaint, fontSize:11, fontWeight:700, ...extra }}>—</td>;
                const color = diff.up ? t.success : t.danger;
                const arrow = diff.up ? "▲" : "▼";
                return (
                  <td style={{ ...tdS, fontWeight:700, ...extra }}>
                    <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                      <span style={{ color, fontSize:14 }}>{arrow}</span>
                      <span style={{ color, fontWeight:800 }}>{Math.abs(diff.d)}</span>
                      {diff.pct !== null &&
                        <span style={{ color, fontSize:10 }}>/ {diff.d >= 0 ? "+" : ""}{diff.pct}%</span>}
                    </div>
                  </td>
                );
              };
              return (
                <tr style={{ background:t.sidebar+"22", borderTop:`3px solid ${t.sidebar}` }}>
                  <td style={{ ...tdS, fontWeight:800, color:t.sidebar, fontSize:12, textTransform:"uppercase", letterSpacing:0.8 }}>
                    TOTAL
                  </td>
                  <td style={{ ...tdS, fontWeight:800, color:t.info }}>{tSSBegin}</td>
                  <td style={{ ...tdS, fontWeight:800, color:t.gold }}>{tSSClose}</td>
                  <td style={{ ...tdS, fontWeight:800, color:"#9B59B6" }}>{tBibBegin}</td>
                  <td style={{ ...tdS, fontWeight:800, color:"#7B3FBE", borderRight:`2px solid ${t.border}` }}>{tBibClose}</td>
                  <TotDiffCell diff={diffs[0]} />
                  <TotDiffCell diff={diffs[1]} />
                  <TotDiffCell diff={diffs[2]} />
                  <TotDiffCell diff={diffs[3]} extra={{ borderRight:`2px solid ${t.border}` }} />
                  <td style={{ ...tdS, color:t.info, fontWeight:700 }}>
                    {tot("male")}M · <span style={{ color:"#E67E22" }}>{tot("female")}F</span>
                  </td>
                  <td style={{ ...tdS }}>
                    <span style={{ color:"#E67E22", fontWeight:800 }}>{tot("firstT")}</span>
                    <span style={{ color:t.textFaint, margin:"0 4px" }}>/</span>
                    <span style={{ color:t.success, fontWeight:700 }}>{tot("visitors")}</span>
                  </td>
                </tr>
              );
            })()}

            {!rows.length && (
              <tr>
                <td colSpan={11} style={{ ...tdS, textAlign:"center", padding:52, color:t.textMuted }}>
                  No records found for {currentDate ? fmtDate(currentDate) : "selected date"}.
                  Submit SS reports to see them here.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Legend */}
      <div style={{ ...card, marginTop:16, display:"flex", gap:20, flexWrap:"wrap" }}>
        <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
          <strong style={{ color:t.text }}>Reading the report:</strong>
        </div>
        {[
          { label:"▲ Green = increase vs previous", color:t.success },
          { label:"▼ Red = decrease vs previous",   color:t.danger },
          { label:"— = no previous data",           color:t.textMuted },
          { label:"Number / % = change amount and percentage", color:t.textMuted },
        ].map(item=>(
          <span key={item.label} style={{ fontSize:11, color:item.color, fontFamily:"'Trebuchet MS',sans-serif" }}>
            {item.label}
          </span>
        ))}
      </div>
    </div>
  );
};

// ─── AI ASSISTANT PAGE ────────────────────────────────────────────────────────
const AIAssistantPage = ({ db, user }) => {
  const { t, card } = useThemeStyles();
  const { mode } = useTheme();
  const { records, classes, churchRecs } = db;

  // API key — stored in sessionStorage so it survives page re-renders but not tab close
  const [apiKey, setApiKey]   = useState(() => sessionStorage.getItem("uct_ai_key") || "");
  const [keyInput, setKeyInput] = useState("");
  const [showKeyInput, setShowKeyInput] = useState(false);

  const saveKey = () => {
    const k = keyInput.trim();
    if (!k.startsWith("sk-ant-")) { alert("Invalid key — must start with sk-ant-"); return; }
    sessionStorage.setItem("uct_ai_key", k);
    setApiKey(k);
    setKeyInput("");
    setShowKeyInput(false);
  };
  const forgetKey = () => { sessionStorage.removeItem("uct_ai_key"); setApiKey(""); };

  const [messages, setMessages] = useState([
    { role:"assistant", content:"Hello! I'm your Sunday School AI Assistant ✨\n\nI have full access to your attendance records, class data, and church reports. Ask me anything — summaries, trends, comparisons, recommendations, or anything else about your ministry data." }
  ]);
  const [input, setInput]         = useState("");
  const [loading, setLoading]     = useState(false);
  const [activeTab, setActiveTab] = useState("chat");
  const bottomRef = useRef(null);

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior:"smooth" }); }, [messages]);

  // Build a compact but rich data context for Claude
  const buildContext = () => {
    const totalRec   = records.length;
    const allDates   = [...new Set(records.map(r=>r.date).filter(Boolean))].sort();
    const latestDate = allDates[allDates.length-1] || "none";
    const prevDate   = allDates[allDates.length-2] || "none";

    const classSummary = classes.filter(c=>c.is_active==="YES").map(cls => {
      const recs   = records.filter(r=>r.class_name===cls.name);
      const latest = recs.filter(r=>r.date===latestDate);
      const prev   = recs.filter(r=>r.date===prevDate);
      const avgClose = recs.length ? Math.round(recs.reduce((s,r)=>s+(Number(r.total_closing)||0),0)/recs.length) : 0;
      return {
        class: cls.name, totalSubmissions: recs.length,
        latestClose: latest.reduce((s,r)=>s+(Number(r.total_closing)||0),0),
        prevClose:   prev.reduce((s,r)=>s+(Number(r.total_closing)||0),0),
        avgClose,
        totalFirstTimers: recs.reduce((s,r)=>s+(Number(r.first_timers)||0),0),
        totalVisitors:    recs.reduce((s,r)=>s+(Number(r.visitors)||0),0),
      };
    });

    const monthMap = {};
    records.forEach(r => {
      const m = (r.date||"").slice(0,7); if(!m) return;
      if(!monthMap[m]) monthMap[m]={month:m,ssTotal:0,bibles:0,firstT:0,sessions:0};
      monthMap[m].ssTotal  += Number(r.total_closing)||0;
      monthMap[m].bibles   += Number(r.bibles_closing)||0;
      monthMap[m].firstT   += Number(r.first_timers)||0;
      monthMap[m].sessions++;
    });
    const monthly = Object.values(monthMap).sort((a,b)=>a.month>b.month?1:-1).slice(-6);

    const churchTotal = churchRecs.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
    const ssTotal     = records.reduce((s,r)=>s+(Number(r.total_closing)||0),0);

    return `You are an AI assistant for ${user?.name||"an admin"} at a church Sunday School management system.
Current user role: ${user?.role||"admin"}
Today: ${new Date().toDateString()}

=== ATTENDANCE DATA SUMMARY ===
Total SS records in database: ${totalRec}
Latest session date: ${latestDate}
Previous session date: ${prevDate}
Total church attendance records: ${churchRecs.length}

=== PER-CLASS SUMMARY (latest vs previous) ===
${JSON.stringify(classSummary, null, 2)}

=== LAST 6 MONTHS TREND ===
${JSON.stringify(monthly, null, 2)}

=== TOTALS ===
Total SS closing attendance (all time): ${ssTotal}
Total Church closing attendance (all time): ${churchTotal}
Active classes: ${classes.filter(c=>c.is_active==="YES").map(c=>c.name).join(", ")}

Give clear, pastoral, ministry-focused insights. Use emojis sparingly. Be concise but thorough. Reference specific numbers from the data when relevant.`;
  };

  const sendMessage = async () => {
    if (!input.trim() || loading || !apiKey) return;
    const userMsg = input.trim();
    setInput("");
    setMessages(m => [...m, { role:"user", content:userMsg }]);
    setLoading(true);

    try {
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          system: buildContext(),
          messages: [
            ...messages.slice(-10).map(m => ({ role:m.role, content:m.content })),
            { role:"user", content:userMsg },
          ],
        }),
      });

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        throw new Error(err?.error?.message || `HTTP ${response.status}`);
      }

      const data  = await response.json();
      const reply = data.content?.map(b => b.text||"").join("") || "Sorry, I couldn't get a response.";
      setMessages(m => [...m, { role:"assistant", content:reply }]);
    } catch (e) {
      setMessages(m => [...m, { role:"assistant", content:`⚠️ ${e.message || "Connection failed. Please try again."}` }]);
    }
    setLoading(false);
  };

  // Quick insight prompts
  const QUICK_PROMPTS = [
    { icon:"📊", label:"Attendance trend", prompt:"Summarise the attendance trend over the last few months. Which classes are growing and which are declining?" },
    { icon:"🏆", label:"Top performing class", prompt:"Which class has the best attendance and engagement? Give me a brief analysis." },
    { icon:"📖", label:"Bible participation", prompt:"How is Bible participation across the classes? Any concerns or highlights?" },
    { icon:"👋", label:"First timers report", prompt:"Give me a report on first-time visitors. Which sessions brought the most new people?" },
    { icon:"⚠️", label:"Areas of concern", prompt:"Are there any attendance drops or concerning trends I should address as a leader?" },
    { icon:"💡", label:"AI recommendations", prompt:"Based on the data, what 3 specific recommendations do you have to improve Sunday School attendance and engagement?" },
  ];

  // --- Insights tab: auto-generated key metrics ---
  const allDates   = [...new Set(records.map(r=>r.date).filter(Boolean))].sort();
  const latestDate = allDates[allDates.length-1] || "";
  const prevDate   = allDates[allDates.length-2] || "";
  const latestRecs = records.filter(r=>r.date===latestDate);
  const prevRecs   = records.filter(r=>r.date===prevDate);
  const latestTotal = latestRecs.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const prevTotal   = prevRecs.reduce((s,r)=>s+(Number(r.total_closing)||0),0);
  const diff = latestTotal - prevTotal;
  const pct  = prevTotal > 0 ? Math.round((diff/prevTotal)*100) : null;

  const inp = { background: isDark(mode) ? "rgba(255,255,255,0.07)" : t.surfaceAlt,
    border:`1px solid ${t.border}`, borderRadius:12, padding:"12px 16px",
    color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:14,
    outline:"none", resize:"none", width:"100%", boxSizing:"border-box" };

  const tabBtn = (id, label) => (
    <button onClick={()=>setActiveTab(id)} style={{
      padding:"8px 20px", borderRadius:8, border:"none", cursor:"pointer",
      fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, fontWeight:600,
      background: activeTab===id ? t.sidebar : "transparent",
      color: activeTab===id ? "#FFFFFF" : t.textMuted,
      transition:"all 0.15s",
    }}>{label}</button>
  );

  return (
    <div>
      {/* Header */}
      <div style={{ marginBottom:20, display:"flex", alignItems:"center", gap:14, flexWrap:"wrap" }}>
        <div style={{ width:46, height:46, borderRadius:12, flexShrink:0,
          background:`linear-gradient(135deg,${t.sidebar},${t.goldLight||t.sidebar}cc)`,
          display:"flex", alignItems:"center", justifyContent:"center",
          boxShadow:`0 4px 18px ${t.sidebar}55` }}>
          <Icon name="ai" size={24} color="#FFFFFF" />
        </div>
        <div>
          <div style={{ fontSize:22, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif" }}>
            AI Assistant
            <span style={{ marginLeft:10, fontSize:11, fontWeight:700, padding:"3px 9px",
              borderRadius:20, background:ACTIVE_COLOR+"22", color:ACTIVE_COLOR,
              border:`1px solid ${ACTIVE_COLOR}44`, verticalAlign:"middle" }}>POWERED BY CLAUDE</span>
          </div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            Ask questions about your attendance data, get insights, and receive ministry recommendations.
          </div>
        </div>
      </div>

      {/* ── API Key Setup Banner ───────────────────────────────────────────── */}
      {!apiKey ? (
        <div style={{ ...card, marginBottom:18, border:`2px dashed ${ACTIVE_COLOR}66`,
          background: isDark(mode) ? `${ACTIVE_COLOR}11` : `${ACTIVE_COLOR}08`,
          display:"flex", flexDirection:"column", gap:14 }}>
          <div style={{ display:"flex", alignItems:"flex-start", gap:12 }}>
            <span style={{ fontSize:28 }}>🔑</span>
            <div>
              <div style={{ fontWeight:700, fontSize:15, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>
                Connect your Anthropic API Key
              </div>
              <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", lineHeight:1.6 }}>
                To use the AI Assistant you need an Anthropic API key. Your key is stored only in this browser session — it's never sent anywhere except directly to Anthropic's API.
                <br/>Get a key at <a href="https://console.anthropic.com/keys" target="_blank" rel="noreferrer"
                  style={{ color:ACTIVE_COLOR, textDecoration:"none", fontWeight:600 }}>console.anthropic.com/keys</a>
              </div>
            </div>
          </div>
          {showKeyInput ? (
            <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
              <input
                type="password"
                placeholder="sk-ant-api03-…"
                value={keyInput}
                onChange={e=>setKeyInput(e.target.value)}
                onKeyDown={e=>e.key==="Enter" && saveKey()}
                style={{ flex:1, minWidth:260, background: isDark(mode)?"rgba(255,255,255,0.07)":t.surfaceAlt,
                  border:`1px solid ${t.border}`, borderRadius:9, padding:"10px 14px",
                  color:t.text, fontFamily:"monospace", fontSize:13, outline:"none" }}
              />
              <button onClick={saveKey} style={{ padding:"10px 20px", borderRadius:9, border:"none",
                background:ACTIVE_COLOR, color:"#fff", fontFamily:"'Trebuchet MS',sans-serif",
                fontWeight:700, fontSize:13, cursor:"pointer" }}>Save Key</button>
              <button onClick={()=>setShowKeyInput(false)} style={{ padding:"10px 14px", borderRadius:9,
                border:`1px solid ${t.border}`, background:"transparent", color:t.textMuted,
                fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, cursor:"pointer" }}>Cancel</button>
            </div>
          ) : (
            <button onClick={()=>setShowKeyInput(true)} style={{ alignSelf:"flex-start", padding:"10px 22px",
              borderRadius:9, border:`1px solid ${ACTIVE_COLOR}`, background:ACTIVE_COLOR,
              color:"#fff", fontFamily:"'Trebuchet MS',sans-serif", fontWeight:700, fontSize:13, cursor:"pointer" }}>
              + Enter API Key
            </button>
          )}
        </div>
      ) : (
        <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:16,
          padding:"8px 14px", borderRadius:9, background: t.success+"18",
          border:`1px solid ${t.success}44`, width:"fit-content" }}>
          <span style={{ color:t.success, fontSize:13 }}>✓</span>
          <span style={{ fontSize:12, color:t.success, fontFamily:"'Trebuchet MS',sans-serif", fontWeight:600 }}>
            API key connected
          </span>
          <span style={{ fontSize:12, color:t.textMuted, fontFamily:"monospace" }}>
            {apiKey.slice(0,14)}…
          </span>
          <button onClick={forgetKey} style={{ padding:"3px 10px", borderRadius:6, border:`1px solid ${t.danger}44`,
            background:t.danger+"11", color:t.danger, fontFamily:"'Trebuchet MS',sans-serif",
            fontSize:11, cursor:"pointer" }}>Remove</button>
        </div>
      )}

      {/* Tabs */}
      <div style={{ display:"flex", gap:4, marginBottom:16, background:t.surfaceAlt, borderRadius:10, padding:4, width:"fit-content" }}>
        {tabBtn("chat",     "💬 Chat")}
        {tabBtn("insights", "📊 Quick Insights")}
      </div>

      {activeTab === "insights" && (
        <div>
          {/* KPI cards */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:14, marginBottom:20 }}>
            {[
              { label:"Latest Session Total", value:latestTotal, sub:latestDate, color:t.gold },
              { label:"vs Previous Session",  value:(diff>=0?"+":"")+diff+(pct!==null?` (${pct}%)`:""), sub:prevDate||"No prev data", color:diff>=0?t.success:t.danger },
              { label:"Total Records",        value:records.length, sub:"all time", color:t.info },
              { label:"Active Classes",       value:classes.filter(c=>c.is_active==="YES").length, sub:"classes", color:"#9B59B6" },
              { label:"All-time First Timers",value:records.reduce((s,r)=>s+(Number(r.first_timers)||0),0), sub:"total", color:ACTIVE_COLOR },
              { label:"Total Visitors",       value:records.reduce((s,r)=>s+(Number(r.visitors)||0),0), sub:"all time", color:"#E67E22" },
            ].map(k=>(
              <div key={k.label} style={{ ...card, padding:"16px 18px", borderTop:`3px solid ${k.color}` }}>
                <div style={{ fontSize:13, fontWeight:800, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                  marginBottom:8, lineHeight:1.3, letterSpacing:0.2 }}>{k.label}</div>
                <div style={{ fontSize:30, fontWeight:900, color:k.color, fontFamily:"'Georgia',serif", lineHeight:1 }}>{k.value}</div>
                <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:5 }}>{k.sub}</div>
              </div>
            ))}
          </div>

          {/* Quick prompt tiles */}
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:12, fontWeight:600 }}>
            Ask the AI about your data:
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))", gap:12 }}>
            {QUICK_PROMPTS.map(p=>(
              <button key={p.label} onClick={()=>{ setActiveTab("chat"); setInput(p.prompt); setTimeout(()=>document.getElementById("ai-input")?.focus(),100); }}
                style={{ ...card, padding:"14px 18px", textAlign:"left", cursor:"pointer", border:`1px solid ${t.border}`,
                  display:"flex", alignItems:"flex-start", gap:12,
                  transition:"all 0.15s", background:t.surface }}>
                <span style={{ fontSize:22 }}>{p.icon}</span>
                <div>
                  <div style={{ fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, marginBottom:4 }}>{p.label}</div>
                  <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", lineHeight:1.5 }}>{p.prompt.slice(0,80)}…</div>
                </div>
              </button>
            ))}
          </div>
        </div>
      )}

      {activeTab === "chat" && (
        <div style={{ display:"flex", flexDirection:"column", gap:0 }}>
          {/* Chat window */}
          <div style={{ ...card, padding:0, height:440, overflowY:"auto", display:"flex", flexDirection:"column" }}>
            <div style={{ flex:1, overflowY:"auto", padding:"18px 20px", display:"flex", flexDirection:"column", gap:14 }}>
              {messages.map((msg, i) => (
                <div key={i} style={{ display:"flex", gap:10, alignItems:"flex-start",
                  flexDirection: msg.role==="user" ? "row-reverse" : "row" }}>
                  {/* Avatar */}
                  <div style={{ width:32, height:32, borderRadius:"50%", flexShrink:0,
                    background: msg.role==="user"
                      ? `linear-gradient(135deg,${ACTIVE_COLOR},${ACTIVE_COLOR}bb)`
                      : `linear-gradient(135deg,${t.sidebar},${t.goldLight||t.sidebar}cc)`,
                    display:"flex", alignItems:"center", justifyContent:"center",
                    fontSize:14, fontWeight:700, color:"#FFFFFF",
                    fontFamily:"'Trebuchet MS',sans-serif" }}>
                    {msg.role==="user" ? (user?.name?.[0]||"U") : "✦"}
                  </div>
                  {/* Bubble */}
                  <div style={{
                    maxWidth:"75%", padding:"11px 16px", borderRadius:
                      msg.role==="user" ? "18px 4px 18px 18px" : "4px 18px 18px 18px",
                    background: msg.role==="user"
                      ? `linear-gradient(135deg,${t.sidebar},${t.goldLight||t.sidebar}dd)`
                      : isDark(mode) ? "rgba(255,255,255,0.07)" : t.surfaceAlt,
                    color: msg.role==="user" ? "#FFFFFF" : t.text,
                    fontFamily:"'Trebuchet MS',sans-serif", fontSize:13.5, lineHeight:1.65,
                    whiteSpace:"pre-wrap", border: msg.role==="assistant" ? `1px solid ${t.border}` : "none",
                  }}>
                    {msg.content}
                  </div>
                </div>
              ))}
              {loading && (
                <div style={{ display:"flex", gap:10, alignItems:"center" }}>
                  <div style={{ width:32, height:32, borderRadius:"50%",
                    background:`linear-gradient(135deg,${t.sidebar},${t.goldLight||t.sidebar}cc)`,
                    display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:14 }}>✦</div>
                  <div style={{ padding:"12px 18px", borderRadius:"4px 18px 18px 18px",
                    background: isDark(mode) ? "rgba(255,255,255,0.07)" : t.surfaceAlt,
                    border:`1px solid ${t.border}`, display:"flex", gap:6, alignItems:"center" }}>
                    {[0,1,2].map(i=>(
                      <div key={i} style={{ width:8, height:8, borderRadius:"50%", background:t.sidebar,
                        animation:`aiBounce 1.2s ease-in-out ${i*0.2}s infinite` }} />
                    ))}
                    <style>{`@keyframes aiBounce{0%,80%,100%{transform:scale(0.6);opacity:0.4}40%{transform:scale(1);opacity:1}}`}</style>
                  </div>
                </div>
              )}
              <div ref={bottomRef} />
            </div>
          </div>

          {/* Input area */}
          <div style={{ ...card, marginTop:12, display:"flex", gap:10, alignItems:"flex-end" }}>
            <textarea id="ai-input" rows={2} style={{ ...inp, flex:1, minHeight:52,
              opacity: apiKey ? 1 : 0.5 }}
              placeholder={apiKey ? "Ask about attendance trends, class performance, recommendations…" : "Enter your API key above to start chatting…"}
              value={input} onChange={e=>setInput(e.target.value)}
              disabled={!apiKey}
              onKeyDown={e=>{ if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); sendMessage(); } }} />
            <button onClick={sendMessage} disabled={loading || !input.trim() || !apiKey}
              style={{ padding:"12px 22px", borderRadius:12, border:"none", cursor:"pointer",
                background: (loading||!input.trim()) ? t.surfaceAlt : `linear-gradient(135deg,${t.sidebar},${t.goldLight||t.sidebar})`,
                color: (loading||!input.trim()) ? t.textFaint : "#FFFFFF",
                fontFamily:"'Trebuchet MS',sans-serif", fontSize:14, fontWeight:700,
                display:"flex", alignItems:"center", gap:7, transition:"all 0.15s",
                boxShadow: (!loading&&input.trim()) ? `0 4px 14px ${t.sidebar}55` : "none" }}>
              <Icon name="ai" size={16} color={(loading||!input.trim()) ? t.textFaint : "#FFFFFF"} />
              Send
            </button>
          </div>

          {/* Quick prompt chips */}
          <div style={{ marginTop:10, display:"flex", gap:8, flexWrap:"wrap" }}>
            {QUICK_PROMPTS.map(p=>(
              <button key={p.label} onClick={()=>{ setInput(p.prompt); document.getElementById("ai-input")?.focus(); }}
                style={{ padding:"6px 12px", borderRadius:20, border:`1px solid ${t.border}`,
                  background:"transparent", color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
                  fontSize:12, cursor:"pointer", display:"flex", alignItems:"center", gap:5 }}>
                {p.icon} {p.label}
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

// ─── LESSONS PAGE ─────────────────────────────────────────────────────────────
const useLessonsStore = () => {
  const KEY = "uct_lessons_v1";
  const load = () => { try { return JSON.parse(localStorage.getItem(KEY)||"[]"); } catch { return []; } };
  const [lessons, setLessons] = useState(load);
  const save = (arr) => { localStorage.setItem(KEY, JSON.stringify(arr)); setLessons(arr); };
  const addLesson    = (l)  => save([{ ...l, id: uid(), createdAt: new Date().toISOString() }, ...lessons]);
  const updateLesson = (id, upd) => save(lessons.map(l => l.id===id ? { ...l, ...upd } : l));
  const deleteLesson = (id) => save(lessons.filter(l => l.id!==id));
  return { lessons, addLesson, updateLesson, deleteLesson };
};

// ─── ALL LESSONS MODAL ───────────────────────────────────────────────────────
const AllLessonsModal = ({ lessons, onClose, deleteLesson, startEdit }) => {
  const { t, card, inp, sel, btnGhost, th, td } = useThemeStyles();
  const [search, setSearch]           = useState("");
  const [filterClass, setFilterClass] = useState("");
  const [filterTeacher, setFilterTeacher] = useState("");
  const [filterMonth, setFilterMonth] = useState("");
  const [filterBible, setFilterBible] = useState("");
  const [sortBy, setSortBy]           = useState("date_desc");
  const [expandedId, setExpandedId]   = useState(null);

  const allClasses  = [...new Set(lessons.map(l=>l.class_name).filter(Boolean))].sort();
  const allTeachers = [...new Set(lessons.map(l=>l.teacher_name).filter(Boolean))].sort();
  const allMonths   = [...new Set(lessons.map(l=>(l.date||"").slice(0,7)).filter(Boolean))].sort().reverse();

  const exportFromModal = (rows, label) => {
    if (!rows.length) { alert("No lessons to export."); return; }
    const headers = ["Date","Class_Name","Teacher_Name","Topic","Bible_References","Memory_Verse","Outline","Key_Points","Assignment","Duration_Mins","Notes","Created_At"];
    const data = rows.map(l => [
      l.date||"", l.class_name||"", l.teacher_name||"", l.topic||"",
      l.bible_references||"", l.memory_verse||"", l.outline||"",
      l.key_points||"", l.assignment||"", l.duration_mins||"",
      l.notes||"", l.createdAt ? l.createdAt.slice(0,10) : "",
    ]);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    ws["!cols"] = [12,18,18,30,22,28,40,30,26,14,26,14].map(wch=>({wch}));
    XLSX.utils.book_append_sheet(wb, ws, "Lessons");
    const classes  = [...new Set(rows.map(r=>r.class_name).filter(Boolean))];
    const teachers = [...new Set(rows.map(r=>r.teacher_name).filter(Boolean))];
    const totalMins = rows.reduce((s,r)=>s+(Number(r.duration_mins)||0),0);
    const summary = [
      ["UCT Lesson Register Export"],
      ["Exported on", new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"})],
      ["Filter", label],
      [""],
      ["Total Lessons",  rows.length],
      ["Total Classes",  classes.length],
      ["Total Teachers", teachers.length],
      ["Total Duration", totalMins ? `${totalMins} mins` : "—"],
      [""],
      ["Classes:", classes.join(", ")||"—"],
      ["Teachers:", teachers.join(", ")||"—"],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summary);
    wsSummary["!cols"] = [{wch:20},{wch:50}];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");
    XLSX.writeFile(wb, `UCT_Lessons_${label.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const filtered = lessons.filter(l => {
    if (filterClass   && l.class_name   !== filterClass)         return false;
    if (filterTeacher && l.teacher_name !== filterTeacher)       return false;
    if (filterMonth   && !(l.date||"").startsWith(filterMonth))  return false;
    if (filterBible   && !(l.bible_references||"").toLowerCase().includes(filterBible.toLowerCase())) return false;
    if (search) {
      const q = search.toLowerCase();
      return (l.topic||"").toLowerCase().includes(q) ||
             (l.teacher_name||"").toLowerCase().includes(q) ||
             (l.class_name||"").toLowerCase().includes(q) ||
             (l.bible_references||"").toLowerCase().includes(q) ||
             (l.memory_verse||"").toLowerCase().includes(q) ||
             (l.outline||"").toLowerCase().includes(q) ||
             (l.notes||"").toLowerCase().includes(q);
    }
    return true;
  }).sort((a,b) => {
    if (sortBy==="date_desc")    return (b.date||"").localeCompare(a.date||"");
    if (sortBy==="date_asc")     return (a.date||"").localeCompare(b.date||"");
    if (sortBy==="topic_asc")    return (a.topic||"").localeCompare(b.topic||"");
    if (sortBy==="class_asc")    return (a.class_name||"").localeCompare(b.class_name||"");
    if (sortBy==="teacher_asc")  return (a.teacher_name||"").localeCompare(b.teacher_name||"");
    return 0;
  });

  const hasFilter = search||filterClass||filterTeacher||filterMonth||filterBible;
  const clearAll  = () => { setSearch(""); setFilterClass(""); setFilterTeacher(""); setFilterMonth(""); setFilterBible(""); };

  const fmtDate = (d) => {
    if (!d) return "—";
    const dt = new Date(d);
    return isNaN(dt) ? d : dt.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"});
  };

  return (
    <div style={{ position:"fixed", inset:0, zIndex:800, display:"flex", flexDirection:"column" }}>
      {/* Backdrop */}
      <div onClick={onClose} style={{ position:"absolute", inset:0, background:"rgba(0,0,0,0.62)" }} />

      {/* Panel */}
      <div style={{ position:"relative", margin:"auto", width:"min(98vw,1050px)", maxHeight:"92vh",
        background:t.surface, borderRadius:18, display:"flex", flexDirection:"column",
        boxShadow:"0 24px 80px rgba(0,0,0,0.45)", overflow:"hidden" }}>

        {/* Header */}
        <div style={{ padding:"18px 24px 14px", borderBottom:`1px solid ${t.border}`,
          background:t.surfaceAlt, display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0 }}>
          <div>
            <div style={{ fontSize:20, fontWeight:800, color:t.gold, fontFamily:"'Georgia',serif" }}>
              📚 All Submitted Lessons
            </div>
            <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:2 }}>
              {filtered.length} of {lessons.length} lessons shown
            </div>
          </div>
          <div style={{ display:"flex", gap:8, alignItems:"center" }}>
            <button
              onClick={()=>exportFromModal(filtered, (filterClass||filterTeacher||filterMonth||filterBible||search) ? `Filtered_(${filtered.length})` : "All_Lessons")}
              style={{ background:t.success+"18", border:`1px solid ${t.success}44`, borderRadius:9,
                padding:"8px 14px", cursor:"pointer", color:t.success, fontFamily:"'Trebuchet MS',sans-serif",
                fontSize:13, display:"flex", alignItems:"center", gap:6, fontWeight:700 }}>
              <Icon name="export" size={14} color={t.success} />
              Export {(filterClass||filterTeacher||filterMonth||filterBible||search) ? `Filtered (${filtered.length})` : `All (${lessons.length})`}
            </button>
            <button onClick={onClose}
              style={{ background:t.danger+"18", border:`1px solid ${t.danger}44`, borderRadius:9,
                padding:"8px 14px", cursor:"pointer", color:t.danger, fontFamily:"'Trebuchet MS',sans-serif",
                fontSize:13, display:"flex", alignItems:"center", gap:6, fontWeight:700 }}>
              <Icon name="close" size={14} color={t.danger} /> Close
            </button>
          </div>
        </div>

        {/* Filters bar */}
        <div style={{ padding:"12px 20px", borderBottom:`1px solid ${t.border}`,
          background:t.surface, display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", flexShrink:0 }}>
          {/* Search */}
          <input style={{ ...inp, minWidth:190, flex:2 }}
            placeholder="🔍 Search topic, teacher, class, verse, notes…"
            value={search} onChange={e=>setSearch(e.target.value)} />
          {/* Class */}
          <select style={{ ...sel, minWidth:140 }} value={filterClass} onChange={e=>setFilterClass(e.target.value)}>
            <option value="">All Classes</option>
            {allClasses.map(c=><option key={c} value={c}>{c}</option>)}
          </select>
          {/* Teacher */}
          <select style={{ ...sel, minWidth:150 }} value={filterTeacher} onChange={e=>setFilterTeacher(e.target.value)}>
            <option value="">All Teachers</option>
            {allTeachers.map(tc=><option key={tc} value={tc}>{tc}</option>)}
          </select>
          {/* Month */}
          <select style={{ ...sel, minWidth:130 }} value={filterMonth} onChange={e=>setFilterMonth(e.target.value)}>
            <option value="">All Months</option>
            {allMonths.map(m=>{
              const [yr,mo]=m.split("-");
              return <option key={m} value={m}>{new Date(yr,mo-1).toLocaleString("default",{month:"short",year:"numeric"})}</option>;
            })}
          </select>
          {/* Bible ref */}
          <input style={{ ...inp, minWidth:130, flex:1 }}
            placeholder="📖 Bible ref…"
            value={filterBible} onChange={e=>setFilterBible(e.target.value)} />
          {/* Sort */}
          <select style={{ ...sel, minWidth:140 }} value={sortBy} onChange={e=>setSortBy(e.target.value)}>
            <option value="date_desc">Newest First</option>
            <option value="date_asc">Oldest First</option>
            <option value="topic_asc">Topic A–Z</option>
            <option value="class_asc">Class A–Z</option>
            <option value="teacher_asc">Teacher A–Z</option>
          </select>
          {hasFilter && (
            <button onClick={clearAll}
              style={{ padding:"8px 13px", borderRadius:9, border:`1px solid ${t.danger}44`,
                background:t.danger+"11", color:t.danger, fontFamily:"'Trebuchet MS',sans-serif",
                fontSize:13, cursor:"pointer", whiteSpace:"nowrap" }}>
              ✕ Clear All
            </button>
          )}
        </div>

        {/* Summary chips */}
        {filtered.length > 0 && (
          <div style={{ padding:"8px 20px 0", display:"flex", gap:8, flexWrap:"wrap", flexShrink:0 }}>
            {[
              { label:`${filtered.length} Lesson${filtered.length!==1?"s":""}`, color:t.gold },
              { label:`${[...new Set(filtered.map(l=>l.class_name))].length} Class${[...new Set(filtered.map(l=>l.class_name))].length!==1?"es":""}`, color:t.info },
              { label:`${[...new Set(filtered.map(l=>l.teacher_name))].length} Teacher${[...new Set(filtered.map(l=>l.teacher_name))].length!==1?"s":""}`, color:"#9B59B6" },
              { label:`${filtered.reduce((s,l)=>s+(Number(l.duration_mins)||0),0)} Total Mins`, color:t.success },
            ].map(s=>(
              <span key={s.label} style={{ fontSize:11, padding:"3px 11px", borderRadius:20,
                background:s.color+"18", color:s.color, fontFamily:"'Trebuchet MS',sans-serif",
                fontWeight:700, border:`1px solid ${s.color}33` }}>{s.label}</span>
            ))}
          </div>
        )}

        {/* Lessons list */}
        <div style={{ flex:1, overflowY:"auto", padding:"12px 20px 20px" }}>
          {filtered.length === 0 ? (
            <div style={{ textAlign:"center", padding:60, color:t.textMuted,
              fontFamily:"'Trebuchet MS',sans-serif", fontSize:14 }}>
              {lessons.length===0 ? "No lessons recorded yet." : "No lessons match your filters."}
            </div>
          ) : (
            <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
              {filtered.map((l,idx) => {
                const expanded = expandedId === l.id;
                return (
                  <div key={l.id} style={{ borderRadius:12, overflow:"hidden",
                    border:`1px solid ${expanded ? t.gold+"66" : t.border}`,
                    background: expanded ? t.gold+"06" : t.surface,
                    transition:"all 0.15s" }}>

                    {/* Row */}
                    <div style={{ display:"flex", alignItems:"center", gap:12, padding:"13px 16px",
                      cursor:"pointer" }}
                      onClick={()=>setExpandedId(expanded ? null : l.id)}>

                      {/* Index badge */}
                      <div style={{ flexShrink:0, width:28, height:28, borderRadius:"50%",
                        background:t.gold+"22", display:"flex", alignItems:"center", justifyContent:"center",
                        fontSize:11, fontWeight:800, color:t.gold, fontFamily:"'Georgia',serif" }}>
                        {idx+1}
                      </div>

                      {/* Date badge */}
                      <div style={{ flexShrink:0, textAlign:"center", width:48,
                        background:t.surfaceAlt, borderRadius:8, padding:"5px 0",
                        border:`1px solid ${t.border}` }}>
                        <div style={{ fontSize:16, fontWeight:800, color:t.gold, fontFamily:"'Georgia',serif", lineHeight:1 }}>
                          {(l.date||"--").slice(8,10)}
                        </div>
                        <div style={{ fontSize:9, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase" }}>
                          {l.date ? new Date(l.date).toLocaleString("default",{month:"short"}) : "--"}
                        </div>
                        <div style={{ fontSize:9, color:t.textFaint, fontFamily:"'Trebuchet MS',sans-serif" }}>
                          {(l.date||"").slice(0,4)}
                        </div>
                      </div>

                      {/* Main info */}
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ fontSize:14, fontWeight:700, color:t.text,
                          fontFamily:"'Trebuchet MS',sans-serif",
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                          {l.topic||"(No topic)"}
                        </div>
                        <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
                          marginTop:3, display:"flex", gap:12, flexWrap:"wrap" }}>
                          <span>🎓 {l.teacher_name||"—"}</span>
                          <span>📚 {l.class_name||"—"}</span>
                          {l.bible_references && <span>📖 {l.bible_references}</span>}
                          {l.memory_verse && <span style={{ fontStyle:"italic" }}>📜 {l.memory_verse.length>30?l.memory_verse.slice(0,30)+"…":l.memory_verse}</span>}
                          {l.duration_mins && <span>⏱ {l.duration_mins} min</span>}
                        </div>
                      </div>

                      {/* Actions */}
                      <div style={{ display:"flex", gap:5, flexShrink:0 }} onClick={e=>e.stopPropagation()}>
                        <button style={{ padding:"6px 8px", borderRadius:7, border:`1px solid ${t.warn}33`,
                          background:t.warn+"11", cursor:"pointer" }}
                          title="Edit" onClick={()=>{ startEdit(l); onClose(); }}>
                          <Icon name="edit" size={13} color={t.warn} />
                        </button>
                        <button style={{ padding:"6px 8px", borderRadius:7, border:`1px solid ${t.danger}33`,
                          background:t.danger+"11", cursor:"pointer" }}
                          title="Delete"
                          onClick={()=>{ if(window.confirm(`Delete "${l.topic}"?`)) deleteLesson(l.id); }}>
                          <Icon name="trash" size={13} color={t.danger} />
                        </button>
                      </div>

                      <Icon name={expanded?"close":"eye"} size={14} color={t.textFaint} />
                    </div>

                    {/* Expanded detail */}
                    {expanded && (
                      <div style={{ borderTop:`1px solid ${t.border}`, padding:"16px 20px",
                        display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))", gap:18 }}>

                        {l.outline && (
                          <div>
                            <div style={{ fontSize:10, fontWeight:700, color:t.gold, textTransform:"uppercase",
                              letterSpacing:1.4, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:6 }}>📋 Lesson Outline</div>
                            <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                              whiteSpace:"pre-wrap", lineHeight:1.7 }}>{l.outline}</div>
                          </div>
                        )}
                        {l.key_points && (
                          <div>
                            <div style={{ fontSize:10, fontWeight:700, color:t.info, textTransform:"uppercase",
                              letterSpacing:1.4, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:6 }}>🎯 Key Points</div>
                            <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                              whiteSpace:"pre-wrap", lineHeight:1.7 }}>{l.key_points}</div>
                          </div>
                        )}
                        {(l.memory_verse||l.assignment||l.notes) && (
                          <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
                            {l.memory_verse && (
                              <div>
                                <div style={{ fontSize:10, fontWeight:700, color:"#9B59B6", textTransform:"uppercase",
                                  letterSpacing:1.4, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>📜 Memory Verse</div>
                                <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                                  fontStyle:"italic", lineHeight:1.6 }}>"{l.memory_verse}"</div>
                              </div>
                            )}
                            {l.assignment && (
                              <div>
                                <div style={{ fontSize:10, fontWeight:700, color:ACTIVE_COLOR, textTransform:"uppercase",
                                  letterSpacing:1.4, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>📝 Assignment</div>
                                <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                                  whiteSpace:"pre-wrap", lineHeight:1.6 }}>{l.assignment}</div>
                              </div>
                            )}
                            {l.notes && (
                              <div>
                                <div style={{ fontSize:10, fontWeight:700, color:t.textMuted, textTransform:"uppercase",
                                  letterSpacing:1.4, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>🗒 Notes</div>
                                <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
                                  whiteSpace:"pre-wrap", lineHeight:1.6 }}>{l.notes}</div>
                              </div>
                            )}
                          </div>
                        )}

                        {/* Full date */}
                        <div style={{ fontSize:11, color:t.textFaint, fontFamily:"'Trebuchet MS',sans-serif",
                          gridColumn:"1 / -1", paddingTop:8, borderTop:`1px solid ${t.border}` }}>
                          Recorded on {fmtDate(l.date)}{l.duration_mins ? ` · Duration: ${l.duration_mins} mins` : ""}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

const LessonsPage = ({ db, user }) => {
  const { t, card, inp, sel, lbl, btnGold, btnOutline, btnGhost, th, td } = useThemeStyles();
  const { lessons, addLesson, updateLesson, deleteLesson } = useLessonsStore();
  const { teachers, classes } = db;
  const activeTeachers = teachers.filter(x => x.is_active === "YES");

  const blank = () => ({
    date: new Date().toISOString().slice(0,10),
    class_name: "", teacher_name: user?.name||"",
    topic: "", bible_references: "", memory_verse: "",
    outline: "", key_points: "", assignment: "",
    duration_mins: "", notes: "",
  });

  const [form, setForm]         = useState(blank);
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId]     = useState(null);
  const [search, setSearch]     = useState("");
  const [filterClass, setFilterClass] = useState("");
  const [filterTeacher, setFilterTeacher] = useState("");
  const [filterMonth, setFilterMonth]   = useState("");
  const [expandedId, setExpandedId]     = useState(null);
  const [importText, setImportText_unused]  = useState(""); // kept for compat, unused
  const [showImport, setShowImport]     = useState(false);
  const [showAllLessons, setShowAllLessons] = useState(false);

  const hc = (e) => setForm(f => ({ ...f, [e.target.name]: e.target.value }));

  const handleSubmit = () => {
    if (!form.topic.trim())   { alert("Topic is required."); return; }
    if (!form.class_name)     { alert("Class is required."); return; }
    if (!form.teacher_name)   { alert("Teacher is required."); return; }
    if (editId) { updateLesson(editId, form); setEditId(null); }
    else        { addLesson(form); }
    setForm(blank()); setShowForm(false);
  };

  const startEdit = (l) => {
    setForm({ ...blank(), ...l });
    setEditId(l.id);
    setShowForm(true);
    window.scrollTo({ top:0, behavior:"smooth" });
  };

  // ── Download Excel template ──────────────────────────────────
  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();

    // ── Instructions sheet ────────────────────────────────────
    const instr = [
      ["UCT Sunday School — Lesson Register Import Template"],
      [""],
      ["INSTRUCTIONS:"],
      ["1. Go to the 'Lessons' sheet and fill in your lesson data starting from row 3."],
      ["2. Do NOT change or delete the column headers in row 2."],
      ["3. DATE must be in YYYY-MM-DD format (e.g. 2026-05-18)."],
      ["4. DURATION_MINS should be a number only (e.g. 45)."],
      ["5. Required columns: Date, Class_Name, Teacher_Name, Topic."],
      ["6. All other columns are optional but recommended."],
      ["7. Save the file as .xlsx and click 'Import Excel File' to upload."],
      [""],
      ["COLUMN GUIDE:"],
      ["Date",           "Date of the lesson (YYYY-MM-DD)"],
      ["Class_Name",     "Name of the class (e.g. Children (5-9), Teen (10-15))"],
      ["Teacher_Name",   "Full name of the teacher"],
      ["Topic",          "Title or topic of the lesson"],
      ["Bible_References","Scripture references (e.g. Luke 10:25-37)"],
      ["Memory_Verse",   "The memory verse for the lesson"],
      ["Outline",        "Full lesson outline / structure"],
      ["Key_Points",     "Key teaching goals or points"],
      ["Assignment",     "Homework or assignment given"],
      ["Duration_Mins",  "Duration of the lesson in minutes (number only)"],
      ["Notes",          "Any additional notes"],
    ];
    const wsInstr = XLSX.utils.aoa_to_sheet(instr);
    wsInstr["!cols"] = [{ wch:22 },{ wch:60 }];
    XLSX.utils.book_append_sheet(wb, wsInstr, "Instructions");

    // ── Lessons data sheet ────────────────────────────────────
    const headers = ["Date","Class_Name","Teacher_Name","Topic","Bible_References","Memory_Verse","Outline","Key_Points","Assignment","Duration_Mins","Notes"];
    const sample = [
      ["2026-05-18","Children (5-9)","Mrs. Asante","The Good Samaritan","Luke 10:25-37","Love your neighbour as yourself","1. Opening Prayer (5 min)\n2. Bible Story (15 min)\n3. Discussion (10 min)\n4. Application (10 min)\n5. Memory Verse & Close (5 min)","• Understand neighbourly love\n• Apply kindness daily","Draw the parable scene in your notebook","45","Went well; children were engaged"],
      ["2026-05-18","Teen (10-15)","Mr. Boateng","Faith and Works","James 2:14-26","Faith without works is dead","1. Recap last week (5 min)\n2. Read passage (10 min)\n3. Group discussion (15 min)\n4. Practical challenge (10 min)\n5. Prayer (5 min)","• Faith must produce action\n• Identify one act of service this week","List 3 ways to serve someone this week","45",""],
    ];
    const wsData = XLSX.utils.aoa_to_sheet([headers, ...sample]);
    wsData["!cols"] = headers.map((h,i) => ({
      wch: [12,18,18,28,20,26,40,28,24,14,24][i]
    }));
    XLSX.utils.book_append_sheet(wb, wsData, "Lessons");

    XLSX.writeFile(wb, "UCT_LessonRegister_Template.xlsx");
  };

  // ── Import from uploaded Excel file ──────────────────────────
  const handleFileImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb    = XLSX.read(ev.target.result, { type:"binary" });
        // Try "Lessons" sheet first, fall back to first sheet
        const sheetName = wb.SheetNames.includes("Lessons") ? "Lessons" : wb.SheetNames[0];
        const rows  = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval:"" });
        if (!rows.length) { alert("No data rows found in the file."); return; }
        let count = 0, skipped = 0;
        rows.forEach(row => {
          // Normalise keys: lowercase + underscores
          const r = {};
          Object.keys(row).forEach(k => { r[k.toLowerCase().replace(/\s+/g,"_")] = String(row[k]||"").trim(); });
          const date     = r["date"]||"";
          const cls      = r["class_name"]||r["class"]||"";
          const teacher  = r["teacher_name"]||r["teacher"]||"";
          const topic    = r["topic"]||"";
          if (!topic || !cls || !teacher) { skipped++; return; }
          addLesson({
            date,
            class_name:       cls,
            teacher_name:     teacher,
            topic,
            bible_references: r["bible_references"]||r["bible_refs"]||"",
            memory_verse:     r["memory_verse"]||"",
            outline:          r["outline"]||"",
            key_points:       r["key_points"]||"",
            assignment:       r["assignment"]||"",
            duration_mins:    r["duration_mins"]||r["duration"]||"",
            notes:            r["notes"]||"",
          });
          count++;
        });
        let msg = `✅ ${count} lesson${count!==1?"s":""} imported successfully.`;
        if (skipped) msg += `\n⚠ ${skipped} row${skipped!==1?"s":""} skipped (missing Topic, Class, or Teacher).`;
        alert(msg);
        setShowImport(false);
      } catch(err) {
        alert("Could not read the file. Make sure it is a valid .xlsx file.\n\n" + err.message);
      }
    };
    reader.readAsBinaryString(file);
    // Reset input so same file can be re-imported
    e.target.value = "";
  };

  const [showExportMenu, setShowExportMenu] = useState(false);

  // ── Export lessons to Excel ───────────────────────────────────
  const exportLessons = (rows, label) => {
    if (!rows.length) { alert("No lessons to export."); return; }
    const headers = ["Date","Class_Name","Teacher_Name","Topic","Bible_References","Memory_Verse","Outline","Key_Points","Assignment","Duration_Mins","Notes","Created_At"];
    const data = rows.map(l => [
      l.date||"", l.class_name||"", l.teacher_name||"", l.topic||"",
      l.bible_references||"", l.memory_verse||"", l.outline||"",
      l.key_points||"", l.assignment||"", l.duration_mins||"",
      l.notes||"", l.createdAt ? l.createdAt.slice(0,10) : "",
    ]);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    ws["!cols"] = [12,18,18,30,22,28,40,30,26,14,26,14].map(wch=>({wch}));
    // Freeze header row
    ws["!freeze"] = { xSplit:0, ySplit:1 };
    XLSX.utils.book_append_sheet(wb, ws, "Lessons");
    // Summary sheet
    const classes  = [...new Set(rows.map(r=>r.class_name).filter(Boolean))];
    const teachers = [...new Set(rows.map(r=>r.teacher_name).filter(Boolean))];
    const totalMins = rows.reduce((s,r)=>s+(Number(r.duration_mins)||0),0);
    const summary = [
      ["UCT Lesson Register Export"],
      ["Exported on", new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"})],
      ["Filter applied", label],
      [""],
      ["Total Lessons",  rows.length],
      ["Total Classes",  classes.length],
      ["Total Teachers", teachers.length],
      ["Total Duration", totalMins ? `${totalMins} mins` : "—"],
      [""],
      ["Classes included:", classes.join(", ")||"—"],
      ["Teachers included:", teachers.join(", ")||"—"],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summary);
    wsSummary["!cols"] = [{wch:20},{wch:50}];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");
    const fileName = `UCT_Lessons_${label.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
    setShowExportMenu(false);
  };

  const allMonths = [...new Set(lessons.map(l=>(l.date||"").slice(0,7)).filter(Boolean))].sort().reverse();

  const filtered = lessons.filter(l => {
    if (filterClass   && l.class_name   !== filterClass)   return false;
    if (filterTeacher && l.teacher_name !== filterTeacher) return false;
    if (filterMonth   && !(l.date||"").startsWith(filterMonth)) return false;
    if (search) {
      const q = search.toLowerCase();
      return (l.topic||"").toLowerCase().includes(q) ||
             (l.teacher_name||"").toLowerCase().includes(q) ||
             (l.class_name||"").toLowerCase().includes(q) ||
             (l.outline||"").toLowerCase().includes(q);
    }
    return true;
  });

  const g2  = { display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 };
  const fw  = (span=1) => ({ display:"flex", flexDirection:"column", gap:5, gridColumn:span===2?"1 / -1":"auto" });
  const sec = { fontSize:11, fontWeight:700, color:t.gold, fontFamily:"'Trebuchet MS',sans-serif",
    textTransform:"uppercase", letterSpacing:1.8, padding:"16px 0 8px",
    borderTop:`1px solid ${t.border}`, marginTop:14 };

  return (
    <div>
      {/* All Lessons Modal */}
      {showAllLessons && (
        <AllLessonsModal
          lessons={lessons}
          onClose={()=>setShowAllLessons(false)}
          deleteLesson={deleteLesson}
          startEdit={startEdit}
        />
      )}

      {/* Header */}
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:12, marginBottom:20 }}>
        <div>
          <div style={{ fontSize:23, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:2 }}>
            Lesson Register
          </div>
          <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
            {lessons.length} lessons recorded · {filtered.length} shown
          </div>
        </div>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
          <button onClick={()=>setShowAllLessons(true)}
            style={{ ...btnOutline, display:"flex", alignItems:"center", gap:6,
              border:`1px solid ${t.info}66`, color:t.info, background:t.info+"11" }}>
            <Icon name="eye" size={14} color={t.info} />
            View All ({lessons.length})
          </button>

          {/* Export dropdown */}
          <div style={{ position:"relative" }}>
            {showExportMenu && (
              <div onClick={()=>setShowExportMenu(false)}
                style={{ position:"fixed", inset:0, zIndex:199 }} />
            )}
            <button onClick={()=>setShowExportMenu(v=>!v)}
              style={{ ...btnOutline, display:"flex", alignItems:"center", gap:6,
                border:`1px solid ${t.success}66`, color:t.success, background:t.success+"11" }}>
              <Icon name="export" size={14} color={t.success} />
              Export ▾
            </button>
            {showExportMenu && (
              <div style={{ position:"absolute", right:0, top:"calc(100% + 6px)", zIndex:200,
                background:t.surface, border:`1px solid ${t.border}`, borderRadius:10,
                boxShadow:`0 8px 28px rgba(0,0,0,0.18)`, minWidth:230, overflow:"hidden" }}>
                <div style={{ padding:"8px 14px 6px", fontSize:10, fontWeight:700,
                  color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
                  textTransform:"uppercase", letterSpacing:1.2, borderBottom:`1px solid ${t.border}` }}>
                  Export Lessons to Excel
                </div>
                <button onClick={()=>exportLessons(lessons,"All_Lessons")}
                  style={{ width:"100%", padding:"11px 16px", display:"flex", alignItems:"center", gap:10,
                    background:"transparent", border:"none", cursor:"pointer", textAlign:"left",
                    borderBottom:`1px solid ${t.border}` }}
                  onMouseEnter={e=>e.currentTarget.style.background=t.surfaceAlt}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <Icon name="export" size={15} color={t.success} />
                  <div>
                    <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>
                      Export All Lessons
                    </div>
                    <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
                      All {lessons.length} lessons
                    </div>
                  </div>
                </button>
                <button onClick={()=>exportLessons(filtered, filterClass||filterTeacher||filterMonth||search ? `Filtered_(${filtered.length})` : "All_Lessons")}
                  style={{ width:"100%", padding:"11px 16px", display:"flex", alignItems:"center", gap:10,
                    background:"transparent", border:"none", cursor:"pointer", textAlign:"left" }}
                  onMouseEnter={e=>e.currentTarget.style.background=t.surfaceAlt}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <Icon name="analytics" size={15} color={t.gold} />
                  <div>
                    <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif" }}>
                      Export Filtered View
                    </div>
                    <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
                      {filtered.length} lesson{filtered.length!==1?"s":""} currently shown
                      {(filterClass||filterTeacher||filterMonth||search) ? " (filters active)" : " (no filters)"}
                    </div>
                  </div>
                </button>
              </div>
            )}
          </div>

          <button onClick={()=>{ setShowImport(v=>!v); setShowForm(false); }}
            style={{ ...btnOutline, display:"flex", alignItems:"center", gap:6 }}>
            <Icon name="upload" size={14} color={t.gold} /> Bulk Import
          </button>
          <button onClick={()=>{ setShowForm(v=>!v); setEditId(null); setForm(blank()); setShowImport(false); }}
            style={{ ...btnGold, display:"flex", alignItems:"center", gap:6 }}>
            <Icon name="plus" size={14} color="#fff" /> {showForm&&!editId ? "Cancel" : "Add Lesson"}
          </button>
        </div>
      </div>

      {/* Bulk Import Panel */}
      {showImport && (
        <div style={{ ...card, marginBottom:18, border:`1px solid ${t.warn}44`, background:t.warn+"08" }}>
          <div style={{ fontSize:15, fontWeight:700, color:t.text, fontFamily:"'Georgia',serif", marginBottom:4 }}>
            📥 Bulk Import Lessons
          </div>
          <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:18, lineHeight:1.7 }}>
            Download the Excel template, fill in your lesson data, then upload it here.
          </div>

          {/* Step 1 */}
          <div style={{ display:"flex", alignItems:"flex-start", gap:14, padding:"14px 16px",
            borderRadius:10, background:t.surfaceAlt, border:`1px solid ${t.border}`, marginBottom:12 }}>
            <div style={{ width:32, height:32, borderRadius:"50%", background:t.gold+"22",
              display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0,
              fontSize:14, fontWeight:800, color:t.gold, fontFamily:"'Georgia',serif" }}>1</div>
            <div style={{ flex:1 }}>
              <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>
                Download the Template
              </div>
              <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:10, lineHeight:1.6 }}>
                Click below to get the Excel template. It includes an <strong>Instructions</strong> sheet and a <strong>Lessons</strong> sheet with sample rows and all column headers pre-filled.
              </div>
              <button onClick={downloadTemplate}
                style={{ display:"inline-flex", alignItems:"center", gap:8, padding:"9px 20px",
                  borderRadius:9, border:`1px solid ${t.success}66`, background:t.success+"18",
                  color:t.success, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13,
                  fontWeight:700, cursor:"pointer" }}>
                <Icon name="export" size={15} color={t.success} />
                ⬇ Download Excel Template
              </button>
            </div>
          </div>

          {/* Step 2 */}
          <div style={{ display:"flex", alignItems:"flex-start", gap:14, padding:"14px 16px",
            borderRadius:10, background:t.surfaceAlt, border:`1px solid ${t.border}`, marginBottom:12 }}>
            <div style={{ width:32, height:32, borderRadius:"50%", background:t.info+"22",
              display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0,
              fontSize:14, fontWeight:800, color:t.info, fontFamily:"'Georgia',serif" }}>2</div>
            <div style={{ flex:1 }}>
              <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>
                Fill In Your Lessons
              </div>
              <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", lineHeight:1.6 }}>
                Open the file, go to the <strong>Lessons</strong> sheet, and enter your data from row 3 downward.
                Do <strong>not</strong> rename or remove the header row. Required: <em>Date, Class_Name, Teacher_Name, Topic</em>.
              </div>
            </div>
          </div>

          {/* Step 3 */}
          <div style={{ display:"flex", alignItems:"flex-start", gap:14, padding:"14px 16px",
            borderRadius:10, background:t.surfaceAlt, border:`1px solid ${t.border}`, marginBottom:16 }}>
            <div style={{ width:32, height:32, borderRadius:"50%", background:ACTIVE_COLOR+"22",
              display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0,
              fontSize:14, fontWeight:800, color:ACTIVE_COLOR, fontFamily:"'Georgia',serif" }}>3</div>
            <div style={{ flex:1 }}>
              <div style={{ fontSize:13, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>
                Upload &amp; Import
              </div>
              <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:10, lineHeight:1.6 }}>
                Save your filled template and upload it below. The system will read the <strong>Lessons</strong> sheet automatically.
              </div>
              <label style={{ display:"inline-flex", alignItems:"center", gap:8, padding:"9px 20px",
                borderRadius:9, border:`1px solid ${ACTIVE_COLOR}66`, background:ACTIVE_COLOR+"18",
                color:ACTIVE_COLOR, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13,
                fontWeight:700, cursor:"pointer" }}>
                <Icon name="upload" size={15} color={ACTIVE_COLOR} />
                📂 Choose Excel File (.xlsx)
                <input type="file" accept=".xlsx,.xls" style={{ display:"none" }} onChange={handleFileImport} />
              </label>
            </div>
          </div>

          <button onClick={()=>setShowImport(false)}
            style={{ padding:"8px 20px", borderRadius:9, border:`1px solid ${t.border}`,
              background:"transparent", color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
              fontSize:13, cursor:"pointer" }}>
            Close
          </button>
        </div>
      )}

      {/* Add / Edit Form */}
      {showForm && (
        <div style={{ ...card, marginBottom:20, border:`2px solid ${t.gold}33` }}>
          <div style={{ fontSize:15, fontWeight:700, color:t.gold, fontFamily:"'Georgia',serif", marginBottom:14 }}>
            {editId ? "✏️ Edit Lesson" : "📝 New Lesson Entry"}
          </div>
          <div style={g2}>
            <div style={fw()}><label style={lbl}>Date *</label><input name="date" type="date" style={inp} value={form.date} onChange={hc} /></div>
            <div style={fw()}><label style={lbl}>Class *</label>
              <select name="class_name" style={{...sel,width:"100%"}} value={form.class_name} onChange={hc}>
                <option value="">Select class…</option>
                {classes.map(c=><option key={c.name} value={c.name}>{c.name}</option>)}
              </select>
            </div>
            <div style={fw()}><label style={lbl}>Teacher *</label>
              <select name="teacher_name" style={{...sel,width:"100%"}} value={form.teacher_name} onChange={hc}>
                <option value="">Select teacher…</option>
                {activeTeachers.map(x=><option key={x.name} value={x.name}>{x.name}</option>)}
              </select>
            </div>
            <div style={fw()}><label style={lbl}>Duration (mins)</label><input name="duration_mins" type="number" style={inp} value={form.duration_mins} onChange={hc} /></div>
          </div>

          <div style={sec}>Lesson Content</div>
          <div style={g2}>
            <div style={fw(2)}><label style={lbl}>Topic / Title *</label><input name="topic" style={inp} value={form.topic} onChange={hc} placeholder="e.g. The Good Samaritan" /></div>
            <div style={fw()}><label style={lbl}>Bible References</label><input name="bible_references" style={inp} value={form.bible_references} onChange={hc} placeholder="e.g. Luke 10:25-37" /></div>
            <div style={fw()}><label style={lbl}>Memory Verse</label><input name="memory_verse" style={inp} value={form.memory_verse} onChange={hc} /></div>
          </div>

          <div style={sec}>Lesson Outline</div>
          <div style={g2}>
            <div style={fw(2)}><label style={lbl}>Full Outline / Structure</label>
              <textarea name="outline" rows={5} style={{...inp,resize:"vertical"}} value={form.outline} onChange={hc}
                placeholder={"1. Introduction (5 mins)\n2. Bible Story (15 mins)\n3. Discussion Questions (10 mins)\n4. Application (10 mins)\n5. Memory Verse & Closing Prayer (5 mins)"} />
            </div>
            <div style={fw(2)}><label style={lbl}>Key Points / Teaching Goals</label>
              <textarea name="key_points" rows={3} style={{...inp,resize:"vertical"}} value={form.key_points} onChange={hc}
                placeholder="• Children understand the meaning of neighbourly love&#10;• Apply kindness in daily life" />
            </div>
            <div style={fw(2)}><label style={lbl}>Assignment / Homework</label>
              <textarea name="assignment" rows={2} style={{...inp,resize:"vertical"}} value={form.assignment} onChange={hc} />
            </div>
            <div style={fw(2)}><label style={lbl}>Additional Notes</label>
              <textarea name="notes" rows={2} style={{...inp,resize:"vertical"}} value={form.notes} onChange={hc} />
            </div>
          </div>

          <div style={{ marginTop:20, display:"flex", gap:12 }}>
            <button onClick={handleSubmit} style={{ ...btnGold, padding:"12px 32px", fontSize:14 }}>
              {editId ? "💾 Save Changes" : "✅ Save Lesson"}
            </button>
            <button onClick={()=>{ setShowForm(false); setEditId(null); setForm(blank()); }} style={{ ...btnOutline }}>Cancel</button>
          </div>
        </div>
      )}

      {/* Filters */}
      <div style={{ ...card, marginBottom:14, display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
        <input style={{ ...inp, minWidth:200, flex:1 }} placeholder="🔍  Search topic, teacher, class, outline…"
          value={search} onChange={e=>setSearch(e.target.value)} />
        <select style={sel} value={filterClass} onChange={e=>setFilterClass(e.target.value)}>
          <option value="">All Classes</option>
          {[...new Set(lessons.map(l=>l.class_name).filter(Boolean))].sort().map(c=><option key={c}>{c}</option>)}
        </select>
        <select style={sel} value={filterTeacher} onChange={e=>setFilterTeacher(e.target.value)}>
          <option value="">All Teachers</option>
          {[...new Set(lessons.map(l=>l.teacher_name).filter(Boolean))].sort().map(t2=><option key={t2}>{t2}</option>)}
        </select>
        <select style={sel} value={filterMonth} onChange={e=>setFilterMonth(e.target.value)}>
          <option value="">All Months</option>
          {allMonths.map(m=>{
            const [yr,mo]=m.split("-");
            return <option key={m} value={m}>{new Date(yr,mo-1).toLocaleString("default",{month:"short",year:"numeric"})}</option>;
          })}
        </select>
        {(search||filterClass||filterTeacher||filterMonth) && (
          <button onClick={()=>{ setSearch(""); setFilterClass(""); setFilterTeacher(""); setFilterMonth(""); }}
            style={{ padding:"8px 14px", borderRadius:9, border:`1px solid ${t.danger}44`,
              background:t.danger+"11", color:t.danger, fontFamily:"'Trebuchet MS',sans-serif", fontSize:13, cursor:"pointer" }}>
            ✕ Clear
          </button>
        )}
      </div>

      {/* Summary strip */}
      {filtered.length > 0 && (
        <div style={{ display:"flex", gap:10, marginBottom:14, flexWrap:"wrap" }}>
          {[
            { label:"Lessons",  value:filtered.length,                             color:t.gold },
            { label:"Classes",  value:[...new Set(filtered.map(l=>l.class_name))].length, color:t.info },
            { label:"Teachers", value:[...new Set(filtered.map(l=>l.teacher_name))].length, color:"#9B59B6" },
          ].map(s=>(
            <div key={s.label} style={{ ...card, padding:"10px 18px", flex:1, minWidth:100 }}>
              <div style={{ fontSize:11, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase", letterSpacing:1 }}>{s.label}</div>
              <div style={{ fontSize:22, fontWeight:800, color:s.color, fontFamily:"'Georgia',serif" }}>{s.value}</div>
            </div>
          ))}
        </div>
      )}

      {/* Lessons list */}
      {filtered.length === 0 ? (
        <div style={{ ...card, textAlign:"center", padding:56, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif" }}>
          {lessons.length === 0 ? "No lessons yet. Click \"Add Lesson\" to get started." : "No lessons match your filters."}
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          {filtered.map(l => {
            const expanded = expandedId === l.id;
            return (
              <div key={l.id} style={{ ...card, padding:0, overflow:"hidden",
                border:`1px solid ${expanded ? t.gold+"66" : t.border}`,
                transition:"border 0.15s" }}>
                {/* Row summary */}
                <div style={{ display:"flex", alignItems:"center", gap:14, padding:"14px 18px",
                  cursor:"pointer", background: expanded ? t.gold+"08" : "transparent" }}
                  onClick={()=>setExpandedId(expanded ? null : l.id)}>
                  {/* Date badge */}
                  <div style={{ flexShrink:0, textAlign:"center", width:52,
                    background:t.gold+"18", borderRadius:8, padding:"6px 0" }}>
                    <div style={{ fontSize:18, fontWeight:800, color:t.gold, fontFamily:"'Georgia',serif", lineHeight:1 }}>
                      {(l.date||"--").slice(8,10)}
                    </div>
                    <div style={{ fontSize:9, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", textTransform:"uppercase" }}>
                      {l.date ? new Date(l.date).toLocaleString("default",{month:"short"}) : "--"}
                    </div>
                  </div>
                  {/* Main info */}
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:15, fontWeight:700, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                      overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{l.topic||"(No topic)"}</div>
                    <div style={{ fontSize:12, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif", marginTop:3, display:"flex", gap:14, flexWrap:"wrap" }}>
                      <span>🎓 {l.teacher_name||"—"}</span>
                      <span>📚 {l.class_name||"—"}</span>
                      {l.bible_references && <span>📖 {l.bible_references}</span>}
                      {l.duration_mins && <span>⏱ {l.duration_mins} min</span>}
                    </div>
                  </div>
                  {/* Actions */}
                  <div style={{ display:"flex", gap:6, flexShrink:0 }} onClick={e=>e.stopPropagation()}>
                    <button style={btnGhost} title="Edit" onClick={()=>startEdit(l)}>
                      <Icon name="edit" size={13} color={t.warn} />
                    </button>
                    <button style={btnGhost} title="Delete"
                      onClick={()=>{ if(window.confirm(`Delete "${l.topic}"?`)) deleteLesson(l.id); }}>
                      <Icon name="trash" size={13} color={t.danger} />
                    </button>
                  </div>
                  <Icon name={expanded?"close":"eye"} size={15} color={t.textFaint} />
                </div>

                {/* Expanded detail */}
                {expanded && (
                  <div style={{ borderTop:`1px solid ${t.border}`, padding:"16px 20px",
                    display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))", gap:20 }}>
                    {l.outline && (
                      <div>
                        <div style={{ fontSize:11, fontWeight:700, color:t.gold, textTransform:"uppercase",
                          letterSpacing:1.2, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:6 }}>📋 Lesson Outline</div>
                        <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                          whiteSpace:"pre-wrap", lineHeight:1.7 }}>{l.outline}</div>
                      </div>
                    )}
                    {l.key_points && (
                      <div>
                        <div style={{ fontSize:11, fontWeight:700, color:t.info, textTransform:"uppercase",
                          letterSpacing:1.2, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:6 }}>🎯 Key Points</div>
                        <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                          whiteSpace:"pre-wrap", lineHeight:1.7 }}>{l.key_points}</div>
                      </div>
                    )}
                    {(l.memory_verse||l.assignment||l.notes) && (
                      <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                        {l.memory_verse && (
                          <div>
                            <div style={{ fontSize:11, fontWeight:700, color:"#9B59B6", textTransform:"uppercase",
                              letterSpacing:1.2, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>📜 Memory Verse</div>
                            <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                              fontStyle:"italic", lineHeight:1.6 }}>"{l.memory_verse}"</div>
                          </div>
                        )}
                        {l.assignment && (
                          <div>
                            <div style={{ fontSize:11, fontWeight:700, color:ACTIVE_COLOR, textTransform:"uppercase",
                              letterSpacing:1.2, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>📝 Assignment</div>
                            <div style={{ fontSize:13, color:t.text, fontFamily:"'Trebuchet MS',sans-serif",
                              whiteSpace:"pre-wrap", lineHeight:1.6 }}>{l.assignment}</div>
                          </div>
                        )}
                        {l.notes && (
                          <div>
                            <div style={{ fontSize:11, fontWeight:700, color:t.textMuted, textTransform:"uppercase",
                              letterSpacing:1.2, fontFamily:"'Trebuchet MS',sans-serif", marginBottom:4 }}>🗒 Notes</div>
                            <div style={{ fontSize:13, color:t.textMuted, fontFamily:"'Trebuchet MS',sans-serif",
                              whiteSpace:"pre-wrap", lineHeight:1.6 }}>{l.notes}</div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

// ─── SIDEBAR ──────────────────────────────────────────────────────────────────
const Sidebar = ({ page, setPage, user, onLogout, collapsed, setCollapsed }) => {
  const { mode, dark, toggle } = useTheme();
  const t = T[mode]||T.light;
  const perms = user?.permissions || [];
  const can = (key) => user?.role==="admin" || perms.includes(key);

  const allAdminNav = [
    {id:"dashboard",  label:"Dashboard",      icon:"dashboard",  perm:"dashboard"},
    {id:"attendance", label:"Sunday Sch. Records", icon:"attendance", perm:"attendance"},
    {id:"ssreport",   label:"Sunday Sch. Attend.", icon:"ssreport",   perm:"ssreport"},
    {id:"church",     label:"Church Attend.", icon:"cross",      perm:"church"},
    {id:"analytics",  label:"Analytics",      icon:"analytics",  perm:"analytics"},
    {id:"ai",         label:"AI Assistant",   icon:"ai",         perm:"ai"},
    {id:"lessons",    label:"Lesson Register", icon:"lesson",     perm:"lessons"},
    {id:"teachers",   label:"Teachers",       icon:"teachers",   perm:"teachers"},
    {id:"classes",    label:"Classes",        icon:"classes",    perm:"classes"},
    {id:"programs",   label:"Programs",       icon:"settings",   perm:"programs"},
    {id:"users",      label:"Users & Access", icon:"users",      perm:"__admin_only__"},
    {id:"roles",      label:"Roles",          icon:"edit",       perm:"__admin_only__"},
    {id:"branding",   label:"Logo & Name",    icon:"bible",      perm:"__admin_only__"},
    {id:"export",     label:"Export",         icon:"export",     perm:"export"},
  ];
  const allTeacherNav = [
    {id:"submit",     label:"Submit Report",  icon:"submit",     perm:"submit"},
    {id:"ssreport",   label:"Sunday Sch. Attend.", icon:"analytics",  perm:"ssreport"},
    {id:"church",     label:"Church Attend.", icon:"cross",      perm:"church"},
    {id:"attendance", label:"My Records",     icon:"attendance", perm:"attendance"},
    {id:"analytics",  label:"Analytics",      icon:"analytics",  perm:"analytics"},
    {id:"ai",         label:"AI Assistant",   icon:"ai",         perm:"ai"},
    {id:"lessons",    label:"Lesson Register", icon:"lesson",     perm:"lessons"},
    {id:"export",     label:"Export",         icon:"export",     perm:"export"},
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
    borderLeft: active ? `3px solid ${ACTIVE_COLOR}` : "3px solid transparent",
    background: active ? t.sidebarActive : "transparent",
    color: active ? "#FFFFFF" : t.sidebarMuted,
    fontSize:13, fontFamily:"'Trebuchet MS',sans-serif", transition:"all 0.15s", whiteSpace:"nowrap",
  });

  const modeLabel = mode==="light"?"☀ Light":mode==="navy"?"🌊 Navy":mode==="dark"?"🌙 Dark":mode==="darkBlue"?"🌊🌙 Dark Blue":mode==="orange"?"🔥 Orange":"🔥🌙 Dark Orange";

  return (
    <div style={{ width:collapsed?60:260, background:t.sidebar, borderRight:`1px solid ${t.sidebarBorder}`,
      display:"flex", flexDirection:"column", flexShrink:0, transition:"width 0.22s", overflow:"hidden", minHeight:"100vh" }}>
      {/* Header with animated bible */}
      <div style={{ padding:collapsed?"18px 8px":"18px 18px 14px", borderBottom:`1px solid ${t.sidebarBorder}`,
        display:"flex", flexDirection:"column", alignItems:"center", gap:6 }}>
        {!collapsed && <AnimatedBible size={56} />}
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          {collapsed && <ChurchLogo size={32} />}
          {!collapsed && (
            <div style={{ textAlign:"center" }}>
              <div style={{ fontSize:9.5, fontWeight:700, color:"#FFFFFF", fontFamily:"'Trebuchet MS',sans-serif", letterSpacing:1.2, lineHeight:1.5 }}>{CHURCH_NAME.toUpperCase()}</div>
              <div style={{ fontSize:8, color:"rgba(255,255,255,0.5)", fontFamily:"'Trebuchet MS',sans-serif", letterSpacing:0.5 }}>SSM SYSTEM</div>
            </div>
          )}
        </div>
      </div>
      <nav style={{ flex:1, paddingTop:8 }}>
        {nav.map(n => (
          <div key={n.id} style={navItem(page===n.id)} onClick={()=>setPage(n.id)} title={collapsed?n.label:""}
            onMouseEnter={e=>{ if(page!==n.id) e.currentTarget.style.background="rgba(255,255,255,0.08)"; }}
            onMouseLeave={e=>{ if(page!==n.id) e.currentTarget.style.background="transparent"; }}>
            <Icon name={n.icon} size={17} color={page===n.id ? ACTIVE_COLOR : t.sidebarMuted} />
            {!collapsed && <span style={{ color: page===n.id ? "#FFFFFF" : t.sidebarMuted }}>{n.label}</span>}
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
        <div style={{ display:"flex", gap:6, justifyContent:collapsed?"center":"flex-start", flexWrap:"wrap" }}>
          <button title={`Theme: ${modeLabel}`} onClick={toggle}
            style={{ padding:"6px 8px", borderRadius:7, border:`1px solid ${ACTIVE_COLOR}55`, background:`${ACTIVE_COLOR}18`, cursor:"pointer",
              display:"flex", alignItems:"center", gap:4, fontSize:10, color:"rgba(255,255,255,0.85)",
              fontFamily:"'Trebuchet MS',sans-serif" }}>
            <Icon name={isDark(mode)?"sun":"moon"} size={13} color={ACTIVE_COLOR} />
            {!collapsed && modeLabel}
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
  const { mode, dark, toggle } = useTheme();
  const t = T[mode]||T.light;
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
  const { mode } = useTheme();
  const t = T[mode]||T.light;
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
      background:t.sidebar, borderTop:`1px solid ${t.sidebarBorder}`,
      display:"flex", alignItems:"center",
      paddingBottom:"env(safe-area-inset-bottom,0px)",
      boxShadow:`0 -3px 16px ${t.sidebar}88` }}>
      {quickTabs.map(tab => {
        const active = page === tab.id;
        return (
          <button key={tab.id} onClick={()=>setPage(tab.id)} style={{
            flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
            padding:"10px 2px 8px", border:"none", background:"transparent", cursor:"pointer",
            borderTop: active ? `2px solid ${ACTIVE_COLOR}` : "2px solid transparent",
          }}>
            <Icon name={tab.icon} size={21} color={active ? ACTIVE_COLOR : "rgba(255,255,255,0.45)"} />
            <span style={{ fontSize:9, color: active ? ACTIVE_COLOR : "rgba(255,255,255,0.45)",
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
  const { mode, dark, toggle } = useTheme();
  const t = T[mode]||T.light;
  const perms   = user?.permissions || [];
  const isAdmin = user?.role === "admin";
  const can = (key) => isAdmin || perms.includes(key);

  const allNav = isAdmin ? [
    {id:"dashboard",  label:"Dashboard",           icon:"dashboard"},
    {id:"attendance", label:"Sunday Sch. Records", icon:"attendance"},
    {id:"ssreport",   label:"Sunday Sch. Attend.", icon:"ssreport"},
    {id:"church",     label:"Church Attend.",      icon:"cross"},
    {id:"analytics",  label:"Analytics",           icon:"analytics"},
    {id:"ai",         label:"AI Assistant",        icon:"ai"},
    {id:"lessons",    label:"Lesson Register",     icon:"lesson"},
    {id:"teachers",   label:"Teachers",            icon:"teachers"},
    {id:"classes",    label:"Classes",             icon:"classes"},
    {id:"programs",   label:"Programs",            icon:"settings"},
    {id:"users",      label:"Users & Access",      icon:"users"},
    {id:"roles",      label:"Roles",               icon:"edit"},
    {id:"branding",   label:"Logo & Name",         icon:"bible"},
    {id:"export",     label:"Export",              icon:"export"},
  ] : [
    {id:"submit",     label:"Submit Report",       icon:"submit",    perm:"submit"},
    {id:"ssreport",   label:"Sunday Sch. Attend.", icon:"ssreport",  perm:"ssreport"},
    {id:"church",     label:"Church Attend.",      icon:"cross",     perm:"church"},
    {id:"attendance", label:"My Records",          icon:"attendance",perm:"attendance"},
    {id:"analytics",  label:"Analytics",           icon:"analytics", perm:"analytics"},
    {id:"ai",         label:"AI Assistant",        icon:"ai",        perm:"ai"},
    {id:"lessons",    label:"Lesson Register",     icon:"lesson",    perm:"lessons"},
    {id:"export",     label:"Export",              icon:"export",    perm:"export"},
  ].filter(n => can(n.perm||n.id));

  if (!open) return null;
  return (
    <>
      <div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", zIndex:300 }} />
      <div style={{ position:"fixed", left:0, top:0, bottom:0, width:280, zIndex:400,
        background:t.sidebar, display:"flex", flexDirection:"column",
        boxShadow:"6px 0 32px rgba(0,0,0,0.4)" }}>
        {/* Header with animated bible */}
        <div style={{ padding:"16px 18px 12px", borderBottom:`1px solid ${t.sidebarBorder}`,
          display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <AnimatedBible size={42} />
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
        <div style={{ padding:"13px 18px", borderBottom:`1px solid ${t.sidebarBorder}`, display:"flex", alignItems:"center", gap:10 }}>
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
                  background: active ? t.sidebarActive : "transparent",
                  borderLeft: active ? `3px solid ${ACTIVE_COLOR}` : "3px solid transparent",
                  color: active ? "#FFFFFF" : t.sidebarMuted,
                  fontFamily:"'Trebuchet MS',sans-serif", fontSize:13 }}>
                <Icon name={n.icon} size={18} color={active ? ACTIVE_COLOR : "rgba(255,255,255,0.5)"} />
                {n.label}
              </div>
            );
          })}
        </nav>
        {/* Footer */}
        <div style={{ padding:"13px 18px", borderTop:`1px solid ${t.sidebarBorder}`, display:"flex", gap:8 }}>
          <button onClick={toggle} style={{ flex:1, padding:"9px 0", borderRadius:8, border:`1px solid ${ACTIVE_COLOR}55`, background:`${ACTIVE_COLOR}18`, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:6, color:"rgba(255,255,255,0.9)", fontFamily:"'Trebuchet MS',sans-serif", fontSize:11 }}>
            <Icon name={isDark(mode)?"sun":"moon"} size={13} color={ACTIVE_COLOR} />
            {mode==="light"?"☀ Light":mode==="navy"?"🌊 Navy":mode==="dark"?"🌙 Dark":mode==="darkBlue"?"🌊🌙 Dark Blue":mode==="orange"?"🔥 Orange":"🔥🌙 Dark Orange"}
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
  const MODES = ["light","navy","dark","darkBlue","orange","darkOrange"];
  const [modeIdx, setModeIdx]         = useState(0);
  const mode = MODES[modeIdx];
  const dark = isDark(mode);
  const cycleTheme = () => setModeIdx(i => (i+1) % MODES.length);
  const [user, setUser]             = useState(null);
  const [page, setPage]             = useState("dashboard");
  const [collapsed, setCollapsed]   = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [toasts, setToasts]         = useState([]);
  const [editingRecord, setEditingRecord] = useState(null); // record being edited
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
    const t = T[mode]||T.light;
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
    attendance:guard("attendance", <AttendancePage db={db} user={user}
      onEditRecord={r=>{ setEditingRecord(r); setPage("submit"); }} />),
    church:    guard("church",     <ChurchAttendancePage db={db} user={user} />),
    analytics: guard("analytics",  <AnalyticsPage db={db} />),
    teachers:  guard("teachers",   <TeachersPage db={db} />),
    classes:   guard("classes",    <ClassesPage db={db} />),
    programs:  guard("programs",   <ProgramsPage db={db} />),
    users:     user?.role==="admin" ? <UsersPage users={users} /> : <AccessDenied />,
    roles:     user?.role==="admin" ? <RolesPage /> : <AccessDenied />,
    branding:  user?.role==="admin" ? <BrandingPage /> : <AccessDenied />,
    export:    guard("export",     <ExportPage db={db} />),
    submit:    guard("submit",     <SubmitPage db={db} user={user}
      onSuccess={()=>toast(editingRecord?"Record updated!":"Report saved!","success")}
      editRecord={editingRecord}
      onCancelEdit={()=>{ setEditingRecord(null); setPage("attendance"); }} />),
    ssreport:  guard("ssreport",   <SSReportPage db={db} />),
    lessons:   guard("lessons",    <LessonsPage db={db} user={user} />),
    ai:        guard("ai",         <AIAssistantPage db={db} user={user} />),
  };

  const t = T[mode]||T.light;

  if (!user) return (
    <ThemeCtx.Provider value={{ mode, dark, toggle: cycleTheme }}>
      <LoginPage onLogin={handleLogin} />
    </ThemeCtx.Provider>
  );

  // Show loading screen while Supabase fetches data
  if (db.loading) {
    const t = T[mode]||T.light;
    return (
      <ThemeCtx.Provider value={{ mode, dark, toggle: cycleTheme }}>
        <div style={{ minHeight:"100vh", background:t.bg, display:"flex", flexDirection:"column",
          alignItems:"center", justifyContent:"center", gap:20, fontFamily:"'Trebuchet MS',sans-serif" }}>
          <AnimatedBible size={72} />
          <div style={{ fontSize:16, color:t.gold, fontWeight:700 }}>Loading data…</div>
          <div style={{ fontSize:13, color:t.textMuted }}>Connecting to Supabase database</div>
          <div style={{ width:200, height:4, background:t.border, borderRadius:4, overflow:"hidden" }}>
            <div style={{ height:"100%", background:t.topbar, borderRadius:4,
              animation:"loadbar 1.4s ease-in-out infinite",
              width:"60%" }} />
          </div>
          <style>{`@keyframes loadbar{0%{transform:translateX(-100%)}100%{transform:translateX(240%)}}`}</style>
        </div>
      </ThemeCtx.Provider>
    );
  }

  return (
    <ThemeCtx.Provider value={{ mode, dark, toggle: cycleTheme }}>
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

          {/* Topbar — themed band */}
          <div style={{ background: t.topbar, padding: mobile ? "11px 14px" : "11px 24px",
            display:"flex", alignItems:"center", justifyContent:"space-between",
            flexShrink:0, position:"sticky", top:0, zIndex:100,
            boxShadow:`0 2px 12px ${t.topbar}55` }}>

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
              {/* 3-mode theme toggle */}
              <button onClick={cycleTheme} title={`Theme: ${mode}`}
                style={{ padding:"6px 10px", borderRadius:7, border:`1px solid ${ACTIVE_COLOR}77`,
                  background:`${ACTIVE_COLOR}22`, cursor:"pointer", display:"flex", alignItems:"center", gap:5,
                  color:"rgba(255,255,255,0.9)", fontFamily:"'Trebuchet MS',sans-serif", fontSize:11 }}>
                <Icon name={isDark(mode)?"sun":"moon"} size={13} color={ACTIVE_COLOR} />
                {!mobile && (mode==="light"?"☀ Light":mode==="navy"?"🌊 Navy":mode==="dark"?"🌙 Dark":mode==="darkBlue"?"🌊🌙 Dark Blue":mode==="orange"?"🔥 Orange":"🔥🌙 Dark Orange")}
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
