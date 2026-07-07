import { useState, useEffect, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { Building2, Users, Clock, Calendar, Plus, Trash2, Check, ChevronDown, ChevronRight, Download, Play, Settings, History, ArrowLeft, Copy, Pencil, X, Save, CircleCheck, Timer, FileSpreadsheet, FileText, Mail } from "lucide-react";

// ---------- constants ----------
// Functional colour-coding for crews/task groups. Kept distinct from the
// Codee brand red/orange so multiple crews stay easy to tell apart on sight,
// the way CC1-CC4 are colour-coded on the site poster.
const PALETTE = [
  { name: "Signal Blue", hex: "#2F6FED" },
  { name: "Field Green", hex: "#16A34A" },
  { name: "Codee Red", hex: "#EE3A23" },
  { name: "Deep Violet", hex: "#7C5CFC" },
  { name: "Codee Orange", hex: "#F47629" },
  { name: "Slate", hex: "#64748B" },
];

const BRAND_RED = "#EE3A23";
const BRAND_ORANGE = "#F47629";
const BRAND_GRADIENT = `linear-gradient(135deg, ${BRAND_RED}, ${BRAND_ORANGE})`;
const CODEE_LOGO_SRC = "data:image/png;base64,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";

const uid = (p = "id") => `${p}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;

// Built-in starter crew templates, taken from the original Hancock Iron Ore
// scope. These are always available as one-click presets in Manage Scope,
// in addition to any groups the user has already created on other sites.
const PRESET_GROUPS_TEMPLATE = [
  {
    code: "PRE", label: "Prestart", color: "#64748B", timeframe: "5:30PM \u2013 5:40PM (10 min)",
    tasks: ["Scope confirmation", "Equipment check", "Chemical check", "Consumables check", "Role allocation"],
  },
  {
    code: "CC1", label: "Bathrooms & Consumables Lead", color: "#2F6FED", timeframe: "All shift",
    tasks: ["Toilets, basins, mirrors", "Soap, paper, paper towel", "Bathroom presentation", "Stock levels"],
  },
  {
    code: "CC2", label: "Floors & Detail Lead \u2014 Levels 5 to Ground", color: "#16A34A", timeframe: "5:40PM \u2013 7:10PM",
    tasks: ["Vacuum all floor areas", "Desk wipe downs", "Meeting rooms", "Spot clean glass", "Lounge areas", "Presentation checks"],
  },
  {
    code: "CC3", label: "Kitchens & Bathrooms Lead \u2014 All Levels", color: BRAND_RED, timeframe: "Kitchens 5:30\u20137:10PM \u00b7 Bathrooms 7:10\u20138:50PM",
    tasks: [
      "Pack & unpack dishwasher", "Wipe microwave inside/out", "Wipe benches & sinks", "Spot mop floors",
      "Wipe cupboards & drawers", "Replenish consumables (kitchen)", "Empty bins (kitchen)", "Wipe tables & chairs", "Coffee machine checks",
      "Replenish consumables (bathroom)", "Clean toilet bowls", "Wipe all surfaces", "Check pamper drawers", "Clean sinks", "Empty bins (bathroom)", "End-of-task bathroom inspection",
    ],
  },
  {
    code: "CC4", label: "Floors, Bins & Common Areas Lead", color: "#7C5CFC", timeframe: "5:30PM \u2013 9:30PM",
    tasks: ["Machine scrub all hard floors, Ground \u2192 Level 5", "Empty all bins", "Cafe table wipe downs", "Flat mop floors", "Cleaner room reset", "End of shift presentation"],
  },
  {
    code: "FIN", label: "Final Walkthrough", color: BRAND_ORANGE, timeframe: "9:10PM \u2013 9:30PM (30 min)",
    tasks: ["Building presentation check", "Toilets complete", "Kitchens complete", "Floors complete", "Bins removed", "Equipment packed away"],
  },
];
function instantiatePresetGroup(tpl) {
  return {
    id: uid("grp"), code: tpl.code, label: tpl.label, color: tpl.color, timeframe: tpl.timeframe || "",
    tasks: tpl.tasks.map(t => ({ id: uid("task"), text: t })),
  };
}

function seedSites() {
  const hancock = {
    id: uid("site"),
    name: "Hancock Iron Ore",
    building: "Hancock Building",
    levels: "Ground + Levels 1\u20135",
    crewSize: 4,
    hours: "5:00 PM \u2013 8:30 PM",
    days: "Monday \u2013 Friday",
    image: null,
    createdAt: Date.now(),
    groups: PRESET_GROUPS_TEMPLATE.map(instantiatePresetGroup),
  };
  return [hancock];
}

// ---------- storage helpers ----------
async function loadJSON(key, fallback) {
  try {
    const res = await window.storage.get(key, false);
    return res ? JSON.parse(res.value) : fallback;
  } catch {
    return fallback;
  }
}
async function saveJSON(key, value) {
  try {
    await window.storage.set(key, JSON.stringify(value), false);
  } catch (e) {
    console.error("storage save failed", key, e);
  }
}

function totalTasks(site) {
  return site.groups.reduce((n, g) => n + g.tasks.length, 0);
}
function fmtDuration(ms) {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  return h > 0 ? `${h}h ${m}m` : `${m}m`;
}
function fmtClock(ts) { return new Date(ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }); }
function fmtDate(ts) { return new Date(ts).toLocaleDateString([], { year: "numeric", month: "short", day: "numeric" }); }

// ---------- phyllotaxis dot mark (echoes the Codee logo's packed dot circle) ----------
function useDotLayout(count, radius) {
  return useMemo(() => {
    const pts = [];
    const golden = 137.50776405003785 * (Math.PI / 180);
    for (let i = 0; i < count; i++) {
      const r = radius * Math.sqrt((i + 0.5) / count);
      const theta = i * golden;
      const x = r * Math.cos(theta);
      const y = r * Math.sin(theta);
      const size = 2.2 + 3.4 * ((i + 0.5) / count);
      pts.push({ x, y, size, order: i });
    }
    return pts;
  }, [count, radius]);
}

function DotMark({ size = 34, lit = null }) {
  // lit: null = fully coloured brand mark (logo use). 0..1 = progress fill.
  // Renders at 100% of its container so it scales responsively; `size` only
  // sets the internal coordinate system / dot density reference.
  const count = 34;
  const R = size / 2 - 2;
  const pts = useDotLayout(count, R);
  const litCount = lit === null ? count : Math.round(lit * count);
  return (
    <svg width="100%" height="100%" viewBox={`${-size / 2} ${-size / 2} ${size} ${size}`} style={{ display: "block" }}>
      {pts.map((p, i) => {
        const t = i / (count - 1);
        const on = i < litCount;
        const color = t < 0.5 ? BRAND_RED : BRAND_ORANGE;
        return <circle key={i} cx={p.x} cy={p.y} r={p.size / 2} fill={on ? color : "#E5E1DD"} opacity={on ? (0.55 + 0.45 * t) : 1} />;
      })}
    </svg>
  );
}

function CodeeWordmark({ height = 22 }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: FONT_DISPLAY, fontWeight: 800, color: "#5B5D63", fontSize: height, letterSpacing: 0.5, lineHeight: 1 }}>
      <span>C</span>
      <div style={{ width: height * 1.35, height: height * 1.35 }}><DotMark size={height * 1.35} /></div>
      <span>DEE</span>
    </div>
  );
}

// ---------- ring gauge + segmented crew bar (the new progress design) ----------
function RingGauge({ pct, gradId }) {
  const V = 100, stroke = 11, r = (V - stroke) / 2, c = 2 * Math.PI * r;
  const offset = c * (1 - Math.max(0, Math.min(1, pct)));
  return (
    <svg width="100%" height="100%" viewBox={`0 0 ${V} ${V}`} style={{ display: "block" }}>
      <defs>
        <linearGradient id={gradId} x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor={BRAND_RED} />
          <stop offset="100%" stopColor={BRAND_ORANGE} />
        </linearGradient>
      </defs>
      <circle cx={V / 2} cy={V / 2} r={r} fill="none" stroke="#F1EFEB" strokeWidth={stroke} />
      <circle cx={V / 2} cy={V / 2} r={r} fill="none" stroke={`url(#${gradId})`} strokeWidth={stroke}
        strokeDasharray={c} strokeDashoffset={offset} strokeLinecap="round"
        transform={`rotate(-90 ${V / 2} ${V / 2})`} style={{ transition: "stroke-dashoffset .5s ease" }} />
      <text x="50%" y="52%" textAnchor="middle" dominantBaseline="central" style={{ fontFamily: FONT_DISPLAY, fontWeight: 700, fontSize: 24 }} fill="#2B2C30">
        {Math.round(pct * 100)}%
      </text>
    </svg>
  );
}

function SegmentedCrewBar({ groups }) {
  const totalTasks = groups.reduce((n, g) => n + g.tasks.length, 0) || 1;
  return (
    <div>
      <div style={styles.segBarTrack}>
        {groups.filter(g => g.tasks.length > 0).map(g => {
          const gTotal = g.tasks.length;
          const gDone = g.tasks.filter(t => t.done).length;
          const widthPct = (gTotal / totalTasks) * 100;
          const fillPct = gTotal ? (gDone / gTotal) * 100 : 0;
          return (
            <div key={g.id} style={{ ...styles.segBarSeg, width: `${widthPct}%` }} title={`${g.code}: ${gDone}/${gTotal}`}>
              <div style={{ ...styles.segBarFill, width: `${fillPct}%`, background: g.color }} />
            </div>
          );
        })}
      </div>
      <div style={styles.segBarLegend}>
        {groups.filter(g => g.tasks.length > 0).map(g => {
          const gDone = g.tasks.filter(t => t.done).length;
          return (
            <div key={g.id} style={styles.segBarLegendItem}>
              <span style={{ ...styles.segBarDot, background: g.color }} />
              {g.code} {gDone}/{g.tasks.length}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ---------- image helper: downscale a photo before storing it ----------
function resizeImageFile(file, maxDim = 720, quality = 0.75) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not read file"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("Could not read image"));
      img.onload = () => {
        let { width, height } = img;
        if (width > height && width > maxDim) { height = Math.round(height * (maxDim / width)); width = maxDim; }
        else if (height >= width && height > maxDim) { width = Math.round(width * (maxDim / height)); height = maxDim; }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

// ============================================================
export default function App() {
  const [ready, setReady] = useState(false);
  const [sites, setSites] = useState([]);
  const [jobHistory, setJobHistory] = useState([]);
  const [activeJob, setActiveJob] = useState(null);
  const [view, setView] = useState("dashboard");
  const [editingSiteId, setEditingSiteId] = useState(null);
  const [completedModalJob, setCompletedModalJob] = useState(null);

  useEffect(() => {
    (async () => {
      const s = await loadJSON("sites", null);
      const initialSites = s && s.length ? s : seedSites();
      if (!s) await saveJSON("sites", initialSites);
      setSites(initialSites);
      setJobHistory(await loadJSON("job-history", []));
      const aj = await loadJSON("active-job", null);
      if (aj) { setActiveJob(aj); setView("shift"); }
      setReady(true);
    })();
  }, []);

  const persistSites = useCallback((next) => { setSites(next); saveJSON("sites", next); }, []);

  const startShift = (site) => {
    const job = {
      id: uid("job"), siteId: site.id, siteName: site.name, siteImage: site.image || null, startedAt: Date.now(), completedAt: null,
      groups: site.groups.map(g => ({
        id: g.id, code: g.code, label: g.label, color: g.color, timeframe: g.timeframe,
        tasks: g.tasks.map(t => ({ id: t.id, text: t.text, done: false, doneAt: null })),
      })),
    };
    setActiveJob(job); saveJSON("active-job", job); setView("shift");
  };
  const updateActiveJob = (next) => { setActiveJob(next); saveJSON("active-job", next); };

  const completeShift = async (job) => {
    const finished = { ...job, completedAt: Date.now() };
    const total = finished.groups.reduce((n, g) => n + g.tasks.length, 0);
    const done = finished.groups.reduce((n, g) => n + g.tasks.filter(t => t.done).length, 0);
    const summary = { id: finished.id, siteId: finished.siteId, siteName: finished.siteName, startedAt: finished.startedAt, completedAt: finished.completedAt, total, done };
    const nextHistory = [summary, ...jobHistory];
    setJobHistory(nextHistory);
    await saveJSON("job-history", nextHistory);
    await saveJSON(`job:${finished.id}`, finished);
    await window.storage.delete("active-job", false).catch(() => {});
    setActiveJob(null);
    setCompletedModalJob(finished);
  };
  const abandonShift = async () => {
    await window.storage.delete("active-job", false).catch(() => {});
    setActiveJob(null); setView("dashboard");
  };

  if (!ready) {
    return (
      <div style={styles.loadingScreen}>
        <style>{globalCss}</style>
        <img src={CODEE_LOGO_SRC} alt="Codee Cleaning Services" style={{ height: 42, width: "auto" }} />
        <div style={{ color: "#8A8D93", fontFamily: FONT_BODY, fontSize: 13, letterSpacing: 1, fontWeight: 600 }}>LOADING SITE OPERATING SYSTEM…</div>
      </div>
    );
  }

  return (
    <div style={styles.app}>
      <style>{globalCss}</style>
      <TopBar view={view} setView={setView} hasActiveJob={!!activeJob} />
      <div style={styles.body}>
        {view === "dashboard" && (
          <Dashboard
            sites={sites}
            onStart={startShift}
            onEdit={(id) => { setEditingSiteId(id); setView("manage"); }}
            onNewSite={() => { setEditingSiteId(null); setView("manage"); }}
            onDeleteSite={(id) => persistSites(sites.filter(s => s.id !== id))}
            onDuplicateSite={(id) => {
              const src = sites.find(s => s.id === id);
              const clone = JSON.parse(JSON.stringify(src));
              clone.id = uid("site"); clone.name = src.name + " (copy)";
              clone.groups.forEach(g => { g.id = uid("grp"); g.tasks.forEach(t => t.id = uid("task")); });
              persistSites([...sites, clone]);
            }}
          />
        )}
        {view === "manage" && (
          <SiteEditor
            site={editingSiteId ? sites.find(s => s.id === editingSiteId) : null}
            sites={sites}
            onCancel={() => setView("dashboard")}
            onSave={(site) => {
              const exists = sites.some(s => s.id === site.id);
              const next = exists ? sites.map(s => s.id === site.id ? site : s) : [...sites, site];
              persistSites(next); setView("dashboard");
            }}
          />
        )}
        {view === "shift" && activeJob && (
          <ShiftRunner
            job={activeJob}
            onToggleTask={(groupId, taskId) => {
              const next = {
                ...activeJob,
                groups: activeJob.groups.map(g => g.id !== groupId ? g : {
                  ...g, tasks: g.tasks.map(t => t.id !== taskId ? t : { ...t, done: !t.done, doneAt: !t.done ? Date.now() : null }),
                }),
              };
              updateActiveJob(next);
            }}
            onComplete={() => completeShift(activeJob)}
            onAbandon={abandonShift}
          />
        )}
        {view === "history" && (
          <HistoryView history={jobHistory} onLoadJob={(id) => loadJSON(`job:${id}`, null)} />
        )}
      </div>

      {completedModalJob && (
        <CompletionModal job={completedModalJob} onClose={() => { setCompletedModalJob(null); setView("dashboard"); }} />
      )}
    </div>
  );
}

// ---------- shared export helpers ----------
function reportBaseName(job) {
  const safeName = job.siteName.replace(/[^a-z0-9]+/gi, "_");
  return `${safeName}_${fmtDate(job.startedAt).replace(/[, ]+/g, "_")}`;
}
function jobStats(job) {
  const total = job.groups.reduce((n, g) => n + g.tasks.length, 0);
  const done = job.groups.reduce((n, g) => n + g.tasks.filter(t => t.done).length, 0);
  return { total, done, pct: total ? Math.round((done / total) * 100) : 0 };
}
function triggerBlobDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

// ---------- Excel ----------
function buildWorkbook(job) {
  const { total, done, pct } = jobStats(job);
  const durationMs = (job.completedAt || Date.now()) - job.startedAt;

  const summaryRows = [
    ["Codee Cleaning Services \u2014 Shift Report"], [],
    ["Site", job.siteName], ["Date", fmtDate(job.startedAt)],
    ["Shift Start", fmtClock(job.startedAt)], ["Shift Completed", job.completedAt ? fmtClock(job.completedAt) : "In progress"],
    ["Duration", fmtDuration(durationMs)], ["Total Tasks", total], ["Tasks Completed", done],
    ["Completion %", pct + "%"],
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  wsSummary["!cols"] = [{ wch: 18 }, { wch: 30 }];

  const detailHeader = ["Group Code", "Group / Crew", "Time Frame", "Task", "Completed", "Completed At"];
  const detailRows = [detailHeader];
  job.groups.forEach(g => {
    g.tasks.forEach(t => {
      detailRows.push([g.code, g.label, g.timeframe || "", t.text, t.done ? "Yes" : "No", t.doneAt ? new Date(t.doneAt).toLocaleString() : ""]);
    });
  });
  const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
  wsDetail["!cols"] = [{ wch: 10 }, { wch: 34 }, { wch: 26 }, { wch: 40 }, { wch: 11 }, { wch: 20 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");
  XLSX.utils.book_append_sheet(wb, wsDetail, "Task Detail");
  return wb;
}
function excelBlob(job) {
  const wb = buildWorkbook(job);
  const arr = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([arr], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
function downloadExcel(job) {
  triggerBlobDownload(excelBlob(job), `${reportBaseName(job)}.xlsx`);
}

// ---------- dependency-free PDF generator ----------
// Builds a minimal but valid multi-page PDF (Helvetica text only) so the
// report can be exported and shared without any external PDF library.
function sanitizePdfText(s) {
  return String(s)
    .replace(/[\u2010-\u2015]/g, "-")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/\u2192/g, "->")
    .replace(/\u00b7/g, "*")
    .replace(/\u2026/g, "...")
    .replace(/[^\x00-\x7E]/g, "");
}
function escapePdfText(s) {
  return s.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}
function hexToRgb01(hex) {
  const h = hex.replace("#", "");
  const r = parseInt(h.substring(0, 2), 16) / 255;
  const g = parseInt(h.substring(2, 4), 16) / 255;
  const b = parseInt(h.substring(4, 6), 16) / 255;
  return [r, g, b];
}
function buildPdfBytes(job) {
  const { total, done, pct } = jobStats(job);
  const duration = fmtDuration((job.completedAt || Date.now()) - job.startedAt);
  const BRAND = hexToRgb01(BRAND_RED);
  const GREY = [0.42, 0.43, 0.46];
  const DARK = [0.14, 0.15, 0.17];

  const lines = [];
  lines.push({ text: "CODEE CLEANING SERVICES", size: 18, bold: true, color: DARK });
  lines.push({ text: "Site Operating System \u2014 Shift Report", size: 11, color: BRAND, bold: true });
  lines.push({ text: "", size: 8 });
  lines.push({ text: `Site: ${job.siteName}`, size: 11, color: DARK });
  lines.push({ text: `Date: ${fmtDate(job.startedAt)}`, size: 11, color: DARK });
  lines.push({ text: `Shift: ${fmtClock(job.startedAt)} - ${job.completedAt ? fmtClock(job.completedAt) : "In progress"}  (${duration})`, size: 11, color: DARK });
  lines.push({ text: `Completion: ${done}/${total} tasks (${pct}%)`, size: 12, bold: true, color: BRAND });
  lines.push({ text: "", size: 10 });

  job.groups.forEach(g => {
    lines.push({ text: `${g.code}   ${g.label}`, size: 12.5, bold: true, color: hexToRgb01(g.color || BRAND_RED) });
    if (g.timeframe) lines.push({ text: g.timeframe, size: 9, color: GREY });
    g.tasks.forEach(t => {
      const mark = t.done ? "[x]" : "[ ]";
      const timeStr = t.done && t.doneAt ? `   (done ${fmtClock(t.doneAt)})` : "";
      lines.push({ text: `   ${mark}  ${t.text}${timeStr}`, size: 10, color: t.done ? GREY : DARK });
    });
    lines.push({ text: "", size: 6 });
  });

  // paginate
  const PAGE_W = 612, PAGE_H = 792, MARGIN = 50, FOOTER_Y = 34;
  const pages = [];
  let ops = [];
  let y = PAGE_H - MARGIN;
  const flushPage = () => { pages.push(ops); ops = []; y = PAGE_H - MARGIN; };

  lines.forEach(line => {
    const size = line.size || 10;
    const leading = size * 1.42;
    if (y - leading < MARGIN + FOOTER_Y) flushPage();
    if (line.text !== "") {
      const font = line.bold ? "F2" : "F1";
      const [r, g, b] = line.color || DARK;
      const txt = escapePdfText(sanitizePdfText(line.text));
      ops.push(`${r.toFixed(3)} ${g.toFixed(3)} ${b.toFixed(3)} rg`);
      ops.push("BT", `/${font} ${size} Tf`, `${MARGIN} ${y} Td`, `(${txt}) Tj`, "ET");
    }
    y -= leading;
  });
  flushPage();

  // footer on every page
  const totalPages = pages.length;
  pages.forEach((pageOps, i) => {
    pageOps.push(`${GREY[0]} ${GREY[1]} ${GREY[2]} rg`, "BT", "/F1 8 Tf", `${MARGIN} ${FOOTER_Y - 12} Td`,
      `(${escapePdfText(`Codee Cleaning Services  \u00b7  Page ${i + 1} of ${totalPages}`)}) Tj`, "ET");
  });

  // assemble PDF object graph
  const numPages = pages.length;
  const pageIds = [], contentIds = [];
  let nextId = 5;
  pages.forEach(() => { pageIds.push(nextId++); contentIds.push(nextId++); });
  const maxId = nextId - 1;

  const objs = {};
  objs[1] = `1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n`;
  objs[2] = `2 0 obj\n<< /Type /Pages /Kids [${pageIds.map(id => `${id} 0 R`).join(" ")}] /Count ${numPages} >>\nendobj\n`;
  objs[3] = `3 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n`;
  objs[4] = `4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>\nendobj\n`;
  pages.forEach((pageOps, i) => {
    const pid = pageIds[i], cid = contentIds[i];
    const stream = pageOps.join("\n");
    objs[pid] = `${pid} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${cid} 0 R >>\nendobj\n`;
    objs[cid] = `${cid} 0 obj\n<< /Length ${stream.length} >>\nstream\n${stream}\nendstream\nendobj\n`;
  });

  let pdf = "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n";
  const offsets = {};
  for (let id = 1; id <= maxId; id++) { offsets[id] = pdf.length; pdf += objs[id]; }
  const xrefStart = pdf.length;
  let xref = `xref\n0 ${maxId + 1}\n0000000000 65535 f \n`;
  for (let id = 1; id <= maxId; id++) xref += `${String(offsets[id]).padStart(10, "0")} 00000 n \n`;
  pdf += xref;
  pdf += `trailer\n<< /Size ${maxId + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;

  const bytes = new Uint8Array(pdf.length);
  for (let i = 0; i < pdf.length; i++) bytes[i] = pdf.charCodeAt(i) & 0xff;
  return bytes;
}
function pdfBlob(job) { return new Blob([buildPdfBytes(job)], { type: "application/pdf" }); }
function downloadPdf(job) { triggerBlobDownload(pdfBlob(job), `${reportBaseName(job)}.pdf`); }

// ---------- email share ----------
// Tries the native share sheet (attaches the real file on supported mobile
// browsers); falls back to a prefilled mailto with a reminder to attach the
// file that was just downloaded.
async function shareJobByEmail(job, kind) {
  const { total, done, pct } = jobStats(job);
  const subject = `Shift Report \u2014 ${job.siteName} \u2014 ${fmtDate(job.startedAt)}`;
  const bodyLines = `Shift report for ${job.siteName} on ${fmtDate(job.startedAt)}.\n${done}/${total} tasks complete (${pct}%).`;

  const isExcel = kind === "excel";
  const blob = isExcel ? excelBlob(job) : pdfBlob(job);
  const filename = `${reportBaseName(job)}.${isExcel ? "xlsx" : "pdf"}`;
  const mime = isExcel ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/pdf";

  try {
    const file = new File([blob], filename, { type: mime });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: subject, text: bodyLines });
      return "shared";
    }
  } catch (e) {
    if (e && e.name === "AbortError") return "cancelled";
    console.warn("Web Share failed, falling back to mailto", e);
  }
  // Fallback: download the file so it's ready to attach, then open mail client.
  triggerBlobDownload(blob, filename);
  const body = encodeURIComponent(`${bodyLines}\n\n(The ${isExcel ? "Excel" : "PDF"} report just downloaded \u2014 attach it to this email before sending.)`);
  window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${body}`;
  return "fallback";
}

// ---------- Top bar ----------
function TopBar({ view, setView, hasActiveJob }) {
  const NavBtn = ({ id, label, icon: Icon }) => (
    <button onClick={() => setView(id)} style={{ ...styles.navBtn, ...(view === id ? styles.navBtnActive : {}) }}>
      <Icon size={15} strokeWidth={2.4} /> <span className="so-nav-label">{label}</span>
    </button>
  );
  return (
    <div style={styles.topbar} className="so-topbar">
      <div style={styles.logoBlock}>
        <img src={CODEE_LOGO_SRC} alt="Codee Cleaning Services" style={styles.logoImg} />
        <div style={styles.logoDivider} />
        <div>
          <div style={styles.logoTitle}>SITE OPERATING SYSTEM</div>
          <div style={styles.logoSub} className="so-logo-sub">Right crew. Right flow. Right standard. Every shift.</div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <NavBtn id="dashboard" label="Sites" icon={Building2} />
        {hasActiveJob && <NavBtn id="shift" label="Active Shift" icon={Play} />}
        <NavBtn id="manage" label="Manage Scope" icon={Settings} />
        <NavBtn id="history" label="History" icon={History} />
      </div>
    </div>
  );
}

// ---------- Dashboard ----------
function Dashboard({ sites, onStart, onEdit, onNewSite, onDeleteSite, onDuplicateSite }) {
  return (
    <div style={styles.page}>
      <div style={styles.pageHeaderRow}>
        <div>
          <div style={styles.eyebrow}>SITE PROFILES</div>
          <h1 style={styles.h1}>Choose a site to run tonight</h1>
        </div>
        <button style={styles.primaryBtn} onClick={onNewSite}><Plus size={16} /> New site</button>
      </div>

      <div style={styles.siteGrid}>
        {sites.map(site => (
          <div key={site.id} style={styles.siteCard}>
            {site.image ? (
              <img src={site.image} alt="" style={styles.siteCardImage} />
            ) : (
              <div style={styles.siteCardImagePlaceholder}><Building2 size={26} color="#D8D4CE" /></div>
            )}
            <div style={styles.siteCardTop}>
              <div style={styles.siteIconBadge}><Building2 size={20} color={BRAND_RED} /></div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={styles.siteCardName}>{site.name}</div>
                <div style={styles.siteCardBuilding}>{site.building}</div>
              </div>
            </div>
            <div style={styles.siteMetaGrid}>
              <MetaRow icon={Building2} text={site.levels} />
              <MetaRow icon={Users} text={`${site.crewSize} crew members \u00b7 ${site.groups.length} task groups`} />
              <MetaRow icon={Clock} text={site.hours} />
              <MetaRow icon={Calendar} text={site.days} />
            </div>
            <div style={styles.siteCardStat}>{totalTasks(site)} tasks scoped</div>
            <div style={styles.siteCardActions}>
              <button style={styles.startBtn} onClick={() => onStart(site)}><Play size={15} /> Start shift</button>
              <button style={styles.iconBtn} onClick={() => onEdit(site.id)} title="Edit scope"><Pencil size={15} /></button>
              <button style={styles.iconBtn} onClick={() => onDuplicateSite(site.id)} title="Duplicate as new site"><Copy size={15} /></button>
              <button style={styles.iconBtnDanger} onClick={() => { if (confirm(`Delete "${site.name}"? This cannot be undone.`)) onDeleteSite(site.id); }} title="Delete site"><Trash2 size={15} /></button>
            </div>
          </div>
        ))}
        {sites.length === 0 && <div style={styles.emptyState}>No sites yet. Add one to define its cleaning scope.</div>}
      </div>
    </div>
  );
}
function MetaRow({ icon: Icon, text }) {
  return <div style={styles.metaRow}><Icon size={13} color="#9A9DA3" /><span>{text}</span></div>;
}

// ---------- Site Editor ----------
function SiteEditor({ site, sites, onCancel, onSave }) {
  const blank = () => ({ id: uid("site"), name: "", building: "", levels: "", crewSize: 1, hours: "", days: "", image: null, createdAt: Date.now(), groups: [] });
  const [draft, setDraft] = useState(() => site ? JSON.parse(JSON.stringify(site)) : blank());
  const [openGroup, setOpenGroup] = useState(draft.groups[0]?.id || null);
  const [imgBusy, setImgBusy] = useState(false);
  const set = (patch) => setDraft(d => ({ ...d, ...patch }));

  // Preset library = the 6 built-in standard crews, plus every distinct
  // task group already scoped on other sites \u2014 so scope built for one
  // site becomes a reusable starting point for the next one.
  const presetLibrary = useMemo(() => {
    const seen = new Set();
    const lib = [];
    PRESET_GROUPS_TEMPLATE.forEach(tpl => {
      const key = `${tpl.code}|${tpl.label}`.toLowerCase();
      if (!seen.has(key)) { seen.add(key); lib.push({ ...tpl, source: "Standard" }); }
    });
    (sites || []).forEach(s => {
      if (s.id === draft.id) return;
      (s.groups || []).forEach(g => {
        const key = `${g.code}|${g.label}`.toLowerCase();
        if (seen.has(key) || !g.tasks.length) return;
        seen.add(key);
        lib.push({ code: g.code, label: g.label, color: g.color, timeframe: g.timeframe, tasks: g.tasks.map(t => t.text), source: s.name });
      });
    });
    return lib;
  }, [sites, draft.id]);

  const handleImageChange = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setImgBusy(true);
    try {
      const dataUrl = await resizeImageFile(file);
      set({ image: dataUrl });
    } catch (err) {
      console.error(err);
      alert("Could not load that photo. Try a different file.");
    } finally {
      setImgBusy(false);
    }
  };

  const addGroup = () => {
    const g = { id: uid("grp"), code: `G${draft.groups.length + 1}`, label: "New task group", color: PALETTE[draft.groups.length % PALETTE.length].hex, timeframe: "", tasks: [] };
    set({ groups: [...draft.groups, g] }); setOpenGroup(g.id);
  };
  const addPresetGroup = (tpl) => {
    const g = instantiatePresetGroup(tpl);
    set({ groups: [...draft.groups, g] });
    setOpenGroup(g.id);
  };
  const addAllStandardPresets = () => {
    const added = PRESET_GROUPS_TEMPLATE.map(instantiatePresetGroup);
    set({ groups: [...draft.groups, ...added] });
    setOpenGroup(added[0]?.id || null);
  };
  const updateGroup = (id, patch) => set({ groups: draft.groups.map(g => g.id === id ? { ...g, ...patch } : g) });
  const removeGroup = (id) => set({ groups: draft.groups.filter(g => g.id !== id) });
  const addTask = (groupId) => set({ groups: draft.groups.map(g => g.id !== groupId ? g : { ...g, tasks: [...g.tasks, { id: uid("task"), text: "" }] }) });
  const updateTask = (groupId, taskId, text) => set({ groups: draft.groups.map(g => g.id !== groupId ? g : { ...g, tasks: g.tasks.map(t => t.id === taskId ? { ...t, text } : t) }) });
  const removeTask = (groupId, taskId) => set({ groups: draft.groups.map(g => g.id !== groupId ? g : { ...g, tasks: g.tasks.filter(t => t.id !== taskId) }) });

  const canSave = draft.name.trim().length > 0 && draft.groups.length > 0;

  return (
    <div style={styles.page}>
      <div style={styles.pageHeaderRow}>
        <div>
          <button style={styles.backLink} onClick={onCancel}><ArrowLeft size={14} /> Back to sites</button>
          <h1 style={styles.h1}>{site ? "Edit scope" : "New site scope"}</h1>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={styles.secondaryBtn} onClick={onCancel}>Cancel</button>
          <button style={{ ...styles.primaryBtn, opacity: canSave ? 1 : 0.4 }} disabled={!canSave} onClick={() => onSave(draft)}><Save size={16} /> Save site</button>
        </div>
      </div>

      <div style={styles.editorProfileCard}>
        <div style={styles.imageUploadRow}>
          <div style={styles.imagePreviewBox}>
            {draft.image ? <img src={draft.image} alt="" style={styles.imagePreviewImg} /> : <Building2 size={22} color="#C9C5BF" />}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <label style={styles.fieldLabel}>Site photo (optional)</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <label style={styles.uploadBtn}>
                {imgBusy ? "Processing\u2026" : (draft.image ? "Change photo" : "Upload photo")}
                <input type="file" accept="image/*" style={{ display: "none" }} onChange={handleImageChange} disabled={imgBusy} />
              </label>
              {draft.image && <button style={styles.iconBtnDanger} onClick={() => set({ image: null })} title="Remove photo"><X size={14} /></button>}
            </div>
          </div>
        </div>
        <div style={styles.editorGridRow}>
          <Field label="Site name"><input style={styles.input} value={draft.name} onChange={e => set({ name: e.target.value })} placeholder="e.g. Hancock Iron Ore" /></Field>
          <Field label="Building"><input style={styles.input} value={draft.building} onChange={e => set({ building: e.target.value })} placeholder="e.g. Hancock Building" /></Field>
        </div>
        <div style={styles.editorGridRow}>
          <Field label="Levels / floors"><input style={styles.input} value={draft.levels} onChange={e => set({ levels: e.target.value })} placeholder="e.g. Ground + Levels 1\u20135" /></Field>
          <Field label="Crew members"><input type="number" min={1} style={styles.input} value={draft.crewSize} onChange={e => set({ crewSize: Number(e.target.value) })} /></Field>
        </div>
        <div style={styles.editorGridRow}>
          <Field label="Shift hours"><input style={styles.input} value={draft.hours} onChange={e => set({ hours: e.target.value })} placeholder="e.g. 5:00 PM \u2013 8:30 PM" /></Field>
          <Field label="Days"><input style={styles.input} value={draft.days} onChange={e => set({ days: e.target.value })} placeholder="e.g. Monday \u2013 Friday" /></Field>
        </div>
      </div>

      <div style={styles.editorSectionHead}>
        <div>
          <div style={styles.eyebrow}>PRESET TASK GROUPS</div>
          <div style={styles.presetHint}>Add a ready-made crew in one click \u2014 pulled from the standard scope and every other site you've built.</div>
        </div>
        <button style={styles.dashedAddBtn} onClick={addAllStandardPresets}><Plus size={14} /> Add all 6 standard crews</button>
      </div>
      <div style={styles.presetChipRow}>
        {presetLibrary.map((tpl, i) => (
          <button key={`${tpl.code}-${tpl.label}-${i}`} style={styles.presetChip} onClick={() => addPresetGroup(tpl)} title={`From: ${tpl.source}`}>
            <span style={{ ...styles.presetChipDot, background: tpl.color }} />
            <span style={styles.presetChipLabel}>{tpl.label}</span>
            <span style={styles.presetChipCount}>{tpl.tasks.length}</span>
          </button>
        ))}
      </div>

      <div style={styles.editorSectionHead}>
        <div style={styles.eyebrow}>TASK GROUPS ({draft.groups.length})</div>
        <button style={styles.dashedAddBtn} onClick={addGroup}><Plus size={14} /> Add blank task group</button>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {draft.groups.map(g => {
          const open = openGroup === g.id;
          return (
            <div key={g.id} style={{ ...styles.groupCard, borderLeft: `4px solid ${g.color}` }}>
              <div style={styles.groupCardHead} onClick={() => setOpenGroup(open ? null : g.id)}>
                {open ? <ChevronDown size={16} color="#9A9DA3" /> : <ChevronRight size={16} color="#9A9DA3" />}
                <span style={{ ...styles.groupCodeBadge, background: g.color }}>{g.code || "??"}</span>
                <span style={styles.groupCardLabel}>{g.label || "Untitled group"}</span>
                <span style={styles.groupCardCount}>{g.tasks.length} tasks</span>
                <button style={styles.iconBtnDanger} onClick={(e) => { e.stopPropagation(); removeGroup(g.id); }}><Trash2 size={14} /></button>
              </div>
              {open && (
                <div style={styles.groupCardBody}>
                  <div style={styles.editorGridRow}>
                    <Field label="Code"><input style={styles.input} value={g.code} onChange={e => updateGroup(g.id, { code: e.target.value })} placeholder="CC1" /></Field>
                    <Field label="Group / crew name"><input style={styles.input} value={g.label} onChange={e => updateGroup(g.id, { label: e.target.value })} placeholder="e.g. Bathrooms & Consumables Lead" /></Field>
                  </div>
                  <div style={styles.editorGridRow}>
                    <Field label="Time frame"><input style={styles.input} value={g.timeframe} onChange={e => updateGroup(g.id, { timeframe: e.target.value })} placeholder="e.g. 5:30PM \u2013 7:10PM" /></Field>
                    <Field label="Color">
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", paddingTop: 4 }}>
                        {PALETTE.map(p => (
                          <button key={p.hex} onClick={() => updateGroup(g.id, { color: p.hex })}
                            style={{ width: 26, height: 26, borderRadius: 6, background: p.hex, border: g.color === p.hex ? "2px solid #2B2C30" : "2px solid transparent", cursor: "pointer" }}
                            title={p.name} />
                        ))}
                      </div>
                    </Field>
                  </div>
                  <div style={{ marginTop: 10 }}>
                    <div style={styles.taskListLabel}>Tasks</div>
                    {g.tasks.map(t => (
                      <div key={t.id} style={styles.taskEditRow}>
                        <input style={{ ...styles.input, flex: 1 }} value={t.text} onChange={e => updateTask(g.id, t.id, e.target.value)} placeholder="Describe the task…" />
                        <button style={styles.iconBtnDanger} onClick={() => removeTask(g.id, t.id)}><X size={14} /></button>
                      </div>
                    ))}
                    <button style={styles.dashedAddBtnSmall} onClick={() => addTask(g.id)}><Plus size={13} /> Add task</button>
                  </div>
                </div>
              )}
            </div>
          );
        })}
        {draft.groups.length === 0 && <div style={styles.emptyState}>No task groups yet. Add one to start scoping this site.</div>}
      </div>
    </div>
  );
}
function Field({ label, children }) {
  return <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 5 }}><label style={styles.fieldLabel}>{label}</label>{children}</div>;
}

// ---------- Shift Runner ----------
function ShiftRunner({ job, onToggleTask, onComplete, onAbandon }) {
  const [now, setNow] = useState(Date.now());
  const gradId = useMemo(() => uid("ring"), []);
  useEffect(() => { const t = setInterval(() => setNow(Date.now()), 1000); return () => clearInterval(t); }, []);

  const total = job.groups.reduce((n, g) => n + g.tasks.length, 0);
  const done = job.groups.reduce((n, g) => n + g.tasks.filter(t => t.done).length, 0);
  const pct = total ? done / total : 0;
  const allDone = total > 0 && done === total;

  return (
    <div style={styles.page}>
      <div style={styles.shiftHeader}>
        {job.siteImage && <img src={job.siteImage} alt="" style={styles.shiftThumb} />}
        <div style={{ flex: 1, minWidth: 180 }}>
          <div style={styles.eyebrow}>ACTIVE SHIFT</div>
          <h1 style={styles.h1}>{job.siteName}</h1>
          <div style={styles.shiftMetaRow}><Timer size={14} color="#9A9DA3" /> <span>Started {fmtClock(job.startedAt)} \u00b7 running {fmtDuration(now - job.startedAt)}</span></div>
        </div>
        <div style={styles.ringWrap}><RingGauge pct={pct} gradId={gradId} /></div>
      </div>

      <div style={styles.overallBarWrap}>
        <SegmentedCrewBar groups={job.groups} />
        <div style={styles.overallBarLabel}>{done} / {total} tasks complete</div>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {job.groups.map(g => <GroupChecklist key={g.id} group={g} onToggleTask={(taskId) => onToggleTask(g.id, taskId)} />)}
      </div>

      <div style={styles.stickyFooter} className="so-sticky-footer">
        <button style={styles.textDangerBtn} onClick={() => { if (confirm("Discard this active shift? Ticked progress will be lost.")) onAbandon(); }}>Discard shift</button>
        <button style={{ ...styles.completeBtn, opacity: allDone ? 1 : 0.55 }} onClick={() => { if (allDone || confirm(`${total - done} tasks are still open. Complete shift anyway?`)) onComplete(); }}>
          <CircleCheck size={17} /> Complete shift & export
        </button>
      </div>
      <div style={{ height: 72 }} />
    </div>
  );
}

function GroupChecklist({ group, onToggleTask }) {
  const [open, setOpen] = useState(true);
  const done = group.tasks.filter(t => t.done).length;
  const total = group.tasks.length;
  const complete = total > 0 && done === total;
  return (
    <div style={{ ...styles.groupPanel, borderLeft: `4px solid ${group.color}`, opacity: complete ? 0.8 : 1 }}>
      <div style={styles.groupPanelHead} onClick={() => setOpen(!open)}>
        {open ? <ChevronDown size={16} color="#9A9DA3" /> : <ChevronRight size={16} color="#9A9DA3" />}
        <span style={{ ...styles.groupCodeBadge, background: group.color }}>{group.code}</span>
        <div style={{ flex: 1 }}>
          <div style={styles.groupPanelLabel}>{group.label}</div>
          {group.timeframe && <div style={styles.groupPanelTimeframe}>{group.timeframe}</div>}
        </div>
        <div style={{ ...styles.groupPanelCount, color: complete ? "#16A34A" : "#9A9DA3" }}>{done}/{total}{complete && <Check size={13} style={{ marginLeft: 4 }} />}</div>
      </div>
      {open && (
        <div style={styles.groupPanelBody}>
          {group.tasks.map(t => (
            <label key={t.id} style={styles.taskRow}>
              <input type="checkbox" checked={t.done} onChange={() => onToggleTask(t.id)} style={styles.checkbox} />
              <span style={{ ...styles.taskText, textDecoration: t.done ? "line-through" : "none", color: t.done ? "#B9BBC0" : "#2B2C30" }}>{t.text}</span>
              {t.done && t.doneAt && <span style={styles.taskDoneAt}>{fmtClock(t.doneAt)}</span>}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Completion Modal ----------
function CompletionModal({ job, onClose }) {
  const [sending, setSending] = useState(false);
  const total = job.groups.reduce((n, g) => n + g.tasks.length, 0);
  const done = job.groups.reduce((n, g) => n + g.tasks.filter(t => t.done).length, 0);
  const duration = fmtDuration(job.completedAt - job.startedAt);

  const email = async () => {
    setSending(true);
    try { await shareJobByEmail(job, "pdf"); } finally { setSending(false); }
  };

  return (
    <div style={styles.modalOverlay}>
      <div style={styles.modalCard}>
        <div style={styles.modalIcon}><div style={{ width: 64, height: 64 }}><DotMark size={64} /></div></div>
        <h2 style={styles.modalTitle}>Shift complete</h2>
        <div style={styles.modalSub}>{job.siteName} \u00b7 {duration}</div>
        <div style={styles.modalStatRow}>
          <div style={styles.modalStat}><div style={styles.modalStatNum}>{done}/{total}</div><div style={styles.modalStatLbl}>Tasks done</div></div>
          <div style={styles.modalStat}><div style={styles.modalStatNum}>{total ? Math.round((done / total) * 100) : 0}%</div><div style={styles.modalStatLbl}>Completion</div></div>
        </div>
        <div style={styles.modalExportRow}>
          <button style={styles.modalExportBtn} onClick={() => downloadExcel(job)}><FileSpreadsheet size={16} /> Excel</button>
          <button style={styles.modalExportBtn} onClick={() => downloadPdf(job)}><FileText size={16} /> PDF</button>
        </div>
        <button style={{ ...styles.primaryBtnWide, opacity: sending ? 0.7 : 1 }} disabled={sending} onClick={email}>
          <Mail size={16} /> {sending ? "Opening share sheet\u2026" : "Share via email"}
        </button>
        <button style={styles.secondaryBtnWide} onClick={onClose}>Back to sites</button>
      </div>
    </div>
  );
}

// ---------- History ----------
function HistoryView({ history, onLoadJob }) {
  const [busyId, setBusyId] = useState(null);

  const run = async (id, action) => {
    setBusyId(id);
    try {
      const job = await onLoadJob(id);
      if (job) await action(job);
    } finally {
      setBusyId(null);
    }
  };

  return (
    <div style={styles.page}>
      <div style={styles.eyebrow}>SHIFT HISTORY</div>
      <h1 style={styles.h1}>Completed shifts</h1>
      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 16 }}>
        {history.map(j => (
          <div key={j.id} style={styles.historyRow}>
            <div style={{ flex: 1, minWidth: 140 }}>
              <div style={{ ...styles.historyName, overflowWrap: "break-word" }}>{j.siteName}</div>
              <div style={{ ...styles.historyMeta, overflowWrap: "break-word" }}>{fmtDate(j.startedAt)} \u00b7 {fmtClock(j.startedAt)}\u2013{fmtClock(j.completedAt)} \u00b7 {j.done}/{j.total} tasks ({j.total ? Math.round((j.done / j.total) * 100) : 0}%)</div>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button style={styles.iconBtn} disabled={busyId === j.id} title="Download Excel" onClick={() => run(j.id, (job) => downloadExcel(job))}><FileSpreadsheet size={15} /></button>
              <button style={styles.iconBtn} disabled={busyId === j.id} title="Download PDF" onClick={() => run(j.id, (job) => downloadPdf(job))}><FileText size={15} /></button>
              <button style={styles.iconBtn} disabled={busyId === j.id} title="Share via email" onClick={() => run(j.id, (job) => shareJobByEmail(job, "pdf"))}><Mail size={15} /></button>
            </div>
          </div>
        ))}
        {history.length === 0 && <div style={styles.emptyState}>No completed shifts yet. Run a shift and export it to see it here.</div>}
      </div>
    </div>
  );
}

// ---------- styles ----------
const FONT_DISPLAY = "'Poppins', 'Segoe UI', sans-serif";
const FONT_BODY = "'Inter', 'Segoe UI', sans-serif";

const globalCss = `
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700;800&family=Inter:wght@400;500;600;700;800&display=swap');
* { box-sizing: border-box; font-family: ${FONT_BODY}; }
input[type=checkbox] { accent-color: ${BRAND_RED}; }
button { font-family: ${FONT_BODY}; }
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-thumb { background: #E2DFDB; border-radius: 4px; }

/* ---- responsive overrides: mobile / tablet ---- */
@media (max-width: 700px) {
  .so-sticky-footer { flex-direction: column-reverse !important; align-items: stretch !important; gap: 10px !important; }
  .so-sticky-footer button { width: 100% !important; justify-content: center !important; }
}
@media (max-width: 560px) {
  .so-logo-sub { display: none !important; }
  .so-topbar { padding: 12px 14px !important; }
}
@media (max-width: 420px) {
  .so-nav-label { display: none; }
}
`;

const styles = {
  app: { background: "#FAF9F7", minHeight: "100vh", color: "#2B2C30" },
  loadingScreen: { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh", background: "#FAF9F7", gap: 14, padding: 20 },
  topbar: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px clamp(14px, 4vw, 24px)", background: "#FFFFFF", borderBottom: "1px solid #ECE9E4", flexWrap: "wrap", gap: 10 },
  logoBlock: { display: "flex", alignItems: "center", gap: 14, minWidth: 0 },
  logoImg: { height: "clamp(24px, 6vw, 30px)", width: "auto", display: "block", flexShrink: 0 },
  logoDivider: { width: 1, height: 28, background: "#E5E2DD", flexShrink: 0 },
  logoTitle: { fontFamily: FONT_DISPLAY, fontSize: 13, letterSpacing: 0.6, color: "#2B2C30", fontWeight: 700, whiteSpace: "nowrap" },
  logoSub: { fontSize: 11, color: "#9A9DA3", marginTop: 2, whiteSpace: "nowrap" },
  navBtn: { display: "flex", alignItems: "center", gap: 6, background: "transparent", border: "1px solid #ECE9E4", color: "#5B5D63", padding: "8px 12px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" },
  navBtnActive: { background: BRAND_GRADIENT, borderColor: "transparent", color: "#fff" },
  body: { maxWidth: 1080, margin: "0 auto", padding: "0 clamp(14px, 4vw, 24px)" },
  page: { padding: "clamp(18px, 4vw, 28px) 0 40px" },
  pageHeaderRow: { display: "flex", alignItems: "flex-end", justifyContent: "space-between", marginBottom: 20, flexWrap: "wrap", gap: 12 },
  eyebrow: { fontSize: 11, letterSpacing: 1.6, color: BRAND_RED, fontWeight: 700, marginBottom: 6 },
  h1: { fontFamily: FONT_DISPLAY, fontSize: "clamp(20px, 4.5vw, 26px)", margin: 0, color: "#2B2C30", letterSpacing: 0.1, fontWeight: 700 },
  primaryBtn: { display: "flex", alignItems: "center", gap: 8, background: BRAND_GRADIENT, color: "#fff", border: "none", padding: "10px 16px", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer", boxShadow: "0 4px 14px rgba(238,58,35,0.25)", whiteSpace: "nowrap" },
  primaryBtnWide: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, width: "100%", background: BRAND_GRADIENT, color: "#fff", border: "none", padding: "12px 16px", borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: "pointer", marginTop: 18, boxShadow: "0 4px 14px rgba(238,58,35,0.25)" },
  secondaryBtn: { display: "flex", alignItems: "center", gap: 6, background: "#FFFFFF", border: "1px solid #E5E2DD", color: "#2B2C30", padding: "9px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" },
  secondaryBtnWide: { width: "100%", background: "#FFFFFF", border: "1px solid #E5E2DD", color: "#5B5D63", padding: "10px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer", marginTop: 8 },
  backLink: { display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", color: "#9A9DA3", fontSize: 12, cursor: "pointer", padding: 0, marginBottom: 10 },
  siteGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(min(280px, 100%), 1fr))", gap: 16 },
  siteCard: { background: "#FFFFFF", border: "1px solid #ECE9E4", borderRadius: 14, padding: 18, display: "flex", flexDirection: "column", gap: 12, boxShadow: "0 1px 3px rgba(20,20,20,0.04)", overflow: "hidden" },
  siteCardImage: { margin: "-18px -18px 0 -18px", width: "calc(100% + 36px)", maxWidth: "calc(100% + 36px)", height: "clamp(100px, 24vw, 150px)", objectFit: "cover", display: "block" },
  siteCardImagePlaceholder: { margin: "-18px -18px 0 -18px", width: "calc(100% + 36px)", height: "clamp(70px, 16vw, 96px)", background: "#FAF9F7", display: "flex", alignItems: "center", justifyContent: "center" },
  siteCardTop: { display: "flex", gap: 12, alignItems: "flex-start" },
  siteIconBadge: { width: 40, height: 40, borderRadius: 9, background: "#FDECE9", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  siteCardName: { fontFamily: FONT_DISPLAY, fontSize: 16, color: "#2B2C30", fontWeight: 700, overflowWrap: "break-word" },
  siteCardBuilding: { fontSize: 12, color: "#9A9DA3", marginTop: 2, overflowWrap: "break-word" },
  siteMetaGrid: { display: "flex", flexDirection: "column", gap: 6, borderTop: "1px solid #F1EFEB", borderBottom: "1px solid #F1EFEB", padding: "10px 0" },
  metaRow: { display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "#5B5D63" },
  siteCardStat: { fontSize: 11, color: BRAND_RED, fontWeight: 700, letterSpacing: 0.4 },
  siteCardActions: { display: "flex", gap: 8, flexWrap: "wrap" },
  startBtn: { flex: 1, minWidth: 120, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: BRAND_GRADIENT, color: "#fff", border: "none", padding: "10px 14px", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer" },
  iconBtn: { display: "flex", alignItems: "center", justifyContent: "center", width: 36, background: "#FFFFFF", border: "1px solid #E5E2DD", color: "#5B5D63", borderRadius: 8, cursor: "pointer", flexShrink: 0 },
  iconBtnDanger: { display: "flex", alignItems: "center", justifyContent: "center", width: 32, height: 32, background: "#FFF5F4", border: "1px solid #F8D9D4", color: "#D93B22", borderRadius: 8, cursor: "pointer", flexShrink: 0 },
  emptyState: { color: "#9A9DA3", fontSize: 13, padding: "30px 0", textAlign: "center", border: "1px dashed #E5E2DD", borderRadius: 12, gridColumn: "1/-1" },
  editorProfileCard: { background: "#FFFFFF", border: "1px solid #ECE9E4", borderRadius: 14, padding: 18, display: "flex", flexDirection: "column", gap: 16, marginBottom: 22, boxShadow: "0 1px 3px rgba(20,20,20,0.04)" },
  editorGridRow: { display: "flex", gap: 14, flexWrap: "wrap" },
  fieldLabel: { fontSize: 11, color: "#9A9DA3", fontWeight: 600, letterSpacing: 0.3 },
  input: { background: "#FAF9F7", border: "1px solid #E5E2DD", color: "#2B2C30", borderRadius: 7, padding: "9px 11px", fontSize: 13.5, width: "100%", outline: "none" },
  editorSectionHead: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10, flexWrap: "wrap", gap: 8 },
  presetHint: { fontSize: 11.5, color: "#9A9DA3", marginTop: 2, maxWidth: 420 },
  presetChipRow: { display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 22 },
  presetChip: { display: "flex", alignItems: "center", gap: 7, background: "#FFFFFF", border: "1px solid #E5E2DD", color: "#2B2C30", padding: "7px 12px 7px 10px", borderRadius: 20, fontSize: 12.5, fontWeight: 600, cursor: "pointer" },
  presetChipDot: { width: 9, height: 9, borderRadius: "50%", flexShrink: 0, display: "inline-block" },
  presetChipLabel: { maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  presetChipCount: { fontSize: 10.5, color: "#9A9DA3", background: "#FAF9F7", borderRadius: 10, padding: "1px 6px", fontWeight: 700 },
  dashedAddBtn: { display: "flex", alignItems: "center", gap: 6, background: "transparent", border: "1px dashed #D8D4CE", color: "#5B5D63", padding: "8px 12px", borderRadius: 8, fontSize: 12.5, cursor: "pointer" },
  dashedAddBtnSmall: { display: "flex", alignItems: "center", gap: 5, background: "transparent", border: "1px dashed #D8D4CE", color: "#9A9DA3", padding: "6px 10px", borderRadius: 7, fontSize: 12, cursor: "pointer", marginTop: 6 },
  imageUploadRow: { display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap" },
  imagePreviewBox: { width: 72, height: 72, borderRadius: 10, background: "#FAF9F7", border: "1px solid #ECE9E4", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0 },
  imagePreviewImg: { width: "100%", height: "100%", objectFit: "cover", display: "block" },
  uploadBtn: { display: "inline-flex", alignItems: "center", gap: 6, background: "#FFFFFF", border: "1px dashed #D8D4CE", color: "#5B5D63", padding: "8px 14px", borderRadius: 8, fontSize: 12.5, fontWeight: 600, cursor: "pointer" },
  groupCard: { background: "#FFFFFF", border: "1px solid #ECE9E4", borderRadius: 10, overflow: "hidden" },
  groupCardHead: { display: "flex", alignItems: "center", gap: 10, padding: "12px 14px", cursor: "pointer" },
  groupCodeBadge: { fontSize: 10.5, fontWeight: 800, color: "#fff", padding: "3px 7px", borderRadius: 5, letterSpacing: 0.3, flexShrink: 0 },
  groupCardLabel: { flex: 1, fontSize: 13.5, fontWeight: 600, color: "#2B2C30", minWidth: 0, overflowWrap: "break-word" },
  groupCardCount: { fontSize: 11.5, color: "#9A9DA3", marginRight: 6, flexShrink: 0 },
  groupCardBody: { padding: "0 14px 16px 14px", borderTop: "1px solid #F1EFEB" },
  taskListLabel: { fontSize: 11, color: "#9A9DA3", fontWeight: 600, margin: "12px 0 6px" },
  taskEditRow: { display: "flex", gap: 8, marginBottom: 6, alignItems: "center" },
  shiftHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, marginBottom: 16, flexWrap: "wrap" },
  shiftThumb: { width: 52, height: 52, borderRadius: 10, objectFit: "cover", flexShrink: 0 },
  shiftMetaRow: { display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "#9A9DA3", marginTop: 6, flexWrap: "wrap" },
  ringWrap: { width: "clamp(64px, 16vw, 84px)", height: "clamp(64px, 16vw, 84px)", flexShrink: 0 },
  overallBarWrap: { marginBottom: 20 },
  segBarTrack: { display: "flex", height: 16, borderRadius: 8, overflow: "hidden", border: "1px solid #ECE9E4", background: "#F1EFEB", gap: 2 },
  segBarSeg: { position: "relative", height: "100%", background: "#F1EFEB", minWidth: 3 },
  segBarFill: { position: "absolute", left: 0, top: 0, bottom: 0, transition: "width .4s" },
  segBarLegend: { display: "flex", flexWrap: "wrap", gap: "6px 14px", marginTop: 8 },
  segBarLegendItem: { display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, color: "#5B5D63", fontWeight: 600 },
  segBarDot: { width: 8, height: 8, borderRadius: "50%", flexShrink: 0, display: "inline-block" },
  overallBarLabel: { fontSize: 12, color: "#9A9DA3", marginTop: 8, fontWeight: 600 },
  groupPanel: { background: "#FFFFFF", border: "1px solid #ECE9E4", borderRadius: 12, overflow: "hidden", boxShadow: "0 1px 3px rgba(20,20,20,0.03)" },
  groupPanelHead: { display: "flex", alignItems: "center", gap: 10, padding: "13px 14px", cursor: "pointer", flexWrap: "wrap" },
  groupPanelLabel: { fontSize: 14, fontWeight: 700, color: "#2B2C30", overflowWrap: "break-word" },
  groupPanelTimeframe: { fontSize: 11.5, color: "#9A9DA3", marginTop: 1 },
  groupPanelCount: { fontSize: 12.5, fontWeight: 700, display: "flex", alignItems: "center", flexShrink: 0 },
  groupPanelBody: { padding: "4px 14px 12px", borderTop: "1px solid #F1EFEB" },
  taskRow: { display: "flex", alignItems: "flex-start", gap: 10, padding: "8px 4px", cursor: "pointer", borderBottom: "1px solid #F5F3EF", flexWrap: "wrap" },
  checkbox: { width: 17, height: 17, cursor: "pointer", flexShrink: 0, marginTop: 2 },
  taskText: { fontSize: 13.5, flex: 1, minWidth: 140, overflowWrap: "break-word" },
  taskDoneAt: { fontSize: 10.5, color: "#16A34A", fontWeight: 600, flexShrink: 0 },
  stickyFooter: { position: "sticky", bottom: 0, display: "flex", justifyContent: "space-between", alignItems: "center", background: "rgba(250,249,247,0.92)", backdropFilter: "blur(6px)", padding: "14px 0", borderTop: "1px solid #ECE9E4", marginTop: 10, gap: 10 },
  textDangerBtn: { background: "none", border: "none", color: "#D93B22", fontSize: 12.5, cursor: "pointer", fontWeight: 600 },
  completeBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "#16A34A", color: "#fff", border: "none", padding: "12px 20px", borderRadius: 9, fontWeight: 800, fontSize: 14, cursor: "pointer", boxShadow: "0 4px 14px rgba(22,163,74,0.25)" },
  modalOverlay: { position: "fixed", inset: 0, background: "rgba(43,44,48,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 },
  modalCard: { background: "#FFFFFF", border: "1px solid #ECE9E4", borderRadius: 16, padding: "clamp(20px, 5vw, 28px)", width: "min(92vw, 340px)", textAlign: "center", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" },
  modalIcon: { display: "flex", justifyContent: "center", marginBottom: 10 },
  modalTitle: { fontFamily: FONT_DISPLAY, fontSize: 20, margin: "0 0 4px", fontWeight: 700, color: "#2B2C30" },
  modalSub: { fontSize: 12.5, color: "#9A9DA3", marginBottom: 16 },
  modalStatRow: { display: "flex", gap: 10, justifyContent: "center" },
  modalExportRow: { display: "flex", gap: 8, marginTop: 16 },
  modalExportBtn: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 7, background: "#FFFFFF", border: "1px solid #E5E2DD", color: "#2B2C30", padding: "10px 10px", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer" },
  modalStat: { flex: 1, background: "#FAF9F7", borderRadius: 10, padding: "10px 6px", border: "1px solid #F1EFEB" },
  modalStatNum: { fontFamily: FONT_DISPLAY, fontSize: 18, color: BRAND_RED, fontWeight: 700 },
  modalStatLbl: { fontSize: 10.5, color: "#9A9DA3", marginTop: 2 },
  historyRow: { display: "flex", alignItems: "center", gap: 12, background: "#FFFFFF", border: "1px solid #ECE9E4", borderRadius: 10, padding: "12px 14px", boxShadow: "0 1px 3px rgba(20,20,20,0.03)", flexWrap: "wrap" },
  historyName: { fontSize: 13.5, fontWeight: 700, color: "#2B2C30" },
  historyMeta: { fontSize: 11.5, color: "#9A9DA3", marginTop: 2 },
};
