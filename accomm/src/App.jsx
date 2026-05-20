import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── Storage helpers (localStorage) ───────────────────────────────────────────
const PROJECTS_INDEX_KEY = "accomm:projects";

function saveProject(id, data) {
  try {
    localStorage.setItem(`accomm:project:${id}`, JSON.stringify(data));
    let index = loadIndex();
    if (!index.find(p => p.id === id)) {
      index.push({ id, name: data.project.name || "Untitled", savedAt: Date.now() });
    } else {
      index = index.map(p => p.id === id ? { ...p, name: data.project.name || "Untitled", savedAt: Date.now() } : p);
    }
    localStorage.setItem(PROJECTS_INDEX_KEY, JSON.stringify(index));
  } catch (e) { console.error("Save failed", e); }
}

function loadProject(id) {
  try { const r = localStorage.getItem(`accomm:project:${id}`); return r ? JSON.parse(r) : null; } catch { return null; }
}

function loadIndex() {
  try { const r = localStorage.getItem(PROJECTS_INDEX_KEY); return r ? JSON.parse(r) : []; } catch { return []; }
}

function deleteProject(id) {
  try {
    localStorage.removeItem(`accomm:project:${id}`);
    const index = loadIndex().filter(p => p.id !== id);
    localStorage.setItem(PROJECTS_INDEX_KEY, JSON.stringify(index));
  } catch {}
}

// ── Constants ────────────────────────────────────────────────────────────────
const DOOR_TEST_OPTIONS = ["PASS", "FAIL", "N/A"];

// ── Helpers ──────────────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2, 8).toUpperCase();
const blankDoor = () => ({ id: uid(), doorRef: "", location: "", opportunityRef: "", installDate: new Date().toISOString().slice(0, 10), readerIn: "", readerOut: "", exitButton: "", breakGlassUnit: "", lock: "", doorHeld: "", doorForced: "", photoTag: "", photoDescription: "", photoDataUrl: "", notes: "" });
const blankCCTV = () => ({ id: uid(), cameraNumber: "", location: "", makeModel: "", ipAddress: "", supplyVolts: "", recordQuality: "", playbackQuality: "", recordFPS: "", macAddress: "", opportunityNumber: "", installDate: new Date().toISOString().slice(0, 10), photoTag: "", photoDescription: "", photoDataUrl: "", notes: "" });

// ── Export to Excel ───────────────────────────────────────────────────────────
function exportToExcel(project, doors, cctvs) {
  const wb = XLSX.utils.book_new();

  const addSheet = (name, rows) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name);
  };

  addSheet("Project Info", [
    ["ISD TECH COMMISSIONING ASSISTANT"],
    [],
    ["Project Name", project.name],
    ["Site Address", project.address],
    ["Client", project.client],
    ["Opportunity No.", project.contractNo],
    ["Engineer(s)", project.engineers],
    ["Start Date", project.startDate],
    ["Completion Date", project.completionDate],
    ["System Description", project.systemDescription],
    [],
    ["Generated", new Date().toLocaleString()],
  ]);

  addSheet("Doors", [
    ["Door Reference", "Location", "Opportunity Ref.", "Installation Date", "Reader In", "Reader Out", "Exit Button", "Break Glass Unit", "Lock", "Monitoring", "Controller","Door Held", "Door Forced", "Photo Ref.", "Photo Description", "Photo Attached", "Notes"],
    ...doors.map(d => [
      d.doorRef,
      d.location,
      d.opportunityRef,
      d.installDate,
      d.readerIn,
      d.readerOut,
      d.exitButton,
      d.breakGlassUnit,
      d.lock,
      d.monitoring,
      d.controller,
      d.doorHeld,
      d.doorForced,
      d.photoTag,
      d.photoDescription,
      d.photoDataUrl ? "Yes" : "No",
      d.notes,
    ]),
  ]);

  addSheet("CCTV", [
    ["Camera Number", "Location", "Camera Make", "Camera Model", "IP Address", "Supply Volts", "Record Quality", "Playback Quality", "Record FPS", "MAC Address", "Opportunity Number", "Installation Date", "Photo Ref.", "Photo Description", "Photo Attached", "Notes"],
    ...cctvs.map(c => [
      c.cameraNumber,
      c.location,
      c.camMake,
      c.camModel,
      c.ipAddress,
      c.supplyVolts,
      c.recordQuality,
      c.playbackQuality,
      c.recordFPS,
      c.macAddress,
      c.opportunityNumber,
      c.installDate,
      c.photoTag,
      c.photoDescription,
      c.photoDataUrl ? "Yes" : "No",
      c.notes,
    ]),
  ]);
  

  addSheet("Summary", [
    ["COMMISSIONING SUMMARY"],
    [],
    ["Total Doors", doors.length],
    ["Total CCTV Cameras", cctvs.length],
    ["Generated", new Date().toLocaleString()],
  ]);

  const fname = `${(project.name || "commissioning").replace(/\s+/g, "_")}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fname);
}

// ── Styles ───────────────────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@600;700;800&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #0d0f14;
    --surface: #161820;
    --surface2: #1e2029;
    --border: #2a2d3a;
    --accent: #e59c00;
    --warn: #ffd166;
    --danger: #ef4444;
    --text: #e8eaf0;
    --muted: #6b7280;
    --pass: #22c55e;
    --fail: #ef4444;
    --radius: 10px;
    --font-head: 'Syne', sans-serif;
    --font-mono: 'DM Mono', monospace;
  }

  body { background: var(--bg); color: var(--text); font-family: var(--font-mono); font-size: 13px; }

  .app { max-width: 480px; margin: 0 auto; min-height: 100vh; display: flex; flex-direction: column; }

  .header { background: var(--surface); border-bottom: 1px solid var(--border); padding: 14px 16px 12px; position: sticky; top: 0; z-index: 100; }
  .header-top { display: flex; align-items: center; justify-content: space-between; }
  .header-title { font-family: var(--font-head); font-size: 18px; font-weight: 800; letter-spacing: -0.5px; display: flex; align-items: center; gap: 8px; }
  .header-title span { color: var(--accent); }
  .project-name { font-size: 11px; color: var(--muted); margin-top: 2px; display: flex; align-items: center; gap: 6px; }
  .export-btn { background: var(--accent); color: #000; border: none; border-radius: 6px; padding: 8px 14px; font-family: var(--font-mono); font-size: 12px; font-weight: 500; cursor: pointer; white-space: nowrap; }
  .export-btn:active { opacity: 0.8; }
  .proj-switch-btn { background: var(--surface2); border: 1px solid var(--border); border-radius: 6px; color: var(--muted); font-size: 11px; padding: 3px 8px; cursor: pointer; font-family: var(--font-mono); font-weight: 400; }

  .nav { display: flex; background: var(--surface); border-bottom: 1px solid var(--border); overflow-x: auto; scrollbar-width: none; }
  .nav::-webkit-scrollbar { display: none; }
  .nav-tab { flex: none; padding: 10px 14px; font-family: var(--font-mono); font-size: 11px; font-weight: 500; color: var(--muted); border: none; background: none; cursor: pointer; border-bottom: 2px solid transparent; white-space: nowrap; transition: color 0.15s, border-color 0.15s; }
  .nav-tab.active { color: var(--accent); border-bottom-color: var(--accent); }
  .nav-count { background: var(--surface2); border-radius: 10px; padding: 1px 5px; margin-left: 4px; font-size: 10px; }
  .nav-tab.active .nav-count { background: var(--accent); color: #000; }

  .content { flex: 1; padding: 12px; overflow-y: auto; }

  .card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); margin-bottom: 10px; overflow: hidden; }
  .card-header { display: flex; align-items: center; justify-content: space-between; padding: 10px 12px; border-bottom: 1px solid var(--border); cursor: pointer; }
  .card-title { font-family: var(--font-head); font-size: 13px; font-weight: 700; display: flex; align-items: center; gap: 8px; }
  .card-tag { background: var(--accent); color: #000; border-radius: 4px; padding: 1px 6px; font-size: 10px; font-weight: 700; }
  .card-body { padding: 12px; display: flex; flex-direction: column; gap: 8px; }
  .card-collapsed .card-body { display: none; }

  .field { display: flex; flex-direction: column; gap: 4px; }
  .field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  label { font-size: 10px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; }
  input, select, textarea { background: var(--surface2); border: 1px solid var(--border); border-radius: 6px; color: var(--text); font-family: var(--font-mono); font-size: 13px; padding: 8px 10px; width: 100%; outline: none; transition: border-color 0.15s; -webkit-appearance: none; }
  input:focus, select:focus, textarea:focus { border-color: var(--accent); }
  textarea { resize: vertical; min-height: 60px; }
  select { background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%236b7280' fill='none' stroke-width='1.5'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 10px center; padding-right: 28px; }

  .test-grid { display: flex; flex-direction: column; gap: 6px; }
  .test-row { display: flex; align-items: center; justify-content: space-between; background: var(--surface2); border-radius: 6px; padding: 8px 10px; }
  .test-label { font-size: 12px; flex: 1; }
  .result-btns { display: flex; gap: 4px; }
  .result-btn { border: 1px solid var(--border); border-radius: 5px; padding: 4px 10px; font-family: var(--font-mono); font-size: 11px; font-weight: 500; cursor: pointer; background: var(--surface); color: var(--muted); transition: all 0.15s; }
  .result-btn.pass.active { background: var(--pass); border-color: var(--pass); color: #000; }
  .result-btn.fail.active { background: var(--fail); border-color: var(--fail); color: #fff; }
  .result-btn.na.active { background: var(--muted); border-color: var(--muted); color: #fff; }

  .result-badge { display: inline-flex; align-items: center; gap: 5px; padding: 4px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; }
  .result-badge.PASS { background: rgba(34,197,94,0.15); color: var(--pass); border: 1px solid rgba(34,197,94,0.3); }
  .result-badge.FAIL { background: rgba(239,68,68,0.15); color: var(--fail); border: 1px solid rgba(239,68,68,0.3); }
  .result-badge.PENDING { background: rgba(107,114,128,0.15); color: var(--muted); border: 1px solid rgba(107,114,128,0.3); }

  .add-btn { display: flex; align-items: center; justify-content: center; gap: 6px; width: 100%; padding: 10px; background: rgba(0,229,160,0.07); border: 1px dashed rgba(0,229,160,0.3); border-radius: var(--radius); color: var(--accent); font-family: var(--font-mono); font-size: 12px; cursor: pointer; margin-top: 4px; transition: background 0.15s; }
  .add-btn:active { background: rgba(0,229,160,0.15); }
  .del-btn { background: none; border: none; color: var(--danger); font-size: 16px; cursor: pointer; padding: 2px 4px; line-height: 1; }

  .setup-card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 14px; margin-bottom: 12px; }
  .setup-title { font-family: var(--font-head); font-size: 14px; font-weight: 700; margin-bottom: 12px; color: var(--accent); }

  .photo-upload { display: flex; align-items: center; justify-content: center; background: var(--surface2); border: 1px dashed var(--border); border-radius: 6px; padding: 14px; cursor: pointer; color: var(--muted); font-size: 12px; gap: 6px; }
  .photo-preview { width: 100%; border-radius: 6px; max-height: 180px; object-fit: cover; }

  .stats-row { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; margin-bottom: 12px; }
  @media (min-width: 520px) { .stats-row { grid-template-columns: repeat(4, 1fr); } }
  .stat-box { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 10px; text-align: center; }
  .stat-num { font-family: var(--font-head); font-size: 24px; font-weight: 800; }
  .stat-label { font-size: 10px; color: var(--muted); margin-top: 2px; }
  .stat-num.green { color: var(--pass); }
  .stat-num.red { color: var(--fail); }
  .stat-num.yellow { color: var(--warn); }

  .check-row { display: flex; align-items: center; gap: 8px; }
  input[type=checkbox] { width: 16px; height: 16px; accent-color: var(--accent); }

  .save-dot { width: 6px; height: 6px; border-radius: 50%; background: var(--muted); display: inline-block; }
  .save-dot.saved { background: var(--pass); }
  .save-dot.saving { background: var(--warn); animation: pulse 0.8s infinite; }
  @keyframes pulse { 0%,100% { opacity:1 } 50% { opacity:0.4 } }

  .picker-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.7); z-index: 200; display: flex; flex-direction: column; justify-content: flex-end; }
  .picker-sheet { background: var(--surface); border-radius: 16px 16px 0 0; padding: 16px; max-height: 80vh; overflow-y: auto; }
  .picker-title { font-family: var(--font-head); font-size: 16px; font-weight: 700; margin-bottom: 14px; }
  .proj-item { display: flex; align-items: center; justify-content: space-between; background: var(--surface2); border: 1px solid var(--border); border-radius: 8px; padding: 12px; margin-bottom: 8px; cursor: pointer; transition: border-color 0.15s; }
  .proj-item:hover { border-color: var(--accent); }
  .proj-item.active { border-color: var(--accent); }
  .proj-name { font-family: var(--font-head); font-size: 13px; font-weight: 700; }
  .proj-meta { font-size: 10px; color: var(--muted); margin-top: 2px; }
  .proj-del { background: none; border: none; color: var(--danger); font-size: 16px; cursor: pointer; padding: 4px 6px; }
  .new-proj-btn { display: flex; align-items: center; justify-content: center; gap: 6px; width: 100%; padding: 12px; background: rgba(0,229,160,0.08); border: 1px dashed rgba(0,229,160,0.4); border-radius: 8px; color: var(--accent); font-family: var(--font-mono); font-size: 13px; cursor: pointer; margin-top: 4px; }
  .picker-close { display: flex; align-items: center; justify-content: center; width: 100%; padding: 12px; background: none; border: none; color: var(--muted); font-family: var(--font-mono); font-size: 13px; cursor: pointer; margin-top: 4px; }

  .toast { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); background: var(--accent); color: #000; padding: 10px 20px; border-radius: 8px; font-weight: 500; font-size: 13px; z-index: 999; animation: fadeup 0.3s ease; pointer-events: none; }
  @keyframes fadeup { from { opacity: 0; transform: translateX(-50%) translateY(10px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }
`;

// ── Sub-components ────────────────────────────────────────────────────────────
function Field({ label, children }) {
  return <div className="field"><label>{label}</label>{children}</div>;
}

function DoorCard({ door, onChange, onDelete }) {
  const [open, setOpen] = useState(true);
  const f = (k) => (e) => onChange({ ...door, [k]: e.target.value });
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onChange({ ...door, photoDataUrl: reader.result });
    reader.readAsDataURL(file);
  };

  return (
    <div className={`card ${open ? "" : "card-collapsed"}`}>
      <div className="card-header" onClick={() => setOpen(o => !o)}>
        <div className="card-title">
          <span className="card-tag">{door.doorRef || "DOOR"}</span>
          {door.location || "Door record"}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ color: "var(--muted)", fontSize: 12 }}>{open ? "▲" : "▼"}</span>
          <button className="del-btn" onClick={e => { e.stopPropagation(); onDelete(); }}>✕</button>
        </div>
      </div>
      <div className="card-body">
        <div className="field-row">
          <Field label="Door Reference"><input value={door.doorRef} onChange={f("doorRef")} placeholder="CR-##.##" /></Field>
          <Field label="Location"><input value={door.location} onChange={f("location")} placeholder="Main Entrance" /></Field>
        </div>
        <div className="field-row">
          <Field label="Reader In"><input value={door.readerIn} onChange={f("readerIn")} placeholder="e.g. SIGNO 20" /></Field>
          <Field label="Reader Out"><input value={door.readerOut} onChange={f("readerOut")} placeholder="e.g. SIGNO 20" /></Field>
        </div>
        <div className="field-row">
          <Field label="Lock"><input value={door.lock} onChange={f("lock")} placeholder="e.g. EL560" /></Field>
          <Field label="Monitoring"><input value={door.monitoring} onChange={f("monitoring")} placeholder="e.g. Yes / No" /></Field>
        </div>
        <div className="field-row">
          <Field label="Exit Button"><input value={door.exitButton} onChange={f("exitButton")} placeholder="e.g. NT200" /></Field>
          <Field label="Break Glass Unit"><input value={door.breakGlassUnit} onChange={f("breakGlassUnit")} placeholder="e.g. Monitored Triple Pole" /></Field>
        </div>
        <Field label="Controller Port & Location"><input value={door.controller} onChange={f("controller")} placeholder="Panel 1 - Comms Room"/></Field>
        <div className="field-row">
          <Field label="Opportunity Ref"><input value={door.opportunityRef} onChange={f("opportunityRef")} placeholder="Opportunity reference" /></Field>
          <Field label="Installation Date"><input type="date" value={door.installDate} onChange={f("installDate")} /></Field>
        </div>
        <div className="field-row">
          <Field label="Door Forced"><select value={door.doorForced} onChange={f("doorForced")}>
            <option value="">— Select —</option>
            {DOOR_TEST_OPTIONS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
          </select></Field>
          <Field label="Door Held"><select value={door.doorHeld} onChange={f("doorHeld")}>
            <option value="">— Select —</option>
            {DOOR_TEST_OPTIONS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
          </select></Field>
        </div>
        <Field label="Photo Description"><input value={door.photoDescription} onChange={f("photoDescription")} placeholder="Photo details" /></Field>
        {door.photoDataUrl ? (
          <img src={door.photoDataUrl} alt="door" className="photo-preview" onClick={() => onChange({ ...door, photoDataUrl: "" })} />
        ) : (
          <label className="photo-upload">
            <input type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={handleFile} />📷 Tap to take / upload door photo
          </label>
        )}
        <Field label="Notes"><textarea value={door.notes} onChange={f("notes")} /></Field>
      </div>
    </div>
  );
}

function CctvCard({ camera, onChange, onDelete }) {
  const [open, setOpen] = useState(true);
  const f = (k) => (e) => onChange({ ...camera, [k]: e.target.value });
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onChange({ ...camera, photoDataUrl: reader.result });
    reader.readAsDataURL(file);
  };

  return (
    <div className={`card ${open ? "" : "card-collapsed"}`}>
      <div className="card-header" onClick={() => setOpen(o => !o)}>
        <div className="card-title">
          <span className="card-tag">{camera.cameraNumber || "CAM"}</span>
          {camera.location || "CCTV record"}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ color: "var(--muted)", fontSize: 12 }}>{open ? "▲" : "▼"}</span>
          <button className="del-btn" onClick={e => { e.stopPropagation(); onDelete(); }}>✕</button>
        </div>
      </div>
      <div className="card-body">
        <div className="field-row">
          <Field label="Camera Number"><input value={camera.cameraNumber} onChange={f("cameraNumber")} placeholder="TV-##.##" /></Field>
          <Field label="Location"><input value={camera.location} onChange={f("location")} placeholder="Main Entrance" /></Field>
        </div>
        <div className="field-row">
          <Field label="Camera Make"><input value={camera.camMake} onChange={f("camMake")} placeholder="e.g. Avigilon" /></Field>
          <Field label="Camera Model"><input value={camera.camModel} onChange={f("camModel")} placeholder="e.g. 8.0C-H6A-D01-IR" /></Field>
        </div>
        <div className="field-row">
          <Field label="Supply Volts"><input value={camera.supplyVolts} onChange={f("supplyVolts")} placeholder="e.g. PoE+" /></Field>
          <Field label="Record Quality"><input value={camera.recordQuality} onChange={f("recordQuality")} placeholder="e.g. 1080p" /></Field>
        </div>
        <div className="field-row">
          <Field label="Playback Quality"><input value={camera.playbackQuality} onChange={f("playbackQuality")} placeholder="e.g. 720p" /></Field>
          <Field label="Record FPS"><input value={camera.recordFPS} onChange={f("recordFPS")} placeholder="e.g. 25" /></Field>
        </div>
        <div className="field-row">
          <Field label="MAC Address"><input value={camera.macAddress} onChange={f("macAddress")} placeholder="AA:BB:CC:DD:EE:FF" /></Field>
          <Field label="IP Address"><input value={camera.ipAddress} onChange={f("ipAddress")} placeholder="192.168.x.x" /></Field>
        </div>
        <div className="field-row">
          <Field label="Installation Date"><input type="date" value={camera.installDate} onChange={f("installDate")} /></Field>
          <Field label="Opportunity Number"><input value={camera.opportunityNumber} onChange={f("opportunityNumber")} placeholder="Opportunity reference" /></Field>
        </div>
        <Field label="Photo Description"><input value={camera.photoDescription} onChange={f("photoDescription")} placeholder="Photo details" /></Field>
        {camera.photoDataUrl ? (
          <img src={camera.photoDataUrl} alt="camera" className="photo-preview" onClick={() => onChange({ ...camera, photoDataUrl: "" })} />
        ) : (
          <label className="photo-upload">
            <input type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={handleFile} />📷 Tap to take / upload camera photo
          </label>
        )}
        <Field label="Notes"><textarea value={camera.notes} onChange={f("notes")} /></Field>
      </div>
    </div>
  );
}

// ── Project Picker ────────────────────────────────────────────────────────────
function ProjectPicker({ currentId, onSelect, onNew, onClose }) {
  const [index, setIndex] = useState(() => loadIndex().sort((a, b) => b.savedAt - a.savedAt));

  const handleDelete = (e, id) => {
    e.stopPropagation();
    if (!window.confirm("Delete this project? This cannot be undone.")) return;
    deleteProject(id);
    setIndex(idx => idx.filter(p => p.id !== id));
  };

  return (
    <div className="picker-overlay" onClick={onClose}>
      <div className="picker-sheet" onClick={e => e.stopPropagation()}>
        <div className="picker-title">Saved Projects</div>
        {index.length === 0 && <div style={{ color: "var(--muted)", fontSize: 12, textAlign: "center", padding: 16 }}>No saved projects yet</div>}
        {index.map(p => (
          <div key={p.id} className={`proj-item ${p.id === currentId ? "active" : ""}`} onClick={() => onSelect(p.id)}>
            <div>
              <div className="proj-name">{p.name || "Untitled Project"}</div>
              <div className="proj-meta">Saved {new Date(p.savedAt).toLocaleString()}</div>
            </div>
            <button className="proj-del" onClick={e => handleDelete(e, p.id)}>✕</button>
          </div>
        ))}
        <button className="new-proj-btn" onClick={onNew}>+ New Project</button>
        <button className="picker-close" onClick={onClose}>Cancel</button>
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
const blankProjectData = () => ({
  project: { name: "", address: "", client: "", contractNo: "", engineers: "", startDate: new Date().toISOString().slice(0, 10), completionDate: "", systemDescription: "" },
  doors: [blankDoor()],
  cctvs: [blankCCTV()],
});

export default function App() {
  const [tab, setTab] = useState("project");
  const [toast, setToast] = useState(null);
  const [saveStatus, setSaveStatus] = useState("saved");
  const [showPicker, setShowPicker] = useState(false);
  const [projectId, setProjectId] = useState(() => uid());

  const [project, setProject] = useState(() => {
    const index = loadIndex();
    if (index.length > 0) {
      const latest = index.sort((a, b) => b.savedAt - a.savedAt)[0];
      const data = loadProject(latest.id);
      if (data) return data.project || blankProjectData().project;
    }
    return blankProjectData().project;
  });

  const [doors, setDoors] = useState(() => {
    const index = loadIndex();
    if (index.length > 0) {
      const latest = index.sort((a, b) => b.savedAt - a.savedAt)[0];
      const data = loadProject(latest.id);
      if (data) return data.doors?.length ? data.doors : [blankDoor()];
    }
    return [blankDoor()];
  });

  const [cctvs, setCctvs] = useState(() => {
    const index = loadIndex();
    if (index.length > 0) {
      const latest = index.sort((a, b) => b.savedAt - a.savedAt)[0];
      const data = loadProject(latest.id);
      if (data) return data.cctvs?.length ? data.cctvs : [blankCCTV()];
    }
    return [blankCCTV()];
  });

  // Also restore the projectId from the latest save on mount
  useEffect(() => {
    const index = loadIndex();
    if (index.length > 0) {
      const latest = index.sort((a, b) => b.savedAt - a.savedAt)[0];
      setProjectId(latest.id);
    }
  }, []);

  const saveTimer = useRef(null);
  const isLoading = useRef(false);

  const triggerSave = useCallback((pid, proj, drs, ccs) => {
    if (isLoading.current) return;
    setSaveStatus("saving");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      saveProject(pid, { project: proj, doors: drs, cctvs: ccs });
      setSaveStatus("saved");
    }, 800);
  }, []);

  const setProjectAndSave = (updater) => setProject(prev => { const next = typeof updater === "function" ? updater(prev) : updater; triggerSave(projectId, next, doors, cctvs); return next; });
  const setDoorsAndSave = (updater) => setDoors(prev => { const next = typeof updater === "function" ? updater(prev) : updater; triggerSave(projectId, project, next, cctvs); return next; });
  const setCctvsAndSave = (updater) => setCctvs(prev => { const next = typeof updater === "function" ? updater(prev) : updater; triggerSave(projectId, project, doors, next); return next; });

  const pf = (k) => (e) => setProjectAndSave(p => ({ ...p, [k]: e.target.value }));

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 2500); };

  const handleExport = () => {
    try { exportToExcel(project, doors, cctvs); showToast("✓ Excel file exported!"); }
    catch (e) { showToast("Export failed – try again"); }
  };

  const handleNewProject = () => {
    isLoading.current = false;
    const newId = uid();
    const blank = blankProjectData();
    setProjectId(newId);
    setProject(blank.project);
    setDoors(blank.doors);
    setCctvs(blank.cctvs);
    setShowPicker(false);
    setTab("project");
    showToast("New project created");
  };

  const handleSelectProject = (id) => {
    isLoading.current = true;
    const data = loadProject(id);
    if (data) {
      setProjectId(id);
      setProject(data.project || blankProjectData().project);
      setDoors(data.doors?.length ? data.doors : [blankDoor()]);
      setCctvs(data.cctvs?.length ? data.cctvs : [blankCCTV()]);
      setTab("project");
      showToast("Project loaded");
    }
    setShowPicker(false);
    setTimeout(() => { isLoading.current = false; }, 100);
  };

  const TABS = [
    { id: "project", label: "Project" },
    { id: "doors", label: "Doors", count: doors.length },
    { id: "cctvs", label: "Cameras", count: cctvs.length },
  ];

  return (
    <>
      <style>{css}</style>
      <div className="app">
        <div className="header">
          <div className="header-top">
            <div style={{ flex: 1, minWidth: 0 }}>
              <div className="header-title">
                ISD<span>TECH</span>
                <button className="proj-switch-btn" onClick={() => setShowPicker(true)}>⇄ Projects</button>
              </div>
              <div className="project-name">
                <span className={`save-dot ${saveStatus}`} />
                {saveStatus === "saving" ? "Saving…" : project.name || "New Project"}
              </div>
            </div>
            <button className="export-btn" onClick={handleExport}>⬇ XLSX</button>
          </div>
        </div>

        <div className="nav">
          {TABS.map(t => (
            <button key={t.id} className={`nav-tab ${tab === t.id ? "active" : ""}`} onClick={() => setTab(t.id)}>
              {t.label}{t.count !== undefined && <span className="nav-count">{t.count}</span>}
            </button>
          ))}
        </div>

        <div className="content">
          {tab === "project" && (
            <>
              <div className="stats-row">
                <div className="stat-box"><div className="stat-num">{doors.length}</div><div className="stat-label">Doors</div></div>
                <div className="stat-box"><div className="stat-num">{cctvs.length}</div><div className="stat-label">Cameras</div></div>
              </div>
              <div className="setup-card">
                <div className="setup-title">Project Details</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  <Field label="Project Name"><input value={project.name} onChange={pf("name")} placeholder="e.g. CCTV Upgrade Phase 1" /></Field>
                  <Field label="Site Address"><input value={project.address} onChange={pf("address")} placeholder="Full site address" /></Field>
                  <div className="field-row">
                    <Field label="Client"><input value={project.client} onChange={pf("client")} /></Field>
                    <Field label="Opportunity No."><input value={project.oppNo} onChange={pf("oppNo")} /></Field>
                  </div>
                  <Field label="Engineer(s)"><input value={project.engineers} onChange={pf("engineers")} placeholder="Names of commissioning engineers" /></Field>
                  <div className="field-row">
                    <Field label="Start Date"><input type="date" value={project.startDate} onChange={pf("startDate")} /></Field>
                    <Field label="Completion Date"><input type="date" value={project.completionDate} onChange={pf("completionDate")} /></Field>
                  </div>
                  <Field label="System Description"><textarea value={project.systemDescription} onChange={pf("systemDescription")} placeholder="Brief description of system installed…" /></Field>
                </div>
              </div>
            </>
          )}

          {tab === "doors" && (
            <>
              {doors.map((d, i) => (
                <DoorCard key={d.id} door={d}
                  onChange={u => setDoorsAndSave(ds => ds.map((x, j) => j === i ? u : x))}
                  onDelete={() => setDoorsAndSave(ds => ds.filter((_, j) => j !== i))} />
              ))}
              <button className="add-btn" onClick={() => setDoorsAndSave(ds => [...ds, blankDoor()])}>+ Add Door</button>
            </>
          )}

          {tab === "cctvs" && (
            <>
              {cctvs.map((c, i) => (
                <CctvCard key={c.id} camera={c}
                  onChange={u => setCctvsAndSave(cs => cs.map((x, j) => j === i ? u : x))}
                  onDelete={() => setCctvsAndSave(cs => cs.filter((_, j) => j !== i))} />
              ))}
              <button className="add-btn" onClick={() => setCctvsAndSave(cs => [...cs, blankCCTV()])}>+ Add Camera</button>
            </>
          )}
        </div>
      </div>

      {showPicker && (
        <ProjectPicker currentId={projectId} onSelect={handleSelectProject} onNew={handleNewProject} onClose={() => setShowPicker(false)} />
      )}

      {toast && <div className="toast">{toast}</div>}
    </>
  );
}
