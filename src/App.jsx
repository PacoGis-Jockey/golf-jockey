import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ── DATA ────────────────────────────────────────────────────────────────────
const HOLES = [
  { num: 1, yards: 134, par: 4, hp: 6,  half: "ida" },
  { num: 2, yards: 152, par: 4, hp: 2,  half: "ida" },
  { num: 3, yards: 75,  par: 3, hp: 8,  half: "ida" },
  { num: 4, yards: 137, par: 4, hp: 4,  half: "ida" },
  { num: 5, yards: 117, par: 4, hp: 1,  half: "vuelta" },
  { num: 6, yards: 113, par: 4, hp: 3,  half: "vuelta" },
  { num: 7, yards: 65,  par: 3, hp: 7,  half: "vuelta" },
  { num: 8, yards: 108, par: 4, hp: 5,  half: "vuelta" },
];

const PLAYERS = [
  { name: "Reynal O'Connor, Alí", hcp: -1 },
  { name: "López Olaciregui, Salvador", hcp: -1 },
  { name: "Gismondi, Simón", hcp: 2 },
  { name: "Cúneo, Beltrán", hcp: 2 },
  { name: "López Olaciregui, Bautista", hcp: 3 },
  { name: "Pisarenko, Belisario", hcp: 3 },
  { name: "Hume Navarro, Alejo", hcp: 3 },
  { name: "Clariá, Marcos", hcp: 3 },
  { name: "Fagalde, Iñaki", hcp: 4 },
  { name: "Pereira, Valentín", hcp: 5 },
  { name: "Torralva, Baltazar", hcp: 6 },
  { name: "Casado, Alfonso", hcp: 6 },
  { name: "Díaz Valdéz, Justo", hcp: 6 },
  { name: "Luchía Puig, Dimas", hcp: 7 },
  { name: "Grau, Felipe", hcp: 7 },
  { name: "Campana, Agustín", hcp: 7 },
  { name: "Crotto, Ramiro", hcp: 8 },
  { name: "Mackinlay, Félix", hcp: 8 },
  { name: "De Achával, Miguel", hcp: 10 },
  { name: "O Farrell, Ambar", hcp: 10 },
  { name: "Dianda, Nicanor", hcp: 11 },
  { name: "Malbrán, Juan Jacinto", hcp: 11 },
  { name: "Lernoud, Tomás", hcp: 12 },
  { name: "Gutierrez Cantilo, Iñaki", hcp: 12 },
  { name: "Plate, Fermín", hcp: 12 },
  { name: "Blousson, Simón", hcp: 12 },
  { name: "Goldaracena, Jaime", hcp: 13 },
  { name: "Bisogno, Marcos", hcp: 13 },
  { name: "Hughes, Constantino", hcp: 14 },
  { name: "Navarra, Nicolás", hcp: 14 },
  { name: "Garzón, Manuel", hcp: 14 },
  { name: "Herrera, Mateo", hcp: 14 },
  { name: "Berri, Mateo", hcp: 15 },
  { name: "Nellen, Félix", hcp: 15 },
  { name: "Daza, Paulina", hcp: 16 },
  { name: "Bisogno, Guillermo", hcp: 16 },
  { name: "Fagalde, Beltrán", hcp: 16 },
  { name: "Marino, Rosa", hcp: 17 },
  { name: "Nicastro Morita, Camilo", hcp: 17 },
  { name: "Ruete, Alfonso", hcp: 17 },
  { name: "Gancedo, Alejandro", hcp: 18 },
  { name: "Almeida, Sofía", hcp: 18 },
  { name: "Ajates, Olivia", hcp: 18 },
  { name: "Ajates, Lucas", hcp: 18 },
  { name: "Navarra, Ignacio", hcp: 20 },
  { name: "Mackinlay, Manuel", hcp: 22 },
  { name: "De Soldati, Clara", hcp: 22 },
  { name: "Diez, Tomás", hcp: 25 },
];

const ADMIN_PASSWORD = "admin2026";
const LOGO = "data:image/jpeg;base64,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";
const SHEETS_URL = "https://script.google.com/macros/s/AKfycbxsm98tiyIr5qQ15Djb51W_LKJAvS3Y8mhCgVMW1iYm5k8DNOmkgmFHqtTsk-MYWRNl/exec";

// ── HELPERS ─────────────────────────────────────────────────────────────────
function findPlayer(query) {
  if (!query || query.length < 2) return [];
  const q = query.toLowerCase();
  return PLAYERS.filter(p => p.name.toLowerCase().includes(q)).slice(0, 6);
}

function emptyScores() {
  return Object.fromEntries(HOLES.map(h => [h.num, ""]));
}

function calcTotal(scores) {
  return Object.values(scores).reduce((s, v) => s + (parseInt(v) || 0), 0);
}

function now() {
  const d = new Date();
  return {
    date: d.toLocaleDateString("es-AR"),
    time: d.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" }),
  };
}

async function sendToSheets(entry) {
  const params = new URLSearchParams({
    date: entry.date,
    time: entry.time,
    name: entry.name,
    h1: entry.scores[1] || "",
    h2: entry.scores[2] || "",
    h3: entry.scores[3] || "",
    h4: entry.scores[4] || "",
    h5: entry.scores[5] || "",
    h6: entry.scores[6] || "",
    h7: entry.scores[7] || "",
    h8: entry.scores[8] || "",
    total: entry.total,
    hcp: entry.hcp !== null ? entry.hcp : "",
    neto: entry.neto !== null ? entry.neto : "",
  });
  await fetch(SHEETS_URL + "?" + params.toString(), {
    method: "GET",
    mode: "no-cors",
  });
}

// ── STYLES ──────────────────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=DM+Sans:wght@300;400;500;600&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --green:   #1a4a2e;
    --green2:  #2d6b45;
    --gold:    #c9a84c;
    --cream:   #f5f0e8;
    --paper:   #faf8f4;
    --ink:     #1a1a1a;
    --muted:   #7a7060;
    --border:  #d4cab8;
    --red:     #b84040;
    --radius:  12px;
  }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--cream);
    color: var(--ink);
    min-height: 100vh;
  }

  .app { max-width: 480px; margin: 0 auto; min-height: 100vh; }

  /* ── HEADER ── */
  .header {
    background: var(--green);
    padding: 20px 20px 18px;
    text-align: center;
    position: relative;
  }
  .header-logo {
    width: 64px;
    height: 64px;
    object-fit: contain;
    margin-bottom: 10px;
    object-fit: contain;
  }
  .header-badge {
    display: inline-block;
    border: 1.5px solid var(--gold);
    color: var(--gold);
    font-size: 10px;
    letter-spacing: 3px;
    text-transform: uppercase;
    padding: 4px 12px;
    border-radius: 20px;
    margin-bottom: 8px;
  }
  .header h1 {
    font-family: 'Playfair Display', serif;
    color: var(--cream);
    font-size: 24px;
    line-height: 1.15;
  }
  .header p {
    color: rgba(245,240,232,0.65);
    font-size: 13px;
    margin-top: 6px;
  }

  /* ── CONTENT ── */
  .content { padding: 20px 16px 40px; }

  /* ── CARD ── */
  .card {
    background: var(--paper);
    border-radius: var(--radius);
    border: 1px solid var(--border);
    margin-bottom: 16px;
    overflow: hidden;
  }
  .card-header {
    background: var(--green);
    padding: 10px 16px;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .card-header span {
    font-family: 'Playfair Display', serif;
    color: var(--cream);
    font-size: 15px;
  }
  .card-header .badge {
    margin-left: auto;
    background: var(--gold);
    color: var(--green);
    font-size: 10px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 10px;
    letter-spacing: 1px;
    text-transform: uppercase;
  }
  .card-body { padding: 16px; }

  /* ── PLAYER ROW ── */
  .player-row {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 12px;
  }
  .player-num {
    width: 24px;
    height: 24px;
    border-radius: 50%;
    background: var(--green);
    color: var(--cream);
    font-size: 12px;
    font-weight: 600;
    display: flex;align-items:center;justify-content:center;
    flex-shrink: 0;
  }
  .input-wrap { position: relative; flex: 1; }
  .input-wrap input {
    width: 100%;
    padding: 10px 12px;
    border: 1.5px solid var(--border);
    border-radius: 8px;
    font-family: 'DM Sans', sans-serif;
    font-size: 14px;
    background: white;
    color: var(--ink);
    outline: none;
    transition: border-color .2s;
  }
  .input-wrap input:focus { border-color: var(--green2); }
  .input-wrap input.has-hcp { border-color: var(--gold); background: #fffbf0; }

  .hcp-chip {
    position: absolute;
    right: 8px; top: 50%;
    transform: translateY(-50%);
    background: var(--gold);
    color: var(--green);
    font-size: 10px;
    font-weight: 700;
    padding: 2px 7px;
    border-radius: 10px;
  }

  .autocomplete {
    position: absolute;
    top: calc(100% + 4px);
    left: 0; right: 0;
    background: white;
    border: 1.5px solid var(--green2);
    border-radius: 8px;
    z-index: 100;
    box-shadow: 0 4px 16px rgba(0,0,0,.12);
    overflow: hidden;
  }
  .autocomplete li {
    padding: 10px 12px;
    font-size: 13px;
    cursor: pointer;
    border-bottom: 1px solid var(--border);
    list-style: none;
    transition: background .15s;
  }
  .autocomplete li:last-child { border-bottom: none; }
  .autocomplete li:hover { background: var(--cream); }
  .autocomplete li .ac-hcp {
    float: right;
    color: var(--gold);
    font-weight: 600;
    font-size: 12px;
  }

  .btn-remove {
    width: 28px; height: 28px;
    border-radius: 50%;
    border: 1.5px solid var(--border);
    background: white;
    color: var(--muted);
    font-size: 16px;
    cursor: pointer;
    display: flex;align-items:center;justify-content:center;
    flex-shrink: 0;
    transition: all .2s;
  }
  .btn-remove:hover { background: #fee; border-color: var(--red); color: var(--red); }

  .btn-add {
    width: 100%;
    padding: 9px;
    border: 1.5px dashed var(--border);
    border-radius: 8px;
    background: transparent;
    color: var(--muted);
    font-family: 'DM Sans', sans-serif;
    font-size: 13px;
    cursor: pointer;
    transition: all .2s;
    margin-top: 4px;
  }
  .btn-add:hover { border-color: var(--green2); color: var(--green2); background: rgba(45,107,69,.04); }

  /* ── SCORECARD TABLE ── */
  .section-label {
    font-size: 10px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--gold);
    font-weight: 600;
    margin-bottom: 10px;
    padding-left: 2px;
  }

  .score-table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }
  .score-table th {
    font-size: 11px;
    color: var(--muted);
    font-weight: 500;
    padding: 6px 4px;
    text-align: center;
    border-bottom: 1px solid var(--border);
  }
  .score-table th.left { text-align: left; }
  .score-table td {
    padding: 4px 3px;
    text-align: center;
    vertical-align: middle;
  }
  .hole-info { font-size: 11px; color: var(--muted); }
  .hole-num {
    width: 26px; height: 26px;
    border-radius: 50%;
    background: var(--green);
    color: var(--cream);
    font-size: 12px;
    font-weight: 700;
    display: flex;align-items:center;justify-content:center;
    margin: 0 auto;
  }
  .score-input {
    width: 36px; height: 36px;
    border-radius: 8px;
    border: 1.5px solid var(--border);
    text-align: center;
    font-family: 'DM Sans', sans-serif;
    font-size: 16px;
    font-weight: 600;
    color: var(--ink);
    background: white;
    outline: none;
    transition: all .2s;
    -moz-appearance: textfield;
  }
  .score-input::-webkit-outer-spin-button,
  .score-input::-webkit-inner-spin-button { -webkit-appearance: none; }
  .score-input:focus { border-color: var(--green2); background: #f0f8f3; }
  .score-input.filled { border-color: var(--green2); background: #f0f8f3; color: var(--green); font-weight: 700; }
  .score-input.max { border-color: var(--red); background: #fff0f0; color: var(--red); }

  .subtotal-row td { padding-top: 8px; border-top: 1.5px solid var(--border); }
  .subtotal-label { font-size: 11px; color: var(--muted); text-align: left; padding-left: 4px; }
  .subtotal-val { font-size: 14px; font-weight: 700; color: var(--green); }

  .divider {
    height: 1px;
    background: var(--border);
    margin: 12px 0;
    border: none;
  }

  /* ── TOTALS ── */
  .totals-box {
    background: var(--green);
    border-radius: 10px;
    padding: 14px 16px;
    margin-top: 12px;
  }
  .totals-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: var(--cream);
    margin-bottom: 6px;
  }
  .totals-row:last-child { margin-bottom: 0; }
  .totals-row span { font-size: 13px; opacity: .75; }
  .totals-row strong { font-size: 18px; font-family: 'Playfair Display', serif; }
  .totals-row.neto strong { color: var(--gold); }

  /* ── BUTTONS ── */
  .btn-primary {
    width: 100%;
    padding: 14px;
    background: var(--green);
    color: var(--cream);
    border: none;
    border-radius: 10px;
    font-family: 'DM Sans', sans-serif;
    font-size: 15px;
    font-weight: 600;
    cursor: pointer;
    transition: background .2s;
    letter-spacing: .3px;
  }
  .btn-primary:hover { background: var(--green2); }
  .btn-primary:disabled { opacity: .45; cursor: not-allowed; }

  .btn-secondary {
    width: 100%;
    padding: 12px;
    background: transparent;
    color: var(--green);
    border: 1.5px solid var(--green);
    border-radius: 10px;
    font-family: 'DM Sans', sans-serif;
    font-size: 14px;
    font-weight: 500;
    cursor: pointer;
    transition: all .2s;
    margin-top: 10px;
  }
  .btn-secondary:hover { background: var(--green); color: var(--cream); }

  /* ── SUCCESS ── */
  .success-screen {
    text-align: center;
    padding: 48px 24px;
  }
  .success-icon {
    width: 72px; height: 72px;
    border-radius: 50%;
    background: var(--green);
    display: flex;align-items:center;justify-content:center;
    margin: 0 auto 20px;
    font-size: 32px;
  }
  .success-screen h2 {
    font-family: 'Playfair Display', serif;
    font-size: 24px;
    margin-bottom: 8px;
  }
  .success-screen p { color: var(--muted); font-size: 14px; line-height: 1.5; }

  /* ── ADMIN ── */
  .admin-bar {
    text-align: center;
    padding: 12px;
    background: var(--paper);
    border-top: 1px solid var(--border);
  }
  .admin-link {
    font-size: 11px;
    color: var(--muted);
    cursor: pointer;
    text-decoration: underline;
    opacity: .5;
  }
  .admin-link:hover { opacity: 1; }

  .admin-screen { padding: 20px 16px 40px; }
  .admin-header {
    display: flex;align-items:center;gap:12px;
    margin-bottom: 20px;
  }
  .admin-title {
    font-family: 'Playfair Display', serif;
    font-size: 22px;
  }
  .btn-back {
    padding: 6px 12px;
    border: 1.5px solid var(--border);
    border-radius: 8px;
    background: white;
    font-family: 'DM Sans', sans-serif;
    font-size: 13px;
    cursor: pointer;
    color: var(--ink);
  }

  .stats-row {
    display: flex;gap:10px;margin-bottom:16px;
  }
  .stat-box {
    flex:1;background:var(--paper);border:1px solid var(--border);
    border-radius:10px;padding:12px;text-align:center;
  }
  .stat-box .num { font-family:'Playfair Display',serif;font-size:28px;color:var(--green); }
  .stat-box .lbl { font-size:11px;color:var(--muted);margin-top:2px; }

  .scorecard-entry {
    background: var(--paper);
    border: 1px solid var(--border);
    border-radius: 10px;
    margin-bottom: 12px;
    overflow: hidden;
  }
  .entry-header {
    background: var(--green);
    padding: 10px 14px;
    display: flex;justify-content:space-between;align-items:center;
  }
  .entry-header .name { color:var(--cream);font-weight:600;font-size:14px; }
  .entry-header .meta { color:rgba(245,240,232,.6);font-size:11px; }
  .entry-scores {
    padding:12px 14px;
    display:flex;flex-wrap:wrap;gap:6px;
  }
  .entry-hole {
    display:flex;flex-direction:column;align-items:center;
    min-width:36px;
  }
  .entry-hole .hn { font-size:9px;color:var(--muted); }
  .entry-hole .hv {
    width:32px;height:32px;border-radius:6px;
    background:var(--green);color:var(--cream);
    font-size:14px;font-weight:700;
    display:flex;align-items:center;justify-content:center;
    margin-top:2px;
  }
  .entry-hole .hv.empty { background:var(--border);color:var(--muted); }
  .entry-totals {
    padding:10px 14px;
    border-top:1px solid var(--border);
    display:flex;gap:16px;
  }
  .entry-totals .t { font-size:13px; }
  .entry-totals .t span { color:var(--muted);margin-right:4px; }
  .entry-totals .t strong { color:var(--green); }

  .btn-export {
    width:100%;padding:13px;
    background:var(--gold);color:var(--green);
    border:none;border-radius:10px;
    font-family:'DM Sans',sans-serif;font-size:15px;font-weight:700;
    cursor:pointer;margin-top:16px;transition:opacity .2s;
  }
  .btn-export:hover { opacity:.85; }

  .btn-clear {
    width:100%;padding:10px;
    background:transparent;color:var(--red);
    border:1.5px solid var(--red);border-radius:10px;
    font-family:'DM Sans',sans-serif;font-size:13px;
    cursor:pointer;margin-top:8px;
  }

  .empty-state {
    text-align:center;padding:40px 20px;color:var(--muted);
  }
  .empty-state .icon { font-size:40px;margin-bottom:12px; }

  .modal-bg {
    position:fixed;inset:0;background:rgba(0,0,0,.5);
    display:flex;align-items:center;justify-content:center;z-index:200;
    padding:20px;
  }
  .modal {
    background:white;border-radius:14px;padding:24px;
    width:100%;max-width:340px;
  }
  .modal h3 { font-family:'Playfair Display',serif;font-size:20px;margin-bottom:8px; }
  .modal p { font-size:13px;color:var(--muted);margin-bottom:16px; }
  .modal input {
    width:100%;padding:10px 12px;
    border:1.5px solid var(--border);border-radius:8px;
    font-family:'DM Sans',sans-serif;font-size:15px;
    outline:none;margin-bottom:12px;
  }
  .modal input:focus { border-color:var(--green2); }
  .modal-err { color:var(--red);font-size:12px;margin-bottom:8px;display:block; }
  .modal-btns { display:flex;gap:8px; }
  .modal-btns button { flex:1;padding:11px;border-radius:8px;font-family:'DM Sans',sans-serif;font-size:14px;cursor:pointer; }
  .modal-cancel { background:white;border:1.5px solid var(--border);color:var(--ink); }
  .modal-ok { background:var(--green);border:none;color:var(--cream);font-weight:600; }
`;

// ── COMPONENT: PlayerInput ───────────────────────────────────────────────────
function PlayerInput({ index, value, onChange, onRemove, canRemove }) {
  const [query, setQuery] = useState(value.name);
  const [suggestions, setSuggestions] = useState([]);
  const ref = useRef();

  useEffect(() => {
    function handler(e) {
      if (ref.current && !ref.current.contains(e.target)) setSuggestions([]);
    }
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  function handleChange(e) {
    const v = e.target.value;
    setQuery(v);
    onChange({ name: v, hcp: null });
    setSuggestions(findPlayer(v));
  }

  function selectPlayer(p) {
    setQuery(p.name);
    onChange({ name: p.name, hcp: p.hcp });
    setSuggestions([]);
  }

  return (
    <div className="player-row">
      <div className="player-num">{index + 1}</div>
      <div className="input-wrap" ref={ref}>
        <input
          type="text"
          placeholder={`Nombre del jugador ${index + 1}`}
          value={query}
          onChange={handleChange}
          className={value.hcp !== null ? "has-hcp" : ""}
          autoComplete="off"
        />
        {value.hcp !== null && (
          <span className="hcp-chip">HCP {value.hcp}</span>
        )}
        {suggestions.length > 0 && (
          <ul className="autocomplete">
            {suggestions.map(p => (
              <li key={p.name} onMouseDown={() => selectPlayer(p)}>
                {p.name}
                <span className="ac-hcp">HCP {p.hcp}</span>
              </li>
            ))}
          </ul>
        )}
      </div>
      {canRemove && (
        <button className="btn-remove" onClick={onRemove}>×</button>
      )}
    </div>
  );
}

// ── COMPONENT: ScoreTable ────────────────────────────────────────────────────
function ScoreTable({ holes, scores, players, onChange }) {
  const idaHoles = holes.filter(h => h.half === "ida");
  const vueltaHoles = holes.filter(h => h.half === "vuelta");

  function renderHalf(halfHoles, label) {
    const subtotals = players.map((_, pi) =>
      halfHoles.reduce((s, h) => s + (parseInt(scores[pi]?.[h.num]) || 0), 0)
    );

    return (
      <>
        <p className="section-label">{label}</p>
        <table className="score-table">
          <thead>
            <tr>
              <th className="left" style={{width:60}}>Hoyo</th>
              <th style={{width:40}}>Yds</th>
              <th style={{width:30}}>Par</th>
              {players.map((p, i) => (
                <th key={i}>{p.name ? p.name.split(",")[0].split(" ")[0] : `J${i+1}`}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {halfHoles.map(h => (
              <tr key={h.num}>
                <td>
                  <div className="hole-num">{h.num}</div>
                </td>
                <td><span className="hole-info">{h.yards}</span></td>
                <td><span className="hole-info">{h.par}</span></td>
                {players.map((_, pi) => {
                  const v = scores[pi]?.[h.num] ?? "";
                  const n = parseInt(v);
                  return (
                    <td key={pi}>
                      <input
                        className={`score-input${v !== "" ? (n === 8 ? " max" : " filled") : ""}`}
                        type="number"
                        min={1} max={8}
                        value={v}
                        onChange={e => {
                          let val = e.target.value;
                          if (val === "") { onChange(pi, h.num, ""); return; }
                          val = Math.max(1, Math.min(8, parseInt(val) || 1));
                          onChange(pi, h.num, val);
                        }}
                        placeholder="–"
                      />
                    </td>
                  );
                })}
              </tr>
            ))}
            <tr className="subtotal-row">
              <td colSpan={3} className="subtotal-label">Subtotal</td>
              {subtotals.map((s, i) => (
                <td key={i} className="subtotal-val">{s || "–"}</td>
              ))}
            </tr>
          </tbody>
        </table>
      </>
    );
  }

  return (
    <>
      {renderHalf(idaHoles, "Cuide su cancha · Hoyos 1–4")}
      <hr className="divider" />
      {renderHalf(vueltaHoles, "Arregle los piques · Hoyos 5–8")}
    </>
  );
}

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [screen, setScreen] = useState("form"); // form | success | admin
  const [players, setPlayers] = useState([{ name: "", hcp: null }]);
  const [scores, setScores] = useState([emptyScores()]);
  const [submitting, setSubmitting] = useState(false);
  const [cards, setCards] = useState([]);
  const [showAdminModal, setShowAdminModal] = useState(false);
  const [adminPwd, setAdminPwd] = useState("");
  const [adminErr, setAdminErr] = useState("");

  // ── handlers ──
  function addPlayer() {
    if (players.length >= 4) return;
    setPlayers(p => [...p, { name: "", hcp: null }]);
    setScores(s => [...s, emptyScores()]);
  }

  function removePlayer(i) {
    setPlayers(p => p.filter((_, idx) => idx !== i));
    setScores(s => s.filter((_, idx) => idx !== i));
  }

  function updatePlayer(i, val) {
    setPlayers(p => p.map((pl, idx) => idx === i ? val : pl));
  }

  function updateScore(pi, hole, val) {
    setScores(s => s.map((sc, idx) =>
      idx === pi ? { ...sc, [hole]: val } : sc
    ));
  }

  function canSubmit() {
    return players.some(p => p.name.trim()) &&
      players.some((_, i) => Object.values(scores[i]).some(v => v !== ""));
  }

  async function handleSubmit() {
    setSubmitting(true);
    const { date, time } = now();
    const newEntries = players
      .filter((p, i) => p.name.trim() && Object.values(scores[i]).some(v => v !== ""))
      .map((p, i) => {
        const pi = players.indexOf(p);
        const total = calcTotal(scores[pi]);
        const neto = p.hcp !== null ? total - p.hcp : null;
        return { date, time, name: p.name.trim(), hcp: p.hcp, scores: scores[pi], total, neto };
      });
    try {
      await Promise.all(newEntries.map(e => sendToSheets(e)));
    } catch(err) {
      console.error("Error sending to Sheets:", err);
    }
    setSubmitting(false);
    setScreen("success");
  }

  function resetForm() {
    setPlayers([{ name: "", hcp: null }]);
    setScores([emptyScores()]);
    setScreen("form");
  }

  function openAdmin() {
    setCards([]);
    setScreen("admin");
    setShowAdminModal(false);
    setAdminPwd("");
    setAdminErr("");
  }

  function tryAdmin() {
    if (adminPwd === ADMIN_PASSWORD) openAdmin();
    else setAdminErr("Contraseña incorrecta");
  }

  function exportExcel() {
    const rows = cards.map(c => ({
      Fecha: c.date,
      Hora: c.time,
      Nombre: c.name,
      H1: c.scores[1] || "",
      H2: c.scores[2] || "",
      H3: c.scores[3] || "",
      H4: c.scores[4] || "",
      H5: c.scores[5] || "",
      H6: c.scores[6] || "",
      H7: c.scores[7] || "",
      H8: c.scores[8] || "",
      Total: c.total || "",
      Handicap: c.hcp !== null ? c.hcp : "",
      "Total Neto": c.neto !== null ? c.neto : "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Scores");
    XLSX.writeFile(wb, `golf_scores_${new Date().toLocaleDateString("es-AR").replace(/\//g,"-")}.xlsx`);
  }

  function clearAll() {
    if (!window.confirm("Para borrar las tarjetas, hacelo directamente en Google Sheets.")) return;
  }

  // ── RENDER ──
  return (
    <>
      <style>{css}</style>
      <div className="app">

        {/* ── FORM SCREEN ── */}
        {screen === "form" && (
          <>
            <div className="header">
              <img
                className="header-logo"
                src={LOGO}
                alt="Jockey Club"
              />
              <div className="header-badge">Golf Infantil</div>
              <h1>Tarjeta de Score</h1>
              <p>Cancha de 8 hoyos · Par 30 · 901 yardas</p>
            </div>

            <div className="content">
              {/* Players */}
              <div className="card">
                <div className="card-header">
                  <span>Jugadores</span>
                  <span className="badge">{players.length}/4</span>
                </div>
                <div className="card-body">
                  {players.map((p, i) => (
                    <PlayerInput
                      key={i}
                      index={i}
                      value={p}
                      onChange={v => updatePlayer(i, v)}
                      onRemove={() => removePlayer(i)}
                      canRemove={players.length > 1}
                    />
                  ))}
                  {players.length < 4 && (
                    <button className="btn-add" onClick={addPlayer}>
                      + Agregar jugador
                    </button>
                  )}
                </div>
              </div>

              {/* Scores */}
              <div className="card">
                <div className="card-header">
                  <span>Scores</span>
                  <span className="badge">Máx. 8 golpes</span>
                </div>
                <div className="card-body">
                  <ScoreTable
                    holes={HOLES}
                    scores={scores}
                    players={players}
                    onChange={updateScore}
                  />

                  {/* Totals */}
                  {players.some((_, i) => calcTotal(scores[i]) > 0) && (
                    <div className="totals-box">
                      {players.map((p, i) => {
                        const t = calcTotal(scores[i]);
                        if (!t) return null;
                        const nombre = p.name ? p.name.split(",")[0].split(" ")[0] : `J${i+1}`;
                        return (
                          <div key={i}>
                            {players.length > 1 && (
                              <div style={{color:"rgba(245,240,232,.5)",fontSize:11,marginBottom:4}}>
                                {p.name || `Jugador ${i+1}`}
                              </div>
                            )}
                            <div className="totals-row">
                              <span>Total golpes</span>
                              <strong>{t}</strong>
                            </div>
                            {p.hcp !== null && (
                              <div className="totals-row neto">
                                <span>Total neto (HCP {p.hcp})</span>
                                <strong>{t - p.hcp}</strong>
                              </div>
                            )}
                            {i < players.length - 1 && <hr className="divider" style={{borderColor:"rgba(255,255,255,.1)",margin:"10px 0"}} />}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>

              <button
                className="btn-primary"
                onClick={handleSubmit}
                disabled={!canSubmit() || submitting}
              >
                {submitting ? "Enviando..." : "Enviar tarjeta ✓"}
              </button>
            </div>

            <div className="admin-bar">
              <span className="admin-link" onClick={() => setShowAdminModal(true)}>
                Acceso administrador
              </span>
            </div>
          </>
        )}

        {/* ── SUCCESS SCREEN ── */}
        {screen === "success" && (
          <>
            <div className="header">
              <img
                className="header-logo"
                src={LOGO}
                alt="Jockey Club"
              />
              <div className="header-badge">Golf Infantil</div>
              <h1>Tarjeta de Score</h1>
            </div>
            <div className="content">
              <div className="success-screen">
                <div className="success-icon">⛳</div>
                <h2>¡Tarjeta enviada!</h2>
                <p>Los scores quedaron guardados.<br />Gracias por cargar la tarjeta.</p>
              </div>
              <button className="btn-primary" onClick={resetForm}>
                Cargar nueva tarjeta
              </button>
            </div>
          </>
        )}

        {/* ── ADMIN SCREEN ── */}
        {screen === "admin" && (
          <>
            <div className="header">
              <img
                className="header-logo"
                src={LOGO}
                alt="Jockey Club"
              />
              <div className="header-badge">Administrador</div>
              <h1>Panel de Scores</h1>
              <p>{cards.length} tarjeta{cards.length !== 1 ? "s" : ""} registrada{cards.length !== 1 ? "s" : ""}</p>
            </div>
            <div className="admin-screen">
              <div style={{display:"flex",gap:8,marginBottom:16}}>
                <button className="btn-back" onClick={() => setScreen("form")}>← Volver</button>
              </div>

              {cards.length === 0 ? (
                <div className="empty-state">
                  <div className="icon">📊</div>
                  <p>Los scores se guardan directamente en Google Sheets.</p>
                  <p style={{marginTop:12,fontSize:12}}>Abrí tu planilla para ver todos los resultados.</p>
                </div>
              ) : (
                <>
                  <div className="stats-row">
                    <div className="stat-box">
                      <div className="num">{cards.length}</div>
                      <div className="lbl">Tarjetas</div>
                    </div>
                    <div className="stat-box">
                      <div className="num">
                        {Math.min(...cards.filter(c=>c.total>0).map(c=>c.total)) || "–"}
                      </div>
                      <div className="lbl">Mejor gross</div>
                    </div>
                    <div className="stat-box">
                      <div className="num">
                        {Math.round(cards.filter(c=>c.total>0).reduce((s,c)=>s+c.total,0)/cards.filter(c=>c.total>0).length) || "–"}
                      </div>
                      <div className="lbl">Promedio</div>
                    </div>
                  </div>

                  {cards.map((c, idx) => (
                    <div className="scorecard-entry" key={idx}>
                      <div className="entry-header">
                        <div>
                          <div className="name">{c.name}</div>
                          <div className="meta">{c.date} · {c.time}</div>
                        </div>
                        {c.hcp !== null && (
                          <div style={{background:"rgba(201,168,76,.2)",color:"#c9a84c",padding:"3px 10px",borderRadius:10,fontSize:12,fontWeight:600}}>
                            HCP {c.hcp}
                          </div>
                        )}
                      </div>
                      <div className="entry-scores">
                        {HOLES.map(h => {
                          const v = c.scores[h.num];
                          return (
                            <div className="entry-hole" key={h.num}>
                              <span className="hn">H{h.num}</span>
                              <div className={`hv${!v ? " empty" : ""}`}>{v || "–"}</div>
                            </div>
                          );
                        })}
                      </div>
                      <div className="entry-totals">
                        <div className="t"><span>Total:</span><strong>{c.total}</strong></div>
                        {c.neto !== null && (
                          <div className="t"><span>Neto:</span><strong>{c.neto}</strong></div>
                        )}
                      </div>
                    </div>
                  ))}
                </>
              )}

              <button className="btn-export" onClick={exportExcel}>
                ⬇ Exportar a Excel
              </button>
              <button className="btn-clear" onClick={clearAll}>
                Borrar todas las tarjetas
              </button>
            </div>
          </>
        )}

        {/* ── ADMIN MODAL ── */}
        {showAdminModal && (
          <div className="modal-bg" onClick={() => setShowAdminModal(false)}>
            <div className="modal" onClick={e => e.stopPropagation()}>
              <h3>Acceso admin</h3>
              <p>Ingresá la contraseña para ver y exportar todos los scores.</p>
              {adminErr && <span className="modal-err">{adminErr}</span>}
              <input
                type="password"
                placeholder="Contraseña"
                value={adminPwd}
                onChange={e => { setAdminPwd(e.target.value); setAdminErr(""); }}
                onKeyDown={e => e.key === "Enter" && tryAdmin()}
                autoFocus
              />
              <div className="modal-btns">
                <button className="modal-cancel" onClick={() => setShowAdminModal(false)}>Cancelar</button>
                <button className="modal-ok" onClick={tryAdmin}>Entrar</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </>
  );
}
