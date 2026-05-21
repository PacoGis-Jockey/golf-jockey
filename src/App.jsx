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
const STORAGE_KEY = "golf_scorecards";

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

async function loadCards() {
  try {
    const r = await window.storage.get(STORAGE_KEY);
    return r ? JSON.parse(r.value) : [];
  } catch { return []; }
}

async function saveCards(cards) {
  await window.storage.set(STORAGE_KEY, JSON.stringify(cards));
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
    padding: 24px 20px 20px;
    text-align: center;
    position: relative;
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
    margin-bottom: 10px;
  }
  .header h1 {
    font-family: 'Playfair Display', serif;
    color: var(--cream);
    font-size: 26px;
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
    const existing = await loadCards();
    const newEntries = players
      .filter((p, i) => p.name.trim() && Object.values(scores[i]).some(v => v !== ""))
      .map((p, i) => {
        const pi = players.indexOf(p);
        const total = calcTotal(scores[pi]);
        const neto = p.hcp !== null ? total - p.hcp : null;
        return { date, time, name: p.name.trim(), hcp: p.hcp, scores: scores[pi], total, neto };
      });
    await saveCards([...existing, ...newEntries]);
    setSubmitting(false);
    setScreen("success");
  }

  function resetForm() {
    setPlayers([{ name: "", hcp: null }]);
    setScores([emptyScores()]);
    setScreen("form");
  }

  async function openAdmin() {
    const c = await loadCards();
    setCards(c);
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

  async function clearAll() {
    if (!window.confirm("¿Borrar todas las tarjetas? Esta acción no se puede deshacer.")) return;
    await saveCards([]);
    setCards([]);
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
                  <div className="icon">📋</div>
                  <p>No hay tarjetas registradas todavía.</p>
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
