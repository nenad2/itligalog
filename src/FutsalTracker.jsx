// ─────────────────────────────────────────────────────────────────────────────
// FutsalTracker.jsx — Futsal Match Statistics App with Card System
// Dependencies: React (hooks), Tailwind CSS, jsPDF (loaded dynamically via CDN)
// ─────────────────────────────────────────────────────────────────────────────

import { useState, useEffect, useRef, useCallback } from "react";

// ── html2pdf loader ───────────────────────────────────────────────────────────
async function loadHtml2pdf() {
  if (window.html2pdf) return window.html2pdf;
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js";
    script.onload = () => resolve(window.html2pdf);
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

// ── SheetJS (xlsx) dynamic loader ────────────────────────────────────────────
async function getXLSX() {
  if (window.XLSX) return window.XLSX;
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src =
      "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error("Failed to load XLSX library"));
    document.head.appendChild(script);
  });
}

// ── localStorage persistence ──────────────────────────────────────────────────
const STORAGE_KEY = "futsal-tracker-match-v1";

function saveMatchToStorage(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    console.warn("Could not save match to localStorage", e);
  }
}

function loadMatchFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    console.warn("Could not load match from localStorage", e);
    return null;
  }
}

function clearMatchFromStorage() {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch (e) {
    console.warn("Could not clear match from localStorage", e);
  }
}

// ── Liga Excel file persistence ───────────────────────────────────────────────
const LIGA_STORAGE_KEY = "futsal-tracker-liga-v1";

function saveLigaToStorage(base64, teams) {
  try {
    localStorage.setItem(LIGA_STORAGE_KEY, JSON.stringify({ base64, teams }));
  } catch (e) {
    console.warn("Could not save liga file to localStorage", e);
  }
}

function loadLigaFromStorage() {
  try {
    const raw = localStorage.getItem(LIGA_STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    console.warn("Could not load liga file from localStorage", e);
    return null;
  }
}

// Parses every sheet in workbook → { teamName: [{ id, num, name }, ...] }
async function parseWorkbook(base64OrFile) {
  const XLSX = await getXLSX();
  let workbook;
  if (typeof base64OrFile === "string") {
    workbook = XLSX.read(base64OrFile, { type: "base64" });
  } else {
    const data = await base64OrFile.arrayBuffer();
    workbook = XLSX.read(data, { type: "array" });
  }
  const teams = {};
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const players = rows
      .map((row) => ({
        num: "",
        name: row[0] !== undefined ? String(row[0]).trim() : "",
      }))
      .filter((p) => p.name !== "")
      .map((p) => ({ id: crypto.randomUUID(), num: p.num, name: p.name }));
    if (players.length > 0) teams[sheetName] = players;
  }
  return teams;
}

// Convert File to base64 string
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── App config ────────────────────────────────────────────────────────────────
const SHEETS_WEBHOOK = "https://script.google.com/macros/s/AKfycbxJjpC29FS4PUCPZRPCN30g8xyB_ddANbm0h12xyvxlJha7bWaiPa27PznZLdzGCNen/exec";
const CONFIRM_FORM_URL = "https://forms.gle/eGHbncnZQ7nZWbaU7";

// ── Constants & helpers ───────────────────────────────────────────────────────
const DEFAULT_HALF_SECS = 20 * 60;
const ROSTER_SIZE = 15;
const HALF_LABELS = ["1. Poluvreme", "2. Poluvreme"];
const pad2 = (n) => String(Math.max(0, n)).padStart(2, "0");
const secsToDisplay = (s) => `${pad2(Math.floor(s / 60))}:${pad2(s % 60)}`;
const makePlayer = () => ({ id: crypto.randomUUID(), num: "", name: "", played: false });
const initRoster = (count = ROSTER_SIZE) => Array.from({ length: count }, makePlayer);

// ── Foul dots ─────────────────────────────────────────────────────────────────
function FoulIndicator({ count, max = 8 }) {
  return (
    <div className="flex items-center gap-1.5">
      {Array.from({ length: max }).map((_, i) => {
        const isBonus = i >= 5; // 6, 7, 8 su bonus (crvene)
        const filled = i < count;
        return (
          <span
            key={i}
            className={`block w-4 h-4 rounded-full border-2 transition-all duration-200 ${
              filled
                ? isBonus
                  ? "bg-red-500 border-red-400 shadow-[0_0_6px_rgba(239,68,68,0.8)]"
                  : "bg-yellow-400 border-yellow-300 shadow-[0_0_6px_rgba(250,204,21,0.7)]"
                : isBonus
                  ? "bg-transparent border-red-900/50"
                  : "bg-transparent border-gray-600"
            }`}
            title={isBonus ? `Bonus faul ${i + 1}` : `Faul ${i + 1}`}
          />
        );
      })}
    </div>
  );
}

// ── Card badge (mini, shown inline in player row) ─────────────────────────────
function CardBadges({ yellowCount, hasRed }) {
  if (yellowCount === 0 && !hasRed) return null;
  return (
    <div className="flex items-center gap-0.5 shrink-0">
      {Array.from({ length: Math.min(yellowCount, 2) }).map((_, i) => (
        <span
          key={i}
          className="inline-block w-3 h-4 rounded-sm bg-yellow-400 shadow-[0_0_4px_rgba(250,204,21,0.7)]"
          title="Yellow card"
        />
      ))}
      {hasRed && (
        <span
          className="inline-block w-3 h-4 rounded-sm bg-red-500 shadow-[0_0_6px_rgba(239,68,68,0.9)] ml-0.5"
          title="Red card"
        />
      )}
    </div>
  );
}

// ── Single player row ─────────────────────────────────────────────────────────
function PlayerRow({ player, goals, cards, onUpdate, onGoal, onCard }) {
  const playerGoalCount = goals.filter((g) => g.playerId === player.id).length;
  const playerCards = cards.filter((c) => c.playerId === player.id);
  const yellowCount = playerCards.filter((c) => c.type === "yellow").length;
  const hasRed = playerCards.some((c) => c.type === "red");
  const isSuspended = hasRed;

  return (
    <div
      className={`flex items-center gap-1.5 py-2 border-b border-gray-800/60 transition-all ${
        hasRed ? "bg-red-950/30"
        : yellowCount >= 1 ? "bg-yellow-950/20"
        : playerGoalCount > 0 ? "bg-green-950/20"
        : !player.played ? "opacity-40"
        : ""
      }`}
    >
      {/* Checkbox — igrao/nije igrao */}
      <button
        onClick={() => onUpdate({ ...player, played: !player.played })}
        title={player.played ? "Igrao — klikni da ukloniš" : "Nije igrao — klikni da označiš"}
        className={`shrink-0 w-5 h-5 rounded border-2 flex items-center justify-center transition-all active:scale-90 ${
          player.played
            ? "bg-emerald-500 border-emerald-400 text-white"
            : "bg-transparent border-gray-600 hover:border-gray-400"
        }`}
      >
        {player.played && <span className="text-[10px] font-black leading-none">✓</span>}
      </button>

      {/* Squad number */}
      <input
        type="text"
        maxLength={3}
        placeholder="#"
        value={player.num}
        onChange={(e) => onUpdate({ ...player, num: e.target.value })}
        className="w-10 shrink-0 bg-gray-800 border border-gray-700 text-yellow-300 font-mono text-xs text-center rounded px-1 py-1.5 focus:outline-none focus:border-yellow-500 focus:ring-1 focus:ring-yellow-500/30 transition-all"
      />

      {/* Name */}
      <input
        type="text"
        placeholder="Player name…"
        value={player.name}
        onChange={(e) => onUpdate({ ...player, name: e.target.value })}
        className="flex-1 min-w-0 bg-gray-800 border border-gray-700 text-white text-xs rounded px-2 py-1.5 focus:outline-none focus:border-yellow-500 focus:ring-1 focus:ring-yellow-500/30 transition-all placeholder-gray-600"
      />

      {/* Card badges (inline indicator) */}
      <CardBadges yellowCount={yellowCount} hasRed={hasRed} />

      {/* Goal button */}
      <button
        onClick={onGoal}
        disabled={isSuspended}
        title={isSuspended ? "Player suspended (red card)" : "Record goal"}
        className={`shrink-0 flex items-center gap-1 px-2 py-1.5 rounded text-xs font-bold transition-all active:scale-95 ${
          isSuspended
            ? "bg-gray-800 text-gray-600 cursor-not-allowed opacity-50"
            : "bg-emerald-700 hover:bg-emerald-500 text-white shadow-[0_2px_8px_rgba(16,185,129,0.25)]"
        }`}
      >
        <span>⚽</span>
        {playerGoalCount > 0 && (
          <span className="bg-emerald-900 text-emerald-300 rounded px-1 text-[10px] font-black">
            {playerGoalCount}
          </span>
        )}
      </button>

      {/* Yellow card button */}
      <button
        onClick={() => onCard(player, "yellow")}
        disabled={hasRed}
        title="Yellow card"
        className={`shrink-0 w-7 h-7 rounded flex items-center justify-center transition-all active:scale-95 ${
          hasRed
            ? "opacity-30 cursor-not-allowed bg-gray-800"
            : "bg-yellow-500/20 hover:bg-yellow-400/40 border border-yellow-500/50 hover:border-yellow-400"
        }`}
      >
        <span className="inline-block w-3 h-4 rounded-sm bg-yellow-400" />
      </button>

      {/* Red card button */}
      <button
        onClick={() => onCard(player, "red")}
        disabled={hasRed}
        title="Red card"
        className={`shrink-0 w-7 h-7 rounded flex items-center justify-center transition-all active:scale-95 ${
          hasRed
            ? "opacity-30 cursor-not-allowed bg-gray-800"
            : "bg-red-500/20 hover:bg-red-400/40 border border-red-500/50 hover:border-red-400"
        }`}
      >
        <span className="inline-block w-3 h-4 rounded-sm bg-red-500" />
      </button>
    </div>
  );
}

// ── Edit Event Modal ──────────────────────────────────────────────────────────
function EditEventModal({ event, players, allPlayers, type, onSave, onClose }) {
  const [clockMin, setClockMin] = useState(event.clockMin);
  const [playerId, setPlayerId] = useState(event.playerId);
  const activePlayers = allPlayers.filter((p) => p.num || p.name);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm">
      <div className="bg-gray-900 border border-gray-700 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 w-80 mx-4">
        <h2 className="text-white font-black text-sm uppercase tracking-[0.15em] text-center">
          {type === "goal" ? "✏ Uredi gol" : "✏ Uredi karton"}
        </h2>

        {/* Minute */}
        <div>
          <label className="block text-[10px] uppercase tracking-widest text-gray-500 mb-1">Minut</label>
          <input
            type="number" min={1} max={120}
            value={clockMin}
            onChange={(e) => setClockMin(Math.max(1, parseInt(e.target.value) || 1))}
            className="w-full text-center text-2xl font-mono font-black text-yellow-400 bg-gray-800 border border-gray-700 rounded-xl p-2 focus:outline-none focus:border-yellow-500"
          />
        </div>

        {/* Player picker */}
        <div>
          <label className="block text-[10px] uppercase tracking-widest text-gray-500 mb-1">Igrač</label>
          <select
            value={playerId}
            onChange={(e) => setPlayerId(e.target.value)}
            className="w-full bg-gray-800 border border-gray-700 text-white rounded-xl px-3 py-2 focus:outline-none focus:border-yellow-500 text-sm"
          >
            {activePlayers.length === 0 && (
              <option value={event.playerId}>{event.snapNum ? `#${event.snapNum} ${event.snapName}` : "Nepoznat igrač"}</option>
            )}
            {activePlayers.map((p) => (
              <option key={p.id} value={p.id}>
                {p.num ? `#${p.num} ` : ""}{p.name || "Bez imena"}
              </option>
            ))}
          </select>
        </div>

        <div className="flex gap-3 mt-1">
          <button onClick={onClose} className="flex-1 py-2.5 rounded-xl bg-gray-800 hover:bg-gray-700 text-gray-300 text-sm font-semibold transition-all active:scale-95">
            Odustani
          </button>
          <button
            onClick={() => {
              const selectedPlayer = activePlayers.find((p) => p.id === playerId);
              onSave({
                ...event,
                clockMin,
                playerId,
                snapNum: selectedPlayer?.num || event.snapNum,
                snapName: selectedPlayer?.name || event.snapName,
              });
            }}
            className="flex-1 py-2.5 rounded-xl bg-yellow-500 hover:bg-yellow-400 text-gray-900 text-sm font-black transition-all active:scale-95"
          >
            Sačuvaj
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Delete confirm modal ──────────────────────────────────────────────────────
function DeleteConfirmModal({ label, onConfirm, onClose }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm">
      <div className="bg-gray-900 border border-gray-700 rounded-2xl p-6 shadow-2xl flex flex-col items-center gap-4 w-72 mx-4">
        <div className="w-12 h-12 rounded-full bg-red-500/15 border border-red-500/40 flex items-center justify-center text-xl">🗑</div>
        <div className="text-center">
          <h2 className="text-white font-black text-sm uppercase tracking-[0.15em] mb-1">Obriši akciju?</h2>
          <p className="text-gray-400 text-xs leading-relaxed">{label}</p>
          <p className="text-red-400 text-xs font-semibold mt-1">Ova akcija se ne može poništiti.</p>
        </div>
        <div className="flex gap-3 w-full">
          <button onClick={onClose} className="flex-1 py-2.5 rounded-xl bg-gray-800 hover:bg-gray-700 text-gray-300 text-sm font-semibold transition-all active:scale-95">
            Odustani
          </button>
          <button onClick={onConfirm} className="flex-1 py-2.5 rounded-xl bg-red-600 hover:bg-red-500 text-white text-sm font-black transition-all active:scale-95">
            Obriši
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Goal log strip ────────────────────────────────────────────────────────────
function GoalStrip({ goals, players, allPlayers, align = "left", onEdit, onDelete }) {
  const [editingIdx, setEditingIdx] = useState(null);
  const [deletingIdx, setDeletingIdx] = useState(null);

  if (goals.length === 0) return null;
  return (
    <>
      <div className="space-y-1 mt-1">
        {goals.map((g, i) => {
          const player = players.find((p) => p.id === g.playerId);
          const num = player?.num || g.snapNum || "?";
          const name = player?.name || g.snapName || "Unknown";
          return (
            <div key={i} className={`flex items-center gap-1.5 text-xs group ${align === "right" ? "flex-row-reverse" : ""}`}>
              <span className="font-mono font-bold text-yellow-400 w-10 shrink-0 text-center bg-yellow-400/10 rounded px-1 py-0.5">
                {g.clockMin}&apos;
              </span>
              <span className="text-gray-300 flex-1">
                <span className="text-white font-semibold">#{num}</span> {name}
              </span>
              <div className={`flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity ${align === "right" ? "flex-row-reverse" : ""}`}>
                <button
                  onClick={() => setEditingIdx(i)}
                  className="w-5 h-5 rounded bg-gray-700 hover:bg-yellow-600 text-gray-300 hover:text-white flex items-center justify-center transition-all text-[10px]"
                  title="Uredi"
                >✏</button>
                <button
                  onClick={() => setDeletingIdx(i)}
                  className="w-5 h-5 rounded bg-gray-700 hover:bg-red-600 text-gray-300 hover:text-white flex items-center justify-center transition-all text-[10px]"
                  title="Obriši"
                >✕</button>
              </div>
            </div>
          );
        })}
      </div>

      {editingIdx !== null && (
        <EditEventModal
          event={goals[editingIdx]}
          players={players}
          allPlayers={allPlayers}
          type="goal"
          onSave={(updated) => { onEdit(editingIdx, updated); setEditingIdx(null); }}
          onClose={() => setEditingIdx(null)}
        />
      )}
      {deletingIdx !== null && (
        <DeleteConfirmModal
          label={`Gol u ${goals[deletingIdx].clockMin}' — #${goals[deletingIdx].snapNum || "?"} ${goals[deletingIdx].snapName || ""}`}
          onConfirm={() => { onDelete(deletingIdx); setDeletingIdx(null); }}
          onClose={() => setDeletingIdx(null)}
        />
      )}
    </>
  );
}

// ── Card log strip ────────────────────────────────────────────────────────────
function CardStrip({ cards, players, allPlayers, align = "left", onEdit, onDelete }) {
  const [editingIdx, setEditingIdx] = useState(null);
  const [deletingIdx, setDeletingIdx] = useState(null);

  if (cards.length === 0) return null;
  return (
    <>
      <div className="space-y-1 mt-1">
        {cards.map((c, i) => {
          const player = players.find((p) => p.id === c.playerId);
          const num = player?.num || c.snapNum || "?";
          const name = player?.name || c.snapName || "Unknown";
          const isRed = c.type === "red";
          return (
            <div key={i} className={`flex items-center gap-1.5 text-xs group ${align === "right" ? "flex-row-reverse" : ""}`}>
              <span className="font-mono font-bold text-yellow-400 w-10 shrink-0 text-center bg-yellow-400/10 rounded px-1 py-0.5">
                {c.clockMin}&apos;
              </span>
              <span className={`inline-block w-2.5 h-3.5 rounded-sm shrink-0 ${isRed ? "bg-red-500" : "bg-yellow-400"}`} />
              <span className="text-gray-300 flex-1">
                <span className="text-white font-semibold">#{num}</span> {name}
                {c.auto && <span className="text-red-400 text-[10px] ml-1">(2Y→R)</span>}
              </span>
              {!c.auto && (
                <div className={`flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity ${align === "right" ? "flex-row-reverse" : ""}`}>
                  <button
                    onClick={() => setEditingIdx(i)}
                    className="w-5 h-5 rounded bg-gray-700 hover:bg-yellow-600 text-gray-300 hover:text-white flex items-center justify-center transition-all text-[10px]"
                    title="Uredi"
                  >✏</button>
                  <button
                    onClick={() => setDeletingIdx(i)}
                    className="w-5 h-5 rounded bg-gray-700 hover:bg-red-600 text-gray-300 hover:text-white flex items-center justify-center transition-all text-[10px]"
                    title="Obriši"
                  >✕</button>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {editingIdx !== null && (
        <EditEventModal
          event={cards[editingIdx]}
          players={players}
          allPlayers={allPlayers}
          type="card"
          onSave={(updated) => { onEdit(editingIdx, updated); setEditingIdx(null); }}
          onClose={() => setEditingIdx(null)}
        />
      )}
      {deletingIdx !== null && (
        <DeleteConfirmModal
          label={`${cards[deletingIdx].type === "red" ? "Crveni" : "Žuti"} karton u ${cards[deletingIdx].clockMin}' — #${cards[deletingIdx].snapNum || "?"} ${cards[deletingIdx].snapName || ""}`}
          onConfirm={() => { onDelete(deletingIdx); setDeletingIdx(null); }}
          onClose={() => setDeletingIdx(null)}
        />
      )}
    </>
  );
}

// ── Foul controls row ─────────────────────────────────────────────────────────
function FoulControls({ fouls, onChange, align = "left" }) {
  return (
    <div className={`flex items-center gap-2 ${align === "right" ? "flex-row-reverse" : ""}`}>
      <FoulIndicator count={fouls} />
      <div className="flex items-center gap-1 ml-1">
        <button
          onClick={() => onChange(-1)}
          disabled={fouls === 0}
          className="w-6 h-6 rounded bg-gray-700 hover:bg-gray-600 disabled:opacity-30 disabled:cursor-not-allowed text-white text-sm font-bold transition-all active:scale-90 flex items-center justify-center"
        >
          −
        </button>
        <span className="font-mono text-sm text-white w-4 text-center">{fouls}</span>
        <button
          onClick={() => onChange(1)}
          disabled={fouls >= 8}
          className="w-6 h-6 rounded bg-red-800 hover:bg-red-600 disabled:opacity-30 disabled:cursor-not-allowed text-white text-sm font-bold transition-all active:scale-90 flex items-center justify-center"
        >
          +
        </button>
      </div>
    </div>
  );
}

// ── Excel roster import handler ───────────────────────────────────────────────
// Reads first column = player name only. No header row.
async function importRosterFromExcel(file) {
  const XLSX = await getXLSX();
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const players = rows
    .map((row) => ({
      num: "",
      name: row[0] !== undefined && row[0] !== null ? String(row[0]).trim() : "",
    }))
    .filter((p) => p.name !== "")
    .map((p) => ({ id: crypto.randomUUID(), num: p.num, name: p.name }));

  if (players.length === 0) {
    throw new Error("Excel fajl je prazan ili nije u očekivanom formatu.");
  }

  return players;
}

// ── Excel Import Button ───────────────────────────────────────────────────────
function ExcelImportButton({ onImport, align = "left" }) {
  const inputRef = useRef(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    setError("");
    try {
      const players = await importRosterFromExcel(file);
      onImport(players);
    } catch (err) {
      console.error(err);
      setError(err.message || "Greška pri učitavanju fajla.");
      setTimeout(() => setError(""), 5000);
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };

  return (
    <div className={`flex flex-col ${align === "right" ? "items-end" : "items-start"}`}>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleFile}
        className="hidden"
      />
      <button
        onClick={() => inputRef.current?.click()}
        disabled={busy}
        className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[10px] font-bold uppercase tracking-widest transition-all active:scale-95 ${
          busy
            ? "bg-gray-800 text-gray-500 cursor-not-allowed"
            : "bg-emerald-700/30 hover:bg-emerald-600/40 border border-emerald-600/50 text-emerald-300"
        }`}
        title="Učitaj rosters iz Excel fajla (1. kolona = broj, 2. kolona = ime i prezime)"
      >
        <span>📊</span>
        <span>{busy ? "Učitavanje…" : "Excel import"}</span>
      </button>
      {error && (
        <span className="text-[10px] text-red-400 mt-1 max-w-[180px] text-right">{error}</span>
      )}
    </div>
  );
}


// ── Team Selector (Google Drive link + dropdown) ──────────────────────────────
// Konvertuje bilo koji Google Drive/Sheets link u direktan download link
function driveShareToDirect(url) {
  // Google Sheets link — export kao xlsx
  const sheetsMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (sheetsMatch) return `https://docs.google.com/spreadsheets/d/${sheetsMatch[1]}/export?format=xlsx`;

  // Google Drive file link
  const driveMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (driveMatch) return `https://drive.google.com/uc?export=download&id=${driveMatch[1]}`;

  // Već direktan link
  if (url.includes("uc?export=download") || url.includes("export?format=xlsx")) return url;

  return null;
}

function TeamSelector({ side, teams, selectedTeam, onSelectTeam, onLigaUpload, ligaLoaded, align = "left" }) {
  const [driveLink, setDriveLink] = useState(() => {
    try { return localStorage.getItem("futsal-drive-link") || ""; } catch { return ""; }
  });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [showInput, setShowInput] = useState(false);
  const teamNames = Object.keys(teams || {});
  const isLeft = align === "left";

  const handleLoad = async () => {
    if (!driveLink.trim()) { setError("Unesi Google Drive link."); return; }
    const directUrl = driveShareToDirect(driveLink.trim());
    if (!directUrl) { setError("Neispravan Google Drive link."); return; }

    setBusy(true);
    setError("");
    try {
      const res = await fetch(directUrl);
      if (!res.ok) throw new Error("Nije moguće učitati fajl.");
      const arrayBuffer = await res.arrayBuffer();
      const XLSX = await getXLSX();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const parsedTeams = {};
      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const players = rows
          .map((row) => ({
            num: "",
            name: row[0] !== undefined ? String(row[0]).trim() : "",
          }))
          .filter((p) => p.name !== "")
          .map((p) => ({ id: crypto.randomUUID(), num: p.num, name: p.name }));
        if (players.length > 0) parsedTeams[sheetName] = players;
      }
      if (Object.keys(parsedTeams).length === 0) throw new Error("Nisu pronađeni timovi u fajlu.");
      try { localStorage.setItem("futsal-drive-link", driveLink.trim()); } catch {}
      saveLigaToStorage(null, parsedTeams);
      onLigaUpload(parsedTeams);
      setShowInput(false);
    } catch (err) {
      setError(err.message || "Greška pri učitavanju.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={`flex flex-col gap-1.5 ${isLeft ? "items-start" : "items-end"}`}>

      {/* Dugme za prikaz input polja */}
      <button
        onClick={() => setShowInput((v) => !v)}
        className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[10px] font-bold uppercase tracking-widest transition-all active:scale-95 ${
          ligaLoaded
            ? "bg-gray-700/50 hover:bg-gray-600/50 border border-gray-600 text-gray-400"
            : "bg-yellow-500/20 hover:bg-yellow-400/30 border border-yellow-500/50 text-yellow-300 animate-pulse"
        }`}
      >
        <span>📋</span>
        <span>{ligaLoaded ? "Zameni ekipe (Drive)" : "Učitaj ekipe (Excel)"}</span>
      </button>

      {/* Google Drive link input */}
      {showInput && (
        <div className="w-full flex flex-col gap-1.5">
          <input
            type="text"
            value={driveLink}
            onChange={(e) => setDriveLink(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleLoad()}
            placeholder="Google Drive link (javni)..."
            className="w-full bg-gray-800 border border-gray-700 text-white text-xs rounded-lg px-2.5 py-2 focus:outline-none focus:border-yellow-500 placeholder-gray-600"
          />
          <div className={`flex gap-1.5 ${isLeft ? "" : "flex-row-reverse"}`}>
            <button
              onClick={handleLoad}
              disabled={busy}
              className={`flex-1 py-1.5 rounded-lg text-[10px] font-bold uppercase tracking-widest transition-all active:scale-95 ${
                busy ? "bg-gray-700 text-gray-500 cursor-not-allowed" : "bg-yellow-500 hover:bg-yellow-400 text-gray-900"
              }`}
            >
              {busy ? "Učitavanje…" : "Učitaj"}
            </button>
            <button
              onClick={() => setShowInput(false)}
              className="px-3 py-1.5 rounded-lg text-[10px] bg-gray-700 hover:bg-gray-600 text-gray-300 transition-all"
            >
              Otkaži
            </button>
          </div>
          {error && <span className="text-[10px] text-red-400">{error}</span>}
          <p className="text-[9px] text-gray-600 leading-relaxed">
            Fajl mora biti javno dostupan na Drive-u.<br />
            Drive → Podeli → "Svi koji imaju link"
          </p>
        </div>
      )}

      {/* Dropdown za izbor tima */}
      {teamNames.length > 0 && (
        <select
          value={selectedTeam}
          onChange={(e) => onSelectTeam(e.target.value, teams[e.target.value] || [])}
          className={`w-full bg-gray-800 border-2 ${
            selectedTeam ? "border-yellow-500/60" : "border-yellow-500 animate-pulse"
          } text-white font-black text-sm uppercase tracking-wider rounded-xl px-3 py-2 focus:outline-none focus:border-yellow-400 transition-all cursor-pointer`}
          style={{ textAlignLast: isLeft ? "left" : "right" }}
        >
          <option value="">— Izaberi tim —</option>
          {teamNames.map((name) => (
            <option key={name} value={name}>{name}</option>
          ))}
        </select>
      )}

      {!ligaLoaded && !showInput && (
        <p className="text-[10px] text-gray-600 italic">
          Postavi Google Drive link da biraš timove
        </p>
      )}
    </div>
  );
}


function TeamPanel({
  side, teamName, onTeamNameChange, players, onPlayerUpdate,
  onGoal, onCard, fouls, onFoulChange, goals, cards,
  onEditGoal, onDeleteGoal, onEditCard, onDeleteCard, onImportRoster,
  ligaTeams, onLigaUpload, ligaLoaded, selectedTeam, onSelectTeam,
}) {
  const isLeft = side === "home";
  const playedPlayers = players.filter((p) => p.played && (p.num || p.name));
  const activePlayers = players.filter((p) => p.num || p.name);

  return (
    <div className="flex flex-col gap-3 h-full">
      {/* Team Selector from Liga Excel */}
      <TeamSelector
        side={side}
        teams={ligaTeams}
        selectedTeam={selectedTeam}
        onSelectTeam={onSelectTeam}
        onLigaUpload={onLigaUpload}
        ligaLoaded={ligaLoaded}
        align={isLeft ? "left" : "right"}
      />

      {/* Team Name (editable, auto-filled from dropdown) */}
      <div>
        <label className="block text-[10px] uppercase tracking-[0.15em] text-gray-500 mb-1">
          Naziv tima
        </label>
        <input
          type="text"
          value={teamName}
          onChange={(e) => onTeamNameChange(e.target.value)}
          placeholder={isLeft ? "Home Team" : "Away Team"}
          className={`w-full bg-transparent border-b-2 border-yellow-500/60 focus:border-yellow-400 text-white font-black text-lg uppercase tracking-widest focus:outline-none pb-1 placeholder-gray-700 transition-colors ${
            isLeft ? "text-left" : "text-right"
          }`}
        />
      </div>

      {/* Fouls */}
      <div>
        <p className={`text-[10px] uppercase tracking-[0.15em] text-gray-500 mb-1.5 ${isLeft ? "" : "text-right"}`}>
          Team Fouls
        </p>
        <FoulControls fouls={fouls} onChange={onFoulChange} align={isLeft ? "left" : "right"} />
      </div>

      {/* Roster */}
      <div className="flex-1 overflow-y-auto">
        <div className={`flex items-center justify-between mb-1 ${isLeft ? "" : "flex-row-reverse"}`}>
          <p className="text-[10px] uppercase tracking-[0.15em] text-gray-500">
            Roster
            {playedPlayers.length > 0 && (
              <span className="ml-2 text-emerald-400 font-bold">
                {playedPlayers.length}/{activePlayers.length} igrali
              </span>
            )}
          </p>
          <div className={`flex items-center gap-2 text-[10px] text-gray-600 ${isLeft ? "" : "flex-row-reverse"}`}>
            <span className="flex items-center gap-1">
              <span className="inline-block w-2.5 h-3 rounded-sm bg-yellow-400" /> Yellow
            </span>
            <span className="flex items-center gap-1">
              <span className="inline-block w-2.5 h-3 rounded-sm bg-red-500" /> Red
            </span>
          </div>
        </div>
        <div>
          {players.map((p) => (
            <PlayerRow
              key={p.id}
              player={p}
              goals={goals}
              cards={cards}
              onUpdate={onPlayerUpdate}
              onGoal={() => onGoal(p)}
              onCard={(player, type) => onCard(player, type)}
            />
          ))}
        </div>
      </div>

      {/* Igrač utakmice — MVP se bira centralno, ovde samo prikazujemo broj */}

      {/* Goal log */}
      {goals.length > 0 && (
        <div>
          <p className={`text-[10px] uppercase tracking-[0.15em] text-gray-500 mb-0.5 ${isLeft ? "" : "text-right"}`}>
            Goals
          </p>
          <GoalStrip
            goals={goals}
            players={players}
            allPlayers={players}
            align={isLeft ? "left" : "right"}
            onEdit={onEditGoal}
            onDelete={onDeleteGoal}
          />
        </div>
      )}

      {/* Card log */}
      {cards.length > 0 && (
        <div>
          <p className={`text-[10px] uppercase tracking-[0.15em] text-gray-500 mb-0.5 ${isLeft ? "" : "text-right"}`}>
            Cards
          </p>
          <CardStrip
            cards={cards}
            players={players}
            allPlayers={players}
            align={isLeft ? "left" : "right"}
            onEdit={onEditCard}
            onDelete={onDeleteCard}
          />
        </div>
      )}
    </div>
  );
}

// ── Edit Clock Modal ──────────────────────────────────────────────────────────
function EditClockModal({ currentSecs, onSave, onClose }) {
  const [m, setM] = useState(Math.floor(currentSecs / 60));
  const [s, setS] = useState(currentSecs % 60);
  const clamp = (val, min, max) => Math.max(min, Math.min(max, val));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-sm">
      <div className="bg-gray-900 border border-gray-700 rounded-2xl p-8 shadow-2xl flex flex-col items-center gap-5 w-72">
        <h2 className="text-white font-black text-sm uppercase tracking-[0.2em]">Set Clock</h2>
        <div className="flex items-center gap-3">
          <div className="flex flex-col items-center gap-1">
            <span className="text-[10px] uppercase tracking-widest text-gray-500">Min</span>
            <input
              type="number" min={0} max={99} value={m}
              onChange={(e) => setM(clamp(parseInt(e.target.value) || 0, 0, 99))}
              className="w-20 text-center text-4xl font-mono font-black text-yellow-400 bg-gray-800 border border-gray-700 rounded-xl p-2 focus:outline-none focus:border-yellow-500"
            />
          </div>
          <span className="text-yellow-400 text-4xl font-mono font-black mt-5">:</span>
          <div className="flex flex-col items-center gap-1">
            <span className="text-[10px] uppercase tracking-widest text-gray-500">Sec</span>
            <input
              type="number" min={0} max={59} value={s}
              onChange={(e) => setS(clamp(parseInt(e.target.value) || 0, 0, 59))}
              className="w-20 text-center text-4xl font-mono font-black text-yellow-400 bg-gray-800 border border-gray-700 rounded-xl p-2 focus:outline-none focus:border-yellow-500"
            />
          </div>
        </div>
        <div className="flex gap-3 w-full mt-1">
          <button onClick={onClose} className="flex-1 py-2 rounded-xl bg-gray-800 hover:bg-gray-700 text-gray-300 text-sm font-semibold transition-all">
            Cancel
          </button>
          <button onClick={() => onSave(m * 60 + s)} className="flex-1 py-2 rounded-xl bg-yellow-500 hover:bg-yellow-400 text-gray-900 text-sm font-black transition-all">
            Save
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Central match timeline ────────────────────────────────────────────────────
function MatchTimeline({ homeGoals, awayGoals, homeCards, awayCards,
  homePlayers, awayPlayers, homeTeam, awayTeam }) {

  const allEvents = [
    ...homeGoals.map((g) => ({ ...g, side: "home", eventType: "goal" })),
    ...awayGoals.map((g) => ({ ...g, side: "away", eventType: "goal" })),
    ...homeCards.map((c) => ({ ...c, side: "home", eventType: "card" })),
    ...awayCards.map((c) => ({ ...c, side: "away", eventType: "card" })),
  ].sort((a, b) => a.elapsedSecs - b.elapsedSecs);

  if (allEvents.length === 0)
    return (
      <div className="flex flex-col items-center justify-center py-8 text-gray-700 text-xs uppercase tracking-widest select-none">
        <span className="text-3xl mb-2 opacity-30">⚽</span>
        No events yet
      </div>
    );

  return (
    <div className="space-y-1.5">
      {allEvents.map((ev, i) => {
        const isHome = ev.side === "home";
        const players = isHome ? homePlayers : awayPlayers;
        const player = players.find((p) => p.id === ev.playerId);
        const num = player?.num || ev.snapNum || "?";
        const name = player?.name || ev.snapName || "?";
        const teamName = isHome ? homeTeam : awayTeam;

        if (ev.eventType === "goal") {
          return (
            /* PROMENJENO: Ostaje isto, plavo za home, narandžasto za away */
            <div key={i} className={`flex items-center gap-2 px-3 py-2 rounded-lg ${
              isHome ? "bg-blue-950/40 border border-blue-900/40" : "bg-orange-950/40 border border-orange-900/30"
            }`}>
              <span className="text-yellow-400 font-mono font-bold text-xs w-9 shrink-0 text-right">{ev.clockMin}&apos;</span>
              <span className={`text-[9px] font-bold px-1 rounded ${ev.half === 2 ? "bg-orange-900/60 text-orange-400" : "bg-blue-900/60 text-blue-400"}`}>
                    {ev.half === 2 ? "P2" : "P1"}
                </span>
              <span className="text-lg shrink-0">⚽</span>
              <div className="flex-1 min-w-0">
                <div className={`text-[10px] font-bold uppercase tracking-widest ${isHome ? "text-blue-400" : "text-orange-400"}`}>{teamName.toUpperCase()}</div>
                <div className="text-white text-xs font-semibold truncate">#{num} {name}</div>
              </div>
            </div>
          );
        }

        // card event
        const isRed = ev.type === "red";
        return (
          /* KLJUČNA PROMENA: Izbačena provera za isRed iz pozadine div-a, sada gleda samo da li je isHome */
          <div key={i} className={`flex items-center gap-2 px-3 py-2 rounded-lg ${
            isHome ? "bg-blue-950/40 border border-blue-900/40" : "bg-orange-950/40 border border-orange-900/30"
          }`}>
            <span className="text-yellow-400 font-mono font-bold text-xs w-9 shrink-0 text-right">{ev.clockMin}&apos;</span>
            <span className={`text-[9px] font-bold px-1 rounded ${ev.half === 2 ? "bg-orange-900/60 text-orange-400" : "bg-blue-900/60 text-blue-400"}`}>
                    {ev.half === 2 ? "P2" : "P1"}
                </span>
            {/* Sam kvadratić kartona naravno ostaje crven ili žut u zavisnosti od isRed */}
            <span className={`inline-block w-3.5 h-5 rounded-sm shrink-0 ${isRed ? "bg-red-500 shadow-[0_0_6px_rgba(239,68,68,0.8)]" : "bg-yellow-400 shadow-[0_0_6px_rgba(250,204,21,0.6)]"}`} />
            <div className="flex-1 min-w-0">
              <div className={`text-[10px] font-bold uppercase tracking-widest ${isHome ? "text-blue-400" : "text-orange-400"}`}>{teamName.toUpperCase()}</div>
              <div className="text-white text-xs font-semibold truncate">
                #{num} {name}
                {ev.auto && <span className="text-red-400 text-[10px] ml-1">(2Y→R)</span>}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ── PDF Generator ─────────────────────────────────────────────────────────────
async function generatePDF({ homeTeam, awayTeam, homePlayers, awayPlayers,
  homeGoals, awayGoals, homeCards, awayCards, homeFouls, awayFouls, matchMvp }) {
  const html2pdf = await loadHtml2pdf();
  const now = new Date();
  const dateStr = now.toLocaleDateString("sr-Latn", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("sr-Latn", { hour: "2-digit", minute: "2-digit" });

  const formatEvents = (goals, cards, players) => {
    const allEvents = [
      ...goals.map((g) => ({ ...g, etype: "goal" })),
      ...cards.map((c) => ({ ...c, etype: "card" })),
    ].sort((a, b) => (a.half - b.half) || a.elapsedSecs - b.elapsedSecs);
    if (allEvents.length === 0) return "<p style='color:#999;font-size:11px;margin:4px 0'>Nema događaja</p>";
    let html = "";
    let lastHalf = null;
    allEvents.forEach((ev) => {
      if (ev.half !== lastHalf) {
        html += `<div style="font-size:10px;font-weight:bold;color:#666;margin:6px 0 2px;border-top:1px solid #eee;padding-top:4px">${ev.half === 1 ? "1. Poluvreme" : "2. Poluvreme"}</div>`;
        lastHalf = ev.half;
      }
      const player = players.find((p) => p.id === ev.playerId);
      const name = player?.name || ev.snapName || "Nepoznat";
      const num = player?.num ? `#${player.num} ` : "";
      if (ev.etype === "goal") {
        html += `<div style="display:flex;align-items:center;gap:6px;padding:2px 0;font-size:12px">
          <span style="font-weight:bold;color:#333;min-width:30px">${ev.clockMin}'</span>
          <span>⚽</span>
          <span>${num}${name}</span>
        </div>`;
      } else {
        const isRed = ev.type === "red";
        const label = ev.auto ? " (2Ž→C)" : "";
        html += `<div style="display:flex;align-items:center;gap:6px;padding:2px 0;font-size:12px">
          <span style="font-weight:bold;color:#333;min-width:30px">${ev.clockMin}'</span>
          <span style="display:inline-block;width:10px;height:14px;background:${isRed ? "#ef4444" : "#eab308"};border-radius:2px;flex-shrink:0"></span>
          <span>${num}${name}${label}</span>
        </div>`;
      }
    });
    return html;
  };

  const formatRoster = (players, goals, cards) => {
    const active = players.filter((p) => p.num || p.name);
    if (active.length === 0) return `<tr><td colspan="4" style="padding:6px;color:#999;font-size:11px;text-align:center">Nema unetih igrača</td></tr>`;
    return active.map((p, idx) => {
      const pg = goals.filter((g) => g.playerId === p.id);
      const pc = cards.filter((c) => c.playerId === p.id);
      const goalStr = pg.map((g) => `⚽ ${g.clockMin}'`).join(" ");
      const cardStr = pc.map((c) =>
        `<span style="display:inline-block;width:8px;height:11px;background:${c.type === "red" ? "#ef4444" : "#eab308"};border-radius:1px;vertical-align:middle;margin-right:1px"></span>${c.clockMin}'`
      ).join(" ");
      const bg = idx % 2 === 0 ? "#f9f9f9" : "#fff";
      const bold = pg.length > 0 ? "font-weight:bold" : "";
      return `<tr style="background:${bg}">
        <td style="padding:4px 6px;font-size:11px;color:#555;white-space:nowrap">${p.num || "–"}</td>
        <td style="padding:4px 6px;font-size:12px;${bold}">${p.name || "–"}</td>
        <td style="padding:4px 6px;font-size:11px;white-space:nowrap">${goalStr}</td>
        <td style="padding:4px 6px;font-size:11px;white-space:nowrap">${cardStr}</td>
      </tr>`;
    }).join("");
  };

  const teamBlock = (teamName, players, goals, cards, fouls) => `
    <div style="flex:1;min-width:0">
      <div style="background:#111;color:#fff;padding:6px 10px;font-size:12px;font-weight:bold;text-transform:uppercase;letter-spacing:1px;border-radius:4px 4px 0 0">${teamName} <span style="font-weight:normal;font-size:10px;opacity:0.6">Fauli: ${fouls}/8</span></div>
      <table style="width:100%;border-collapse:collapse;border:1px solid #ddd;border-top:none">
        <thead>
          <tr style="background:#eee">
            <th style="padding:3px 6px;font-size:10px;text-align:left;color:#666;font-weight:bold">#</th>
            <th style="padding:3px 6px;font-size:10px;text-align:left;color:#666;font-weight:bold">Igrač</th>
            <th style="padding:3px 6px;font-size:10px;text-align:left;color:#666;font-weight:bold">Golovi</th>
            <th style="padding:3px 6px;font-size:10px;text-align:left;color:#666;font-weight:bold">Kartoni</th>
          </tr>
        </thead>
        <tbody>${formatRoster(players, goals, cards)}</tbody>
      </table>
      <div style="margin-top:8px">
        <div style="font-size:10px;font-weight:bold;color:#333;margin-bottom:2px;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid #eee;padding-bottom:2px">Događaji</div>
        ${formatEvents(goals, cards, players)}
      </div>
    </div>`;

  const html = `
    <div style="font-family:Arial,Helvetica,sans-serif;background:#fff;color:#111;padding:16px;max-width:794px;box-sizing:border-box">
      <div style="text-align:center;border-bottom:2px solid #111;padding-bottom:8px;margin-bottom:12px">
        <div style="font-size:18px;font-weight:900;letter-spacing:3px;text-transform:uppercase">IT LIGA++ NIŠ</div>
        <div style="font-size:12px;font-weight:bold;margin-top:2px">Zapisnik meča</div>
        <div style="font-size:10px;color:#666;margin-top:2px">${dateStr} · ${timeStr}</div>
      </div>
      <div style="text-align:center;margin-bottom:14px;padding:10px;background:#f5f5f5;border-radius:6px;border:1px solid #ddd">
        <div style="display:flex;align-items:center;justify-content:center;gap:20px">
          <div style="text-align:right">
            <div style="font-size:11px;color:#666;text-transform:uppercase;letter-spacing:1px">${homeTeam}</div>
            <div style="font-size:44px;font-weight:900;line-height:1.1">${homeGoals.length}</div>
          </div>
          <div style="font-size:28px;color:#999;font-weight:bold">–</div>
          <div style="text-align:left">
            <div style="font-size:11px;color:#666;text-transform:uppercase;letter-spacing:1px">${awayTeam}</div>
            <div style="font-size:44px;font-weight:900;line-height:1.1">${awayGoals.length}</div>
          </div>
        </div>
      </div>
      <div style="display:flex;gap:14px">
        ${teamBlock(homeTeam, homePlayers, homeGoals, homeCards, homeFouls)}
        ${teamBlock(awayTeam, awayPlayers, awayGoals, awayCards, awayFouls)}
      </div>
      ${(() => {
        const allPlayers = [
          ...homePlayers.map(p => ({ ...p, team: homeTeam })),
          ...awayPlayers.map(p => ({ ...p, team: awayTeam })),
        ];
        const mvpPlayer = allPlayers.find(p => p.id === matchMvp);
        if (!mvpPlayer) return "";
        return `<div style="margin-top:14px;padding:10px;background:#fffbeb;border:1px solid #fde68a;border-radius:6px;text-align:center">
          <div style="font-size:11px;font-weight:bold;color:#92400e;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px">⭐ Igrač utakmice</div>
          <div style="font-size:18px;font-weight:900;color:#111">${mvpPlayer.num ? "#"+mvpPlayer.num+" " : ""}${mvpPlayer.name}</div>
          <div style="font-size:11px;color:#92400e;margin-top:2px">${mvpPlayer.team}</div>
        </div>`;
      })()}
      <div style="margin-top:14px;border-top:1px solid #ddd;padding-top:6px;text-align:center;font-size:10px;color:#999">
        IT LIGA++ NIŠ · Automatski generisano
      </div>
    </div>`;

  const el = document.createElement("div");
  el.innerHTML = html;
  document.body.appendChild(el);

  const safeHome = homeTeam.replace(/\s+/g, "-").toLowerCase();
  const safeAway = awayTeam.replace(/\s+/g, "-").toLowerCase();

  await html2pdf().set({
    margin: 6,
    filename: `itliga-${safeHome}-vs-${safeAway}.pdf`,
    html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff" },
    jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
  }).from(el).save();

  document.body.removeChild(el);
}

// ── Match Close Modal (zaključi meč + QR za potvrdu) ─────────────────────────
function MatchCloseModal({ homeTeam, awayTeam, homeGoals, awayGoals,
  homeCards, awayCards, homePlayers, awayPlayers, homeFouls, awayFouls,
  onClose, onConfirm }) {
  const [sending, setSending] = useState(false);
  const [sent, setSent] = useState(false);
  const [error, setError] = useState("");

  const matchLabel = `${homeTeam} vs ${awayTeam} — ${new Date().toLocaleDateString("sr")}`;

  // Format goals for sheet
  const formatGoals = (goals, players) =>
    goals.map((g) => {
      const p = players.find((pl) => pl.id === g.playerId);
      return `${g.clockMin}' #${p?.num || g.snapNum || "?"} ${p?.name || g.snapName || "?"}`;
    }).join(", ") || "—";

  const formatCards = (cards, players) =>
    cards.map((c) => {
      const p = players.find((pl) => pl.id === c.playerId);
      const type = c.type === "red" ? (c.auto ? "RC(2Y)" : "RC") : "YC";
      return `${c.clockMin}' ${type} #${p?.num || c.snapNum || "?"} ${p?.name || c.snapName || "?"}`;
    }).join(", ") || "—";

  const sendToSheets = async () => {
    setSending(true);
    setError("");
    try {
      const payload = {
        homeTeam,
        awayTeam,
        homeScore: homeGoals.length,
        awayScore: awayGoals.length,
        goalsDetail: `${homeTeam}: ${formatGoals(homeGoals, homePlayers)} | ${awayTeam}: ${formatGoals(awayGoals, awayPlayers)}`,
        cardsDetail: `${homeTeam}: ${formatCards(homeCards, homePlayers)} | ${awayTeam}: ${formatCards(awayCards, awayPlayers)}`,
        homeFouls,
        awayFouls,
        status: "Zaključan — čeka potvrdu",
        matchLabel,
      };
      await fetch(SHEETS_WEBHOOK, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      setSent(true);
      onConfirm();
    } catch (e) {
      setError("Greška pri slanju. Provjeri internet konekciju.");
    } finally {
      setSending(false);
    }
  };

  // QR code via Google Charts API
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(CONFIRM_FORM_URL)}`;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 backdrop-blur-sm p-4">
      <div className="bg-gray-900 border border-gray-700 rounded-2xl shadow-2xl flex flex-col gap-5 w-full max-w-md">
        {/* Header */}
        <div className="bg-gray-800 rounded-t-2xl px-6 py-4 border-b border-gray-700">
          <h2 className="text-white font-black text-base uppercase tracking-[0.15em] text-center">
            Zaključi meč
          </h2>
          <p className="text-yellow-400 text-xs text-center font-semibold mt-1">{matchLabel}</p>
        </div>

        <div className="px-6 flex flex-col gap-4">
          {/* Rezultat summary */}
          <div className="bg-gray-800 rounded-xl p-4 flex items-center justify-center gap-6">
            <div className="text-center">
              <div className="text-xs text-gray-400 uppercase tracking-wider mb-1 truncate max-w-24">{homeTeam}</div>
              <div className="text-5xl font-black text-white">{homeGoals.length}</div>
            </div>
            <div className="text-gray-600 text-2xl font-bold">–</div>
            <div className="text-center">
              <div className="text-xs text-gray-400 uppercase tracking-wider mb-1 truncate max-w-24">{awayTeam}</div>
              <div className="text-5xl font-black text-white">{awayGoals.length}</div>
            </div>
          </div>

          {!sent ? (
            <>
              <p className="text-gray-400 text-xs text-center leading-relaxed">
                Klikom na <span className="text-yellow-400 font-bold">"Zaključi i pošalji"</span> meč se šalje u Google Sheets bazu. Kapiten protivničke ekipe potom skenira QR kod i potvrđuje rezultat.
              </p>
              {error && <p className="text-red-400 text-xs text-center">{error}</p>}
              <div className="flex gap-3">
                <button
                  onClick={onClose}
                  className="flex-1 py-2.5 rounded-xl bg-gray-800 hover:bg-gray-700 text-gray-300 text-sm font-semibold transition-all active:scale-95"
                >
                  Odustani
                </button>
                <button
                  onClick={sendToSheets}
                  disabled={sending}
                  className={`flex-1 py-2.5 rounded-xl text-sm font-black transition-all active:scale-95 ${
                    sending ? "bg-gray-700 text-gray-500 cursor-not-allowed"
                    : "bg-yellow-500 hover:bg-yellow-400 text-gray-900"
                  }`}
                >
                  {sending ? "Slanje…" : "Zaključi i pošalji ✓"}
                </button>
              </div>
            </>
          ) : (
            <>
              {/* QR kod za potvrdu */}
              <div className="flex flex-col items-center gap-3">
                <div className="w-6 h-6 rounded-full bg-emerald-500 flex items-center justify-center text-white font-black text-sm">✓</div>
                <p className="text-emerald-400 text-sm font-bold text-center">Meč uspešno poslat u bazu!</p>
                <p className="text-gray-400 text-xs text-center leading-relaxed">
                  Kapiten protivničke ekipe treba da skenira QR kod i potvrdi rezultat:
                </p>
                <div className="bg-white p-3 rounded-xl">
                  <img src={qrUrl} alt="QR kod za potvrdu" className="w-44 h-44" />
                </div>
                <p className="text-gray-500 text-[10px] text-center">ili otvori direktno:<br/>
                  <a href={CONFIRM_FORM_URL} target="_blank" rel="noreferrer"
                    className="text-yellow-400 underline break-all">{CONFIRM_FORM_URL}</a>
                </p>
              </div>
              <button
                onClick={onClose}
                className="w-full py-2.5 rounded-xl bg-emerald-700 hover:bg-emerald-600 text-white text-sm font-black transition-all active:scale-95"
              >
                Zatvori
              </button>
            </>
          )}
        </div>
        <div className="h-2"></div>
      </div>
    </div>
  );
}

// ── App Root ──────────────────────────────────────────────────────────────────
export default function FutsalTracker() {
  const saved = useRef(loadMatchFromStorage());
  const s = saved.current;

  const [showMatchClose, setShowMatchClose] = useState(false);
  const [matchLocked, setMatchLocked] = useState(s?.matchLocked ?? false);

  const [running, setRunning] = useState(false);
  const [secs, setSecs] = useState(s?.secs ?? DEFAULT_HALF_SECS);
  const [halfDuration, setHalfDuration] = useState(s?.halfDuration ?? DEFAULT_HALF_SECS);
  const [showEditClock, setShowEditClock] = useState(false);
  const [pdfBusy, setPdfBusy] = useState(false);
  const [pdfMsg, setPdfMsg] = useState("");
  const [showNewMatchConfirm, setShowNewMatchConfirm] = useState(false);
  const [half, setHalf] = useState(s?.half ?? 1);

  const [homeTeam, setHomeTeam] = useState(s?.homeTeam ?? "");
  const [awayTeam, setAwayTeam] = useState(s?.awayTeam ?? "");
  const [homePlayers, setHomePlayers] = useState(s?.homePlayers ?? initRoster(ROSTER_SIZE));
  const [awayPlayers, setAwayPlayers] = useState(s?.awayPlayers ?? initRoster(ROSTER_SIZE));
  const [homeGoals, setHomeGoals] = useState(s?.homeGoals ?? []);
  const [awayGoals, setAwayGoals] = useState(s?.awayGoals ?? []);
  const [homeCards, setHomeCards] = useState(s?.homeCards ?? []);
  const [awayCards, setAwayCards] = useState(s?.awayCards ?? []);
  const [homeFouls, setHomeFouls] = useState(s?.homeFouls ?? 0);
  const [awayFouls, setAwayFouls] = useState(s?.awayFouls ?? 0);
  const [matchMvp, setMatchMvp] = useState(s?.matchMvp ?? "");
  const [restoredMsg, setRestoredMsg] = useState(s ? "Meč učitan iz memorije" : "");

  // Liga Excel state — shared between both team panels
  const [ligaTeams, setLigaTeams] = useState(() => {
    const liga = loadLigaFromStorage();
    return liga?.teams ?? {};
  });
  const [homeSelectedTeam, setHomeSelectedTeam] = useState(s?.homeTeam ?? "");
  const [awaySelectedTeam, setAwaySelectedTeam] = useState(s?.awayTeam ?? "");
  const ligaLoaded = Object.keys(ligaTeams).length > 0;

  const fileInputRef = useRef(null);

  const timerRef = useRef(null);
  const secsRef = useRef(secs);
  secsRef.current = secs;
  const halfRef = useRef(half);
  halfRef.current = half;
  const halfDurationRef = useRef(halfDuration);
  halfDurationRef.current = halfDuration;

  // Auto-save to localStorage on every relevant change
  useEffect(() => {
    saveMatchToStorage({
      secs, half, halfDuration, homeTeam, awayTeam, homePlayers, awayPlayers,
      homeGoals, awayGoals, homeCards, awayCards, homeFouls, awayFouls,
      matchMvp, matchLocked,
    });
  }, [secs, half, halfDuration, homeTeam, awayTeam, homePlayers, awayPlayers,
      homeGoals, awayGoals, homeCards, awayCards, homeFouls, awayFouls,
      matchMvp, matchLocked]);

  // Hide "restored" message after a few seconds
  useEffect(() => {
    if (restoredMsg) {
      const t = setTimeout(() => setRestoredMsg(""), 4000);
      return () => clearTimeout(t);
    }
  }, [restoredMsg]);

  const startTimeRef = useRef(null);
  const startSecsRef = useRef(null);

  // Timer — koristi Date.now() da radi i sa zaključanim ekranom
  useEffect(() => {
    if (running) {
      startTimeRef.current = Date.now();
      startSecsRef.current = secsRef.current;
      timerRef.current = setInterval(() => {
        const elapsed = Math.floor((Date.now() - startTimeRef.current) / 1000);
        const newSecs = startSecsRef.current - elapsed;
        if (newSecs <= 0) {
          clearInterval(timerRef.current);
          setSecs(0);
          setRunning(false);
        } else {
          setSecs(newSecs);
        }
      }, 500);
    } else {
      clearInterval(timerRef.current);
    }
    return () => clearInterval(timerRef.current);
  }, [running]);

  // Resync odmah kad se ekran otključa / tab postane aktivan
  useEffect(() => {
    const handleVisibility = () => {
      if (document.visibilityState === "visible" && running && startTimeRef.current) {
        const elapsed = Math.floor((Date.now() - startTimeRef.current) / 1000);
        const newSecs = startSecsRef.current - elapsed;
        if (newSecs <= 0) {
          clearInterval(timerRef.current);
          setSecs(0);
          setRunning(false);
        } else {
          setSecs(newSecs);
        }
      }
    };
    document.addEventListener("visibilitychange", handleVisibility);
    return () => document.removeEventListener("visibilitychange", handleVisibility);
  }, [running]);

  const toggleTimer = () => { if (secs > 0) setRunning((r) => !r); };
  const resetTimer = () => { setRunning(false); setSecs(halfDuration); };


  // Goal recording
  const recordGoal = useCallback((side, player) => {
    const elapsed = halfDurationRef.current - secsRef.current;
    const minuteInHalf = Math.ceil(elapsed / 60) || 1;
    const cumulativeMin = halfRef.current === 1
      ? minuteInHalf
      : Math.floor(halfDurationRef.current / 60) + minuteInHalf;
    const entry = {
      playerId: player.id,
      snapNum: player.num,
      snapName: player.name,
      clockMin: cumulativeMin,
      half: halfRef.current,
      elapsedSecs: elapsed,
      ts: Date.now(),
    };
    if (side === "home") setHomeGoals((g) => [...g, entry]);
    else setAwayGoals((g) => [...g, entry]);
  }, []);

  // Card recording — auto red if 2nd yellow
  const recordCard = useCallback((side, player, type) => {
    const elapsed = halfDurationRef.current - secsRef.current;
    const minuteInHalf = Math.ceil(elapsed / 60) || 1;
    const cumulativeMin = halfRef.current === 1
      ? minuteInHalf
      : Math.floor(halfDurationRef.current / 60) + minuteInHalf;

    const existingCards = side === "home" ? homeCards : awayCards;
    const playerYellows = existingCards.filter(
      (c) => c.playerId === player.id && c.type === "yellow"
    ).length;

    const setCards = side === "home" ? setHomeCards : setAwayCards;

    if (type === "yellow") {
      const newCard = {
        playerId: player.id, snapNum: player.num, snapName: player.name,
        type: "yellow", clockMin: cumulativeMin, half: halfRef.current,
        elapsedSecs: elapsed, ts: Date.now(), auto: false,
      };
      if (playerYellows >= 1) {
        const autoRed = {
          playerId: player.id, snapNum: player.num, snapName: player.name,
          type: "red", clockMin: cumulativeMin, half: halfRef.current,
          elapsedSecs: elapsed, ts: Date.now() + 1, auto: true,
        };
        setCards((c) => [...c, newCard, autoRed]);
      } else {
        setCards((c) => [...c, newCard]);
      }
    } else {
      const newCard = {
        playerId: player.id, snapNum: player.num, snapName: player.name,
        type: "red", clockMin: cumulativeMin, half: halfRef.current,
        elapsedSecs: elapsed, ts: Date.now(), auto: false,
      };
      setCards((c) => [...c, newCard]);
    }
  }, [homeCards, awayCards]);

  const updatePlayer = (side, updated) => {
    if (side === "home") setHomePlayers((ps) => ps.map((p) => p.id === updated.id ? updated : p));
    else setAwayPlayers((ps) => ps.map((p) => p.id === updated.id ? updated : p));
  };

  const changeFoul = (side, delta) => {
    const clamp = (v) => Math.max(0, Math.min(8, v));
    if (side === "home") setHomeFouls((f) => clamp(f + delta));
    else setAwayFouls((f) => clamp(f + delta));
  };

  // Goal edit / delete
  const editGoal = (side, idx, updated) => {
    if (side === "home") setHomeGoals((g) => g.map((item, i) => i === idx ? updated : item));
    else setAwayGoals((g) => g.map((item, i) => i === idx ? updated : item));
  };
  const deleteGoal = (side, idx) => {
    if (side === "home") setHomeGoals((g) => g.filter((_, i) => i !== idx));
    else setAwayGoals((g) => g.filter((_, i) => i !== idx));
  };

  // Card edit / delete
  const editCard = (side, idx, updated) => {
    if (side === "home") setHomeCards((c) => c.map((item, i) => i === idx ? updated : item));
    else setAwayCards((c) => c.map((item, i) => i === idx ? updated : item));
  };
  const deleteCard = (side, idx) => {
    // Ako brišemo žuti karton, brišemo i eventualni auto-crveni koji je nastao od njega
    const cards = side === "home" ? homeCards : awayCards;
    const deletedCard = cards[idx];
    const filtered = cards.filter((_, i) => i !== idx);
    // Ako je obrisan žuti, proveravamo da li postoji auto-crveni za istog igrača i brišemo ga
    const cleaned = deletedCard.type === "yellow"
      ? filtered.filter((c, i) => !(c.playerId === deletedCard.playerId && c.type === "red" && c.auto && c.ts === deletedCard.ts + 1))
      : filtered;
    if (side === "home") setHomeCards(cleaned);
    else setAwayCards(cleaned);
  };

  const handlePDF = async () => {
    setPdfBusy(true);
    setPdfMsg("Building report…");
    try {
      await generatePDF({
        homeTeam, awayTeam, homePlayers, awayPlayers,
        homeGoals, awayGoals, homeCards, awayCards,
        homeFouls, awayFouls, matchMvp,
      });
      setPdfMsg("Downloaded ✓");
      setTimeout(() => setPdfMsg(""), 3000);
    } catch (e) {
      console.error(e);
      setPdfMsg("Error — try again");
      setTimeout(() => setPdfMsg(""), 4000);
    } finally {
      setPdfBusy(false);
    }
  };

  const handleNewMatch = () => {
    setRunning(false);
    setSecs(DEFAULT_HALF_SECS);
    setHalfDuration(DEFAULT_HALF_SECS);
    setHomeTeam("");
    setAwayTeam("");
    setHomePlayers(initRoster(ROSTER_SIZE));
    setAwayPlayers(initRoster(ROSTER_SIZE));
    setHomeGoals([]);
    setAwayGoals([]);
    setHomeCards([]);
    setAwayCards([]);
    setHomeFouls(0);
    setAwayFouls(0);
    setMatchMvp("");
    setPdfMsg("");
    setShowNewMatchConfirm(false);
    setHalf(1);
    setMatchLocked(false);
    setHomeSelectedTeam("");
    setAwaySelectedTeam("");
    clearMatchFromStorage();
  };

  // Liga Excel upload — shared, saves to localStorage
  const handleLigaUpload = (parsedTeams) => {
    setLigaTeams(parsedTeams);
  };

  // When a team is selected from dropdown — set name + load players
  const handleSelectTeam = (side, teamName, players) => {
    if (!teamName) return;
    const padded = players.length >= ROSTER_SIZE
      ? players
      : [...players, ...initRoster(ROSTER_SIZE - players.length)];
    if (side === "home") {
      setHomeTeam(teamName);
      setHomePlayers(padded);
      setHomeSelectedTeam(teamName);
    } else {
      setAwayTeam(teamName);
      setAwayPlayers(padded);
      setAwaySelectedTeam(teamName);
    }
  };

  // Excel roster import — replaces entire roster for the given team
  const handleImportRoster = (side, importedPlayers) => {
    // Ako Excel ima manje igrača od ROSTER_SIZE, dopuni prazninama; ako ima vise, proširi listu
    const finalRoster = importedPlayers.length >= ROSTER_SIZE
      ? importedPlayers
      : [...importedPlayers, ...initRoster(ROSTER_SIZE - importedPlayers.length)];

    if (side === "home") setHomePlayers(finalRoster);
    else setAwayPlayers(finalRoster);
  };

  // Export current match state as a JSON file (backup)
  const handleExportJSON = () => {
    const state = {
      secs, half, homeTeam, awayTeam, homePlayers, awayPlayers,
      homeGoals, awayGoals, homeCards, awayCards, homeFouls, awayFouls,
      exportedAt: new Date().toISOString(),
    };
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const safeHome = (homeTeam || "home").replace(/[^a-z0-9]/gi, "-").toLowerCase();
    const safeAway = (awayTeam || "away").replace(/[^a-z0-9]/gi, "-").toLowerCase();
    a.href = url;
    a.download = `futsal-mec-${safeHome}-vs-${safeAway}-${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Import match state from a JSON file
  const handleImportJSON = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        setRunning(false);
        setSecs(data.secs ?? DEFAULT_HALF_SECS);
        setHalf(data.half ?? 1);
        setHomeTeam(data.homeTeam ?? "Home");
        setAwayTeam(data.awayTeam ?? "Away");
        setHomePlayers(data.homePlayers ?? initRoster(ROSTER_SIZE));
        setAwayPlayers(data.awayPlayers ?? initRoster(ROSTER_SIZE));
        setHomeGoals(data.homeGoals ?? []);
        setAwayGoals(data.awayGoals ?? []);
        setHomeCards(data.homeCards ?? []);
        setAwayCards(data.awayCards ?? []);
        setHomeFouls(data.homeFouls ?? 0);
        setAwayFouls(data.awayFouls ?? 0);
        setRestoredMsg("Meč učitan iz fajla");
      } catch (err) {
        console.error(err);
        alert("Greška: fajl nije validan JSON meč fajl.");
      } finally {
        e.target.value = "";
      }
    };
    reader.readAsText(file);
  };

  const pdfEnabled = !running && !pdfBusy;
  const isFullTime = secs === 0;

  return (
    <div className="min-h-screen bg-gray-950 text-white flex flex-col select-none"
      style={{ fontFamily: "'Inter', 'Segoe UI', system-ui, sans-serif" }}>

      {/* Restored from storage notice */}
      {restoredMsg && (
        <div className="bg-emerald-900/40 border-b border-emerald-700/40 text-emerald-300 text-[11px] font-semibold uppercase tracking-widest text-center py-1.5 shrink-0">
          ✓ {restoredMsg}
        </div>
      )}

      {/* Topbar */}
      <header className="flex items-center justify-between px-4 py-2.5 bg-gray-900 border-b border-gray-800 shrink-0">
        <div className="flex items-center gap-2">
          <img src="/src/assets/logo-liga.png" alt="" className="h-6 w-auto object-contain" />
          <span className="text-white font-black tracking-[0.2em] uppercase text-xs hidden sm:block">
            IT LIGA++ NIŠ
          </span>
        </div>
        {isFullTime && (
          <span className="px-3 py-0.5 bg-red-600 text-white text-[10px] font-black rounded-full uppercase tracking-widest animate-pulse">
            Full Time
          </span>
        )}
        <div className="hidden sm:flex items-center gap-2">
          <input
            ref={fileInputRef}
            type="file"
            accept=".json"
            onChange={handleImportJSON}
            className="hidden"
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 px-3 py-2 rounded-lg text-xs font-bold tracking-wide transition-all bg-gray-700 hover:bg-gray-600 active:scale-95 text-gray-200"
            title="Učitaj meč iz JSON fajla"
          >
            <span>📂</span>
            <span>Učitaj</span>
          </button>
          <button
            onClick={handleExportJSON}
            className="flex items-center gap-2 px-3 py-2 rounded-lg text-xs font-bold tracking-wide transition-all bg-gray-700 hover:bg-gray-600 active:scale-95 text-gray-200"
            title="Sačuvaj meč kao JSON fajl (bekap)"
          >
            <span>💾</span>
            <span>Sačuvaj</span>
          </button>
          <button
            onClick={() => setShowNewMatchConfirm(true)}
            className="flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold tracking-wide transition-all bg-gray-700 hover:bg-gray-600 active:scale-95 text-gray-200"
          >
            <span>🔄</span>
            <span>Novi meč</span>
          </button>
          {!matchLocked ? (
            <button
              onClick={() => setShowMatchClose(true)}
              disabled={!homeTeam || !awayTeam}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold tracking-wide transition-all active:scale-95 ${
                !homeTeam || !awayTeam
                  ? "bg-gray-800 text-gray-600 cursor-not-allowed"
                  : "bg-emerald-700 hover:bg-emerald-600 text-white"
              }`}
            >
              <span>✓</span>
              <span>Zaključi meč</span>
            </button>
          ) : (
            <span className="flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold bg-emerald-900/40 border border-emerald-700/50 text-emerald-400">
              <span>✓</span>
              <span>Meč zaključan</span>
            </span>
          )}
          <button
            onClick={handlePDF}
            disabled={!pdfEnabled}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold tracking-wide transition-all ${
              pdfEnabled
                ? "bg-yellow-500 hover:bg-yellow-400 active:scale-95 text-gray-900"
                : "bg-gray-800 text-gray-500 cursor-not-allowed"
            }`}
          >
            <span>📄</span>
            <span>{pdfBusy ? "Generating…" : pdfMsg || "PDF Report"}</span>
          </button>
        </div>
      </header>

      {/* SAT */}
      <div className="bg-gray-900 border-b border-gray-800 px-4 py-3 flex flex-col items-center">
        <div
          className={`font-mono font-black tabular-nums leading-none transition-all ${
            running
              ? "text-yellow-400 [text-shadow:0_0_40px_rgba(234,179,8,0.5)]"
              : isFullTime ? "text-red-400" : "text-yellow-300"
          }`}
          /* VELIČINA SATA */
          style={{ fontSize: "clamp(3.5rem, 20vw, 30rem)", letterSpacing: "0.05em" }}
        >
          {secsToDisplay(secs)}
        </div>
      </div>

      {/* 3-column layout */}
      <div className="flex-1 grid grid-cols-1 lg:grid-cols-[1fr_300px_1fr] xl:grid-cols-[1fr_340px_1fr] gap-0">

        {/* Kontrole sata, rezultat i oznaka poluvremena — skroluju zajedno sa sadržajem */}
        <div className="lg:col-span-3 shrink-0 bg-gray-900 border-b border-gray-800 px-4 py-4 flex flex-col items-center gap-1">
          <div className="flex items-center gap-2">
            <button
              onClick={() => { setRunning(false); setShowEditClock(true); }}
              className="px-3 py-1.5 bg-gray-800 hover:bg-gray-700 text-gray-300 text-[10px] font-semibold uppercase tracking-widest rounded-lg transition-all active:scale-95"
            >
              ✏ Edit
            </button>
            <button
              onClick={toggleTimer}
              disabled={secs === 0}
              className={`px-7 py-2 rounded-xl font-black text-sm tracking-widest transition-all active:scale-95 ${
                secs === 0 ? "bg-gray-800 text-gray-500 cursor-not-allowed"
                  : running ? "bg-amber-600 hover:bg-amber-500 text-white"
                  : "bg-emerald-600 hover:bg-emerald-500 text-white"
              }`}
            >
              {running ? "⏸ PAUSE" : "▶ START"}
            </button>
            <button
              onClick={resetTimer}
              className="px-3 py-1.5 bg-gray-800 hover:bg-gray-700 text-gray-300 text-[10px] font-semibold uppercase tracking-widest rounded-lg transition-all active:scale-95"
            >
              ↺ Reset
            </button>
          </div>

          {/* Poluvreme dugme */}
          <div className="flex items-center gap-3 mt-3">
            <button
              onClick={() => {
                if (half === 1) {
                  setRunning(false);
                  setSecs(halfDuration);
                  setHalf(2);
                }
              }}
              disabled={half === 2}
              className={`px-5 py-1.5 rounded-lg text-xs font-bold tracking-widest transition-all active:scale-95 ${
                half === 2
                  ? "bg-gray-800 text-gray-500 cursor-not-allowed"
                  : "bg-blue-700 hover:bg-blue-600 text-white"
              }`}
            >
              {half === 2 ? "2. Poluvreme" : "→ Počni 2. poluvreme"}
            </button>
          </div>

          {/* Live score */}
          <div className="flex items-center gap-6 mt-4">
            <div className="text-right">
              <div className="text-[10px] uppercase tracking-[0.2em] text-gray-400 mb-0.5 truncate max-w-28">{homeTeam}</div>
              <div className="font-black tabular-nums text-white" style={{ fontSize: "clamp(2rem, 6vw, 4rem)" }}>
                {homeGoals.length}
              </div>
            </div>
            <div className="text-gray-600 font-bold text-3xl pb-1">–</div>
            <div className="text-left">
              <div className="text-[10px] uppercase tracking-[0.2em] text-gray-400 mb-0.5 truncate max-w-28">{awayTeam}</div>
              <div className="font-black tabular-nums text-white" style={{ fontSize: "clamp(2rem, 6vw, 4rem)" }}>
                {awayGoals.length}
              </div>
            </div>
          </div>

          {/* Oznaka poluvremena */}
          <div className="mt-2 flex items-center gap-2">
            <span className={`px-3 py-1 rounded-full text-[11px] font-black uppercase tracking-widest ${
              half === 1
                ? "bg-blue-600/20 border border-blue-500/40 text-blue-300"
                : "bg-orange-600/20 border border-orange-500/40 text-orange-300"
            }`}>
              {HALF_LABELS[half - 1]}
            </span>
          </div>
        </div>

        {/* Home */}
        <div className="p-4 lg:p-5 border-b lg:border-b-0 lg:border-r border-gray-800">
          <TeamPanel
            side="home"
            teamName={homeTeam}
            onTeamNameChange={setHomeTeam}
            players={homePlayers}
            onPlayerUpdate={(p) => updatePlayer("home", p)}
            onGoal={(p) => recordGoal("home", p)}
            onCard={(p, t) => recordCard("home", p, t)}
            fouls={homeFouls}
            onFoulChange={(d) => changeFoul("home", d)}
            goals={homeGoals}
            cards={homeCards}
            onEditGoal={(idx, updated) => editGoal("home", idx, updated)}
            onDeleteGoal={(idx) => deleteGoal("home", idx)}
            onEditCard={(idx, updated) => editCard("home", idx, updated)}
            onDeleteCard={(idx) => deleteCard("home", idx)}
            onImportRoster={(players) => handleImportRoster("home", players)}
            ligaTeams={ligaTeams}
            onLigaUpload={handleLigaUpload}
            ligaLoaded={ligaLoaded}
            selectedTeam={homeSelectedTeam}
            onSelectTeam={(name, players) => handleSelectTeam("home", name, players)}
          />
        </div>

        {/* Centre timeline */}
        <div className="p-4 lg:p-5 border-b lg:border-b-0 lg:border-r border-gray-800 flex flex-col gap-4">
          <p className="text-[10px] uppercase tracking-[0.2em] text-gray-500 text-center">
            Match Timeline
          </p>
          <MatchTimeline
            homeGoals={homeGoals}
            awayGoals={awayGoals}
            homeCards={homeCards}
            awayCards={awayCards}
            homePlayers={homePlayers}
            awayPlayers={awayPlayers}
            homeTeam={homeTeam}
            awayTeam={awayTeam}
          />

          {/* Stats summary */}
          <div className="mt-auto pt-4 border-t border-gray-800 grid grid-cols-2 gap-3 text-center">
            <div>
              <div className="text-red-400 font-black text-xl">
                {homeFouls}<span className="text-gray-600 font-normal text-sm">/5</span>
              </div>
              <div className="text-[10px] uppercase tracking-widest text-gray-500 mt-0.5 truncate">{homeTeam} Fouls</div>
            </div>
            <div>
              <div className="text-red-400 font-black text-xl">
                {awayFouls}<span className="text-gray-600 font-normal text-sm">/5</span>
              </div>
              <div className="text-[10px] uppercase tracking-widest text-gray-500 mt-0.5 truncate">{awayTeam} Fouls</div>
            </div>
            {/* Card summary */}
            <div className="col-span-2 flex justify-center gap-6 pt-2 border-t border-gray-800">
              <div className="text-center">
                <div className="flex items-center justify-center gap-1">
                  <span className="inline-block w-3 h-4 rounded-sm bg-yellow-400" />
                  <span className="font-black text-yellow-400">{homeCards.filter(c => c.type === "yellow").length}</span>
                </div>
                <div className="text-[10px] text-gray-600 truncate">{homeTeam}</div>
              </div>
              <div className="text-center">
                <div className="flex items-center justify-center gap-1">
                  <span className="inline-block w-3 h-4 rounded-sm bg-red-500" />
                  <span className="font-black text-red-400">{homeCards.filter(c => c.type === "red").length}</span>
                </div>
                <div className="text-[10px] text-gray-600 truncate">{homeTeam}</div>
              </div>
              <div className="text-center">
                <div className="flex items-center justify-center gap-1">
                  <span className="inline-block w-3 h-4 rounded-sm bg-yellow-400" />
                  <span className="font-black text-yellow-400">{awayCards.filter(c => c.type === "yellow").length}</span>
                </div>
                <div className="text-[10px] text-gray-600 truncate">{awayTeam}</div>
              </div>
              <div className="text-center">
                <div className="flex items-center justify-center gap-1">
                  <span className="inline-block w-3 h-4 rounded-sm bg-red-500" />
                  <span className="font-black text-red-400">{awayCards.filter(c => c.type === "red").length}</span>
                </div>
                <div className="text-[10px] text-gray-600 truncate">{awayTeam}</div>
              </div>
            </div>

            {/* MVP — Igrač utakmice */}
            {(() => {
              const allPlayed = [
                ...homePlayers.filter(p => p.played && (p.num || p.name)).map(p => ({ ...p, team: homeTeam })),
                ...awayPlayers.filter(p => p.played && (p.num || p.name)).map(p => ({ ...p, team: awayTeam })),
              ];
              if (allPlayed.length === 0) return null;
              const mvpPlayer = allPlayed.find(p => p.id === matchMvp);
              return (
                <div className="col-span-2 pt-3 border-t border-gray-800">
                  <p className="text-[10px] uppercase tracking-widest text-yellow-500 text-center mb-2">⭐ Igrač utakmice</p>
                  <select
                    value={matchMvp}
                    onChange={(e) => setMatchMvp(e.target.value)}
                    className={`w-full bg-gray-800 border-2 ${
                      matchMvp ? "border-yellow-500/60" : "border-gray-700"
                    } text-white text-xs rounded-xl px-3 py-2 focus:outline-none focus:border-yellow-500 transition-all cursor-pointer`}
                  >
                    <option value="">— Izaberi igrača utakmice —</option>
                    <optgroup label={homeTeam}>
                      {homePlayers.filter(p => p.played && (p.num || p.name)).map(p => (
                        <option key={p.id} value={p.id}>{p.num ? `#${p.num} ` : ""}{p.name}</option>
                      ))}
                    </optgroup>
                    <optgroup label={awayTeam}>
                      {awayPlayers.filter(p => p.played && (p.num || p.name)).map(p => (
                        <option key={p.id} value={p.id}>{p.num ? `#${p.num} ` : ""}{p.name}</option>
                      ))}
                    </optgroup>
                  </select>
                  {mvpPlayer && (
                    <div className="mt-2 text-center">
                      <span className="text-yellow-400 font-black text-sm">
                        {mvpPlayer.num ? `#${mvpPlayer.num} ` : ""}{mvpPlayer.name}
                      </span>
                      <span className="text-gray-500 text-xs ml-1">({mvpPlayer.team})</span>
                    </div>
                  )}
                </div>
              );
            })()}
          </div>
        </div>
        <div className="p-4 lg:p-5">
          <TeamPanel
            side="away"
            teamName={awayTeam}
            onTeamNameChange={setAwayTeam}
            players={awayPlayers}
            onPlayerUpdate={(p) => updatePlayer("away", p)}
            onGoal={(p) => recordGoal("away", p)}
            onCard={(p, t) => recordCard("away", p, t)}
            fouls={awayFouls}
            onFoulChange={(d) => changeFoul("away", d)}
            goals={awayGoals}
            cards={awayCards}
            onEditGoal={(idx, updated) => editGoal("away", idx, updated)}
            onDeleteGoal={(idx) => deleteGoal("away", idx)}
            onEditCard={(idx, updated) => editCard("away", idx, updated)}
            onDeleteCard={(idx) => deleteCard("away", idx)}
            onImportRoster={(players) => handleImportRoster("away", players)}
            ligaTeams={ligaTeams}
            onLigaUpload={handleLigaUpload}
            ligaLoaded={ligaLoaded}
            selectedTeam={awaySelectedTeam}
            onSelectTeam={(name, players) => handleSelectTeam("away", name, players)}
          />
        </div>
      </div>

      {/* Mobile PDF + New Match bar */}
      <footer className="sm:hidden shrink-0 p-3 bg-gray-900 border-t border-gray-800 flex flex-col gap-2">
        <button
          onClick={handlePDF}
          disabled={!pdfEnabled}
          className={`w-full py-3 rounded-xl text-sm font-black tracking-widest transition-all ${
            pdfEnabled ? "bg-yellow-500 hover:bg-yellow-400 text-gray-900"
              : "bg-gray-800 text-gray-500 cursor-not-allowed"
          }`}
        >
          {pdfBusy ? "Generating…" : pdfMsg || "📄 Generate PDF Report"}
        </button>
        <div className="grid grid-cols-2 gap-2">
          <button
            onClick={() => fileInputRef.current?.click()}
            className="py-2.5 rounded-xl text-sm font-bold tracking-widest transition-all bg-gray-700 hover:bg-gray-600 text-gray-200 active:scale-95"
          >
            📂 Učitaj
          </button>
          <button
            onClick={handleExportJSON}
            className="py-2.5 rounded-xl text-sm font-bold tracking-widest transition-all bg-gray-700 hover:bg-gray-600 text-gray-200 active:scale-95"
          >
            💾 Sačuvaj
          </button>
        </div>
        <button
          onClick={() => setShowNewMatchConfirm(true)}
          className="w-full py-2.5 rounded-xl text-sm font-bold tracking-widest transition-all bg-gray-700 hover:bg-gray-600 text-gray-200 active:scale-95"
        >
          🔄 Novi meč
        </button>
        {running && (
          <p className="text-center text-[10px] text-gray-600">
            Pauziraj sat za PDF izveštaj
          </p>
        )}
      </footer>

      {/* Edit clock modal */}
      {showEditClock && (
        <EditClockModal
          currentSecs={secs}
          onSave={(newSecs) => { setSecs(newSecs); setHalfDuration(newSecs); setShowEditClock(false); }}
          onClose={() => setShowEditClock(false)}
        />
      )}

      {/* New Match confirmation modal */}
      {showNewMatchConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm">
          <div className="bg-gray-900 border border-gray-700 rounded-2xl p-8 shadow-2xl flex flex-col items-center gap-5 w-80 mx-4">
            <div className="w-14 h-14 rounded-full bg-orange-500/15 border border-orange-500/40 flex items-center justify-center">
              <span className="text-2xl">🔄</span>
            </div>
            <div className="text-center">
              <h2 className="text-white font-black text-base uppercase tracking-[0.15em] mb-2">
                Novi meč?
              </h2>
              <p className="text-gray-400 text-sm leading-relaxed">
                Svi podaci će biti obrisani — timovi, igrači, golovi, kartoni i sat.
                <span className="text-red-400 font-semibold"> Ova akcija se ne može poništiti.</span>
              </p>
            </div>
            <div className="flex gap-3 w-full">
              <button
                onClick={() => setShowNewMatchConfirm(false)}
                className="flex-1 py-2.5 rounded-xl bg-gray-800 hover:bg-gray-700 text-gray-300 text-sm font-semibold transition-all active:scale-95"
              >
                Odustani
              </button>
              <button
                onClick={handleNewMatch}
                className="flex-1 py-2.5 rounded-xl bg-orange-600 hover:bg-orange-500 text-white text-sm font-black transition-all active:scale-95"
              >
                Novi meč
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Match Close Modal */}
      {showMatchClose && (
        <MatchCloseModal
          homeTeam={homeTeam}
          awayTeam={awayTeam}
          homeGoals={homeGoals}
          awayGoals={awayGoals}
          homeCards={homeCards}
          awayCards={awayCards}
          homePlayers={homePlayers}
          awayPlayers={awayPlayers}
          homeFouls={homeFouls}
          awayFouls={awayFouls}
          onClose={() => setShowMatchClose(false)}
          onConfirm={() => { setMatchLocked(true); setRunning(false); }}
        />
      )}
    </div>
  );
}