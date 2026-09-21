// ─────────────────────────────────────────────────────────────────────────────
// FutsalTracker.jsx — Futsal Match Statistics App with Card System
// Dependencies: React (hooks), Tailwind CSS, jsPDF (loaded dynamically via CDN)
// ─────────────────────────────────────────────────────────────────────────────

import { useState, useEffect, useRef, useCallback } from "react";

// ── jsPDF dynamic loader ──────────────────────────────────────────────────────
async function getJsPDF() {
  if (window.jspdf?.jsPDF) return window.jspdf.jsPDF;
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src =
      "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    script.onload = () => resolve(window.jspdf.jsPDF);
    script.onerror = () => reject(new Error("Failed to load jsPDF"));
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
const makePlayer = () => ({ id: crypto.randomUUID(), num: "", name: "" });
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
  const isSuspended = hasRed; // suspended if red card

  return (
    <div
      className={`flex items-center gap-1.5 py-2 border-b border-gray-800/60 transition-colors ${
        hasRed
          ? "bg-red-950/30"
          : yellowCount >= 1
          ? "bg-yellow-950/20"
          : playerGoalCount > 0
          ? "bg-green-950/20"
          : ""
      }`}
    >
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
          <p className="text-[10px] uppercase tracking-[0.15em] text-gray-500">Roster</p>
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
  homeGoals, awayGoals, homeCards, awayCards, homeFouls, awayFouls }) {
  const JsPDF = await getJsPDF();
  const doc = new JsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const now = new Date();
  const dateStr = now.toLocaleDateString("en-GB", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });

  doc.setFillColor(10, 15, 30);
  doc.rect(0, 0, W, H, "F");

  // Header
  doc.setFillColor(20, 28, 50);
  doc.rect(0, 0, W, 28, "F");
  doc.setDrawColor(234, 179, 8);
  doc.setLineWidth(0.5);
  doc.line(0, 28, W, 28);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(234, 179, 8);
  doc.text("IT LIGA++ MATCH REPORT", W / 2, 11, { align: "center" });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(156, 163, 175);
  doc.text(`${dateStr}  ·  ${timeStr}`, W / 2, 19, { align: "center" });
  doc.text("Generated by IT LIGA++ NIŠ", W / 2, 24, { align: "center" });

  // Score block
  doc.setFillColor(15, 23, 42);
  doc.roundedRect(15, 34, W - 30, 34, 3, 3, "F");
  doc.setDrawColor(55, 65, 81);
  doc.setLineWidth(0.4);
  doc.roundedRect(15, 34, W - 30, 34, 3, 3, "S");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(200, 200, 200);
  doc.text(homeTeam.toUpperCase(), W / 2 - 28, 43, { align: "right" });
  doc.text(awayTeam.toUpperCase(), W / 2 + 28, 43, { align: "left" });
  doc.setFontSize(30);
  doc.setTextColor(255, 255, 255);
  doc.text(String(homeGoals.length), W / 2 - 14, 60, { align: "center" });
  doc.text(String(awayGoals.length), W / 2 + 14, 60, { align: "center" });
  doc.setFontSize(20);
  doc.setTextColor(100, 116, 139);
  doc.text("–", W / 2, 60, { align: "center" });
  doc.setFontSize(7);
  doc.setTextColor(107, 114, 128);
  doc.text(`Fouls: ${homeFouls}/5`, W / 2 - 28, 64, { align: "right" });
  doc.text(`Fouls: ${awayFouls}/5`, W / 2 + 28, 64, { align: "left" });

  doc.setDrawColor(55, 65, 81);
  doc.setLineWidth(0.3);
  doc.line(15, 76, W - 15, 76);

  let y = 82;
  const colLeft = 15;
  const colRight = W / 2 + 4;
  const colW = W / 2 - 20;

  const renderTeamSection = (startX, teamName, players, goals, cards) => {
    let ty = y;

    // Section header
    doc.setFillColor(30, 42, 70);
    doc.rect(startX, ty - 4, colW, 10, "F");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(234, 179, 8);
    doc.text(teamName.toUpperCase(), startX + colW / 2, ty + 2.5, { align: "center" });
    ty += 10;

    // Column headers
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7);
    doc.setTextColor(156, 163, 175);
    doc.text("#", startX + 2, ty);
    doc.text("Player", startX + 12, ty);
    doc.text("G", startX + colW - 14, ty, { align: "right" });
    doc.text("Y", startX + colW - 8, ty, { align: "right" });
    doc.text("R", startX + colW - 2, ty, { align: "right" });
    ty += 3;
    doc.setDrawColor(55, 65, 81);
    doc.line(startX, ty, startX + colW, ty);
    ty += 4;

    const active = players.filter((p) => p.num || p.name);
    if (active.length === 0) {
      doc.setFont("helvetica", "italic");
      doc.setTextColor(75, 85, 99);
      doc.text("No players entered", startX + colW / 2, ty + 2, { align: "center" });
      ty += 8;
    } else {
      active.forEach((p, idx) => {
        const gc = goals.filter((g) => g.playerId === p.id).length;
        const pc = cards.filter((c) => c.playerId === p.id);
        const yc = pc.filter((c) => c.type === "yellow").length;
        const rc = pc.filter((c) => c.type === "red").length;
        if (idx % 2 === 0) {
          doc.setFillColor(18, 26, 45);
          doc.rect(startX, ty - 3, colW, 6, "F");
        }
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7.5);
        doc.setTextColor(200, 210, 230);
        doc.text(p.num || "–", startX + 2, ty);
        doc.text(p.name || "–", startX + 12, ty);
        // Goals
        if (gc > 0) { doc.setTextColor(134, 239, 172); doc.setFont("helvetica", "bold"); }
        else { doc.setTextColor(75, 85, 99); doc.setFont("helvetica", "normal"); }
        doc.text(String(gc), startX + colW - 14, ty, { align: "right" });
        // Yellow cards
        if (yc > 0) { doc.setTextColor(250, 204, 21); doc.setFont("helvetica", "bold"); }
        else { doc.setTextColor(75, 85, 99); doc.setFont("helvetica", "normal"); }
        doc.text(String(yc), startX + colW - 8, ty, { align: "right" });
        // Red cards
        if (rc > 0) { doc.setTextColor(239, 68, 68); doc.setFont("helvetica", "bold"); }
        else { doc.setTextColor(75, 85, 99); doc.setFont("helvetica", "normal"); }
        doc.text(String(rc), startX + colW - 2, ty, { align: "right" });
        ty += 5.5;
      });
    }

    // Events log
    const allEvents = [
      ...goals.map((g) => ({ ...g, etype: "goal" })),
      ...cards.map((c) => ({ ...c, etype: "card" })),
    ].sort((a, b) => (a.half - b.half) || (a.elapsedSecs - b.elapsedSecs));

    if (allEvents.length > 0) {
      ty += 2;
      doc.setDrawColor(55, 65, 81);
      doc.line(startX, ty, startX + colW, ty);
      ty += 5;
      doc.setFont("helvetica", "bold");
      doc.setFontSize(7.5);
      doc.setTextColor(234, 179, 8);
      doc.text("Match Events", startX, ty);
      ty += 4.5;

      let lastHalf = null;
      allEvents.forEach((ev) => {
        // Ubaci oznaku poluvremena kada se promeni
        if (ev.half !== lastHalf) {
          doc.setFont("helvetica", "bold");
          doc.setFontSize(6.5);
          doc.setTextColor(ev.half === 1 ? 96 : 251, ev.half === 1 ? 165 : 146, ev.half === 1 ? 250 : 60);
          doc.text(`── ${ev.half === 1 ? "1. Poluvreme" : "2. Poluvreme"} ──`, startX, ty);
          ty += 4;
          lastHalf = ev.half;
        }

        const player = players.find((p) => p.id === ev.playerId);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7);
        const halfLabel = ev.half === 2 ? " [P2]" : " [P1]";
        if (ev.etype === "goal") {
          doc.setTextColor(134, 239, 172);
          doc.text(`${ev.clockMin}'${halfLabel}  G  #${player?.num || "?"} ${player?.name || ev.snapName || "Unknown"}`, startX, ty);
        } else {
          const isRed = ev.type === "red";
          doc.setTextColor(isRed ? 239 : 250, isRed ? 68 : 204, isRed ? 68 : 21);
          const label = isRed ? (ev.auto ? "RC(2Y→R)" : "RC") : "YC";
          doc.text(`${ev.clockMin}'${halfLabel}  ${label}  #${player?.num || "?"} ${player?.name || ev.snapName || "Unknown"}`, startX, ty);
        }
        ty += 4.5;
      });
    }

    return ty;
  };

  const endLeft = renderTeamSection(colLeft, homeTeam, homePlayers, homeGoals, homeCards);
  const endRight = renderTeamSection(colRight, awayTeam, awayPlayers, awayGoals, awayCards);
  let finalY = Math.max(endLeft, endRight) + 8;

  doc.setDrawColor(55, 65, 81);
  doc.setLineWidth(0.3);
  doc.line(15, finalY, W - 15, finalY);
  finalY += 5;
  doc.setFont("helvetica", "italic");
  doc.setFontSize(7);
  doc.setTextColor(75, 85, 99);
  doc.text("IT LIGA++", W / 2, finalY, { align: "center" });

  const safeHome = homeTeam.replace(/[^a-z0-9]/gi, "-").toLowerCase();
  const safeAway = awayTeam.replace(/[^a-z0-9]/gi, "-").toLowerCase();
  doc.save(`futsal-${safeHome}-vs-${safeAway}-${Date.now()}.pdf`);
}

// ── Login Screen ──────────────────────────────────────────────────────────────
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

  const [running, setRunning] = useState(false); // never auto-resume running clock
  const [secs, setSecs] = useState(s?.secs ?? DEFAULT_HALF_SECS);
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

  // Auto-save to localStorage on every relevant change
  useEffect(() => {
    saveMatchToStorage({
      secs, half, homeTeam, awayTeam, homePlayers, awayPlayers,
      homeGoals, awayGoals, homeCards, awayCards, homeFouls, awayFouls, matchLocked,
    });
  }, [secs, half, homeTeam, awayTeam, homePlayers, awayPlayers,
      homeGoals, awayGoals, homeCards, awayCards, homeFouls, awayFouls, matchLocked]);

  // Hide "restored" message after a few seconds
  useEffect(() => {
    if (restoredMsg) {
      const t = setTimeout(() => setRestoredMsg(""), 4000);
      return () => clearTimeout(t);
    }
  }, [restoredMsg]);

  // Timer
  useEffect(() => {
    if (running) {
      timerRef.current = setInterval(() => {
        setSecs((prev) => {
          if (prev <= 1) { clearInterval(timerRef.current); setRunning(false); return 0; }
          return prev - 1;
        });
      }, 1000);
    } else {
      clearInterval(timerRef.current);
    }
    return () => clearInterval(timerRef.current);
  }, [running]);

  const toggleTimer = () => { if (secs > 0) setRunning((r) => !r); };
  const resetTimer = () => { setRunning(false); setSecs(DEFAULT_HALF_SECS); };


  // Goal recording
  const recordGoal = useCallback((side, player) => {
    const elapsed = DEFAULT_HALF_SECS - secsRef.current;
    const minuteInHalf = Math.ceil(elapsed / 60) || 1;
    // Kumulativni minut: u 2. poluvremenu dodajemo trajanje 1. poluvremena
    const cumulativeMin = halfRef.current === 1
      ? minuteInHalf
      : Math.floor(DEFAULT_HALF_SECS / 60) + minuteInHalf;
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
    const elapsed = DEFAULT_HALF_SECS - secsRef.current;
    const minuteInHalf = Math.ceil(elapsed / 60) || 1;
    const cumulativeMin = halfRef.current === 1
      ? minuteInHalf
      : Math.floor(DEFAULT_HALF_SECS / 60) + minuteInHalf;

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
        homeFouls, awayFouls,
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
                  setSecs(DEFAULT_HALF_SECS);
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
          </div>
        </div>

        {/* Away */}
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
          onSave={(newSecs) => { setSecs(newSecs); setShowEditClock(false); }}
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