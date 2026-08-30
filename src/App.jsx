import * as XLSX from "xlsx";
import { useState, useEffect, useMemo, useRef } from "react";
import {
  Camera, Mic, Lightbulb, Plus, X, Check, AlertTriangle,
  Package, Users, ClipboardList, LayoutGrid, ChevronDown,
  Trash2, Calendar, Clock, Search, Folder, CalendarDays,
  Battery, Triangle, Joystick, Aperture, Rows3, StickyNote
} from "lucide-react";

const MESI_IT = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
const GIORNI_IT = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"];
const EVENT_PALETTE = [
  "#E1523D", "#F2A93B", "#D9C24E", "#8FB93F", "#3FB6A8", "#3E9BD6",
  "#7C7FE8", "#B168D6", "#E0629E", "#E88A5A", "#5FA8E0", "#6FB07A",
  "#C9645A", "#9AA85E", "#5B9EA6", "#A87FD1",
];
/* Colore stabile per evento (stesso colore in Calendario e nella card
   dell'evento), calcolato dall'id — non dipende dall'ordine dell'elenco */
function getEventColor(eventId) {
  // FNV-1a: distribuisce bene anche ID quasi identici (es. eventi creati a
  // pochi millisecondi di distanza), a differenza di un hash "somma*31"
  // che con timestamp ravvicinati produceva pattern ripetitivi.
  let hash = 0x811c9dc5;
  for (let i = 0; i < eventId.length; i++) {
    hash ^= eventId.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return EVENT_PALETTE[(hash >>> 0) % EVENT_PALETTE.length];
}

/* ---------------------------------------------------------
   TOKENS — "sala regia": tavolo di regia, gaffer tape, tally light
--------------------------------------------------------- */
const TOKENS = {
  bg: "#17191A",
  panel: "#1F2224",
  panelRaised: "#262A2C",
  line: "#33383A",
  amber: "#F2A93B",
  teal: "#3FB6A8",
  red: "#E1523D",
  text: "#EDEAE3",
  textMute: "#9AA0A3",
};

const CATEGORY_META = {
  camera: { label: "Videocamera", icon: Camera, color: TOKENS.amber },
  microfono: { label: "Microfono", icon: Mic, color: TOKENS.teal },
  luce: { label: "Luce", icon: Lightbulb, color: "#D9C24E" },
  batterie: { label: "Batterie", icon: Battery, color: "#8FB93F" },
  cavalletti: { label: "Cavalletti", icon: Triangle, color: "#8A97A6" },
  ronin: { label: "Ronin", icon: Joystick, color: "#B168D6" },
  obiettivi: { label: "Obiettivi", icon: Aperture, color: "#5FA8E0" },
  vario: { label: "Vario", icon: Package, color: "#C9645A" },
};

/* Per l'export/import Excel usiamo la stessa etichetta che si vede nel menù
   a tendina dell'app (es. "Videocamera"), non la chiave tecnica interna
   (es. "camera"), per evitare confusione a chi modifica il file. Questa
   mappatura permette di riconoscere la categoria in fase di importazione
   sia dall'etichetta che, per tolleranza, dalla vecchia chiave tecnica. */
const CATEGORY_LABEL_TO_KEY = Object.fromEntries(
  Object.entries(CATEGORY_META).map(([key, meta]) => [meta.label.toLowerCase(), key])
);
function resolveCategoryFromImport(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "vario";
  const byLabel = CATEGORY_LABEL_TO_KEY[raw.toLowerCase()];
  if (byLabel) return byLabel;
  const byKey = Object.keys(CATEGORY_META).find((k) => k.toLowerCase() === raw.toLowerCase());
  return byKey || "vario";
}

const STATUS_META = {
  disponibile: { label: "Disponibile", color: TOKENS.teal },
  assegnato: { label: "In uso ora", color: TOKENS.amber },
  manutenzione: { label: "Manutenzione", color: TOKENS.red },
};

/* ---------------------------------------------------------
   DATI DI ESEMPIO (in memoria — nessun salvataggio reale)
   item.status: solo "disponibile" | "manutenzione" (manuale).
   Lo stato "in uso ora" è calcolato dalle date/orari degli eventi.
--------------------------------------------------------- */
const INITIAL_ITEMS = [
  { id: "CAM-014", name: "Sony FX6", category: "camera", status: "disponibile", note: "" },
  { id: "CAM-015", name: "Sony FX6", category: "camera", status: "disponibile", note: "" },
  { id: "CAM-021", name: "Canon C70", category: "camera", status: "manutenzione", note: "Sensore da pulire, in officina" },
  { id: "MIC-003", name: "Rode NTG5", category: "microfono", status: "disponibile", note: "" },
  { id: "MIC-007", name: "Sennheiser G4", category: "microfono", status: "disponibile", note: "" },
  { id: "MIC-011", name: "Zoom H6", category: "microfono", status: "disponibile", note: "" },
  { id: "LUC-002", name: "Aputure 300D", category: "luce", status: "disponibile", note: "" },
  { id: "LUC-006", name: "Aputure 300D", category: "luce", status: "disponibile", note: "" },
  { id: "LUC-009", name: "Nanlite Pavotube", category: "luce", status: "disponibile", note: "" },
];

const INITIAL_CAMERAMEN = [
  { id: "cm-1", name: "Marco Rossi" },
  { id: "cm-2", name: "Giulia Bianchi" },
  { id: "cm-3", name: "Luca Ferrari" },
];

const INITIAL_EVENTS = [
  { id: "ev-1", name: "Matrimonio Villa Erba", cameramanId: "cm-1", fromDate: "2026-08-29", fromTime: "09:00", toDate: "2026-08-29", toTime: "23:00" },
  { id: "ev-2", name: "Intervista aziendale", cameramanId: "cm-2", fromDate: "2026-08-29", fromTime: "14:00", toDate: "2026-08-29", toTime: "16:00" },
];

const INITIAL_ASSIGNMENTS = [
  { id: "a1", itemId: "CAM-015", eventId: "ev-1" },
  { id: "a2", itemId: "MIC-007", eventId: "ev-1" },
  { id: "a3", itemId: "LUC-002", eventId: "ev-2" },
];

/* ---------------------------------------------------------
   HELPER — date/orari e sovrapposizioni
--------------------------------------------------------- */
function toDateTime(dateStr, timeStr, fallback) {
  if (!dateStr) return null;
  return new Date(`${dateStr}T${timeStr || fallback}`);
}
function eventRange(ev) {
  return {
    from: toDateTime(ev.fromDate, ev.fromTime, "00:00"),
    to: toDateTime(ev.toDate || ev.fromDate, ev.toTime, "23:59"),
  };
}
function rangesOverlap(aFrom, aTo, bFrom, bTo) {
  if (!aFrom || !aTo || !bFrom || !bTo) return false;
  return aFrom <= bTo && bFrom <= aTo;
}
function formatEventWhen(ev) {
  const sameDay = ev.toDate === ev.fromDate || !ev.toDate;
  if (sameDay) {
    return `${ev.fromDate}${ev.fromTime ? ` · ${ev.fromTime}` : ""}${ev.toTime ? ` → ${ev.toTime}` : ""}`;
  }
  return `${ev.fromDate}${ev.fromTime ? ` ${ev.fromTime}` : ""} → ${ev.toDate}${ev.toTime ? ` ${ev.toTime}` : ""}`;
}

function dateOnly(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function sameDate(a, b) {
  return a.getTime() === b.getTime();
}
/* Costruisce le settimane (lun-dom) necessarie a coprire un mese, incluse
   le code dei mesi adiacenti, come nella vista mensile di Google Calendar */
function getMonthMatrix(year, month) {
  const firstOfMonth = new Date(year, month, 1);
  const lastOfMonth = new Date(year, month + 1, 0);
  const startWeekday = (firstOfMonth.getDay() + 6) % 7; // 0 = lunedì
  const cursor = new Date(year, month, 1 - startWeekday);
  const weeks = [];
  while (true) {
    const week = [];
    for (let d = 0; d < 7; d++) {
      week.push(new Date(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
    weeks.push(week);
    if (cursor > lastOfMonth) break;
  }
  return weeks;
}
/* Individua tutti i mesi (anno+mese) attraversati da almeno un evento */
function getMonthsWithEvents(events) {
  const set = new Map();
  events.forEach((ev) => {
    if (!ev.fromDate) return;
    const from = dateOnly(new Date(`${ev.fromDate}T00:00`));
    const to = dateOnly(new Date(`${ev.toDate || ev.fromDate}T00:00`));
    const cursor = new Date(from.getFullYear(), from.getMonth(), 1);
    const last = new Date(to.getFullYear(), to.getMonth(), 1);
    while (cursor <= last) {
      const key = `${cursor.getFullYear()}-${cursor.getMonth()}`;
      set.set(key, { year: cursor.getFullYear(), month: cursor.getMonth() });
      cursor.setMonth(cursor.getMonth() + 1);
    }
  });
  return Array.from(set.values()).sort((a, b) => a.year - b.year || a.month - b.month);
}
/* Per ogni settimana calcola le "barre" evento (con eventuale accatastamento
   su più righe se più eventi si sovrappongono negli stessi giorni) */
function computeWeekBars(week, events, cameramanName) {
  const weekStart = dateOnly(week[0]);
  const weekEnd = dateOnly(week[6]);
  const overlapping = events
    .filter((ev) => ev.fromDate)
    .map((ev) => {
      const evFrom = dateOnly(new Date(`${ev.fromDate}T00:00`));
      const evTo = dateOnly(new Date(`${ev.toDate || ev.fromDate}T00:00`));
      const start = evFrom > weekStart ? evFrom : weekStart;
      const end = evTo < weekEnd ? evTo : weekEnd;
      if (start > end) return null;
      const startCol = week.findIndex((d) => sameDate(dateOnly(d), start));
      const endCol = week.findIndex((d) => sameDate(dateOnly(d), end));
      return {
        event: ev,
        startCol,
        endCol,
        continuesBefore: evFrom < weekStart,
        continuesAfter: evTo > weekEnd,
        cameraman: cameramanName(ev.cameramanId),
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.startCol - b.startCol);

  const rowEnds = [];
  const bars = overlapping.map((bar) => {
    let rowIndex = rowEnds.findIndex((end) => end < bar.startCol);
    if (rowIndex === -1) {
      rowIndex = rowEnds.length;
      rowEnds.push(bar.endCol);
    } else {
      rowEnds[rowIndex] = bar.endCol;
    }
    return { ...bar, rowIndex };
  });
  return bars;
}

/* Singola barra evento nel calendario: colore proprio + tooltip al passaggio
   del mouse con l'elenco del materiale prenotato per quell'evento */
function EventBar({ bar, style, materialForEvent }) {
  const [hover, setHover] = useState(false);
  const color = getEventColor(bar.event.id);
  const material = materialForEvent(bar.event.id);

  return (
    <div style={{ minWidth: 0, minHeight: 0, ...style, position: "relative" }} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      <div
        style={{
          background: `${color}CC`,
          color: "#161616",
          fontSize: 16,
          fontWeight: 700,
          padding: "1px 6px",
          height: "100%",
          borderTopLeftRadius: bar.continuesBefore ? 0 : 4,
          borderBottomLeftRadius: bar.continuesBefore ? 0 : 4,
          borderTopRightRadius: bar.continuesAfter ? 0 : 4,
          borderBottomRightRadius: bar.continuesAfter ? 0 : 4,
          whiteSpace: "nowrap",
          overflow: "hidden",
          textOverflow: "ellipsis",
          cursor: "default",
        }}
      >
        {bar.event.name} · {bar.cameraman || "nessun cameraman"}
      </div>

      {hover && (
        <div
          style={{
            position: "absolute", top: "100%", left: 0, marginTop: 4, zIndex: 50,
            background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`,
            borderRadius: 7, padding: "10px 12px", minWidth: 190,
            boxShadow: "0 10px 26px rgba(0,0,0,0.45)",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
            <div style={{ width: 9, height: 9, borderRadius: 3, background: color }} />
            <div style={{ fontWeight: 700, fontSize: 18 }}>{bar.event.name}</div>
          </div>
          {bar.cameraman ? (
            <div style={{ display: "inline-flex", alignItems: "center", gap: 5, marginBottom: 9, padding: "3px 9px", background: `${color}22`, border: `1px solid ${color}55`, borderRadius: 20 }}>
              <Users size={12} color={color} />
              <span style={{ fontSize: 14, fontWeight: 700, color: TOKENS.text }}>{bar.cameraman}</span>
            </div>
          ) : (
            <div style={{ display: "inline-flex", alignItems: "center", gap: 5, marginBottom: 9, padding: "3px 9px", background: `${TOKENS.red}22`, border: `1px solid ${TOKENS.red}55`, borderRadius: 20 }}>
              <AlertTriangle size={12} color={TOKENS.red} />
              <span style={{ fontSize: 14, fontWeight: 700, color: TOKENS.red }}>Cameraman non assegnato</span>
            </div>
          )}
          {material.length === 0 ? (
            <div style={{ fontSize: 15, color: TOKENS.textMute }}>Nessun materiale assegnato.</div>
          ) : (
            <ul style={{ margin: 0, padding: 0, listStyle: "none", display: "flex", flexDirection: "column", gap: 3 }}>
              {material.map(({ item }) => (
                <li key={item.id} style={{ fontSize: 15, display: "flex", gap: 6 }}>
                  <span style={{ fontFamily: "ui-monospace, monospace", color: TOKENS.textMute, fontSize: 13 }}>{item.id}</span>
                  <span>{item.name}</span>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
    </div>
  );
}

/* Vista mensile di sola visualizzazione, in stile Google Calendar:
   ogni evento ha un colore proprio (non più uno fisso per mese) così
   gli eventi che si susseguono o si sovrappongono si distinguono a colpo
   d'occhio; al passaggio del mouse mostra il materiale prenotato. */
function MonthCalendar({ year, month, events, cameramanName, materialForEvent }) {
  const weeks = getMonthMatrix(year, month);

  return (
    <div style={{ background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderRadius: 8, padding: 16, marginBottom: 18 }}>
      <div style={{ fontSize: 26, fontWeight: 800, marginBottom: 12 }}>{MESI_IT[month]} {year}</div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 4 }}>
        {GIORNI_IT.map((g) => (
          <div key={g} style={{ fontSize: 16, fontWeight: 700, color: TOKENS.textMute, textTransform: "uppercase", letterSpacing: "0.05em", textAlign: "center", padding: "2px 0" }}>
            {g}
          </div>
        ))}
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
        {weeks.map((week, wi) => {
          const bars = computeWeekBars(week, events, cameramanName);
          const maxRow = bars.reduce((m, b) => Math.max(m, b.rowIndex), -1);
          return (
            <div
              key={wi}
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(7, 1fr)",
                gridTemplateRows: maxRow >= 0 ? `20px repeat(${maxRow + 1}, 19px)` : "20px",
                gap: 3,
                position: "relative",
                background: TOKENS.panelRaised,
                borderRadius: 5,
                padding: 4,
              }}
            >
              {week.map((day, di) => (
                <div
                  key={di}
                  style={{
                    gridColumn: di + 1,
                    gridRow: 1,
                    minWidth: 0,
                    fontSize: 16,
                    color: day.getMonth() === month ? TOKENS.textMute : "#55595B",
                    fontWeight: day.getMonth() === month ? 700 : 400,
                    textAlign: "right",
                    paddingRight: 3,
                  }}
                >
                  {day.getDate()}
                </div>
              ))}
              {bars.map((bar, bi) => (
                <EventBar
                  key={bi}
                  bar={bar}
                  materialForEvent={materialForEvent}
                  style={{ gridColumn: `${bar.startCol + 1} / ${bar.endCol + 2}`, gridRow: bar.rowIndex + 2 }}
                />
              ))}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   PICCOLI COMPONENTI (tutti a livello di modulo — MAI ridefiniti
   dentro App, altrimenti perdono lo stato/focus ad ogni render)
--------------------------------------------------------- */
function Tag({ color, children }) {
  return (
    <span
      style={{
        display: "inline-flex", alignItems: "center", gap: 6,
        padding: "3px 9px", borderRadius: 3, fontSize: 16, fontWeight: 600,
        letterSpacing: "0.04em", textTransform: "uppercase", color,
        border: `1px solid ${color}55`, background: `${color}14`,
        fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
      }}
    >
      {children}
    </span>
  );
}

function GearChip({ item, onRemove }) {
  const meta = CATEGORY_META[item.category];
  const Icon = meta.icon;
  return (
    <div
      style={{
        display: "flex", alignItems: "center", gap: 7,
        background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`,
        borderLeft: `3px solid ${meta.color}`, borderRadius: 5,
        padding: "6px 8px 6px 10px", fontSize: 17.5,
      }}
    >
      <Icon size={13} color={meta.color} strokeWidth={2} />
      <span style={{ fontFamily: "ui-monospace, monospace", color: TOKENS.textMute, fontSize: 16 }}>{item.id}</span>
      <span style={{ fontWeight: 600 }}>{item.name}</span>
      {onRemove && (
        <button onClick={onRemove} title="Rimuovi dall'evento" style={{ background: "transparent", border: "none", color: TOKENS.textMute, cursor: "pointer", padding: 2, display: "flex" }}>
          <X size={12} />
        </button>
      )}
    </div>
  );
}

function GearTag({ item, status }) {
  const meta = CATEGORY_META[item.category];
  const statusMeta = STATUS_META[status];
  const Icon = meta.icon;
  return (
    <div style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderLeft: `3px solid ${meta.color}`, borderRadius: 4, padding: "12px 14px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <div style={{ fontFamily: "ui-monospace, monospace", fontSize: 17, letterSpacing: "0.08em", color: TOKENS.textMute }}>{item.id}</div>
          <div style={{ fontSize: 20, fontWeight: 600, color: TOKENS.text, marginTop: 2 }}>{item.name}</div>
        </div>
        <Icon size={18} color={meta.color} strokeWidth={1.75} />
      </div>
      <div style={{ marginTop: 10, display: "flex", gap: 6, flexWrap: "wrap" }}>
        <Tag color={statusMeta.color}>{statusMeta.label}</Tag>
        <Tag color={TOKENS.textMute}>{meta.label}</Tag>
      </div>
      {item.note && (
        <div style={{ display: "flex", alignItems: "flex-start", gap: 5, marginTop: 9, padding: "6px 8px", background: `${TOKENS.amber}14`, border: `1px solid ${TOKENS.amber}40`, borderRadius: 5 }}>
          <StickyNote size={13} color={TOKENS.amber} style={{ flexShrink: 0, marginTop: 1 }} />
          <span style={{ fontSize: 14, color: TOKENS.text, lineHeight: 1.3 }}>{item.note}</span>
        </div>
      )}
    </div>
  );
}

/* Password condivise per ruolo — non sono una vera misura di sicurezza
   (chiunque sappia leggere il codice del sito le trova), servono solo ad
   evitare accessi/errori casuali (es. un cameraman che clicca per sbaglio
   su "Responsabile"). Per cambiarle, modifica semplicemente questi due
   valori e ripubblica l'app. */
const RESPONSABILE_PASSWORD = "sky-responsabile-2026";
const CAMERAMAN_PASSWORD = "sky-cameraman-2026";

function LoginScreen({ onLogin }) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  function handleSubmit(e) {
    e.preventDefault();
    if (password === RESPONSABILE_PASSWORD) onLogin("responsabile");
    else if (password === CAMERAMAN_PASSWORD) onLogin("cameraman");
    else setError("Password non corretta.");
  }

  return (
    <div style={{ minHeight: 600, display: "flex", alignItems: "center", justifyContent: "center", background: TOKENS.bg, borderRadius: 10, border: `1px solid ${TOKENS.line}` }}>
      <form
        onSubmit={handleSubmit}
        style={{ background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderRadius: 12, padding: 32, width: 320, display: "flex", flexDirection: "column", gap: 14 }}
      >
        <div style={{ textAlign: "center", marginBottom: 6 }}>
          <div style={{ fontSize: 24, fontWeight: 800, color: TOKENS.text }}>SkySportGear</div>
          <div style={{ fontSize: 15, color: TOKENS.textMute, marginTop: 4 }}>Inserisci la password per accedere</div>
        </div>
        <input
          type="password"
          autoFocus
          value={password}
          onChange={(e) => { setPassword(e.target.value); setError(""); }}
          placeholder="Password"
          style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 8, padding: "10px 12px", color: TOKENS.text, fontSize: 16 }}
        />
        {error && <div style={{ color: TOKENS.red, fontSize: 14 }}>{error}</div>}
        <button
          type="submit"
          style={{ background: TOKENS.amber, color: "#1A1A1A", border: "none", borderRadius: 8, padding: "10px 14px", fontWeight: 700, fontSize: 16, cursor: "pointer" }}
        >
          Accedi
        </button>
      </form>
    </div>
  );
}

function RoleSwitcher({ role, onLogout, cameramanId, setCameramanId, cameramen }) {
  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
      <div
        style={{
          padding: "7px 14px", borderRadius: 8, fontSize: 18, fontWeight: 700,
          textTransform: "capitalize", background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, color: TOKENS.amber,
        }}
      >
        {role}
      </div>
      {role === "cameraman" && cameramen.length > 0 && (
        <div style={{ position: "relative" }}>
          <select
            value={cameramanId}
            onChange={(e) => setCameramanId(e.target.value)}
            style={{ appearance: "none", background: TOKENS.panelRaised, color: TOKENS.text, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "7px 28px 7px 10px", fontSize: 18, cursor: "pointer" }}
          >
            {cameramen.map((c) => (<option key={c.id} value={c.id}>{c.name}</option>))}
          </select>
          <ChevronDown size={14} color={TOKENS.textMute} style={{ position: "absolute", right: 8, top: 9, pointerEvents: "none" }} />
        </div>
      )}
      <button
        onClick={onLogout}
        title="Esci e torna alla schermata di accesso"
        style={{ background: "transparent", border: `1px solid ${TOKENS.line}`, color: TOKENS.textMute, borderRadius: 6, padding: "7px 12px", fontSize: 15, cursor: "pointer" }}
      >
        Esci
      </button>
    </div>
  );
}

function NavButton({ active, onClick, icon: Icon, label }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "10px 12px",
        borderRadius: 6, border: "none", background: active ? TOKENS.panelRaised : "transparent",
        color: active ? TOKENS.amber : TOKENS.textMute, fontSize: 18.5, fontWeight: 600,
        cursor: "pointer", textAlign: "left",
        borderLeft: active ? `3px solid ${TOKENS.amber}` : "3px solid transparent",
      }}
    >
      <Icon size={16} strokeWidth={2} />
      {label}
    </button>
  );
}

/* Card che rappresenta un evento con tutto il materiale accorpato */
function EventCard({ event, items, availableForThisEvent, cameramanLabel, onAddItem, onRemoveItem, onDeleteEvent, readOnly }) {
  const [adding, setAdding] = useState(false);
  const [pick, setPick] = useState("");
  const color = getEventColor(event.id);

  return (
    <div style={{ background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderLeft: `4px solid ${color}`, borderRadius: 8, padding: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 11, height: 11, borderRadius: 3, background: color, flexShrink: 0 }} title="Colore evento nel calendario" />
            <span style={{ fontSize: 20, fontWeight: 700 }}>{event.name}</span>
          </div>
          {cameramanLabel ? (
            <div style={{ display: "inline-flex", alignItems: "center", gap: 5, marginTop: 6, padding: "3px 9px", background: `${color}22`, border: `1px solid ${color}55`, borderRadius: 20 }}>
              <Users size={13} color={color} />
              <span style={{ fontSize: 15, fontWeight: 700, color: TOKENS.text }}>{cameramanLabel}</span>
            </div>
          ) : (
            <div style={{ display: "inline-flex", alignItems: "center", gap: 5, marginTop: 6, padding: "3px 9px", background: `${TOKENS.red}22`, border: `1px solid ${TOKENS.red}55`, borderRadius: 20 }}>
              <AlertTriangle size={13} color={TOKENS.red} />
              <span style={{ fontSize: 15, fontWeight: 700, color: TOKENS.red }}>Cameraman non assegnato</span>
            </div>
          )}
          <div style={{ fontSize: 17, color: TOKENS.textMute, marginTop: 6, display: "flex", gap: 12, flexWrap: "wrap" }}>
            <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <Calendar size={12} /> {formatEventWhen(event)}
            </span>
          </div>
        </div>
        {!readOnly && onDeleteEvent && (
          <button onClick={() => onDeleteEvent(event.id)} title="Elimina evento e libera il materiale" style={{ background: "transparent", border: `1px solid ${TOKENS.line}`, borderRadius: 5, color: TOKENS.red, padding: "5px 8px", cursor: "pointer" }}>
            <Trash2 size={13} />
          </button>
        )}
      </div>

      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 12 }}>
        {items.length === 0 && <span style={{ fontSize: 17.5, color: TOKENS.textMute }}>Nessun materiale in questo evento.</span>}
        {items.map(({ assignment, item }) => (
          <GearChip key={assignment.id} item={item} onRemove={readOnly ? null : () => onRemoveItem(assignment.id)} />
        ))}
      </div>

      {!readOnly && onAddItem && (
        <div style={{ marginTop: 12 }}>
          {!adding ? (
            <button onClick={() => setAdding(true)} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: `1px dashed ${TOKENS.line}`, color: TOKENS.textMute, borderRadius: 6, padding: "6px 10px", fontSize: 17, cursor: "pointer" }}>
              <Plus size={12} /> Aggiungi materiale a questo evento
            </button>
          ) : (
            <div style={{ display: "flex", gap: 6 }}>
              <select value={pick} onChange={(e) => setPick(e.target.value)} style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "6px 8px", color: TOKENS.text, fontSize: 17.5, flex: 1 }}>
                <option value="">Materiale libero in queste date…</option>
                {availableForThisEvent.map((i) => (<option key={i.id} value={i.id}>{i.id} — {i.name}</option>))}
              </select>
              <button onClick={() => { if (pick) { onAddItem(event.id, pick); setPick(""); setAdding(false); } }} style={{ background: TOKENS.amber, color: "#1A1A1A", border: "none", borderRadius: 6, padding: "6px 10px", fontSize: 17, fontWeight: 700, cursor: "pointer" }}>
                Aggiungi
              </button>
              <button onClick={() => { setAdding(false); setPick(""); }} style={{ background: "transparent", border: `1px solid ${TOKENS.line}`, color: TOKENS.textMute, borderRadius: 6, padding: "6px 8px", fontSize: 17, cursor: "pointer" }}>
                Annulla
              </button>
              {availableForThisEvent.length === 0 && (
                <div style={{ fontSize: 16.5, color: TOKENS.red, alignSelf: "center" }}>Nessun materiale libero per queste date/orari.</div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* Form "assegna materiale a un evento" — componente stabile a livello di modulo:
   definirlo dentro App ne causava la ricreazione ad ogni render, con perdita del
   focus sugli input a ogni tasto premuto e chiusura dei date-picker nativi. */
function EventAssignForm({ forCameramanId, eventsPool, cameramen, cameramanName, eventForm, setEventForm, emptyEventForm, getAvailableItems, onSubmit }) {
  const showCameramanPicker = !forCameramanId;
  const selectedExistingEvent = eventForm.mode === "existing" ? eventsPool.find((e) => e.id === eventForm.eventId) : null;

  const availableItems =
    eventForm.mode === "existing"
      ? selectedExistingEvent
        ? getAvailableItems(selectedExistingEvent.fromDate, selectedExistingEvent.fromTime, selectedExistingEvent.toDate, selectedExistingEvent.toTime, selectedExistingEvent.id)
        : []
      : getAvailableItems(eventForm.fromDate, eventForm.fromTime, eventForm.toDate, eventForm.toTime, null);

  return (
    <div style={{ background: TOKENS.panel, border: `1px dashed ${TOKENS.line}`, borderRadius: 8, padding: 14, marginBottom: 20 }}>
      <div style={{ fontSize: 17.5, fontWeight: 700, color: TOKENS.textMute, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.05em" }}>
        Assegna materiale a un evento
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
        <button
          onClick={() => setEventForm({ ...emptyEventForm, mode: "new" })}
          style={{
            padding: "5px 12px", borderRadius: 6, fontSize: 17, fontWeight: 600, cursor: "pointer",
            border: `1px solid ${eventForm.mode === "new" ? TOKENS.amber : TOKENS.line}`,
            background: eventForm.mode === "new" ? `${TOKENS.amber}1A` : "transparent",
            color: eventForm.mode === "new" ? TOKENS.amber : TOKENS.textMute,
          }}
        >
          Nuovo evento
        </button>
        <button
          onClick={() => setEventForm({ ...emptyEventForm, mode: "existing" })}
          disabled={eventsPool.length === 0}
          style={{
            padding: "5px 12px", borderRadius: 6, fontSize: 17, fontWeight: 600,
            cursor: eventsPool.length === 0 ? "not-allowed" : "pointer",
            opacity: eventsPool.length === 0 ? 0.5 : 1,
            border: `1px solid ${eventForm.mode === "existing" ? TOKENS.amber : TOKENS.line}`,
            background: eventForm.mode === "existing" ? `${TOKENS.amber}1A` : "transparent",
            color: eventForm.mode === "existing" ? TOKENS.amber : TOKENS.textMute,
          }}
        >
          Evento esistente
        </button>
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        {eventForm.mode === "existing" ? (
          <select
            value={eventForm.eventId}
            onChange={(e) => setEventForm({ ...eventForm, eventId: e.target.value })}
            style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, minWidth: 200 }}
          >
            <option value="">Scegli evento…</option>
            {eventsPool.map((e) => (
              <option key={e.id} value={e.id}>{e.name}{!forCameramanId ? ` — ${cameramanName(e.cameramanId)}` : ""}</option>
            ))}
          </select>
        ) : (
          <>
            <input
              placeholder="Nome evento (es. Matrimonio Villa Erba)"
              value={eventForm.name}
              onChange={(e) => setEventForm({ ...eventForm, name: e.target.value })}
              style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, minWidth: 190 }}
            />
            {showCameramanPicker && (
              <select
                value={eventForm.cameramanId}
                onChange={(e) => setEventForm({ ...eventForm, cameramanId: e.target.value })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18 }}
              >
                <option value="">Cameraman…</option>
                {cameramen.map((c) => (<option key={c.id} value={c.id}>{c.name}</option>))}
              </select>
            )}
            <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <input type="date" value={eventForm.fromDate} onChange={(e) => setEventForm({ ...eventForm, fromDate: e.target.value })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18 }} />
              <input type="time" value={eventForm.fromTime} onChange={(e) => setEventForm({ ...eventForm, fromTime: e.target.value })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, width: 92 }} />
            </div>
            <span style={{ color: TOKENS.textMute, fontSize: 17 }}>→</span>
            <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <input type="date" value={eventForm.toDate} onChange={(e) => setEventForm({ ...eventForm, toDate: e.target.value })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18 }} />
              <input type="time" value={eventForm.toTime} onChange={(e) => setEventForm({ ...eventForm, toTime: e.target.value })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, width: 92 }} />
            </div>
          </>
        )}

        <select
          value={eventForm.itemId}
          onChange={(e) => setEventForm({ ...eventForm, itemId: e.target.value })}
          style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, minWidth: 190 }}
        >
          <option value="">Materiale libero in queste date…</option>
          {availableItems.map((i) => (<option key={i.id} value={i.id}>{i.id} — {i.name}</option>))}
        </select>

        <button onClick={onSubmit} style={{ display: "flex", alignItems: "center", gap: 6, background: TOKENS.amber, color: "#1A1A1A", border: "none", borderRadius: 6, padding: "8px 14px", fontWeight: 700, fontSize: 18, cursor: "pointer" }}>
          <Check size={14} /> Assegna
        </button>
      </div>
      {(eventForm.mode === "new" ? eventForm.fromDate : selectedExistingEvent) && availableItems.length === 0 && (
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10, fontSize: 17, color: TOKENS.red }}>
          <AlertTriangle size={13} /> Nessun materiale libero per queste date/orari.
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   APP
--------------------------------------------------------- */
/* Database condiviso (Firebase Realtime Database): qui vengono letti e
   scritti i dati reali, visibili a tutti quelli che usano l'app. */
const FIREBASE_DATA_URL = "https://skysportgear-default-rtdb.europe-west1.firebasedatabase.app/skysportgear.json";

/* Valorizzata da Vite al momento della build (vedi vite.config.js e il
   workflow GitHub Actions), è diversa ad ogni pubblicazione. Serve per
   accorgersi quando è uscita una versione più recente dell'app, senza
   dover chiedere all'utente di fare un "refresh forzato" a mano. */
const CURRENT_APP_VERSION = typeof __APP_VERSION__ !== "undefined" ? __APP_VERSION__ : "dev";

/* Come useState, ma salva automaticamente il valore nel localStorage del
   browser e lo ricarica al successivo avvio: così i dati sopravvivono al
   refresh della pagina e alla chiusura del browser (sullo stesso dispositivo).
   Fa da cache locale/di riserva: il vero dato condiviso vive su Firebase
   (vedi gli effect dentro App più sotto). */
function usePersistentState(key, initialValue) {
  const [state, setState] = useState(() => {
    try {
      const stored = window.localStorage.getItem(key);
      return stored ? JSON.parse(stored) : initialValue;
    } catch {
      return initialValue;
    }
  });
  useEffect(() => {
    try {
      window.localStorage.setItem(key, JSON.stringify(state));
    } catch {
      // localStorage non disponibile (es. modalità privata): l'app continua
      // a funzionare, semplicemente senza salvataggio persistente.
    }
  }, [key, state]);
  return [state, setState];
}

export default function App() {
  const [items, setItems] = usePersistentState("skysportgear_items", INITIAL_ITEMS);
  const [cameramen, setCameramen] = usePersistentState("skysportgear_cameramen", INITIAL_CAMERAMEN);
  const [events, setEvents] = usePersistentState("skysportgear_events", INITIAL_EVENTS);
  const [assignments, setAssignments] = usePersistentState("skysportgear_assignments", INITIAL_ASSIGNMENTS);

  const [authRole, setAuthRole] = useState(() => {
    try { return window.localStorage.getItem("skysportgear_auth_role") || null; } catch { return null; }
  });
  useEffect(() => {
    try {
      if (authRole) window.localStorage.setItem("skysportgear_auth_role", authRole);
      else window.localStorage.removeItem("skysportgear_auth_role");
    } catch {}
  }, [authRole]);
  const role = authRole;
  function handleLogout() {
    setAuthRole(null);
  }
  const [cameramanId, setCameramanId] = useState(INITIAL_CAMERAMEN[0].id);
  const [tab, setTab] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [materialView, setMaterialView] = useState("grid");
  const excelInputRef = useRef(null);
  const [selectedDashboardStatus, setSelectedDashboardStatus] = useState(null);
  const [newItem, setNewItem] = useState({ id: "", name: "", category: "camera" });
  const [newCameraman, setNewCameraman] = useState("");
  const [toast, setToast] = useState(null);

  const emptyEventForm = { mode: "new", eventId: "", name: "", cameramanId: "", fromDate: "", fromTime: "", toDate: "", toTime: "", itemId: "" };
  const [eventForm, setEventForm] = useState(emptyEventForm);

  const [syncStatus, setSyncStatus] = useState("connessione"); // connessione | pronto | in-corso | offline
  const [updateAvailable, setUpdateAvailable] = useState(false);

  /* Controlla periodicamente (e ogni volta che si torna su questa scheda)
     se è stata pubblicata una versione più recente dell'app, confrontando
     con un piccolo file che GitHub Actions rigenera ad ogni build. Evita
     di dover spiegare agli utenti come fare un "refresh forzato" a mano. */
  useEffect(() => {
    function checkForUpdate() {
      fetch(`./version.json?t=${Date.now()}`, { cache: "no-store" })
        .then((res) => res.json())
        .then((data) => {
          if (data?.version && data.version !== CURRENT_APP_VERSION) {
            setUpdateAvailable(true);
          }
        })
        .catch(() => {}); // se il file non c'è (es. in sviluppo locale), ignora silenziosamente
    }
    checkForUpdate();
    const interval = setInterval(checkForUpdate, 2 * 60 * 1000); // ogni 2 minuti
    function onVisible() {
      if (document.visibilityState === "visible") checkForUpdate();
    }
    document.addEventListener("visibilitychange", onVisible);
    return () => {
      clearInterval(interval);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, []);
  const [lastSyncAt, setLastSyncAt] = useState(null);
  const didLoadRef = useRef(false);

  /* Legge un blocco dati da Firebase, trattando in modo esplicito le liste
     mancanti come "vuote": Firebase non conserva gli array vuoti (li
     cancella), quindi l'assenza di una chiave qui significa "lista svuotata
     di proposito", non "nessun dato ancora salvato". */
  function normalizeRemote(remote) {
    return {
      items: remote?.items || [],
      cameramen: remote?.cameramen || [],
      events: remote?.events || [],
      assignments: remote?.assignments || [],
    };
  }

  /* Controlla se, tra le assegnazioni presenti sul server ma non ancora
     viste in locale, ce n'è qualcuna che usa lo stesso materiale in un
     periodo che si sovrappone a un'assegnazione fatta qui in locale: è il
     caso classico di due persone che assegnano lo stesso pezzo nello stesso
     momento. Restituisce un elenco di conflitti leggibili, vuoto se nessuno. */
  function findBookingConflicts(remote) {
    const conflicts = [];
    const remoteEvents = remote.events || [];
    const localEventsById = new Map(events.map((e) => [e.id, e]));
    const remoteEventsById = new Map(remoteEvents.map((e) => [e.id, e]));

    assignments.forEach((localA) => {
      const localEvent = localEventsById.get(localA.eventId);
      if (!localEvent) return;
      const localRange = eventRange(localEvent);

      (remote.assignments || []).forEach((remoteA) => {
        if (remoteA.itemId !== localA.itemId) return;
        if (remoteA.eventId === localA.eventId) return; // stessa assegnazione, non è un conflitto
        const remoteEvent = remoteEventsById.get(remoteA.eventId);
        if (!remoteEvent) return;
        const remoteRange = eventRange(remoteEvent);
        if (rangesOverlap(localRange.from, localRange.to, remoteRange.from, remoteRange.to)) {
          const already = conflicts.some((c) => c.itemId === localA.itemId && c.remoteEventId === remoteA.eventId);
          if (!already) {
            conflicts.push({
              itemId: localA.itemId,
              itemName: items.find((i) => i.id === localA.itemId)?.name || localA.itemId,
              localEventName: localEvent.name,
              remoteEventName: remoteEvent.name,
              remoteEventId: remoteA.eventId,
            });
          }
        }
      });
    });
    return conflicts;
  }

  /* Al primo avvio, scarica i dati condivisi da Firebase (se presenti) e
     sostituisce quelli locali/di esempio. Da qui in poi la sincronizzazione
     è sempre manuale (pulsanti "Carica" e "Condividi"), mai automatica. */
  useEffect(() => {
    let cancelled = false;
    fetch(FIREBASE_DATA_URL)
      .then((res) => res.json())
      .then((remote) => {
        if (cancelled) return;
        if (remote) {
          const n = normalizeRemote(remote);
          setItems(n.items);
          setCameramen(n.cameramen);
          setEvents(n.events);
          setAssignments(n.assignments);
        }
        setSyncStatus("pronto");
        setLastSyncAt(new Date());
        didLoadRef.current = true;
      })
      .catch(() => {
        setSyncStatus("offline");
        didLoadRef.current = true;
      });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* Pulsante "Carica dati condivisi": sostituisce lo stato locale con
     l'ultima versione salvata da chiunque altro. Le modifiche locali non
     ancora condivise andrebbero perse, quindi chiede conferma. */
  function pullSharedData() {
    if (window.confirm("Caricare gli ultimi dati condivisi? Eventuali modifiche fatte qui e non ancora condivise andranno perse.") === false) return;
    setSyncStatus("in-corso");
    fetch(FIREBASE_DATA_URL)
      .then((res) => res.json())
      .then((remote) => {
        const n = normalizeRemote(remote);
        setItems(n.items);
        setCameramen(n.cameramen);
        setEvents(n.events);
        setAssignments(n.assignments);
        setSyncStatus("pronto");
        setLastSyncAt(new Date());
        showToast("Dati condivisi caricati.");
      })
      .catch(() => {
        setSyncStatus("offline");
        showToast("Impossibile raggiungere il database condiviso.");
      });
  }

  /* Pulsante "Condividi le mie modifiche": prima controlla eventuali
     conflitti di prenotazione materiale avvenuti nel frattempo su altri
     dispositivi; se ne trova, blocca l'invio e li segnala chiaramente
     invece di sovrascrivere silenziosamente. */
  function pushSharedData() {
    setSyncStatus("in-corso");
    fetch(FIREBASE_DATA_URL)
      .then((res) => res.json())
      .then((remoteRaw) => {
        const remote = normalizeRemote(remoteRaw);
        const conflicts = findBookingConflicts(remote);
        if (conflicts.length > 0) {
          setSyncStatus("pronto");
          const details = conflicts
            .map((c) => `• ${c.itemName} (${c.itemId}): assegnato qui a "${c.localEventName}", ma nel frattempo anche a "${c.remoteEventName}" da qualcun altro, con date che si sovrappongono`)
            .join("\n");
          window.alert(
            "Impossibile condividere: c'è un conflitto di prenotazione materiale.\n\n" +
            details +
            "\n\nCarica prima i dati condivisi (pulsante 'Carica dati condivisi'), correggi l'assegnazione in conflitto, poi riprova a condividere."
          );
          return;
        }
        return fetch(FIREBASE_DATA_URL, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ items, cameramen, events, assignments }),
        }).then(() => {
          setSyncStatus("pronto");
          setLastSyncAt(new Date());
          showToast("Modifiche condivise con tutti.");
        });
      })
      .catch(() => {
        setSyncStatus("offline");
        showToast("Impossibile raggiungere il database condiviso.");
      });
  }


  function showToast(msg) {
    setToast(msg);
    setTimeout(() => setToast(null), 2500);
  }

  const cameramanName = (id) => cameramen.find((c) => c.id === id)?.name || null;

  function itemsForEvent(eventId) {
    return assignments
      .filter((a) => a.eventId === eventId)
      .map((assignment) => ({ assignment, item: items.find((i) => i.id === assignment.itemId) }))
      .filter((x) => x.item);
  }

  /* Materiale libero per un dato intervallo data/ora, escludendo eventualmente
     l'evento che si sta modificando (per non "auto-bloccarsi" il proprio materiale) */
  function getAvailableItems(fromDate, fromTime, toDate, toTime, excludeEventId) {
    if (!fromDate) return items.filter((i) => i.status !== "manutenzione");
    const from = toDateTime(fromDate, fromTime, "00:00");
    const to = toDateTime(toDate || fromDate, toTime, "23:59");
    return items.filter((i) => {
      if (i.status === "manutenzione") return false;
      const clash = assignments.some((a) => {
        if (a.itemId !== i.id) return false;
        if (a.eventId === excludeEventId) return false;
        const ev = events.find((e) => e.id === a.eventId);
        if (!ev) return false;
        const r = eventRange(ev);
        return rangesOverlap(from, to, r.from, r.to);
      });
      return !clash;
    });
  }

  function isCurrentlyInUse(itemId) {
    const now = new Date();
    return assignments.some((a) => {
      if (a.itemId !== itemId) return false;
      const ev = events.find((e) => e.id === a.eventId);
      if (!ev) return false;
      const r = eventRange(ev);
      return r.from && r.to && r.from <= now && now <= r.to;
    });
  }

  function computeStatus(item) {
    if (item.status === "manutenzione") return "manutenzione";
    return isCurrentlyInUse(item.id) ? "assegnato" : "disponibile";
  }

  const counts = useMemo(() => {
    const byStatus = { disponibile: 0, assegnato: 0, manutenzione: 0 };
    items.forEach((i) => byStatus[computeStatus(i)]++);
    return byStatus;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items, assignments, events]);

  const filteredItems = items.filter((i) => i.name.toLowerCase().includes(search.toLowerCase()) || i.id.toLowerCase().includes(search.toLowerCase()));

  /* --- Crea/aggiorna evento e assegna un materiale (responsabile o self-service) --- */
  function submitEventAssignment(forCameramanId) {
    const { mode, eventId, name, fromDate, fromTime, toDate, toTime, itemId } = eventForm;
    if (!itemId) { showToast("Scegli un materiale."); return; }

    let targetEventId = eventId;

    if (mode === "new") {
      if (!name.trim() || !fromDate) { showToast("Dai un nome all'evento e una data di inizio."); return; }
      const camId = forCameramanId || eventForm.cameramanId;
      if (!camId) { showToast("Scegli il cameraman."); return; }
      targetEventId = "ev-" + Date.now();
      setEvents((prev) => [...prev, {
        id: targetEventId, name: name.trim(), cameramanId: camId,
        fromDate, fromTime: fromTime || "00:00",
        toDate: toDate || fromDate, toTime: toTime || "23:59",
      }]);
    } else if (!eventId) {
      showToast("Scegli un evento esistente.");
      return;
    }

    const assignId = "a-" + Date.now();
    setAssignments((prev) => [...prev, { id: assignId, itemId, eventId: targetEventId }]);
    setEventForm(emptyEventForm);
    showToast("Materiale assegnato all'evento.");
  }

  function addItemToEvent(eventId, itemId) {
    const assignId = "a-" + Date.now();
    setAssignments((prev) => [...prev, { id: assignId, itemId, eventId }]);
    showToast("Materiale aggiunto all'evento.");
  }

  function removeItemFromEvent(assignmentId) {
    const a = assignments.find((x) => x.id === assignmentId);
    if (!a) return;
    setAssignments((prev) => prev.filter((x) => x.id !== assignmentId));
    showToast(`${a.itemId} rientrato in magazzino.`);
  }

  function deleteEvent(eventId) {
    const event = events.find((e) => e.id === eventId);
    if (!event) return;

    if (role === "cameraman") {
      const ok = window.confirm(`Vuoi davvero eliminare l'evento "${event.name}"? Il materiale assegnato tornerà disponibile.`);
      if (!ok) return;
    } else {
      const r = eventRange(event);
      const now = new Date();
      const isOngoing = r.from && r.to && r.from <= now && now <= r.to;
      const isFuture = r.from && r.from > now;
      if (isOngoing) {
        const ok = window.confirm(
          `Attenzione: l'evento "${event.name}" è attualmente IN CORSO.\n\nSei sicuro di volerlo eliminare? Il materiale assegnato tornerà disponibile immediatamente.`
        );
        if (!ok) return;
      } else if (isFuture) {
        const ok = window.confirm(
          `L'evento "${event.name}" è programmato per il futuro (${formatEventWhen(event)}).\n\nSei sicuro di volerlo eliminare?`
        );
        if (!ok) return;
      }
      // evento già concluso: nessuna conferma richiesta
    }

    setAssignments((prev) => prev.filter((a) => a.eventId !== eventId));
    setEvents((prev) => prev.filter((e) => e.id !== eventId));
    showToast("Evento chiuso, materiale rientrato.");
  }

  function setItemManualStatus(itemId, status) {
    setItems((prev) => prev.map((i) => (i.id === itemId ? { ...i, status } : i)));
  }

  function setItemNote(itemId, note) {
    setItems((prev) => prev.map((i) => (i.id === itemId ? { ...i, note } : i)));
  }

  function addItem() {
    if (!newItem.id || !newItem.name) { showToast("Inserisci codice e nome del materiale."); return; }
    if (items.some((i) => i.id === newItem.id)) { showToast("Codice già esistente."); return; }
    setItems([...items, { ...newItem, status: "disponibile", note: "" }]);
    setNewItem({ id: "", name: "", category: "camera" });
    showToast("Materiale aggiunto al magazzino.");
  }

  function removeItem(id) {
    setAssignments((prev) => prev.filter((a) => a.itemId !== id));
    setItems((prev) => prev.filter((i) => i.id !== id));
  }

  function exportItemsToExcel() {
    const rows = items.map((i) => ({
      Codice: i.id,
      Nome: i.name,
      Categoria: CATEGORY_META[i.category]?.label || i.category,
      Stato: STATUS_META[i.status]?.label || i.status,
      Nota: i.note || "",
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Materiale");
    XLSX.writeFile(workbook, "materiale-skysportgear.xlsx");
  }

  /* Importa da un file Excel: aggiorna gli oggetti con Codice già esistente
     e aggiunge quelli nuovi, senza mai cancellare pezzi non presenti nel
     file (per evitare perdite di dati accidentali). Riconosce la categoria
     sia dall'etichetta leggibile (es. "Videocamera") sia, per tolleranza,
     dalla vecchia chiave tecnica (es. "camera"). */
  function importItemsFromExcel(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);
        let added = 0, updated = 0, skipped = 0;

        setItems((prev) => {
          const map = new Map(prev.map((i) => [i.id, i]));
          rows.forEach((row) => {
            const codice = String(row.Codice ?? row.codice ?? "").trim();
            if (!codice) { skipped++; return; }
            const category = resolveCategoryFromImport(row.Categoria ?? row.categoria);
            const status = String(row.Stato ?? "").trim().toLowerCase() === "manutenzione" ? "manutenzione" : "disponibile";
            const newItem = {
              id: codice,
              name: String(row.Nome ?? row.nome ?? "").trim() || codice,
              category,
              status,
              note: String(row.Nota ?? row.nota ?? ""),
            };
            if (map.has(codice)) updated++; else added++;
            map.set(codice, newItem);
          });
          return Array.from(map.values());
        });

        showToast(`Importazione completata: ${added} aggiunti, ${updated} aggiornati${skipped ? `, ${skipped} righe senza codice ignorate` : ""}.`);
      } catch (err) {
        showToast("Il file non sembra un Excel valido (colonne attese: Codice, Nome, Categoria, Stato, Nota).");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function addCameraman() {
    if (!newCameraman.trim()) return;
    setCameramen((prev) => [...prev, { id: "cm-" + Date.now(), name: newCameraman.trim() }]);
    setNewCameraman("");
    showToast("Cameraman aggiunto.");
  }

  function deleteCameraman(id) {
    const theirEvents = events.filter((e) => e.cameramanId === id);
    if (theirEvents.length > 0) {
      const cam = cameramen.find((c) => c.id === id);
      const elenco = theirEvents.map((e) => `"${e.name}"`).join(", ");
      const ok = window.confirm(
        `Attenzione: ${cam?.name || "questo cameraman"} ha ${theirEvents.length} evento/i assegnato/i (${elenco}).\n\nEliminandolo, questi eventi verranno chiusi e il relativo materiale tornerà disponibile.\n\nVuoi procedere comunque?`
      );
      if (!ok) return;
    }
    const theirEventIds = theirEvents.map((e) => e.id);
    setAssignments((prev) => prev.filter((a) => !theirEventIds.includes(a.eventId)));
    setEvents((prev) => prev.filter((e) => e.cameramanId !== id));
    const remaining = cameramen.filter((c) => c.id !== id);
    setCameramen(remaining);
    if (cameramanId === id && remaining.length > 0) setCameramanId(remaining[0].id);
    showToast("Cameraman eliminato, suoi eventi chiusi.");
  }

  const canManage = role === "responsabile";
  const myEvents = events.filter((e) => e.cameramanId === cameramanId);

  if (!role) {
    return <LoginScreen onLogin={setAuthRole} />;
  }

  const NAV = [
    { key: "dashboard", label: "Dashboard", icon: LayoutGrid, roles: ["responsabile", "cameraman"] },
    { key: "calendario", label: "Calendario", icon: CalendarDays, roles: ["responsabile", "cameraman"] },
    { key: "materiale", label: "Materiale", icon: Package, roles: ["responsabile"] },
    { key: "eventi", label: "Eventi", icon: ClipboardList, roles: ["responsabile"] },
    { key: "cameramen", label: "Cameraman", icon: Users, roles: ["responsabile"] },
    { key: "mie", label: "I miei eventi", icon: Folder, roles: ["cameraman"] },
  ];
  const visibleNav = NAV.filter((n) => n.roles.includes(role));
  const activeTab = visibleNav.some((n) => n.key === tab) ? tab : visibleNav[0].key;

  return (
    <div style={{ display: "flex", minHeight: 600, background: TOKENS.bg, color: TOKENS.text, fontFamily: "'Inter','Helvetica Neue',Arial,sans-serif", borderRadius: 10, overflow: "hidden", border: `1px solid ${TOKENS.line}` }}>
      {/* SIDEBAR */}
      <div style={{ width: 200, background: TOKENS.panel, borderRight: `1px solid ${TOKENS.line}`, padding: "18px 12px", display: "flex", flexDirection: "column", gap: 4 }}>
        <div style={{ padding: "0 8px 18px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 8, height: 8, borderRadius: "50%", background: TOKENS.red }} />
            <span style={{ fontSize: 16, letterSpacing: "0.12em", color: TOKENS.textMute, fontWeight: 700 }}>ON AIR</span>
          </div>
          <div style={{ fontSize: 21, fontWeight: 800, marginTop: 6, letterSpacing: "-0.01em" }}>SkySportGear</div>
          <div style={{ fontSize: 16, color: TOKENS.textMute, marginTop: 2 }}>gestione materiale</div>
        </div>
        {visibleNav.map((n) => (
          <NavButton key={n.key} active={activeTab === n.key} onClick={() => { setTab(n.key); setEventForm(emptyEventForm); }} icon={n.icon} label={n.label} />
        ))}
        <div style={{ flex: 1 }} />
        <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "0 4px 8px", fontSize: 12, color: TOKENS.textMute }}>
          <div
            style={{
              width: 7, height: 7, borderRadius: "50%",
              background: syncStatus === "pronto" ? TOKENS.teal : syncStatus === "in-corso" ? TOKENS.amber : syncStatus === "offline" ? TOKENS.red : TOKENS.textMute,
              flexShrink: 0,
            }}
          />
          {syncStatus === "pronto" && lastSyncAt && `Aggiornato alle ${lastSyncAt.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" })}`}
          {syncStatus === "in-corso" && "Sincronizzazione…"}
          {syncStatus === "offline" && "Offline: solo su questo dispositivo"}
          {syncStatus === "connessione" && "Connessione…"}
        </div>
        <button
          onClick={pullSharedData}
          title="Scarica gli ultimi dati condivisi da tutti (sovrascrive le modifiche locali non ancora condivise)"
          style={{ background: "transparent", border: `1px solid ${TOKENS.line}`, color: TOKENS.text, borderRadius: 6, padding: "8px 10px", fontSize: 13, cursor: "pointer", marginTop: 4 }}
        >
          ⭳ Carica dati condivisi
        </button>
        <button
          onClick={pushSharedData}
          title="Condividi le modifiche fatte qui con tutti gli altri (controlla prima eventuali conflitti sul materiale)"
          style={{ background: TOKENS.amber, border: "none", color: "#1A1A1A", borderRadius: 6, padding: "8px 10px", fontSize: 13, fontWeight: 700, cursor: "pointer", marginTop: 6 }}
        >
          ⭱ Condividi le mie modifiche
        </button>
      </div>

      {/* MAIN */}
      <div style={{ flex: 1, padding: "20px 26px", overflow: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
          <div>
            <div style={{ fontSize: 25, fontWeight: 800 }}>{visibleNav.find((n) => n.key === activeTab)?.label}</div>
            <div style={{ fontSize: 17.5, color: TOKENS.textMute, marginTop: 2 }}>
              {role === "cameraman" ? `Visualizzazione come ${cameramanName(cameramanId)}` : `Visualizzazione: ${role}`}
            </div>
          </div>
          <RoleSwitcher role={role} onLogout={handleLogout} cameramanId={cameramanId} setCameramanId={setCameramanId} cameramen={cameramen} />
        </div>

        {/* ---------------- DASHBOARD ---------------- */}
        {activeTab === "dashboard" && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 14 }}>
              {Object.entries(STATUS_META).map(([key, meta]) => {
                const active = selectedDashboardStatus === key;
                return (
                  <button
                    key={key}
                    onClick={() => setSelectedDashboardStatus(active ? null : key)}
                    style={{
                      textAlign: "left", cursor: "pointer", background: TOKENS.panel,
                      border: `1px solid ${active ? meta.color : TOKENS.line}`, borderRadius: 8, padding: "16px 18px",
                      boxShadow: active ? `0 0 0 1px ${meta.color}` : "none",
                    }}
                  >
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ width: 8, height: 8, borderRadius: "50%", background: meta.color }} />
                      <span style={{ fontSize: 17, color: TOKENS.textMute, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em" }}>{meta.label}</span>
                    </div>
                    <div style={{ fontSize: 37, fontWeight: 800, marginTop: 8, fontFamily: "ui-monospace, monospace" }}>{counts[key]}</div>
                  </button>
                );
              })}
            </div>

            {selectedDashboardStatus && (
              <div style={{ background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderRadius: 8, padding: 16, marginBottom: 22 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                  <span style={{ fontSize: 16, fontWeight: 700, color: TOKENS.textMute, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                    Materiale — {STATUS_META[selectedDashboardStatus].label}
                  </span>
                  <button onClick={() => setSelectedDashboardStatus(null)} style={{ background: "transparent", border: "none", color: TOKENS.textMute, cursor: "pointer", display: "flex" }}>
                    <X size={16} />
                  </button>
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {items.filter((i) => computeStatus(i) === selectedDashboardStatus).map((item) => (
                    <div key={item.id} title={item.note || undefined}>
                      <GearChip item={item} />
                    </div>
                  ))}
                  {items.filter((i) => computeStatus(i) === selectedDashboardStatus).length === 0 && (
                    <span style={{ fontSize: 14, color: TOKENS.textMute }}>Nessun materiale in questo stato.</span>
                  )}
                </div>
              </div>
            )}

            <div style={{ fontSize: 18, fontWeight: 700, color: TOKENS.textMute, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>
              Eventi programmati
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {events.length === 0 && <div style={{ color: TOKENS.textMute, fontSize: 18 }}>Nessun evento attivo.</div>}
              {events.map((ev) => (
                <EventCard key={ev.id} event={ev} items={itemsForEvent(ev.id)} cameramanLabel={cameramanName(ev.cameramanId)} readOnly />
              ))}
            </div>
          </div>
        )}

        {/* ---------------- CALENDARIO (sola visualizzazione) ---------------- */}
        {activeTab === "calendario" && (
          <div>
            {role === "cameraman" && (
              <div style={{ fontSize: 15, color: TOKENS.textMute, marginBottom: 14 }}>
                Vedi gli eventi di tutti i cameraman, non solo i tuoi.
              </div>
            )}
            {(() => {
              const months = getMonthsWithEvents(events);
              if (months.length === 0) {
                return <div style={{ color: TOKENS.textMute, fontSize: 18 }}>Nessun evento da mostrare in calendario.</div>;
              }
              return months.map(({ year, month }) => (
                <MonthCalendar key={`${year}-${month}`} year={year} month={month} events={events} cameramanName={cameramanName} materialForEvent={itemsForEvent} />
              ));
            })()}
          </div>
        )}

        {/* ---------------- MATERIALE ---------------- */}
        {activeTab === "materiale" && canManage && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 18, alignItems: "center", justifyContent: "space-between", flexWrap: "wrap" }}>
              <div style={{ position: "relative", flex: 1, maxWidth: 320 }}>
                <Search size={14} color={TOKENS.textMute} style={{ position: "absolute", left: 10, top: 10 }} />
                <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Cerca per nome o codice…"
                  style={{ width: "100%", background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px 8px 32px", color: TOKENS.text, fontSize: 18 }} />
              </div>
              <div style={{ display: "flex", gap: 4, background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderRadius: 8, padding: 4 }}>
                <button
                  onClick={() => setMaterialView("grid")}
                  title="Vista a riquadri"
                  style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", borderRadius: 6, border: "none", fontSize: 16, fontWeight: 600, cursor: "pointer", background: materialView === "grid" ? TOKENS.amber : "transparent", color: materialView === "grid" ? "#1A1A1A" : TOKENS.textMute }}
                >
                  <LayoutGrid size={14} /> Riquadri
                </button>
                <button
                  onClick={() => setMaterialView("list")}
                  title="Vista a lista"
                  style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", borderRadius: 6, border: "none", fontSize: 16, fontWeight: 600, cursor: "pointer", background: materialView === "list" ? TOKENS.amber : "transparent", color: materialView === "list" ? "#1A1A1A" : TOKENS.textMute }}
                >
                  <Rows3 size={14} /> Lista
                </button>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button
                  onClick={exportItemsToExcel}
                  title="Scarica l'elenco materiale come file Excel"
                  style={{ display: "flex", alignItems: "center", gap: 6, background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, color: TOKENS.text, borderRadius: 8, padding: "9px 14px", fontSize: 16, fontWeight: 600, cursor: "pointer" }}
                >
                  Esporta Excel
                </button>
                <button
                  onClick={() => excelInputRef.current?.click()}
                  title="Importa/aggiorna materiale da un file Excel (colonne: Codice, Nome, Categoria, Stato, Nota)"
                  style={{ display: "flex", alignItems: "center", gap: 6, background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, color: TOKENS.text, borderRadius: 8, padding: "9px 14px", fontSize: 16, fontWeight: 600, cursor: "pointer" }}
                >
                  Importa Excel
                </button>
                <input
                  ref={excelInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  style={{ display: "none" }}
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) importItemsFromExcel(file);
                    e.target.value = "";
                  }}
                />
              </div>
            </div>

            <div style={{ display: "flex", gap: 8, marginBottom: 22, flexWrap: "wrap", background: TOKENS.panel, border: `1px dashed ${TOKENS.line}`, borderRadius: 8, padding: 12 }}>
              <input placeholder="Codice (es. CAM-030)" value={newItem.id} onChange={(e) => setNewItem({ ...newItem, id: e.target.value.toUpperCase() })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, width: 150 }} />
              <input placeholder="Nome / modello" value={newItem.name} onChange={(e) => setNewItem({ ...newItem, name: e.target.value })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, flex: 1, minWidth: 140 }} />
              <select value={newItem.category} onChange={(e) => setNewItem({ ...newItem, category: e.target.value })}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18 }}>
                {Object.entries(CATEGORY_META).map(([k, m]) => (<option key={k} value={k}>{m.label}</option>))}
              </select>
              <button onClick={addItem} style={{ display: "flex", alignItems: "center", gap: 6, background: TOKENS.amber, color: "#1A1A1A", border: "none", borderRadius: 6, padding: "8px 14px", fontWeight: 700, fontSize: 18, cursor: "pointer" }}>
                <Plus size={14} /> Aggiungi
              </button>
            </div>

            {Object.entries(CATEGORY_META).map(([catKey, catMeta]) => {
              const catItems = filteredItems.filter((i) => i.category === catKey);
              if (catItems.length === 0) return null;
              const CatIcon = catMeta.icon;
              return (
                <div key={catKey} style={{ marginBottom: 26 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                    <CatIcon size={16} color={catMeta.color} strokeWidth={2} />
                    <span style={{ fontSize: 16, fontWeight: 700, color: TOKENS.textMute, textTransform: "uppercase", letterSpacing: "0.05em" }}>{catMeta.label}</span>
                    <span style={{ fontSize: 14, color: TOKENS.textMute }}>({catItems.length})</span>
                  </div>

                  {materialView === "grid" ? (
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 10 }}>
                      {catItems.map((item) => (
                        <div key={item.id}>
                          <GearTag item={item} status={computeStatus(item)} />
                          <div style={{ display: "flex", gap: 4, marginTop: 6 }}>
                            <select
                              value={item.status}
                              onChange={(e) => setItemManualStatus(item.id, e.target.value)}
                              title="Stato manuale (il rientro 'in uso' è automatico in base agli eventi)"
                              style={{ flex: 1, fontSize: 15, background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, color: TOKENS.textMute, borderRadius: 5, padding: "4px 6px" }}
                            >
                              <option value="disponibile">Disponibile</option>
                              <option value="manutenzione">Manutenzione</option>
                            </select>
                            <button onClick={() => removeItem(item.id)} title="Rimuovi materiale" style={{ background: "transparent", border: `1px solid ${TOKENS.line}`, borderRadius: 5, color: TOKENS.red, padding: "4px 7px", cursor: "pointer" }}>
                              <Trash2 size={12} />
                            </button>
                          </div>
                          <input
                            value={item.note || ""}
                            onChange={(e) => setItemNote(item.id, e.target.value)}
                            placeholder="Nota (es. da controllare, graffio, ecc.)"
                            style={{ width: "100%", marginTop: 4, fontSize: 14, background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, color: TOKENS.text, borderRadius: 5, padding: "5px 7px", boxSizing: "border-box" }}
                          />
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {catItems.map((item) => {
                        const st = STATUS_META[computeStatus(item)];
                        return (
                          <div
                            key={item.id}
                            style={{
                              display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap",
                              background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderLeft: `3px solid ${catMeta.color}`,
                              borderRadius: 6, padding: "8px 12px",
                            }}
                          >
                            <span style={{ fontFamily: "ui-monospace, monospace", fontSize: 15, color: TOKENS.textMute, width: 90, flexShrink: 0 }}>{item.id}</span>
                            <span style={{ fontSize: 17, fontWeight: 600, minWidth: 150 }}>{item.name}</span>
                            <Tag color={st.color}>{st.label}</Tag>
                            <input
                              value={item.note || ""}
                              onChange={(e) => setItemNote(item.id, e.target.value)}
                              placeholder="Nota…"
                              style={{ flex: 1, minWidth: 140, fontSize: 14, background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, color: TOKENS.text, borderRadius: 5, padding: "5px 8px" }}
                            />
                            <select
                              value={item.status}
                              onChange={(e) => setItemManualStatus(item.id, e.target.value)}
                              title="Stato manuale (il rientro 'in uso' è automatico in base agli eventi)"
                              style={{ fontSize: 14, background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, color: TOKENS.textMute, borderRadius: 5, padding: "5px 6px" }}
                            >
                              <option value="disponibile">Disponibile</option>
                              <option value="manutenzione">Manutenzione</option>
                            </select>
                            <button onClick={() => removeItem(item.id)} title="Rimuovi materiale" style={{ background: "transparent", border: `1px solid ${TOKENS.line}`, borderRadius: 5, color: TOKENS.red, padding: "5px 8px", cursor: "pointer", flexShrink: 0 }}>
                              <Trash2 size={13} />
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })}
            {filteredItems.length === 0 && <div style={{ color: TOKENS.textMute, fontSize: 17 }}>Nessun materiale trovato.</div>}
          </div>
        )}

        {/* ---------------- EVENTI (responsabile) ---------------- */}
        {activeTab === "eventi" && canManage && (
          <div>
            <EventAssignForm
              forCameramanId={null}
              eventsPool={events}
              cameramen={cameramen}
              cameramanName={cameramanName}
              eventForm={eventForm}
              setEventForm={setEventForm}
              emptyEventForm={emptyEventForm}
              getAvailableItems={getAvailableItems}
              onSubmit={() => submitEventAssignment(null)}
            />
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {events.length === 0 && <div style={{ color: TOKENS.textMute, fontSize: 18 }}>Nessun evento creato.</div>}
              {events.map((ev) => (
                <EventCard
                  key={ev.id}
                  event={ev}
                  items={itemsForEvent(ev.id)}
                  availableForThisEvent={getAvailableItems(ev.fromDate, ev.fromTime, ev.toDate, ev.toTime, ev.id)}
                  cameramanLabel={cameramanName(ev.cameramanId)}
                  onAddItem={addItemToEvent}
                  onRemoveItem={removeItemFromEvent}
                  onDeleteEvent={deleteEvent}
                />
              ))}
            </div>
          </div>
        )}

        {/* ---------------- CAMERAMEN ---------------- */}
        {activeTab === "cameramen" && role === "responsabile" && (
          <div>
            <div style={{ display: "flex", gap: 8, marginBottom: 18 }}>
              <input placeholder="Nome cameraman" value={newCameraman} onChange={(e) => setNewCameraman(e.target.value)}
                style={{ background: TOKENS.panelRaised, border: `1px solid ${TOKENS.line}`, borderRadius: 6, padding: "8px 10px", color: TOKENS.text, fontSize: 18, width: 220 }} />
              <button onClick={addCameraman} style={{ display: "flex", alignItems: "center", gap: 6, background: TOKENS.amber, color: "#1A1A1A", border: "none", borderRadius: 6, padding: "8px 14px", fontWeight: 700, fontSize: 18, cursor: "pointer" }}>
                <Plus size={14} /> Aggiungi
              </button>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 10 }}>
              {cameramen.map((c) => {
                const evCount = events.filter((e) => e.cameramanId === c.id).length;
                return (
                  <div key={c.id} style={{ background: TOKENS.panel, border: `1px solid ${TOKENS.line}`, borderRadius: 8, padding: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                      <div>
                        <div style={{ fontWeight: 700, fontSize: 19.5 }}>{c.name}</div>
                        <div style={{ fontSize: 17, color: TOKENS.textMute, marginTop: 4 }}>{evCount} evento/i attivo/i</div>
                      </div>
                      <button
                        onClick={() => deleteCameraman(c.id)}
                        title="Elimina cameraman (chiude i suoi eventi e libera il materiale)"
                        style={{ background: "transparent", border: `1px solid ${TOKENS.line}`, borderRadius: 5, color: TOKENS.red, padding: "5px 7px", cursor: "pointer" }}
                      >
                        <Trash2 size={13} />
                      </button>
                    </div>
                  </div>
                );
              })}
              {cameramen.length === 0 && <div style={{ color: TOKENS.textMute, fontSize: 18 }}>Nessun cameraman in elenco.</div>}
            </div>
          </div>
        )}

        {/* ---------------- I MIEI EVENTI (cameraman) ---------------- */}
        {activeTab === "mie" && role === "cameraman" && (
          <div>
            {cameramen.length === 0 ? (
              <div style={{ color: TOKENS.textMute, fontSize: 18.5 }}>Nessun cameraman registrato.</div>
            ) : (
              <>
                <EventAssignForm
                  forCameramanId={cameramanId}
                  eventsPool={myEvents}
                  cameramen={cameramen}
                  cameramanName={cameramanName}
                  eventForm={eventForm}
                  setEventForm={setEventForm}
                  emptyEventForm={emptyEventForm}
                  getAvailableItems={getAvailableItems}
                  onSubmit={() => submitEventAssignment(cameramanId)}
                />
                <div style={{ fontSize: 18, fontWeight: 700, color: TOKENS.textMute, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>
                  I tuoi eventi
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {myEvents.length === 0 && <div style={{ color: TOKENS.textMute, fontSize: 18.5 }}>Nessun evento attivo al momento.</div>}
                  {myEvents.map((ev) => (
                    <EventCard
                      key={ev.id}
                      event={ev}
                      items={itemsForEvent(ev.id)}
                      availableForThisEvent={getAvailableItems(ev.fromDate, ev.fromTime, ev.toDate, ev.toTime, ev.id)}
                      cameramanLabel={cameramanName(ev.cameramanId)}
                      onAddItem={addItemToEvent}
                      onRemoveItem={removeItemFromEvent}
                      onDeleteEvent={deleteEvent}
                    />
                  ))}
                </div>
              </>
            )}
          </div>
        )}
      </div>

      {updateAvailable && (
        <div
          style={{
            position: "fixed", top: 0, left: 0, right: 0, zIndex: 100,
            display: "flex", alignItems: "center", justifyContent: "center", gap: 14,
            background: TOKENS.amber, color: "#1A1A1A", padding: "10px 16px", fontSize: 15, fontWeight: 700,
          }}
        >
          È disponibile una versione più recente di SkySportGear.
          <button
            onClick={() => window.location.reload()}
            style={{ background: "#1A1A1A", color: TOKENS.amber, border: "none", borderRadius: 6, padding: "6px 14px", fontWeight: 700, fontSize: 14, cursor: "pointer" }}
          >
            Aggiorna ora
          </button>
        </div>
      )}

      {toast && (
        <div style={{ position: "fixed", bottom: 20, left: "50%", transform: "translateX(-50%)", background: TOKENS.panelRaised, border: `1px solid ${TOKENS.amber}`, color: TOKENS.text, padding: "10px 18px", borderRadius: 8, fontSize: 18, boxShadow: "0 8px 24px rgba(0,0,0,0.4)" }}>
          {toast}
        </div>
      )}
    </div>
  );
}
