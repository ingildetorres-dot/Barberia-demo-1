/**
 * ============================================================================
 *  AUTHENTIC BARBERSHOP & MORE — Panel de operación
 * ============================================================================
 *  Agenda · Clientes · Barberos · Servicios · Productos · Gastos · Reportes
 *
 *  ARQUITECTURA (leer antes de modificar)
 *  --------------------------------------------------------------------------
 *  1) CAPA DE DATOS  ......... sección [DB]
 *     Todo pasa por loadDB() / mutate(). Es el ÚNICO punto que toca
 *     almacenamiento. Para migrar a PostgreSQL basta reimplementar esas dos
 *     funciones contra tu API REST (ver schema.sql y types.ts adjuntos).
 *     Ninguna vista lee ni escribe almacenamiento directamente.
 *
 *  2) REGLAS DE NEGOCIO ...... sección [LOGICA]
 *     Disponibilidad de horarios, bloqueos, ingresos, gastos. Funciones puras,
 *     sin React. Si cambias una, revisa quién la usa (se indica en cada JSDoc).
 *
 *  3) VISTAS ................. sección [VISTAS]
 *     Cada vista es un componente independiente que recibe { state, actions }.
 *     Agregar un módulo nuevo = crear componente + registrarlo en NAV.
 *     No hay dependencias cruzadas entre vistas.
 *
 *  4) PERMISOS ............... can(rol, permiso) en sección [PERMISOS]
 *
 *  CONVENCIONES
 *  - Fechas: 'YYYY-MM-DD' (string). Horas: 'HH:mm' (string, 24h).
 *  - Dinero: número entero en la moneda local (sin decimales).
 *  - Todo id: string generado por uid().
 *  - Al confirmar una cita se congela el precio del servicio en la cita
 *    (appointment.price) para que reportes históricos no cambien.
 * ============================================================================
 */

import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid,
  LineChart, Line, Legend,
} from "recharts";
import {
  doc, collection, query, where, getDoc, getDocs, setDoc, deleteDoc, deleteField,
  onSnapshot, writeBatch, runTransaction,
} from "firebase/firestore";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { db as firestore, auth } from "./firebase";
import {
  Calendar, Users, Scissors, Package, Wallet, BarChart3, Settings as Cog,
  Clock, Search, Plus, Check, X, Bell, Trash2, Pencil, Download, Printer,
  LogOut, Lock, MessageCircle, Ban, ChevronLeft, ChevronRight, Menu,
  AlertTriangle, UserPlus, Volume2, Save, RefreshCw, ListPlus, Wrench, ShoppingCart,
  CreditCard, Receipt, Minus, ShieldCheck, KeyRound,
} from "lucide-react";

/* El logo ya no está escrito en el código: cada negocio sube el suyo en Ajustes → Apariencia (ver [MARCA]). */

/* ==========================================================================
   [TOKENS] Estilos base. Los colores de abajo son los de la paleta «Negro y
   dorado»; la paleta elegida en Ajustes → Apariencia los reemplaza (ver [MARCA]).
   ========================================================================== */
const BASE_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600&family=Barlow:wght@400;500;600&family=Playfair+Display:wght@500;600;700&family=Poppins:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

.bt {
  --ink:#14161B; --ink-2:#1E2228; --line:#DCD7CD;
  --porcelain:#F4F2ED; --paper:#FFFFFF;
  --brass:#B08725; --brass-soft:#F0E4C4;
  --pole:#B5322E; --mint:#2F7D68; --slate:#5A6169;
  --display:'Oswald','Arial Narrow',ui-sans-serif,sans-serif;
  --body:'Barlow',ui-sans-serif,system-ui,sans-serif;
  --mono:'IBM Plex Mono',ui-monospace,monospace;
  font-family:var(--body); color:var(--ink); background:var(--porcelain);
  -webkit-font-smoothing:antialiased;
}
.bt *,.bt *::before,.bt *::after{ box-sizing:border-box; }

/* --- tipografía --- */
.bt-display{ font-family:var(--display); text-transform:uppercase; letter-spacing:.06em; font-weight:600; line-height:1.05; }
.bt-eyebrow{ font-family:var(--display); text-transform:uppercase; letter-spacing:.18em; font-size:11px; font-weight:500; color:var(--slate); }
.bt-mono{ font-family:var(--mono); font-variant-numeric:tabular-nums; }

/* --- superficies --- */
.bt-card{ background:var(--paper); border:1px solid var(--line); border-radius:2px; }
.bt-inset{ background:var(--porcelain); border:1px solid var(--line); border-radius:2px; }
.bt-hair{ border-color:var(--line); }

/* --- controles --- */
.bt-btn{ font-family:var(--display); text-transform:uppercase; letter-spacing:.08em; font-size:12px; font-weight:500;
  display:inline-flex; align-items:center; justify-content:center; gap:6px; padding:9px 14px; border-radius:2px;
  border:1px solid var(--ink,#14161B); background:var(--ink,#14161B); color:#fff; cursor:pointer; transition:opacity .15s, background .15s; }
.bt-btn:hover{ opacity:.86; }
.bt-btn:disabled{ opacity:.4; cursor:not-allowed; }
.bt-btn.ghost{ background:transparent; color:var(--ink,#14161B); border-color:var(--line,#DCD7CD); }
.bt-btn.ghost:hover{ background:#fff; border-color:var(--ink); }
.bt-btn.brass{ background:var(--brass,#B08725); border-color:var(--brass,#B08725); color:#fff; }
.bt-btn.danger{ background:transparent; color:var(--pole,#B5322E); border-color:var(--pole,#B5322E); }
.bt-btn.mint{ background:var(--mint,#2F7D68); border-color:var(--mint,#2F7D68); color:#fff; }
.bt-btn.sm{ padding:6px 10px; font-size:11px; }
.bt-icon{ display:inline-flex; align-items:center; justify-content:center; width:32px; height:32px; border-radius:2px;
  border:1px solid var(--line); background:var(--paper); color:var(--slate); cursor:pointer; }
.bt-icon:hover{ color:var(--ink); border-color:var(--ink); }

.bt-input, .bt-select, .bt-area{ width:100%; font-family:var(--body); font-size:14px; color:var(--ink);
  background:var(--paper); border:1px solid var(--line); border-radius:2px; padding:9px 11px; }
.bt-input:focus,.bt-select:focus,.bt-area:focus{ outline:2px solid var(--brass); outline-offset:-1px; border-color:var(--brass); }
.bt-label{ font-family:var(--display); text-transform:uppercase; letter-spacing:.12em; font-size:10px; font-weight:500; color:var(--slate); display:block; margin-bottom:5px; }

/* --- etiquetas de estado --- */
.bt-tag{ font-family:var(--display); text-transform:uppercase; letter-spacing:.1em; font-size:10px; font-weight:500;
  padding:3px 7px; border-radius:2px; border:1px solid currentColor; display:inline-block; }
.bt-tag.pend{ color:var(--brass); background:var(--brass-soft); }
.bt-tag.conf{ color:var(--mint); background:#E6F0EC; background:color-mix(in srgb, var(--mint) 14%, white); }
.bt-tag.done{ color:var(--slate); background:#EAE7E1; background:color-mix(in srgb, var(--slate) 12%, white); }
.bt-tag.canc{ color:var(--pole); background:#F6E3E2; background:color-mix(in srgb, var(--pole) 12%, white); }

/* --- FIRMA VISUAL: el poste de barbero como columna de horas de la agenda --- */
.bt-pole{ width:14px; flex:0 0 14px; border-left:1px solid var(--line); border-right:1px solid var(--line);
  background-image:repeating-linear-gradient(135deg,
    var(--pole) 0 7px, #fff 7px 14px, var(--ink) 14px 21px, #fff 21px 28px); }

/* --- ticket de comanda: la cita se ve como un talón perforado --- */
.bt-ticket{ position:relative; background:var(--paper); border:1px solid var(--line); border-left:3px solid var(--brass);
  border-radius:2px; padding:9px 11px; }
.bt-ticket.is-conf{ border-left-color:var(--mint); }
.bt-ticket.is-pend{ border-left-color:var(--brass); border-top:1px dashed var(--brass); }
.bt-ticket.is-canc{ border-left-color:var(--pole); opacity:.6; }
.bt-ticket.is-done{ border-left-color:var(--slate); }

/* --- navegación --- */
.bt-nav{ background:var(--ink); color:rgba(255,255,255,.72); }
.bt-navitem{ font-family:var(--display); text-transform:uppercase; letter-spacing:.1em; font-size:12px;
  display:flex; align-items:center; gap:10px; padding:11px 16px; width:100%; background:none; border:0;
  color:rgba(255,255,255,.72); cursor:pointer; border-left:2px solid transparent; text-align:left; }
.bt-navitem:hover{ color:#fff; background:var(--ink-2); }
.bt-navitem.on{ color:#fff; background:var(--ink-2); border-left-color:var(--brass); }
.bt-logo{ height:46px; width:auto; max-width:160px; object-fit:contain; }
.bt-logoplate{ display:inline-flex; background:var(--porcelain); border-radius:14px; padding:6px; }

/* --- avisos --- */
.bt-dot{ background:var(--pole); color:#fff; font-family:var(--mono); font-size:10px; font-weight:600;
  min-width:18px; height:18px; border-radius:9px; display:inline-flex; align-items:center; justify-content:center; padding:0 5px; }
.bt-pulse{ animation:btp 1.6s ease-in-out infinite; }
@keyframes btp{ 0%,100%{ transform:scale(1); } 50%{ transform:scale(1.14); } }
.bt-toast{ animation:btslide .22s ease-out; }
@keyframes btslide{ from{ transform:translateY(10px); opacity:0; } to{ transform:translateY(0); opacity:1; } }
@media (prefers-reduced-motion: reduce){ .bt-pulse,.bt-toast{ animation:none; } }

/* --- slots libres para reservar --- */
.bt-slot{ font-family:var(--mono); font-size:13px; padding:9px 4px; border:1px solid var(--line); background:var(--paper);
  border-radius:2px; cursor:pointer; text-align:center; }
.bt-slot:hover{ border-color:var(--ink); }
.bt-slot.on{ background:var(--ink); color:#fff; border-color:var(--ink); }
.bt-scroll{ overflow:auto; }
.bt-scroll::-webkit-scrollbar{ height:8px; width:8px; }
.bt-scroll::-webkit-scrollbar-thumb{ background:#CFC9BE; border-radius:4px; }
@media print{ .bt-noprint{ display:none !important; } }
`;
/** Estilos activos = base + paleta. applyBranding() lo reemplaza antes de pintar. */
let CSS = BASE_CSS;

/* ==========================================================================
   [UTILS] Funciones puras de uso general.
   ========================================================================== */
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const DOW = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
const DOW_S = ["Dom", "Lun", "Mar", "Mié", "Jue", "Vie", "Sáb"];
const MONTHS = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];

/** Fecha local -> 'YYYY-MM-DD' (nunca usar toISOString: desplaza por zona horaria). */
function ymd(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
/** 'YYYY-MM-DD' -> Date local a medianoche. */
function fromYmd(s) { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); }
function addDays(s, n) { const d = fromYmd(s); d.setDate(d.getDate() + n); return ymd(d); }
function dowOf(s) { return fromYmd(s).getDay(); }
function niceDate(s) { const d = fromYmd(s); return `${DOW[d.getDay()]} ${d.getDate()} de ${MONTHS[d.getMonth()]}`; }
function shortDate(s) { const d = fromYmd(s); return `${DOW_S[d.getDay()]} ${d.getDate()}/${d.getMonth() + 1}`; }
const toMin = (hhmm) => { const [h, m] = hhmm.split(":").map(Number); return h * 60 + m; };
const toHHmm = (mins) => `${String(Math.floor(mins / 60)).padStart(2, "0")}:${String(mins % 60).padStart(2, "0")}`;
/** 09:30 -> "9:30 AM" (más legible para el cliente en WhatsApp). */
function ampm(hhmm) {
  const [h, m] = hhmm.split(":").map(Number);
  const p = h < 12 ? "AM" : "PM"; const hh = h % 12 === 0 ? 12 : h % 12;
  return `${hh}:${String(m).padStart(2, "0")} ${p}`;
}
const money = (n, cur = "$") => cur + (Number(n) || 0).toLocaleString("es-CO", { maximumFractionDigits: 0 });
/** Normaliza texto para búsquedas: minúsculas, sin acentos. */
const norm = (s) => (s || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
/** Deja solo dígitos de un teléfono (formato requerido por wa.me). */
const digits = (s) => (s || "").replace(/\D/g, "");
const clone = (o) => JSON.parse(JSON.stringify(o));
const sum = (arr, f) => arr.reduce((a, b) => a + (Number(f(b)) || 0), 0);

/* ==========================================================================
   [PERMISOS] Roles del personal.
   admin     : todo
   recepcion : agenda, solicitudes, clientes, espera (sin dinero ni ajustes)
   barbero   : su propia agenda y sus clientes
   Usado por: App (menú visible) y cada vista (guardas de acción).
   ========================================================================== */
const PERMS = {
  admin:     ["agenda","caja","ventas","solicitudes","espera","clientes","barberos","servicios","productos","mantenimiento","insumos","gastos","reportes","ajustes"],
  recepcion: ["agenda","caja","ventas","solicitudes","espera","clientes","servicios","productos","mantenimiento","insumos"],
  barbero:   ["agenda","caja","solicitudes","espera","clientes"],
};
const can = (rol, perm) => (PERMS[rol] || []).includes(perm);

/* ==========================================================================
   [DB] Capa de datos — ÚNICO punto de contacto con Firestore.
   --------------------------------------------------------------------------
   Los datos están separados según quién puede tocarlos (ver firestore.rules):

   · shop/main        PRIVADO. Configuración, clientes, productos, gastos…
                      Solo personal con sesión. Toda escritura pasa por mutate().
   · pub/catalog      PÚBLICO (solo lectura). Copia mínima para la portada:
                      horarios, servicios, barberos, productos. SIN teléfonos
                      de clientes ni costos. mutate() la regenera sola.
   · appointments/    Citas. Un cliente SIN cuenta solo puede CREAR una
                      solicitud 'pendiente'; leer o cambiar es del personal.
   · locks/           "Candados" de horario (uno por cada 5 min ocupados).
                      Los lee cualquiera (para ocultar horas tomadas) y se
                      crean de forma atómica: dos clientes no pueden quedarse
                      con la misma hora.
   · waitlist/        Lista de espera (el cliente solo crea).
   · sales/           Ventas de la caja. Solo personal.
   · staff/<correo>   Acceso del personal (rol y si está activo).
   ========================================================================== */
const SHOP = doc(firestore, "shop", "main");
const PUB = doc(firestore, "pub", "catalog");
/** Marca del negocio (logo, paleta): pública, la escribe el personal. */
const BRAND = doc(firestore, "pub", "brand");
const fcol = (name) => collection(firestore, name);
const fref = (name, id) => doc(firestore, name, id);

/** Error con mensaje listo para mostrarle al usuario. */
const userErr = (msg) => { const e = new Error(msg); e.userMessage = msg; return e; };
const nowHHmm = () => { const n = new Date(); return toHHmm(n.getHours() * 60 + n.getMinutes()); };

/** Estado inicial de una instalación nueva (solo shop/main). */
function seed() {
  const hours = [1,2,3,4,5,6].map((dow) => ({ dow, start: "09:00", end: "19:00" }));
  return {
    rev: 1,
    salesSeq: 0,               // consecutivo de ventas (V-00001…)
    license: null,             // { code, client, validUntil, blocked, checkedAt }
    settings: {
      businessName: "Mi negocio",
      businessPhone: "",          // WhatsApp del negocio (con indicativo, solo dígitos)
      currency: "$",
      slotMinutes: 30,            // granularidad de la agenda
      bookingLeadMinutes: 0,       // anticipación mínima para que el cliente reserve (0 = sin restricción extra)
      openTime: "09:00",
      closeTime: "19:00",
      workDays: [1,2,3,4,5,6],    // 0=domingo
      soundEnabled: true,
      payMethods: ["Efectivo", "Nequi", "Daviplata", "Transferencia", "Tarjeta"],
      receiptFooter: "¡Gracias por tu visita!",
      tplClient: "Hola {cliente}, tu cita en {negocio} quedó CONFIRMADA.\n\nFecha: {fecha}\nHora: {hora}\nBarbero: {barbero}\nServicio: {servicio} ({duracion} min)\nValor: {precio}\n{comentario}\nTe esperamos. Si necesitas cambiarla, responde a este mensaje.",
      tplBusiness: "CITA CONFIRMADA\n\nCliente: {cliente}\nTeléfono: {telefono}\nFecha: {fecha}\nHora: {hora}\nBarbero: {barbero}\nServicio: {servicio}\nValor: {precio}\n{comentario}",
      tplRequestBusiness: "NUEVA SOLICITUD DE CITA\n\nCliente: {cliente}\nTeléfono: {telefono}\nFecha: {fecha}\nHora: {hora}\nBarbero: {barbero}\nServicio: {servicio}\nValor: {precio}\n{comentario}\nEntra al panel para aceptarla o rechazarla.",
    },
    barbers: [
      { id: "b_1", name: "Barbero 1", phone: "", hours: clone(hours), active: true },
    ],
    services: [
      { id: "s_1", name: "Corte clásico",     durationMin: 30, price: 25000, active: true },
      { id: "s_2", name: "Corte + barba",     durationMin: 45, price: 38000, active: true },
      { id: "s_3", name: "Perfilado de barba",durationMin: 20, price: 15000, active: true },
    ],
    clients: [],        // { id, name, phone, notes, createdAt, visits, spent, lastVisit }
    blocks: [],         // { id, barberId, from, to, reason, allDay, timeStart, timeEnd }
    expenses: [],       // { id, date, category, description, amount, method }
    products: [],       // { id, name, sku, stock, minStock, cost, price, unit, description, showPublic }
    equipment: [],      // { id, name, notes }
    maintenance: [],    // { id, equipmentId, date, description, createdAt }
    supplies: [],       // { id, name, dateToBuy, status: 'pendiente'|'comprado', createdAt }
    moves: [],          // { id, productId, type, qty, date, note }
  };
}

/** Garantiza que existan todos los campos (documentos viejos o recién creados). */
function normalize(d) {
  const base = seed();
  ["barbers","services","clients","blocks","expenses","products","moves","equipment","maintenance","supplies"]
    .forEach((k) => { if (!Array.isArray(d[k])) d[k] = base[k]; });
  d.settings = { ...base.settings, ...(d.settings || {}) };
  if (!Array.isArray(d.settings.payMethods) || !d.settings.payMethods.length) d.settings.payMethods = base.settings.payMethods;
  if (typeof d.salesSeq !== "number") d.salesSeq = 0;
  return d;
}

/**
 * Lo único que ve un visitante sin cuenta. Se arma SOLO con campos seguros:
 * nada de teléfonos de clientes, costos, gastos ni motivos de bloqueo.
 */
function publicView(d) {
  const s = d.settings;
  return {
    settings: {
      businessName: s.businessName, businessPhone: s.businessPhone || "", currency: s.currency,
      slotMinutes: s.slotMinutes, bookingLeadMinutes: s.bookingLeadMinutes || 0,
      openTime: s.openTime, closeTime: s.closeTime, workDays: s.workDays,
      tplRequestBusiness: s.tplRequestBusiness || "",
    },
    barbers: d.barbers.map((b) => ({ id: b.id, name: b.name, hours: b.hours || [], active: b.active !== false })),
    services: d.services.map((x) => ({ id: x.id, name: x.name, durationMin: x.durationMin, price: x.price, active: x.active !== false })),
    blocks: d.blocks.map((b) => ({
      id: b.id, barberId: b.barberId, from: b.from, to: b.to, allDay: !!b.allDay,
      timeStart: b.timeStart || "", timeEnd: b.timeEnd || "",
    })),
    products: d.products.filter((p) => p.showPublic !== false).map((p) => ({
      id: p.id, name: p.name, description: p.description || "", price: p.price || 0, available: (p.stock || 0) > 0,
    })),
    suspended: licenseSuspended(d.license),
    licensed: LICENSE_ON,                                                    // esta instalación exige licencia
    licenseUntilMs: d.license && d.license.d && !d.license.d.permanente ? d.license.d.hastaMs || 0 : 0, // para suspender sola al vencer
    updatedAt: new Date().toISOString(),
  };
}

/**
 * Cambio seguro sobre shop/main: corre dentro de una TRANSACCIÓN, así dos
 * dispositivos escribiendo a la vez no se pisan (Firestore reintenta solo).
 *   fn(d)            modifica el documento fresco.
 *   extra(tx,next)   (opcional) otras escrituras que deben ocurrir JUNTO con
 *                    el cambio (una venta, una cita, candados…). Primero las
 *                    lecturas, después las escrituras.
 * Cada vez regenera pub/catalog para que la portada nunca quede desfasada.
 */
async function mutate(fn, extra) {
  return runTransaction(firestore, async (tx) => {
    const snap = await tx.get(SHOP);
    const cur = snap.exists() ? normalize(snap.data()) : seed();
    const next = fn(cur) || cur;
    if (extra) await extra(tx, next, cur);
    next.rev = (cur.rev || 0) + 1;
    next.updatedAt = new Date().toISOString();
    tx.set(SHOP, next);
    tx.set(PUB, publicView(next));
    return next;
  });
}

/** Escribe muchas operaciones en tandas (Firestore admite 500 por lote). */
async function writeChunks(ops, size = 400) {
  for (let i = 0; i < ops.length; i += size) {
    const batch = writeBatch(firestore);
    ops.slice(i, i + size).forEach((op) => op(batch));
    await batch.commit();
  }
}

/* --- Candados de horario ------------------------------------------------- */
const LOCK_UNIT = 5; // minutos por candado
const padMin = (m) => String(m).padStart(4, "0");
const lockId = (barberId, date, m) => `${barberId}_${date}_${padMin(m)}`;

/** Candados que ocupa una cita: uno por cada 5 minutos. */
function lockDocs(a) {
  const s = toMin(a.time), e = s + (a.durationMin || 30);
  const out = [];
  for (let m = Math.floor(s / LOCK_UNIT) * LOCK_UNIT; m < e; m += LOCK_UNIT) {
    out.push({ id: lockId(a.barberId, a.date, m), data: { barberId: a.barberId, date: a.date, min: padMin(m), apptId: a.id || "" } });
  }
  return out;
}

/**
 * Dentro de una transacción: verifica que ningún candado del horario sea de
 * OTRA cita y devuelve los candados a escribir. Lanza error legible si hay choque.
 */
async function claimLocks(tx, appt) {
  const locks = lockDocs(appt);
  const snaps = await Promise.all(locks.map((l) => tx.get(fref("locks", l.id))));
  snaps.forEach((snap) => {
    if (snap.exists() && snap.data().apptId !== appt.id) throw userErr("Ese horario acaba de ocuparse. Elige otro.");
  });
  return locks.map((l) => ({ ref: fref("locks", l.id), data: l.data }));
}

/** Suma/resta visitas y gasto acumulado de un cliente (se usa al cobrar o anular). */
function bumpClient(d, clientId, { visits = 0, spent = 0, last = "" }) {
  const c = d.clients.find((x) => x.id === clientId);
  if (!c) return;
  c.visits = Math.max(0, (c.visits || 0) + visits);
  c.spent = Math.max(0, (c.spent || 0) + spent);
  if (last && last > (c.lastVisit || "")) c.lastVisit = last;
}

/* ==========================================================================
   [LOGICA] Reglas de negocio. Puras y testeables: no tocan React ni la DB.
   ========================================================================== */

/** Jornada del barbero en una fecha, o null si no trabaja ese día. */
function shiftFor(db, barber, date) {
  const dow = dowOf(date);
  const h = (barber.hours || []).find((x) => x.dow === dow);
  if (h) return h;
  return db.settings.workDays.includes(dow)
    ? { dow, start: db.settings.openTime, end: db.settings.closeTime }
    : null;
}

/** ¿El barbero está bloqueado (vacaciones/permiso) en ese rango horario? */
function isBlocked(db, barberId, date, startMin, endMin) {
  return (db.blocks || []).some((b) => {
    if (b.barberId !== barberId) return false;
    if (date < b.from || date > b.to) return false;
    if (b.allDay) return true;
    return startMin < toMin(b.timeEnd) && endMin > toMin(b.timeStart);
  });
}

/** Citas que ocupan agenda (pendientes y confirmadas). Canceladas liberan el cupo. */
const OCCUPY = ["pendiente", "confirmada", "completada"];

/**
 * ¿Cabe una cita de `durationMin` a las `time` para ese barbero?
 * Usado por: freeSlots(), aceptar solicitud, crear cita manual.
 */
function isFree(db, barberId, date, time, durationMin, ignoreApptId = null) {
  const s = toMin(time), e = s + durationMin;
  const barber = db.barbers.find((b) => b.id === barberId);
  if (!barber || !barber.active) return false;
  const shift = shiftFor(db, barber, date);
  if (!shift) return false;
  if (s < toMin(shift.start) || e > toMin(shift.end)) return false;
  if (isBlocked(db, barberId, date, s, e)) return false;
  // Vista del cliente: no ve las citas (son privadas), ve los candados.
  if (db.locks) {
    for (let m = Math.floor(s / LOCK_UNIT) * LOCK_UNIT; m < e; m += LOCK_UNIT) {
      if (db.locks.has(lockId(barberId, date, m))) return false;
    }
  }
  return !db.appointments.some((a) => {
    if (a.id === ignoreApptId) return false;
    if (a.barberId !== barberId || a.date !== date) return false;
    if (!OCCUPY.includes(a.status)) return false;
    const as = toMin(a.time), ae = as + (a.durationMin || 30);
    return s < ae && e > as;
  });
}

/**
 * Lista de horas libres de un barbero en una fecha para un servicio dado.
 * Si la fecha es hoy, además descarta las horas que ya pasaron: un horario
 * de las 9am no debería ofrecerse si ya son las 3pm. Días futuros no se ven
 * afectados por esto.
 *
 * `forClient=true` aplica además la anticipación mínima configurada en
 * Ajustes (settings.bookingLeadMinutes). Solo aplica a la reserva del
 * cliente — el mostrador y el reprogramado de una cita existente ven
 * cualquier horario libre, sin esa restricción.
 *
 * `ignoreApptId` excluye esa cita del cálculo de choques: así, al mover una
 * cita ya existente, su propio horario actual no se cuenta como "ocupado".
 * Usado por: BookingView, NewAppointmentModal, RescheduleModal.
 */
function freeSlots(db, barberId, date, durationMin, forClient = false, ignoreApptId = null) {
  const barber = db.barbers.find((b) => b.id === barberId);
  if (!barber) return [];
  const shift = shiftFor(db, barber, date);
  if (!shift) return [];
  const step = db.settings.slotMinutes || 30;
  const out = [];
  const now = new Date();
  const isToday = date === ymd(now);
  const nowMin = now.getHours() * 60 + now.getMinutes();
  const leadMin = forClient ? (db.settings.bookingLeadMinutes || 0) : 0;
  const cutoff = new Date(now.getTime() + leadMin * 60000);
  for (let m = toMin(shift.start); m + durationMin <= toMin(shift.end); m += step) {
    if (isToday && m <= nowMin) continue;
    if (leadMin > 0) {
      const slotStart = fromYmd(date);
      slotStart.setHours(Math.floor(m / 60), m % 60, 0, 0);
      if (slotStart < cutoff) continue;
    }
    const t = toHHmm(m);
    if (isFree(db, barberId, date, t, durationMin, ignoreApptId)) out.push(t);
  }
  return out;
}

/** Cita "hidratada" con los nombres reales, lista para pintar o exportar. */
function hydrate(db, a) {
  const c = db.clients.find((x) => x.id === a.clientId);
  const b = db.barbers.find((x) => x.id === a.barberId);
  const s = db.services.find((x) => x.id === a.serviceId);
  return {
    ...a,
    clientName: c ? c.name : (a.clientName || "Cliente eliminado"),
    clientPhone: c ? c.phone : (a.clientPhone || ""),
    barberName: b ? b.name : "—",
    serviceName: s ? s.name : "—",
    endTime: toHHmm(toMin(a.time) + (a.durationMin || 30)),
  };
}

function expensesBetween(db, from, to) {
  return db.expenses.filter((e) => e.date >= from && e.date <= to);
}

/** Reemplaza {marcadores} de las plantillas de WhatsApp. */
function fillTemplate(tpl, vars) {
  return Object.keys(vars).reduce((t, k) => t.split(`{${k}}`).join(vars[k] ?? ""), tpl || "");
}

/** Construye el enlace de WhatsApp. Sin teléfono devuelve null. */
function waLink(phone, text) {
  const p = digits(phone);
  if (!p) return null;
  return `https://wa.me/${p}?text=${encodeURIComponent(text)}`;
}

/** Variables de mensajería para una cita. Usado por: modal de aceptar cita. */
function msgVars(db, appt, comment) {
  const h = hydrate(db, appt);
  return {
    cliente: h.clientName,
    telefono: h.clientPhone,
    negocio: db.settings.businessName,
    fecha: niceDate(appt.date),
    hora: ampm(appt.time),
    barbero: h.barberName,
    servicio: h.serviceName,
    duracion: String(appt.durationMin),
    precio: money(appt.price, db.settings.currency),
    comentario: comment ? `Nota: ${comment}` : "",
  };
}

/* ==========================================================================
   [LICENCIA] Candado de licencia — se administra desde Conty Licencias.
   --------------------------------------------------------------------------
   Habla el mismo idioma que las demás apps Conty (conty-licencia.js): lee
   verificacion/<código> del proyecto Firebase "conty-licencia" y aplica sus
   mismas reglas (suspendida, cancelada, vencida, días de gracia, aviso previo).

   · Producto en Conty Licencias: "Conty Barber", código corto BRB
     (las licencias salen como BRB-XXXXX).
   · La licencia se guarda en shop/main (privado): el administrador la activa
     UNA vez y vale para todos los celulares del negocio.
   · Direcciones de demo y pruebas NO piden licencia: localhost, o cualquier
     dirección que contenga la palabra "demo".
   · Si no hay internet se respeta lo último que se supo (igual que Conty POS).
   ========================================================================== */
const LICENSE = {
  product: "BRB",
  appName: "Conty Barber",
  project: "conty-licencia",
  apiKey: "AIzaSyAwU-NP-_E-lSxaccdCT1ZQ6Sm8WvByiEI",   // llave pública de Conty Licencias (la misma de conty-licencia.js)
  freeHosts: ["localhost", "127.0.0.1"],
  freeHostContains: ["demo"],
  warnDays: 3,
  recheckHours: 6,
};
const LICENSE_ON = (() => {
  if (typeof location === "undefined") return false;
  const h = String(location.hostname || "").toLowerCase();
  return !(LICENSE.freeHosts.includes(h) || LICENSE.freeHostContains.some((w) => h.includes(w)));
})();
const DAY_MS = 86400000;
const endOfDay = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d + 1).getTime(); };
const licDate = (ms) => { const d = new Date(ms); return `${d.getDate()} ${MONTHS[d.getMonth()].slice(0, 3)} ${d.getFullYear()}`; };

/** Decodifica un campo de la API REST de Firestore. */
function fsVal(v) {
  if (!v) return undefined;
  if ("stringValue" in v) return v.stringValue;
  if ("integerValue" in v) return Number(v.integerValue);
  if ("doubleValue" in v) return v.doubleValue;
  if ("booleanValue" in v) return v.booleanValue;
  return undefined;
}

/** Lee la licencia en Conty Licencias. { noExiste:true } si el código no existe. Lanza error si no hay red. */
async function readLicenseRemote(code) {
  const url = `https://firestore.googleapis.com/v1/projects/${LICENSE.project}/databases/(default)/documents/verificacion/${encodeURIComponent(code)}?key=${LICENSE.apiKey}`;
  const r = await fetch(url, { cache: "no-store" });
  if (r.status === 404 || r.status === 403) return { noExiste: true };
  if (!r.ok) throw new Error("http " + r.status);
  const f = (await r.json()).fields || {};
  return {
    estado: fsVal(f.estado) || "activa", productoCodigo: String(fsVal(f.productoCodigo) || "").toUpperCase(),
    producto: fsVal(f.producto) || "", negocio: fsVal(f.negocio) || "", vence: fsVal(f.vence) || "",
    hastaMs: Number(fsVal(f.hastaMs)) || 0, permanente: !!fsVal(f.permanente),
  };
}

/** WhatsApp de cobros de Conty (se configura en Conty Licencias → Más → Ajustes). */
async function readVendorWhatsapp() {
  try {
    const cached = localStorage.getItem("conty_lic_whatsapp");
    if (cached) return cached;
  } catch (e) { /* sin almacenamiento */ }
  try {
    const r = await fetch(`https://firestore.googleapis.com/v1/projects/${LICENSE.project}/databases/(default)/documents/publico/general?key=${LICENSE.apiKey}`, { cache: "no-store" });
    if (!r.ok) return "";
    const j = await r.json();
    const w = String(fsVal((j.fields || {}).whatsapp) || "").replace(/\D/g, "");
    if (w) { try { localStorage.setItem("conty_lic_whatsapp", w); } catch (e) { /* ignorar */ } }
    return w;
  } catch (e) { return ""; }
}

/** Mismas reglas que conty-licencia.js, sobre el documento ya leído. */
function evaluateLicense(d) {
  if (!d) return { tipo: "invalido" };
  if (d.productoCodigo !== LICENSE.product) return { tipo: "otraApp", d };
  if (d.estado === "suspendida") return { tipo: "bloqueo", titulo: "Licencia suspendida", texto: "Tu licencia está suspendida. Escríbenos para reactivarla.", d };
  if (d.estado === "cancelada") return { tipo: "bloqueo", titulo: "Licencia cancelada", texto: "Esta licencia fue cancelada. Escríbenos si quieres volver a activarla.", d };
  if (d.permanente) return { tipo: "ok", d };
  const now = Date.now();
  if (!d.hastaMs) return { tipo: "bloqueo", titulo: "Falta el pago", texto: "Tu licencia todavía no tiene el pago registrado.", d };
  if (now >= d.hastaMs) return { tipo: "bloqueo", titulo: "Licencia vencida", texto: "Tu licencia venció" + (d.vence ? " el " + licDate(endOfDay(d.vence) - DAY_MS) : "") + ". Renueva para seguir usando la app.", d };
  const fin = d.vence ? endOfDay(d.vence) : d.hastaMs;
  if (now >= fin) return { tipo: "gracia", texto: "Tu licencia venció. Puedes usar la app hasta el " + licDate(d.hastaMs - 1) + ".", d };
  if (fin - now <= LICENSE.warnDays * DAY_MS) return { tipo: "pronto", texto: "Tu licencia vence el " + licDate(fin - 1) + ".", d };
  return { tipo: "ok", d };
}

/** Estado de la licencia guardada en shop/main. tipo: falta | invalido | otraApp | bloqueo | gracia | pronto | ok */
function licenseStatus(l) {
  if (!l || !l.code) return { tipo: "falta" };
  if (l.invalid || !l.d) return { tipo: "invalido" };
  return evaluateLicense(l.d);
}
const LIC_BLOCKING = ["falta", "invalido", "otraApp", "bloqueo"];
/** ¿Hay que bloquear el panel por licencia? Solo en instalaciones de clientes. */
const licenseBlocked = (shop) => LICENSE_ON && LIC_BLOCKING.includes(licenseStatus(shop.license).tipo);
/** ¿Hay que suspender la reserva pública? (una instalación aún sin activar NO se suspende) */
const licenseSuspended = (l) => LICENSE_ON && ["invalido", "otraApp", "bloqueo"].includes(licenseStatus(l).tipo);

/* ==========================================================================
   [MARCA] Logo, nombre y paleta de colores — editables desde Ajustes → Apariencia.
   --------------------------------------------------------------------------
   Cada negocio guarda su marca en pub/brand (lectura pública, escritura del
   personal), así la ven también los clientes que reservan. El código es el
   mismo para todos: lo único que cambia entre negocios es esta marca y su
   propia base de datos de Firebase.
   Todas las paletas se revisaron para que botones y textos se lean bien.
   ========================================================================== */
const FONTS = {
  industrial: { name: "Industrial", display: "'Oswald','Arial Narrow',ui-sans-serif,sans-serif", body: "'Barlow',ui-sans-serif,system-ui,sans-serif" },
  elegante:   { name: "Elegante",   display: "'Playfair Display',Georgia,serif",                  body: "'Poppins',ui-sans-serif,system-ui,sans-serif" },
};
const PALETTES = [
  { id: "clasica",   name: "Negro y dorado",     ink: "#14161B", ink2: "#1E2228", line: "#DCD7CD", porcelain: "#F4F2ED", paper: "#FFFFFF", brass: "#B08725", brassSoft: "#F0E4C4", pole: "#B5322E", mint: "#2F7D68", slate: "#5A6169", font: "industrial" },
  { id: "rosa",      name: "Rosa empolvado",     ink: "#3D2B33", ink2: "#4E3743", line: "#EAD9D6", porcelain: "#FBF2F0", paper: "#FFFFFF", brass: "#B07257", brassSoft: "#F5E4DC", pole: "#B83A5A", mint: "#5E8A71", slate: "#7A6268", font: "elegante" },
  { id: "marino",    name: "Azul marino",        ink: "#14213D", ink2: "#1F2F52", line: "#D5DBE6", porcelain: "#F1F4F9", paper: "#FFFFFF", brass: "#4C6A92", brassSoft: "#DCE6F4", pole: "#B5322E", mint: "#2F7D68", slate: "#55607A", font: "industrial" },
  { id: "bosque",    name: "Verde bosque y cobre", ink: "#1B2E24", ink2: "#263F32", line: "#D9DCCF", porcelain: "#F3F4EC", paper: "#FFFFFF", brass: "#A55C30", brassSoft: "#F1DFD0", pole: "#B5322E", mint: "#3A7F59", slate: "#5A6560", font: "industrial" },
  { id: "vino",      name: "Vino y crema",       ink: "#3B1620", ink2: "#4F2230", line: "#E6D8D0", porcelain: "#F8F1EA", paper: "#FFFFFF", brass: "#9A6A26", brassSoft: "#F0E1C8", pole: "#A82B3B", mint: "#3A7A5E", slate: "#6B5A58", font: "elegante" },
  { id: "turquesa",  name: "Grafito y turquesa", ink: "#1F2428", ink2: "#2B3238", line: "#D6DEDF", porcelain: "#F0F5F5", paper: "#FFFFFF", brass: "#1B7F7F", brassSoft: "#D3ECEB", pole: "#B5322E", mint: "#2F7D68", slate: "#566066", font: "industrial" },
  { id: "lila",      name: "Lila suave",         ink: "#2E2540", ink2: "#3E3256", line: "#E3DCEB", porcelain: "#F6F2FA", paper: "#FFFFFF", brass: "#7E5AA3", brassSoft: "#E9DDF4", pole: "#B83A5A", mint: "#5E8A71", slate: "#6F6480", font: "elegante" },
  { id: "terracota", name: "Terracota y arena",  ink: "#3A2A22", ink2: "#4A372D", line: "#E6DACB", porcelain: "#F8F2E8", paper: "#FFFFFF", brass: "#B35634", brassSoft: "#F4DDCF", pole: "#A8322E", mint: "#4F8250", slate: "#7A6A5F", font: "industrial" },
];

/* --- utilidades de color (para la paleta «Personalizada») --- */
const hex2rgb = (h) => { const n = parseInt(String(h).replace("#", ""), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
const rgb2hex = (r, g, b) => "#" + [r, g, b].map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, "0")).join("");
function rgb2hsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
  if (mx === mn) return [0, 0, l];
  const d = mx - mn, s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  const h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h * 60, s, l];
}
function hsl(h, s, l) {
  const k = (n) => (n + h / 30) % 12, a = s * Math.min(l, 1 - l);
  const f = (n) => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
  return rgb2hex(f(0) * 255, f(8) * 255, f(4) * 255);
}
function luminance(hex) {
  const [r, g, b] = hex2rgb(hex).map((c) => { c /= 255; return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
function contrastRatio(a, b) { const la = luminance(a), lb = luminance(b); return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05); }
const isHex = (v) => /^#[0-9a-fA-F]{6}$/.test(String(v || ""));

/** Arma una paleta completa a partir de solo 2 colores. Oscurece lo necesario para que se lea. */
function deriveTokens(primary, accent, font) {
  const p = isHex(primary) ? primary : "#1B3A5C", a0 = isHex(accent) ? accent : "#B08725";
  const [h, s] = rgb2hsl(...hex2rgb(p));
  const [ah, as, al0] = rgb2hsl(...hex2rgb(a0));
  let al = al0, acc = a0;
  while (contrastRatio(acc, "#FFFFFF") < 3 && al > 0.12) { al -= 0.03; acc = hsl(ah, as, al); } // el acento lleva texto blanco encima
  return {
    ink: hsl(h, Math.min(s, 0.5), 0.14), ink2: hsl(h, Math.min(s, 0.5), 0.2),
    line: hsl(h, Math.min(s, 0.25), 0.87), porcelain: hsl(h, Math.min(s, 0.3), 0.96), paper: "#FFFFFF",
    brass: acc, brassSoft: hsl(ah, Math.min(as, 0.6), 0.88),
    pole: "#B5322E", mint: "#2F7D68", slate: hsl(h, Math.min(s, 0.12), 0.36),
    font: font === "elegante" ? "elegante" : "industrial",
  };
}

/** Colores efectivos de una marca guardada. */
function resolveTokens(brand) {
  const b = brand || {};
  const base = b.palette === "custom"
    ? deriveTokens((b.custom || {}).primary, (b.custom || {}).accent, (b.custom || {}).font)
    : { ...(PALETTES.find((x) => x.id === b.palette) || PALETTES[0]) };
  if (b.font && FONTS[b.font]) base.font = b.font;
  return base;
}
/** Variables CSS de una paleta, en forma de bloque de estilos. */
function themeCss(t) {
  const f = FONTS[t.font] || FONTS.industrial;
  return `.bt{ --ink:${t.ink}; --ink-2:${t.ink2}; --line:${t.line}; --porcelain:${t.porcelain}; --paper:${t.paper}; --brass:${t.brass}; --brass-soft:${t.brassSoft}; --pole:${t.pole}; --mint:${t.mint}; --slate:${t.slate}; --display:${f.display}; --body:${f.body}; }`;
}
/** Lo mismo, como objeto de estilo de React (para la vista previa en Ajustes). */
function tokensStyle(t) {
  const f = FONTS[t.font] || FONTS.industrial;
  return { "--ink": t.ink, "--ink-2": t.ink2, "--line": t.line, "--porcelain": t.porcelain, "--paper": t.paper, "--brass": t.brass, "--brass-soft": t.brassSoft, "--pole": t.pole, "--mint": t.mint, "--slate": t.slate, "--display": f.display, "--body": f.body };
}

/** Marca activa. App la actualiza en cada render, antes de pintar las pantallas. */
const BRANDING = { logo: "", plate: "none", name: "", theme: PALETTES[0] };
/** Logo por defecto: monograma con las iniciales del negocio. */
function defaultLogo(name, theme) {
  const t = theme || BRANDING.theme;
  const words = String(name || "").replace(/[^\p{L}\p{N}\s]/gu, "").trim().split(/\s+/).filter(Boolean);
  const ini = (words.length > 1 ? words[0][0] + words[1][0] : (words[0] || "B").slice(0, 2)).toUpperCase();
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120"><circle cx="60" cy="60" r="56" fill="${t.ink}" stroke="${t.brass}" stroke-width="4"/><text x="60" y="77" text-anchor="middle" font-family="Georgia,serif" font-size="46" font-weight="700" fill="${t.brass}">${ini}</text></svg>`;
  return "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svg)));
}
const currentLogo = () => BRANDING.logo || defaultLogo(BRANDING.name, BRANDING.theme);
/** Aplica la marca guardada (y el nombre del negocio) a todo lo que se pinte después. */
function applyBranding(brand, name) {
  const b = brand || {};
  BRANDING.logo = typeof b.logo === "string" ? b.logo : "";
  BRANDING.plate = b.plate === "light" ? "light" : "none";
  BRANDING.name = name || "";
  BRANDING.theme = resolveTokens(b);
  CSS = BASE_CSS + themeCss(BRANDING.theme);
}

/* --- logo: de la galería a un PNG liviano --- */
function loadImage(src) {
  return new Promise((ok, bad) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => bad(new Error("img")); i.src = src; });
}
/** Reduce la imagen elegida y decide si necesita un fondo claro sobre la barra oscura. */
async function fileToLogo(file) {
  if (!file || !/^image\//.test(file.type)) throw userErr("Elige una imagen (JPG o PNG).");
  const url = URL.createObjectURL(file);
  let img;
  try { img = await loadImage(url); } catch (e) { throw userErr("No se pudo abrir esa imagen. Prueba con otra."); } finally { URL.revokeObjectURL(url); }
  let max = 420, out = "", w = 0, h = 0, cv, ctx;
  for (let k = 0; k < 4; k++) {
    const r = Math.min(1, max / Math.max(img.width, img.height));
    w = Math.max(1, Math.round(img.width * r)); h = Math.max(1, Math.round(img.height * r));
    cv = document.createElement("canvas"); cv.width = w; cv.height = h;
    ctx = cv.getContext("2d");
    if (!ctx) throw userErr("Tu navegador no puede procesar la imagen. Prueba con otro.");
    ctx.drawImage(img, 0, 0, w, h);
    out = cv.toDataURL("image/png");
    if (out.length < 300000) break;       // cabe cómodo en la base de datos
    max = Math.round(max * 0.72);
  }
  if (out.length >= 300000) throw userErr("La imagen es muy pesada. Prueba con una más simple.");
  // ¿Es un logo oscuro sobre fondo transparente? Entonces en la barra oscura se pierde: se le pone un fondo claro.
  const px = ctx.getImageData(0, 0, w, h).data;
  let sum = 0, n = 0, clear = 0;
  for (let i = 0; i < px.length; i += 4) {
    if (px[i + 3] < 40) { clear++; continue; }
    sum += (0.299 * px[i] + 0.587 * px[i + 1] + 0.114 * px[i + 2]) / 255; n++;
  }
  const avg = n ? sum / n : 1;
  return { logo: out, plate: avg < 0.45 && clear / (px.length / 4) > 0.12 ? "light" : "none" };
}

/** Icono cuadrado (para el ícono de la app instalada). */
async function squareIcon(src, size, bg) {
  const img = await loadImage(src);
  const cv = document.createElement("canvas"); cv.width = size; cv.height = size;
  const ctx = cv.getContext("2d");
  if (!ctx) throw new Error("canvas");
  ctx.fillStyle = bg; ctx.fillRect(0, 0, size, size);
  const box = size * 0.72, r = Math.min(box / img.width, box / img.height);
  ctx.drawImage(img, (size - img.width * r) / 2, (size - img.height * r) / 2, img.width * r, img.height * r);
  return cv.toDataURL("image/png");
}
let manifestUrl = "";
function setHeadLink(rel, href) {
  let el = document.querySelector(`link[rel="${rel}"]`);
  if (!el) { el = document.createElement("link"); el.setAttribute("rel", rel); document.head.appendChild(el); }
  el.setAttribute("href", href);
}
/** Título de la pestaña, color de la barra, ícono y manifiesto de la app instalada: todo con la marca del negocio. */
async function applyDocumentBranding() {
  const t = BRANDING.theme, logo = currentLogo(), name = BRANDING.name || "Reserva tu turno";
  document.title = name;
  let tc = document.querySelector('meta[name="theme-color"]');
  if (!tc) { tc = document.createElement("meta"); tc.setAttribute("name", "theme-color"); document.head.appendChild(tc); }
  tc.setAttribute("content", t.ink);
  setHeadLink("icon", logo);
  setHeadLink("apple-touch-icon", logo);
  try {
    const [i192, i512] = await Promise.all([squareIcon(logo, 192, t.porcelain), squareIcon(logo, 512, t.porcelain)]);
    const manifest = {
      name, short_name: name.slice(0, 12), start_url: location.origin + "/", scope: location.origin + "/",
      display: "standalone", orientation: "portrait", background_color: t.porcelain, theme_color: t.ink,
      icons: [{ src: i192, sizes: "192x192", type: "image/png" }, { src: i512, sizes: "512x512", type: "image/png", purpose: "any maskable" }],
    };
    const url = URL.createObjectURL(new Blob([JSON.stringify(manifest)], { type: "application/manifest+json" }));
    setHeadLink("manifest", url);
    if (manifestUrl) URL.revokeObjectURL(manifestUrl);
    manifestUrl = url;
  } catch (e) { /* el ícono instalado queda con el genérico: no es grave */ }
}

/* ==========================================================================
   [SONIDO] Aviso audible cuando entra una solicitud nueva.
   El navegador exige un gesto previo del usuario: se desbloquea con el
   primer clic (ver App -> useEffect de desbloqueo).
   ========================================================================== */
const Sound = {
  ctx: null,
  unlock() {
    try {
      if (!this.ctx) this.ctx = new (window.AudioContext || window.webkitAudioContext)();
      if (this.ctx.state === "suspended") this.ctx.resume();
    } catch (e) { /* sin audio disponible */ }
  },
  /** Campanilla de dos tonos, discreta. */
  chime() {
    try {
      this.unlock();
      if (!this.ctx) return;
      [[880, 0], [1318.5, 0.16]].forEach(([f, delay]) => {
        const o = this.ctx.createOscillator(), g = this.ctx.createGain();
        o.type = "sine"; o.frequency.value = f;
        const t = this.ctx.currentTime + delay;
        g.gain.setValueAtTime(0, t);
        g.gain.linearRampToValueAtTime(0.22, t + 0.02);
        g.gain.exponentialRampToValueAtTime(0.001, t + 0.42);
        o.connect(g); g.connect(this.ctx.destination);
        o.start(t); o.stop(t + 0.45);
      });
    } catch (e) { /* ignorar */ }
  },
};

/* ==========================================================================
   [EXPORTAR] Excel (.xlsx real) y PDF (vía diálogo de impresión del sistema).
   ========================================================================== */
/** Escapa texto para meterlo en HTML. Los nombres y notas los escriben los
 *  clientes, así que NUNCA se inserta texto suyo en HTML sin pasar por aquí. */
const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

/** Descarga una o varias hojas en un mismo archivo .xlsx. */
function toExcel(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ "Sin datos": "" }]);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

/** Abre una ventana con el reporte formateado y lanza Imprimir → "Guardar como PDF". */
function toPDF(title, subtitle, sections) {
  const w = window.open("", "_blank");
  if (!w) { alert("El navegador bloqueó la ventana. Permite ventanas emergentes para exportar a PDF."); return; }
  const table = (rows) => {
    if (!rows.length) return "<p class='empty'>Sin registros en este periodo.</p>";
    const cols = Object.keys(rows[0]);
    return `<table><thead><tr>${cols.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>${
      rows.map((r) => `<tr>${cols.map((c) => `<td>${esc(r[c])}</td>`).join("")}</tr>`).join("")
    }</tbody></table>`;
  };
  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title>
  <style>
    @page{ margin:16mm; }
    body{ font-family:Barlow,Helvetica,Arial,sans-serif; color:${BRANDING.theme.ink}; }
    header{ display:flex; align-items:center; gap:14px; border-bottom:2px solid ${BRANDING.theme.ink}; padding-bottom:10px; margin-bottom:18px; }
    header img{ height:56px; }
    h1{ font-family:Oswald,Arial Narrow,sans-serif; text-transform:uppercase; letter-spacing:.06em; font-size:20px; margin:0; }
    .sub{ color:${BRANDING.theme.slate}; font-size:12px; margin-top:3px; }
    h2{ font-family:Oswald,Arial Narrow,sans-serif; text-transform:uppercase; letter-spacing:.1em; font-size:13px;
        margin:22px 0 8px; padding-bottom:4px; border-bottom:1px solid ${BRANDING.theme.line}; }
    table{ width:100%; border-collapse:collapse; font-size:11px; }
    th{ text-align:left; background:${BRANDING.theme.porcelain}; border-bottom:1px solid ${BRANDING.theme.line}; padding:6px; font-weight:600; }
    td{ border-bottom:1px solid ${BRANDING.theme.line}; padding:6px; }
    .empty{ color:${BRANDING.theme.slate}; font-size:12px; font-style:italic; }
    footer{ margin-top:24px; border-top:1px solid ${BRANDING.theme.line}; padding-top:8px; color:${BRANDING.theme.slate}; font-size:10px; }
  </style></head><body>
  <header><img src="${currentLogo()}"/><div><h1>${esc(title)}</h1><div class="sub">${esc(subtitle)}</div></div></header>
  ${sections.map((s) => `<h2>${esc(s.title)}</h2>${s.rows ? table(s.rows) : s.html || ""}`).join("")}
  <footer>Generado el ${new Date().toLocaleString("es-CO")}</footer>
  </body></html>`);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 450);
}

/** Respaldo completo: descarga un .json restaurable (incluye citas, espera y ventas). */
async function downloadFullBackup(db) {
  const [a, w, v] = await Promise.all([getDocs(fcol("appointments")), getDocs(fcol("waitlist")), getDocs(fcol("sales"))]);
  const { appointments, waitlist, staff, legacy, ...shopPart } = db;
  const payload = {
    ...shopPart,
    appointments: a.docs.map((d) => ({ ...d.data(), id: d.id })),
    waitlist: w.docs.map((d) => ({ ...d.data(), id: d.id })),
    sales: v.docs.map((d) => ({ ...d.data(), id: d.id })),
    exportedAt: new Date().toISOString(),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `respaldo-${ymd()}.json`;
  link.click(); URL.revokeObjectURL(link.href);
  return payload;
}

/* ==========================================================================
   [UI] Piezas reutilizables.
   ========================================================================== */
function Logo({ dark = false, className = "" }) {
  const img = <img src={currentLogo()} alt={BRANDING.name || "Logo"} className={`bt-logo ${className}`} />;
  // Logo oscuro sobre fondo transparente: en la barra oscura se pone sobre una placa clara para que se vea.
  return dark && BRANDING.plate === "light" ? <span className="bt-logoplate">{img}</span> : img;
}

function Field({ label, children, hint }) {
  return (
    <div>
      <label className="bt-label">{label}</label>
      {children}
      {hint && <div style={{ fontSize: 11, color: "var(--slate)", marginTop: 4 }}>{hint}</div>}
    </div>
  );
}

function Modal({ title, onClose, children, wide = false }) {
  useEffect(() => {
    const esc = (e) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", esc);
    return () => window.removeEventListener("keydown", esc);
  }, [onClose]);
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4"
         style={{ background: "rgba(20,22,27,.55)" }} onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="bt-card w-full sm:max-w-lg bt-scroll" style={{ maxWidth: wide ? 780 : undefined, maxHeight: "92vh" }}>
        <div className="flex items-center justify-between px-4 py-3" style={{ borderBottom: "1px solid var(--line)" }}>
          <h3 className="bt-display" style={{ fontSize: 15 }}>{title}</h3>
          <button className="bt-icon" onClick={onClose} aria-label="Cerrar"><X size={15} /></button>
        </div>
        <div className="p-4">{children}</div>
      </div>
    </div>
  );
}

function Empty({ icon: Icon, title, action }) {
  return (
    <div className="bt-inset flex flex-col items-center justify-center text-center py-12 px-6">
      <Icon size={26} style={{ color: "var(--brass)" }} />
      <p className="bt-display mt-3" style={{ fontSize: 14 }}>{title}</p>
      {action && <div className="mt-4">{action}</div>}
    </div>
  );
}

function Stat({ label, value, sub, tone }) {
  const color = tone === "pole" ? "var(--pole)" : tone === "mint" ? "var(--mint)" : "var(--ink)";
  return (
    <div className="bt-card p-4">
      <div className="bt-eyebrow">{label}</div>
      <div className="bt-mono" style={{ fontSize: 24, fontWeight: 600, color, marginTop: 6, lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: "var(--slate)", marginTop: 5 }}>{sub}</div>}
    </div>
  );
}

const TAG = { pendiente: "pend", confirmada: "conf", completada: "done", cancelada: "canc", rechazada: "canc" };
function StatusTag({ s }) { return <span className={`bt-tag ${TAG[s] || "done"}`}>{s}</span>; }

/** Buscador de clientes por nombre o teléfono, con creación en línea. */
function ClientPicker({ db, value, onChange, allowCreate = true }) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const isNew = value && typeof value === "object";
  const sel = isNew ? { name: value.newName, phone: "por registrar" } : db.clients.find((c) => c.id === value);
  const results = useMemo(() => {
    const n = norm(q);
    if (!n) return db.clients.slice(-6).reverse();
    return db.clients.filter((c) => norm(c.name).includes(n) || digits(c.phone).includes(digits(q))).slice(0, 8);
  }, [q, db.clients]);

  if (sel && !open) {
    return (
      <div className="bt-inset flex items-center justify-between px-3 py-2">
        <div>
          <div style={{ fontWeight: 600, fontSize: 14 }}>{sel.name}</div>
          <div className="bt-mono" style={{ fontSize: 11, color: "var(--slate)" }}>{sel.phone || "sin teléfono"}</div>
        </div>
        <button className="bt-btn ghost sm" onClick={() => { setOpen(true); setQ(""); onChange(""); }}>Cambiar</button>
      </div>
    );
  }
  return (
    <div>
      <div className="relative">
        <Search size={14} style={{ position: "absolute", left: 10, top: 12, color: "var(--slate)" }} />
        <input className="bt-input" style={{ paddingLeft: 30 }} autoFocus placeholder="Buscar por nombre o teléfono"
               value={q} onChange={(e) => setQ(e.target.value)} />
      </div>
      <div className="mt-2" style={{ maxHeight: 190, overflow: "auto" }}>
        {results.map((c) => (
          <button key={c.id} className="w-full text-left px-3 py-2 bt-hair"
                  style={{ borderBottom: "1px solid var(--line)", background: "transparent", border: 0, cursor: "pointer" }}
                  onClick={() => { onChange(c.id); setOpen(false); }}>
            <div style={{ fontSize: 14 }}>{c.name}</div>
            <div className="bt-mono" style={{ fontSize: 11, color: "var(--slate)" }}>{c.phone}</div>
          </button>
        ))}
        {!results.length && <p style={{ fontSize: 12, color: "var(--slate)", padding: "8px 2px" }}>Ningún cliente coincide.</p>}
      </div>
      {allowCreate && q.trim().length > 1 && (
        <button className="bt-btn ghost sm mt-2 w-full" onClick={() => onChange({ newName: q.trim() })}>
          <UserPlus size={13} /> Crear cliente “{q.trim()}”
        </button>
      )}
    </div>
  );
}

/* ==========================================================================
   [CLIENTE] Lo que puede hacer alguien SIN cuenta: pedir una cita y anotarse
   en la lista de espera. No lee ni cambia nada más (lo exigen las reglas).
   ========================================================================== */
/**
 * Crea una solicitud 'pendiente' y sus candados en UNA transacción: si otro
 * cliente tomó esa hora un segundo antes, falla con code 'slot-taken' y no
 * se crea nada (nunca hay dos clientes con la misma hora).
 */
async function requestAppointment(pubDb, p) {
  const svc = pubDb.services.find((s) => s.id === p.serviceId);
  if (!svc) throw new Error("servicio");
  const id = uid();
  const appt = {
    clientName: p.name.trim().slice(0, 80), clientPhone: p.phone.trim().slice(0, 20), clientId: "",
    barberId: p.barberId, serviceId: p.serviceId, date: p.date, time: p.time,
    durationMin: svc.durationMin, price: svc.price,
    status: "pendiente", source: "cliente", clientNote: (p.note || "").slice(0, 300), comment: "",
    createdAt: new Date().toISOString(),
  };
  const locks = lockDocs({ ...appt, id });
  await runTransaction(firestore, async (tx) => {
    const snaps = await Promise.all(locks.map((l) => tx.get(fref("locks", l.id))));
    if (snaps.some((x) => x.exists())) { const e = new Error("taken"); e.code = "slot-taken"; throw e; }
    tx.set(fref("appointments", id), appt);
    locks.forEach((l) => tx.set(fref("locks", l.id), l.data));
  });
  return id;
}

/* ==========================================================================
   [APP] Estado global, sincronización y acciones.
   ========================================================================== */
const NAV = [
  { key: "agenda",      label: "Agenda",      icon: Calendar },
  { key: "caja",        label: "Caja",        icon: CreditCard },
  { key: "ventas",      label: "Ventas",      icon: Receipt },
  { key: "solicitudes", label: "Solicitudes", icon: Bell },
  { key: "espera",      label: "Lista espera",icon: ListPlus },
  { key: "clientes",    label: "Clientes",    icon: Users },
  { key: "barberos",    label: "Barberos",    icon: Scissors },
  { key: "servicios",   label: "Servicios",   icon: Clock },
  { key: "productos",   label: "Productos",   icon: Package },
  { key: "mantenimiento", label: "Mantenimiento", icon: Wrench },
  { key: "insumos",     label: "Insumos",     icon: ShoppingCart },
  { key: "gastos",      label: "Gastos",      icon: Wallet },
  { key: "reportes",    label: "Reportes",    icon: BarChart3 },
  { key: "ajustes",     label: "Ajustes",     icon: Cog },
];

function LoadingScreen({ err }) {
  return (
    <div className="bt min-h-screen flex flex-col items-center justify-center gap-4 px-5">
      <style>{CSS}</style>
      <Logo />
      <p className="bt-eyebrow">{err ? "Sin conexión" : "Cargando"}</p>
      {err && <p style={{ fontSize: 12, color: "var(--slate)", textAlign: "center", maxWidth: 320 }}>{err}</p>}
      {err && <button className="bt-btn ghost" onClick={() => window.location.reload()}><RefreshCw size={13} /> Reintentar</button>}
    </div>
  );
}

export default function App() {
  const [pub, setPub] = useState(undefined);       // catálogo público: undefined=cargando, null=aún no existe
  // Marca (logo, paleta). Se recuerda en el celular para que la próxima vez se vea de una, sin parpadeo.
  const [brand, setBrand] = useState(() => { try { const c = JSON.parse(localStorage.getItem("bt_brand") || "null"); return c || undefined; } catch (e) { return undefined; } });
  const [fbUser, setFbUser] = useState(undefined); // cuenta de Firebase Auth: undefined=resolviendo
  const [user, setUser] = useState(null);          // registro de personal ya validado
  const [authErr, setAuthErr] = useState("");
  const [shop, setShop] = useState(null);          // shop/main (solo personal)
  const [appts, setAppts] = useState([]);          // citas desde winFrom en adelante
  const [pendAppts, setPendAppts] = useState([]);  // todas las pendientes (sin importar la fecha)
  const [waitlist, setWaitlist] = useState([]);
  const [staffList, setStaffList] = useState([]);
  const [winFrom, setWinFrom] = useState(() => addDays(ymd(), -3));
  const [err, setErr] = useState(null);
  const [mode, setMode] = useState("gate");        // gate | catalogo | login | reservar | panel
  const [view, setView] = useState("agenda");
  const [preset, setPreset] = useState(null);      // venta precargada desde una cita
  const [menuOpen, setMenuOpen] = useState(false);
  const [toasts, setToasts] = useState([]);
  const pendingRef = useRef(null);                 // ids de solicitudes ya vistas
  const shopRef = useRef(null);
  const dbRef = useRef(null);

  const toast = useCallback((text, tone = "ink") => {
    const id = uid();
    setToasts((t) => [...t, { id, text, tone }]);
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 3800);
  }, []);

  /** Crea el manejador de error de una escucha y muestra el código real de Firebase. */
  const onFail = useCallback((what) => (e) => {
    console.error(e);
    setErr(`Sin conexión con ${what}. Código: ${e.code || e.message || "desconocido"}`);
  }, []);

  /* 1) Catálogo público: cualquiera puede leerlo (portada, horarios, productos). */
  useEffect(() => onSnapshot(PUB, (s) => { setPub(s.exists() ? s.data() : null); setErr(null); }, onFail("el catálogo")), [onFail]);

  /* 1b) Marca del negocio (también pública: los clientes que reservan ven el logo y los colores). */
  useEffect(() => onSnapshot(BRAND, (s) => {
    const b = s.exists() ? s.data() : {};
    setBrand(b);
    try { localStorage.setItem("bt_brand", JSON.stringify(b)); } catch (e) { /* sin almacenamiento */ }
  }, (e) => { console.error(e); setBrand((b) => b || {}); }), []);

  /* 2) Sesión de Firebase Auth (se restaura sola al volver a abrir la app). */
  useEffect(() => onAuthStateChanged(auth, (u) => setFbUser(u || null)), []);

  /* 3) Vincula la cuenta con su registro en staff/<correo>. */
  useEffect(() => {
    if (fbUser === undefined) return undefined;
    if (!fbUser) {
      setUser(null);
      setMode((m) => (m === "panel" ? "gate" : m));
      return undefined;
    }
    let dead = false;
    setAuthErr("");
    (async () => {
      try {
        const mail = (fbUser.email || "").toLowerCase();
        const ref = fref("staff", mail);
        let snap = await getDoc(ref);
        if (!snap.exists()) {
          // Primer ingreso del dueño: las reglas solo dejan crear este registro al correo del dueño.
          await setDoc(ref, { name: "Administrador", rol: "admin", email: mail, active: true });
          snap = await getDoc(ref);
        }
        const rec = { ...snap.data(), id: mail };
        if (rec.active === false) throw userErr("Tu acceso está desactivado. Habla con el administrador.");
        if (dead) return;
        setUser(rec);
        setMode((m) => (m === "gate" || m === "login" ? "panel" : m));
        setView(can(rec.rol, "agenda") ? "agenda" : "clientes");
      } catch (e) {
        console.error(e);
        if (dead) return;
        setAuthErr(e.userMessage || "Esta cuenta no está vinculada a ningún acceso del negocio. Pide al administrador que te agregue en Ajustes. (Si eres el dueño y es la primera vez, revisa que tu correo esté en las reglas de Firestore.)");
        setMode((m) => (m === "gate" ? "login" : m));
        signOut(auth);
      }
    })();
    return () => { dead = true; };
  }, [fbUser]);

  /* 4) Datos privados: solo se escuchan con sesión de personal válida. */
  useEffect(() => {
    if (!user) { setShop(null); shopRef.current = null; return undefined; }
    return onSnapshot(SHOP, (snap) => {
      if (snap.exists()) { const d = normalize(snap.data()); shopRef.current = d; setShop(d); setErr(null); }
      else if (user.rol === "admin") mutate(() => {}).catch(onFail("la base de datos")); // primera vez: crea la configuración
      else setErr("El administrador debe iniciar sesión primero para configurar el negocio.");
    }, onFail("la base de datos"));
  }, [user, onFail]);

  useEffect(() => {
    if (!user) { setAppts([]); return undefined; }
    const q = query(fcol("appointments"), where("date", ">=", winFrom));
    return onSnapshot(q, (s) => setAppts(s.docs.map((d) => ({ ...d.data(), id: d.id }))), onFail("las citas"));
  }, [user, winFrom, onFail]);

  useEffect(() => {
    if (!user) { setPendAppts([]); pendingRef.current = null; return undefined; }
    const q = query(fcol("appointments"), where("status", "==", "pendiente"));
    return onSnapshot(q, (s) => {
      const list = s.docs.map((d) => ({ ...d.data(), id: d.id }));
      setPendAppts(list);
      const ids = list.map((a) => a.id);
      const nuevas = pendingRef.current ? ids.filter((i) => !pendingRef.current.includes(i)) : [];
      if (nuevas.length) {
        if (!shopRef.current || shopRef.current.settings.soundEnabled !== false) Sound.chime();
        toast(`${nuevas.length} solicitud${nuevas.length > 1 ? "es" : ""} de cita por confirmar`, "brass");
      }
      pendingRef.current = ids;
    }, onFail("las solicitudes"));
  }, [user, onFail, toast]);

  useEffect(() => {
    if (!user) { setWaitlist([]); return undefined; }
    const q = query(fcol("waitlist"), where("status", "==", "esperando"));
    return onSnapshot(q, (s) => setWaitlist(s.docs.map((d) => ({ ...d.data(), id: d.id }))), onFail("la lista de espera"));
  }, [user, onFail]);

  useEffect(() => {
    if (!user) { setStaffList([]); return undefined; }
    if (user.rol !== "admin") { setStaffList([user]); return undefined; }
    return onSnapshot(fcol("staff"), (s) => setStaffList(s.docs.map((d) => ({ ...d.data(), id: d.id }))), onFail("el personal"));
  }, [user, onFail]);

  // El audio requiere un gesto del usuario antes de poder sonar.
  useEffect(() => {
    const g = () => Sound.unlock();
    window.addEventListener("pointerdown", g, { once: true });
    return () => window.removeEventListener("pointerdown", g);
  }, []);

  /** Vista completa que usan las pantallas del panel: shop/main + colecciones. */
  const db = useMemo(() => {
    if (!shop || !user) return null;
    const map = new Map();
    appts.forEach((a) => map.set(a.id, a));
    pendAppts.forEach((a) => map.set(a.id, a));
    return {
      ...shop, appointments: [...map.values()], waitlist, staff: staffList,
      legacy: (shop.appointments || []).length + (shop.waitlist || []).length, // datos del sistema anterior sin migrar
    };
  }, [shop, user, appts, pendAppts, waitlist, staffList]);
  dbRef.current = db;

  /** Lo que ve un visitante sin cuenta (portada, reservas, productos). */
  const pubDb = useMemo(() => {
    const base = { settings: seed().settings, barbers: [], services: [], blocks: [], products: [], clients: [], appointments: [], waitlist: [], expenses: [] };
    return pub ? { ...base, ...pub, settings: { ...base.settings, ...pub.settings } } : base;
  }, [pub]);

  // La marca (logo + paleta + nombre) se aplica ANTES de pintar cualquier pantalla.
  const bizName = db ? db.settings.businessName : pub ? pub.settings.businessName : "";
  applyBranding(brand, bizName);
  const brandKey = [brand && brand.updatedAt, bizName, brand && brand.palette].join("|");
  useEffect(() => { applyDocumentBranding(); }, [brandKey]); // título, color de la barra e ícono de la app instalada

  /* --- escritura: toda acción pasa por aquí --- */
  const fail = useCallback((e) => {
    console.error(e);
    toast(e.userMessage || "No se guardó el cambio. Revisa la conexión e inténtalo otra vez.", "pole");
    return null;
  }, [toast]);

  /** Cambio en shop/main (+ escrituras extra atómicas). Devuelve null si falló. */
  const apply = useCallback(async (fn, okMsg, extra) => {
    try { const next = await mutate(fn, extra); if (okMsg) toast(okMsg, "mint"); return next; }
    catch (e) { return fail(e); }
  }, [toast, fail]);

  /** Escritura que NO toca shop/main (citas, espera, personal…). */
  const run = useCallback(async (job, okMsg) => {
    try { const r = await job(); if (okMsg) toast(okMsg, "mint"); return r === undefined ? true : r; }
    catch (e) { return fail(e); }
  }, [toast, fail]);

  const actions = useMemo(() => {
    const findAppt = (id) => (dbRef.current ? dbRef.current.appointments.find((a) => a.id === id) : null);
    const stamp = () => new Date().toISOString();
    return {
      toast,

      /** La agenda amplía hacia atrás la ventana de citas que se escucha. */
      ensureFrom: (date) => setWinFrom((w) => (date < w ? date : w)),

      /** Alta o edición. Citas, espera y personal viven en su propia colección. */
      save: (coll, item, msg) => {
        if (coll === "appointments" || coll === "waitlist") {
          const { id, ...data } = item;
          return run(() => setDoc(fref(coll, id || uid()), data, { merge: true }), msg);
        }
        if (coll === "staff") {
          const { id: oldId, pin, email, ...rest } = item;
          const mail = (email || "").trim().toLowerCase();
          return run(async () => {
            await setDoc(fref("staff", mail), { ...rest, email: mail }, { merge: true });
            if (oldId && oldId !== mail) await deleteDoc(fref("staff", oldId));
          }, msg);
        }
        return apply((d) => {
          const list = d[coll];
          const i = list.findIndex((x) => x.id === item.id);
          if (i >= 0) list[i] = { ...list[i], ...item }; else list.push({ ...item, id: item.id || uid() });
        }, msg);
      },

      /** Baja física. */
      remove: (coll, id, msg) => {
        if (coll === "appointments" || coll === "waitlist" || coll === "staff") return run(() => deleteDoc(fref(coll, id)), msg);
        return apply((d) => { d[coll] = d[coll].filter((x) => x.id !== id); }, msg);
      },

      /** Cita creada por el personal en el mostrador: queda confirmada y con candados. */
      bookDirect: (p) => {
        const id = uid();
        const newClientId = typeof p.clientId === "string" ? null : uid();
        let appt = null;
        return apply((d) => {
          const svc = d.services.find((s) => s.id === p.serviceId);
          let clientId = p.clientId;
          if (typeof clientId !== "string") {
            clientId = newClientId;
            d.clients.push({ id: clientId, name: p.clientId.newName, phone: p.phone || "", notes: "", createdAt: stamp() });
          }
          appt = {
            id, clientId, barberId: p.barberId, serviceId: p.serviceId, date: p.date, time: p.time,
            durationMin: svc ? svc.durationMin : 30, price: svc ? svc.price : 0,
            status: "confirmada", clientNote: "", comment: p.comment || "", createdAt: stamp(), source: "mostrador",
          };
        }, "Cita agendada", async (tx) => {
          const refs = await claimLocks(tx, appt);
          const { id: _x, ...data } = appt;
          tx.set(fref("appointments", id), data);
          refs.forEach((r) => tx.set(r.ref, r.data));
        });
      },

      /** Acepta una solicitud: la vincula con su ficha de cliente y confirma. */
      accept: (id, comment) => {
        const orig = findAppt(id);
        if (!orig) return Promise.resolve(null);
        let a2 = null;
        return apply((d) => {
          const svc = d.services.find((s) => s.id === orig.serviceId);
          let clientId = orig.clientId;
          if (!clientId || !d.clients.some((c) => c.id === clientId)) {
            const ex = d.clients.find((c) => digits(c.phone) && digits(c.phone) === digits(orig.clientPhone));
            if (ex) clientId = ex.id;
            else {
              clientId = uid();
              d.clients.push({ id: clientId, name: orig.clientName || "Cliente", phone: orig.clientPhone || "", notes: "", createdAt: stamp() });
            }
          }
          // El precio y la duración los manda el servicio, no lo que traiga la solicitud.
          a2 = {
            ...orig, clientId, durationMin: svc ? svc.durationMin : orig.durationMin, price: svc ? svc.price : orig.price,
            status: "confirmada", comment: comment || "", confirmedAt: stamp(),
          };
        }, "Cita confirmada", async (tx) => {
          const refs = await claimLocks(tx, a2);
          const { id: _x, ...data } = a2;
          tx.set(fref("appointments", id), data, { merge: true });
          refs.forEach((r) => tx.set(r.ref, r.data));
        });
      },

      reject: (id, comment) => {
        const a = findAppt(id);
        return run(async () => {
          const b = writeBatch(firestore);
          b.set(fref("appointments", id), { status: "rechazada", comment: comment || "" }, { merge: true });
          if (a) lockDocs(a).forEach((l) => b.delete(fref("locks", l.id)));
          await b.commit();
        }, "Solicitud rechazada");
      },

      /** Cambia el estado. Cancelar libera el horario; atender suma la visita al cliente. */
      setStatus: (id, status) => {
        const a = findAppt(id);
        if (!a) return Promise.resolve(null);
        if (status === "completada") {
          return apply((d) => { bumpClient(d, a.clientId, { visits: 1, spent: a.price, last: a.date }); }, "Cita actualizada",
            async (tx) => { tx.set(fref("appointments", id), { status }, { merge: true }); });
        }
        return run(async () => {
          const b = writeBatch(firestore);
          b.set(fref("appointments", id), { status }, { merge: true });
          if (status === "cancelada" || status === "rechazada") lockDocs(a).forEach((l) => b.delete(fref("locks", l.id)));
          await b.commit();
        }, "Cita actualizada");
      },

      /** Mueve una cita a otra fecha/hora: cambia los candados en la misma transacción. */
      reschedule: (id, { date, time }) => {
        const a = findAppt(id);
        if (!a) return Promise.resolve(null);
        const moved = { ...a, date, time };
        return run(() => runTransaction(firestore, async (tx) => {
          const refs = await claimLocks(tx, moved);
          const keep = new Set(refs.map((r) => r.ref.id));
          lockDocs(a).filter((l) => !keep.has(l.id)).forEach((l) => tx.delete(fref("locks", l.id)));
          refs.forEach((r) => tx.set(r.ref, r.data));
          tx.set(fref("appointments", id), { date, time }, { merge: true });
        }), "Cita reprogramada");
      },

      /** Movimiento de inventario: ajusta el stock y deja rastro. */
      stockMove: (m) => apply((d) => {
        const p = d.products.find((x) => x.id === m.productId);
        if (!p) return;
        const q = Number(m.qty) || 0;
        p.stock = m.type === "entrada" ? p.stock + q : m.type === "salida" ? p.stock - q : q;
        d.moves.push({ ...m, id: uid(), date: m.date || ymd() });
      }, "Inventario actualizado"),

      saveSettings: (s) => apply((d) => { d.settings = { ...d.settings, ...s }; }, "Ajustes guardados"),

      /** Guarda logo y paleta. Es público (lo ven los clientes), por eso va aparte de shop/main. */
      saveBrand: (look) => run(() => setDoc(BRAND, {
        logo: look.logo || "", plate: look.plate === "light" ? "light" : "none",
        palette: look.palette || "clasica", font: look.font || "", custom: look.custom || null,
        updatedAt: stamp(),
      }), "Apariencia guardada"),

      /**
       * COBRO (caja). En una sola transacción: crea la venta con su número
       * consecutivo, descuenta el inventario, suma al cliente y marca la cita
       * como cobrada. Si algo falla, no se guarda nada.
       */
      sell: async (p) => {
        const id = uid();
        let sale = null;
        const res = await apply((d) => {
          const items = p.items.map((i) => ({ type: i.type, refId: i.refId, name: i.name, qty: Number(i.qty) || 0, unitPrice: Number(i.unitPrice) || 0 }));
          if (!items.length || items.some((i) => i.qty <= 0)) throw userErr("Agrega al menos un ítem con cantidad válida.");
          const prods = items.filter((i) => i.type === "product");
          prods.forEach((i) => {
            const pr = d.products.find((x) => x.id === i.refId);
            if (!pr) throw userErr(`El producto "${i.name}" ya no existe.`);
            if ((pr.stock || 0) < i.qty) throw userErr(`No hay existencias suficientes de "${i.name}" (quedan ${pr.stock || 0}).`);
          });
          d.salesSeq = (d.salesSeq || 0) + 1;
          const number = "V-" + String(d.salesSeq).padStart(5, "0");
          const subtotal = sum(items, (i) => i.qty * i.unitPrice);
          const discount = Math.min(Math.max(Number(p.discount) || 0, 0), subtotal);
          const tip = Math.max(Number(p.tip) || 0, 0);
          const total = subtotal - discount + tip;
          const paid = sum(p.payments, (x) => x.amount);
          if (Math.abs(paid - total) > 0.5) throw userErr("Los pagos no suman el total de la venta.");
          const client = p.clientId ? d.clients.find((c) => c.id === p.clientId) : null;
          sale = {
            number, date: ymd(), time: nowHHmm(), createdAt: stamp(),
            clientId: client ? client.id : "", clientName: client ? client.name : "",
            barberId: p.barberId || "", apptId: p.apptId || "",
            items, subtotal, discount, tip, total,
            payments: p.payments.map((x) => ({ method: x.method, amount: Number(x.amount) || 0 })),
            received: Number(p.received) || 0, change: Number(p.change) || 0,
            status: "pagada", staffName: p.staffName || "", note: p.note || "",
          };
          prods.forEach((i) => {
            const pr = d.products.find((x) => x.id === i.refId);
            pr.stock = (pr.stock || 0) - i.qty;
            d.moves.push({ id: uid(), productId: i.refId, type: "salida", qty: i.qty, date: ymd(), note: `Venta ${number}` });
          });
          if (client) {
            if (p.apptCounted) bumpClient(d, client.id, { spent: total - (p.apptPrice || 0), last: ymd() }); // la visita ya se contó al marcarla atendida
            else bumpClient(d, client.id, { visits: 1, spent: total, last: ymd() });
          }
        }, "Venta registrada", async (tx) => {
          tx.set(fref("sales", id), sale);
          if (p.apptId) tx.set(fref("appointments", p.apptId), { status: "completada", saleId: id }, { merge: true });
        });
        return res ? { ...sale, id } : null;
      },

      /** Anula una venta: devuelve el inventario, resta al cliente y reabre la cita. */
      voidSale: (sale, reason) => apply((d) => {
        sale.items.filter((i) => i.type === "product").forEach((i) => {
          const pr = d.products.find((x) => x.id === i.refId);
          if (pr) {
            pr.stock = (pr.stock || 0) + i.qty;
            d.moves.push({ id: uid(), productId: i.refId, type: "entrada", qty: i.qty, date: ymd(), note: `Anula ${sale.number}` });
          }
        });
        if (sale.clientId) bumpClient(d, sale.clientId, { visits: -1, spent: -sale.total });
      }, "Venta anulada", async (tx) => {
        const cur = await tx.get(fref("sales", sale.id));
        if (cur.exists() && cur.data().status === "anulada") throw userErr("Esta venta ya estaba anulada.");
        tx.set(fref("sales", sale.id), { status: "anulada", voidReason: reason || "", voidedAt: stamp() }, { merge: true });
        if (sale.apptId) tx.set(fref("appointments", sale.apptId), { status: "confirmada", saleId: deleteField() }, { merge: true });
      }),

      /** Activa (o cambia) la licencia: verifica el código en Conty Licencias y lo guarda para todo el negocio. */
      activateLicense: (code) => run(async () => {
        const c = String(code).trim().toUpperCase().replace(/\s+/g, "");
        if (!c) throw userErr("Escribe el código de la licencia.");
        let d;
        try { d = await readLicenseRemote(c); }
        catch (e) { throw userErr("No hay conexión para verificar el código. Conéctate a internet e inténtalo de nuevo."); }
        if (d.noExiste) throw userErr("No encontramos ese código. Revisa letras y números.");
        if (d.productoCodigo !== LICENSE.product) throw userErr(`Ese código es de ${d.producto || "otra app"}, no de ${LICENSE.appName}.`);
        await mutate((x) => { x.license = { code: c, d, checkedAt: stamp() }; });
        return true;
      }, "Licencia verificada"),

      /** Vuelve a consultar la licencia guardada (por ejemplo, después de pagar). */
      recheckLicense: () => run(async () => {
        const lic = dbRef.current && dbRef.current.license;
        if (!lic || !lic.code) throw userErr("Todavía no hay un código de licencia.");
        let d;
        try { d = await readLicenseRemote(lic.code); }
        catch (e) { throw userErr("No hay conexión para revisar la licencia. Inténtalo de nuevo."); }
        await mutate((x) => { x.license = d.noExiste ? { code: lic.code, d: null, invalid: true, checkedAt: stamp() } : { code: lic.code, d, checkedAt: stamp() }; });
        return true;
      }, "Licencia revisada"),

      /** Pasa citas y lista de espera del sistema anterior (todo en un documento) a colecciones. */
      migrateLegacy: () => run(async () => {
        const snap = await getDoc(SHOP);
        const raw = snap.exists() ? snap.data() : {};
        const oldAppts = Array.isArray(raw.appointments) ? raw.appointments : [];
        const oldWait = Array.isArray(raw.waitlist) ? raw.waitlist : [];
        const today = ymd();
        const ops = [];
        oldAppts.forEach((a) => {
          const { id, ...data } = a;
          ops.push((b) => b.set(fref("appointments", id), data));
          if ((a.status === "pendiente" || a.status === "confirmada") && a.date >= today) {
            lockDocs({ ...a, id }).forEach((l) => ops.push((b) => b.set(fref("locks", l.id), l.data)));
          }
        });
        oldWait.forEach((w) => { const { id, ...data } = w; ops.push((b) => b.set(fref("waitlist", id), data)); });
        await writeChunks(ops);
        await mutate((d) => {
          d.clients.forEach((c) => {
            if (c.visits !== undefined) return;
            const done = oldAppts.filter((a) => a.clientId === c.id && a.status === "completada");
            c.visits = done.length; c.spent = sum(done, (a) => a.price);
            c.lastVisit = done.map((a) => a.date).sort().pop() || "";
          });
          delete d.appointments; delete d.waitlist; delete d.staff;
        });
        return true;
      }, "Datos migrados al sistema nuevo"),

      /** Restaura un respaldo: reemplaza la configuración y reescribe citas, espera y ventas. */
      restore: (json) => run(async () => {
        const { appointments = [], waitlist = [], sales = [], staff, legacy, exportedAt, ...shopPart } = json;
        await mutate(() => normalize({ ...shopPart }));
        const today = ymd();
        const ops = [];
        appointments.forEach((a) => {
          const { id, ...data } = a;
          ops.push((b) => b.set(fref("appointments", id), data));
          if ((a.status === "pendiente" || a.status === "confirmada") && a.date >= today) {
            lockDocs({ ...a, id }).forEach((l) => ops.push((b) => b.set(fref("locks", l.id), l.data)));
          }
        });
        waitlist.forEach((w) => { const { id, ...data } = w; ops.push((b) => b.set(fref("waitlist", id), data)); });
        sales.forEach((v) => { const { id, ...data } = v; ops.push((b) => b.set(fref("sales", id), data)); });
        await writeChunks(ops);
        return true;
      }, "Respaldo restaurado"),
    };
  }, [apply, run, toast]);

  /* Licencia: mientras el panel está abierto se vuelve a consultar cada 6 h (y al volver a la app). */
  const licCode = shop && shop.license ? shop.license.code : "";
  useEffect(() => {
    if (!LICENSE_ON || !user || !licCode) return undefined;
    let stop = false;
    const check = async () => {
      try {
        const lic = shopRef.current && shopRef.current.license;
        if (!lic || !lic.code) return;
        const age = Date.now() - new Date(lic.checkedAt || 0).getTime();
        if (age < 3600000) return;                                  // se revisó hace menos de 1 hora
        const d = await readLicenseRemote(lic.code);
        if (stop) return;
        const next = d.noExiste ? { code: lic.code, d: null, invalid: true } : { code: lic.code, d };
        const changed = JSON.stringify([lic.d || null, !!lic.invalid]) !== JSON.stringify([next.d || null, !!next.invalid]);
        if (changed || age > DAY_MS) await mutate((x) => { x.license = { ...next, checkedAt: new Date().toISOString() }; });
      } catch (e) { /* sin internet: se respeta lo último que se supo */ }
    };
    check();
    const t = setInterval(check, LICENSE.recheckHours * 3600000);
    const vis = () => { if (document.visibilityState === "visible") check(); };
    document.addEventListener("visibilitychange", vis);
    return () => { stop = true; clearInterval(t); document.removeEventListener("visibilitychange", vis); };
  }, [user, licCode]);

  /* Limpieza discreta: borra candados de días pasados (solo lo hace un administrador). */
  useEffect(() => {
    if (!user || user.rol !== "admin") return;
    (async () => {
      try {
        const q = query(fcol("locks"), where("date", "<", addDays(ymd(), -1)));
        const snap = await getDocs(q);
        if (snap.empty) return;
        await writeChunks(snap.docs.slice(0, 2000).map((d) => (b) => b.delete(d.ref)));
      } catch (e) { console.error(e); }
    })();
  }, [user]);

  /* --- pantallas de carga y portada pública --- */
  if (pub === undefined || fbUser === undefined || brand === undefined) return <LoadingScreen err={err} />;

  if (mode === "gate") {
    const lapsed = !!(pub && (pub.suspended || (pub.licensed && pub.licenseUntilMs && Date.now() >= pub.licenseUntilMs)));
    return <Gate db={pubDb} configured={!!pub} suspended={lapsed}
                 onBook={() => setMode("reservar")} onCatalog={() => setMode("catalogo")} onStaff={() => setMode("login")} />;
  }
  if (mode === "catalogo") {
    return (
      <div className="bt">
        <style>{CSS}</style>
        <CatalogView db={pubDb} onExit={() => setMode("gate")} />
      </div>
    );
  }
  if (mode === "login") return <Login authErr={authErr} onBack={() => setMode("gate")} />;
  if (mode === "reservar") {
    return (
      // El contenedor .bt es obligatorio: ahí viven las variables de color.
      <div className="bt">
        <style>{CSS}</style>
        <BookingView db={pubDb} actions={actions} onExit={() => setMode("gate")} toasts={toasts} />
      </div>
    );
  }

  /* --- panel del personal --- */
  if (!user || !db) return <LoadingScreen err={err || authErr} />;
  if (licenseBlocked(db)) return <LicenseScreen db={db} actions={actions} user={user} toasts={toasts} />;

  const pendientes = db.appointments.filter((a) => a.status === "pendiente").length;
  const allowed = NAV.filter((n) => can(user.rol, n.key));
  const Current = {
    agenda: AgendaView, caja: PosView, ventas: SalesView, solicitudes: RequestsView, espera: WaitlistView, clientes: ClientsView,
    barberos: BarbersView, servicios: ServicesView, productos: ProductsView,
    mantenimiento: MaintenanceView, insumos: SuppliesView, gastos: ExpensesView,
    reportes: ReportsView, ajustes: SettingsView,
  }[view] || AgendaView;
  const startCharge = (pr) => { setPreset(pr); setView("caja"); };

  return (
    <div className="bt min-h-screen flex">
      <style>{CSS}</style>

      {/* --- navegación lateral (escritorio) --- */}
      <aside className="bt-nav hidden md:flex flex-col" style={{ width: 208, flex: "0 0 208px" }}>
        <div className="px-4 py-5 flex flex-col items-center gap-2" style={{ borderBottom: "1px solid rgba(255,255,255,.14)" }}>
          <Logo dark />
          <span className="bt-eyebrow" style={{ color: "rgba(255,255,255,.55)", fontSize: 9 }}>Panel de operación</span>
        </div>
        <nav className="flex-1 py-2 bt-scroll">
          {allowed.map((n) => (
            <button key={n.key} className={`bt-navitem ${view === n.key ? "on" : ""}`} onClick={() => setView(n.key)}>
              <n.icon size={15} /> <span className="flex-1">{n.label}</span>
              {n.key === "solicitudes" && pendientes > 0 && <span className="bt-dot bt-pulse">{pendientes}</span>}
            </button>
          ))}
        </nav>
        <div className="px-4 py-3" style={{ borderTop: "1px solid rgba(255,255,255,.14)" }}>
          <div className="bt-display" style={{ color: "#fff", fontSize: 12 }}>{user.name}</div>
          <div className="bt-eyebrow" style={{ fontSize: 9 }}>{user.rol}</div>
          <button className="bt-btn ghost sm mt-3 w-full" style={{ color: "rgba(255,255,255,.72)", borderColor: "rgba(255,255,255,.14)" }}
                  onClick={() => signOut(auth)}><LogOut size={12} /> Salir</button>
        </div>
      </aside>

      {/* --- contenido --- */}
      <main className="flex-1 min-w-0 flex flex-col">
        <header className="md:hidden bt-nav flex items-center justify-between px-3 py-2 bt-noprint">
          <button className="bt-icon" style={{ background: "transparent", borderColor: "rgba(255,255,255,.14)", color: "#fff" }}
                  onClick={() => setMenuOpen(true)} aria-label="Menú"><Menu size={16} /></button>
          <Logo dark className="" />
          <button className="bt-icon relative" style={{ background: "transparent", borderColor: "rgba(255,255,255,.14)", color: "#fff" }}
                  onClick={() => setView("solicitudes")} aria-label="Solicitudes">
            <Bell size={16} />
            {pendientes > 0 && <span className="bt-dot bt-pulse" style={{ position: "absolute", top: -6, right: -6 }}>{pendientes}</span>}
          </button>
        </header>

        {err && (
          <div className="px-4 py-2 flex items-center gap-2" style={{ background: "#F6E3E2", color: "var(--pole)", fontSize: 12 }}>
            <AlertTriangle size={14} /> {err}
          </div>
        )}
        {LICENSE_ON && ["pronto", "gracia"].includes(licenseStatus(db.license).tipo) && (
          <LicenseBanner status={licenseStatus(db.license)} />
        )}
        {user.rol === "admin" && db.legacy > 0 && (
          <div className="px-4 py-2 flex items-center justify-between gap-2" style={{ background: "#FFF4D6", color: "#7A5B00", fontSize: 12 }}>
            <span><AlertTriangle size={13} style={{ display: "inline", marginRight: 6 }} />Hay citas del sistema anterior sin migrar. No se ven hasta migrarlas.</span>
            <button className="bt-btn sm" onClick={() => setView("ajustes")}>Ir a Ajustes</button>
          </div>
        )}

        <div className="flex-1 p-3 sm:p-5 bt-scroll" style={{ maxHeight: "100vh" }}>
          <Current db={db} actions={actions} user={user} setView={setView} brand={brand}
                   preset={preset} clearPreset={() => setPreset(null)} startCharge={startCharge} />
        </div>
      </main>

      {/* --- menú móvil --- */}
      {menuOpen && (
        <div className="fixed inset-0 z-50 md:hidden" style={{ background: "rgba(20,22,27,.55)" }} onClick={() => setMenuOpen(false)}>
          <div className="bt-nav h-full flex flex-col" style={{ width: 232 }} onClick={(e) => e.stopPropagation()}>
            <div className="px-4 py-5 flex justify-center" style={{ borderBottom: "1px solid rgba(255,255,255,.14)" }}><Logo dark /></div>
            <nav className="flex-1 py-2 bt-scroll">
              {allowed.map((n) => (
                <button key={n.key} className={`bt-navitem ${view === n.key ? "on" : ""}`}
                        onClick={() => { setView(n.key); setMenuOpen(false); }}>
                  <n.icon size={15} /> <span className="flex-1">{n.label}</span>
                  {n.key === "solicitudes" && pendientes > 0 && <span className="bt-dot">{pendientes}</span>}
                </button>
              ))}
            </nav>
            <button className="bt-navitem" onClick={() => signOut(auth)}><LogOut size={15} /> Salir</button>
          </div>
        </div>
      )}

      <Toasts toasts={toasts} />
    </div>
  );
}

function Toasts({ toasts }) {
  return (
    <div className="fixed z-50 bottom-4 left-1/2 flex flex-col gap-2 bt-noprint"
         style={{ transform: "translateX(-50%)", width: "min(92vw,380px)" }}>
      {toasts.map((t) => (
        <div key={t.id} className="bt-toast px-3 py-2.5" style={{
          background: "var(--ink)", color: "#fff", borderRadius: 2, fontSize: 13,
          borderLeft: `3px solid ${t.tone === "pole" ? "var(--pole)" : t.tone === "brass" ? "var(--brass)" : "var(--mint)"}`,
        }}>{t.text}</div>
      ))}
    </div>
  );
}

/* --- puerta de entrada: reservar (cliente) o entrar al panel (personal) --- */
function Gate({ db, configured, suspended, onBook, onCatalog, onStaff }) {
  const off = !configured || suspended;
  return (
    <div className="bt min-h-screen flex flex-col items-center justify-center px-5 py-10">
      <style>{CSS}</style>
      <img src={currentLogo()} alt="" style={{ height: 150, width: "auto", maxWidth: "70vw", objectFit: "contain" }} />
      <h1 className="bt-display mt-5 text-center" style={{ fontSize: 20 }}>{db.settings.businessName}</h1>
      <div className="flex items-center gap-3 mt-2">
        <span style={{ width: 26, height: 1, background: "var(--line)" }} />
        <span className="bt-eyebrow">Reserva tu turno</span>
        <span style={{ width: 26, height: 1, background: "var(--line)" }} />
      </div>
      <div className="flex flex-col gap-2 mt-8 w-full" style={{ maxWidth: 300 }}>
        <button className="bt-btn brass" style={{ padding: "13px" }} disabled={off} onClick={onBook}><Calendar size={15} /> Pedir una cita</button>
        <button className="bt-btn ghost" style={{ padding: "13px" }} disabled={off} onClick={onCatalog}><Package size={14} /> Ver productos</button>
        {!configured && <p style={{ fontSize: 12, color: "var(--slate)", textAlign: "center" }}>El negocio aún no está configurado. El administrador debe entrar al panel una vez para terminarlo.</p>}
        {suspended && <p style={{ fontSize: 12, color: "var(--pole)", textAlign: "center" }}>Servicio temporalmente suspendido. Escríbenos por WhatsApp.</p>}
        <button className="bt-btn ghost sm mt-2" style={{ border: 0, color: "var(--slate)" }} onClick={onStaff}>
          <Lock size={12} /> Entrar al panel
        </button>
      </div>
    </div>
  );
}

/** Acceso del personal con correo y contraseña reales (Firebase Authentication). */
function Login({ authErr, onBack }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  // Si la cuenta entró pero no está vinculada a un acceso, App avisa por aquí.
  useEffect(() => { if (authErr) { setBusy(false); setErr(authErr); } }, [authErr]);

  const tryLogin = async () => {
    if (!email.trim() || !password || busy) return;
    setBusy(true); setErr("");
    try {
      await signInWithEmailAndPassword(auth, email.trim(), password);
      // El paso al panel lo hace App cuando valida el acceso.
    } catch (e) {
      const known = ["auth/invalid-credential", "auth/wrong-password", "auth/user-not-found", "auth/invalid-email"];
      setErr(known.includes(e.code) ? "Correo o contraseña incorrectos." : e.code === "auth/too-many-requests"
        ? "Demasiados intentos. Espera unos minutos e inténtalo de nuevo." : `No se pudo iniciar sesión (${e.code || e.message}).`);
      setBusy(false);
    }
  };

  return (
    <div className="bt min-h-screen flex flex-col items-center justify-center px-5">
      <style>{CSS}</style>
      <Logo className="" />
      <div className="bt-card p-5 mt-6 w-full" style={{ maxWidth: 320 }}>
        <span className="bt-eyebrow">Acceso del personal</span>
        <div className="mt-3 flex flex-col gap-3">
          <Field label="Correo">
            <input className="bt-input" type="email" autoCapitalize="none" autoComplete="username" autoFocus
                   value={email} onChange={(e) => { setEmail(e.target.value); setErr(""); }}
                   onKeyDown={(e) => e.key === "Enter" && tryLogin()} />
          </Field>
          <Field label="Contraseña">
            <input className="bt-input" type="password" autoComplete="current-password" value={password}
                   onChange={(e) => { setPassword(e.target.value); setErr(""); }}
                   onKeyDown={(e) => e.key === "Enter" && tryLogin()} />
          </Field>
        </div>
        {err && <p style={{ color: "var(--pole)", fontSize: 12, marginTop: 8 }}>{err}</p>}
        <button className="bt-btn w-full mt-4" onClick={tryLogin} disabled={busy || !email.trim() || !password}>
          {busy ? "Entrando…" : "Entrar"}
        </button>
        <button className="bt-btn ghost w-full mt-2" onClick={onBack}>Volver</button>
        <p style={{ fontSize: 11, color: "var(--slate)", marginTop: 12 }}>
          Las cuentas las crea el administrador en Firebase → Authentication → Users.
        </p>
      </div>
    </div>
  );
}

/** WhatsApp de cobros de Conty (para «Pagar»/«Renovar»). Se pide una vez y se recuerda. */
function useVendorWhatsapp() {
  const [wa, setWa] = useState("");
  useEffect(() => { let dead = false; readVendorWhatsapp().then((w) => { if (!dead) setWa(w); }); return () => { dead = true; }; }, []);
  return wa;
}
const vendorLink = (wa, text) => (wa ? `https://wa.me/${wa}?text=${encodeURIComponent(text)}` : null);

/** Franja de aviso arriba del panel: la licencia vence pronto o ya venció pero hay días de gracia. */
function LicenseBanner({ status }) {
  const wa = useVendorWhatsapp();
  const link = vendorLink(wa, `Hola, quiero renovar mi licencia de ${LICENSE.appName}${status.d && status.d.negocio ? " para " + status.d.negocio : ""}.`);
  const gracia = status.tipo === "gracia";
  return (
    <div className="px-4 py-2 flex items-center justify-between gap-2 bt-noprint"
         style={{ background: gracia ? "#FBE3D6" : "#FFF4D6", color: gracia ? "#9A3F10" : "#7A5B00", fontSize: 12 }}>
      <span><KeyRound size={13} style={{ display: "inline", marginRight: 6 }} />{status.texto}</span>
      {link && <a className="bt-btn sm" href={link} target="_blank" rel="noreferrer">Renovar</a>}
    </div>
  );
}

/** Pantalla que bloquea el panel cuando la licencia falta, venció, fue suspendida o cancelada. */
function LicenseScreen({ db, actions, user, toasts }) {
  const [code, setCode] = useState("");
  const [busy, setBusy] = useState("");
  const wa = useVendorWhatsapp();
  const st = licenseStatus(db.license);
  const admin = user.rol === "admin";
  const info = {
    falta:    { t: `Activa ${LICENSE.appName}`, x: "Escribe el código de tu licencia. Te lo dio tu proveedor." },
    invalido: { t: "Licencia no válida", x: "Ese código ya no está activo. Escribe uno nuevo o escríbenos para revisarlo." },
    otraApp:  { t: "Código de otra app", x: `Ese código es de ${(st.d && st.d.producto) || "otra app"}, no de ${LICENSE.appName}.` },
    bloqueo:  { t: st.titulo, x: st.texto },
  }[st.tipo] || { t: "Licencia", x: "" };
  const link = vendorLink(wa, `Hola, quiero renovar mi licencia de ${LICENSE.appName}${db.license && db.license.code ? " (" + db.license.code + ")" : ""}${st.d && st.d.negocio ? " para " + st.d.negocio : ""}.`);
  const askCode = st.tipo === "falta" || st.tipo === "invalido" || st.tipo === "otraApp";

  return (
    <div className="bt min-h-screen flex flex-col items-center justify-center px-5">
      <style>{CSS}</style>
      <Logo />
      <div className="bt-card p-5 mt-6 w-full" style={{ maxWidth: 380 }}>
        <div className="flex items-center gap-2"><KeyRound size={16} style={{ color: "var(--brass)" }} /><span className="bt-eyebrow">Licencia</span></div>
        <h2 className="bt-display mt-2" style={{ fontSize: 17 }}>{info.t}</h2>
        <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 6 }}>{info.x}</p>
        {db.license && db.license.code && !askCode && <div className="bt-inset p-2 mt-3 bt-mono text-center" style={{ fontSize: 13 }}>{db.license.code}</div>}

        {admin && (askCode || st.tipo === "bloqueo") && (
          <div className="mt-3">
            <Field label={askCode ? "Código de licencia" : "¿Tienes un código nuevo?"}>
              <input className="bt-input bt-mono" placeholder="BRB-XXXXX" autoCapitalize="characters" value={code} onChange={(e) => setCode(e.target.value)} />
            </Field>
            <button className="bt-btn w-full mt-2" disabled={!code.trim() || busy === "act"}
                    onClick={async () => { setBusy("act"); await actions.activateLicense(code); setBusy(""); setCode(""); }}>
              <ShieldCheck size={13} /> {busy === "act" ? "Verificando…" : "Activar"}
            </button>
          </div>
        )}
        {!admin && askCode && <p style={{ fontSize: 12, marginTop: 10 }}>Pídele al administrador que active la licencia.</p>}

        {link && st.tipo !== "falta" && <a className="bt-btn mint w-full mt-3" href={link} target="_blank" rel="noreferrer"><MessageCircle size={13} /> Pagar por WhatsApp</a>}
        {st.tipo === "bloqueo" && (
          <button className="bt-btn ghost w-full mt-2" disabled={busy === "re"}
                  onClick={async () => { setBusy("re"); await actions.recheckLicense(); setBusy(""); }}>
            <RefreshCw size={13} /> {busy === "re" ? "Revisando…" : "Ya pagué, revisar otra vez"}
          </button>
        )}
        <button className="bt-btn ghost w-full mt-2" onClick={() => signOut(auth)}><LogOut size={12} /> Salir</button>
      </div>
      <Toasts toasts={toasts} />
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 1. Reservar — lo que ve el cliente. Solo crea SOLICITUDES.
   ========================================================================== */
function BookingView({ db, actions, onExit, toasts }) {
  const services = db.services.filter((s) => s.active);
  const barbers = db.barbers.filter((b) => b.active);
  const [serviceId, setServiceId] = useState(services[0] ? services[0].id : "");
  const [barberId, setBarberId] = useState(barbers[0] ? barbers[0].id : "");
  const [date, setDate] = useState(ymd());
  const [time, setTime] = useState("");
  const [form, setForm] = useState({ name: "", phone: "", note: "" });
  const [done, setDone] = useState(false);
  const [waited, setWaited] = useState(false);
  const [bizLink, setBizLink] = useState(null); // enlace de respaldo si el navegador bloquea el auto-envío
  const [busy, setBusy] = useState(false);
  const [locks, setLocks] = useState(() => new Set());

  // Horas ya tomadas ese día: se leen de los "candados" públicos (las citas son privadas).
  useEffect(() => {
    const q = query(fcol("locks"), where("date", "==", date));
    return onSnapshot(q, (snap) => setLocks(new Set(snap.docs.map((d) => d.id))), (e) => console.error(e));
  }, [date]);
  const dbv = useMemo(() => ({ ...db, locks }), [db, locks]);

  const service = services.find((s) => s.id === serviceId);
  const slots = useMemo(
    () => (service && barberId ? freeSlots(dbv, barberId, date, service.durationMin, true) : []),
    [dbv, barberId, date, service]
  );
  const days = useMemo(() => Array.from({ length: 14 }, (_, i) => addDays(ymd(), i)), []);

  const send = async () => {
    if (!form.name.trim() || digits(form.phone).length < 7 || !time || busy) return;
    setBusy(true);
    try {
      // Una transacción crea la solicitud y toma la hora; si otro cliente ganó, falla sin guardar nada.
      await requestAppointment(db, { ...form, serviceId, barberId, date, time });
    } catch (e) {
      setBusy(false);
      if (e.code === "slot-taken") { actions.toast("Ese horario acaba de ocuparse. Elige otro, por favor.", "pole"); setTime(""); }
      else { console.error(e); actions.toast("No se pudo enviar la solicitud. Revisa tu conexión e inténtalo otra vez.", "pole"); }
      return;
    }

    // Aviso inmediato al negocio: se abre WhatsApp en el propio celular del
    // cliente, con el mensaje ya escrito, para que el negocio se entere aunque
    // el panel esté cerrado (la campanilla solo suena con el panel abierto).
    const bar = barbers.find((b) => b.id === barberId);
    const vars = {
      cliente: form.name, telefono: form.phone, negocio: db.settings.businessName,
      fecha: niceDate(date), hora: ampm(time), barbero: bar ? bar.name : "",
      servicio: service.name, duracion: String(service.durationMin),
      precio: money(service.price, db.settings.currency),
      comentario: form.note ? `Nota: ${form.note}` : "",
    };
    const link = waLink(db.settings.businessPhone, fillTemplate(db.settings.tplRequestBusiness, vars));
    setBizLink(link);
    if (link) window.open(link, "_blank");

    setDone(true);
  };

  const joinWaitlist = async () => {
    if (!form.name.trim() || digits(form.phone).length < 7) { actions.toast("Escribe tu nombre y un teléfono válido para anotarte.", "pole"); return; }
    try {
      await setDoc(fref("waitlist", uid()), {
        name: form.name.trim().slice(0, 80), phone: form.phone.trim().slice(0, 20), date, barberId, serviceId,
        note: (form.note || "").slice(0, 300), status: "esperando", createdAt: new Date().toISOString(),
      });
      setWaited(true);
    } catch (e) { console.error(e); actions.toast("No se pudo anotar en la lista de espera. Inténtalo otra vez.", "pole"); }
  };

  if (done) {
    return (
      <div className="bt min-h-screen flex flex-col items-center justify-center px-5 text-center">
        <Logo />
        <div className="bt-card p-6 mt-6 w-full" style={{ maxWidth: 380 }}>
          <Check size={30} style={{ color: "var(--mint)", margin: "0 auto" }} />
          <h2 className="bt-display mt-3" style={{ fontSize: 17 }}>Solicitud enviada</h2>
          <p style={{ fontSize: 14, color: "var(--slate)", marginTop: 8 }}>
            Pedimos <strong>{niceDate(date)}</strong> a las <strong>{ampm(time)}</strong>.
            La barbería la revisará y, al confirmarla, recibirás el mensaje por WhatsApp.
          </p>
          {bizLink && (
            <a className="bt-btn mint w-full mt-4" target="_blank" rel="noreferrer" href={bizLink}>
              <MessageCircle size={13} /> Si no se abrió solo, avisar por WhatsApp
            </a>
          )}
          {!db.settings.businessPhone && (
            <p style={{ fontSize: 11, color: "var(--pole)", marginTop: 8 }}>
              El negocio no tiene un WhatsApp configurado todavía, así que no se pudo avisar automáticamente.
            </p>
          )}
          <button className="bt-btn ghost w-full mt-5" onClick={onExit}>Volver al inicio</button>
        </div>
        <Toasts toasts={toasts} />
      </div>
    );
  }

  return (
    <div className="bt min-h-screen">
      <header className="flex items-center justify-between px-4 py-3" style={{ borderBottom: "1px solid var(--line)", background: "var(--paper)" }}>
        <Logo />
        <button className="bt-btn ghost sm" onClick={onExit}>Salir</button>
      </header>

      <div className="px-4 py-5 mx-auto" style={{ maxWidth: 620 }}>
        <span className="bt-eyebrow">Paso 1 · El servicio</span>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-2">
          {services.map((s) => (
            <button key={s.id} className={`bt-ticket text-left ${serviceId === s.id ? "is-conf" : ""}`}
                    style={{ cursor: "pointer", outline: serviceId === s.id ? "1px solid var(--mint)" : "none" }}
                    onClick={() => { setServiceId(s.id); setTime(""); }}>
              <div style={{ fontWeight: 600, fontSize: 14 }}>{s.name}</div>
              <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>
                {s.durationMin} min · {money(s.price, db.settings.currency)}
              </div>
            </button>
          ))}
        </div>

        <span className="bt-eyebrow block mt-6">Paso 2 · Con quién</span>
        <div className="flex gap-2 mt-2 bt-scroll pb-1">
          {barbers.map((b) => (
            <button key={b.id} className={`bt-slot ${barberId === b.id ? "on" : ""}`} style={{ minWidth: 116, fontFamily: "var(--body)" }}
                    onClick={() => { setBarberId(b.id); setTime(""); }}>{b.name}</button>
          ))}
        </div>

        <span className="bt-eyebrow block mt-6">Paso 3 · El día</span>
        <div className="flex gap-2 mt-2 bt-scroll pb-1">
          {days.map((d) => {
            const b = barbers.find((x) => x.id === barberId);
            const closed = b ? !shiftFor(db, b, d) : true;
            return (
              <button key={d} disabled={closed} className={`bt-slot ${date === d ? "on" : ""}`}
                      style={{ minWidth: 78, opacity: closed ? .35 : 1 }}
                      onClick={() => { setDate(d); setTime(""); }}>
                <div style={{ fontSize: 10, letterSpacing: ".08em" }}>{DOW_S[dowOf(d)].toUpperCase()}</div>
                <div style={{ fontSize: 16, fontWeight: 600 }}>{fromYmd(d).getDate()}</div>
              </button>
            );
          })}
        </div>

        <span className="bt-eyebrow block mt-6">Paso 4 · La hora</span>
        {slots.length > 0 ? (
          <div className="grid grid-cols-3 sm:grid-cols-5 gap-2 mt-2">
            {slots.map((t) => (
              <button key={t} className={`bt-slot ${time === t ? "on" : ""}`} onClick={() => setTime(t)}>{ampm(t)}</button>
            ))}
          </div>
        ) : waited ? (
          <div className="bt-card p-4 mt-2" style={{ borderLeft: "3px solid var(--mint)" }}>
            <p className="bt-display" style={{ fontSize: 13 }}>Quedaste en la lista de espera</p>
            <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 4 }}>
              Si se libera un cupo el {shortDate(date)}, la barbería te escribe.
            </p>
          </div>
        ) : (
          <div className="bt-inset p-4 mt-2">
            <p className="bt-display" style={{ fontSize: 13 }}>Ese día está lleno</p>
            <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 4 }}>
              Prueba otro día, u déjanos tus datos y te avisamos si se libera un cupo.
            </p>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-3">
              <Field label="Tu nombre"><input className="bt-input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
              <Field label="WhatsApp"><input className="bt-input bt-mono" inputMode="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
            </div>
            <button className="bt-btn brass mt-3" onClick={joinWaitlist}><ListPlus size={13} /> Anotarme en la lista de espera</button>
          </div>
        )}

        {time && (
          <div className="bt-card p-4 mt-6">
            <span className="bt-eyebrow">Paso 5 · Tus datos</span>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-2">
              <Field label="Nombre y apellido">
                <input className="bt-input" placeholder="Como te llamamos" value={form.name}
                       onChange={(e) => setForm({ ...form, name: e.target.value })} />
              </Field>
              <Field label="WhatsApp" hint="Con indicativo del país. Ahí llega la confirmación.">
                <input className="bt-input bt-mono" inputMode="tel" placeholder="573001234567" value={form.phone}
                       onChange={(e) => setForm({ ...form, phone: e.target.value })} />
              </Field>
            </div>
            <div className="mt-3">
              <Field label="¿Algo que debamos saber? (opcional)">
                <textarea className="bt-area" rows={2} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
              </Field>
            </div>
            <div className="bt-inset p-3 mt-4" style={{ fontSize: 13 }}>
              <strong>{service.name}</strong> con <strong>{barbers.find((b) => b.id === barberId)?.name}</strong><br />
              {niceDate(date)} · <span className="bt-mono">{ampm(time)}</span> · {money(service.price, db.settings.currency)}
            </div>
            <button className="bt-btn brass w-full mt-4" style={{ padding: 13 }}
                    disabled={busy || !form.name.trim() || digits(form.phone).length < 7} onClick={send}>
              {busy ? "Enviando…" : "Enviar solicitud de cita"}
            </button>
            <p style={{ fontSize: 11, color: "var(--slate)", marginTop: 8, textAlign: "center" }}>
              La cita queda en firme cuando la barbería la confirme.
            </p>
          </div>
        )}
      </div>
      <Toasts toasts={toasts} />
    </div>
  );
}


/* ==========================================================================
   [VISTAS] 1b. Catálogo — los productos que el cliente puede ver.
   Se alimenta del módulo Productos del panel: aparece aquí todo el que esté
   marcado como "Mostrar en el catálogo". No hay carrito ni cobro: el cliente
   pregunta por WhatsApp o en el local.
   ========================================================================== */
function CatalogView({ db, onExit }) {
  const items = db.products.filter((p) => p.showPublic !== false);
  const cur = db.settings.currency;

  return (
    <div className="bt min-h-screen">
      <header className="flex items-center justify-between px-4 py-3"
              style={{ borderBottom: "1px solid var(--line)", background: "var(--paper)" }}>
        <Logo />
        <button className="bt-btn ghost sm" onClick={onExit}>Volver</button>
      </header>

      <div className="px-4 py-5 mx-auto" style={{ maxWidth: 720 }}>
        <h1 className="bt-display" style={{ fontSize: 20 }}>Productos</h1>
        <p className="bt-eyebrow">Lo que tenemos en el local</p>

        {items.length === 0 ? (
          <div className="mt-5"><Empty icon={Package} title="Todavía no hay productos publicados" /></div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-5">
            {items.map((p) => {
              const agotado = !p.available;
              const texto = `Hola, quiero preguntar por ${p.name} en ${db.settings.businessName}.`;
              return (
                <div key={p.id} className="bt-ticket">
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0">
                      <div className="bt-display" style={{ fontSize: 14 }}>{p.name}</div>
                      {p.description && (
                        <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 4 }}>{p.description}</p>
                      )}
                    </div>
                    <span className={`bt-tag ${agotado ? "canc" : "conf"}`}>{agotado ? "agotado" : "disponible"}</span>
                  </div>
                  <div className="flex items-end justify-between gap-2 mt-3">
                    <span className="bt-mono" style={{ fontSize: 17, fontWeight: 600 }}>
                      {p.price ? money(p.price, cur) : "Consultar"}
                    </span>
                    <a className="bt-btn ghost sm" target="_blank" rel="noreferrer"
                       href={waLink(db.settings.businessPhone, texto) || undefined}
                       style={{ opacity: db.settings.businessPhone ? 1 : .4, pointerEvents: db.settings.businessPhone ? "auto" : "none" }}>
                      <MessageCircle size={12} /> Preguntar
                    </a>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        <div className="bt-inset p-4 mt-6 text-center">
          <p style={{ fontSize: 13 }}>¿Prefieres separar tu turno primero?</p>
          <button className="bt-btn brass mt-3" onClick={onExit}><Calendar size={14} /> Pedir una cita</button>
        </div>
      </div>
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 2. Agenda del día — con el poste de barbero como eje de horas.
   ========================================================================== */
function AgendaView({ db, actions, user, startCharge }) {
  const [date, setDate] = useState(ymd());
  // Las citas se escuchan desde unos días atrás: si navegas más atrás, se amplía la ventana.
  useEffect(() => { actions.ensureFrom(date); }, [date, actions]);
  const [newOpen, setNewOpen] = useState(false);
  const [detail, setDetail] = useState(null);

  // Un barbero solo ve su propia agenda.
  const mine = user.rol === "barbero" ? db.barbers.filter((b) => norm(b.name) === norm(user.name)) : db.barbers.filter((b) => b.active);
  const barbers = mine.length ? mine : db.barbers.filter((b) => b.active);
  const [focus, setFocus] = useState("todos");
  const [showDiscarded, setShowDiscarded] = useState(false);

  const shown = focus === "todos" ? barbers : barbers.filter((b) => b.id === focus);
  // Canceladas y rechazadas no se muestran por defecto: siguen guardadas
  // (para reportes e historial), solo dejan de acumularse a la vista.
  const allDay = db.appointments
    .filter((a) => a.date === date && shown.some((b) => b.id === a.barberId))
    .map((a) => hydrate(db, a))
    .sort((x, y) => x.time.localeCompare(y.time));
  const day = allDay.filter((a) => a.status !== "cancelada" && a.status !== "rechazada");
  const discarded = allDay.filter((a) => a.status === "cancelada" || a.status === "rechazada");

  const hours = [];
  for (let m = toMin(db.settings.openTime); m <= toMin(db.settings.closeTime); m += 60) hours.push(toHHmm(m));
  const blocksToday = db.blocks.filter((b) => date >= b.from && date <= b.to && shown.some((x) => x.id === b.barberId));

  return (
    <div>
      <div className="flex flex-wrap items-center gap-2 justify-between">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Agenda</h1>
          <p className="bt-eyebrow">{niceDate(date)}</p>
        </div>
        <div className="flex items-center gap-2">
          <button className="bt-icon" onClick={() => setDate(addDays(date, -1))} aria-label="Día anterior"><ChevronLeft size={16} /></button>
          <input className="bt-input bt-mono" type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ width: 148 }} />
          <button className="bt-icon" onClick={() => setDate(addDays(date, 1))} aria-label="Día siguiente"><ChevronRight size={16} /></button>
          <button className="bt-btn ghost sm" onClick={() => setDate(ymd())}>Hoy</button>
          <button className="bt-btn" onClick={() => setNewOpen(true)}><Plus size={14} /> Nueva cita</button>
        </div>
      </div>

      {barbers.length > 1 && (
        <div className="flex gap-2 mt-4 bt-scroll pb-1">
          <button className={`bt-slot ${focus === "todos" ? "on" : ""}`} style={{ minWidth: 92, fontFamily: "var(--body)" }}
                  onClick={() => setFocus("todos")}>Todos</button>
          {barbers.map((b) => (
            <button key={b.id} className={`bt-slot ${focus === b.id ? "on" : ""}`} style={{ minWidth: 108, fontFamily: "var(--body)" }}
                    onClick={() => setFocus(b.id)}>{b.name}</button>
          ))}
        </div>
      )}

      {blocksToday.length > 0 && (
        <div className="bt-card p-3 mt-4" style={{ borderLeft: "3px solid var(--pole)" }}>
          {blocksToday.map((b) => (
            <div key={b.id} style={{ fontSize: 13 }}>
              <Ban size={12} style={{ display: "inline", color: "var(--pole)", marginRight: 6 }} />
              <strong>{db.barbers.find((x) => x.id === b.barberId)?.name}</strong>: {b.reason}
              {" — "}{b.allDay ? "todo el día" : `${ampm(b.timeStart)} a ${ampm(b.timeEnd)}`}
            </div>
          ))}
        </div>
      )}

      <div className="flex gap-3 mt-4">
        {/* FIRMA: el poste de barbero marca el transcurso del día */}
        <div className="hidden sm:flex flex-col items-stretch" style={{ width: 62, flex: "0 0 62px" }}>
          <div className="flex gap-2 h-full">
            <div className="flex flex-col justify-between py-1">
              {hours.map((h) => <span key={h} className="bt-mono" style={{ fontSize: 11, color: "var(--slate)" }}>{h}</span>)}
            </div>
            <div className="bt-pole" />
          </div>
        </div>

        <div className="flex-1 min-w-0">
          {day.length === 0 ? (
            <Empty icon={Calendar} title="Ningún turno para este día"
                   action={<button className="bt-btn ghost sm" onClick={() => setNewOpen(true)}><Plus size={13} /> Agendar uno</button>} />
          ) : (
            <div className="flex flex-col gap-2">
              {day.map((a) => (
                <button key={a.id} className={`bt-ticket is-${TAG[a.status] || "done"} text-left w-full`}
                        style={{ cursor: "pointer" }} onClick={() => setDetail(a)}>
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <div className="bt-mono" style={{ fontSize: 15, fontWeight: 600 }}>{ampm(a.time)} <span style={{ color: "var(--slate)", fontSize: 12 }}>→ {ampm(a.endTime)}</span></div>
                      <div style={{ fontSize: 14, fontWeight: 600, marginTop: 2 }}>{a.clientName}</div>
                      <div style={{ fontSize: 12, color: "var(--slate)" }}>{a.serviceName} · {a.barberName}</div>
                      {a.comment && <div style={{ fontSize: 12, color: "var(--brass)", marginTop: 3 }}>Nota: {a.comment}</div>}
                    </div>
                    <StatusTag s={a.status} />
                  </div>
                </button>
              ))}
            </div>
          )}

          {discarded.length > 0 && (
            <div className="mt-4">
              <button className="bt-btn ghost sm" onClick={() => setShowDiscarded(!showDiscarded)}>
                {showDiscarded ? "Ocultar" : "Ver"} canceladas y rechazadas ({discarded.length})
              </button>
              {showDiscarded && (
                <div className="flex flex-col gap-2 mt-2">
                  {discarded.map((a) => (
                    <button key={a.id} className={`bt-ticket is-${TAG[a.status] || "done"} text-left w-full`}
                            style={{ cursor: "pointer", opacity: .6 }} onClick={() => setDetail(a)}>
                      <div className="flex items-start justify-between gap-3">
                        <div className="min-w-0">
                          <div className="bt-mono" style={{ fontSize: 14, fontWeight: 600 }}>{ampm(a.time)}</div>
                          <div style={{ fontSize: 13, fontWeight: 600 }}>{a.clientName}</div>
                          <div style={{ fontSize: 11, color: "var(--slate)" }}>{a.serviceName} · {a.barberName}</div>
                        </div>
                        <StatusTag s={a.status} />
                      </div>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {newOpen && <NewAppointmentModal db={db} actions={actions} date={date} onClose={() => setNewOpen(false)} />}
      {detail && <AppointmentModal db={db} actions={actions} appt={detail} startCharge={startCharge} onClose={() => setDetail(null)} />}
    </div>
  );
}

/** Cita creada en el mostrador: queda confirmada de inmediato. */
function NewAppointmentModal({ db, actions, date: initial, onClose }) {
  const services = db.services.filter((s) => s.active);
  const barbers = db.barbers.filter((b) => b.active);
  const [f, setF] = useState({
    clientId: "", phone: "", serviceId: services[0]?.id || "", barberId: barbers[0]?.id || "",
    date: initial, time: "", comment: "",
  });
  const svc = services.find((s) => s.id === f.serviceId);
  const slots = useMemo(() => (svc ? freeSlots(db, f.barberId, f.date, svc.durationMin) : []), [db, f.barberId, f.date, svc]);
  const ready = f.clientId && f.time && svc;

  return (
    <Modal title="Nueva cita" onClose={onClose}>
      <div className="flex flex-col gap-3">
        <Field label="Cliente"><ClientPicker db={db} value={f.clientId} onChange={(v) => setF({ ...f, clientId: v })} /></Field>
        {typeof f.clientId === "object" && (
          <Field label="WhatsApp del nuevo cliente"><input className="bt-input bt-mono" value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field>
        )}
        <div className="grid grid-cols-2 gap-3">
          <Field label="Servicio">
            <select className="bt-select" value={f.serviceId} onChange={(e) => setF({ ...f, serviceId: e.target.value, time: "" })}>
              {services.map((s) => <option key={s.id} value={s.id}>{s.name} · {s.durationMin}m</option>)}
            </select>
          </Field>
          <Field label="Barbero">
            <select className="bt-select" value={f.barberId} onChange={(e) => setF({ ...f, barberId: e.target.value, time: "" })}>
              {barbers.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
            </select>
          </Field>
        </div>
        <Field label="Fecha"><input className="bt-input bt-mono" type="date" min={ymd()} value={f.date} onChange={(e) => setF({ ...f, date: e.target.value, time: "" })} /></Field>
        <Field label="Hora disponible">
          {slots.length ? (
            <div className="grid grid-cols-4 gap-2">
              {slots.map((t) => <button key={t} className={`bt-slot ${f.time === t ? "on" : ""}`} onClick={() => setF({ ...f, time: t })}>{t}</button>)}
            </div>
          ) : <p style={{ fontSize: 12, color: "var(--pole)" }}>Sin cupos libres: revisa la jornada del barbero o sus bloqueos.</p>}
        </Field>
        <Field label="Comentario (opcional)"><textarea className="bt-area" rows={2} value={f.comment} onChange={(e) => setF({ ...f, comment: e.target.value })} /></Field>
        <button className="bt-btn" disabled={!ready} onClick={async () => { const ok = await actions.bookDirect(f); if (ok) onClose(); }}>Agendar cita</button>
      </div>
    </Modal>
  );
}

/** Detalle de una cita ya existente: cambiar estado o reenviar WhatsApp. */
function AppointmentModal({ db, actions, appt, onClose, startCharge }) {
  const h = hydrate(db, appt);
  const canCharge = (appt.status === "confirmada" || appt.status === "completada") && !appt.saleId;
  const [comment, setComment] = useState(appt.comment || "");
  const [rescheduling, setRescheduling] = useState(false);
  const canReschedule = appt.status === "pendiente" || appt.status === "confirmada";

  return (
    <Modal title="Detalle de la cita" onClose={onClose}>
      <div className="bt-inset p-3" style={{ fontSize: 14 }}>
        <div className="bt-mono" style={{ fontSize: 16, fontWeight: 600 }}>{ampm(appt.time)} → {ampm(h.endTime)}</div>
        <div className="mt-1"><strong>{h.clientName}</strong> · <span className="bt-mono">{h.clientPhone || "sin teléfono"}</span></div>
        <div style={{ color: "var(--slate)" }}>{h.serviceName} con {h.barberName} · {niceDate(appt.date)}</div>
        <div className="mt-2"><StatusTag s={appt.status} /></div>
        {appt.clientNote && <div style={{ fontSize: 13, marginTop: 8 }}>El cliente escribió: “{appt.clientNote}”</div>}
      </div>

      {canReschedule && (
        <button className="bt-btn ghost sm w-full mt-3" onClick={() => setRescheduling(true)}>
          <Calendar size={12} /> Cambiar fecha u hora
        </button>
      )}

      <div className="mt-3">
        <Field label="Comentario interno / para el mensaje">
          <textarea className="bt-area" rows={2} value={comment} onChange={(e) => setComment(e.target.value)} />
        </Field>
        <button className="bt-btn ghost sm mt-2" onClick={() => actions.save("appointments", { id: appt.id, comment })}>
          <Save size={12} /> Guardar comentario
        </button>
      </div>
      <WhatsAppPanel db={db} appt={appt} comment={comment} />
      {appt.saleId && (
        <div className="bt-inset p-2 mt-4" style={{ fontSize: 12 }}>
          <Receipt size={12} style={{ display: "inline", marginRight: 6 }} />Cobrada en caja. Para deshacerlo, anula la venta desde Ventas.
        </div>
      )}
      {canCharge && startCharge && (
        <button className="bt-btn brass w-full mt-4" onClick={() => {
          onClose();
          startCharge({
            apptId: appt.id, clientId: appt.clientId || "", barberId: appt.barberId,
            apptCounted: appt.status === "completada", apptPrice: appt.price || 0,
            items: [{ type: "service", refId: appt.serviceId, name: h.serviceName, qty: 1, unitPrice: appt.price || 0 }],
          });
        }}>
          <CreditCard size={14} /> Cobrar
        </button>
      )}
      <div className="grid grid-cols-2 gap-2 mt-3">
        {appt.status === "confirmada" && (
          <button className="bt-btn ghost" onClick={async () => { await actions.setStatus(appt.id, "completada"); onClose(); }}>
            <Check size={13} /> Atendida (sin cobrar)
          </button>
        )}
        {appt.status !== "cancelada" && !appt.saleId && (
          <button className="bt-btn danger" onClick={async () => { await actions.setStatus(appt.id, "cancelada"); onClose(); }}>
            <X size={13} /> Cancelar cita
          </button>
        )}
      </div>

      {rescheduling && (
        <RescheduleModal db={db} actions={actions} appt={appt}
                          onClose={() => setRescheduling(false)}
                          onSaved={onClose} />
      )}
    </Modal>
  );
}

/**
 * Cambia la fecha y/o la hora de una cita ya existente (pendiente o
 * confirmada). Reutiliza freeSlots con ignoreApptId: la propia cita no
 * cuenta como choque contra sí misma, así su horario actual sigue apareciendo
 * como disponible si nada más lo ocupó mientras tanto.
 */
function RescheduleModal({ db, actions, appt, onClose, onSaved }) {
  const h = hydrate(db, appt);
  const [date, setDate] = useState(appt.date);
  const [time, setTime] = useState(appt.time);
  const [busy, setBusy] = useState(false);

  const slots = useMemo(
    () => freeSlots(db, appt.barberId, date, appt.durationMin, false, appt.id),
    [db, appt, date]
  );

  return (
    <Modal title="Cambiar fecha u hora" onClose={onClose}>
      <div className="bt-inset p-3" style={{ fontSize: 13 }}>
        <strong>{h.clientName}</strong> · {h.serviceName} con {h.barberName}<br />
        Actual: {niceDate(appt.date)} · <span className="bt-mono">{ampm(appt.time)}</span>
      </div>
      <div className="mt-3">
        <Field label="Nueva fecha">
          <input className="bt-input bt-mono" type="date" min={ymd()} value={date}
                 onChange={(e) => { setDate(e.target.value); if (e.target.value !== appt.date) setTime(""); }} />
        </Field>
      </div>
      <div className="mt-3">
        <Field label="Nueva hora">
          {slots.length ? (
            <div className="grid grid-cols-4 gap-2">
              {slots.map((t) => (
                <button key={t} className={`bt-slot ${time === t ? "on" : ""}`} onClick={() => setTime(t)}>{t}</button>
              ))}
            </div>
          ) : <p style={{ fontSize: 12, color: "var(--pole)" }}>Sin cupos libres ese día para este profesional.</p>}
        </Field>
      </div>
      <button className="bt-btn w-full mt-4" disabled={!date || !time || busy}
              onClick={async () => {
                setBusy(true);
                const ok = await actions.reschedule(appt.id, { date, time });
                if (ok) onSaved(); else setBusy(false);
              }}>
        <Save size={13} /> {busy ? "Guardando…" : "Guardar nuevo horario"}
      </button>
    </Modal>
  );
}

/** Botones de envío por WhatsApp (cliente y negocio) para una cita. */
function WhatsAppPanel({ db, appt, comment, auto = false }) {
  const vars = msgVars(db, appt, comment);
  const h = hydrate(db, appt);
  const linkClient = waLink(h.clientPhone, fillTemplate(db.settings.tplClient, vars));
  const linkBiz = waLink(db.settings.businessPhone, fillTemplate(db.settings.tplBusiness, vars));
  const [sent, setSent] = useState({});

  useEffect(() => { if (auto && linkClient) { window.open(linkClient, "_blank"); setSent((s) => ({ ...s, c: true })); } }, [auto]); // eslint-disable-line

  return (
    <div className="bt-card p-3 mt-4" style={{ borderLeft: "3px solid var(--mint)" }}>
      <span className="bt-eyebrow">Confirmación por WhatsApp</span>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-2">
        <a className="bt-btn mint" href={linkClient || undefined} target="_blank" rel="noreferrer"
           style={{ opacity: linkClient ? 1 : .4, pointerEvents: linkClient ? "auto" : "none" }}
           onClick={() => setSent((s) => ({ ...s, c: true }))}>
          <MessageCircle size={13} /> {sent.c ? "Reenviar al cliente" : "Enviar al cliente"}
        </a>
        <a className="bt-btn ghost" href={linkBiz || undefined} target="_blank" rel="noreferrer"
           style={{ opacity: linkBiz ? 1 : .4, pointerEvents: linkBiz ? "auto" : "none" }}
           onClick={() => setSent((s) => ({ ...s, b: true }))}>
          <MessageCircle size={13} /> {sent.b ? "Reenviado al negocio" : "Enviar al negocio"}
        </a>
      </div>
      {!linkClient && <p style={{ fontSize: 11, color: "var(--pole)", marginTop: 6 }}>Este cliente no tiene teléfono registrado.</p>}
      {!linkBiz && <p style={{ fontSize: 11, color: "var(--pole)", marginTop: 6 }}>Falta el WhatsApp del negocio: agrégalo en Ajustes.</p>}
      <details className="mt-2">
        <summary style={{ fontSize: 11, color: "var(--slate)", cursor: "pointer" }}>Ver el texto que se enviará</summary>
        <pre className="bt-inset p-2 mt-2" style={{ fontSize: 11, whiteSpace: "pre-wrap", fontFamily: "var(--body)" }}>
{fillTemplate(db.settings.tplClient, vars)}
        </pre>
      </details>
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 3. Solicitudes — aceptar o rechazar, con comentario opcional.
   ========================================================================== */
function RequestsView({ db, actions }) {
  const pend = db.appointments.filter((a) => a.status === "pendiente")
    .map((a) => hydrate(db, a))
    .sort((x, y) => (x.date + x.time).localeCompare(y.date + y.time));
  const [accepting, setAccepting] = useState(null);
  const [confirmed, setConfirmed] = useState(null); // cita recién aceptada -> panel WhatsApp

  return (
    <div>
      <h1 className="bt-display" style={{ fontSize: 20 }}>Solicitudes</h1>
      <p className="bt-eyebrow">{pend.length ? `${pend.length} por revisar` : "Todo al día"}</p>

      <div className="flex flex-col gap-2 mt-4">
        {pend.length === 0 && <Empty icon={Bell} title="No hay solicitudes por confirmar" />}
        {pend.map((a) => (
          <div key={a.id} className="bt-ticket is-pend">
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div className="min-w-0">
                <div style={{ fontSize: 15, fontWeight: 600 }}>{a.clientName}</div>
                <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>{a.clientPhone}</div>
                <div style={{ fontSize: 13, marginTop: 4 }}>
                  {shortDate(a.date)} · <span className="bt-mono">{ampm(a.time)}</span> · {a.serviceName} · {a.barberName}
                </div>
                {a.clientNote && <div style={{ fontSize: 12, color: "var(--slate)", marginTop: 3 }}>“{a.clientNote}”</div>}
              </div>
              <div className="flex gap-2">
                <button className="bt-btn mint sm" onClick={() => setAccepting(a)}><Check size={13} /> Aceptar</button>
                <button className="bt-btn danger sm" onClick={() => actions.reject(a.id, "")}><X size={13} /> Rechazar</button>
              </div>
            </div>
          </div>
        ))}
      </div>

      {accepting && (
        <AcceptModal db={db} actions={actions} appt={accepting}
                     onClose={() => setAccepting(null)}
                     onDone={(withComment) => { setAccepting(null); setConfirmed({ appt: accepting, comment: withComment }); }} />
      )}
      {confirmed && (
        <Modal title="Cita confirmada" onClose={() => setConfirmed(null)}>
          <p style={{ fontSize: 14 }}>
            Quedó agendada para <strong>{niceDate(confirmed.appt.date)}</strong> a las <strong>{ampm(confirmed.appt.time)}</strong>.
            Envía los avisos y listo.
          </p>
          <WhatsAppPanel db={db} appt={{ ...confirmed.appt, comment: confirmed.comment }} comment={confirmed.comment} auto />
          <button className="bt-btn ghost w-full mt-3" onClick={() => setConfirmed(null)}>Cerrar</button>
        </Modal>
      )}
    </div>
  );
}

/** Aceptar: valida que el cupo siga libre y permite añadir un comentario. */
function AcceptModal({ db, actions, appt, onClose, onDone }) {
  const [comment, setComment] = useState("");
  const h = hydrate(db, appt);
  const free = isFree(db, appt.barberId, appt.date, appt.time, appt.durationMin, appt.id);
  return (
    <Modal title="Aceptar la cita" onClose={onClose}>
      <div className="bt-inset p-3" style={{ fontSize: 14 }}>
        <strong>{h.clientName}</strong> · <span className="bt-mono">{h.clientPhone}</span><br />
        {niceDate(appt.date)} · <span className="bt-mono">{ampm(appt.time)}</span><br />
        <span style={{ color: "var(--slate)" }}>{h.serviceName} con {h.barberName} · {money(appt.price, db.settings.currency)}</span>
      </div>
      {!free && (
        <p className="mt-3" style={{ color: "var(--pole)", fontSize: 13 }}>
          <AlertTriangle size={13} style={{ display: "inline", marginRight: 4 }} />
          Ese espacio ya está ocupado o bloqueado. Rechaza la solicitud o mueve la cita desde la agenda.
        </p>
      )}
      <div className="mt-3">
        <Field label="Comentario para el cliente (opcional)" hint="Se incluye en el WhatsApp. Ej: llega 5 minutos antes.">
          <textarea className="bt-area" rows={3} value={comment} onChange={(e) => setComment(e.target.value)} />
        </Field>
      </div>
      <button className="bt-btn mint w-full mt-4" disabled={!free}
              onClick={async () => { const ok = await actions.accept(appt.id, comment); if (ok) onDone(comment); }}>
        <Check size={14} /> Confirmar y avisar por WhatsApp
      </button>
    </Modal>
  );
}

/* ==========================================================================
   [VISTAS] 4. Lista de espera.
   ========================================================================== */
function WaitlistView({ db, actions }) {
  const list = (db.waitlist || []).filter((w) => w.status === "esperando")
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  return (
    <div>
      <h1 className="bt-display" style={{ fontSize: 20 }}>Lista de espera</h1>
      <p className="bt-eyebrow">Clientes que quedaron sin cupo</p>
      <div className="flex flex-col gap-2 mt-4">
        {list.length === 0 && <Empty icon={ListPlus} title="Nadie esperando cupo" />}
        {list.map((w) => {
          const svc = db.services.find((s) => s.id === w.serviceId);
          const bar = db.barbers.find((b) => b.id === w.barberId);
          const text = `Hola ${w.name}, se liberó un cupo en ${db.settings.businessName} para el ${niceDate(w.date)}. ¿Lo tomas?`;
          return (
            <div key={w.id} className="bt-card p-3 flex flex-wrap items-center justify-between gap-3">
              <div>
                <div style={{ fontSize: 14, fontWeight: 600 }}>{w.name}</div>
                <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>{w.phone}</div>
                <div style={{ fontSize: 12, marginTop: 3 }}>
                  Quería {shortDate(w.date)} · {svc?.name || "—"}{bar ? ` · ${bar.name}` : ""}
                </div>
                {w.note && <div style={{ fontSize: 12, color: "var(--slate)" }}>“{w.note}”</div>}
              </div>
              <div className="flex gap-2">
                <a className="bt-btn mint sm" target="_blank" rel="noreferrer" href={waLink(w.phone, text) || undefined}>
                  <MessageCircle size={12} /> Avisar
                </a>
                <button className="bt-btn ghost sm" onClick={() => actions.save("waitlist", { id: w.id, status: "atendido" }, "Retirado de la lista")}>
                  <Check size={12} /> Resuelto
                </button>
                <button className="bt-icon" onClick={() => actions.remove("waitlist", w.id)} aria-label="Borrar"><Trash2 size={13} /></button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 5. Clientes — búsqueda instantánea e historial.
   ========================================================================== */
function ClientsView({ db, actions }) {
  const [q, setQ] = useState("");
  const [edit, setEdit] = useState(null);
  const list = useMemo(() => {
    const n = norm(q);
    const base = db.clients.map((c) => {
      // Visitas y gasto se acumulan en la ficha al cobrar/atender (las citas viejas ya no se cargan todas).
      if (c.visits !== undefined) return { ...c, visits: c.visits || 0, spent: c.spent || 0, last: c.lastVisit || "" };
      const done = db.appointments.filter((a) => a.clientId === c.id && a.status === "completada");
      return { ...c, visits: done.length, spent: sum(done, (a) => a.price), last: done.map((a) => a.date).sort().pop() || "" };
    });
    const f = n ? base.filter((c) => norm(c.name).includes(n) || digits(c.phone).includes(digits(q))) : base;
    return f.sort((a, b) => (b.last || "").localeCompare(a.last || "") || norm(a.name).localeCompare(norm(b.name)));
  }, [q, db]);

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Clientes</h1>
          <p className="bt-eyebrow">{db.clients.length} registrados</p>
        </div>
        <div className="flex gap-2">
          <button className="bt-btn ghost sm" onClick={() => toExcel(`clientes-${ymd()}.xlsx`, [{ name: "Clientes", rows: list.map((c) => ({ Nombre: c.name, Teléfono: c.phone, Visitas: c.visits, Total: c.spent, "Última visita": c.last, Notas: c.notes || "" })) }])}>
            <Download size={13} /> Excel
          </button>
          <button className="bt-btn" onClick={() => setEdit({ id: "", name: "", phone: "", notes: "" })}><Plus size={14} /> Nuevo</button>
        </div>
      </div>

      <div className="relative mt-4" style={{ maxWidth: 420 }}>
        <Search size={15} style={{ position: "absolute", left: 11, top: 12, color: "var(--slate)" }} />
        <input className="bt-input" style={{ paddingLeft: 32 }} placeholder="Buscar por nombre o teléfono"
               value={q} onChange={(e) => setQ(e.target.value)} />
      </div>

      <div className="flex flex-col gap-2 mt-4">
        {list.length === 0 && <Empty icon={Users} title={q ? "Ningún cliente coincide" : "Aún no hay clientes"} />}
        {list.map((c) => (
          <div key={c.id} className="bt-card p-3 flex items-center justify-between gap-3">
            <div className="min-w-0">
              <div style={{ fontSize: 14, fontWeight: 600 }}>{c.name}</div>
              <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>{c.phone || "sin teléfono"}</div>
              <div style={{ fontSize: 11, color: "var(--slate)", marginTop: 2 }}>
                {c.visits} visita{c.visits === 1 ? "" : "s"} · {money(c.spent, db.settings.currency)}
                {c.last && ` · última ${shortDate(c.last)}`}
              </div>
            </div>
            <div className="flex gap-1">
              {c.phone && <a className="bt-icon" target="_blank" rel="noreferrer" href={waLink(c.phone, `Hola ${c.name}, te escribimos de ${db.settings.businessName}.`)} aria-label="WhatsApp"><MessageCircle size={14} /></a>}
              <button className="bt-icon" onClick={() => setEdit(c)} aria-label="Editar"><Pencil size={13} /></button>
            </div>
          </div>
        ))}
      </div>

      {edit && (
        <Modal title={edit.id ? "Editar cliente" : "Nuevo cliente"} onClose={() => setEdit(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Nombre"><input className="bt-input" value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} /></Field>
            <Field label="WhatsApp" hint="Con indicativo del país, solo números."><input className="bt-input bt-mono" inputMode="tel" value={edit.phone} onChange={(e) => setEdit({ ...edit, phone: e.target.value })} /></Field>
            <Field label="Notas (preferencias, alergias, etc.)"><textarea className="bt-area" rows={3} value={edit.notes || ""} onChange={(e) => setEdit({ ...edit, notes: e.target.value })} /></Field>
            <div className="grid grid-cols-2 gap-2">
              <button className="bt-btn" disabled={!edit.name.trim()}
                      onClick={async () => { await actions.save("clients", { ...edit, id: edit.id || uid(), createdAt: edit.createdAt || new Date().toISOString() }, "Cliente guardado"); setEdit(null); }}>
                <Save size={13} /> Guardar
              </button>
              {edit.id && (
                <button className="bt-btn danger" onClick={async () => {
                  if (db.appointments.some((a) => a.clientId === edit.id)) { actions.toast("No se puede borrar: tiene citas registradas.", "pole"); return; }
                  await actions.remove("clients", edit.id, "Cliente eliminado"); setEdit(null);
                }}><Trash2 size={13} /> Eliminar</button>
              )}
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 6. Barberos — jornada semanal y bloqueos por vacaciones/permisos.
   ========================================================================== */
function BarbersView({ db, actions }) {
  const [edit, setEdit] = useState(null);
  const [blocking, setBlocking] = useState(null);
  return (
    <div>
      <div className="flex items-center justify-between">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Barberos</h1>
          <p className="bt-eyebrow">Jornadas y ausencias</p>
        </div>
        <button className="bt-btn" onClick={() => setEdit({ id: "", name: "", phone: "", active: true, hours: [1,2,3,4,5,6].map((dow) => ({ dow, start: "09:00", end: "19:00" })) })}>
          <Plus size={14} /> Nuevo
        </button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3 mt-4">
        {db.barbers.map((b) => {
          const blocks = db.blocks.filter((x) => x.barberId === b.id && x.to >= ymd()).sort((x, y) => x.from.localeCompare(y.from));
          return (
            <div key={b.id} className="bt-card p-4">
              <div className="flex items-start justify-between gap-2">
                <div>
                  <div className="bt-display" style={{ fontSize: 15 }}>{b.name}</div>
                  <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>{b.phone || "sin teléfono"}</div>
                  {!b.active && <span className="bt-tag canc mt-1">inactivo</span>}
                </div>
                <div className="flex gap-1">
                  <button className="bt-icon" onClick={() => setEdit(clone(b))} aria-label="Editar"><Pencil size={13} /></button>
                  <button className="bt-icon" onClick={() => setBlocking({ id: "", barberId: b.id, from: ymd(), to: ymd(), reason: "", allDay: true, timeStart: "09:00", timeEnd: "13:00" })} aria-label="Bloquear"><Ban size={13} /></button>
                </div>
              </div>

              <div className="bt-inset p-2 mt-3" style={{ fontSize: 12 }}>
                {(b.hours || []).length === 0 ? <span style={{ color: "var(--slate)" }}>Sin jornada definida (usa el horario general).</span> :
                  (b.hours || []).sort((x, y) => x.dow - y.dow).map((h) => (
                    <div key={h.dow} className="flex justify-between">
                      <span>{DOW[h.dow]}</span><span className="bt-mono">{h.start} – {h.end}</span>
                    </div>
                  ))}
              </div>

              {blocks.length > 0 && (
                <div className="mt-3">
                  <span className="bt-eyebrow">Ausencias próximas</span>
                  {blocks.map((x) => (
                    <div key={x.id} className="flex items-center justify-between mt-1" style={{ fontSize: 12 }}>
                      <span>
                        <Ban size={11} style={{ display: "inline", color: "var(--pole)", marginRight: 4 }} />
                        {shortDate(x.from)}{x.from !== x.to ? ` – ${shortDate(x.to)}` : ""} · {x.reason}
                        {!x.allDay && <span className="bt-mono"> {x.timeStart}-{x.timeEnd}</span>}
                      </span>
                      <button className="bt-icon" style={{ width: 24, height: 24 }} onClick={() => actions.remove("blocks", x.id, "Bloqueo eliminado")}><X size={11} /></button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {edit && (
        <Modal title={edit.id ? "Editar barbero" : "Nuevo barbero"} onClose={() => setEdit(null)} wide>
          <div className="flex flex-col gap-3">
            <div className="grid grid-cols-2 gap-3">
              <Field label="Nombre"><input className="bt-input" value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} /></Field>
              <Field label="WhatsApp"><input className="bt-input bt-mono" value={edit.phone || ""} onChange={(e) => setEdit({ ...edit, phone: e.target.value })} /></Field>
            </div>
            <div>
              <span className="bt-label">Jornada semanal</span>
              {DOW.map((d, i) => {
                const h = (edit.hours || []).find((x) => x.dow === i);
                return (
                  <div key={i} className="flex items-center gap-2 mt-1">
                    <label className="flex items-center gap-2" style={{ width: 118, fontSize: 13 }}>
                      <input type="checkbox" checked={!!h} onChange={(e) => {
                        const hours = (edit.hours || []).filter((x) => x.dow !== i);
                        if (e.target.checked) hours.push({ dow: i, start: db.settings.openTime, end: db.settings.closeTime });
                        setEdit({ ...edit, hours });
                      }} />{d}
                    </label>
                    {h && <>
                      <input className="bt-input bt-mono" type="time" style={{ width: 118 }} value={h.start}
                             onChange={(e) => setEdit({ ...edit, hours: edit.hours.map((x) => x.dow === i ? { ...x, start: e.target.value } : x) })} />
                      <input className="bt-input bt-mono" type="time" style={{ width: 118 }} value={h.end}
                             onChange={(e) => setEdit({ ...edit, hours: edit.hours.map((x) => x.dow === i ? { ...x, end: e.target.value } : x) })} />
                    </>}
                  </div>
                );
              })}
            </div>
            <label className="flex items-center gap-2" style={{ fontSize: 13 }}>
              <input type="checkbox" checked={edit.active !== false} onChange={(e) => setEdit({ ...edit, active: e.target.checked })} />
              Recibe citas
            </label>
            <button className="bt-btn" disabled={!edit.name.trim()}
                    onClick={async () => { await actions.save("barbers", { ...edit, id: edit.id || uid() }, "Barbero guardado"); setEdit(null); }}>
              <Save size={13} /> Guardar
            </button>
          </div>
        </Modal>
      )}

      {blocking && (
        <Modal title="Bloquear agenda" onClose={() => setBlocking(null)}>
          <div className="flex flex-col gap-3">
            <div className="grid grid-cols-2 gap-3">
              <Field label="Desde"><input className="bt-input bt-mono" type="date" value={blocking.from} onChange={(e) => setBlocking({ ...blocking, from: e.target.value })} /></Field>
              <Field label="Hasta"><input className="bt-input bt-mono" type="date" value={blocking.to} onChange={(e) => setBlocking({ ...blocking, to: e.target.value })} /></Field>
            </div>
            <Field label="Motivo"><input className="bt-input" placeholder="Vacaciones, permiso, cita médica…" value={blocking.reason} onChange={(e) => setBlocking({ ...blocking, reason: e.target.value })} /></Field>
            <label className="flex items-center gap-2" style={{ fontSize: 13 }}>
              <input type="checkbox" checked={blocking.allDay} onChange={(e) => setBlocking({ ...blocking, allDay: e.target.checked })} /> Todo el día
            </label>
            {!blocking.allDay && (
              <div className="grid grid-cols-2 gap-3">
                <Field label="Desde"><input className="bt-input bt-mono" type="time" value={blocking.timeStart} onChange={(e) => setBlocking({ ...blocking, timeStart: e.target.value })} /></Field>
                <Field label="Hasta"><input className="bt-input bt-mono" type="time" value={blocking.timeEnd} onChange={(e) => setBlocking({ ...blocking, timeEnd: e.target.value })} /></Field>
              </div>
            )}
            <p style={{ fontSize: 11, color: "var(--slate)" }}>Los horarios bloqueados desaparecen de las opciones que ve el cliente.</p>
            <button className="bt-btn" disabled={!blocking.reason.trim() || blocking.to < blocking.from}
                    onClick={async () => { await actions.save("blocks", { ...blocking, id: uid() }, "Agenda bloqueada"); setBlocking(null); }}>
              <Ban size={13} /> Bloquear
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 7. Servicios.
   ========================================================================== */
function ServicesView({ db, actions }) {
  const [edit, setEdit] = useState(null);
  return (
    <div>
      <div className="flex items-center justify-between">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Servicios</h1>
          <p className="bt-eyebrow">Duración y precio</p>
        </div>
        <button className="bt-btn" onClick={() => setEdit({ id: "", name: "", durationMin: 30, price: 0, active: true })}><Plus size={14} /> Nuevo</button>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 mt-4">
        {db.services.map((s) => (
          <div key={s.id} className="bt-card p-4 flex items-start justify-between">
            <div>
              <div className="bt-display" style={{ fontSize: 14 }}>{s.name}</div>
              <div className="bt-mono" style={{ fontSize: 13, color: "var(--slate)", marginTop: 3 }}>
                {s.durationMin} min · {money(s.price, db.settings.currency)}
              </div>
              {!s.active && <span className="bt-tag canc mt-2">oculto</span>}
            </div>
            <button className="bt-icon" onClick={() => setEdit(clone(s))} aria-label="Editar"><Pencil size={13} /></button>
          </div>
        ))}
      </div>
      {edit && (
        <Modal title={edit.id ? "Editar servicio" : "Nuevo servicio"} onClose={() => setEdit(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Nombre"><input className="bt-input" value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} /></Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Duración (min)" hint="Define cuánto ocupa en la agenda.">
                <input className="bt-input bt-mono" type="number" min="5" step="5" value={edit.durationMin} onChange={(e) => setEdit({ ...edit, durationMin: Number(e.target.value) })} />
              </Field>
              <Field label="Precio"><input className="bt-input bt-mono" type="number" min="0" value={edit.price} onChange={(e) => setEdit({ ...edit, price: Number(e.target.value) })} /></Field>
            </div>
            <label className="flex items-center gap-2" style={{ fontSize: 13 }}>
              <input type="checkbox" checked={edit.active !== false} onChange={(e) => setEdit({ ...edit, active: e.target.checked })} /> Visible para los clientes
            </label>
            <div className="grid grid-cols-2 gap-2">
              <button className="bt-btn" disabled={!edit.name.trim()}
                      onClick={async () => { await actions.save("services", { ...edit, id: edit.id || uid() }, "Servicio guardado"); setEdit(null); }}>
                <Save size={13} /> Guardar
              </button>
              {edit.id && (
                <button className="bt-btn danger" onClick={async () => {
                  if (db.appointments.some((a) => a.serviceId === edit.id)) { actions.toast("Tiene citas asociadas: márcalo como oculto en lugar de borrarlo.", "pole"); return; }
                  await actions.remove("services", edit.id, "Servicio eliminado"); setEdit(null);
                }}><Trash2 size={13} /> Eliminar</button>
              )}
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 8. Productos — control de existencias (esta app no vende).
   ========================================================================== */
function ProductsView({ db, actions }) {
  const [edit, setEdit] = useState(null);
  const [move, setMove] = useState(null);
  const low = db.products.filter((p) => p.stock <= (p.minStock || 0));
  return (
    <div>
      <div className="flex items-center justify-between">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Productos</h1>
          <p className="bt-eyebrow">Existencias de insumos</p>
        </div>
        <div className="flex gap-2">
          <button className="bt-btn ghost sm" onClick={() => toExcel(`inventario-${ymd()}.xlsx`, [
            { name: "Existencias", rows: db.products.map((p) => ({ Producto: p.name, Código: p.sku || "", Existencias: p.stock, Mínimo: p.minStock, Costo: p.cost })) },
            { name: "Movimientos", rows: db.moves.map((m) => ({ Fecha: m.date, Producto: db.products.find((p) => p.id === m.productId)?.name || "", Tipo: m.type, Cantidad: m.qty, Nota: m.note || "" })) },
          ])}><Download size={13} /> Excel</button>
          <button className="bt-btn" onClick={() => setEdit({ id: "", name: "", sku: "", stock: 0, minStock: 2, cost: 0, price: 0, unit: "unidad", description: "", showPublic: true })}><Plus size={14} /> Nuevo</button>
        </div>
      </div>

      {low.length > 0 && (
        <div className="bt-card p-3 mt-4" style={{ borderLeft: "3px solid var(--pole)", fontSize: 13 }}>
          <AlertTriangle size={13} style={{ display: "inline", color: "var(--pole)", marginRight: 5 }} />
          Toca reponer: {low.map((p) => p.name).join(", ")}
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 mt-4">
        {db.products.length === 0 && <div className="sm:col-span-2 lg:col-span-3"><Empty icon={Package} title="Sin productos registrados" /></div>}
        {db.products.map((p) => (
          <div key={p.id} className="bt-card p-4">
            <div className="flex items-start justify-between">
              <div>
                <div className="bt-display" style={{ fontSize: 14 }}>{p.name}</div>
                {p.sku && <div className="bt-mono" style={{ fontSize: 11, color: "var(--slate)" }}>{p.sku}</div>}
              </div>
              <button className="bt-icon" onClick={() => setEdit(clone(p))} aria-label="Editar"><Pencil size={13} /></button>
            </div>
            <div className="bt-mono mt-3" style={{ fontSize: 22, fontWeight: 600, color: p.stock <= (p.minStock || 0) ? "var(--pole)" : "var(--ink)" }}>
              {p.stock} <span style={{ fontSize: 12, color: "var(--slate)" }}>{p.unit}</span>
            </div>
            <div className="flex items-center gap-2 mt-1" style={{ fontSize: 12, color: "var(--slate)" }}>
              <span className="bt-mono">{p.price ? money(p.price, db.settings.currency) : "sin precio"}</span>
              {p.showPublic !== false && <span className="bt-tag conf">en catálogo</span>}
            </div>
            <div className="flex gap-2 mt-3">
              <button className="bt-btn ghost sm" onClick={() => setMove({ productId: p.id, type: "entrada", qty: 1, note: "" })}>Entrada</button>
              <button className="bt-btn ghost sm" onClick={() => setMove({ productId: p.id, type: "salida", qty: 1, note: "" })}>Salida</button>
            </div>
          </div>
        ))}
      </div>

      {edit && (
        <Modal title={edit.id ? "Editar producto" : "Nuevo producto"} onClose={() => setEdit(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Nombre"><input className="bt-input" value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} /></Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Código interno"><input className="bt-input bt-mono" value={edit.sku || ""} onChange={(e) => setEdit({ ...edit, sku: e.target.value })} /></Field>
              <Field label="Unidad"><input className="bt-input" value={edit.unit || ""} onChange={(e) => setEdit({ ...edit, unit: e.target.value })} /></Field>
            </div>
            <div className="grid grid-cols-3 gap-3">
              <Field label="Existencias"><input className="bt-input bt-mono" type="number" value={edit.stock} onChange={(e) => setEdit({ ...edit, stock: Number(e.target.value) })} /></Field>
              <Field label="Mínimo"><input className="bt-input bt-mono" type="number" value={edit.minStock} onChange={(e) => setEdit({ ...edit, minStock: Number(e.target.value) })} /></Field>
              <Field label="Costo" hint="Lo que te cuesta."><input className="bt-input bt-mono" type="number" value={edit.cost} onChange={(e) => setEdit({ ...edit, cost: Number(e.target.value) })} /></Field>
            </div>
            <Field label="Precio al público" hint="Se muestra en el catálogo. En 0 aparece como “Consultar”.">
              <input className="bt-input bt-mono" type="number" min="0" value={edit.price || 0} onChange={(e) => setEdit({ ...edit, price: Number(e.target.value) })} />
            </Field>
            <Field label="Descripción corta">
              <textarea className="bt-area" rows={2} placeholder="Cera mate, fijación fuerte, 100 g" value={edit.description || ""} onChange={(e) => setEdit({ ...edit, description: e.target.value })} />
            </Field>
            <label className="flex items-center gap-2" style={{ fontSize: 13 }}>
              <input type="checkbox" checked={edit.showPublic !== false} onChange={(e) => setEdit({ ...edit, showPublic: e.target.checked })} />
              Mostrar en el catálogo que ven los clientes
            </label>
            <div className="grid grid-cols-2 gap-2">
              <button className="bt-btn" disabled={!edit.name.trim()}
                      onClick={async () => { await actions.save("products", { ...edit, id: edit.id || uid() }, "Producto guardado"); setEdit(null); }}><Save size={13} /> Guardar</button>
              {edit.id && <button className="bt-btn danger" onClick={async () => { await actions.remove("products", edit.id, "Producto eliminado"); setEdit(null); }}><Trash2 size={13} /> Eliminar</button>}
            </div>
          </div>
        </Modal>
      )}

      {move && (
        <Modal title={move.type === "entrada" ? "Registrar entrada" : "Registrar salida"} onClose={() => setMove(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Cantidad"><input className="bt-input bt-mono" type="number" min="1" autoFocus value={move.qty} onChange={(e) => setMove({ ...move, qty: Number(e.target.value) })} /></Field>
            <Field label="Nota"><input className="bt-input" placeholder="Compra al proveedor, uso en sala…" value={move.note} onChange={(e) => setMove({ ...move, note: e.target.value })} /></Field>
            <button className="bt-btn" onClick={async () => { await actions.stockMove(move); setMove(null); }}>Registrar</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 8b. Mantenimiento — equipos y su historial de mantenimientos.
   Cada equipo (máquina, silla, secador...) tiene una lista de intervenciones:
   qué se le hizo y cuándo. Sin fechas de "próximo mantenimiento": es un
   registro histórico simple, no un planificador.
   ========================================================================== */
function MaintenanceView({ db, actions }) {
  const [editEquip, setEditEquip] = useState(null);   // {id,name,notes} — alta/edición de equipo
  const [historyFor, setHistoryFor] = useState(null); // equipo cuyo historial se está viendo
  const [newRecord, setNewRecord] = useState(null);   // {equipmentId,date,description} — nuevo registro

  const recordsOf = (equipId) =>
    db.maintenance.filter((m) => m.equipmentId === equipId).sort((a, b) => b.date.localeCompare(a.date));

  return (
    <div>
      <div className="flex items-center justify-between">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Mantenimiento</h1>
          <p className="bt-eyebrow">Equipos y su historial</p>
        </div>
        <button className="bt-btn" onClick={() => setEditEquip({ id: "", name: "", notes: "" })}>
          <Plus size={14} /> Nuevo equipo
        </button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 mt-4">
        {db.equipment.length === 0 && (
          <div className="sm:col-span-2 lg:col-span-3">
            <Empty icon={Wrench} title="Sin equipos registrados todavía"
                   action={<button className="bt-btn ghost sm" onClick={() => setEditEquip({ id: "", name: "", notes: "" })}><Plus size={13} /> Agregar el primero</button>} />
          </div>
        )}
        {db.equipment.map((eq) => {
          const recs = recordsOf(eq.id);
          const last = recs[0];
          return (
            <div key={eq.id} className="bt-card p-4">
              <div className="flex items-start justify-between">
                <div>
                  <div className="bt-display" style={{ fontSize: 14 }}>{eq.name}</div>
                  {eq.notes && <div style={{ fontSize: 12, color: "var(--slate)", marginTop: 2 }}>{eq.notes}</div>}
                </div>
                <button className="bt-icon" onClick={() => setEditEquip(clone(eq))} aria-label="Editar"><Pencil size={13} /></button>
              </div>
              <div className="bt-inset p-2 mt-3" style={{ fontSize: 12 }}>
                {last ? (
                  <>
                    <span style={{ color: "var(--slate)" }}>Último: {shortDate(last.date)}</span>
                    <div style={{ marginTop: 2 }}>{last.description}</div>
                  </>
                ) : <span style={{ color: "var(--slate)" }}>Sin mantenimientos registrados</span>}
              </div>
              <div className="flex gap-2 mt-3">
                <button className="bt-btn ghost sm" onClick={() => setHistoryFor(eq)}>
                  Ver historial {recs.length > 0 && `(${recs.length})`}
                </button>
                <button className="bt-btn ghost sm" onClick={() => setNewRecord({ equipmentId: eq.id, date: ymd(), description: "" })}>
                  <Plus size={12} /> Mantenimiento
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {editEquip && (
        <Modal title={editEquip.id ? "Editar equipo" : "Nuevo equipo"} onClose={() => setEditEquip(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Nombre" hint="Máquina de corte, secador, silla 1…">
              <input className="bt-input" value={editEquip.name} onChange={(e) => setEditEquip({ ...editEquip, name: e.target.value })} />
            </Field>
            <Field label="Notas (opcional)">
              <textarea className="bt-area" rows={2} value={editEquip.notes || ""} onChange={(e) => setEditEquip({ ...editEquip, notes: e.target.value })} />
            </Field>
            <div className="grid grid-cols-2 gap-2">
              <button className="bt-btn" disabled={!editEquip.name.trim()}
                      onClick={async () => { await actions.save("equipment", { ...editEquip, id: editEquip.id || uid() }, "Equipo guardado"); setEditEquip(null); }}>
                <Save size={13} /> Guardar
              </button>
              {editEquip.id && (
                <button className="bt-btn danger" onClick={async () => {
                  if (db.maintenance.some((m) => m.equipmentId === editEquip.id)) { actions.toast("Tiene mantenimientos registrados: bórralos primero desde el historial.", "pole"); return; }
                  await actions.remove("equipment", editEquip.id, "Equipo eliminado"); setEditEquip(null);
                }}><Trash2 size={13} /> Eliminar</button>
              )}
            </div>
          </div>
        </Modal>
      )}

      {historyFor && (
        <Modal title={`Historial · ${historyFor.name}`} onClose={() => setHistoryFor(null)}>
          <button className="bt-btn sm w-full mb-3" onClick={() => setNewRecord({ equipmentId: historyFor.id, date: ymd(), description: "" })}>
            <Plus size={13} /> Nuevo mantenimiento
          </button>
          <div className="flex flex-col gap-2" style={{ maxHeight: 360, overflow: "auto" }}>
            {recordsOf(historyFor.id).length === 0 && <p style={{ fontSize: 13, color: "var(--slate)" }}>Sin registros todavía.</p>}
            {recordsOf(historyFor.id).map((m) => (
              <div key={m.id} className="bt-ticket">
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>{shortDate(m.date)}</div>
                    <div style={{ fontSize: 14, marginTop: 2 }}>{m.description}</div>
                  </div>
                  <button className="bt-icon" style={{ width: 26, height: 26 }}
                          onClick={() => actions.remove("maintenance", m.id, "Registro eliminado")} aria-label="Borrar">
                    <Trash2 size={12} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        </Modal>
      )}

      {newRecord && (
        <Modal title="Nuevo mantenimiento" onClose={() => setNewRecord(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Fecha en que se realizó">
              <input className="bt-input bt-mono" type="date" value={newRecord.date} onChange={(e) => setNewRecord({ ...newRecord, date: e.target.value })} />
            </Field>
            <Field label="¿Qué se le hizo?">
              <textarea className="bt-area" rows={4} autoFocus placeholder="Cambio de cuchillas, limpieza, lubricación…"
                        value={newRecord.description} onChange={(e) => setNewRecord({ ...newRecord, description: e.target.value })} />
            </Field>
            <button className="bt-btn" disabled={!newRecord.description.trim()}
                    onClick={async () => {
                      await actions.save("maintenance", { ...newRecord, id: uid(), createdAt: new Date().toISOString() }, "Mantenimiento registrado");
                      setNewRecord(null);
                    }}>
              <Save size={13} /> Guardar
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 8c. Insumos a comprar — lista simple de compras pendientes.
   ========================================================================== */
function SuppliesView({ db, actions }) {
  const [newItem, setNewItem] = useState(null); // {name, dateToBuy}
  const pending = db.supplies.filter((s) => s.status !== "comprado")
    .sort((a, b) => (a.dateToBuy || "").localeCompare(b.dateToBuy || ""));
  const bought = db.supplies.filter((s) => s.status === "comprado")
    .sort((a, b) => (b.dateToBuy || "").localeCompare(a.dateToBuy || ""));
  const [showBought, setShowBought] = useState(false);

  return (
    <div>
      <div className="flex items-center justify-between">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Insumos</h1>
          <p className="bt-eyebrow">{pending.length ? `${pending.length} por comprar` : "Todo comprado"}</p>
        </div>
        <button className="bt-btn" onClick={() => setNewItem({ name: "", dateToBuy: ymd() })}>
          <Plus size={14} /> Nueva compra
        </button>
      </div>

      <div className="flex flex-col gap-2 mt-4">
        {pending.length === 0 && <Empty icon={ShoppingCart} title="No hay compras pendientes" />}
        {pending.map((s) => (
          <div key={s.id} className="bt-card p-3 flex items-center justify-between gap-3">
            <div>
              <div style={{ fontSize: 14, fontWeight: 600 }}>{s.name}</div>
              {s.dateToBuy && <div className="bt-mono" style={{ fontSize: 11, color: "var(--slate)" }}>Comprar antes del {shortDate(s.dateToBuy)}</div>}
            </div>
            <div className="flex gap-1">
              <button className="bt-btn mint sm" onClick={() => actions.save("supplies", { id: s.id, status: "comprado" }, "Marcado como comprado")}>
                <Check size={12} /> Comprado
              </button>
              <button className="bt-icon" onClick={() => actions.remove("supplies", s.id, "Eliminado")} aria-label="Borrar"><Trash2 size={13} /></button>
            </div>
          </div>
        ))}
      </div>

      {bought.length > 0 && (
        <div className="mt-5">
          <button className="bt-btn ghost sm" onClick={() => setShowBought(!showBought)}>
            {showBought ? "Ocultar" : "Ver"} comprados ({bought.length})
          </button>
          {showBought && (
            <div className="flex flex-col gap-2 mt-2">
              {bought.map((s) => (
                <div key={s.id} className="bt-card p-3 flex items-center justify-between gap-3" style={{ opacity: .6 }}>
                  <div style={{ fontSize: 14 }}>{s.name}</div>
                  <button className="bt-icon" onClick={() => actions.remove("supplies", s.id, "Eliminado")} aria-label="Borrar"><Trash2 size={13} /></button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {newItem && (
        <Modal title="Nueva compra" onClose={() => setNewItem(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Insumo" hint="Toallas, shampoo, guantes, cuchillas…">
              <input className="bt-input" autoFocus value={newItem.name} onChange={(e) => setNewItem({ ...newItem, name: e.target.value })} />
            </Field>
            <Field label="Fecha a comprar">
              <input className="bt-input bt-mono" type="date" value={newItem.dateToBuy} onChange={(e) => setNewItem({ ...newItem, dateToBuy: e.target.value })} />
            </Field>
            <button className="bt-btn" disabled={!newItem.name.trim()}
                    onClick={async () => {
                      await actions.save("supplies", { ...newItem, id: uid(), status: "pendiente", createdAt: new Date().toISOString() }, "Añadido a la lista");
                      setNewItem(null);
                    }}>
              <Save size={13} /> Guardar
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 8d. Caja (POS) — cobrar servicios y productos.
   Una venta guarda: ítems, descuento, propina, total, cómo se pagó (uno o
   varios métodos), quién cobró y a quién se le atendió. Al cobrar descuenta
   el inventario y suma al cliente; todo en UNA transacción (ver actions.sell).
   ========================================================================== */
const isCash = (m) => norm(m).includes("efectivo");

/** Texto del recibo para WhatsApp. */
function saleText(sale, settings) {
  const cur = settings.currency;
  const lines = [
    `*${settings.businessName}*`, `Recibo ${sale.number} · ${shortDate(sale.date)} ${ampm(sale.time)}`, "",
    ...sale.items.map((i) => `${i.qty} x ${i.name} — ${money(i.qty * i.unitPrice, cur)}`),
  ];
  if (sale.discount) lines.push(`Descuento: -${money(sale.discount, cur)}`);
  if (sale.tip) lines.push(`Propina: ${money(sale.tip, cur)}`);
  lines.push(`*Total: ${money(sale.total, cur)}*`);
  lines.push(`Pago: ${(sale.payments || []).map((p) => `${p.method} ${money(p.amount, cur)}`).join(" + ") || "—"}`);
  if (settings.receiptFooter) lines.push("", settings.receiptFooter);
  return lines.join("\n");
}

/** Abre el recibo en una ventana angosta (tipo tirilla) y lanza Imprimir. */
function printReceipt(sale, settings) {
  const w = window.open("", "_blank");
  if (!w) { alert("El navegador bloqueó la ventana. Permite ventanas emergentes para imprimir."); return; }
  const cur = settings.currency;
  const rows = sale.items.map((i) =>
    `<tr><td>${esc(i.name)}<br><small>${esc(i.qty)} × ${esc(money(i.unitPrice, cur))}</small></td><td class="r">${esc(money(i.qty * i.unitPrice, cur))}</td></tr>`).join("");
  const pays = (sale.payments || []).map((p) => `<tr><td>${esc(p.method)}</td><td class="r">${esc(money(p.amount, cur))}</td></tr>`).join("");
  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(sale.number)}</title>
  <style>
    @page{ margin:4mm; size:80mm auto; }
    body{ font-family:Arial,Helvetica,sans-serif; font-size:12px; color:#000; width:72mm; margin:0 auto; }
    h1{ font-size:15px; text-align:center; margin:6px 0 2px; }
    .c{ text-align:center; } .r{ text-align:right; white-space:nowrap; }
    table{ width:100%; border-collapse:collapse; } td{ padding:2px 0; vertical-align:top; }
    hr{ border:0; border-top:1px dashed #000; margin:6px 0; }
    .tot td{ font-weight:bold; font-size:14px; }
  </style></head><body>
  <div class="c"><img src="${currentLogo()}" style="height:54px"/></div>
  <h1>${esc(settings.businessName)}</h1>
  <div class="c">Recibo ${esc(sale.number)}<br>${esc(shortDate(sale.date))} ${esc(ampm(sale.time))}</div>
  ${sale.clientName ? `<div class="c">Cliente: ${esc(sale.clientName)}</div>` : ""}
  <hr><table>${rows}</table><hr>
  <table>
    <tr><td>Subtotal</td><td class="r">${esc(money(sale.subtotal, cur))}</td></tr>
    ${sale.discount ? `<tr><td>Descuento</td><td class="r">-${esc(money(sale.discount, cur))}</td></tr>` : ""}
    ${sale.tip ? `<tr><td>Propina</td><td class="r">${esc(money(sale.tip, cur))}</td></tr>` : ""}
    <tr class="tot"><td>TOTAL</td><td class="r">${esc(money(sale.total, cur))}</td></tr>
  </table><hr>
  <table>${pays}${sale.change ? `<tr><td>Cambio</td><td class="r">${esc(money(sale.change, cur))}</td></tr>` : ""}</table>
  <p class="c">${esc(settings.receiptFooter || "")}</p>
  </body></html>`);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 350);
}

function PosView({ db, actions, user, preset, clearPreset }) {
  const cur = db.settings.currency;
  const services = db.services.filter((s) => s.active !== false);
  const barbers = db.barbers.filter((b) => b.active !== false);
  const mine = user.rol === "barbero" ? barbers.find((b) => norm(b.name) === norm(user.name)) : null;

  const [tab, setTab] = useState("servicios");
  const [q, setQ] = useState("");
  const [cart, setCart] = useState([]);            // { key, type, refId, name, qty, unitPrice }
  const [clientId, setClientId] = useState("");
  const [barberId, setBarberId] = useState(mine ? mine.id : "");
  const [apptId, setApptId] = useState("");
  const [apptMeta, setApptMeta] = useState({ counted: false, price: 0 });
  const [discount, setDiscount] = useState("");
  const [tip, setTip] = useState("");
  const [paying, setPaying] = useState(false);
  const [receipt, setReceipt] = useState(null);

  // Llegó desde una cita ("Cobrar"): precarga servicio, cliente y barbero.
  useEffect(() => {
    if (!preset) return;
    setCart(preset.items.map((i) => ({ ...i, key: uid() })));
    setClientId(preset.clientId || "");
    setBarberId(preset.barberId || "");
    setApptId(preset.apptId || "");
    setApptMeta({ counted: !!preset.apptCounted, price: preset.apptPrice || 0 });
    clearPreset();
  }, [preset]); // eslint-disable-line react-hooks/exhaustive-deps

  const reset = () => {
    setCart([]); setClientId(""); setBarberId(mine ? mine.id : ""); setApptId("");
    setApptMeta({ counted: false, price: 0 }); setDiscount(""); setTip("");
  };

  const n = norm(q);
  const list = (tab === "servicios" ? services : db.products).filter((x) => !n || norm(x.name).includes(n));

  const add = (type, item) => {
    const line = cart.find((x) => x.type === type && x.refId === item.id);
    if (type === "product" && (line ? line.qty : 0) >= (item.stock || 0)) {
      actions.toast(`Solo quedan ${item.stock || 0} de "${item.name}".`, "pole"); return;
    }
    if (line) setCart(cart.map((x) => (x === line ? { ...x, qty: x.qty + 1 } : x)));
    else setCart([...cart, { key: uid(), type, refId: item.id, name: item.name, qty: 1, unitPrice: item.price || 0 }]);
  };
  const setQty = (key, delta) => {
    const l = cart.find((x) => x.key === key);
    if (!l) return;
    const qn = l.qty + delta;
    if (qn <= 0) { setCart(cart.filter((x) => x.key !== key)); return; }
    if (l.type === "product") {
      const pr = db.products.find((x) => x.id === l.refId);
      if (pr && qn > (pr.stock || 0)) { actions.toast(`Solo quedan ${pr.stock || 0} de "${l.name}".`, "pole"); return; }
    }
    setCart(cart.map((x) => (x.key === key ? { ...x, qty: qn } : x)));
  };
  const setPrice = (key, v) => setCart(cart.map((x) => (x.key === key ? { ...x, unitPrice: Math.max(Number(v) || 0, 0) } : x)));

  const subtotal = sum(cart, (l) => l.qty * l.unitPrice);
  const disc = Math.min(Math.max(Number(discount) || 0, 0), subtotal);
  const tipN = Math.max(Number(tip) || 0, 0);
  const total = subtotal - disc + tipN;

  return (
    <div>
      <div className="flex items-center justify-between gap-2">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Caja</h1>
          <p className="bt-eyebrow">Cobrar servicios y productos</p>
        </div>
        {cart.length > 0 && <button className="bt-btn ghost sm" onClick={reset}><X size={13} /> Vaciar</button>}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3 mt-4">
        {/* ----- qué se vende ----- */}
        <div className="bt-card p-3">
          <div className="flex gap-2">
            {["servicios", "productos"].map((t) => (
              <button key={t} className={`bt-slot ${tab === t ? "on" : ""}`} style={{ flex: 1, fontFamily: "var(--body)", textTransform: "capitalize" }}
                      onClick={() => setTab(t)}>{t}</button>
            ))}
          </div>
          <div className="relative mt-3">
            <Search size={14} style={{ position: "absolute", left: 10, top: 12, color: "var(--slate)" }} />
            <input className="bt-input" style={{ paddingLeft: 30 }} placeholder="Buscar" value={q} onChange={(e) => setQ(e.target.value)} />
          </div>
          <div className="grid grid-cols-2 gap-2 mt-3 bt-scroll" style={{ maxHeight: 380 }}>
            {list.length === 0 && <p style={{ fontSize: 12, color: "var(--slate)" }}>Nada para mostrar.</p>}
            {list.map((it) => {
              const isP = tab === "productos";
              const out = isP && (it.stock || 0) <= 0;
              return (
                <button key={it.id} className="bt-ticket text-left" disabled={out}
                        style={{ cursor: out ? "not-allowed" : "pointer", opacity: out ? 0.45 : 1 }}
                        onClick={() => add(isP ? "product" : "service", it)}>
                  <div style={{ fontSize: 13, fontWeight: 600 }}>{it.name}</div>
                  <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>
                    {it.price ? money(it.price, cur) : "Precio libre"}{isP && ` · ${it.stock || 0} disp.`}
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        {/* ----- la cuenta ----- */}
        <div className="bt-card p-3">
          <span className="bt-eyebrow">Cuenta</span>
          {cart.length === 0 ? (
            <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 8 }}>Toca un servicio o producto para agregarlo.</p>
          ) : (
            <div className="flex flex-col gap-2 mt-2">
              {cart.map((l) => (
                <div key={l.key} className="bt-inset p-2">
                  <div className="flex items-center justify-between gap-2">
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{l.name}</div>
                    <div className="flex items-center gap-1">
                      <button className="bt-icon" style={{ width: 26, height: 26 }} onClick={() => setQty(l.key, -1)} aria-label="Menos"><Minus size={12} /></button>
                      <span className="bt-mono" style={{ minWidth: 22, textAlign: "center" }}>{l.qty}</span>
                      <button className="bt-icon" style={{ width: 26, height: 26 }} onClick={() => setQty(l.key, 1)} aria-label="Más"><Plus size={12} /></button>
                    </div>
                  </div>
                  <div className="flex items-center justify-between gap-2 mt-1">
                    <input className="bt-input bt-mono" type="number" min="0" style={{ width: 120, padding: "5px 8px" }}
                           value={l.unitPrice} onChange={(e) => setPrice(l.key, e.target.value)} aria-label="Precio" />
                    <span className="bt-mono" style={{ fontWeight: 600 }}>{money(l.qty * l.unitPrice, cur)}</span>
                  </div>
                </div>
              ))}
            </div>
          )}

          <div className="flex flex-col gap-3 mt-3">
            <Field label="Cliente (opcional)">
              <ClientPicker db={db} value={clientId} allowCreate={false} onChange={(v) => setClientId(typeof v === "string" ? v : "")} />
            </Field>
            <Field label="Atendió">
              <select className="bt-select" value={barberId} onChange={(e) => setBarberId(e.target.value)}>
                <option value="">— Sin asignar —</option>
                {barbers.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
              </select>
            </Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Descuento"><input className="bt-input bt-mono" type="number" min="0" value={discount} onChange={(e) => setDiscount(e.target.value)} /></Field>
              <Field label="Propina"><input className="bt-input bt-mono" type="number" min="0" value={tip} onChange={(e) => setTip(e.target.value)} /></Field>
            </div>
          </div>

          <div className="bt-inset p-3 mt-3" style={{ fontSize: 13 }}>
            <div className="flex justify-between"><span>Subtotal</span><span className="bt-mono">{money(subtotal, cur)}</span></div>
            {disc > 0 && <div className="flex justify-between"><span>Descuento</span><span className="bt-mono">-{money(disc, cur)}</span></div>}
            {tipN > 0 && <div className="flex justify-between"><span>Propina</span><span className="bt-mono">{money(tipN, cur)}</span></div>}
            <div className="flex justify-between mt-1" style={{ fontSize: 18, fontWeight: 600 }}><span>Total</span><span className="bt-mono">{money(total, cur)}</span></div>
          </div>
          <button className="bt-btn brass w-full mt-3" style={{ padding: 13 }} disabled={cart.length === 0} onClick={() => setPaying(true)}>
            <CreditCard size={15} /> Cobrar {money(total, cur)}
          </button>
        </div>
      </div>

      {paying && (
        <CheckoutModal db={db} total={total} onClose={() => setPaying(false)}
          onConfirm={async (payments, received, change) => {
            const sale = await actions.sell({
              items: cart, clientId, barberId, apptId, apptCounted: apptMeta.counted, apptPrice: apptMeta.price,
              discount: disc, tip: tipN, payments, received, change, staffName: user.name,
            });
            if (sale) { setPaying(false); setReceipt(sale); }
            return !!sale;
          }} />
      )}
      {receipt && <ReceiptModal db={db} sale={receipt} onClose={() => { setReceipt(null); reset(); }} />}
    </div>
  );
}

/** Paso de pago: un método o varios (pago dividido), con vuelto en efectivo. */
function CheckoutModal({ db, total, onClose, onConfirm }) {
  const cur = db.settings.currency;
  const methods = db.settings.payMethods;
  const [rows, setRows] = useState([{ method: methods[0], amount: total }]);
  const [received, setReceived] = useState("");
  const [busy, setBusy] = useState(false);

  const paid = sum(rows, (r) => r.amount);
  const diff = total - paid;
  const cashPaid = sum(rows.filter((r) => isCash(r.method)), (r) => r.amount);
  const recv = Number(received) || 0;
  const change = cashPaid > 0 && recv > cashPaid ? recv - cashPaid : 0;
  const ok = Math.abs(diff) < 0.5 && rows.every((r) => r.amount >= 0);
  const setRow = (i, patch) => setRows(rows.map((r, k) => (k === i ? { ...r, ...patch } : r)));

  return (
    <Modal title="Cobrar" onClose={onClose}>
      <div className="bt-inset p-3 text-center">
        <div className="bt-eyebrow">Total a pagar</div>
        <div className="bt-mono" style={{ fontSize: 28, fontWeight: 600 }}>{money(total, cur)}</div>
      </div>

      <span className="bt-label mt-4" style={{ display: "block" }}>Método de pago</span>
      <div className="grid grid-cols-3 gap-2">
        {methods.map((m) => (
          <button key={m} className={`bt-slot ${rows.length === 1 && rows[0].method === m ? "on" : ""}`}
                  style={{ fontFamily: "var(--body)", fontSize: 12 }} onClick={() => setRows([{ method: m, amount: total }])}>{m}</button>
        ))}
      </div>

      {rows.length > 1 && (
        <div className="mt-3">
          {rows.map((r, i) => (
            <div key={i} className="flex items-center gap-2 mt-2">
              <select className="bt-select" value={r.method} onChange={(e) => setRow(i, { method: e.target.value })}>
                {methods.map((m) => <option key={m}>{m}</option>)}
              </select>
              <input className="bt-input bt-mono" type="number" min="0" style={{ width: 120 }} value={r.amount}
                     onChange={(e) => setRow(i, { amount: Math.max(Number(e.target.value) || 0, 0) })} />
              <button className="bt-icon" onClick={() => setRows(rows.filter((_, k) => k !== i))} aria-label="Quitar"><X size={13} /></button>
            </div>
          ))}
        </div>
      )}
      <button className="bt-btn ghost sm mt-3" onClick={() => setRows([...rows, { method: methods[0], amount: Math.max(diff, 0) }])}>
        <Plus size={12} /> Dividir el pago
      </button>

      {cashPaid > 0 && (
        <div className="mt-3">
          <Field label="Efectivo recibido (para calcular el vuelto)">
            <input className="bt-input bt-mono" type="number" min="0" value={received} onChange={(e) => setReceived(e.target.value)} />
          </Field>
          {recv > 0 && recv < cashPaid && <p style={{ fontSize: 12, color: "var(--pole)", marginTop: 4 }}>Falta {money(cashPaid - recv, cur)} en efectivo.</p>}
          {change > 0 && <p className="bt-mono" style={{ fontSize: 16, fontWeight: 600, marginTop: 6 }}>Vuelto: {money(change, cur)}</p>}
        </div>
      )}

      {Math.abs(diff) >= 0.5 && (
        <p style={{ fontSize: 12, color: "var(--pole)", marginTop: 10 }}>
          {diff > 0 ? `Faltan ${money(diff, cur)} por asignar.` : `Sobran ${money(-diff, cur)}: los pagos pasan del total.`}
        </p>
      )}
      <button className="bt-btn brass w-full mt-4" style={{ padding: 13 }} disabled={!ok || busy}
              onClick={async () => {
                setBusy(true);
                const done = await onConfirm(rows.filter((r) => r.amount > 0), recv, change);
                if (!done) setBusy(false);
              }}>
        <Check size={15} /> {busy ? "Registrando…" : "Confirmar pago"}
      </button>
    </Modal>
  );
}

/** Recibo de una venta: ver, imprimir, mandar por WhatsApp y (admin) anular. */
function ReceiptModal({ db, sale, onClose, onVoid }) {
  const cur = db.settings.currency;
  const client = sale.clientId ? db.clients.find((c) => c.id === sale.clientId) : null;
  const wa = client ? waLink(client.phone, saleText(sale, db.settings)) : null;
  return (
    <Modal title={`Recibo ${sale.number}`} onClose={onClose}>
      <div className="bt-inset p-3" style={{ fontSize: 13 }}>
        <div className="bt-mono" style={{ color: "var(--slate)" }}>{shortDate(sale.date)} · {ampm(sale.time)}{sale.staffName ? ` · ${sale.staffName}` : ""}</div>
        {sale.clientName && <div style={{ fontWeight: 600, marginTop: 4 }}>{sale.clientName}</div>}
        <div className="mt-2">
          {sale.items.map((i, k) => (
            <div key={k} className="flex justify-between gap-2"><span>{i.qty} × {i.name}</span><span className="bt-mono">{money(i.qty * i.unitPrice, cur)}</span></div>
          ))}
        </div>
        {sale.discount > 0 && <div className="flex justify-between mt-1"><span>Descuento</span><span className="bt-mono">-{money(sale.discount, cur)}</span></div>}
        {sale.tip > 0 && <div className="flex justify-between"><span>Propina</span><span className="bt-mono">{money(sale.tip, cur)}</span></div>}
        <div className="flex justify-between mt-2" style={{ fontSize: 17, fontWeight: 600 }}><span>Total</span><span className="bt-mono">{money(sale.total, cur)}</span></div>
        <div className="mt-2" style={{ color: "var(--slate)" }}>
          {(sale.payments || []).map((p, k) => <div key={k} className="flex justify-between"><span>{p.method}</span><span className="bt-mono">{money(p.amount, cur)}</span></div>)}
          {sale.change > 0 && <div className="flex justify-between"><span>Vuelto</span><span className="bt-mono">{money(sale.change, cur)}</span></div>}
        </div>
        {sale.status === "anulada" && <div className="mt-2"><span className="bt-tag canc">anulada</span>{sale.voidReason ? ` ${sale.voidReason}` : ""}</div>}
      </div>
      <div className="grid grid-cols-2 gap-2 mt-3">
        <button className="bt-btn ghost" onClick={() => printReceipt(sale, db.settings)}><Printer size={13} /> Imprimir</button>
        <a className="bt-btn mint" target="_blank" rel="noreferrer" href={wa || undefined}
           style={{ opacity: wa ? 1 : 0.4, pointerEvents: wa ? "auto" : "none" }}><MessageCircle size={13} /> WhatsApp</a>
      </div>
      {onVoid && sale.status !== "anulada" && (
        <button className="bt-btn danger w-full mt-3" onClick={() => onVoid(sale)}><Ban size={13} /> Anular venta</button>
      )}
      <button className="bt-btn ghost w-full mt-2" onClick={onClose}>Cerrar</button>
    </Modal>
  );
}

/* ==========================================================================
   [VISTAS] 8e. Ventas — historial del día y cierre de caja.
   ========================================================================== */
function SalesView({ db, actions, user }) {
  const cur = db.settings.currency;
  const [date, setDate] = useState(ymd());
  const [sales, setSales] = useState([]);
  const [err, setErr] = useState("");
  const [open, setOpen] = useState(null);

  useEffect(() => {
    const q = query(fcol("sales"), where("date", "==", date));
    return onSnapshot(q, (snap) => {
      setSales(snap.docs.map((d) => ({ ...d.data(), id: d.id })).sort((a, b) => (a.time || "").localeCompare(b.time || "")));
      setErr("");
    }, (e) => setErr(e.code || e.message));
  }, [date]);

  const valid = sales.filter((v) => v.status !== "anulada");
  const byMethod = {};
  valid.forEach((v) => (v.payments || []).forEach((p) => { byMethod[p.method] = (byMethod[p.method] || 0) + (Number(p.amount) || 0); }));
  const total = sum(valid, (v) => v.total);
  const tips = sum(valid, (v) => v.tip);
  const cash = Object.entries(byMethod).filter(([m]) => isCash(m)).reduce((a, [, v]) => a + v, 0);

  const closeCash = () => toPDF("Cierre de caja", `${db.settings.businessName} · ${niceDate(date)}`, [
    { title: "Resumen", rows: [{ Ventas: valid.length, Total: money(total, cur), Propinas: money(tips, cur), "Efectivo en caja": money(cash, cur), Anuladas: sales.length - valid.length }] },
    { title: "Por método de pago", rows: Object.entries(byMethod).map(([m, v]) => ({ Método: m, Total: money(v, cur) })) },
    { title: "Ventas del día", rows: valid.map((v) => ({ "N°": v.number, Hora: v.time, Cliente: v.clientName || "—", Total: money(v.total, cur), Pago: (v.payments || []).map((p) => p.method).join(" + ") })) },
  ]);

  const voidIt = async (sale) => {
    const reason = window.prompt("Motivo de la anulación (opcional):", "");
    if (reason === null) return;
    const res = await actions.voidSale(sale, reason);
    if (res) setOpen(null);
  };

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Ventas</h1>
          <p className="bt-eyebrow">{niceDate(date)}</p>
        </div>
        <div className="flex items-center gap-2">
          <button className="bt-icon" onClick={() => setDate(addDays(date, -1))} aria-label="Día anterior"><ChevronLeft size={16} /></button>
          <input className="bt-input bt-mono" type="date" value={date} onChange={(e) => e.target.value && setDate(e.target.value)} style={{ width: 148 }} />
          <button className="bt-icon" onClick={() => setDate(addDays(date, 1))} aria-label="Día siguiente"><ChevronRight size={16} /></button>
          <button className="bt-btn ghost sm" onClick={() => setDate(ymd())}>Hoy</button>
        </div>
      </div>
      {err && <p style={{ color: "var(--pole)", fontSize: 12, marginTop: 8 }}>No se pudieron cargar las ventas ({err}).</p>}

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mt-4">
        <Stat label="Vendido" value={money(total, cur)} sub={`${valid.length} venta${valid.length === 1 ? "" : "s"}`} tone="mint" />
        <Stat label="Efectivo en caja" value={money(cash, cur)} sub="según lo cobrado en efectivo" />
        <Stat label="Propinas" value={money(tips, cur)} />
        <Stat label="Anuladas" value={String(sales.length - valid.length)} />
      </div>

      {Object.keys(byMethod).length > 0 && (
        <div className="bt-card p-4 mt-3">
          <span className="bt-eyebrow">Por método de pago</span>
          {Object.entries(byMethod).map(([m, v]) => (
            <div key={m} className="flex items-center justify-between mt-2" style={{ fontSize: 13 }}>
              <span>{m}</span><span className="bt-mono" style={{ fontWeight: 600 }}>{money(v, cur)}</span>
            </div>
          ))}
          <button className="bt-btn ghost sm mt-3" onClick={closeCash}><Printer size={13} /> Cierre de caja (PDF)</button>
        </div>
      )}

      <div className="flex flex-col gap-2 mt-4">
        {sales.length === 0 && <Empty icon={Receipt} title="Sin ventas este día" />}
        {sales.map((v) => (
          <button key={v.id} className="bt-card p-3 text-left w-full" style={{ cursor: "pointer", opacity: v.status === "anulada" ? 0.5 : 1 }} onClick={() => setOpen(v)}>
            <div className="flex items-center justify-between gap-2">
              <div>
                <div className="bt-mono" style={{ fontSize: 13, fontWeight: 600 }}>{v.number} · {ampm(v.time)}</div>
                <div style={{ fontSize: 12, color: "var(--slate)" }}>
                  {v.clientName || "Sin cliente"} · {(v.payments || []).map((p) => p.method).join(" + ") || "—"}
                </div>
              </div>
              <div className="text-right">
                <div className="bt-mono" style={{ fontSize: 15, fontWeight: 600 }}>{money(v.total, cur)}</div>
                {v.status === "anulada" && <span className="bt-tag canc">anulada</span>}
              </div>
            </div>
          </button>
        ))}
      </div>

      {open && <ReceiptModal db={db} sale={open} onClose={() => setOpen(null)} onVoid={user.rol === "admin" ? voidIt : null} />}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 9. Gastos — arriendo, servicios, compras, nómina, otros.
   ========================================================================== */
const CATEGORIES = ["Arriendo", "Servicios públicos", "Compras / insumos", "Nómina", "Publicidad", "Mantenimiento", "Impuestos", "Otros"];

function ExpensesView({ db, actions }) {
  const [edit, setEdit] = useState(null);
  const [from, setFrom] = useState(ymd().slice(0, 8) + "01");
  const [to, setTo] = useState(ymd());
  const list = expensesBetween(db, from, to).sort((a, b) => b.date.localeCompare(a.date));
  const total = sum(list, (e) => e.amount);
  const byCat = CATEGORIES.map((c) => ({ name: c, total: sum(list.filter((e) => e.category === c), (e) => e.amount) })).filter((x) => x.total > 0);

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Gastos</h1>
          <p className="bt-eyebrow">{money(total, db.settings.currency)} en el periodo</p>
        </div>
        <div className="flex gap-2">
          <button className="bt-btn ghost sm" onClick={() => toExcel(`gastos-${from}_${to}.xlsx`, [{ name: "Gastos", rows: list.map((e) => ({ Fecha: e.date, Categoría: e.category, Descripción: e.description, Monto: e.amount, Pago: e.method || "" })) }])}>
            <Download size={13} /> Excel
          </button>
          <button className="bt-btn" onClick={() => setEdit({ id: "", date: ymd(), category: CATEGORIES[0], description: "", amount: 0, method: "Efectivo" })}><Plus size={14} /> Registrar</button>
        </div>
      </div>

      <div className="flex flex-wrap items-end gap-2 mt-4">
        <Field label="Desde"><input className="bt-input bt-mono" type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></Field>
        <Field label="Hasta"><input className="bt-input bt-mono" type="date" value={to} onChange={(e) => setTo(e.target.value)} /></Field>
      </div>

      {byCat.length > 0 && (
        <div className="bt-card p-4 mt-4">
          <span className="bt-eyebrow">Por categoría</span>
          <div style={{ height: 200, marginTop: 10 }}>
            <ResponsiveContainer>
              <BarChart data={byCat} margin={{ left: -10, right: 8, top: 6 }}>
                <CartesianGrid vertical={false} stroke={BRANDING.theme.line} />
                <XAxis dataKey="name" tick={{ fontSize: 10 }} interval={0} angle={-18} textAnchor="end" height={52} />
                <YAxis tick={{ fontSize: 10 }} />
                <Tooltip formatter={(v) => money(v, db.settings.currency)} />
                <Bar dataKey="total" fill={BRANDING.theme.pole} radius={[2, 2, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      <div className="flex flex-col gap-2 mt-4">
        {list.length === 0 && <Empty icon={Wallet} title="Sin gastos en este periodo" />}
        {list.map((e) => (
          <div key={e.id} className="bt-card p-3 flex items-center justify-between gap-3">
            <div>
              <div style={{ fontSize: 14, fontWeight: 600 }}>{e.description || e.category}</div>
              <div style={{ fontSize: 11, color: "var(--slate)" }}>{shortDate(e.date)} · {e.category} · {e.method}</div>
            </div>
            <div className="flex items-center gap-2">
              <span className="bt-mono" style={{ fontSize: 15, fontWeight: 600, color: "var(--pole)" }}>{money(e.amount, db.settings.currency)}</span>
              <button className="bt-icon" onClick={() => setEdit(clone(e))} aria-label="Editar"><Pencil size={13} /></button>
            </div>
          </div>
        ))}
      </div>

      {edit && (
        <Modal title={edit.id ? "Editar gasto" : "Registrar gasto"} onClose={() => setEdit(null)}>
          <div className="flex flex-col gap-3">
            <div className="grid grid-cols-2 gap-3">
              <Field label="Fecha"><input className="bt-input bt-mono" type="date" value={edit.date} onChange={(e) => setEdit({ ...edit, date: e.target.value })} /></Field>
              <Field label="Monto"><input className="bt-input bt-mono" type="number" min="0" value={edit.amount} onChange={(e) => setEdit({ ...edit, amount: Number(e.target.value) })} /></Field>
            </div>
            <Field label="Categoría">
              <select className="bt-select" value={edit.category} onChange={(e) => setEdit({ ...edit, category: e.target.value })}>
                {CATEGORIES.map((c) => <option key={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="Descripción"><input className="bt-input" value={edit.description} onChange={(e) => setEdit({ ...edit, description: e.target.value })} /></Field>
            <Field label="Forma de pago">
              <select className="bt-select" value={edit.method} onChange={(e) => setEdit({ ...edit, method: e.target.value })}>
                {["Efectivo", "Transferencia", "Tarjeta", "Otro"].map((m) => <option key={m}>{m}</option>)}
              </select>
            </Field>
            <div className="grid grid-cols-2 gap-2">
              <button className="bt-btn" disabled={!edit.amount}
                      onClick={async () => { await actions.save("expenses", { ...edit, id: edit.id || uid() }, "Gasto registrado"); setEdit(null); }}><Save size={13} /> Guardar</button>
              {edit.id && <button className="bt-btn danger" onClick={async () => { await actions.remove("expenses", edit.id, "Gasto eliminado"); setEdit(null); }}><Trash2 size={13} /> Eliminar</button>}
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 10. Reportes — ingresos vs. gastos, barberos, servicios.
   Ingresos = citas marcadas como atendidas (precio congelado en la cita).
   ========================================================================== */
function ReportsView({ db }) {
  const cur = db.settings.currency;
  const [from, setFrom] = useState(ymd().slice(0, 8) + "01");
  const [to, setTo] = useState(ymd());
  const [data, setData] = useState({ sales: [], appts: [], loading: true, err: "" });
  const [tick, setTick] = useState(0);

  // Ventas y citas del rango se piden solo al abrir/cambiar fechas (no se cargan todas siempre).
  useEffect(() => {
    let dead = false;
    setData((d) => ({ ...d, loading: true, err: "" }));
    (async () => {
      try {
        const [sv, ap] = await Promise.all([
          getDocs(query(fcol("sales"), where("date", ">=", from), where("date", "<=", to))),
          getDocs(query(fcol("appointments"), where("date", ">=", from), where("date", "<=", to))),
        ]);
        if (dead) return;
        setData({
          sales: sv.docs.map((d) => ({ ...d.data(), id: d.id })),
          appts: ap.docs.map((d) => ({ ...d.data(), id: d.id })),
          loading: false, err: "",
        });
      } catch (e) {
        console.error(e);
        if (!dead) setData({ sales: [], appts: [], loading: false, err: e.code || e.message || "error" });
      }
    })();
    return () => { dead = true; };
  }, [from, to, tick]);

  const sales = data.sales.filter((v) => v.status !== "anulada");
  // Citas atendidas SIN cobro en caja (de antes de la caja, o marcadas "sin cobrar"): cuentan por su precio.
  const legacy = data.appts.filter((a) => a.status === "completada" && !a.saleId);
  const net = (v) => (v.subtotal || 0) - (v.discount || 0);
  const income = sum(sales, net) + sum(legacy, (a) => a.price);
  const tips = sum(sales, (v) => v.tip);
  const exps = expensesBetween(db, from, to);
  const spent = sum(exps, (e) => e.amount);
  const balance = income - spent;
  const cancel = data.appts.filter((a) => a.status === "cancelada").length;
  const count = sales.length + legacy.length;
  const ticket = count ? income / count : 0;

  /** Serie diaria de ingresos vs. gastos para la gráfica de línea. */
  const daily = useMemo(() => {
    const map = {};
    for (let d = from; d <= to; d = addDays(d, 1)) map[d] = { day: shortDate(d), Ingresos: 0, Gastos: 0 };
    sales.forEach((v) => { if (map[v.date]) map[v.date].Ingresos += net(v); });
    legacy.forEach((a) => { if (map[a.date]) map[a.date].Ingresos += a.price || 0; });
    exps.forEach((e) => { if (map[e.date]) map[e.date].Gastos += e.amount; });
    return Object.values(map);
  }, [from, to, data, db.expenses]); // eslint-disable-line react-hooks/exhaustive-deps

  const barberRows = [...db.barbers.map((b) => ({ id: b.id, name: b.name })), { id: "", name: "Sin asignar" }].map((b) => {
    const rs = sales.filter((v) => (v.barberId || "") === b.id);
    const rl = legacy.filter((a) => (a.barberId || "") === b.id);
    return { name: b.name, ventas: rs.length + rl.length, total: sum(rs, net) + sum(rl, (a) => a.price) };
  }).filter((x) => x.ventas > 0).sort((a, b) => b.total - a.total);

  const itemMap = {};
  sales.forEach((v) => (v.items || []).forEach((i) => {
    const k = `${i.type}|${i.name}`;
    itemMap[k] = itemMap[k] || { name: i.name, type: i.type, qty: 0, total: 0 };
    itemMap[k].qty += i.qty; itemMap[k].total += i.qty * i.unitPrice;
  }));
  legacy.forEach((a) => {
    const sv = db.services.find((x) => x.id === a.serviceId);
    const k = `service|${sv ? sv.name : "Servicio"}`;
    itemMap[k] = itemMap[k] || { name: sv ? sv.name : "Servicio", type: "service", qty: 0, total: 0 };
    itemMap[k].qty += 1; itemMap[k].total += a.price || 0;
  });
  const items = Object.values(itemMap).sort((a, b) => b.total - a.total);

  const methodMap = {};
  sales.forEach((v) => (v.payments || []).forEach((p) => { methodMap[p.method] = (methodMap[p.method] || 0) + (Number(p.amount) || 0); }));

  const period = `${shortDate(from)} — ${shortDate(to)}`;
  const typeName = (t) => (t === "product" ? "Producto" : "Servicio");

  const excel = () => toExcel(`reporte-${from}_${to}.xlsx`, [
    { name: "Resumen", rows: [{ Desde: from, Hasta: to, Ingresos: income, Propinas: tips, Gastos: spent, Balance: balance, Ventas: count, "Ticket promedio": Math.round(ticket), Canceladas: cancel }] },
    { name: "Ventas", rows: sales.map((v) => ({ "N°": v.number, Fecha: v.date, Hora: v.time, Cliente: v.clientName || "", Barbero: (db.barbers.find((b) => b.id === v.barberId) || {}).name || "", Subtotal: v.subtotal, Descuento: v.discount, Propina: v.tip, Total: v.total, Pago: (v.payments || []).map((p) => `${p.method} ${p.amount}`).join(" + "), Cobró: v.staffName || "" })) },
    { name: "Por ítem", rows: items.map((i) => ({ Tipo: typeName(i.type), Ítem: i.name, Cantidad: i.qty, Total: i.total })) },
    { name: "Por método de pago", rows: Object.entries(methodMap).map(([m, v]) => ({ Método: m, Total: v })) },
    { name: "Por barbero", rows: barberRows.map((b) => ({ Barbero: b.name, Ventas: b.ventas, Total: b.total })) },
    { name: "Gastos", rows: exps.map((e) => ({ Fecha: e.date, Categoría: e.category, Descripción: e.description, Monto: e.amount })) },
    { name: "Citas", rows: data.appts.map((a) => { const h = hydrate(db, a); return { Fecha: a.date, Hora: a.time, Cliente: h.clientName, Teléfono: h.clientPhone, Barbero: h.barberName, Servicio: h.serviceName, Valor: a.price, Estado: a.status, Origen: a.source }; }) },
  ]);

  const pdf = () => toPDF("Reporte de operación", `${db.settings.businessName} · ${period}`, [
    { title: "Resumen", rows: [{ Ingresos: money(income, cur), Propinas: money(tips, cur), Gastos: money(spent, cur), Balance: money(balance, cur), Ventas: count, "Ticket promedio": money(Math.round(ticket), cur), Canceladas: cancel }] },
    { title: "Por método de pago", rows: Object.entries(methodMap).map(([m, v]) => ({ Método: m, Total: money(v, cur) })) },
    { title: "Por barbero", rows: barberRows.map((b) => ({ Barbero: b.name, Ventas: b.ventas, Total: money(b.total, cur) })) },
    { title: "Servicios y productos", rows: items.map((i) => ({ Tipo: typeName(i.type), Ítem: i.name, Cantidad: i.qty, Total: money(i.total, cur) })) },
    { title: "Gastos del periodo", rows: exps.map((e) => ({ Fecha: e.date, Categoría: e.category, Descripción: e.description, Monto: money(e.amount, cur) })) },
  ]);

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-2 bt-noprint">
        <div>
          <h1 className="bt-display" style={{ fontSize: 20 }}>Reportes</h1>
          <p className="bt-eyebrow">{period}</p>
        </div>
        <div className="flex gap-2">
          <button className="bt-btn ghost sm" onClick={() => setTick(tick + 1)} aria-label="Actualizar"><RefreshCw size={13} /></button>
          <button className="bt-btn ghost sm" onClick={excel} disabled={data.loading}><Download size={13} /> Excel</button>
          <button className="bt-btn ghost sm" onClick={pdf} disabled={data.loading}><Printer size={13} /> PDF</button>
        </div>
      </div>

      <div className="flex flex-wrap items-end gap-2 mt-4 bt-noprint">
        <Field label="Desde"><input className="bt-input bt-mono" type="date" value={from} onChange={(e) => e.target.value && setFrom(e.target.value)} /></Field>
        <Field label="Hasta"><input className="bt-input bt-mono" type="date" value={to} onChange={(e) => e.target.value && setTo(e.target.value)} /></Field>
        <button className="bt-btn ghost sm" onClick={() => { setFrom(ymd().slice(0, 8) + "01"); setTo(ymd()); }}>Este mes</button>
        <button className="bt-btn ghost sm" onClick={() => { setFrom(addDays(ymd(), -6)); setTo(ymd()); }}>7 días</button>
      </div>

      {data.err && <p style={{ color: "var(--pole)", fontSize: 12, marginTop: 10 }}>No se pudieron cargar los datos ({data.err}).</p>}
      {data.loading && <p className="bt-eyebrow" style={{ marginTop: 10 }}>Cargando…</p>}

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mt-4">
        <Stat label="Ingresos" value={money(income, cur)} sub={`${count} venta${count === 1 ? "" : "s"} · sin propinas`} tone="mint" />
        <Stat label="Gastos" value={money(spent, cur)} sub={`${exps.length} registros`} tone="pole" />
        <Stat label="Balance" value={money(balance, cur)} sub={balance >= 0 ? "a favor" : "en rojo"} tone={balance >= 0 ? "mint" : "pole"} />
        <Stat label="Propinas" value={money(tips, cur)} sub={`ticket promedio ${money(Math.round(ticket), cur)}`} />
      </div>

      <div className="bt-card p-4 mt-3">
        <span className="bt-eyebrow">Ingresos vs. gastos por día</span>
        <div style={{ height: 240, marginTop: 10 }}>
          <ResponsiveContainer>
            <LineChart data={daily} margin={{ left: -12, right: 8, top: 8 }}>
              <CartesianGrid vertical={false} stroke={BRANDING.theme.line} />
              <XAxis dataKey="day" tick={{ fontSize: 10 }} interval="preserveStartEnd" />
              <YAxis tick={{ fontSize: 10 }} />
              <Tooltip formatter={(v) => money(v, cur)} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Line type="monotone" dataKey="Ingresos" stroke={BRANDING.theme.mint} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="Gastos" stroke={BRANDING.theme.pole} strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3 mt-3">
        <div className="bt-card p-4">
          <span className="bt-eyebrow">Por método de pago</span>
          {Object.keys(methodMap).length === 0 ? <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 8 }}>Sin cobros en el periodo.</p> :
            Object.entries(methodMap).map(([m, v]) => (
              <div key={m} className="flex items-center justify-between mt-2" style={{ fontSize: 13 }}>
                <span>{m}</span><span className="bt-mono" style={{ fontWeight: 600 }}>{money(v, cur)}</span>
              </div>
            ))}
        </div>
        <div className="bt-card p-4">
          <span className="bt-eyebrow">Rendimiento por barbero</span>
          {barberRows.length === 0 ? <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 8 }}>Sin ventas en el periodo.</p> :
            barberRows.map((b) => (
              <div key={b.name} className="flex items-center justify-between mt-2" style={{ fontSize: 13 }}>
                <span>{b.name} <span style={{ color: "var(--slate)" }}>· {b.ventas}</span></span>
                <span className="bt-mono" style={{ fontWeight: 600 }}>{money(b.total, cur)}</span>
              </div>
            ))}
        </div>
        <div className="bt-card p-4 lg:col-span-2">
          <span className="bt-eyebrow">Servicios y productos más vendidos</span>
          {items.length === 0 ? <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 8 }}>Sin datos todavía.</p> :
            items.slice(0, 15).map((i) => (
              <div key={i.type + i.name} className="flex items-center justify-between mt-2" style={{ fontSize: 13 }}>
                <span>{i.name} <span style={{ color: "var(--slate)" }}>· {i.qty} · {typeName(i.type).toLowerCase()}</span></span>
                <span className="bt-mono" style={{ fontWeight: 600 }}>{money(i.total, cur)}</span>
              </div>
            ))}
        </div>
      </div>
      <p style={{ fontSize: 11, color: "var(--slate)", marginTop: 12 }}>
        Ingresos = ventas cobradas en caja (antes de propinas, con el descuento aplicado) + citas marcadas «atendidas» sin cobro, por su precio.
        Las ventas anuladas no cuentan.
      </p>
    </div>
  );
}

/* ==========================================================================
   [VISTAS] 11. Ajustes — negocio, horario, plantillas, personal, respaldos.
   ========================================================================== */
function SettingsView({ db, actions, user, brand }) {
  const [s, setS] = useState(db.settings);
  const b0 = brand || {};
  const [look, setLook] = useState(() => ({
    logo: b0.logo || "", plate: b0.plate === "light" ? "light" : "none", palette: b0.palette || "clasica", font: b0.font || "",
    custom: { primary: "#1B3A5C", accent: "#B08725", font: "industrial", ...(b0.custom || {}) },
  }));
  const lookTokens = resolveTokens(look);
  const logoRef = useRef(null);
  const pickLogo = async (file) => {
    setBusy("logo");
    try { const r = await fileToLogo(file); setLook((l) => ({ ...l, ...r })); }
    catch (e) { actions.toast(e.userMessage || "No se pudo usar esa imagen.", "pole"); }
    setBusy("");
  };
  const [payText, setPayText] = useState((db.settings.payMethods || []).join(", "));
  const [staff, setStaff] = useState(null);
  const [licCode, setLicCode] = useState("");
  const [busy, setBusy] = useState("");
  const fileRef = useRef(null);

  const importBackup = (file) => {
    const r = new FileReader();
    r.onload = async () => {
      try {
        const json = JSON.parse(r.result);
        if (!json.settings || !json.appointments) throw new Error("formato");
        if (!confirm("Esto reemplaza la configuración, clientes, productos y gastos por los del respaldo, y vuelve a cargar sus citas, lista de espera y ventas. ¿Continuar?")) return;
        setBusy("restore");
        await actions.restore(json);
        setBusy("");
      } catch (e) { actions.toast("El archivo no es un respaldo válido de esta app.", "pole"); }
    };
    r.readAsText(file);
  };

  return (
    <div style={{ maxWidth: 720 }}>
      <h1 className="bt-display" style={{ fontSize: 20 }}>Ajustes</h1>
      <p className="bt-eyebrow">Configuración del negocio</p>

      <div className="bt-card p-4 mt-4 flex flex-col gap-3">
        <span className="bt-eyebrow">El negocio</span>
        <Field label="Nombre"><input className="bt-input" value={s.businessName} onChange={(e) => setS({ ...s, businessName: e.target.value })} /></Field>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <Field label="WhatsApp del negocio" hint="Con indicativo, solo números.">
            <input className="bt-input bt-mono" value={s.businessPhone} onChange={(e) => setS({ ...s, businessPhone: e.target.value })} />
          </Field>
          <Field label="Símbolo de moneda"><input className="bt-input" value={s.currency} onChange={(e) => setS({ ...s, currency: e.target.value })} /></Field>
          <Field label="Intervalo de agenda (min)">
            <select className="bt-select" value={s.slotMinutes} onChange={(e) => setS({ ...s, slotMinutes: Number(e.target.value) })}>
              {[15, 20, 30, 45, 60].map((n) => <option key={n} value={n}>{n}</option>)}
            </select>
          </Field>
        </div>
        <Field label="Anticipación mínima para reservar" hint="Ej.: con 1 hora, una cita de la 1:30pm deja de poder pedirse desde la 12:30pm. Solo aplica a lo que reserva el cliente; el mostrador sigue viendo todo.">
          <select className="bt-select" value={s.bookingLeadMinutes} onChange={(e) => setS({ ...s, bookingLeadMinutes: Number(e.target.value) })}>
            <option value={0}>Sin anticipación (hasta el último minuto)</option>
            <option value={15}>15 minutos antes</option>
            <option value={30}>30 minutos antes</option>
            <option value={60}>1 hora antes</option>
            <option value={120}>2 horas antes</option>
            <option value={180}>3 horas antes</option>
            <option value={240}>4 horas antes</option>
            <option value={360}>6 horas antes</option>
            <option value={720}>12 horas antes</option>
            <option value={1440}>1 día antes</option>
          </select>
        </Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Abre"><input className="bt-input bt-mono" type="time" value={s.openTime} onChange={(e) => setS({ ...s, openTime: e.target.value })} /></Field>
          <Field label="Cierra"><input className="bt-input bt-mono" type="time" value={s.closeTime} onChange={(e) => setS({ ...s, closeTime: e.target.value })} /></Field>
        </div>
        <div>
          <span className="bt-label">Días que abre (por defecto)</span>
          <div className="flex flex-wrap gap-1">
            {DOW_S.map((d, i) => (
              <button key={i} className={`bt-slot ${s.workDays.includes(i) ? "on" : ""}`} style={{ width: 56, fontFamily: "var(--body)" }}
                      onClick={() => setS({ ...s, workDays: s.workDays.includes(i) ? s.workDays.filter((x) => x !== i) : [...s.workDays, i].sort() })}>{d}</button>
            ))}
          </div>
        </div>
        <label className="flex items-center gap-2" style={{ fontSize: 13 }}>
          <input type="checkbox" checked={s.soundEnabled} onChange={(e) => setS({ ...s, soundEnabled: e.target.checked })} />
          Sonar cuando entre una solicitud
          <button className="bt-btn ghost sm" onClick={() => Sound.chime()}><Volume2 size={12} /> Probar</button>
        </label>
      </div>

      <div className="bt-card p-4 mt-4">
        <span className="bt-eyebrow">Apariencia</span>
        <p style={{ fontSize: 12, color: "var(--slate)", marginTop: 4 }}>El logo y los colores que ven tus clientes al reservar y tu equipo en el panel.</p>

        {/* Vista previa: se pinta con la paleta elegida AHORA, aunque todavía no la hayas guardado */}
        <div className="bt mt-3" style={{ ...tokensStyle(lookTokens), border: "1px solid var(--line)", borderRadius: 6, overflow: "hidden" }}>
          <div className="bt-nav flex items-center gap-3 px-3 py-2">
            <span className={look.logo && look.plate === "light" ? "bt-logoplate" : ""}>
              <img className="bt-logo" src={look.logo || defaultLogo(db.settings.businessName, lookTokens)} alt="" style={{ height: 34 }} />
            </span>
            <span className="bt-display" style={{ color: "#fff", fontSize: 13 }}>{db.settings.businessName}</span>
          </div>
          <div style={{ padding: 12, background: "var(--porcelain)" }}>
            <div className="bt-ticket is-conf">
              <div style={{ fontWeight: 600, fontSize: 14 }}>Corte clásico</div>
              <div className="bt-mono" style={{ fontSize: 12, color: "var(--slate)" }}>10:00 AM · {money(25000, db.settings.currency)}</div>
            </div>
            <div className="flex flex-wrap items-center gap-2 mt-2">
              <button type="button" className="bt-btn brass sm">Pedir una cita</button>
              <button type="button" className="bt-btn sm">Entrar</button>
              <span className="bt-tag conf">confirmada</span>
              <span className="bt-tag pend">pendiente</span>
            </div>
          </div>
        </div>

        <span className="bt-label mt-4" style={{ display: "block" }}>Logo</span>
        <div className="flex items-center gap-3">
          <div className="bt-inset p-2" style={{ minWidth: 72, textAlign: "center" }}>
            <img src={look.logo || defaultLogo(db.settings.businessName, lookTokens)} alt="" style={{ height: 56, maxWidth: 120, objectFit: "contain" }} />
          </div>
          <div className="flex flex-col gap-2">
            <button type="button" className="bt-btn ghost sm" disabled={busy === "logo"} onClick={() => logoRef.current.click()}>
              <Plus size={13} /> {busy === "logo" ? "Procesando…" : "Elegir imagen"}
            </button>
            {look.logo && <button type="button" className="bt-btn ghost sm" onClick={() => setLook({ ...look, logo: "", plate: "none" })}>Quitar logo</button>}
          </div>
          <input ref={logoRef} type="file" accept="image/*" style={{ display: "none" }}
                 onChange={(e) => { const f = e.target.files && e.target.files[0]; e.target.value = ""; if (f) pickLogo(f); }} />
        </div>
        <label className="flex items-start gap-2 mt-2" style={{ fontSize: 12 }}>
          <input type="checkbox" checked={look.plate === "light"} onChange={(e) => setLook({ ...look, plate: e.target.checked ? "light" : "none" })} />
          <span>Fondo claro detrás del logo en la barra oscura (actívalo si el logo es negro u oscuro y no se ve).</span>
        </label>
        <p style={{ fontSize: 11, color: "var(--slate)", marginTop: 6 }}>Mejor un PNG con fondo transparente. Se reduce solo para que pese poco.</p>

        <span className="bt-label mt-4" style={{ display: "block" }}>Paleta de colores</span>
        <div className="grid grid-cols-2 gap-2">
          {PALETTES.map((p) => (
            <button key={p.id} type="button" className="bt-card p-2 text-left"
                    style={{ cursor: "pointer", outline: look.palette === p.id ? "2px solid var(--brass)" : "none" }}
                    onClick={() => setLook({ ...look, palette: p.id })}>
              <div className="flex gap-1">
                {[p.ink, p.brass, p.porcelain, p.mint].map((c, i) => <span key={i} style={{ width: 22, height: 22, borderRadius: 6, background: c, border: "1px solid rgba(0,0,0,.12)" }} />)}
              </div>
              <div style={{ fontSize: 12, marginTop: 6, fontWeight: 600 }}>{p.name}{look.palette === p.id ? " ✓" : ""}</div>
            </button>
          ))}
          <button type="button" className="bt-card p-2 text-left"
                  style={{ cursor: "pointer", outline: look.palette === "custom" ? "2px solid var(--brass)" : "none" }}
                  onClick={() => setLook({ ...look, palette: "custom" })}>
            <div className="flex gap-1">
              {[deriveTokens(look.custom.primary, look.custom.accent).ink, deriveTokens(look.custom.primary, look.custom.accent).brass, deriveTokens(look.custom.primary, look.custom.accent).porcelain].map((c, i) => (
                <span key={i} style={{ width: 22, height: 22, borderRadius: 6, background: c, border: "1px solid rgba(0,0,0,.12)" }} />
              ))}
            </div>
            <div style={{ fontSize: 12, marginTop: 6, fontWeight: 600 }}>Personalizada{look.palette === "custom" ? " ✓" : ""}</div>
          </button>
        </div>

        {look.palette === "custom" && (
          <div className="grid grid-cols-2 gap-3 mt-3">
            <Field label="Color principal" hint="Barra lateral y botones oscuros.">
              <input type="color" className="bt-input" style={{ height: 44, padding: 4 }} value={look.custom.primary}
                     onChange={(e) => setLook({ ...look, custom: { ...look.custom, primary: e.target.value } })} />
            </Field>
            <Field label="Color de acento" hint="Botones de acción y detalles.">
              <input type="color" className="bt-input" style={{ height: 44, padding: 4 }} value={look.custom.accent}
                     onChange={(e) => setLook({ ...look, custom: { ...look.custom, accent: e.target.value } })} />
            </Field>
          </div>
        )}

        <div className="mt-3">
          <Field label="Tipo de letra">
            <select className="bt-select" value={lookTokens.font} onChange={(e) => setLook({ ...look, font: e.target.value })}>
              <option value="industrial">Industrial — condensada, estilo barbería</option>
              <option value="elegante">Elegante — con serifa, estilo salón</option>
            </select>
          </Field>
        </div>

        <button type="button" className="bt-btn mt-4" disabled={busy === "look"}
                onClick={async () => { setBusy("look"); await actions.saveBrand(look); setBusy(""); }}>
          <Save size={13} /> {busy === "look" ? "Guardando…" : "Guardar apariencia"}
        </button>
      </div>

      <div className="bt-card p-4 mt-3 flex flex-col gap-3">
        <span className="bt-eyebrow">Caja</span>
        <Field label="Métodos de pago" hint="Sepáralos con coma. «Efectivo» activa el cálculo de vuelto.">
          <input className="bt-input" value={payText} onChange={(e) => setPayText(e.target.value)} />
        </Field>
        <Field label="Mensaje al pie del recibo">
          <input className="bt-input" value={s.receiptFooter || ""} onChange={(e) => setS({ ...s, receiptFooter: e.target.value })} />
        </Field>
      </div>

      <div className="bt-card p-4 mt-3 flex flex-col gap-3">
        <span className="bt-eyebrow">Mensajes de WhatsApp</span>
        <p style={{ fontSize: 12, color: "var(--slate)" }}>
          Marcadores disponibles: {"{cliente} {telefono} {negocio} {fecha} {hora} {barbero} {servicio} {duracion} {precio} {comentario}"}
        </p>
        <Field label="Mensaje al cliente"><textarea className="bt-area" rows={7} value={s.tplClient} onChange={(e) => setS({ ...s, tplClient: e.target.value })} /></Field>
        <Field label="Aviso al negocio · cita confirmada"><textarea className="bt-area" rows={6} value={s.tplBusiness} onChange={(e) => setS({ ...s, tplBusiness: e.target.value })} /></Field>
        <Field label="Aviso al negocio · solicitud nueva" hint="Este es el que se abre solo en el celular del cliente apenas pide la cita.">
          <textarea className="bt-area" rows={6} value={s.tplRequestBusiness} onChange={(e) => setS({ ...s, tplRequestBusiness: e.target.value })} />
        </Field>
      </div>

      <button className="bt-btn mt-3" onClick={() => {
        const pm = payText.split(",").map((x) => x.trim()).filter(Boolean);
        actions.saveSettings({ ...s, payMethods: pm.length ? pm : db.settings.payMethods });
      }}><Save size={13} /> Guardar ajustes</button>

      <div className="bt-card p-4 mt-5">
        <div className="flex items-center justify-between">
          <span className="bt-eyebrow">Personal y accesos</span>
          <button className="bt-btn ghost sm" onClick={() => setStaff({ id: "", name: "", rol: "barbero", email: "", active: true })}><Plus size={13} /> Agregar</button>
        </div>
        {db.staff.map((u) => (
          <div key={u.id} className="flex items-center justify-between mt-2" style={{ fontSize: 13, opacity: u.active === false ? 0.5 : 1 }}>
            <span className="min-w-0">{u.name} <span className="bt-tag done" style={{ marginLeft: 6 }}>{u.rol}</span>
              <div className="bt-mono" style={{ fontSize: 11, color: "var(--slate)" }}>{u.email}{u.active === false ? " · desactivado" : ""}</div>
            </span>
            <div className="flex gap-1">
              <button className="bt-icon" onClick={() => setStaff(clone(u))} aria-label="Editar"><Pencil size={13} /></button>
              {u.id !== user.id && <button className="bt-icon" onClick={() => { if (confirm(`¿Quitar el acceso de ${u.name}?`)) actions.remove("staff", u.id, "Acceso eliminado"); }} aria-label="Borrar"><Trash2 size={13} /></button>}
            </div>
          </div>
        ))}
        <p style={{ fontSize: 11, color: "var(--slate)", marginTop: 10 }}>
          Cada persona entra con su correo y contraseña reales, creados en Firebase → Authentication → Users.
          Aquí defines su nombre y su rol. Desactivar a alguien le cierra el panel sin tocar su contraseña.
          Los roles ocultan pantallas; no son una barrera entre datos del mismo negocio.
        </p>
      </div>

      <div className="bt-card p-4 mt-3">
        <div className="flex items-center gap-2"><KeyRound size={14} style={{ color: "var(--brass)" }} /><span className="bt-eyebrow">Licencia</span></div>
        {!LICENSE_ON ? (
          <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 8 }}>
            Modo demostración: esta dirección no exige licencia. En la instalación de un cliente se pide el código que entrega Conty Licencias.
          </p>
        ) : (() => {
          const st = licenseStatus(db.license);
          const d = db.license && db.license.d;
          const okTag = st.tipo === "ok" || st.tipo === "pronto" || st.tipo === "gracia";
          return (
            <>
              <div className="bt-inset p-3 mt-2" style={{ fontSize: 13 }}>
                <div>Producto: <strong>{LICENSE.appName}</strong></div>
                <div>Código: <strong className="bt-mono">{(db.license && db.license.code) || "—"}</strong></div>
                <div>Negocio: <strong>{(d && d.negocio) || "—"}</strong></div>
                <div>Pagado hasta: <strong>{d ? (d.permanente ? "Para siempre" : d.vence || "—") : "—"}</strong></div>
                <div>Estado: <span className={`bt-tag ${okTag ? "conf" : "canc"}`}>{okTag ? "activa" : st.tipo === "falta" ? "sin activar" : "bloqueada"}</span></div>
                {st.texto && <div style={{ color: "var(--slate)", marginTop: 4 }}>{st.texto}</div>}
              </div>
              <div className="flex flex-wrap gap-2 mt-3">
                <button className="bt-btn ghost sm" disabled={!db.license || busy === "re"}
                        onClick={async () => { setBusy("re"); await actions.recheckLicense(); setBusy(""); }}>
                  <RefreshCw size={13} /> {busy === "re" ? "Revisando…" : "Revisar ahora"}
                </button>
              </div>
              <div className="flex gap-2 mt-3">
                <input className="bt-input bt-mono" placeholder="Cambiar código (BRB-XXXXX)" value={licCode} onChange={(e) => setLicCode(e.target.value)} />
                <button className="bt-btn" disabled={!licCode.trim() || busy === "lic"}
                        onClick={async () => { setBusy("lic"); await actions.activateLicense(licCode); setBusy(""); setLicCode(""); }}>
                  <ShieldCheck size={13} /> Verificar
                </button>
              </div>
            </>
          );
        })()}
      </div>

      {db.legacy > 0 && (
        <div className="bt-card p-4 mt-3" style={{ borderLeft: "3px solid var(--brass)" }}>
          <span className="bt-eyebrow">Datos del sistema anterior</span>
          <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 6 }}>
            Hay {db.legacy} registro{db.legacy === 1 ? "" : "s"} (citas y lista de espera) guardados con el sistema viejo. Migrarlos los pasa al
            sistema nuevo y seguro. Haz antes un respaldo. El proceso se puede repetir sin duplicar nada.
          </p>
          <button className="bt-btn mt-3" disabled={busy === "mig"}
                  onClick={async () => { setBusy("mig"); await actions.migrateLegacy(); setBusy(""); }}>
            <RefreshCw size={13} /> {busy === "mig" ? "Migrando…" : "Migrar datos ahora"}
          </button>
        </div>
      )}

      <div className="bt-card p-4 mt-3">
        <span className="bt-eyebrow">Respaldo de datos</span>
        <p style={{ fontSize: 13, color: "var(--slate)", marginTop: 6 }}>
          Descarga una copia completa (clientes, citas, ventas, gastos, inventario). Guárdala fuera del dispositivo una vez por semana.
        </p>
        <div className="flex flex-wrap gap-2 mt-3">
          <button className="bt-btn ghost sm" disabled={busy === "bk"}
                  onClick={async () => { setBusy("bk"); try { await downloadFullBackup(db); } catch (e) { console.error(e); actions.toast("No se pudo generar el respaldo.", "pole"); } setBusy(""); }}>
            <Download size={13} /> {busy === "bk" ? "Preparando…" : "Descargar respaldo"}
          </button>
          <button className="bt-btn ghost sm" disabled={busy === "restore"} onClick={() => fileRef.current.click()}><RefreshCw size={13} /> Restaurar respaldo</button>
          <input ref={fileRef} type="file" accept="application/json" style={{ display: "none" }}
                 onChange={(e) => e.target.files[0] && importBackup(e.target.files[0])} />
        </div>
        <div className="bt-inset p-2 mt-3 bt-mono" style={{ fontSize: 11, color: "var(--slate)" }}>
          {db.clients.length} clientes · {db.expenses.length} gastos · {db.products.length} productos · versión {db.rev}
        </div>
      </div>

      {staff && (
        <Modal title={staff.id ? "Editar acceso" : "Nuevo acceso"} onClose={() => setStaff(null)}>
          <div className="flex flex-col gap-3">
            <Field label="Nombre" hint="Para un barbero, escríbelo igual que en la ficha de Barberos: así solo verá su agenda.">
              <input className="bt-input" value={staff.name} onChange={(e) => setStaff({ ...staff, name: e.target.value })} />
            </Field>
            <Field label="Rol">
              <select className="bt-select" value={staff.rol} onChange={(e) => setStaff({ ...staff, rol: e.target.value })}>
                <option value="admin">Administrador — todo</option>
                <option value="recepcion">Recepción — agenda y clientes</option>
                <option value="barbero">Barbero — solo su agenda</option>
              </select>
            </Field>
            <Field label="Correo de acceso" hint="Debe existir primero como usuario en Firebase → Authentication → Users (ahí se le pone la contraseña).">
              <input className="bt-input" type="email" autoCapitalize="none" value={staff.email || ""} onChange={(e) => setStaff({ ...staff, email: e.target.value })} />
            </Field>
            <label className="flex items-center gap-2" style={{ fontSize: 13 }}>
              <input type="checkbox" checked={staff.active !== false} onChange={(e) => setStaff({ ...staff, active: e.target.checked })} />
              Acceso activo
            </label>
            <button className="bt-btn" disabled={!staff.name.trim() || !(staff.email || "").includes("@")}
                    onClick={async () => { const ok = await actions.save("staff", staff, "Acceso guardado"); if (ok) setStaff(null); }}>
              <Save size={13} /> Guardar
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}
