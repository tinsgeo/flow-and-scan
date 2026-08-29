/* ============================================================
   Flow & Scan Practice Site — shared app logic
   This file powers the interactive, visual dashboards.
   It does NOT power the plain data tables under /data/ —
   those are static HTML, regenerated on a schedule, and are
   meant to be pulled with VBA/Excel web queries instead of
   read by a browser running JavaScript.
   ============================================================ */

const POSITIONS = ["Pick", "Pack", "Sort", "Induct", "Stow"];

const FIRST_NAMES = ["Jordan","Casey","Morgan","Taylor","Riley","Avery","Devon","Reese",
  "Sydney","Cameron","Peyton","Quinn","Rowan","Emerson","Harper","Elliot"];
const LAST_NAMES = ["Bell","Reyes","Whitfield","Nakamura","Alvarez","Sutton","Okafor",
  "Petrov","Lindgren","Marsh","Delgado","Osei","Kowalski","Fontaine"];

function seededRandom(seed) {
  // simple deterministic PRNG so a given session is stable until "regenerate" is clicked
  let s = seed % 2147483647;
  if (s <= 0) s += 2147483646;
  return function () {
    s = (s * 16807) % 2147483647;
    return (s - 1) / 2147483646;
  };
}

function pad(n) { return n.toString().padStart(2, "0"); }

function fmtTime(minutesFromMidnight) {
  let h = Math.floor(minutesFromMidnight / 60);
  const m = Math.floor(minutesFromMidnight % 60);
  const ampm = h >= 12 ? "PM" : "AM";
  let h12 = h % 12;
  if (h12 === 0) h12 = 12;
  return `${h12}:${pad(m)} ${ampm}`;
}

function generateAssociates(rand, count = 14) {
  const shiftStart = 6 * 60;      // 6:00 AM
  const shiftEnd = 14.5 * 60;     // 2:30 PM
  const associates = [];

  for (let i = 0; i < count; i++) {
    const id = 1000 + i;
    const name = `${FIRST_NAMES[Math.floor(rand() * FIRST_NAMES.length)]} ${LAST_NAMES[Math.floor(rand() * LAST_NAMES.length)]}`;
    const position = POSITIONS[Math.floor(rand() * POSITIONS.length)];

    const scans = [];
    let t = shiftStart + Math.floor(rand() * 10);
    let unitCounter = 1;
    while (t < shiftEnd) {
      // normal interval between scans: 2-9 minutes
      let interval = 2 + Math.floor(rand() * 7);
      // occasionally inject a real gap (18-45 min) to flag
      if (rand() < 0.12) {
        interval = 18 + Math.floor(rand() * 27);
      }
      t += interval;
      if (t >= shiftEnd) break;
      scans.push({ time: t, unit: `SKU-${10000 + Math.floor(rand() * 89999)}` });
      unitCounter++;
    }

    associates.push({ id, name, position, scans });
  }
  return associates;
}

function computeGaps(scans, thresholdMinutes = 15) {
  const gaps = [];
  for (let i = 1; i < scans.length; i++) {
    const delta = scans[i].time - scans[i - 1].time;
    gaps.push({
      from: scans[i - 1].time,
      to: scans[i].time,
      minutes: delta,
      flagged: delta >= thresholdMinutes,
      unitBefore: scans[i - 1].unit,
      unitAfter: scans[i].unit,
    });
  }
  return gaps;
}

function ratePerHour(scans, shiftStart = 6 * 60, shiftEnd = 14.5 * 60) {
  const hours = (shiftEnd - shiftStart) / 60;
  return (scans.length / hours).toFixed(1);
}

/* ---------------- Productivity dashboard ---------------- */

let ASSOCIATES = [];
let SELECTED_ID = null;

function initProductivity() {
  const seed = Date.now() % 100000;
  ASSOCIATES = generateAssociates(seededRandom(seed));
  populatePositionFilter();
  renderProductivityTable();
  document.getElementById("regenBtn").addEventListener("click", () => {
    ASSOCIATES = generateAssociates(seededRandom(Date.now() % 100000));
    SELECTED_ID = null;
    renderProductivityTable();
    document.getElementById("detailPanel").innerHTML = "";
  });
  document.getElementById("posFilter").addEventListener("change", renderProductivityTable);
  document.getElementById("periodFilter").addEventListener("change", renderProductivityTable);
}

function populatePositionFilter() {
  const sel = document.getElementById("posFilter");
  sel.innerHTML = `<option value="all">All positions</option>` +
    POSITIONS.map(p => `<option value="${p}">${p}</option>`).join("");
}

function periodBounds(periodValue) {
  if (periodValue === "full") return [6 * 60, 14.5 * 60];
  const startHour = parseInt(periodValue, 10);
  return [startHour * 60, (startHour + 1) * 60];
}

function renderProductivityTable() {
  const posVal = document.getElementById("posFilter").value;
  const periodVal = document.getElementById("periodFilter").value;
  const [pStart, pEnd] = periodBounds(periodVal);

  const rows = ASSOCIATES
    .filter(a => posVal === "all" || a.position === posVal)
    .map(a => {
      const scansInPeriod = a.scans.filter(s => s.time >= pStart && s.time < pEnd);
      const gaps = computeGaps(a.scans);
      const flaggedCount = gaps.filter(g => g.flagged).length;
      return { a, scansInPeriod, flaggedCount };
    });

  const tbody = document.getElementById("prodTbody");
  tbody.innerHTML = rows.map(({ a, scansInPeriod, flaggedCount }) => `
    <tr data-id="${a.id}" class="${SELECTED_ID === a.id ? 'selected' : ''}">
      <td>${a.id}</td>
      <td>${a.name}</td>
      <td>${a.position}</td>
      <td class="num">${scansInPeriod.length}</td>
      <td class="num">${ratePerHour(a.scans)}</td>
      <td>${flaggedCount > 0
        ? `<span class="badge flag">${flaggedCount} gap${flaggedCount > 1 ? "s" : ""}</span>`
        : `<span class="badge ok">on pace</span>`}</td>
    </tr>
  `).join("");

  tbody.querySelectorAll("tr").forEach(tr => {
    tr.addEventListener("click", () => {
      SELECTED_ID = parseInt(tr.dataset.id, 10);
      renderProductivityTable();
      renderDetail(SELECTED_ID);
    });
  });
}

function renderDetail(id) {
  const a = ASSOCIATES.find(x => x.id === id);
  const panel = document.getElementById("detailPanel");
  if (!a) { panel.innerHTML = ""; return; }

  const gaps = computeGaps(a.scans);
  const shiftStart = 6 * 60, shiftEnd = 14.5 * 60, total = shiftEnd - shiftStart;

  const segs = gaps.map(g => {
    const widthPct = ((g.to - g.from) / total) * 100;
    return `<div class="seg ${g.flagged ? 'gapflag' : ''}" style="width:${widthPct}%" title="${fmtTime(g.from)} → ${fmtTime(g.to)}"></div>`;
  }).join("");

  const list = gaps.filter(g => g.flagged).map(g => `
    <li class="flag">
      <span>${fmtTime(g.from)} (scanned ${g.unitBefore}) → ${fmtTime(g.to)} (scanned ${g.unitAfter})</span>
      <span class="len">${g.minutes} min gap</span>
    </li>
  `).join("");

  panel.innerHTML = `
    <div class="panel detail">
      <h2 style="margin-top:0">${a.name} — ${a.position} <span class="plain-note">(ID ${a.id})</span></h2>
      <p class="lede">Scan timeline across the shift. Shaded/hatched segments are gaps of 15 minutes or more.</p>
      <div class="timeline">${segs}</div>
      ${list ? `<ul class="gap-list">${list}</ul>` : `<p class="plain-note">No gaps of 15+ minutes recorded for this associate.</p>`}
    </div>
  `;
}

/* ---------------- Orders flow dashboard ---------------- */

let ORDER_STATE = { pending: 0, created: 0, inProgress: 0, completed: 0 };
let ORDER_LOG = [];
let orderRand = seededRandom(Date.now() % 100000);

function initOrders() {
  ORDER_STATE = { pending: 40, created: 0, inProgress: 6, completed: 0 };
  renderOrderMetrics();
  tickOrders();
  setInterval(tickOrders, 2200);
}

function tickOrders() {
  const newOrders = Math.floor(orderRand() * 5);
  ORDER_STATE.created += newOrders;
  ORDER_STATE.pending += newOrders;

  const started = Math.min(ORDER_STATE.pending, Math.floor(orderRand() * 4));
  ORDER_STATE.pending -= started;
  ORDER_STATE.inProgress += started;

  const finished = Math.min(ORDER_STATE.inProgress, Math.floor(orderRand() * 4));
  ORDER_STATE.inProgress -= finished;
  ORDER_STATE.completed += finished;

  const now = new Date();
  ORDER_LOG.unshift({
    time: now.toLocaleTimeString(),
    pending: ORDER_STATE.pending,
    created: ORDER_STATE.created,
    inProgress: ORDER_STATE.inProgress,
    completed: ORDER_STATE.completed,
  });
  ORDER_LOG = ORDER_LOG.slice(0, 12);

  renderOrderMetrics();
  renderOrderLog();
}

function renderOrderMetrics() {
  document.getElementById("mPending").textContent = ORDER_STATE.pending;
  document.getElementById("mCreated").textContent = ORDER_STATE.created;
  document.getElementById("mProgress").textContent = ORDER_STATE.inProgress;
  document.getElementById("mCompleted").textContent = ORDER_STATE.completed;
}

function renderOrderLog() {
  const tbody = document.getElementById("orderTbody");
  tbody.innerHTML = ORDER_LOG.map(r => `
    <tr>
      <td>${r.time}</td>
      <td class="num">${r.pending}</td>
      <td class="num">${r.created}</td>
      <td class="num">${r.inProgress}</td>
      <td class="num">${r.completed}</td>
    </tr>
  `).join("");
}
