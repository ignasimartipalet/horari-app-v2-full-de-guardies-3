import { useState, useRef, useCallback, useMemo, useEffect } from "react";
import * as XLSX from "xlsx-js-style";

// ─── Constants ────────────────────────────────────────────────────────────────

const DAYS = ["DILLUNS", "DIMARTS", "DIMECRES", "DIJOUS", "DIVENDRES"];
const DAY_LABELS = {
  DILLUNS: "Dilluns",
  DIMARTS: "Dimarts",
  DIMECRES: "Dimecres",
  DIJOUS: "Dijous",
  DIVENDRES: "Divendres",
};

const MORNING_SLOTS = [
  "8:00-8:55",
  "8:55-9:50",
  "9:50-10:45",
  "10:45-11:15",
  "11:15-11:45",
  "11:45-12:40",
  "12:40-13:35",
  "13:35-14:30",
];

const SLOT_LABEL = {
  "8:00-8:55": "1a hora",
  "8:55-9:50": "2a hora",
  "9:50-10:45": "3a hora",
  "10:45-11:15": "Esbarjo/lectura 1r torn",
  "11:15-11:45": "Esbarjo/lectura 2n torn",
  "11:45-12:40": "4a hora",
  "12:40-13:35": "5a hora",
  "13:35-14:30": "6a hora",
};

const YEAR_GROUPS = [
  {
    label: "1r ESO",
    values: ["A", "B", "C", "D", "E"].map((l) => ({
      label: l,
      value: "1 ESO " + l,
    })),
  },
  {
    label: "2n ESO",
    values: ["A", "B", "C", "D", "E"].map((l) => ({
      label: l,
      value: "2 ESO " + l,
    })),
  },
  {
    label: "3r ESO",
    values: ["A", "B", "C", "D", "E"].map((l) => ({
      label: l,
      value: "3 ESO " + l,
    })),
  },
  {
    label: "4t ESO",
    values: ["A", "B", "C", "D", "E"].map((l) => ({
      label: l,
      value: "4 ESO " + l,
    })),
  },
  {
    label: "1r BAT",
    values: ["A", "B"].map((l) => ({ label: l, value: "BAT 1" + l })),
  },
  {
    label: "2n BAT",
    values: ["A", "B"].map((l) => ({ label: l, value: "BAT 2" + l })),
  },
];

// ─── Compactació de grups ────────────────────────────────────────────────────
// ["1 ESO A","1 ESO B","1 ESO C"] → "1r ESO A, B i C"
// ["1 ESO A","1 ESO B","1 ESO C","1 ESO D","1 ESO E"] → "Tot 1r d'ESO"
const CURS_LABELS = {
  "1 ESO": "1r ESO", "2 ESO": "2n ESO", "3 ESO": "3r ESO", "4 ESO": "4t ESO",
  "BAT 1": "1r BAT", "BAT 2": "2n BAT",
};
const CURS_TOT = {
  "1 ESO": "Tot 1r d'ESO", "2 ESO": "Tot 2n d'ESO",
  "3 ESO": "Tot 3r d'ESO", "4 ESO": "Tot 4t d'ESO",
  "BAT 1": "Tot 1r de BAT", "BAT 2": "Tot 2n de BAT",
};
const MAX_LLETRES = { "1 ESO":5,"2 ESO":5,"3 ESO":5,"4 ESO":5,"BAT 1":2,"BAT 2":2 };

function compactaGrups(grups) {
  if (!grups || grups.length === 0) return "";
  // Agrupa per curs
  const byCurs = {};
  grups.forEach(g => {
    const m = g.match(/^(\d ESO|BAT [12])\s+([A-E])$/i);
    if (m) {
      const curs = m[1].toUpperCase();
      if (!byCurs[curs]) byCurs[curs] = [];
      byCurs[curs].push(m[2].toUpperCase());
    }
  });
  if (Object.keys(byCurs).length === 0) return grups.join(", ");
  return Object.entries(byCurs).map(([curs, lletres]) => {
    lletres.sort();
    if (lletres.length === MAX_LLETRES[curs]) return CURS_TOT[curs] || curs;
    const label = CURS_LABELS[curs] || curs;
    if (lletres.length === 1) return `${label} ${lletres[0]}`;
    const darrers = lletres[lletres.length - 1];
    const resta = lletres.slice(0, -1).join(", ");
    return `${label} ${resta} i ${darrers}`;
  }).join(" · ");
}



function fixTime(t) {
  if (!t) return t;
  return t.replace(/^0(\d:)/, "$1").replace(/-0(\d:)/, "-$1");
}
function fixGroup(g) {
  if (!g) return g;
  const s = g.trim();
  if (/^\d ESO [A-E]$/i.test(s)) return s.toUpperCase().replace(/eso/i, "ESO");
  if (/^BAT [12][AB]$/i.test(s)) return s.toUpperCase();
  let m = s.match(/ESO\s+(\d)[rntèé°]*\s*([A-E])/i);
  if (m) return `${m[1]} ESO ${m[2].toUpperCase()}`;
  m = s.match(/(\d)[rntèé°\s]*\s*ESO\s*([A-E])/i);
  if (m) return `${m[1]} ESO ${m[2].toUpperCase()}`;
  m = s.match(/BAT\s*([12])\s*([AB])/i);
  if (m) return `BAT ${m[1]}${m[2].toUpperCase()}`;
  return s;
}
// ─── Clean subject strings ────────────────────────────────────────────────────
// Scraper sometimes appends teacher name to subject: "GUÀRDIA BVanessa Casanova" → "GUÀRDIA B"
// Strategy: remove everything after the last all-caps word boundary where a proper name starts
function cleanSubject(subject, teacherName) {
  if (!subject) return subject;
  let s = subject.trim();
  // If teacher name is known, strip it directly
  if (teacherName) {
    // Try full name
    s = s.replace(teacherName.trim(), "").trim();
    // Try last name only
    const lastName = teacherName.trim().split(" ").slice(-1)[0];
    if (lastName && lastName.length > 2)
      s = s.replace(new RegExp(lastName + ".*$"), "").trim();
    // Try first name only
    const firstName = teacherName.trim().split(" ")[0];
    if (firstName && firstName.length > 2)
      s = s.replace(new RegExp(firstName + ".*$"), "").trim();
  }
  // Generic: strip trailing capitalized words that look like proper names (Uppercase + lowercase)
  // e.g. "GUÀRDIA AVanessa" → strip "Vanessa"
  s = s
    .replace(
      /[A-ZÀÁÈÉÍÏÒÓÚÜ][a-zàáèéíïòóúü][a-zàáèéíïòóúüA-ZÀÁÈÉÍÏÒÓÚÜ\s]*$/,
      ""
    )
    .trim();
  return s || subject.trim(); // fallback to original if we stripped too much
}

function normalizeTeacher(t) {
  return {
    ...t,
    schedule: Object.fromEntries(
      Object.entries(t.schedule).map(([day, slots]) => [
        day,
        slots
          .map((s) => ({
            ...s,
            time: fixTime(s.time),
            group: fixGroup(s.group),
            subject: cleanSubject(s.subject, t.name),
          }))
          .filter((s) => MORNING_SLOTS.includes(s.time)),
      ])
    ),
  };
}

// ─── Group matching ───────────────────────────────────────────────────────────

function ngs(g) {
  return (g || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function groupMatches(slotGroup, baseGroup, excludedSubs = []) {
  if (!slotGroup) return false;
  const ns = ngs(slotGroup),
    nb = ngs(baseGroup);
  if (ns === nb) return true;
  // Subgrup: el slot ha de CONTENIR el baseGroup com a prefix/sufix complet
  // Evitem falsos positius: "aa" no ha de coincidir amb "a" (Aula Acollida vs grup A)
  if (nb.length >= 4 && ns.includes(nb))
    return !excludedSubs.some((e) => ngs(e) === ns);
  return false;
}
function slotWithTripGroups(slotGroup, selectedGroups, excludedSubs) {
  if (!selectedGroups.length) return false; // sense grups seleccionats → tot genera forat
  return selectedGroups.some((g) => groupMatches(slotGroup, g, excludedSubs));
}
function slotWithExcludedSub(slotGroup, excludedSubs) {
  return (excludedSubs || []).some((e) => ngs(e) === ngs(slotGroup));
}
function detectSubgroups(teachers, baseGroup) {
  const nb = ngs(baseGroup);
  const found = new Set();
  teachers.forEach((t) =>
    Object.values(t.schedule).forEach((slots) =>
      slots.forEach((s) => {
        if (!s.group) return;
        const ns = ngs(s.group);
        if (ns !== nb && ns.includes(nb) && s.group !== baseGroup)
          found.add(s.group);
      })
    )
  );
  return [...found].sort();
}

// ─── Utilities ────────────────────────────────────────────────────────────────

function fileToBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("Error llegint arxiu"));
    r.readAsDataURL(file);
  });
}
function getMediaType(file) {
  if (file.type === "application/pdf" || file.name.endsWith(".pdf"))
    return "application/pdf";
  if (file.type === "image/png") return "image/png";
  if (file.type === "image/webp") return "image/webp";
  return "image/jpeg";
}
function extractSubjects(teachers) {
  const s = new Set();
  teachers.forEach((t) =>
    Object.values(t.schedule).forEach((slots) =>
      slots.forEach((slot) => {
        if (slot.type === "class" && slot.subject && slot.subject.length > 2)
          s.add(slot.subject.trim());
      })
    )
  );
  return [...s].sort();
}
function slotsInRange(start, end) {
  const si = MORNING_SLOTS.indexOf(start),
    ei = MORNING_SLOTS.indexOf(end);
  return MORNING_SLOTS.filter((_, i) => i >= si && i <= ei);
}
function slotIdx(t) {
  return MORNING_SLOTS.indexOf(t);
}
function daySpan(daySchedule) {
  // Sort by slot index first — JSON may not be in order
  const occ = daySchedule
    .filter((s) => s.type !== "free")
    .sort((a, b) => slotIdx(a.time) - slotIdx(b.time));
  if (!occ.length) return null;
  return {
    first: slotIdx(occ[0].time),
    last: slotIdx(occ[occ.length - 1].time),
  };
}

function isGuardSlot(slot) {
  if (isPatiSlot(slot)) return false;
  if (isPatiEspecialSlot(slot)) return false; // biblioteca/música no és guàrdia normal
  if (slot.type === "guard") return true;
  const subj = (slot.subject || "")
    .toUpperCase()
    .replace(/À/g, "A")
    .replace(/È/g, "E");
  return subj.includes("GUARDIA");
}
function guardType(slot) {
  const subj = (slot.subject || "").toUpperCase().replace(/À/g, "A");
  if (subj.match(/GUARDIA\s*B\b/) || subj.endsWith(" B")) return "B";
  return "A";
}
function isPatiSlot(slot) {
  const subj = (slot.subject || "")
    .toUpperCase()
    .replace(/À/g, "A")
    .replace(/È/g, "E")
    .replace(/Ï/g, "I");
  return subj.includes("PATI") && !subj.includes("LECTURA");
}

function isLecturaGuard(slot) {
  const subj = (slot.subject || "")
    .toUpperCase()
    .replace(/À/g, "A")
    .replace(/È/g, "E");
  return subj.includes("GUARDIA") && subj.includes("LECTURA");
}
function isPatiEspecialSlot(slot) {
  const subj = (slot.subject || "")
    .toUpperCase()
    .replace(/À/g, "A")
    .replace(/È/g, "E")
    .replace(/Ï/g, "I");
  // Detecta tant la versió completa com l'abreujada: BIBLIOTECA/BIBL, MUSICA/MUSIC/MÚSICA
  return subj.includes("BIBL") || subj.includes("MUSIC");
}
function isOccupiedUnavailable(slot) {
  return isPatiSlot(slot);
}

// ─── API ──────────────────────────────────────────────────────────────────────

const PROMPT_RULES = `Franges: 8:00-8:55,8:55-9:50,9:50-10:45,10:45-11:15,11:15-11:45,11:45-12:40,12:40-13:35,13:35-14:30. type: class/guard/meeting/tutoring/free. Inclou totes les franges.`;
const SCHEMA_ONE = `{"name":"Nom","schedule":{"DILLUNS":[{"time":"8:00-8:55","subject":"Bio","group":"1 ESO D","room":"209","type":"class"},...],...}}`;
const SCHEMA_ALL = `[${SCHEMA_ONE},{...}]`;

async function parseSingleImage(base64, mediaType) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 4000,
      messages: [
        {
          role: "user",
          content: [
            {
              type: "image",
              source: { type: "base64", media_type: mediaType, data: base64 },
            },
            {
              type: "text",
              text: `Extreu l'horari. ÚNICAMENT JSON.\nFormat:${SCHEMA_ONE}\n${PROMPT_RULES}`,
            },
          ],
        },
      ],
    }),
  });
  const d = await res.json();
  if (d.error) throw new Error(d.error.message);
  return JSON.parse(
    d.content[0].text
      .replace(/```json\n?/g, "")
      .replace(/```\n?/g, "")
      .trim()
  );
}
async function parseAllFromPDF(base64) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 16000,
      messages: [
        {
          role: "user",
          content: [
            {
              type: "document",
              source: {
                type: "base64",
                media_type: "application/pdf",
                data: base64,
              },
            },
            {
              type: "text",
              text: `Extreu TOTS els professors. ÚNICAMENT array JSON.\nFormat:${SCHEMA_ALL}\n${PROMPT_RULES}`,
            },
          ],
        },
      ],
    }),
  });
  const d = await res.json();
  if (d.error) throw new Error(d.error.message);
  return JSON.parse(
    d.content[0].text
      .replace(/```json\n?/g, "")
      .replace(/```\n?/g, "")
      .trim()
  );
}

function isReadingSlot(slot) {
  return (slot.subject || "").toUpperCase().includes("LECTURA");
}
// Weight: reading = 0.5, normal = 1
function slotWeight(slot) {
  return isReadingSlot(slot) ? 0.5 : 1;
}

// Human-readable entry/exit time from span
function spanLabel(span, daySchedule) {
  if (!span) return null;
  const firstSlot = MORNING_SLOTS[span.first];
  const lastSlot = MORNING_SLOTS[span.last];
  const entryTime = firstSlot ? firstSlot.split("-")[0] : null;
  const exitTime = lastSlot ? lastSlot.split("-")[1] : null;
  if (!entryTime || !exitTime) return null;
  return `Entra ${entryTime} · Surt ${exitTime}`;
}

// How many slots does this teacher actually "span" on trip day (first occupied → last occupied)?
// hoursShort = how many range slots fall outside [first,last] occupied
function rangeHoursShort(span, si, ei) {
  if (!span) return ei - si + 1; // completely absent
  const effectiveFirst = Math.max(span.first, si);
  const effectiveLast = Math.min(span.last, ei);
  if (effectiveLast < effectiveFirst) return ei - si + 1;
  return effectiveFirst - si + (ei - effectiveLast);
}

// Does teacher teach any of the trip groups in ANY day of the week?
function teachesAnyTripGroup(teacher, selectedGroups, excludedSubs) {
  if (!selectedGroups.length) return false;
  return Object.values(teacher.schedule).some((slots) =>
    slots.some(
      (s) =>
        (s.type === "class" ||
          s.type === "tutoring" ||
          s.type === "guard" ||
          s.type === "reading") &&
        slotWithTripGroups(s.group, selectedGroups, excludedSubs)
    )
  );
}

function analyzeTeacher(teacher, trip) {
  const {
    day,
    startSlot,
    endSlot,
    selectedGroups,
    excludedSubs,
    halfGroups,
    subject,
  } = trip;
  const daySchedule = teacher.schedule[day] || [];
  const rangeSlots = slotsInRange(startSlot, endSlot);
  const si = slotIdx(startSlot),
    ei = slotIdx(endSlot);
  const totalRangeSlots = ei - si + 1;
  const inRange = daySchedule.filter((s) => rangeSlots.includes(s.time));
  const span = daySpan(daySchedule);
  const coversFullRange = span ? span.first <= si && span.last >= ei : false;
  const hoursShort = rangeHoursShort(span, si, ei); // 0=perfect, 1=one hour short, etc.
  const morningOccupied = daySchedule.filter((s) => s.type !== "free").length;

  let tripClasses = [],
    stayClasses = [],
    halfClasses = [],
    guardSlots = [],
    freeCount = 0,
    subjectMatch = false;
  if (subject)
    Object.values(teacher.schedule).forEach((slots) =>
      slots.forEach((s) => {
        if (
          s.type === "class" &&
          s.subject.toLowerCase().includes(subject.toLowerCase())
        )
          subjectMatch = true;
      })
    );

  inRange.forEach((s) => {
    if (s.type === "free") {
      freeCount++;
      return;
    }
    if (isGuardSlot(s)) {
      guardSlots.push(s);
      return;
    }
    if (s.type === "meeting" || (s.type === "tutoring" && !s.group)) return; // occupied but not a class
    if (s.type === "class" || s.type === "tutoring") {
      if (!selectedGroups.length) {
        tripClasses.push(s);
        return;
      }
      if (slotWithExcludedSub(s.group, excludedSubs)) {
        halfClasses.push(s);
        return;
      }
      const isHalfGroup = (halfGroups || []).some((g) =>
        groupMatches(s.group, g, [])
      );
      if (isHalfGroup) {
        halfClasses.push(s);
        return;
      }
      if (slotWithTripGroups(s.group, selectedGroups, excludedSubs))
        tripClasses.push(s);
      else stayClasses.push(s);
    }
  });

  // Does teacher teach the trip group at all (any day)?
  const teachesGroup = teachesAnyTripGroup(
    teacher,
    selectedGroups,
    excludedSubs
  );

  // Weighted covered slots (reading = 0.5)
  const coveredSlots = tripClasses.reduce((acc, s) => acc + slotWeight(s), 0);

  // Range coverage score: starts at MAX, decays heavily per hour short
  // Full range = 60pts, -15 per hour short (so 1h short=45, 2h=30, 3h=15, 4h+=0)
  const rangeCoverageScore = Math.max(0, 60 - hoursShort * 15);

  // Group teacher bonus (teaches this group any day of the week)
  const groupTeacherBonus = teachesGroup ? 50 : 0;

  const score = Math.max(
    0,
    rangeCoverageScore +
      groupTeacherBonus +
      coveredSlots * 20 +
      (subjectMatch ? 15 : 0) +
      morningOccupied * 3 +
      freeCount * 1 -
      stayClasses.length * 12 -
      halfClasses.length * 6
  );

  const label = spanLabel(span, daySchedule);

  return {
    name: teacher.name,
    score,
    coveredSlots,
    totalRange: totalRangeSlots,
    coversFullRange,
    hoursShort,
    teachesGroup,
    spanLabel: label,
    tripClasses,
    stayClasses,
    halfClasses,
    guardSlots,
    freeCount,
    morningOccupied,
    subjectMatch,
    daySchedule,
    inRange,
    span,
  };
}

// ─── Coverage Engine ──────────────────────────────────────────────────────────

function computeCoverage(teachers, confirmedNames, trip, franges, teacherFranjaMap) {
  const { day, startSlot, endSlot, selectedGroups, excludedSubs, halfGroups } = trip;

  const rangeSlots = slotsInRange(startSlot, endSlot);
  const confirmedSet = new Set(confirmedNames);

  const goTeachers = teachers.filter((t) => confirmedSet.has(t.name));
  const stayTeachers = teachers.filter((t) => !confirmedSet.has(t.name));

  const gapsBySlot = {};
  const coversBySlot = {};

  rangeSlots.forEach((time) => {
    gapsBySlot[time] = [];
    coversBySlot[time] = { freed: [], guardL: [], guardE: [], guardA: [], guardB: [], free: [] };
  });

  // Helper: slots del torn assignat a un professor (o tots si no hi ha assignació per torn)
  const getSlotsForTeacher = (teacherName) => {
    if (!franges || !teacherFranjaMap) return rangeSlots;
    const assignedIds = teacherFranjaMap[teacherName] || [];
    if (assignedIds.length === 0) return rangeSlots;
    const tornSlots = assignedIds.flatMap(id => {
      const f = franges.find(f => f.id === id);
      return f ? slotsInRange(f.startSlot, f.endSlot) : [];
    });
    return [...new Set(tornSlots)];
  };

  // CLASSES QUE QUEDEN SENSE COBRIR
  goTeachers.forEach((teacher) => {
    const daySlots = teacher.schedule[day] || [];
    const teacherSlots = getSlotsForTeacher(teacher.name);

    teacherSlots.forEach((time) => {
      if (!gapsBySlot[time]) return; // fora del rang global
      const slot = daySlots.find((s) => s.time === time);
      if (!slot) return;
      if (slot.type === "free") return;
      // Reunions normals → no generen forat
      // PERÒ: AA (Aula Acollida) i BIBLIOTECA/GUÀRDIA BIBLIOTECA → sí cal tractar-los
      const subj = (slot.subject || "").toUpperCase().replace(/À/g,"A").replace(/È/g,"E").replace(/Ï/g,"I");
      const isAA = subj === "AA" || subj.startsWith("AA ");
      const isBiblioteca = subj.includes("BIBL");
      const isMusica = subj.includes("MUSIC");

      if (slot.type === "meeting" && !isAA && !isBiblioteca && !isMusica) return;

      // Biblioteca/música: només en hora de pati generen forat; en altres hores → feina personal
      const PATI_TIMES = new Set(["10:45-11:15", "11:15-11:45"]);
      if ((isBiblioteca || isMusica) && !PATI_TIMES.has(time)) return;

      // Guàrdies normals (A, B) → no apareixen
      if (isGuardSlot(slot)) return;
      // Pati normal → no apareix
      if (isPatiSlot(slot)) return;

      const isClassOrTutoring = slot.type === "class" || slot.type === "tutoring";
      // AA sense grup → sempre genera forat (no pertany a cap grup de sortida)
      const withTrip = !isAA && isClassOrTutoring && slotWithTripGroups(slot.group, selectedGroups, excludedSubs);
      const isHalfExcluded = slotWithExcludedSub(slot.group, excludedSubs);
      const isHalfGroup = (halfGroups || []).some((g) => groupMatches(slot.group, g, []));
      const leavesGap = !withTrip || isHalfExcluded || isHalfGroup;

      if (leavesGap) {
        gapsBySlot[time].push({
          teacherName: teacher.name,
          subject: slot.subject,
          group: slot.group,
          room: slot.room,
          isHalf: isHalfExcluded || isHalfGroup,
        });
      }
    });
  });

  // ─────────────────────────────────────────────
  // PROFESSORS DISPONIBLES PER COBRIR
  // ─────────────────────────────────────────────
  stayTeachers.forEach((teacher) => {
    const daySlots = teacher.schedule[day] || [];

    rangeSlots.forEach((time) => {
      const slot = daySlots.find((s) => s.time === time);

      // hora lliure
      if (!slot || slot.type === "free") {
        coversBySlot[time].free.push(teacher.name);
        return;
      }

      // pati → no disponible
      if (isPatiSlot(slot)) return;

      // guàrdies
      if (isGuardSlot(slot)) {
        if (isLecturaGuard(slot)) {
          coversBySlot[time].guardL.push(teacher.name);
        } else if (isPatiEspecialSlot(slot)) {
          coversBySlot[time].guardE.push(teacher.name);
        } else if (guardType(slot) === "B") {
          coversBySlot[time].guardB.push(teacher.name);
        } else {
          coversBySlot[time].guardA.push(teacher.name);
        }
        return;
      }

      // si fa classe amb grup que marxa → queda alliberat
      if (slot.type === "class" || slot.type === "tutoring") {
        const withTrip = slotWithTripGroups(
          slot.group,
          selectedGroups,
          excludedSubs
        );

        const isHalfExcluded = slotWithExcludedSub(slot.group, excludedSubs);

        const isHalfGroup = (halfGroups || []).some((g) =>
          groupMatches(slot.group, g, [])
        );

        if (withTrip && !isHalfExcluded && !isHalfGroup) {
          coversBySlot[time].freed.push(teacher.name);
        }
      }
    });
  });

  return { gapsBySlot, coversBySlot, rangeSlots };
}

// ─── UI Primitives ────────────────────────────────────────────────────────────

function Spinner() {
  return (
    <div
      style={{
        width: 34,
        height: 34,
        borderRadius: "50%",
        border: "3px solid #e5e7eb",
        borderTopColor: "#e8451e",
        animation: "spin 0.8s linear infinite",
      }}
    />
  );
}
function Avatar({ name, color }) {
  const ini = name
    .split(" ")
    .map((w) => w[0])
    .join("")
    .slice(0, 2)
    .toUpperCase();
  return (
    <div
      style={{
        width: 42,
        height: 42,
        borderRadius: 10,
        background: color || "#1a2744",
        color: "white",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontSize: 14,
        fontWeight: 700,
        flexShrink: 0,
      }}
    >
      {ini}
    </div>
  );
}
function Pill({ color, children, small }) {
  const c = {
    green: { bg: "#dcfce7", fg: "#166534" },
    blue: { bg: "#dbeafe", fg: "#1e40af" },
    gray: { bg: "#f3f4f6", fg: "#4b5563" },
    red: { bg: "#fee2e2", fg: "#991b1b" },
    navy: { bg: "#e8edf5", fg: "#1a2744" },
    orange: { bg: "#fff7ed", fg: "#9a3412" },
    yellow: { bg: "#fef9c3", fg: "#854d0e" },
    purple: { bg: "#f3e8ff", fg: "#6b21a8" },
    teal: { bg: "#ccfbf1", fg: "#0f766e" },
  }[color] || { bg: "#f3f4f6", fg: "#4b5563" };
  return (
    <span
      style={{
        fontSize: small ? 10 : 12,
        padding: small ? "2px 7px" : "3px 10px",
        borderRadius: 99,
        backgroundColor: c.bg,
        color: c.fg,
        fontWeight: 500,
        whiteSpace: "nowrap",
      }}
    >
      {children}
    </span>
  );
}
function Card({ title, hint, children, accent }) {
  return (
    <div
      style={{
        backgroundColor: "white",
        borderRadius: 12,
        padding: "20px 22px",
        border: `1px solid ${accent ? "#e8451e" : "#e5e7eb"}`,
        marginBottom: 12,
      }}
    >
      {title && (
        <p
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: accent ? "#e8451e" : "#6b7280",
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            margin: `0 0 ${hint ? "4px" : "14px"}`,
          }}
        >
          {title}
        </p>
      )}
      {hint && (
        <p style={{ fontSize: 12, color: "#9ca3af", margin: "0 0 12px" }}>
          {hint}
        </p>
      )}
      {children}
    </div>
  );
}
function TimeRangePicker({ startSlot, endSlot, onChange }) {
  const si = MORNING_SLOTS.indexOf(startSlot),
    ei = MORNING_SLOTS.indexOf(endSlot);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          flexWrap: "wrap",
        }}
      >
        <span
          style={{
            fontSize: 12,
            fontWeight: 600,
            color: "#6b7280",
            minWidth: 48,
          }}
        >
          Inici
        </span>
        <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
          {MORNING_SLOTS.map((slot, i) => {
            const a = i === si,
              d = i > ei;
            return (
              <button
                key={slot}
                disabled={d}
                onClick={() => onChange(slot, endSlot)}
                style={{
                  padding: "5px 10px",
                  borderRadius: 6,
                  border: `1.5px solid ${a ? "#1a2744" : "#e5e7eb"}`,
                  backgroundColor: a ? "#1a2744" : "white",
                  color: a ? "white" : d ? "#d1d5db" : "#374151",
                  fontSize: 12,
                  fontWeight: 500,
                  cursor: d ? "not-allowed" : "pointer",
                }}
              >
                {slot.split("-")[0]}
              </button>
            );
          })}
        </div>
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          flexWrap: "wrap",
        }}
      >
        <span
          style={{
            fontSize: 12,
            fontWeight: 600,
            color: "#6b7280",
            minWidth: 48,
          }}
        >
          Final
        </span>
        <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
          {MORNING_SLOTS.map((slot, i) => {
            const a = i === ei,
              d = i < si;
            return (
              <button
                key={slot}
                disabled={d}
                onClick={() => onChange(startSlot, slot)}
                style={{
                  padding: "5px 10px",
                  borderRadius: 6,
                  border: `1.5px solid ${a ? "#e8451e" : "#e5e7eb"}`,
                  backgroundColor: a ? "#e8451e" : "white",
                  color: a ? "white" : d ? "#d1d5db" : "#374151",
                  fontSize: 12,
                  fontWeight: 500,
                  cursor: d ? "not-allowed" : "pointer",
                }}
              >
                {slot.split("-")[1]}
              </button>
            );
          })}
        </div>
      </div>
      <div style={{ display: "flex", gap: 3 }}>
        {MORNING_SLOTS.map((slot, i) => (
          <div
            key={slot}
            style={{
              flex: 1,
              height: 4,
              borderRadius: 2,
              backgroundColor: i >= si && i <= ei ? "#e8451e" : "#e5e7eb",
              transition: "background 0.15s",
            }}
          />
        ))}
      </div>
      <p style={{ fontSize: 11, color: "#9ca3af", margin: 0 }}>
        Sortida de <strong>{startSlot.split("-")[0]}</strong> a{" "}
        <strong>{endSlot.split("-")[1]}</strong> · {ei - si + 1} franges
      </p>
    </div>
  );
}

function GroupSelector({
  selected,
  onChange,
  teachers,
  excludedSubs,
  onExcludedSubs,
  halfGroups,
  onHalfGroups,
}) {
  const subgroupsPerGroup = useMemo(() => {
    const m = {};
    selected.forEach((g) => {
      m[g] = detectSubgroups(teachers || [], g);
    });
    return m;
  }, [selected, teachers]);
  const toggle = (v) =>
    onChange(
      selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v]
    );
  const toggleSub = (sub) =>
    onExcludedSubs(
      excludedSubs.includes(sub)
        ? excludedSubs.filter((x) => x !== sub)
        : [...excludedSubs, sub]
    );
  const toggleHalf = (g) =>
    onHalfGroups(
      (halfGroups || []).includes(g)
        ? (halfGroups || []).filter((x) => x !== g)
        : [...(halfGroups || []), g]
    );

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {YEAR_GROUPS.map((yr) => (
        <div
          key={yr.label}
          style={{ display: "flex", alignItems: "center", gap: 8 }}
        >
          <span
            style={{
              fontSize: 12,
              fontWeight: 600,
              color: "#6b7280",
              width: 56,
              flexShrink: 0,
            }}
          >
            {yr.label}
          </span>
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
            {yr.values.map(({ label, value }) => {
              const a = selected.includes(value);
              return (
                <button
                  key={value}
                  onClick={() => toggle(value)}
                  style={{
                    width: 34,
                    height: 30,
                    borderRadius: 6,
                    border: `1.5px solid ${a ? "#1a2744" : "#e5e7eb"}`,
                    backgroundColor: a ? "#1a2744" : "white",
                    color: a ? "white" : "#374151",
                    fontSize: 13,
                    fontWeight: 600,
                    cursor: "pointer",
                  }}
                >
                  {label}
                </button>
              );
            })}
          </div>
        </div>
      ))}

      {/* Per each selected group: subgroup toggles OR half-group toggle */}
      {selected.length > 0 && (
        <div
          style={{
            marginTop: 4,
            display: "flex",
            flexDirection: "column",
            gap: 8,
          }}
        >
          {selected.map((g) => {
            const subs = subgroupsPerGroup[g] || [];
            const isHalf = (halfGroups || []).includes(g);
            return (
              <div
                key={g}
                style={{
                  padding: "10px 14px",
                  backgroundColor: "#f8fafc",
                  borderRadius: 8,
                  border: "1px solid #e2e8f0",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    marginBottom: subs.length > 0 ? 8 : 0,
                  }}
                >
                  <span
                    style={{ fontSize: 12, fontWeight: 700, color: "#1a2744" }}
                  >
                    {g}
                  </span>
                  {subs.length === 0 && (
                    <button
                      onClick={() => toggleHalf(g)}
                      style={{
                        padding: "3px 10px",
                        borderRadius: 6,
                        fontSize: 11,
                        fontWeight: 500,
                        cursor: "pointer",
                        border: `1.5px solid ${isHalf ? "#9a3412" : "#e5e7eb"}`,
                        backgroundColor: isHalf ? "#fff7ed" : "white",
                        color: isHalf ? "#9a3412" : "#6b7280",
                      }}
                    >
                      {isHalf
                        ? "½ Mig grup va de sortida"
                        : "Tot el grup va de sortida"}
                    </button>
                  )}
                </div>
                {subs.length > 0 && (
                  <>
                    <p
                      style={{
                        fontSize: 11,
                        color: "#854d0e",
                        fontWeight: 600,
                        margin: "0 0 6px",
                      }}
                    >
                      ⚠ Grups partits — indica quins van:
                    </p>
                    <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                      {subs.map((sub) => {
                        const goes = !excludedSubs.includes(sub);
                        return (
                          <button
                            key={sub}
                            onClick={() => toggleSub(sub)}
                            style={{
                              padding: "3px 10px",
                              borderRadius: 6,
                              fontSize: 12,
                              fontWeight: 500,
                              cursor: "pointer",
                              border: `1.5px solid ${
                                goes ? "#166534" : "#9ca3af"
                              }`,
                              backgroundColor: goes ? "#dcfce7" : "#f3f4f6",
                              color: goes ? "#166534" : "#6b7280",
                            }}
                          >
                            {goes ? "🚌" : "🏫"} {sub}
                          </button>
                        );
                      })}
                    </div>
                  </>
                )}
              </div>
            );
          })}
        </div>
      )}

      {selected.length > 0 && (
        <button
          onClick={() => {
            onChange([]);
            onExcludedSubs([]);
            onHalfGroups([]);
          }}
          style={{
            alignSelf: "flex-start",
            background: "none",
            border: "none",
            color: "#9ca3af",
            fontSize: 12,
            cursor: "pointer",
            textDecoration: "underline",
            marginTop: 2,
          }}
        >
          Netejar selecció
        </button>
      )}
    </div>
  );
}

function SubjectSelector({ subjects, selected, onChange }) {
  const [custom, setCustom] = useState("");
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {subjects.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {subjects.map((s) => {
            const a = selected === s;
            return (
              <button
                key={s}
                onClick={() => onChange(a ? "" : s)}
                style={{
                  padding: "5px 12px",
                  borderRadius: 99,
                  border: `1.5px solid ${a ? "#1a2744" : "#e5e7eb"}`,
                  backgroundColor: a ? "#1a2744" : "white",
                  color: a ? "white" : "#374151",
                  fontSize: 12,
                  fontWeight: 500,
                  cursor: "pointer",
                }}
              >
                {a && "✓ "}
                {s}
              </button>
            );
          })}
        </div>
      )}
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <input
          style={{
            flex: 1,
            padding: "8px 12px",
            border: "1.5px solid #e5e7eb",
            borderRadius: 8,
            fontSize: 13,
            color: "#1a1a2e",
          }}
          placeholder="Escriu una matèria (Enter)"
          value={custom}
          onChange={(e) => setCustom(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && custom.trim()) {
              onChange(custom.trim());
              setCustom("");
            }
          }}
        />
        {selected && (
          <button
            onClick={() => onChange("")}
            style={{
              background: "none",
              border: "none",
              color: "#9ca3af",
              fontSize: 12,
              cursor: "pointer",
              textDecoration: "underline",
            }}
          >
            Treure
          </button>
        )}
      </div>
      {selected && <Pill color="navy">★ {selected}</Pill>}
    </div>
  );
}

// ─── Helpers de format de noms ────────────────────────────────────────────────
// "Casas Molist, Marta" → "Marta Casas Molist"
function nomComplet(nom) {
  if (!nom) return nom;
  const parts = nom.split(",").map(s => s.trim());
  if (parts.length >= 2) return `${parts[1]} ${parts[0]}`;
  return nom;
}
// "Casas Molist, Marta" → "Marta C."
function nomAbreujat(nom) {
  if (!nom) return nom;
  const complet = nomComplet(nom);
  const words = complet.trim().split(/\s+/);
  if (words.length < 2) return complet;
  return `${words[0]} ${words[1][0]}.`;
}

// ─── Coverage Panel ───────────────────────────────────────────────────────────

// assignments: { "time|gapIdx": teacherName }
function CoveragePanel({
  teachers,
  confirmedNames,
  trip,
  franges,
  teacherFranjaMap,
  assignments,
  onAssign,
}) {
  const { gapsBySlot, coversBySlot, rangeSlots } = useMemo(
    () => computeCoverage(teachers, confirmedNames, trip, franges, teacherFranjaMap),
    [teachers, confirmedNames, trip, franges, teacherFranjaMap]
  );
  const [showFreeSlots, setShowFreeSlots] = useState({});
  const hasAnyGap = rangeSlots.some((t) => gapsBySlot[t]?.length > 0);

  if (!hasAnyGap)
    return (
      <div
        style={{
          padding: "14px 18px",
          backgroundColor: "#f0fdf4",
          borderRadius: 10,
          border: "1px solid #bbf7d0",
        }}
      >
        <p
          style={{ fontSize: 13, color: "#166534", margin: 0, fontWeight: 600 }}
        >
          ✓ Cap forat. Els professors confirmats no deixen classes sense cobrir.
        </p>
      </div>
    );

  // Which teachers are already assigned in a given time slot (can't double-assign)
  function assignedInSlot(time) {
    return Object.entries(assignments)
      .filter(([k]) => k.startsWith(time + "|"))
      .map(([, v]) => v)
      .filter(Boolean);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {rangeSlots.map((time) => {
        const gaps = gapsBySlot[time] || [];
        const covers = coversBySlot[time] || {
          freed: [],
          guardL: [],
          guardE: [],
          guardA: [],
          guardB: [],
          free: [],
        };
        if (!gaps.length)
          return (
            <div
              key={time}
              style={{
                padding: "8px 14px",
                backgroundColor: "#f0fdf4",
                borderRadius: 8,
                border: "1px solid #bbf7d0",
                marginBottom: 4,
                display: "flex",
                gap: 10,
                alignItems: "center",
              }}
            >
              <span
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: "#166534",
                  minWidth: 70,
                }}
              >
                {SLOT_LABEL[time]}
              </span>
              <span style={{ fontSize: 11, opacity: 0.65 }}>{time}</span>
              <span
                style={{ fontSize: 11, color: "#166534", marginLeft: "auto" }}
              >
                ✓ 0 forats per cobrir
              </span>
            </div>
          );

        const assignedHere = assignedInSlot(time);
        const allAssigned = gaps.every((_, gi) => assignments[`${time}|${gi}`]);
        const borderColor = allAssigned ? "#bbf7d0" : "#fecaca";
        const bgColor = allAssigned ? "#f0fdf4" : "#fef2f2";
        const showFree = showFreeSlots[time] || false;

        const totalAvail =
          covers.freed.length + covers.guardA.length + covers.guardB.length;

        return (
          <div
            key={time}
            style={{
              borderRadius: 10,
              border: `1.5px solid ${borderColor}`,
              backgroundColor: bgColor,
              overflow: "hidden",
            }}
          >
            {/* Header */}
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: "9px 14px",
                borderBottom: `1px solid ${borderColor}`,
                backgroundColor: allAssigned ? "#dcfce7" : "#fee2e2",
              }}
            >
              <span
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: allAssigned ? "#166534" : "#991b1b",
                  minWidth: 70,
                }}
              >
                {SLOT_LABEL[time]}
              </span>
              <span style={{ fontSize: 11, opacity: 0.65 }}>{time}</span>
              <span
                style={{
                  marginLeft: "auto",
                  fontSize: 11,
                  fontWeight: 700,
                  color: allAssigned ? "#166534" : "#991b1b",
                }}
              >
                {allAssigned
                  ? "✓ Cobert"
                  : `⚠ ${
                      gaps.filter((_, gi) => !assignments[`${time}|${gi}`])
                        .length
                    } forat${gaps.length > 1 ? "s" : ""} per cobrir`}
              </span>
            </div>

            <div
              style={{
                padding: "12px 14px",
                display: "flex",
                gap: 16,
                flexWrap: "wrap",
              }}
            >
              {/* Gaps with assignment UI */}
              <div style={{ flex: "1 1 200px" }}>
                <p
                  style={{
                    fontSize: 10,
                    fontWeight: 700,
                    color: "#991b1b",
                    textTransform: "uppercase",
                    letterSpacing: "0.06em",
                    margin: "0 0 8px",
                  }}
                >
                  Classes sense professor
                </p>
                {gaps.map((g, gi) => {
                  const key = `${time}|${gi}`;
                  const assigned = assignments[key];
                  return (
                    <div
                      key={gi}
                      style={{
                        marginBottom: 10,
                        padding: "8px 10px",
                        borderRadius: 8,
                        backgroundColor: assigned ? "#f0fdf4" : "white",
                        border: `1px solid ${assigned ? "#bbf7d0" : "#e5e7eb"}`,
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: assigned ? 4 : 6, flexWrap: "wrap" }}>
                        <span style={{ fontSize: 11, padding: "2px 7px", borderRadius: 4, backgroundColor: g.isHalf ? "#fff7ed" : "#fee2e2", color: g.isHalf ? "#9a3412" : "#991b1b", fontWeight: 600 }}>
                          {g.isHalf ? "½ " : ""}{g.subject || "—"}
                        </span>
                        {g.group && <span style={{ fontSize: 11, color: "#374151", fontWeight: 500 }}>{g.group}</span>}
                        {g.room && <span style={{ fontSize: 10, color: "#9ca3af" }}>· {g.room}</span>}
                        <span style={{ fontSize: 10, color: "#9ca3af", marginLeft: "auto" }}>
                          {nomAbreujat(g.teacherName) || g.teacherName}
                        </span>
                      </div>
                      {assigned ? (
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{
                            fontSize: 11, padding: "2px 8px", borderRadius: 5, fontWeight: 600,
                            backgroundColor: assigned === "NO_CAL_COBRIR" ? "#f0fdf4" : assigned === "PROF. DE GUÀRDIA" ? "#fef9c3" : "#166534",
                            color: assigned === "NO_CAL_COBRIR" ? "#166534" : assigned === "PROF. DE GUÀRDIA" ? "#854d0e" : "white",
                          }}>
                            {assigned === "NO_CAL_COBRIR" ? "✓ No cal cobrir" : assigned === "PROF. DE GUÀRDIA" ? assigned : (nomAbreujat(assigned) || assigned)}
                          </span>
                          <button onClick={() => onAssign(key, null)} style={{ fontSize: 10, padding: "1px 7px", borderRadius: 4, border: "1px solid #d1d5db", backgroundColor: "white", color: "#6b7280", cursor: "pointer" }}>
                            ✕ Desassignar
                          </button>
                        </div>
                      ) : (
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <p style={{ fontSize: 10, color: "#9ca3af", margin: 0, fontStyle: "italic", flex: 1 }}>↓ Clica un professor per assignar</p>
                          <button
                            onClick={() => onAssign(key, "PROF. DE GUÀRDIA")}
                            style={{ fontSize: 10, padding: "2px 8px", borderRadius: 4, border: "1px solid #d1d5db", backgroundColor: "#fef9c3", color: "#854d0e", cursor: "pointer", fontWeight: 600, whiteSpace: "nowrap" }}
                          >
                            + Prof. de guàrdia
                          </button>
                          <button
                            onClick={() => onAssign(key, "NO_CAL_COBRIR")}
                            style={{ fontSize: 10, padding: "2px 8px", borderRadius: 4, border: "1px solid #bbf7d0", backgroundColor: "#f0fdf4", color: "#166534", cursor: "pointer", fontWeight: 600, whiteSpace: "nowrap" }}
                          >
                            ✓ No cal cobrir
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Available teachers */}
              <div style={{ flex: "2 1 260px" }}>
                <p
                  style={{
                    fontSize: 10,
                    fontWeight: 700,
                    color: "#166534",
                    textTransform: "uppercase",
                    letterSpacing: "0.06em",
                    margin: "0 0 8px",
                  }}
                >
                  Professors disponibles
                </p>

                {/* Freed */}
                {covers.freed.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <span
                      style={{
                        fontSize: 10,
                        color: "#0f766e",
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: "0.05em",
                      }}
                    >
                      🟢 Alliberats per la sortida
                    </span>
                    <p
                      style={{
                        fontSize: 10,
                        color: "#9ca3af",
                        margin: "1px 0 5px",
                      }}
                    >
                      Tenien classe amb alumnes que han marxat
                    </p>
                    <TeacherButtons
                      teachers={covers.freed}
                      time={time}
                      gaps={gaps}
                      assignments={assignments}
                      assignedHere={assignedHere}
                      onAssign={onAssign}
                      colorBg="#ccfbf1"
                      colorFg="#0f766e"
                      colorBorder="#99f6e4"
                    />
                  </div>
                )}
                {covers.guardL?.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <span
                      style={{
                        fontSize: 10,
                        color: "#0369a1",
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: "0.05em",
                      }}
                    >
                      📖 Guàrdia de lectura
                    </span>
                    <p
                      style={{
                        fontSize: 10,
                        color: "#9ca3af",
                        margin: "1px 0 5px",
                      }}
                    >
                      Guàrdia assignada durant la lectura
                    </p>
                    <TeacherButtons
                      teachers={covers.guardL}
                      time={time}
                      gaps={gaps}
                      assignments={assignments}
                      assignedHere={assignedHere}
                      onAssign={onAssign}
                      colorBg="#e0f2fe"
                      colorFg="#0369a1"
                      colorBorder="#7dd3fc"
                    />
                  </div>
                )}
                {/* Guard A */}
                {covers.guardA.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <span
                      style={{
                        fontSize: 10,
                        color: "#854d0e",
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: "0.05em",
                      }}
                    >
                      🟡 Guàrdia A
                    </span>
                    <p
                      style={{
                        fontSize: 10,
                        color: "#9ca3af",
                        margin: "1px 0 5px",
                      }}
                    >
                      Guàrdia assignada (A)
                    </p>
                    <TeacherButtons
                      teachers={covers.guardA}
                      time={time}
                      gaps={gaps}
                      assignments={assignments}
                      assignedHere={assignedHere}
                      onAssign={onAssign}
                      colorBg="#fef9c3"
                      colorFg="#854d0e"
                      colorBorder="#fde68a"
                    />
                  </div>
                )}

                {/* Guard B */}
                {covers.guardB.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <span
                      style={{
                        fontSize: 10,
                        color: "#92400e",
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: "0.05em",
                      }}
                    >
                      🟠 Guàrdia B
                    </span>
                    <p
                      style={{
                        fontSize: 10,
                        color: "#9ca3af",
                        margin: "1px 0 5px",
                      }}
                    >
                      Guàrdia assignada (B)
                    </p>
                    <TeacherButtons
                      teachers={covers.guardB}
                      time={time}
                      gaps={gaps}
                      assignments={assignments}
                      assignedHere={assignedHere}
                      onAssign={onAssign}
                      colorBg="#ffedd5"
                      colorFg="#92400e"
                      colorBorder="#fed7aa"
                    />
                  </div>
                )}
                {/* Guard E */}
                {covers.guardE?.length > 0 && (
                  <div style={{ marginBottom: 10 }}>
                    <span
                      style={{
                        fontSize: 10,
                        color: "#7c3aed",
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: "0.05em",
                      }}
                    >
                      🎵 Guàrdia Pati Especial
                    </span>

                    <p
                      style={{
                        fontSize: 10,
                        color: "#9ca3af",
                        margin: "1px 0 5px",
                      }}
                    >
                      Música / Biblioteca
                    </p>

                    <TeacherButtons
                      teachers={covers.guardE}
                      time={time}
                      gaps={gaps}
                      assignments={assignments}
                      assignedHere={assignedHere}
                      onAssign={onAssign}
                      colorBg="#ede9fe"
                      colorFg="#7c3aed"
                      colorBorder="#c4b5fd"
                    />
                  </div>
                )}

                {/* Free — hidden by default, toggle */}
                {covers.free.length > 0 && (
                  <div>
                    <button
                      onClick={() =>
                        setShowFreeSlots((p) => ({ ...p, [time]: !p[time] }))
                      }
                      style={{
                        fontSize: 10,
                        fontWeight: 600,
                        color: "#6b7280",
                        background: "none",
                        border: "1px solid #e5e7eb",
                        borderRadius: 6,
                        padding: "3px 10px",
                        cursor: "pointer",
                        marginBottom: showFree ? 6 : 0,
                      }}
                    >
                      {showFree ? "▲ Amagar" : "▼ Mostrar"} hores lliures (
                      {covers.free.length})
                    </button>
                    {showFree && (
                      <>
                        <p
                          style={{
                            fontSize: 10,
                            color: "#9ca3af",
                            margin: "4px 0 5px",
                          }}
                        >
                          No tenen classe ni guàrdia programada
                        </p>
                        <TeacherButtons
                          teachers={covers.free}
                          time={time}
                          gaps={gaps}
                          assignments={assignments}
                          assignedHere={assignedHere}
                          onAssign={onAssign}
                          colorBg="#f3f4f6"
                          colorFg="#4b5563"
                          colorBorder="#d1d5db"
                        />
                      </>
                    )}
                  </div>
                )}

                {totalAvail === 0 && covers.free.length === 0 && (
                  <span
                    style={{
                      fontSize: 12,
                      color: "#991b1b",
                      fontStyle: "italic",
                      fontWeight: 600,
                    }}
                  >
                    ⚠ Cap professor disponible
                  </span>
                )}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// Clickable teacher buttons for assignment
function TeacherButtons({
  teachers,
  time,
  gaps,
  assignments,
  assignedHere,
  onAssign,
  colorBg,
  colorFg,
  colorBorder,
}) {
  // Find first unassigned gap to assign to when clicked
  function handleClick(name) {
    // If already assigned somewhere in this slot, unassign
    const existingKey = Object.keys(assignments).find(
      (k) => k.startsWith(time + "|") && assignments[k] === name
    );
    if (existingKey) {
      onAssign(existingKey, null);
      return;
    }
    // Find first gap not yet assigned
    const firstFree = gaps.findIndex((_, gi) => !assignments[`${time}|${gi}`]);
    if (firstFree === -1) return; // all gaps covered
    onAssign(`${time}|${firstFree}`, name);
  }

  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
      {teachers.map((n) => {
        const isAssignedHere = Object.keys(assignments).some(
          (k) => k.startsWith(time + "|") && assignments[k] === n
        );
        const isUsedElsewhere = assignedHere.includes(n) && !isAssignedHere;
        return (
          <button
            key={n}
            onClick={() => !isUsedElsewhere && handleClick(n)}
            disabled={isUsedElsewhere}
            title={
              isAssignedHere
                ? "Clic per desassignar"
                : isUsedElsewhere
                ? "Ja assignat a un altre grup"
                : "Clic per assignar"
            }
            style={{
              fontSize: 11,
              padding: "3px 9px",
              borderRadius: 5,
              cursor: isUsedElsewhere ? "not-allowed" : "pointer",
              fontWeight: isAssignedHere ? 700 : 500,
              border: `1.5px solid ${
                isAssignedHere
                  ? "#166534"
                  : isUsedElsewhere
                  ? "#e5e7eb"
                  : colorBorder
              }`,
              backgroundColor: isAssignedHere
                ? "#166534"
                : isUsedElsewhere
                ? "#f9fafb"
                : colorBg,
              color: isAssignedHere
                ? "white"
                : isUsedElsewhere
                ? "#d1d5db"
                : colorFg,
              textDecoration: isAssignedHere ? "none" : "none",
              transition: "all 0.1s",
            }}
          >
            {isAssignedHere ? "✓ " : ""}
            {n}
          </button>
        );
      })}
    </div>
  );
}

// ─── Ranking Card ─────────────────────────────────────────────────────────────

function RankingCard({ r, i, trip, confirmed, onToggleConfirm, teacherFranjaMap, onAssignFranja }) {
  const isConfirmed = confirmed.has(r.name);
  const medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : null;
  const tripTimes = new Set(r.tripClasses.map((c) => c.time));
  const multiTorn = trip.franges && trip.franges.length > 1;

  return (
    <div
      style={{
        display: "flex",
        alignItems: "flex-start",
        gap: 14,
        backgroundColor: isConfirmed ? "#f0fdf4" : "white",
        borderRadius: 12,
        padding: "15px 17px",
        border: isConfirmed
          ? "2px solid #166534"
          : i === 0
          ? "2px solid #e8451e"
          : "1px solid #e5e7eb",
        animation: "fadeUp 0.2s ease",
        transition: "all 0.2s",
      }}
    >
      <div
        style={{ width: 30, textAlign: "center", flexShrink: 0, paddingTop: 2 }}
      >
        {medal ? (
          <span style={{ fontSize: 20 }}>{medal}</span>
        ) : (
          <span
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: "#9ca3af",
              fontFamily: "monospace",
            }}
          >
            #{i + 1}
          </span>
        )}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <p
          style={{
            fontSize: 15,
            fontWeight: 700,
            color: "#1a2744",
            margin: "0 0 6px",
          }}
        >
          {r.name}
        </p>
        <div
          style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 8 }}
        >
          {r.teachesGroup && trip.franges.some(f => f.selectedGroups.length > 0) && (
            <Pill color="teal">👥 Coneix el grup</Pill>
          )}
          {r.coversFullRange && (
            <Pill color="green">✓ Cobreix tot el rang</Pill>
          )}
          {r.spanLabel && !r.coversFullRange && (
            <Pill color="gray">🕐 {r.spanLabel}</Pill>
          )}
          {r.spanLabel && r.coversFullRange && (
            <Pill color="gray">🕐 {r.spanLabel}</Pill>
          )}
          {r.tripClasses.length > 0 && (
            <Pill color="navy">
              📚{" "}
              {r.coveredSlots === Math.floor(r.coveredSlots)
                ? r.coveredSlots
                : `${r.coveredSlots}`}
              /{r.totalRange} h. amb el grup
            </Pill>
          )}
          {r.subjectMatch && trip.subject && (
            <Pill color="blue">★ {trip.subject}</Pill>
          )}
          {r.stayClasses.length > 0 && (
            <Pill color="red">
              ✗ {r.stayClasses.length} classe
              {r.stayClasses.length > 1 ? "s" : ""} sense cobrir
            </Pill>
          )}
          {r.halfClasses.length > 0 && (
            <Pill color="orange">½ {r.halfClasses.length} mig grup</Pill>
          )}
          {r.guardSlots.length > 0 && (
            <Pill color="yellow">
              G {r.guardSlots.length} guàrd
              {r.guardSlots.length > 1 ? "ies" : "ia"}
            </Pill>
          )}
          {r.freeCount > 0 && (
            <Pill color="gray">
              ◯ {r.freeCount} h. lliure{r.freeCount > 1 ? "s" : ""}
            </Pill>
          )}
        </div>
        {/* Slot timeline */}
        <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
          {r.inRange.map((slot, j) => {
            let bg,
              fg,
              label,
              prefix = "";
            if (slot.type === "free") {
              bg = "#f3f4f6";
              fg = "#9ca3af";
              label = "Lliure";
            } else if (isGuardSlot(slot)) {
              bg = "#fef9c3";
              fg = "#854d0e";
              label = slot.subject || "Guàrdia";
            } else if (slot.type === "meeting") {
              bg = "#f3e8ff";
              fg = "#6b21a8";
              label = slot.subject || "Reunió";
            } else if (tripTimes.has(slot.time)) {
              bg = "#dcfce7";
              fg = "#166534";
              label = slot.subject;
            } else if (r.halfClasses.some((c) => c.time === slot.time)) {
              bg = "#fff7ed";
              fg = "#9a3412";
              label = slot.subject;
              prefix = "½ ";
            } else {
              bg = "#fee2e2";
              fg = "#991b1b";
              label = slot.subject;
            }
            const sl =
              (label || "").length > 13
                ? (label || "").slice(0, 12) + "…"
                : label || "";
            return (
              <span
                key={j}
                title={`${SLOT_LABEL[slot.time] || slot.time}: ${label}${
                  slot.group ? " · " + slot.group : ""
                }`}
                style={{
                  fontSize: 10,
                  padding: "2px 7px",
                  borderRadius: 4,
                  backgroundColor: bg,
                  color: fg,
                  fontWeight: 500,
                }}
              >
                <span style={{ opacity: 0.5, marginRight: 2, fontSize: 9 }}>
                  {slot.time.split("-")[0]}
                </span>
                {prefix}
                {sl}
                {slot.group && slot.type !== "free" && (
                  <span style={{ opacity: 0.55 }}> · {slot.group}</span>
                )}
              </span>
            );
          })}
        </div>
        {r.freeCount > 0 && (
          <p
            style={{
              fontSize: 10,
              color: "#9ca3af",
              margin: "5px 0 0",
              fontStyle: "italic",
            }}
          >
            ◯ Hora lliure = cap activitat programada (ni classe, ni guàrdia, ni
            reunió)
          </p>
        )}
      </div>
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          gap: 8,
          flexShrink: 0,
        }}
      >
        <div style={{ textAlign: "center" }}>
          <span
            style={{
              display: "block",
              fontSize: 26,
              fontWeight: 800,
              color: isConfirmed ? "#166534" : i === 0 ? "#e8451e" : "#1a2744",
              lineHeight: 1,
              fontFamily: "monospace",
            }}
          >
            {r.score}
          </span>
          <span
            style={{
              display: "block",
              fontSize: 9,
              color: "#9ca3af",
              textTransform: "uppercase",
              letterSpacing: "0.07em",
              marginTop: 2,
            }}
          >
            pts
          </span>
        </div>
        <button
          onClick={() => {
            if (multiTorn) return; // en mode multi-torn, els botons de torn fan la feina
            onToggleConfirm(r.name);
          }}
          style={{
            padding: "5px 10px",
            borderRadius: 7,
            fontSize: 11,
            fontWeight: 600,
            cursor: multiTorn ? "default" : "pointer",
            border: `1.5px solid ${isConfirmed ? "#166534" : "#d1d5db"}`,
            backgroundColor: isConfirmed ? "#166534" : "white",
            color: isConfirmed ? "white" : "#6b7280",
            whiteSpace: "nowrap",
            display: multiTorn ? "none" : "block",
          }}
        >
          {isConfirmed ? "✓ Confirmat" : "+ Confirmar"}
        </button>
        {multiTorn && (
          <div style={{ display: "flex", flexDirection: "column", gap: 4, minWidth: 90 }}>
            {trip.franges.map((franja, fi) => {
              const assignedFranjaIds = teacherFranjaMap[r.name] || [];
              const isAssignedHere = assignedFranjaIds.includes(franja.id);
              // Comprovar solapament amb torns JA assignats (excloent aquest mateix)
              const otherAssignedFranjes = trip.franges.filter(f =>
                f.id !== franja.id && assignedFranjaIds.includes(f.id)
              );
              const slotsA = new Set(slotsInRange(franja.startSlot, franja.endSlot));
              const overlappingTorn = otherAssignedFranjes.find(f =>
                slotsInRange(f.startSlot, f.endSlot).some(s => slotsA.has(s))
              );
              return (
                <button
                  key={franja.id}
                  onClick={() => {
                    if (isAssignedHere) {
                      // Desassignar d'aquest torn
                      const newIds = assignedFranjaIds.filter(id => id !== franja.id);
                      onAssignFranja(r.name, newIds);
                      if (newIds.length === 0) onToggleConfirm(r.name);
                    } else {
                      if (overlappingTorn) {
                        const tornIdx = trip.franges.findIndex(f => f.id === overlappingTorn.id) + 1;
                        alert(`⚠ Atenció: ${r.name} ja té assignat el Torn ${tornIdx} que se solapa en horari amb aquest torn.`);
                        return;
                      }
                      if (!isConfirmed) onToggleConfirm(r.name);
                      onAssignFranja(r.name, [...assignedFranjaIds, franja.id]);
                    }
                  }}
                  style={{
                    padding: "4px 8px",
                    borderRadius: 6,
                    fontSize: 10,
                    fontWeight: 600,
                    cursor: "pointer",
                    border: `1.5px solid ${isAssignedHere ? "#166534" : overlappingTorn ? "#991b1b" : "#d1d5db"}`,
                    backgroundColor: isAssignedHere ? "#166534" : overlappingTorn ? "#fee2e2" : "white",
                    color: isAssignedHere ? "white" : overlappingTorn ? "#991b1b" : "#6b7280",
                    whiteSpace: "nowrap",
                    textAlign: "center",
                  }}
                >
                  {isAssignedHere ? `✓ Torn ${fi + 1}` : overlappingTorn ? `⚠ Torn ${fi + 1}` : `+ Torn ${fi + 1}`}
                </button>
              );
            })}
            {isConfirmed && (
              <button
                onClick={() => {
                  onToggleConfirm(r.name);
                  onAssignFranja(r.name, []);
                }}
                style={{
                  padding: "3px 8px", borderRadius: 6, fontSize: 10,
                  fontWeight: 600, cursor: "pointer",
                  border: "1px solid #e5e7eb", backgroundColor: "transparent",
                  color: "#9ca3af",
                }}
              >
                ✕ Treure
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
async function iEducaScraperFn() {
  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  const ui = document.createElement("div");
  ui.style.cssText = `
    position:fixed;bottom:24px;right:24px;z-index:999999;
    background:linear-gradient(135deg,#1a2744,#243a6b);
    color:white;padding:18px 22px;border-radius:14px;
    font-family:system-ui,sans-serif;font-size:13px;
    min-width:320px;box-shadow:0 10px 30px rgba(0,0,0,0.4);
    border-left:4px solid #e8451e;
    opacity:0;transform:translateY(10px);
    transition:all .3s ease;
  `;
  document.body.appendChild(ui);
  setTimeout(() => {
    ui.style.opacity = 1;
    ui.style.transform = "translateY(0)";
  }, 10);

  let i = 0,
    total = 0,
    errors = [];

  function setStatus(msg, pct) {
    ui.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center">
        <div style="font-weight:700;color:#e8451e">📊 iEduca Scraper</div>
        <div style="cursor:pointer;font-size:14px" onclick="this.parentElement.parentElement.remove()">❌</div>
      </div>
      <div style="margin-top:8px">${msg}</div>
      ${
        pct !== undefined
          ? `
        <div style="margin-top:12px;background:rgba(255,255,255,0.15);height:6px;border-radius:4px;overflow:hidden">
          <div style="width:${pct}%;background:#e8451e;height:6px;border-radius:4px;transition:width .3s ease"></div>
        </div>
        <div style="display:flex;justify-content:space-between;margin-top:6px;font-size:11px;color:#ccc">
          <span>👥 ${i + 1} / ${total}</span>
          <span>${pct}%</span>
        </div>
      `
          : ""
      }
      ${
        errors.length
          ? `<div style="margin-top:6px;font-size:11px;color:#ffb3b3">⚠️ ${errors.length} errors</div>`
          : ""
      }
    `;
  }

  function removeUI() {
    ui.style.opacity = 0;
    ui.style.transform = "translateY(10px)";
    setTimeout(() => ui.remove(), 500);
  }

  const teacherSelect = [...document.querySelectorAll("select")].find((s) =>
    [...s.options].some((o) => o.value.includes("professor="))
  );

  if (!teacherSelect) {
    setStatus("🚫 No trobat selector de professors");
    setTimeout(removeUI, 3000);
    return;
  }

  const teacherOptions = [...teacherSelect.options]
    .filter((o) => o.value.includes("professor="))
    .map((o) => ({ name: o.text.trim(), url: location.origin + o.value }));

  total = teacherOptions.length;
  setStatus(`🔍 Trobats ${total} professors...`);
  await sleep(800);

  function cleanTeacherName(text) {
    return text.replace(/[A-ZÀ-Ú][a-zà-ú]+\s[A-ZÀ-Ú][a-zà-ú]+$/, "").trim();
  }

  function isMeeting(full, subject) {
    return (
      full.includes("REUNIÓ") ||
      full.includes("REUNIO") ||
      full.includes("EQUIP") ||
      full.includes("CLAUSTRE") ||
      full.includes("DEPARTAMENT") ||
      full.includes("DEP") ||
      /\bR\./.test(full) ||
      full.includes("DIRECCIO") ||
      full.includes("DIRECCIÓ") ||
      full.includes("COORD") ||
      full.includes("CONSELL") ||
      full.includes("PEDAGÒGIC") ||
      full.includes("PEDAGOGIC") ||
      full.includes("STEAM") ||
      full.includes("DIGITAL") ||
      full.includes("BIBLIOTECA") ||
      /^[A-ZÀÈÉÍÏÒÓÚÜ]{2,6}\d*$/.test(subject.trim())
    );
  }

  function parseCell(cell, time) {
    const empty = { time, subject: "", group: "", room: "", type: "free" };
    if (!cell) return empty;

    const dts = cell.querySelector('[id^="dts_"]');
    if (dts) {
      const rawText = dts.innerText
        .split("\n")
        .map((t) => t.trim())
        .filter(Boolean);
      if (!rawText.length) return empty;

      const subject = (
        dts.querySelector("strong")?.textContent ||
        rawText[0] ||
        ""
      ).trim();
      const lines = dts.innerHTML
        .replace(/<strong[^>]*>.*?<\/strong>/i, "")
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<[^>]+>/g, "")
        .split("\n")
        .map((t) => t.trim())
        .filter(Boolean);

      const room = lines[1] || "";
      const group = lines[2] || "";
      let type = "class";
      if (subject.toUpperCase().includes("TUTORIA")) {
        type = group ? "tutoring" : "meeting";
      }
      return { time, subject, group, room, type };
    }

    const tooltip = cell.querySelector(".tooltip_sortida");
    const rawText = (tooltip || cell).innerText
      .split("\n")
      .map((t) => t.trim())
      .filter(Boolean);
    if (!rawText.length) return empty;

    const full = rawText.join(" ").toUpperCase();
    const subject = (
      (tooltip || cell).querySelector("strong")?.textContent ||
      rawText[0] ||
      ""
    ).trim();

    if (/\bCAP\b/.test(full) || full.includes("TUT."))
      return { time, subject, group: "", room: "", type: "personal" };
    if (isMeeting(full, subject))
      return { time, subject, group: "", room: "", type: "meeting" };
    if (full.includes("GUÀRDIA") || full.includes("GUARDIA"))
      return {
        time,
        subject: cleanTeacherName(rawText.join(" ")),
        group: "",
        room: "",
        type: "guard",
      };
    if (full.includes("PATI"))
      return { time, subject: "Pati", group: "", room: "", type: "break" };

    return empty;
  }

  function parseHTML(html, teacherName) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (!table) return null;

    const DAYS = ["DILLUNS", "DIMARTS", "DIMECRES", "DIJOUS", "DIVENDRES"];
    const schedule = {};
    DAYS.forEach((d) => (schedule[d] = []));

    const rows = [...table.querySelectorAll("tr")];
    let headerIdx = 0;
    for (let r = 0; r < rows.length; r++) {
      const t = rows[r].textContent.toLowerCase();
      if (t.includes("dl") || t.includes("dilluns")) {
        headerIdx = r;
        break;
      }
    }

    for (let ri = headerIdx + 1; ri < rows.length; ri++) {
      const cells = [...rows[ri].querySelectorAll("td,th")];
      if (cells.length < 2) continue;
      const m = cells[0].textContent.match(/(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})/);
      if (!m) continue;
      const timeSlot = `${m[1]}-${m[2]}`;
      for (let di = 0; di < 5; di++) {
        schedule[DAYS[di]].push(parseCell(cells[di + 1] || null, timeSlot));
      }
    }

    return { name: teacherName.trim().replace(/\s*,\s*/, ", "), schedule };
  }

  const allTeachers = [];
  for (i = 0; i < teacherOptions.length; i++) {
    const t = teacherOptions[i];
    setStatus(`🔄 Descarregant ${t.name}`, Math.round(((i + 1) / total) * 100));
    try {
      const res = await fetch(t.url, { credentials: "include" });
      const html = await res.text();
      const data = parseHTML(html, t.name);
      if (data) allTeachers.push(data);
    } catch (e) {
      errors.push({ name: t.name, error: e.message });
    }
    await sleep(300);
  }

  const result = allTeachers.map((t) => ({
    ...t,
    schedule: Object.fromEntries(
      Object.entries(t.schedule).map(([day, slots]) => [
        day,
        slots.map((s) => ({
          ...s,
          time: s.time.replace(/^0/, "").replace("-0", "-"),
        })),
      ])
    ),
  }));

  const blob = new Blob([JSON.stringify(result, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "horaris_professors.json";
  a.click();
  URL.revokeObjectURL(url);

  setStatus(`✅ ${result.length} professors exportats`);
  setTimeout(removeUI, 2500);
  console.log("RESULTAT:", result);
  console.log("ERRORS:", errors);
}

const bookmarklet = `javascript:(${iEducaScraperFn.toString()})();`;
export default function SortidesApp() {
  const [step, setStep] = useState("upload");
  const [teachers, setTeachers] = useState(() => {
    try {
      const saved = localStorage.getItem("sortides_teachers");
      if (saved) return JSON.parse(saved);
    } catch {}
    return [];
  });
  const [savedAt, setSavedAt] = useState(() => {
    return localStorage.getItem("sortides_teachers_date") || null;
  });
  const [processing, setProcessing] = useState(false);
  const [processingName, setProcessingName] = useState("");
  const [errors, setErrors] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef();

  const newFranja = () => ({
    id: Date.now() + Math.random(),
    startSlot: "8:00-8:55",
    endSlot: "13:35-14:30",
    selectedGroups: [],
    excludedSubs: [],
    halfGroups: [],
    neededCount: 2,
  });
  const [trip, setTrip] = useState({
    day: "DILLUNS",
    titol: "",
    data: "",
    lloc: "",
    subject: "",
    franges: [newFranja()],
  });
  const [ranking, setRanking] = useState([]);
  const [confirmed, setConfirmed] = useState(new Set());
  const [activeFilters, setActiveFilters] = useState(new Set());
  // assignments: { "time|gapIdx": teacherName | null }
  const [assignments, setAssignments] = useState({});
  const [teacherFranjaMap, setTeacherFranjaMap] = useState({});

  const FILTERS = [
    {
      id: "fullRange",
      label: "Cobreix tot el rang",
      desc: "Entra i surt com la sortida",
    },
    {
      id: "hasGroup",
      label: "Té classe amb el grup",
      desc: "≥1 classe amb els grups seleccionats",
    },
    {
      id: "noLoss",
      label: "No deixa forats",
      desc: "No té classes amb grups que es queden",
    },
    {
      id: "subject",
      label: "Imparteix la matèria",
      desc: "Coincideix amb la matèria triada",
    },
  ];

  const subjects = extractSubjects(teachers);

  const sortedRanking = useMemo(() => {
    let r = [...ranking];
    if (activeFilters.has("fullRange")) r = r.filter((x) => x.coversFullRange);
    if (activeFilters.has("hasGroup"))
      r = r.filter((x) => x.tripClasses.length > 0);
    if (activeFilters.has("noLoss"))
      r = r.filter(
        (x) => x.stayClasses.length === 0 && x.halfClasses.length === 0
      );
    if (activeFilters.has("subject")) r = r.filter((x) => x.subjectMatch);
    r.sort((a, b) =>
      a.hoursShort !== b.hoursShort
        ? a.hoursShort - b.hoursShort
        : b.score - a.score
    );
    return r;
  }, [ranking, activeFilters]);

  const toggleFilter = (id) =>
    setActiveFilters((prev) => {
      const n = new Set(prev);
      n.has(id) ? n.delete(id) : n.add(id);
      return n;
    });
  const toggleConfirm = (name) =>
    setConfirmed((prev) => {
      const n = new Set(prev);
      n.has(name) ? n.delete(name) : n.add(name);
      return n;
    });

  const mergeTeachers = useCallback((incoming) => {
    setTeachers((prev) => {
      let u = [...prev];
      incoming.forEach((t) => {
        u = [...u.filter((x) => x.name !== t.name), t];
      });
      try {
        localStorage.setItem("sortides_teachers", JSON.stringify(u));
        const now = new Date().toLocaleString("ca-ES");
        localStorage.setItem("sortides_teachers_date", now);
        setSavedAt(now);
      } catch {}
      return u;
    });
  }, []);

  const processFiles = useCallback(
    async (files) => {
      const all = Array.from(files);
      const pdfs = all.filter(
        (f) => f.type === "application/pdf" || f.name.endsWith(".pdf")
      );
      const imgs = all.filter(
        (f) =>
          f.type.startsWith("image/") || /\.(jpe?g|png|webp)$/i.test(f.name)
      );
      const jsons = all.filter(
        (f) => f.name.endsWith(".json") || f.type === "application/json"
      );
      if (!pdfs.length && !imgs.length && !jsons.length) {
        setErrors(["Puja un PDF, imatges o un JSON del scraper."]);
        return;
      }
      setProcessing(true);
      setErrors([]);
      for (const file of jsons) {
        setProcessingName(file.name);
        try {
          const text = await file.text();
          const data = JSON.parse(text);
          const arr = Array.isArray(data) ? data : [data];
          if (!arr[0]?.name || !arr[0]?.schedule)
            throw new Error("Format no reconegut");
          mergeTeachers(arr.map(normalizeTeacher));
        } catch (e) {
          setErrors((p) => [...p, `Error JSON "${file.name}": ${e.message}`]);
        }
      }
      for (const file of pdfs) {
        setProcessingName(`${file.name}…`);
        try {
          const b64 = await fileToBase64(file);
          const arr = await parseAllFromPDF(b64);
          mergeTeachers(Array.isArray(arr) ? arr : [arr]);
        } catch (e) {
          setErrors((p) => [...p, `Error PDF "${file.name}": ${e.message}`]);
        }
      }
      for (const file of imgs) {
        setProcessingName(file.name);
        try {
          const b64 = await fileToBase64(file);
          const data = await parseSingleImage(b64, getMediaType(file));
          mergeTeachers([data]);
        } catch (e) {
          setErrors((p) => [...p, `Error imatge "${file.name}": ${e.message}`]);
        }
      }
      setProcessing(false);
      setProcessingName("");
    },
    [mergeTeachers]
  );

  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      setDragOver(false);
      processFiles(e.dataTransfer.files);
    },
    [processFiles]
  );

  const handleCompute = () => {
    // Per al ranking, unim tots els grups de totes les franges
    // i agafem el rang horari global (startSlot més aviat, endSlot més tard)
    const allGroups = [...new Set(trip.franges.flatMap((f) => f.selectedGroups))];
    const allExcludedSubs = [...new Set(trip.franges.flatMap((f) => f.excludedSubs || []))];
    const allHalfGroups = [...new Set(trip.franges.flatMap((f) => f.halfGroups || []))];
    const allSlots = trip.franges.flatMap((f) => slotsInRange(f.startSlot, f.endSlot));
    const sortedSlots = [...new Set(allSlots)].sort((a, b) => MORNING_SLOTS.indexOf(a) - MORNING_SLOTS.indexOf(b));
    const globalStart = sortedSlots[0] || "8:00-8:55";
    const globalEnd = sortedSlots[sortedSlots.length - 1] || "13:35-14:30";
    const tripForRanking = {
      day: trip.day,
      startSlot: globalStart,
      endSlot: globalEnd,
      selectedGroups: allGroups,
      excludedSubs: allExcludedSubs,
      halfGroups: allHalfGroups,
      subject: trip.subject,
    };
    const r = teachers.map((t) => analyzeTeacher(t, tripForRanking));
    r.sort((a, b) =>
      a.hoursShort !== b.hoursShort
        ? a.hoursShort - b.hoursShort
        : b.score - a.score
    );
    setRanking(r);
    setConfirmed(new Set());
    setAssignments({});
    setTeacherFranjaMap({});
    setStep("ranking");
  };

  // ── Exportació de dades ──────────────────────────────────────────────────
  const buildExportData = () => {
    // Unim tots els grups per computeCoverage global
    const allGroups = [...new Set(trip.franges.flatMap((f) => f.selectedGroups))];
    const allExcludedSubs = [...new Set(trip.franges.flatMap((f) => f.excludedSubs || []))];
    const allHalfGroups = [...new Set(trip.franges.flatMap((f) => f.halfGroups || []))];
    const allSlots = trip.franges.flatMap((f) => slotsInRange(f.startSlot, f.endSlot));
    const sortedSlots = [...new Set(allSlots)].sort((a, b) => MORNING_SLOTS.indexOf(a) - MORNING_SLOTS.indexOf(b));
    const globalStart = sortedSlots[0] || "8:00-8:55";
    const globalEnd = sortedSlots[sortedSlots.length - 1] || "13:35-14:30";
    const tripForCoverage = {
      day: trip.day,
      startSlot: globalStart,
      endSlot: globalEnd,
      selectedGroups: allGroups,
      excludedSubs: allExcludedSubs,
      halfGroups: allHalfGroups,
      subject: trip.subject,
    };
    const coverage = computeCoverage(teachers, [...confirmed], tripForCoverage, trip.franges, teacherFranjaMap);
    const { gapsBySlot } = coverage;

    const frangesExport = Object.entries(gapsBySlot).map(([time, gaps]) => {
      const cobertures = gaps.map((gap, gi) => {
        const assignat = assignments[`${time}|${gi}`] || null;
        const esNoCal = assignat === "NO_CAL_COBRIR";
        const esGuardia = assignat === "PROF. DE GUÀRDIA";
        // Evitar que el titular aparegui com a substitut
        const substitut = esNoCal ? null : assignat;
        return {
          assignatura: gap.subject || "",
          grup: gap.group || "",
          aula: gap.room || "",
          professorOriginal: gap.teacherName || "",
          substitut: substitut,
          nota: esNoCal ? "no_cal" : "",
        };
      });

      // Afegir professors alliberats sense cobrir (a disposició del centre)
      // Són professors confirmats que en aquesta franja no fan res
      const cobertsNoms = new Set(cobertures.map(c => c.substitut).filter(Boolean));
      const assignatsAquestSlot = [...confirmed].filter(name => {
        // Professor confirmat que cobria alguna classe en aquest slot
        return cobertures.some(c => c.substitut === name);
      });
      const assignatsEnAlgunaCobertura = new Set(
        Object.values(assignments).filter(Boolean)
      );
      // Professors confirmats que estan lliures en aquest slot (alliberats)
      const alliberats = [...confirmed].filter(name => {
        if (cobertsNoms.has(name)) return false; // ja cobreix aquí
        // Comprova si té alguna classe en aquest slot que hauria de deixar lliure
        const teacher = teachers.find(t => t.name === name);
        if (!teacher) return false;
        const daySlots = teacher.schedule[trip.day] || [];
        const slot = daySlots.find(s => s.time === time);
        if (!slot || slot.type === "free" || slot.type === "meeting") return false;
        return true; // té classe però no la cobreix ningú → a disposició
      });

      if (alliberats.length > 0) {
        const nomAlliberats = alliberats.map(n => {
          // Usar nomAbreujat inline
          const nc = (nm) => { const p = nm.split(",").map(s=>s.trim()); return p.length>=2?`${p[1]} ${p[0]}`:nm; };
          const c = nc(n).trim().split(/\s+/);
          return c.length < 2 ? nc(n) : `${c[0]} ${c[1][0]}.`;
        });
        cobertures.push({
          assignatura: "",
          grup: "",
          aula: "",
          professorOriginal: "",
          substitut: `${nomAlliberats.join(", ")} a disposició del centre`,
          nota: "alliberat",
        });
      }

      return { nom: SLOT_LABEL[time] || time, hora: time, cobertures };
    });

    // Acompanyants per franja
    const acompanyants = trip.franges.map((f) => {
      const assignedProfs = [...confirmed].filter(
        (name) => (teacherFranjaMap[name] || []).includes(f.id)
      );
      const profsToShow = assignedProfs.length > 0 ? assignedProfs : [...confirmed];
      return {
        hora: `${f.startSlot.split("-")[0]} – ${f.endSlot.split("-")[1]}`,
        grups: compactaGrups(f.selectedGroups),
        professors: profsToShow,
        responsables: "",
      };
    });

    return {
      event: {
        title: trip.titol || allGroups.join(", "),
        subtitle: trip.lloc || "",
        date: trip.data || DAY_LABELS[trip.day] || trip.day,
      },
      acompanyants,
      franges: frangesExport,
    };
  };

  const generateDocHTML = (data) => {
    const ev = data.event || {};
    const acomp = data.acompanyants || [];
    const franges = data.franges || [];
    const esc = (s) => String(s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    const nc = (n) => { // nomComplet inline
      if (!n) return "";
      const p = n.split(",").map(s => s.trim());
      return p.length >= 2 ? `${p[1]} ${p[0]}` : n;
    };
    const na = (n) => { // nomAbreujat inline
      if (!n) return "";
      const c = nc(n).trim().split(/\s+/);
      return c.length < 2 ? nc(n) : `${c[0]} ${c[1][0]}.`;
    };
    let acompRows = "";
    acomp.forEach((a) => {
      const profs = (a.professors || []).map((p) => `<li>${esc(nc(p))}</li>`).join("");
      acompRows += `<tr>
        <td class="cell-hora">${esc(a.hora)}</td>
        <td class="cell-grups">${esc(a.grups)}</td>
        <td class="cell-profs"><ul class="prof-list">${profs}</ul></td>
        <td class="cell-resp">${esc(nc(a.responsables))}</td>
      </tr>`;
    });
    let subsRows = "";
    franges.forEach((f) => {
      subsRows += `<tr class="franja-row"><td colspan="5">${esc(f.hora)}</td></tr>`;
      if (f.cobertures.length === 0) {
        subsRows += `<tr class="cob-row"><td colspan="5" style="color:#888;font-style:italic;text-align:center">— Sense incidències</td></tr>`;
      }
      f.cobertures.forEach((c) => {
        const subNom = c.substitut === "PROF. DE GUÀRDIA" ? "Prof. de guàrdia" : na(c.substitut);
        const sub = c.substitut
          ? `<span class="substitut">${esc(subNom)}</span>`
          : `<span class="uncovered">⚠ Sense cobrir</span>`;
        const titularNom = na(c.professorOriginal) || esc(c.professorOriginal);
        subsRows += `<tr class="cob-row">
          <td class="col-hora">${esc(f.hora)}</td>
          <td class="col-ass">${esc(c.assignatura)}</td>
          <td class="col-grp">${esc(c.grup)}</td>
          <td class="col-aul">${esc(c.aula)}</td>
          <td class="col-ori">${esc(titularNom)}</td>
          <td>${sub}</td>
        </tr>`;
      });
    });
    const now = new Date().toLocaleString("ca-ES");
    return `<!DOCTYPE html><html lang="ca"><head><meta charset="UTF-8"><title>${esc(ev.title)} — ${esc(ev.date)}</title><style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:Arial,Helvetica,sans-serif;font-size:10pt;color:#111;background:#e8e8e8;}
.page-wrap{padding:20px;display:flex;flex-direction:column;align-items:center;}
.document{width:210mm;min-height:297mm;background:#fff;box-shadow:0 4px 24px rgba(0,0,0,.18);padding:14mm;}
.doc-header{text-align:center;border-bottom:3px solid #1a1a2e;padding-bottom:8px;margin-bottom:14px;}
.doc-date{font-size:11pt;font-weight:bold;color:#1a1a2e;}
.doc-title{font-size:15pt;font-weight:bold;text-transform:uppercase;letter-spacing:.04em;}
.doc-subtitle{font-size:10pt;color:#444;margin-top:2px;}
.section-title{font-size:9pt;font-weight:bold;text-transform:uppercase;letter-spacing:.08em;background:#1a1a2e;color:#fff;padding:4px 8px;margin-bottom:0;text-align:center;}
table.acomp{width:100%;border-collapse:collapse;margin-bottom:14px;font-size:9.5pt;}
table.acomp th{background:#dde4ef;padding:4px 8px;text-align:center;font-weight:bold;border:1px solid #aab;font-size:8.5pt;}
table.acomp td{padding:5px 8px;border:1px solid #ccc;vertical-align:middle;text-align:center;}
.cell-hora{font-weight:bold;white-space:nowrap;width:90px;background:#f4f6fb;text-align:center!important;vertical-align:middle!important;}
.cell-grups{width:55px;background:#f4f6fb;text-align:center!important;}
.cell-profs{text-align:justify!important;min-width:200px;}
.cell-resp{font-size:8pt;color:#444;width:100px;text-align:center!important;}
table.subs{width:100%;border-collapse:collapse;font-size:9pt;}
table.subs th{background:#1a1a2e;color:#fff;padding:5px 8px;text-align:center;font-size:8.5pt;letter-spacing:.03em;}
tr.franja-row td{background:#dde4ef;font-weight:bold;padding:4px 8px;border:1px solid #aab;font-size:9pt;text-align:center;}
tr.cob-row td{padding:4px 8px;border:1px solid #ccc;vertical-align:middle;text-align:center;}
tr.cob-row:nth-child(even) td{background:#f9f9fb;}
.col-hora{width:90px;font-weight:bold;background:#f4f6fb;text-align:center;vertical-align:middle;}
.col-ass{width:110px;font-weight:bold;}
.col-grp{width:70px;}
.col-aul{width:50px;}
.col-ori{width:80px;color:#555;font-style:italic;font-size:8.5pt;}
.substitut{font-weight:bold;color:#c0392b;}
.uncovered{color:#c0392b;font-weight:bold;}
.doc-footer{margin-top:18px;font-size:7.5pt;color:#999;text-align:right;border-top:1px solid #ddd;padding-top:5px;}
.controls{position:sticky;top:0;z-index:100;background:#1a1a2e;color:#fff;padding:10px 24px;display:flex;align-items:center;gap:12px;}
.controls h1{font-size:14px;font-weight:600;flex:1;color:#a0c4ff;text-transform:uppercase;letter-spacing:.05em;}
.btn-print{padding:7px 16px;border:none;border-radius:5px;font-size:13px;font-weight:600;cursor:pointer;background:#27ae60;color:#fff;}
@media print{body{background:#fff;}.controls{display:none!important;}.page-wrap{padding:0;}.document{box-shadow:none;width:100%;min-height:unset;}@page{margin:10mm;size:A4;}}
</style></head><body>
<div class="controls"><h1>${esc(ev.title)} — ${esc(ev.date)}</h1><button class="btn-print" onclick="window.print()">🖨 Imprimir / Desar PDF</button></div>
<div class="page-wrap"><div class="document">
<div class="doc-header"><div class="doc-date">${esc(ev.date)}</div><div class="doc-title">${esc(ev.title)}</div>${ev.subtitle ? `<div class="doc-subtitle">${esc(ev.subtitle)}</div>` : ""}</div>
${acomp.length > 0 ? `
<div class="section-title">Professors/es acompanyants</div>
<table class="acomp"><thead><tr><th>Hora</th><th>Alumnes</th><th>Professors/es acompanyants</th><th>Responsables</th></tr></thead>
<tbody>${acompRows}</tbody></table>` : ""}
<div class="section-title" style="margin-top:10px">Professorat que substituirà</div>
<table class="subs"><thead><tr><th class="col-hora">Hora</th><th class="col-ass">Assignatura</th><th class="col-grp">Grup</th><th class="col-aul">Aula</th><th class="col-ori">Titular</th><th>Substitut/a</th></tr></thead>
<tbody>${subsRows}</tbody></table>
<div class="doc-footer">Document generat el ${now}</div>
</div></div></body></html>`;
  };

  const handleAssign = (key, name) => {
    setAssignments((prev) => {
      const n = { ...prev };
      if (name === null) delete n[key];
      else n[key] = name;
      return n;
    });
  };

  const totalClasses = teachers.reduce(
    (acc, t) =>
      acc +
      Object.values(t.schedule)
        .flat()
        .filter((s) => s.type === "class").length,
    0
  );
  // Per cada torn, comptem quants professors únics té assignats
  // Un professor que va a dos torns no se solapa compta una sola vegada per torn
  const totalNeeded = trip.franges.reduce((acc, f) => acc + (f.neededCount || 2), 0);
  const totalCovered = trip.franges.reduce((acc, f) => {
    const assignedToFranja = [...confirmed].filter(n => (teacherFranjaMap[n] || []).includes(f.id));
    // Si no hi ha assignació per torn (un sol torn), tots els confirmats compten
    const count = trip.franges.length === 1 ? confirmed.size : assignedToFranja.length;
    return acc + Math.min(count, f.neededCount || 2);
  }, 0);
  const remaining = Math.max(0, totalNeeded - totalCovered);

  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans',system-ui,sans-serif",
        minHeight: "100vh",
        backgroundColor: "#f5f4f0",
        color: "#1a1a2e",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@700&display=swap');
        @keyframes spin{to{transform:rotate(360deg);}}
        @keyframes fadeUp{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:translateY(0);}}
        button{font-family:inherit;transition:opacity 0.12s,transform 0.08s;}
        input{font-family:inherit;}
        button:hover:not(:disabled){opacity:0.82;}
        button:active:not(:disabled){transform:scale(0.97);}
        input:focus{outline:none;border-color:#1a2744!important;box-shadow:0 0 0 3px rgba(26,39,68,0.1);}
        *{box-sizing:border-box;}
      `}</style>

      <header
        style={{
          backgroundColor: "#1a2744",
          borderBottom: "3px solid #e8451e",
        }}
      >
        <div
          style={{
            maxWidth: 940,
            margin: "0 auto",
            padding: "0 24px",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            height: 62,
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 9,
              color: "white",
            }}
          >
            <span style={{ fontSize: 20, color: "#e8451e" }}>⬡</span>
            <span
              style={{ fontSize: 15, fontWeight: 700, letterSpacing: "-0.3px" }}
            >
              Assignador de Sortides
            </span>
          </div>
          <nav style={{ display: "flex", gap: 2 }}>
            {[
              ["upload", "Professors"],
              ["trip", "Sortida"],
              ["ranking", "Ranking & Cobertura"],
            ].map(([s, label], idx) => {
              const locked = s === "ranking" && ranking.length === 0;
              return (
                <button
                  key={s}
                  disabled={locked}
                  onClick={() => !locked && setStep(s)}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 6,
                    background: "transparent",
                    border: "none",
                    color: step === s ? "white" : "rgba(255,255,255,0.5)",
                    padding: "7px 13px",
                    borderRadius: 6,
                    fontSize: 13,
                    fontWeight: 500,
                    cursor: locked ? "not-allowed" : "pointer",
                    backgroundColor:
                      step === s ? "rgba(255,255,255,0.12)" : "transparent",
                  }}
                >
                  <span
                    style={{
                      width: 18,
                      height: 18,
                      borderRadius: "50%",
                      border: `1.5px solid ${
                        step === s ? "white" : "rgba(255,255,255,0.4)"
                      }`,
                      display: "inline-flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 10,
                      fontWeight: 700,
                    }}
                  >
                    {idx + 1}
                  </span>
                  {label}
                </button>
              );
            })}
          </nav>
        </div>
      </header>

      <main
        style={{
          maxWidth: 940,
          margin: "0 auto",
          padding: "32px 24px",
          animation: "fadeUp 0.3s ease",
        }}
      >
        {/* ══ STEP 1 ══ */}
        {step === "upload" && (
          <div>
            <h2
              style={{
                fontSize: 26,
                fontWeight: 700,
                color: "#1a2744",
                margin: "0 0 6px",
              }}
            >
              Horaris dels professors
            </h2>
            <p
              style={{
                fontSize: 13,
                color: "#6b7280",
                margin: "0 0 10px",
                fontWeight: 500,
              }}
            >
              Arrossega el botó a favorits / adreces d'interès.
            </p>

            {/* Banner de dades guardades */}
            {savedAt && teachers.length > 0 && (
              <div style={{
                background: "#f0fdf4",
                border: "1.5px solid #86efac",
                borderRadius: 10,
                padding: "12px 16px",
                marginBottom: 14,
                display: "flex",
                alignItems: "center",
                gap: 12,
                flexWrap: "wrap",
              }}>
                <span style={{ fontSize: 20 }}>💾</span>
                <div style={{ flex: 1 }}>
                  <p style={{ fontSize: 13, fontWeight: 700, color: "#166534", margin: 0 }}>
                    Horaris carregats des de la memòria del navegador
                  </p>
                  <p style={{ fontSize: 12, color: "#4b7c59", margin: "2px 0 0" }}>
                    {teachers.length} professors guardats · Actualitzat el {savedAt}
                  </p>
                </div>
                <button
                  onClick={() => {
                    if (window.confirm("Vols esborrar els horaris guardats? Hauràs de tornar a pujar el JSON.")) {
                      localStorage.removeItem("sortides_teachers");
                      localStorage.removeItem("sortides_teachers_date");
                      setTeachers([]);
                      setSavedAt(null);
                    }
                  }}
                  style={{
                    padding: "6px 12px",
                    borderRadius: 7,
                    border: "1px solid #86efac",
                    background: "white",
                    color: "#c0392b",
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: "pointer",
                  }}
                >
                  🗑 Esborrar
                </button>
              </div>
            )}

            <div style={{ margin: "0 0 10px" }}>
              <a
                ref={(el) => el && el.setAttribute("href", bookmarklet)}
                draggable="true"
                style={{
                  display: "inline-block",
                  padding: "10px 18px",
                  backgroundColor: "#e8451e",
                  color: "white",
                  borderRadius: 10,
                  fontSize: 14,
                  fontWeight: 700,
                  textDecoration: "none",
                  cursor: "grab",
                }}
              >
                Extractor d'horaris a iEduca
              </a>
            </div>

            <p
              style={{
                fontSize: 12,
                color: "#6b7280",
                margin: "0 0 18px",
                lineHeight: 1.45,
              }}
            >
              Entra a iEduca a la pàgina d'horaris i clica al botó a favorits
              per a extreure els horaris dels professors.
            </p>
            <p style={{ fontSize: 14, color: "#6b7280", margin: "0 0 22px" }}>
              Puja el{" "}
              <strong style={{ color: "#374151" }}>
                JSON del scraper iEduca
              </strong>
              , un PDF o imatges individuals.
            </p>
            <div
              onClick={() => !processing && fileInputRef.current?.click()}
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={handleDrop}
              style={{
                border: `2px dashed ${dragOver ? "#e8451e" : "#d1d5db"}`,
                borderRadius: 14,
                padding: "40px 32px",
                textAlign: "center",
                cursor: processing ? "default" : "pointer",
                backgroundColor: dragOver ? "#fff7f5" : "white",
                transition: "all 0.2s",
              }}
            >
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept="image/*,.pdf,application/pdf,.json,application/json"
                style={{ display: "none" }}
                onChange={(e) => processFiles(e.target.files)}
              />
              {processing ? (
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    gap: 12,
                  }}
                >
                  <Spinner />
                  <p style={{ fontSize: 14, color: "#374151", margin: 0 }}>
                    Processant <strong>{processingName}</strong>
                  </p>
                </div>
              ) : (
                <>
                  <div
                    style={{ fontSize: 28, marginBottom: 8, color: "#d1d5db" }}
                  >
                    ↑
                  </div>
                  <p
                    style={{
                      fontSize: 15,
                      color: "#374151",
                      margin: "0 0 12px",
                    }}
                  >
                    Arrossega aquí o{" "}
                    <span
                      style={{
                        color: "#e8451e",
                        fontWeight: 600,
                        textDecoration: "underline",
                      }}
                    >
                      fes clic
                    </span>
                  </p>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      gap: 8,
                      flexWrap: "wrap",
                    }}
                  >
                    {[
                      { icon: "📋", label: "JSON scraper", sub: "Recomanat" },
                      {
                        icon: "📄",
                        label: "PDF tots els horaris",
                        sub: "Via IA",
                      },
                      {
                        icon: "🖼️",
                        label: "Imatges JPG/PNG",
                        sub: "Un per un",
                      },
                    ].map(({ icon, label, sub }) => (
                      <div
                        key={label}
                        style={{
                          padding: "8px 14px",
                          borderRadius: 8,
                          border: "1px solid #e5e7eb",
                          backgroundColor: "#fafafa",
                          textAlign: "center",
                          minWidth: 120,
                        }}
                      >
                        <div style={{ fontSize: 18, marginBottom: 2 }}>
                          {icon}
                        </div>
                        <div
                          style={{
                            fontSize: 12,
                            fontWeight: 600,
                            color: "#374151",
                          }}
                        >
                          {label}
                        </div>
                        <div style={{ fontSize: 10, color: "#9ca3af" }}>
                          {sub}
                        </div>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>
            {errors.length > 0 && (
              <div
                style={{
                  marginTop: 8,
                  padding: "10px 14px",
                  backgroundColor: "#fef2f2",
                  borderRadius: 8,
                  border: "1px solid #fecaca",
                }}
              >
                {errors.map((e, i) => (
                  <p
                    key={i}
                    style={{ fontSize: 12, color: "#991b1b", margin: "1px 0" }}
                  >
                    ⚠ {e}
                  </p>
                ))}
              </div>
            )}
            {teachers.length > 0 && (
              <div style={{ marginTop: 26 }}>
                <p
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#9ca3af",
                    textTransform: "uppercase",
                    letterSpacing: "0.08em",
                    margin: "0 0 10px",
                  }}
                >
                  {teachers.length} professors · {totalClasses} classes
                  setmanals
                </p>
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: 6,
                    marginBottom: 18,
                  }}
                >
                  {teachers.map((t) => {
                    const nC = Object.values(t.schedule)
                      .flat()
                      .filter((s) => s.type === "class").length;
                    const grps = [
                      ...new Set(
                        Object.values(t.schedule)
                          .flat()
                          .filter((s) => s.type === "class")
                          .map((s) => s.group)
                          .filter(Boolean)
                      ),
                    ];
                    return (
                      <div
                        key={t.name}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 12,
                          backgroundColor: "white",
                          borderRadius: 10,
                          padding: "10px 14px",
                          border: "1px solid #e5e7eb",
                        }}
                      >
                        <Avatar name={t.name} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <p
                            style={{
                              fontSize: 14,
                              fontWeight: 600,
                              color: "#1a2744",
                              margin: "0 0 1px",
                            }}
                          >
                            {t.name}
                          </p>
                          <p
                            style={{
                              fontSize: 11,
                              color: "#9ca3af",
                              margin: 0,
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {nC} classes · {grps.slice(0, 5).join(", ")}
                            {grps.length > 5 ? "…" : ""}
                          </p>
                        </div>
                        <button
                          onClick={() =>
                            setTeachers((p) =>
                              p.filter((x) => x.name !== t.name)
                            )
                          }
                          style={{
                            background: "none",
                            border: "none",
                            color: "#d1d5db",
                            cursor: "pointer",
                            fontSize: 16,
                            padding: "2px 5px",
                          }}
                        >
                          ✕
                        </button>
                      </div>
                    );
                  })}
                </div>
                <button
                  onClick={() => setStep("trip")}
                  style={{
                    backgroundColor: "#e8451e",
                    color: "white",
                    border: "none",
                    borderRadius: 9,
                    padding: "12px 26px",
                    fontSize: 14,
                    fontWeight: 600,
                    cursor: "pointer",
                  }}
                >
                  Continua → Definir sortida
                </button>
              </div>
            )}
          </div>
        )}

        {/* ══ STEP 2 ══ */}
        {step === "trip" && (
          <div>
            <h2
              style={{
                fontSize: 26,
                fontWeight: 700,
                color: "#1a2744",
                margin: "0 0 6px",
              }}
            >
              Definir la sortida
            </h2>
            <p style={{ fontSize: 14, color: "#6b7280", margin: "0 0 18px" }}>
              Configura dia, horari, professors necessaris, grups i matèria.
            </p>
            <Card title="Informació de la sortida" hint="Títol, data i lloc que apareixeran al document.">
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {[
                  { key: "titol", label: "Títol", placeholder: "p.ex. Turó de Can Mates · 2n ESO" },
                  { key: "data",  label: "Data",  placeholder: "p.ex. 31 d'octubre" },
                  { key: "lloc",  label: "Lloc",  placeholder: "p.ex. Turó de Can Mates" },
                ].map(({ key, label, placeholder }) => (
                  <div key={key} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontSize: 13, color: "#374151", fontWeight: 500, minWidth: 40 }}>{label}</span>
                    <input
                      type="text"
                      value={trip[key]}
                      onChange={(e) => setTrip((t) => ({ ...t, [key]: e.target.value }))}
                      placeholder={placeholder}
                      style={{
                        flex: 1, padding: "7px 11px", borderRadius: 7,
                        border: "1.5px solid #e5e7eb", fontSize: 13,
                        color: "#111827", backgroundColor: "#ffffff",
                      }}
                    />
                  </div>
                ))}
              </div>
            </Card>
            <Card title="Dia de la sortida">
              <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
                {DAYS.map((day) => (
                  <button
                    key={day}
                    onClick={() => setTrip((t) => ({ ...t, day }))}
                    style={{
                      padding: "9px 18px",
                      borderRadius: 8,
                      border: `1.5px solid ${
                        trip.day === day ? "#1a2744" : "#e5e7eb"
                      }`,
                      backgroundColor: trip.day === day ? "#1a2744" : "white",
                      color: trip.day === day ? "white" : "#374151",
                      fontSize: 13,
                      fontWeight: 500,
                      cursor: "pointer",
                    }}
                  >
                    {DAY_LABELS[day]}
                  </button>
                ))}
              </div>
            </Card>

            {/* ── Franges ── */}
            {trip.franges.map((franja, fi) => {
              const updateFranja = (patch) =>
                setTrip((t) => ({
                  ...t,
                  franges: t.franges.map((f) =>
                    f.id === franja.id ? { ...f, ...patch } : f
                  ),
                }));
              return (
                <div
                  key={franja.id}
                  style={{
                    border: "1.5px solid #e5e7eb",
                    borderRadius: 12,
                    padding: "16px",
                    marginBottom: 14,
                    background: "#f9fafb",
                    position: "relative",
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      marginBottom: 12,
                    }}
                  >
                    <span
                      style={{
                        fontSize: 13,
                        fontWeight: 700,
                        color: "#1a2744",
                        textTransform: "uppercase",
                        letterSpacing: "0.06em",
                      }}
                    >
                      Torn {fi + 1}
                      {trip.franges.length > 1 && (
                        <span style={{ fontWeight: 400, color: "#6b7280", marginLeft: 6 }}>
                          — {franja.startSlot.split("-")[0]}–{franja.endSlot.split("-")[1]}
                        </span>
                      )}
                    </span>
                    {trip.franges.length > 1 && (
                      <button
                        onClick={() =>
                          setTrip((t) => ({
                            ...t,
                            franges: t.franges.filter((f) => f.id !== franja.id),
                          }))
                        }
                        style={{
                          background: "none",
                          border: "none",
                          color: "#c0392b",
                          cursor: "pointer",
                          fontSize: 13,
                          fontWeight: 600,
                          padding: "2px 6px",
                        }}
                      >
                        ✕ Eliminar torn
                      </button>
                    )}
                  </div>

                  <Card title="Horari del torn" hint="Inici i final d'aquest torn.">
                    <TimeRangePicker
                      startSlot={franja.startSlot}
                      endSlot={franja.endSlot}
                      onChange={(s, e) => updateFranja({ startSlot: s, endSlot: e })}
                    />
                  </Card>

                  <Card
                    title="Professors acompanyants necessaris"
                    hint="Quants professors han d'anar en aquest torn?"
                  >
                    <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                      <button
                        onClick={() =>
                          updateFranja({ neededCount: Math.max(1, (franja.neededCount || 2) - 1) })
                        }
                        style={{
                          width: 36, height: 36, borderRadius: 8,
                          border: "1.5px solid #e5e7eb", backgroundColor: "white",
                          fontSize: 20, cursor: "pointer", display: "flex",
                          alignItems: "center", justifyContent: "center",
                          fontWeight: 700, color: "#374151",
                        }}
                      >
                        −
                      </button>
                      <span
                        style={{
                          fontSize: 28, fontWeight: 800, color: "#1a2744",
                          fontFamily: "monospace", minWidth: 36, textAlign: "center",
                        }}
                      >
                        {franja.neededCount || 2}
                      </span>
                      <button
                        onClick={() =>
                          updateFranja({ neededCount: (franja.neededCount || 2) + 1 })
                        }
                        style={{
                          width: 36, height: 36, borderRadius: 8,
                          border: "1.5px solid #e5e7eb", backgroundColor: "white",
                          fontSize: 20, cursor: "pointer", display: "flex",
                          alignItems: "center", justifyContent: "center",
                          fontWeight: 700, color: "#374151",
                        }}
                      >
                        +
                      </button>
                      <span style={{ fontSize: 13, color: "#6b7280" }}>
                        professor{(franja.neededCount || 2) > 1 ? "s" : ""} acompanyant
                        {(franja.neededCount || 2) > 1 ? "s" : ""}
                      </span>
                    </div>
                  </Card>

                  <Card
                    title="Grups d'alumnes"
                    hint="Selecciona els grups d'aquest torn."
                  >
                    <GroupSelector
                      selected={franja.selectedGroups}
                      onChange={(v) => updateFranja({ selectedGroups: v })}
                      teachers={teachers}
                      excludedSubs={franja.excludedSubs}
                      onExcludedSubs={(v) => updateFranja({ excludedSubs: v })}
                      halfGroups={franja.halfGroups}
                      onHalfGroups={(v) => updateFranja({ halfGroups: v })}
                    />
                  </Card>
                </div>
              );
            })}

            {/* Botó afegir torn */}
            {trip.franges.length < 6 && (
              <button
                onClick={() =>
                  setTrip((t) => ({
                    ...t,
                    franges: [...t.franges, newFranja()],
                  }))
                }
                style={{
                  width: "100%",
                  padding: "11px",
                  border: "1.5px dashed #1a2744",
                  borderRadius: 10,
                  background: "white",
                  color: "#1a2744",
                  fontSize: 13,
                  fontWeight: 600,
                  cursor: "pointer",
                  marginBottom: 14,
                }}
              >
                + Afegeix un nou torn
              </button>
            )}

            <Card
              title="Matèria relacionada"
              hint="Opcional. Puntua extra als professors d'aquesta matèria."
            >
              <SubjectSelector
                subjects={subjects}
                selected={trip.subject}
                onChange={(v) => setTrip((t) => ({ ...t, subject: v }))}
              />
            </Card>
            <div
              style={{
                display: "flex",
                gap: 10,
                justifyContent: "flex-end",
                marginTop: 4,
              }}
            >
              <button
                onClick={() => setStep("upload")}
                style={{
                  backgroundColor: "white",
                  color: "#374151",
                  border: "1px solid #e5e7eb",
                  borderRadius: 9,
                  padding: "11px 20px",
                  fontSize: 13,
                  fontWeight: 500,
                  cursor: "pointer",
                }}
              >
                ← Enrere
              </button>
              <button
                onClick={handleCompute}
                style={{
                  backgroundColor: "#e8451e",
                  color: "white",
                  border: "none",
                  borderRadius: 9,
                  padding: "11px 26px",
                  fontSize: 14,
                  fontWeight: 600,
                  cursor: "pointer",
                }}
              >
                Calcular ranking →
              </button>
            </div>
          </div>
        )}

        {/* ══ STEP 3 ══ */}
        {step === "ranking" && (
          <div>
            {/* Page header */}
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "flex-start",
                marginBottom: 18,
                gap: 12,
                flexWrap: "wrap",
              }}
            >
              <div>
                <h2
                  style={{
                    fontSize: 26,
                    fontWeight: 700,
                    color: "#1a2744",
                    margin: "0 0 4px",
                  }}
                >
                  Ranking & Cobertura
                </h2>
                <p style={{ fontSize: 13, color: "#6b7280", margin: 0 }}>
                  <strong style={{ color: "#374151" }}>
                    {DAY_LABELS[trip.day]}
                  </strong>{" "}
                  ·{" "}
                  {trip.franges.map((f, fi) => (
                    <span key={f.id}>
                      {fi > 0 && " · "}
                      <strong style={{ color: "#374151" }}>
                        {f.startSlot.split("-")[0]}–{f.endSlot.split("-")[1]}
                      </strong>
                      {f.selectedGroups.length > 0 && (
                        <span style={{ color: "#374151" }}> ({f.selectedGroups.join(", ")})</span>
                      )}
                    </span>
                  ))}
                  {trip.subject && (
                    <>
                      {" "}·{" "}
                      <strong style={{ color: "#374151" }}>{trip.subject}</strong>
                    </>
                  )}
                </p>
              </div>
              <button
                onClick={() => setStep("trip")}
                style={{
                  backgroundColor: "white",
                  color: "#374151",
                  border: "1px solid #e5e7eb",
                  borderRadius: 9,
                  padding: "8px 14px",
                  fontSize: 12,
                  fontWeight: 500,
                  cursor: "pointer",
                  flexShrink: 0,
                }}
              >
                ← Canviar
              </button>
            </div>

            {/* ── Confirmed status bar ── */}
            <div style={{ backgroundColor: "#1a2744", borderRadius: 12, padding: "14px 20px", marginBottom: 14 }}>
              <p style={{ fontSize: 10, fontWeight: 700, color: "rgba(255,255,255,0.45)", textTransform: "uppercase", letterSpacing: "0.08em", margin: "0 0 10px" }}>
                Professors confirmats per la sortida
              </p>
              {confirmed.size === 0 ? (
                <p style={{ fontSize: 13, color: "rgba(255,255,255,0.4)", margin: 0, fontStyle: "italic" }}>
                  {trip.franges.length > 1
                    ? "Fes clic a \"+ Torn 1\" o \"+ Torn 2\" per assignar professors a cada torn"
                    : "Fes clic a \"+ Confirmar\" per marcar els professors que van de sortida"}
                </p>
              ) : (
                <>
                  {trip.franges.length > 1 ? (
                    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                      {trip.franges.map((franja, fi) => {
                        const profsInFranja = [...confirmed].filter(n => (teacherFranjaMap[n] || []).includes(franja.id));
                        return (
                          <div key={franja.id} style={{ background: "rgba(255,255,255,0.07)", borderRadius: 8, padding: "8px 12px" }}>
                            <p style={{ fontSize: 11, fontWeight: 700, color: "rgba(255,255,255,0.55)", margin: "0 0 5px" }}>
                              Torn {fi + 1} — {franja.startSlot.split("-")[0]}–{franja.endSlot.split("-")[1]}
                              {franja.selectedGroups.length > 0 && ` · ${franja.selectedGroups.join(", ")}`}
                            </p>
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                              {profsInFranja.length === 0
                                ? <span style={{ fontSize: 12, color: "rgba(255,255,255,0.3)", fontStyle: "italic" }}>Cap professor assignat</span>
                                : profsInFranja.map(n => (
                                  <span key={n} style={{ fontSize: 12, padding: "3px 10px", borderRadius: 99, backgroundColor: "#166534", color: "white", fontWeight: 500 }}>
                                    ✓ {n}
                                  </span>
                                ))
                              }
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                      {[...confirmed].map(n => (
                        <span key={n} style={{ fontSize: 12, padding: "3px 10px", borderRadius: 99, backgroundColor: "#166534", color: "white", fontWeight: 500, display: "flex", alignItems: "center", gap: 5 }}>
                          ✓ {n}
                          <button onClick={() => toggleConfirm(n)} style={{ background: "none", border: "none", color: "rgba(255,255,255,0.55)", cursor: "pointer", fontSize: 13, padding: "0 0 0 2px", lineHeight: 1 }}>✕</button>
                        </span>
                      ))}
                    </div>
                  )}
                  <div style={{ textAlign: "right", marginTop: 8 }}>
                    <span style={{ fontSize: 13, color: remaining === 0 ? "#4ade80" : "#e8451e", fontWeight: 700 }}>
                      {confirmed.size} confirmat{confirmed.size !== 1 ? "s" : ""}
                      {remaining > 0 ? ` · Falten ${remaining}` : " · ✓ Cobert"}
                    </span>
                  </div>
                </>
              )}
            </div>

            {/* ── Coverage Panel ── */}
            {confirmed.size > 0 && (
              <div style={{ marginBottom: 16 }}>
                <p
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#6b7280",
                    textTransform: "uppercase",
                    letterSpacing: "0.08em",
                    margin: "0 0 8px",
                  }}
                >
                  Cobertura de classes — franges horàries de la sortida
                </p>
                <CoveragePanel
                  teachers={teachers}
                  confirmedNames={[...confirmed]}
                  trip={(() => {
                    const allGroups = [...new Set(trip.franges.flatMap((f) => f.selectedGroups))];
                    const allExcludedSubs = [...new Set(trip.franges.flatMap((f) => f.excludedSubs || []))];
                    const allHalfGroups = [...new Set(trip.franges.flatMap((f) => f.halfGroups || []))];
                    const allSlots = trip.franges.flatMap((f) => slotsInRange(f.startSlot, f.endSlot));
                    const sortedSlots = [...new Set(allSlots)].sort((a, b) => MORNING_SLOTS.indexOf(a) - MORNING_SLOTS.indexOf(b));
                    return {
                      day: trip.day,
                      startSlot: sortedSlots[0] || "8:00-8:55",
                      endSlot: sortedSlots[sortedSlots.length - 1] || "13:35-14:30",
                      selectedGroups: allGroups,
                      excludedSubs: allExcludedSubs,
                      halfGroups: allHalfGroups,
                      subject: trip.subject,
                    };
                  })()}
                  franges={trip.franges}
                  teacherFranjaMap={teacherFranjaMap}
                  assignments={assignments}
                  onAssign={handleAssign}
                />

                {/* ── Botons d'exportació ── */}
                <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
                  <button
                    onClick={() => {
                      const data = buildExportData();
                      const html = generateDocHTML(data);
                      const win = window.open("", "_blank");
                      win.document.write(html);
                      win.document.close();
                    }}
                    style={{ padding: "8px 16px", background: "#c0392b", color: "white", border: "none", borderRadius: 7, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
                  >
                    🖨 Generar document PDF
                  </button>
                  <button
                    onClick={() => {
                      if (typeof XLSX === "undefined") {
                        alert("La llibreria Excel no s'ha carregat. Comprova la connexió.");
                        return;
                      }
                      const data = buildExportData();
                      const ev = data.event || {};
                      const acomp = data.acompanyants || [];
                      const franges = data.franges || [];

                      // Colors exactes de l'Excel model
                      const LILA     = "FF674EA7";
                      const FOSC     = "FF20124D";
                      const BLAU_CL1 = "FFD9E1F2";
                      const BLAU_CL2 = "FFF0F3FA";
                      const BLAU_CL3 = "FFCFE2F3";
                      const BLANC    = "FFFFFFFF";
                      const BLAU_TEXT = "FF0000FF"; // color text professors alliberats i substituts

                      // Helpers de nom
                      const nc = (n) => { if (!n) return ""; const p = n.split(",").map(s=>s.trim()); return p.length>=2?`${p[1]} ${p[0]}`:n; };
                      const na = (n) => { if (!n) return ""; const c=nc(n).trim().split(/\s+/); return c.length<2?nc(n):`${c[0]} ${c[1][0]}.`; };

                      // Funció de cel·la — 4 columnes (A=0,B=1,C=2,D=3)
                      const s = (v, bold, sz, bgRGB, fontRGB, italic, ha) => ({
                        v: v||"", t: "s",
                        s: {
                          font: { bold: !!bold, sz: sz||12, color: { rgb: fontRGB||"FF000000" }, italic: !!italic },
                          fill: bgRGB && bgRGB !== "00000000" ? { patternType:"solid", fgColor:{ rgb: bgRGB } } : undefined,
                          alignment: { wrapText: true, vertical:"center", horizontal: ha||"center" },
                          border: {
                            top:{style:"thin",color:{rgb:"FFCCCCCC"}},
                            bottom:{style:"thin",color:{rgb:"FFCCCCCC"}},
                            left:{style:"thin",color:{rgb:"FFCCCCCC"}},
                            right:{style:"thin",color:{rgb:"FFCCCCCC"}},
                          }
                        }
                      });

                      const wb = XLSX.utils.book_new();
                      const ws = {};
                      const merges = [];
                      let r = 0;

                      // F1: Títol gran — A:D fusionat
                      ws["A1"] = s(ev.title||"Sortida", false, 27, null, "FF000000", false, "center");
                      ["B","C","D"].forEach(c => { ws[`${c}1`] = s("",false,10,null,null); });
                      merges.push({s:{r:0,c:0},e:{r:0,c:3}});
                      r++;

                      // F2: Data | PROFESSORS ACOMPANYANTS | Hora | Alumnes
                      // Estructura: A=data, B=professors(B:B), C=hora, D=alumnes
                      ws[XLSX.utils.encode_cell({r,c:0})] = s(ev.date||"", true, 11, BLAU_CL1, "FF000000");
                      ws[XLSX.utils.encode_cell({r,c:1})] = s("PROFESSORS/ES  ACOMPANYANTS", true, 14, LILA, BLANC);
                      ws[XLSX.utils.encode_cell({r,c:2})] = s("Hora", true, 14, LILA, BLANC);
                      ws[XLSX.utils.encode_cell({r,c:3})] = s("Alumnes", true, 14, LILA, BLANC);
                      r++;

                      // Acompanyants: A=grups, B=professors (B únicament, no fusionat), C=hora, D=alumnes
                      const acompStartR = r;
                      acomp.forEach((a) => {
                        const profsText = (a.professors||[]).map(nc).join(", ");
                        ws[XLSX.utils.encode_cell({r,c:0})] = s(a.grups||"", false, 11, null, "FF000000");
                        ws[XLSX.utils.encode_cell({r,c:1})] = s(profsText, false, 11, null, "FF000000", false, "left");
                        ws[XLSX.utils.encode_cell({r,c:2})] = s(a.hora||"", true, 11, null, "FF000000");
                        ws[XLSX.utils.encode_cell({r,c:3})] = s(a.responsables||"", false, 11, null, "FF000000");
                        r++;
                      });
                      if (acomp.length > 1) merges.push({s:{r:acompStartR,c:0},e:{r:r-1,c:0}});

                      // Capçalera "PROFESSORAT QUE SUBSTITUIRÀ..." — A:D fusionat, 2 files
                      ws[XLSX.utils.encode_cell({r,c:0})] = s("PROFESSORAT QUE SUBSTITUIRÀ ALS PROFESSORS/ES QUE MARXEN DE SORTIDA", true, 14, FOSC, BLANC, false, "center");
                      [1,2,3].forEach(ci => { ws[XLSX.utils.encode_cell({r,c:ci})] = s("",false,10,FOSC,BLANC); });
                      merges.push({s:{r,c:0},e:{r,c:3}});
                      r++;
                      // Fila buida sota
                      [0,1,2,3].forEach(ci => { ws[XLSX.utils.encode_cell({r,c:ci})] = s("",false,10,FOSC,BLANC); });
                      merges.push({s:{r,c:0},e:{r,c:3}});
                      r++;

                      // Franges de substitució — A=hora, B:D=contingut fusionat
                      const PATI_SLOTS = new Set(["10:45-11:15", "11:15-11:45"]);
                      // Funció per rich text (nom en vermell, resta en negre)
                      const richSub = (subNom, resta) => ({
                        v: `${subNom}${resta}`, t: "s",
                        s: {
                          font: { bold: true, sz: 12, color: { rgb: "FFCC0000" } },
                          fill: undefined,
                          alignment: { wrapText: true, vertical: "center", horizontal: "left" },
                          border: {
                            top:{style:"thin",color:{rgb:"FFCCCCCC"}},
                            bottom:{style:"thin",color:{rgb:"FFCCCCCC"}},
                            left:{style:"thin",color:{rgb:"FFCCCCCC"}},
                            right:{style:"thin",color:{rgb:"FFCCCCCC"}},
                          }
                        },
                        // Rich text: nom en vermell, resta en negre
                        r: [
                          { t: subNom, s: { font: { bold: true, sz: 12, color: { rgb: "FFCC0000" } } } },
                          { t: resta,  s: { font: { bold: false, sz: 12, color: { rgb: "FF000000" } } } },
                        ]
                      });

                      franges.forEach((f, fi) => {
                        const bgFranja = fi % 2 === 0 ? BLAU_CL2 : BLAU_CL1;
                        const cobertures = f.cobertures || [];
                        const franjaStart = r;
                        const esPati = PATI_SLOTS.has(f.hora);
                        const horaLabel = esPati ? `${f.hora}\n(Pati)` : f.hora;

                        if (cobertures.length === 0) {
                          ws[XLSX.utils.encode_cell({r,c:0})] = s(horaLabel, true, 12, bgFranja, "FF000000");
                          [1,2,3].forEach(ci => ws[XLSX.utils.encode_cell({r,c:ci})] = s("",false,12,null,"FF000000"));
                          merges.push({s:{r,c:1},e:{r,c:3}});
                          r++;
                        } else {
                          cobertures.forEach((c, ci) => {
                            const bg = ci % 2 === 0 ? null : BLAU_CL3;
                            const bgStyle = bg && bg !== "00000000" ? { patternType:"solid", fgColor:{ rgb: bg } } : undefined;
                            const borderStyle = {
                              top:{style:"thin",color:{rgb:"FFCCCCCC"}},
                              bottom:{style:"thin",color:{rgb:"FFCCCCCC"}},
                              left:{style:"thin",color:{rgb:"FFCCCCCC"}},
                              right:{style:"thin",color:{rgb:"FFCCCCCC"}},
                            };

                            ws[XLSX.utils.encode_cell({r,c:0})] = s(ci===0?horaLabel:"", true, 12, bgFranja, "FF000000");

                            let cel;
                            if (c.nota === "alliberat") {
                              cel = s(c.substitut, true, 12, bg, BLAU_TEXT, false, "left");
                            } else if (c.nota === "no_cal") {
                              const titNom = na(c.professorOriginal) || c.professorOriginal || "";
                              const txt = `✓ No cal cobrir — ${titNom}${c.assignatura ? ` (${c.assignatura})` : ""}`;
                              cel = s(txt, true, 12, bg, "FF2E7D32", false, "left");
                            } else if (!c.substitut) {
                              const titNom = na(c.professorOriginal) || c.professorOriginal || "";
                              const txt = `⚠ SENSE COBRIR — ${titNom}${c.assignatura?` · ${c.assignatura}`:""}${c.grup?` · ${c.grup}`:""}`;
                              cel = s(txt, true, 12, bg, "FFCC0000", false, "left");
                            } else {
                              // Rich text: substitut en vermell, resta en negre
                              const subNom = c.substitut === "PROF. DE GUÀRDIA"
                                ? "Prof. de guàrdia"
                                : (na(c.substitut) || c.substitut);
                              const titNom = na(c.professorOriginal) || c.professorOriginal || "";
                              const detall = [
                                titNom ? `substitueix ${titNom}` : "",
                                c.assignatura || "",
                                c.grup || "",
                                c.aula ? `aula ${c.aula}` : "",
                              ].filter(Boolean).join("_");
                              const resta = detall ? `_${detall}` : "";
                              cel = richSub(subNom, resta);
                              if (bgStyle) cel.s.fill = bgStyle;
                            }
                            ws[XLSX.utils.encode_cell({r,c:1})] = cel;
                            ws[XLSX.utils.encode_cell({r,c:2})] = s("", false, 12, bg, "FF000000");
                            ws[XLSX.utils.encode_cell({r,c:3})] = s("", false, 12, bg, "FF000000");
                            merges.push({s:{r,c:1},e:{r,c:3}});
                            r++;
                          });
                          if (cobertures.length > 1) merges.push({s:{r:franjaStart,c:0},e:{r:r-1,c:0}});
                        }
                      });

                      ws["!ref"] = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:r-1,c:3}});
                      ws["!merges"] = merges;
                      // Amplades de columna com l'original: A=~12, B=~98, C=~16, D=~32
                      ws["!cols"] = [{wch:12},{wch:98},{wch:16},{wch:32}];
                      ws["!rows"] = [{hpt:46},...Array(r-1).fill({hpt:16})];

                      const sheetName = (DAY_LABELS[trip.day] || trip.day || "Sortida").substring(0,31);
                      XLSX.utils.book_append_sheet(wb, ws, sheetName);
                      const filename = `sortida_${(ev.date||trip.day||"doc").replace(/[\s/]/g,"_")}.xlsx`;
                      XLSX.writeFile(wb, filename);
                    }}
                    style={{ padding: "8px 16px", background: "#27ae60", color: "white", border: "none", borderRadius: 7, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
                  >
                    📊 Exportar Excel (.xlsx)
                  </button>
                </div>
              </div>
            )}

            {/* ── Filters ── */}
            <div
              style={{
                backgroundColor: "white",
                borderRadius: 10,
                padding: "14px 18px",
                marginBottom: 12,
                border: "1px solid #e5e7eb",
              }}
            >
              <p
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: "#6b7280",
                  textTransform: "uppercase",
                  letterSpacing: "0.08em",
                  margin: "0 0 8px",
                }}
              >
                Filtres del ranking
              </p>
              <div
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  gap: 6,
                  marginBottom: 10,
                }}
              >
                {FILTERS.map((f) => {
                  const a = activeFilters.has(f.id);
                  return (
                    <button
                      key={f.id}
                      onClick={() => toggleFilter(f.id)}
                      title={f.desc}
                      style={{
                        padding: "6px 12px",
                        borderRadius: 99,
                        border: `1.5px solid ${a ? "#1a2744" : "#e5e7eb"}`,
                        backgroundColor: a ? "#1a2744" : "white",
                        color: a ? "white" : "#4b5563",
                        fontSize: 12,
                        fontWeight: 500,
                        cursor: "pointer",
                      }}
                    >
                      {a && "✓ "}
                      {f.label}
                    </button>
                  );
                })}
                {activeFilters.size > 0 && (
                  <button
                    onClick={() => setActiveFilters(new Set())}
                    style={{
                      padding: "6px 12px",
                      borderRadius: 99,
                      border: "1.5px solid #fecaca",
                      backgroundColor: "#fef2f2",
                      color: "#991b1b",
                      fontSize: 12,
                      fontWeight: 500,
                      cursor: "pointer",
                    }}
                  >
                    ✕ Netejar filtres
                  </button>
                )}
              </div>
              <p style={{ fontSize: 11, color: "#9ca3af", margin: 0 }}>
                Ordenat per:{" "}
                <strong style={{ color: "#4b5563" }}>
                  hores cobertes del grup ↓
                </strong>{" "}
                · després per puntuació. Un professor que cobreix totes les
                hores té prioritat sobre un que en té menys, fins i tot si té
                hores lliures entremig.
              </p>
            </div>

            {/* ── Legend ── */}
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: "3px 14px",
                fontSize: 11,
                color: "#6b7280",
                marginBottom: 10,
                padding: "10px 14px",
                backgroundColor: "white",
                borderRadius: 8,
                border: "1px solid #e5e7eb",
              }}
            >
              <span
                style={{
                  fontWeight: 700,
                  color: "#374151",
                  width: "100%",
                  marginBottom: 2,
                }}
              >
                Colors als horaris:
              </span>
              {[
                {
                  bg: "#dcfce7",
                  c: "#166534",
                  t: "Classe amb el grup de la sortida",
                },
                {
                  bg: "#fee2e2",
                  c: "#991b1b",
                  t: "Classe amb grup que es queda (genera forat)",
                },
                {
                  bg: "#fff7ed",
                  c: "#9a3412",
                  t: "Mig grup que es queda (½ forat)",
                },
                { bg: "#fef9c3", c: "#854d0e", t: "Guàrdia assignada" },
                { bg: "#f3e8ff", c: "#6b21a8", t: "Reunió / tutoria tècnica" },
                {
                  bg: "#f3f4f6",
                  c: "#9ca3af",
                  t: "Hora lliure (sense classe programada)",
                },
              ].map(({ bg, c, t }) => (
                <span key={t}>
                  <span
                    style={{
                      display: "inline-block",
                      width: 10,
                      height: 10,
                      borderRadius: 2,
                      backgroundColor: bg,
                      border: `1px solid ${c}22`,
                      marginRight: 4,
                      verticalAlign: "middle",
                    }}
                  />
                  {t}
                </span>
              ))}
            </div>

            {/* Results count */}
            <p style={{ fontSize: 12, color: "#9ca3af", margin: "0 0 8px" }}>
              {sortedRanking.length} de {ranking.length} professors
              {activeFilters.size > 0 && (
                <>
                  {" "}
                  ·{" "}
                  <span style={{ color: "#e8451e" }}>
                    {activeFilters.size} filtre
                    {activeFilters.size > 1 ? "s" : ""} actiu
                    {activeFilters.size > 1 ? "s" : ""}
                  </span>
                </>
              )}
            </p>

            {/* Ranking list */}
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {sortedRanking.length === 0 ? (
                <div
                  style={{
                    padding: "32px",
                    textAlign: "center",
                    backgroundColor: "white",
                    borderRadius: 12,
                    border: "1px solid #e5e7eb",
                  }}
                >
                  <p
                    style={{
                      fontSize: 14,
                      color: "#6b7280",
                      margin: "0 0 10px",
                    }}
                  >
                    Cap professor compleix els filtres actius.
                  </p>
                  <button
                    onClick={() => setActiveFilters(new Set())}
                    style={{
                      padding: "7px 14px",
                      borderRadius: 8,
                      border: "1px solid #e5e7eb",
                      backgroundColor: "white",
                      color: "#374151",
                      fontSize: 13,
                      cursor: "pointer",
                    }}
                  >
                    Treure filtres
                  </button>
                </div>
              ) : (
                sortedRanking.map((r, i) => (
                  <RankingCard
                    key={r.name}
                    r={r}
                    i={i}
                    trip={trip}
                    confirmed={confirmed}
                    onToggleConfirm={toggleConfirm}
                    teacherFranjaMap={teacherFranjaMap}
                    onAssignFranja={(name, franjaIds) =>
                      setTeacherFranjaMap(prev =>
                        franjaIds === null || franjaIds.length === 0
                          ? Object.fromEntries(Object.entries(prev).filter(([k]) => k !== name))
                          : { ...prev, [name]: franjaIds }
                      )
                    }
                  />
                ))
              )}
            </div>

            {/* Score legend */}
            <div
              style={{
                marginTop: 18,
                padding: "12px 16px",
                backgroundColor: "#f9fafb",
                borderRadius: 10,
                border: "1px solid #e5e7eb",
              }}
            >
              <p
                style={{
                  fontSize: 10,
                  fontWeight: 700,
                  color: "#9ca3af",
                  textTransform: "uppercase",
                  letterSpacing: "0.08em",
                  margin: "0 0 6px",
                }}
              >
                Com es calcula la puntuació
              </p>
              <div
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  gap: "3px 18px",
                  fontSize: 11,
                  color: "#4b5563",
                }}
              >
                <span>👥 +50 coneix el grup (li fa classe algun dia)</span>
                <span>
                  ✓ +60 cobreix tot el rang horari · −15 per cada hora que falta
                </span>
                <span>
                  📚 +20 per hora de classe amb el grup aquell dia (lectura =
                  +10)
                </span>
                <span>★ +15 imparteix la matèria seleccionada</span>
                <span>🔴 −12 per classe amb grup que es queda</span>
                <span>🟠 −6 per classe de mig grup que es queda</span>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
