import React, { useState, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, ZAxis, CartesianGrid, Tooltip,
  Cell, ScatterChart, Scatter, ReferenceLine, Legend,
} from "recharts";
import {
  Upload, AlertTriangle, Download, RotateCcw,
  TrendingDown, CheckCircle2, ChevronRight, Info,
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/*  BISR logo (inline, so the file stays fully self-contained)            */
/* ---------------------------------------------------------------------- */

const BISR_LOGO =
  "data:image/png;base64,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";

/* ---------------------------------------------------------------------- */
/*  Grade scale definitions                                               */
/* ---------------------------------------------------------------------- */

// Every grade scale gets converted to a plain number so VA (actual - predicted)
// can be calculated the same way regardless of qualification type.
// GCSE grades are already numeric (9 down to 1, plus U = unclassified = 0),
// so they don't need a lookup map - see gradeToNumeric() below.
const ALEVEL_MAP = { "A*": 6, A: 5, B: 4, C: 3, D: 2, E: 1, U: 0 };
const AS_MAP = { A: 5, B: 4, C: 3, D: 2, E: 1, U: 0 }; // AS has no A* grade

// Threshold bands used for the "% at grade X or above" summary stats.
// `min` is the numeric cutoff (from the maps above) a student's grade must
// reach or beat to count in that band.
const GCSE_BANDS = [
  { label: "% 9", min: 9 },
  { label: "% 9-8", min: 8 },
  { label: "% 9-7", min: 7 },
  { label: "% 9-6", min: 6 },
  { label: "% 9-4", min: 4 },
];
const AS_BANDS = [
  { label: "% A", min: 5 },
  { label: "% A-B", min: 4 },
  { label: "% A-C", min: 3 },
  { label: "% A-E", min: 1 },
];
const ALEVEL_BANDS = [
  { label: "% A*-A", min: 5 },
  { label: "% A*-B", min: 4 },
  { label: "% A*-C", min: 3 },
  { label: "% A*-E", min: 1 },
];

// Highest-to-lowest grade order, used to lay out the grade distribution chart
// left-to-right in the order teachers expect (best grade first).
const GCSE_ORDER = ["9", "8", "7", "6", "5", "4", "3", "2", "1", "U"];
const ALEVEL_ORDER = ["A*", "A", "B", "C", "D", "E", "U"];
const AS_ORDER = ["A", "B", "C", "D", "E", "U"];

const QUAL_LABELS = { GCSE: "GCSE (9–1)", AS: "AS Level (A–E)", ALEVEL: "A Level (A*–E)" };

// Fixed axis range for the predicted-vs-actual scatter chart. Both axes use
// the same 0-9 scale (rather than auto-fitting to the data) so every grade
// number is labelled and the y = x "on target" reference line is always a
// true diagonal, regardless of qualification type.
const SCATTER_DOMAIN = [0, 9];
const SCATTER_TICKS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];

// Small lookup helpers so the rest of the code can just call bandsFor(qualType)
// etc. instead of repeating this three-way branch everywhere.
function bandsFor(qualType) {
  return qualType === "GCSE" ? GCSE_BANDS : qualType === "AS" ? AS_BANDS : ALEVEL_BANDS;
}
function orderFor(qualType) {
  return qualType === "GCSE" ? GCSE_ORDER : qualType === "AS" ? AS_ORDER : ALEVEL_ORDER;
}
function passMinFor(qualType) {
  return qualType === "GCSE" ? 4 : 3; // grade 4 / grade C
}

// Converts a raw grade value from the spreadsheet (e.g. "7", "A*", "u", " C ")
// into a plain number Claude/the app can do maths with. Returns null if the
// value can't be parsed, so callers can tell "no grade" apart from "grade 0".
function gradeToNumeric(raw, qualType) {
  if (raw === null || raw === undefined) return null;
  const v = String(raw).trim().toUpperCase();
  if (v === "") return null;
  if (qualType === "GCSE") {
    if (v === "U") return 0;
    const n = parseFloat(v);
    return isNaN(n) ? null : n;
  }
  // AS / A Level: try the letter-grade lookup first (A*, A, B...). If the
  // cell instead contains a number (some schools store CAT4 predictions as
  // decimals on the same 0-6 scale), fall back to parsing it directly.
  const map = qualType === "ALEVEL" ? ALEVEL_MAP : AS_MAP;
  if (map[v] !== undefined) return map[v];
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

// Used purely for display - shows the student's original grade text (e.g. "A*"),
// not the converted number, so the tables read the way a teacher expects.
function rawGradeLabel(raw) {
  if (raw === null || raw === undefined) return "U";
  const v = String(raw).trim().toUpperCase();
  return v === "" ? "U" : v;
}

// Looks at a sample of values from the "actual grade" column and guesses
// whether this is a GCSE (numeric), AS, or A Level (letter) file. AS and
// A Level share the same letters (A-E), so the only signal that tells them
// apart is whether an A* ever appears - if it does, it must be A Level. The
// user can still override this guess on the mapping screen either way.
function detectQualType(sampleValues) {
  const letterRe = /^(A\*|[A-E])$/;
  const numRe = /^[0-9](\.[0-9]+)?$/;
  let letters = 0, hasStar = false, numbers = 0, total = 0;
  sampleValues.forEach((raw) => {
    if (raw === null || raw === undefined) return;
    const v = String(raw).trim().toUpperCase();
    if (v === "" || v === "U") return;
    total++;
    if (letterRe.test(v)) { letters++; if (v === "A*") hasStar = true; }
    else if (numRe.test(v)) numbers++;
  });
  if (total === 0) return "GCSE";
  if (letters >= numbers) return hasStar ? "ALEVEL" : "AS";
  return "GCSE";
}

/* ---------------------------------------------------------------------- */
/*  Column auto-detection                                                 */
/* ---------------------------------------------------------------------- */

// One entry per field the app needs. `patterns` are the header phrases we'll
// look for (in priority order) when guessing which spreadsheet column maps
// to each field - see guessColumn() below. To recognise a new header
// variant (e.g. your MIS exports "Pupil Ref" instead of "Student ID"), just
// add another string to the relevant patterns array.
const FIELD_DEFS = [
  { key: "student", label: "Student name / ID", required: true,
    patterns: ["student name", "student id", "pupil name", "pupil", "full name", "name", "student"] },
  { key: "cls", label: "Class (optional)", required: false,
    patterns: ["class", "teaching group", "set", "group"] },
  { key: "staff", label: "Staff initials", required: false,
    patterns: ["staff initials", "initials", "teacher", "staff"] },
  { key: "cat4", label: "CAT4 target grade", required: true,
    patterns: ["cat4 predicted", "cat4", "predicted grade", "target grade", "predicted", "target"] },
  { key: "actual", label: "Actual exam grade", required: true,
    patterns: ["exam grade", "actual grade", "actual", "result", "final grade", "grade"] },
  { key: "gender", label: "Gender", required: false, patterns: ["gender", "sex"] },
  { key: "send", label: "SEND status", required: false, patterns: ["send status", "send", "sen"] },
  { key: "eal", label: "EAL status", required: false, patterns: ["eal status", "eal"] },
  { key: "moreAble", label: "More able", required: false,
    patterns: ["more able", "gifted and talented", "g&t", "ma"] },
];

// Lowercases, trims, and collapses spacing/underscores/hyphens so headers
// like "Student_Name", "student  name", and "Student Name" all compare equal.
function normalize(s) {
  return String(s || "").trim().toLowerCase().replace(/[_\-]+/g, " ").replace(/\s+/g, " ");
}

// Tries to find the best-matching column header for one field. Two passes:
// first look for an exact match against any pattern (most confident), then
// fall back to a "header contains this pattern" partial match. Patterns are
// checked in the order they're listed in FIELD_DEFS, so put more specific
// phrases first (e.g. "cat4 predicted" before the bare "cat4").
function guessColumn(headers, patterns) {
  const normHeaders = headers.map((h) => ({ h, n: normalize(h) }));
  for (const p of patterns) {
    const exact = normHeaders.find((x) => x.n === p);
    if (exact) return exact.h;
  }
  for (const p of patterns) {
    const partial = normHeaders.find((x) => x.n.includes(p));
    if (partial) return partial.h;
  }
  return ""; // no match - shows as "Not in this file" on the mapping screen
}

// Runs guessColumn() for every field in FIELD_DEFS, producing the initial
// { student: "...", cls: "...", ... } mapping shown on the confirm-columns
// screen. The user can still override any of these before continuing.
function autoMapColumns(headers) {
  const mapping = {};
  FIELD_DEFS.forEach((f) => { mapping[f.key] = guessColumn(headers, f.patterns); });
  return mapping;
}

/* ---------------------------------------------------------------------- */
/*  Small presentational helpers                                          */
/* ---------------------------------------------------------------------- */

// Formats a VA number for display: "+1.5", "-0.33", "0" - always signed
// (except zero), and trims trailing zeros so "2.00" shows as just "2".
function fmtVA(v) {
  if (v === null || v === undefined || isNaN(v)) return "—";
  const fixed = v.toFixed(2);
  const trimmed = fixed.replace(/\.?0+$/, "");
  return (v > 0 ? "+" : "") + (trimmed === "" || trimmed === "-" ? "0" : trimmed);
}

// Sage green for VA >= 0, clay red for VA < 0 - used to colour stat cards,
// table cells, and chart bars consistently throughout the dashboard.
function vaColor(v) {
  if (v === null || v === undefined || isNaN(v)) return "var(--slate)";
  return v < 0 ? "var(--clay)" : "var(--sage)";
}

// Used for SEND / EAL / More Able columns, which schools tend to fill in
// inconsistently (Y/N, Yes/No, TRUE/FALSE, 1/0, or the field's own name).
// Anything that looks like an affirmative value counts as "flagged".
function isPositiveFlag(raw) {
  if (raw === null || raw === undefined) return false;
  const v = String(raw).trim().toLowerCase();
  return ["y", "yes", "true", "1", "send", "eal"].includes(v);
}

// Custom tooltip for the predicted-vs-actual scatter chart. Each bubble can
// represent more than one student (see scatterData below), so this shows the
// count and always lists who they are - rather than the default Recharts
// tooltip, which would just print the raw x/y/z numbers. Very large groups
// are truncated so the tooltip itself doesn't take over the screen.
const SCATTER_TOOLTIP_MAX_NAMES = 12;
function ScatterTooltip({ active, payload }) {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  const shown = d.names.slice(0, SCATTER_TOOLTIP_MAX_NAMES);
  const extra = d.names.length - shown.length;
  return (
    <div style={{
      background: "#fffdf8", border: "1px solid var(--line)", borderRadius: 4,
      padding: "8px 11px", fontSize: 12.5, maxWidth: 260, boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
    }}>
      <div><b>Target {d.x}</b> → <b>Actual {d.y}</b></div>
      <div style={{ marginTop: 3, color: "var(--slate)" }}>
        {d.count} student{d.count === 1 ? "" : "s"}
      </div>
      <div style={{ marginTop: 3, fontSize: 11.5, color: "var(--slate)" }}>
        {shown.join(", ")}{extra > 0 ? `, +${extra} more` : ""}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Main component                                                        */
/* ---------------------------------------------------------------------- */

export default function ExamAnalysisDashboard() {
  // `stage` drives which of the three screens is shown - see the render
  // section near the bottom. Everything else here is raw input state; the
  // *derived* dashboard numbers (VA, bands, charts...) are computed further
  // down with useMemo rather than stored directly.
  const [stage, setStage] = useState("upload"); // upload | mapping | dashboard
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);          // column names from the uploaded file
  const [rows, setRows] = useState([]);                 // raw parsed rows, one object per student
  const [mapping, setMapping] = useState({});           // field key -> chosen column header, e.g. { student: "Name", cls: "Class" }
  const [qualType, setQualType] = useState("GCSE");     // GCSE | AS | ALEVEL
  const [subjectLabel, setSubjectLabel] = useState("");
  const [parseError, setParseError] = useState("");
  const [notes, setNotes] = useState({});               // intervention notes, keyed by student.idx
  const fileInputRef = useRef(null);

  /* ---------------- file handling ---------------- */

  // Parses whatever file the user picked/dropped. CSV goes through PapaParse;
  // Excel (.xlsx/.xls) is read as an ArrayBuffer and converted with SheetJS.
  // Either path ends by calling finishLoad() with the same shape of data
  // (column headers + row objects), so the rest of the app doesn't need to
  // care which file type was uploaded.
  const handleFile = useCallback((file) => {
    if (!file) return;
    setParseError("");
    setFileName(file.name);
    const ext = file.name.split(".").pop().toLowerCase();

    if (ext === "csv") {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const flds = results.meta.fields || [];
          if (!flds.length) { setParseError("Couldn't find any columns in that file."); return; }
          finishLoad(flds, results.data);
        },
        error: () => setParseError("Couldn't read that CSV file. Please check it and try again."),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          if (!data.length) { setParseError("That spreadsheet looks empty."); return; }
          finishLoad(Object.keys(data[0]), data);
        } catch (err) {
          setParseError("Couldn't read that Excel file. Please check it and try again.");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setParseError("Please upload a .csv, .xlsx, or .xls file.");
    }
  }, []);

  // Called once parsing succeeds (from either branch above). Stores the raw
  // data, runs auto-detection for both the column mapping and the
  // qualification type (sampling up to 50 rows of the actual-grade column
  // is enough to reliably tell GCSE/AS/A Level apart), then moves to the
  // mapping-confirmation screen.
  function finishLoad(flds, data) {
    setHeaders(flds);
    setRows(data);
    const guessed = autoMapColumns(flds);
    setMapping(guessed);
    if (guessed.actual) {
      const sample = data.slice(0, Math.min(50, data.length)).map((r) => r[guessed.actual]);
      setQualType(detectQualType(sample));
    }
    setStage("mapping");
  }

  function onDrop(e) {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    handleFile(file);
  }

  function resetAll() {
    setStage("upload"); setFileName(""); setHeaders([]); setRows([]);
    setMapping({}); setQualType("GCSE"); setSubjectLabel(""); setParseError(""); setNotes({});
  }

  /* ---------------- computed dashboard data ---------------- */
  // Everything below is derived (with useMemo) from `rows` + `mapping` +
  // `qualType`. Nothing here is stored in state directly - change a mapping
  // or the qualification type and all of these recalculate automatically.

  // The core per-student record. Every other stat on the dashboard is built
  // by filtering/grouping/averaging this array - this is the one place that
  // reads raw spreadsheet values and converts them into numbers.
  const students = useMemo(() => {
    if (stage !== "dashboard") return [];
    return rows.map((row, idx) => {
      const predictedRaw = mapping.cat4 ? row[mapping.cat4] : null;
      const actualRaw = mapping.actual ? row[mapping.actual] : null;
      const predicted = gradeToNumeric(predictedRaw, qualType);
      const actual = gradeToNumeric(actualRaw, qualType);
      // VA = actual minus predicted, on the numeric scale. null if either
      // grade couldn't be parsed (e.g. the cell was blank).
      const va = predicted !== null && actual !== null ? Math.round((actual - predicted) * 100) / 100 : null;
      return {
        idx, // stable row reference, used as the React key and as the notes-table lookup key
        name: (mapping.student ? row[mapping.student] : "") || `Student ${idx + 1}`,
        cls: mapping.cls ? ((row[mapping.cls] || "").toString().trim() || "Unassigned") : null,
        staff: mapping.staff ? row[mapping.staff] : "",
        predictedRaw, actualRaw, predicted, actual, va,
        gender: mapping.gender ? row[mapping.gender] : null,
        send: mapping.send ? row[mapping.send] : null,
        eal: mapping.eal ? row[mapping.eal] : null,
        moreAble: mapping.moreAble ? row[mapping.moreAble] : null,
      };
    });
  }, [rows, mapping, qualType, stage]);

  // Students whose actual grade parsed successfully - the denominator for
  // "Entries" and the grade-band percentages. Rows with an unreadable grade
  // (blank, typo, etc.) are excluded and counted in skippedCount instead.
  const validGrades = useMemo(() => students.filter((s) => s.actual !== null), [students]);
  const skippedCount = students.length - validGrades.length;
  // Students with both a predicted AND actual grade, i.e. a usable VA figure.
  const withVA = useMemo(() => students.filter((s) => s.va !== null), [students]);

  // "% of students at grade X or above", for each band defined in
  // GCSE_BANDS / AS_BANDS / ALEVEL_BANDS.
  const bandStats = useMemo(() => {
    const bands = bandsFor(qualType);
    const total = validGrades.length || 1; // avoid divide-by-zero
    return bands.map((b) => ({
      label: b.label,
      pct: Math.round((validGrades.filter((s) => s.actual >= b.min).length / total) * 1000) / 10,
    }));
  }, [validGrades, qualType]);

  // Count of students at each individual grade (9, 8, 7... or A*, A, B...),
  // for both actual and predicted grades - used for the grade distribution
  // bar chart. `pass` flags whether that grade clears the pass mark (based
  // on the actual scale), so the chart can colour the actual-grade bars.
  const gradeDistribution = useMemo(() => {
    const order = orderFor(qualType);
    const passMin = passMinFor(qualType);
    const actualCounts = {};
    const predictedCounts = {};
    order.forEach((g) => { actualCounts[g] = 0; predictedCounts[g] = 0; });
    validGrades.forEach((s) => {
      const label = rawGradeLabel(s.actualRaw);
      if (actualCounts[label] === undefined) actualCounts[label] = 0;
      actualCounts[label]++;
    });
    students.forEach((s) => {
      if (s.predicted === null) return;
      const label = rawGradeLabel(s.predictedRaw);
      if (predictedCounts[label] === undefined) predictedCounts[label] = 0;
      predictedCounts[label]++;
    });
    return order.map((g) => ({
      grade: g,
      actualCount: actualCounts[g] || 0,
      predictedCount: predictedCounts[g] || 0,
      pass: gradeToNumeric(g, qualType) >= passMin,
    }));
  }, [validGrades, students, qualType]);

  // Groups students by class and averages their VA. Returns an empty array
  // (and every class-related chart/table quietly hides itself) when no class
  // column was mapped - see the `mapping.cls &&` checks in the render below.
  const byClass = useMemo(() => {
    if (!mapping.cls) return [];
    const groups = {};
    students.forEach((s) => {
      if (!groups[s.cls]) groups[s.cls] = { cls: s.cls, staff: s.staff || "", n: 0, vaSum: 0, vaN: 0 };
      groups[s.cls].n++;
      if (s.staff && !groups[s.cls].staff) groups[s.cls].staff = s.staff;
      if (s.va !== null) { groups[s.cls].vaSum += s.va; groups[s.cls].vaN++; }
    });
    return Object.values(groups)
      .map((g) => ({ ...g, va: g.vaN ? Math.round((g.vaSum / g.vaN) * 100) / 100 : null }))
      .sort((a, b) => a.cls.localeCompare(b.cls));
  }, [students, mapping.cls]);

  // Cohort-wide VA plus a breakdown for each subgroup column that's actually
  // present in the file. Gender is treated as multi-category (whatever
  // values appear - M/F/Other, etc.), while SEND/EAL/More Able are treated
  // as yes/no flags (see isPositiveFlag). Any subgroup with no mapped column,
  // or with zero flagged students, is simply left out of the result.
  const subgroupStats = useMemo(() => {
    const out = [];
    const allVA = withVA.length ? withVA.reduce((a, s) => a + s.va, 0) / withVA.length : null;
    out.push({ name: "All", va: allVA !== null ? Math.round(allVA * 100) / 100 : null, n: withVA.length });

    if (mapping.gender) {
      const cats = {};
      withVA.forEach((s) => {
        const g = (s.gender || "").toString().trim();
        if (!g) return;
        if (!cats[g]) cats[g] = { sum: 0, n: 0 };
        cats[g].sum += s.va; cats[g].n++;
      });
      Object.entries(cats).forEach(([g, v]) => out.push({ name: g, va: Math.round((v.sum / v.n) * 100) / 100, n: v.n }));
    }
    [["send", "SEND"], ["eal", "EAL"], ["moreAble", "More able"]].forEach(([key, label]) => {
      if (!mapping[key]) return;
      const flagged = withVA.filter((s) => isPositiveFlag(s[key]));
      if (!flagged.length) return;
      const avg = flagged.reduce((a, s) => a + s.va, 0) / flagged.length;
      out.push({ name: label, va: Math.round(avg * 100) / 100, n: flagged.length });
    });
    return out;
  }, [withVA, mapping]);

  // Students who underperformed their prediction, worst-first - feeds the
  // "below CAT4" table where staff add intervention notes.
  const negativeStudents = useMemo(
    () => withVA.filter((s) => s.va < 0).sort((a, b) => a.va - b.va),
    [withVA]
  );

  // Predicted/actual pairs for the scatter chart. Since grades are whole
  // numbers, it's common for several students to land on the exact same
  // (predicted, actual) pair - if we plotted one dot per student those would
  // sit exactly on top of each other and effectively disappear. Instead we
  // group by the pair and size each dot by how many students it represents
  // (via the `count` field, used as the bubble size below), so nobody gets
  // silently hidden. Students with a missing predicted or actual grade are
  // excluded here (and are already reflected in "rows skipped" elsewhere).
  const scatterData = useMemo(() => {
    const groups = {};
    students.forEach((s) => {
      if (s.predicted === null || s.actual === null) return;
      const key = `${s.predicted}|${s.actual}`;
      if (!groups[key]) groups[key] = { x: s.predicted, y: s.actual, count: 0, names: [] };
      groups[key].count++;
      groups[key].names.push(s.name);
    });
    return Object.values(groups);
  }, [students]);

  const entriesCount = validGrades.length;
  const passRate = useMemo(() => {
    const passMin = passMinFor(qualType);
    if (!entriesCount) return null;
    return Math.round((validGrades.filter((s) => s.actual >= passMin).length / entriesCount) * 1000) / 10;
  }, [validGrades, entriesCount, qualType]);

  // Updates one field (intervention / reasons / steps) for one student's row
  // in the negative-VA table. Notes are keyed by student.idx and live only
  // in this component's state - see the footer note in the UI about them
  // not persisting beyond the browser tab.
  function updateNote(idx, field, value) {
    setNotes((prev) => ({ ...prev, [idx]: { ...prev[idx], [field]: value } }));
  }

  /* ---------------- Word export ---------------- */

  // Escapes text that gets interpolated into the exported HTML/Word document,
  // since it can come from uploaded spreadsheet cells or free-typed notes -
  // without this, a student name or note containing < > & " could corrupt
  // the document's markup.
  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (c) => (
      { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
    ));
  }

  // Turns the dashboard's numbers into a few plain-English sentences per
  // section, so the exported report isn't just raw data - it gives the
  // teacher a starting read on what the numbers mean. Deliberately written
  // in advisory language ("may be worth", "consider") rather than firm
  // conclusions, since a VA figure alone can't explain *why* something
  // happened - it's a prompt to look closer, not a verdict.
  function buildInsights() {
    const cohortVA = subgroupStats[0] ? subgroupStats[0].va : null;
    const summary = [];
    if (cohortVA === null) {
      summary.push("Value added couldn't be calculated for this cohort — check that CAT4 target and actual grades are present.");
    } else if (cohortVA > 0.15) {
      summary.push(`On average, students outperformed their CAT4 targets by ${fmtVA(cohortVA)} grades — a positive result for this cohort.`);
    } else if (cohortVA < -0.15) {
      summary.push(`On average, students underperformed their CAT4 targets by ${fmtVA(Math.abs(cohortVA))} grades. It may be worth reviewing teaching, intervention, or target-setting for this cohort.`);
    } else {
      summary.push(`Students performed broadly in line with their CAT4 targets overall (cohort VA of ${fmtVA(cohortVA)}).`);
    }
    if (passRate !== null) {
      summary.push(`${passRate}% of entries met the pass threshold for ${QUAL_LABELS[qualType]}.`);
    }
    if (skippedCount > 0) {
      summary.push(`${skippedCount} row${skippedCount === 1 ? " was" : "s were"} excluded from these figures because ${skippedCount === 1 ? "it didn't" : "they didn't"} contain a valid grade.`);
    }

    const byClassInsights = [];
    if (mapping.cls && byClass.length) {
      const negClasses = byClass.filter((c) => c.va !== null && c.va < 0);
      byClassInsights.push(
        negClasses.length === 0
          ? "All classes met or exceeded their CAT4 targets on average."
          : `${negClasses.length} of ${byClass.length} class${byClass.length === 1 ? "" : "es"} ${negClasses.length === 1 ? "has" : "have"} a negative VA: ${negClasses.map((c) => c.cls).join(", ")}. These are highlighted below and may be worth a closer look.`
      );
    }

    const subgroupInsights = [];
    if (subgroupStats.length > 1) {
      const others = subgroupStats.slice(1).filter((g) => g.va !== null);
      const lowest = others.reduce((min, g) => (min === null || g.va < min.va ? g : min), null);
      if (lowest && cohortVA !== null && lowest.va < cohortVA - 0.3) {
        subgroupInsights.push(`The ${lowest.name} subgroup's VA (${fmtVA(lowest.va)}) is notably below the cohort average (${fmtVA(cohortVA)}) and may be worth a closer look.`);
      } else {
        subgroupInsights.push("No subgroup shows a VA notably different from the cohort average.");
      }
    } else {
      subgroupInsights.push("No subgroup breakdown is available for this file — gender, SEND, EAL, or more-able columns weren't provided.");
    }

    const negativeInsights = [
      negativeStudents.length === 0
        ? "No students are currently performing below their CAT4 target based on this data."
        : `${negativeStudents.length} student${negativeStudents.length === 1 ? " is" : "s are"} performing below ${negativeStudents.length === 1 ? "their" : "their"} CAT4 target, sorted worst-first below. Consider reviewing attendance, engagement, and any existing intervention plans for these students.`,
    ];

    return { summary, byClass: byClassInsights, subgroup: subgroupInsights, negative: negativeInsights };
  }

  // Builds the report as an HTML document and downloads it with a .doc
  // extension - Word, Google Docs, and LibreOffice all open HTML content
  // saved this way as a normal, editable document. This avoids needing any
  // extra library just to generate a Word file client-side.
  function exportWord() {
    const title = subjectLabel ? subjectLabel : "Subject Exam Analysis";
    const generatedDate = new Date().toLocaleDateString(undefined, { day: "numeric", month: "long", year: "numeric" });
    const cohortVA = subgroupStats[0] ? subgroupStats[0].va : null;
    const negClassCount = byClass.filter((c) => c.va !== null && c.va < 0).length;
    const insights = buildInsights();

    const gradeRows = gradeDistribution.map((d) => `
      <tr><td>${escapeHtml(d.grade)}</td><td>${d.predictedCount}</td><td>${d.actualCount}</td></tr>
    `).join("");

    const bandRow = bandStats.map((b) => `<td>${b.pct}%</td>`).join("");
    const bandHeaderRow = bandStats.map((b) => `<th>${escapeHtml(b.label)}</th>`).join("");

    const byClassSection = mapping.cls ? `
      <h2>By Class</h2>
      <table>
        <tr><th>Class</th><th>Staff</th><th>Students</th><th>VA</th></tr>
        ${byClass.map((c) => `
          <tr>
            <td>${escapeHtml(c.cls)}</td>
            <td>${escapeHtml(c.staff || "—")}</td>
            <td>${c.n}</td>
            <td class="${c.va !== null && c.va < 0 ? "neg" : "pos"}">${fmtVA(c.va)}</td>
          </tr>
        `).join("")}
      </table>
      ${insights.byClass.map((s) => `<p class="insight">${escapeHtml(s)}</p>`).join("")}
    ` : "";

    const negativeTable = negativeStudents.length ? `
      <table>
        <tr>
          <th>#</th><th>Student</th>${mapping.cls ? "<th>Class</th>" : ""}<th>CAT4 Target</th><th>Actual</th><th>VA</th>
          <th>Intervention Notes</th><th>Possible Reasons</th><th>Steps Taken</th>
        </tr>
        ${negativeStudents.map((s, i) => {
          const n = notes[s.idx] || {};
          return `
            <tr>
              <td>${i + 1}</td>
              <td>${escapeHtml(s.name)}</td>
              ${mapping.cls ? `<td>${escapeHtml(s.cls)}</td>` : ""}
              <td>${escapeHtml(rawGradeLabel(s.predictedRaw))}</td>
              <td>${escapeHtml(rawGradeLabel(s.actualRaw))}</td>
              <td class="neg">${fmtVA(s.va)}</td>
              <td>${escapeHtml(n.intervention || "")}</td>
              <td>${escapeHtml(n.reasons || "")}</td>
              <td>${escapeHtml(n.steps || "")}</td>
            </tr>
          `;
        }).join("")}
      </table>
    ` : "";

    // The xmlns attributes below are what make Word recognise this as a
    // Word document rather than a generic web page when opened.
    const html = `<!DOCTYPE html>
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(title)}</title>
        <style>
          body { font-family: Calibri, Arial, sans-serif; color: #1f2a3c; font-size: 11pt; }
          h1 { color: #1f2a3c; border-bottom: 3px solid #b8874f; padding-bottom: 6px; font-size: 20pt; }
          h2 { color: #9c6f3d; border-bottom: 1px solid #e3d9c3; padding-bottom: 3px; margin-top: 26px; font-size: 14pt; }
          p.meta { color: #6b7686; font-size: 9.5pt; }
          table { border-collapse: collapse; width: 100%; margin: 8px 0 14px; }
          th { background: #1f2a3c; color: #ffffff; padding: 5px 8px; text-align: left; font-size: 10.5pt; }
          td { padding: 5px 8px; border: 1px solid #e3d9c3; font-size: 10.5pt; }
          td.neg { color: #bd5b3d; font-weight: bold; }
          td.pos { color: #5f8362; font-weight: bold; }
          p.insight { background: #f7f2e7; border-left: 4px solid #b8874f; padding: 7px 12px; margin: 8px 0; font-size: 10.5pt; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(title)}</h1>
        <p class="meta">${QUAL_LABELS[qualType]} · Generated ${generatedDate}${fileName ? " · Source file: " + escapeHtml(fileName) : ""}</p>

        <h2>Summary</h2>
        <table>
          <tr>
            <th>Entries</th><th>Pass rate</th><th>Cohort VA</th>
            ${mapping.cls ? "<th>Classes below VA 0</th>" : ""}
            <th>Students below CAT4 target</th>
          </tr>
          <tr>
            <td>${entriesCount}</td>
            <td>${passRate === null ? "—" : passRate + "%"}</td>
            <td class="${cohortVA !== null && cohortVA < 0 ? "neg" : "pos"}">${fmtVA(cohortVA)}</td>
            ${mapping.cls ? `<td>${negClassCount} / ${byClass.length}</td>` : ""}
            <td>${negativeStudents.length}</td>
          </tr>
        </table>
        ${insights.summary.map((s) => `<p class="insight">${escapeHtml(s)}</p>`).join("")}

        <h2>Grade Distribution</h2>
        <table>
          <tr><th>Grade</th><th>CAT4 Target</th><th>Actual</th></tr>
          ${gradeRows}
        </table>
        <table>
          <tr>${bandHeaderRow}</tr>
          <tr>${bandRow}</tr>
        </table>

        ${byClassSection}

        <h2>Value Added by Subgroup</h2>
        <table>
          <tr><th>Group</th><th>VA</th><th>N</th></tr>
          ${subgroupStats.map((g) => `
            <tr>
              <td>${escapeHtml(g.name)}</td>
              <td class="${g.va !== null && g.va < 0 ? "neg" : "pos"}">${fmtVA(g.va)}</td>
              <td>${g.n}</td>
            </tr>
          `).join("")}
        </table>
        ${insights.subgroup.map((s) => `<p class="insight">${escapeHtml(s)}</p>`).join("")}

        <h2>Students Below CAT4 Target (Negative VA)</h2>
        ${insights.negative.map((s) => `<p class="insight">${escapeHtml(s)}</p>`).join("")}
        ${negativeTable}

        <p class="meta">Generated by the Subject Exam Analysis Dashboard.</p>
      </body>
      </html>
    `;

    const blob = new Blob(["\ufeff", html], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const safeName = (subjectLabel || "subject").replace(/[^a-z0-9]+/gi, "_").toLowerCase();
    a.download = `${safeName}_exam_analysis.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /* ---------------------------------------------------------------------- */
  /*  Render                                                                 */
  /* ---------------------------------------------------------------------- */

  return (
    <div className="eax-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

        .eax-root {
          --ink: #1f2a3c;
          --parchment: #f7f2e7;
          --paper: #fffdf8;
          --brass: #b8874f;
          --brass-deep: #9c6f3d;
          --sage: #5f8362;
          --sage-bg: #e8efe4;
          --clay: #bd5b3d;
          --clay-bg: #f6e6df;
          --slate: #6b7686;
          --line: #e3d9c3;
          font-family: 'Inter', system-ui, sans-serif;
          color: var(--ink);
          background: var(--parchment);
          min-height: 100%;
          padding: 0 0 60px 0;
        }
        .eax-root * { box-sizing: border-box; }
        .eax-serif { font-family: 'Source Serif 4', Georgia, serif; }
        .eax-mono { font-family: 'IBM Plex Mono', monospace; }

        .eax-header {
          background: var(--ink);
          color: var(--parchment);
          padding: 28px 32px 24px;
          display: flex;
          align-items: flex-end;
          justify-content: space-between;
          flex-wrap: wrap;
          gap: 16px;
          border-bottom: 3px solid var(--brass);
        }
        .eax-header-title { display: flex; align-items: center; gap: 14px; }
        .eax-logo-badge {
          width: 44px; height: 44px; flex-shrink: 0; background: #fff; border-radius: 8px;
          display: flex; align-items: center; justify-content: center; padding: 3px;
          border: 1px solid rgba(184,135,79,0.5);
        }
        .eax-logo-badge img { width: 100%; height: 100%; object-fit: contain; }
        .eax-header h1 { font-size: 22px; font-weight: 600; margin: 0; letter-spacing: 0.2px; }
        .eax-header .eax-sub { font-size: 12.5px; opacity: 0.72; margin-top: 2px; }
        .eax-header-actions { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }

        .eax-btn {
          display: inline-flex; align-items: center; gap: 7px;
          padding: 9px 16px; border-radius: 3px; font-size: 13.5px; font-weight: 600;
          border: 1px solid transparent; cursor: pointer; transition: all 0.15s ease;
          font-family: inherit;
        }
        .eax-btn-brass { background: var(--brass); color: #241a0f; }
        .eax-btn-brass:hover { background: #c8964f; }
        .eax-btn-ghost { background: transparent; border-color: rgba(247,242,231,0.35); color: var(--parchment); }
        .eax-btn-ghost:hover { border-color: var(--parchment); }
        .eax-btn-outline { background: var(--paper); border-color: var(--line); color: var(--ink); }
        .eax-btn-outline:hover { border-color: var(--brass); color: var(--brass-deep); }
        .eax-btn:disabled { opacity: 0.45; cursor: not-allowed; }

        .eax-body { max-width: 1180px; margin: 0 auto; padding: 32px 32px 0; }

        /* ---------- Upload stage ---------- */
        .eax-upload-wrap { max-width: 640px; margin: 60px auto; padding: 0 24px; }
        .eax-upload-hero { text-align: center; margin-bottom: 28px; }
        .eax-upload-hero h2 { font-size: 26px; margin: 0 0 8px; }
        .eax-upload-hero p { color: var(--slate); font-size: 14.5px; margin: 0; }
        .eax-dropzone {
          background: var(--paper); border: 2px dashed var(--line); border-radius: 8px;
          padding: 48px 24px; text-align: center; cursor: pointer; transition: border-color 0.15s;
        }
        .eax-dropzone:hover, .eax-dropzone.drag { border-color: var(--brass); background: #fffaf0; }
        .eax-dropzone svg { color: var(--brass); margin-bottom: 14px; }
        .eax-dropzone .eax-drop-title { font-weight: 600; font-size: 15.5px; margin-bottom: 4px; }
        .eax-dropzone .eax-drop-sub { color: var(--slate); font-size: 13px; }
        .eax-error {
          margin-top: 16px; background: var(--clay-bg); color: var(--clay);
          border: 1px solid var(--clay); border-radius: 4px; padding: 12px 14px;
          font-size: 13.5px; display: flex; gap: 8px; align-items: flex-start;
        }
        .eax-hint-list { margin-top: 24px; font-size: 12.5px; color: var(--slate); line-height: 1.7; }
        .eax-hint-list b { color: var(--ink); }

        /* ---------- Mapping stage ---------- */
        .eax-card {
          background: var(--paper); border: 1px solid var(--line); border-radius: 6px;
          padding: 24px 26px; margin-bottom: 20px;
        }
        .eax-card h3 {
          font-size: 15px; text-transform: uppercase; letter-spacing: 0.06em;
          color: var(--brass-deep); margin: 0 0 4px; font-weight: 700;
        }
        .eax-card .eax-card-desc { color: var(--slate); font-size: 13px; margin: 0 0 18px; }
        .eax-map-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 16px; }
        .eax-field label {
          display: block; font-size: 12.5px; font-weight: 600; margin-bottom: 5px; color: var(--ink);
        }
        .eax-field .req { color: var(--clay); margin-left: 3px; }
        .eax-field select, .eax-field input[type=text] {
          width: 100%; padding: 8px 10px; border: 1px solid var(--line); border-radius: 4px;
          font-size: 13.5px; background: #fff; font-family: inherit; color: var(--ink);
        }
        .eax-field select:focus, .eax-field input:focus { outline: 2px solid var(--brass); outline-offset: 1px; }
        .eax-qual-row { display: flex; gap: 22px; align-items: center; margin-top: 6px; flex-wrap: wrap; }
        .eax-radio { display: flex; align-items: center; gap: 6px; font-size: 13.5px; cursor: pointer; }
        .eax-subject-input { max-width: 320px; }

        /* ---------- Dashboard stage ---------- */
        .eax-stats-grid {
          display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
          gap: 14px; margin-bottom: 22px;
        }
        .eax-stat {
          background: var(--paper); border: 1px solid var(--line); border-top: 3px solid var(--brass);
          border-radius: 4px; padding: 14px 16px;
        }
        .eax-stat .eax-stat-label {
          font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--slate); font-weight: 600;
        }
        .eax-stat .eax-stat-val { font-size: 25px; margin-top: 4px; }

        .eax-section-title {
          display: flex; align-items: center; gap: 10px; margin: 34px 0 14px;
        }
        .eax-section-title .diag {
          width: 18px; height: 18px; flex-shrink: 0;
        }
        .eax-section-title h2 { font-size: 18px; margin: 0; }
        .eax-section-title .eax-rule { flex: 1; height: 1px; background: var(--line); }

        .eax-chart-card {
          background: var(--paper); border: 1px solid var(--line); border-radius: 6px; padding: 18px 20px 8px;
          margin-bottom: 18px;
        }
        .eax-chart-card h4 { font-size: 13.5px; margin: 0 0 4px; font-weight: 700; }
        .eax-chart-card .eax-chart-note { font-size: 12px; color: var(--slate); margin: 0 0 10px; }
        .eax-two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
        @media (max-width: 860px) { .eax-two-col { grid-template-columns: 1fr; } }

        table.eax-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        table.eax-table th {
          text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em;
          color: var(--slate); border-bottom: 2px solid var(--line); padding: 8px 10px; font-weight: 700;
        }
        table.eax-table td { padding: 8px 10px; border-bottom: 1px solid var(--line); vertical-align: middle; }
        table.eax-table tr.eax-row-warn td:first-child { box-shadow: inset 3px 0 0 var(--clay); }
        .eax-va-neg { color: var(--clay); font-weight: 700; }
        .eax-va-pos { color: var(--sage); font-weight: 700; }
        .eax-pill {
          display: inline-block; padding: 2px 9px; border-radius: 20px; font-size: 11px; font-weight: 700;
        }
        .eax-pill-warn { background: var(--clay-bg); color: var(--clay); }

        .eax-note-input {
          width: 100%; min-width: 140px; border: 1px solid var(--line); border-radius: 3px;
          padding: 5px 7px; font-size: 12.5px; font-family: inherit; background: #fffef9;
        }
        .eax-note-input:focus { outline: 2px solid var(--brass); }

        .eax-empty {
          text-align: center; color: var(--slate); font-size: 13px; padding: 26px 0;
        }
        .eax-table-scroll { overflow-x: auto; }

        .eax-footer-actions {
          display: flex; justify-content: space-between; align-items: center; margin-top: 30px;
          padding-top: 18px; border-top: 1px solid var(--line); flex-wrap: wrap; gap: 12px;
        }
        .eax-footer-note { font-size: 12px; color: var(--slate); max-width: 520px; }
      `}</style>

      {/* ------------------------------- HEADER ------------------------------- */}
      <div className="eax-header">
        <div className="eax-header-title">
          <div className="eax-logo-badge">
            <img src={BISR_LOGO} alt="British International School Riyadh" />
          </div>
          <div>
            <h1 className="eax-serif">{subjectLabel || "Subject Exam Analysis"}</h1>
            <div className="eax-sub">{fileName ? `Source file: ${fileName}` : "Value-added analysis, class by class"}</div>
          </div>
        </div>
        {stage === "dashboard" && (
          <div className="eax-header-actions">
            <button className="eax-btn eax-btn-ghost" onClick={resetAll}>
              <RotateCcw size={15} /> Start over
            </button>
            <button className="eax-btn eax-btn-brass" onClick={exportWord}>
              <Download size={15} /> Export to Word
            </button>
          </div>
        )}
      </div>

      {/* ------------------------------- UPLOAD ------------------------------- */}
      {stage === "upload" && (
        <div className="eax-upload-wrap">
          <div className="eax-upload-hero">
            <h2 className="eax-serif">Start with a class list</h2>
            <p>Upload a CSV or Excel file — one row per student. Everything stays in your browser; nothing is uploaded anywhere.</p>
          </div>
          <div
            className="eax-dropzone"
            onClick={() => fileInputRef.current?.click()}
            onDragOver={(e) => e.preventDefault()}
            onDrop={onDrop}
          >
            <Upload size={34} />
            <div className="eax-drop-title">Click to choose a file, or drag it in here</div>
            <div className="eax-drop-sub">.csv, .xlsx, or .xls</div>
            <input
              ref={fileInputRef}
              type="file"
              accept=".csv,.xlsx,.xls"
              style={{ display: "none" }}
              onChange={(e) => handleFile(e.target.files?.[0])}
            />
          </div>
          {parseError && (
            <div className="eax-error"><AlertTriangle size={16} /> {parseError}</div>
          )}
          <div className="eax-hint-list">
            <b>What to include:</b> student name/ID, a CAT4 target grade, and the actual exam grade are needed.
            Staff initials, gender, SEND, EAL, and "more able" columns are optional — include what you have, and the dashboard
            will adapt to it.
          </div>
        </div>
      )}

      {/* ------------------------------- MAPPING ------------------------------- */}
      {stage === "mapping" && (
        <div className="eax-body">
          <div className="eax-card">
            <h3>Confirm your columns</h3>
            <p className="eax-card-desc">
              We matched these automatically from your headers. Check they look right, adjust anything that isn't, then continue.
            </p>
            <div className="eax-field eax-subject-input" style={{ marginBottom: 20 }}>
              <label>Subject / class label (used in titles and the export file name)</label>
              <input
                type="text"
                placeholder="e.g. GCSE Biology — Al Waha"
                value={subjectLabel}
                onChange={(e) => setSubjectLabel(e.target.value)}
              />
            </div>
            <div className="eax-map-grid">
              {FIELD_DEFS.map((f) => (
                <div className="eax-field" key={f.key}>
                  <label>{f.label}{f.required && <span className="req">*</span>}</label>
                  <select
                    value={mapping[f.key] || ""}
                    onChange={(e) => setMapping((m) => ({ ...m, [f.key]: e.target.value }))}
                  >
                    {!f.required && <option value="">— Not in this file —</option>}
                    {f.required && <option value="">— Select a column —</option>}
                    {headers.map((h) => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          </div>

          <div className="eax-card">
            <h3>Qualification type</h3>
            <p className="eax-card-desc">
              We guessed this from the grades in your actual-grade column. AS and A Level share the same letters, so please confirm which one this is.
            </p>
            <div className="eax-qual-row">
              <label className="eax-radio">
                <input type="radio" checked={qualType === "GCSE"} onChange={() => setQualType("GCSE")} />
                GCSE (9–1 scale)
              </label>
              <label className="eax-radio">
                <input type="radio" checked={qualType === "AS"} onChange={() => setQualType("AS")} />
                AS Level (A–E scale)
              </label>
              <label className="eax-radio">
                <input type="radio" checked={qualType === "ALEVEL"} onChange={() => setQualType("ALEVEL")} />
                A Level (A*–E scale)
              </label>
            </div>
          </div>

          <div className="eax-footer-actions">
            <button className="eax-btn eax-btn-outline" onClick={resetAll}>
              <RotateCcw size={15} /> Choose a different file
            </button>
            <button
              className="eax-btn eax-btn-brass"
              disabled={!mapping.student || !mapping.cat4 || !mapping.actual}
              onClick={() => setStage("dashboard")}
            >
              Continue to dashboard <ChevronRight size={15} />
            </button>
          </div>
          {(!mapping.student || !mapping.cat4 || !mapping.actual) && (
            <p style={{ fontSize: 12.5, color: "var(--clay)", marginTop: -6 }}>
              Please map the required fields (marked *) before continuing.
            </p>
          )}
        </div>
      )}

      {/* ------------------------------- DASHBOARD ------------------------------- */}
      {stage === "dashboard" && (
        <div className="eax-body">
          <p style={{ color: "var(--slate)", fontSize: 13, marginTop: 0 }}>
            {QUAL_LABELS[qualType]}
            {skippedCount > 0 && ` · ${skippedCount} row${skippedCount === 1 ? "" : "s"} skipped (no valid grade)`}
          </p>

          {/* Stat cards */}
          <div className="eax-stats-grid">
            <div className="eax-stat">
              <div className="eax-stat-label">Entries</div>
              <div className="eax-stat-val eax-mono">{entriesCount}</div>
            </div>
            <div className="eax-stat">
              <div className="eax-stat-label">Pass rate</div>
              <div className="eax-stat-val eax-mono">{passRate === null ? "—" : `${passRate}%`}</div>
            </div>
            <div className="eax-stat">
              <div className="eax-stat-label">Cohort VA</div>
              <div className="eax-stat-val eax-mono" style={{ color: vaColor(subgroupStats[0]?.va) }}>
                {fmtVA(subgroupStats[0]?.va)}
              </div>
            </div>
            {mapping.cls && (
              <div className="eax-stat">
                <div className="eax-stat-label">Classes below VA 0</div>
                <div className="eax-stat-val eax-mono" style={{ color: byClass.some(c => c.va < 0) ? "var(--clay)" : "var(--ink)" }}>
                  {byClass.filter((c) => c.va !== null && c.va < 0).length} / {byClass.length}
                </div>
              </div>
            )}
            <div className="eax-stat">
              <div className="eax-stat-label">Students below CAT4</div>
              <div className="eax-stat-val eax-mono" style={{ color: negativeStudents.length ? "var(--clay)" : "var(--ink)" }}>
                {negativeStudents.length}
              </div>
            </div>
          </div>

          {/* Grade band table */}
          <div className="eax-chart-card">
            <h4>Grade distribution</h4>
            <p className="eax-chart-note">CAT4 target vs. actual, threshold bands for {QUAL_LABELS[qualType]}</p>
            <div style={{ display: "flex", gap: 22, flexWrap: "wrap", marginBottom: 14 }}>
              {bandStats.map((b) => (
                <div key={b.label} style={{ minWidth: 70 }}>
                  <div className="eax-mono" style={{ fontSize: 20, fontWeight: 600 }}>{b.pct}%</div>
                  <div style={{ fontSize: 11.5, color: "var(--slate)" }}>{b.label}</div>
                </div>
              ))}
            </div>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={gradeDistribution} margin={{ top: 4, right: 10, left: -10, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" vertical={false} />
                <XAxis dataKey="grade" tick={{ fontSize: 12, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false} />
                <YAxis allowDecimals={false} tick={{ fontSize: 12, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false} />
                <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 4, border: "1px solid var(--line)" }} />
                <Legend wrapperStyle={{ fontSize: 12.5 }} iconType="circle" iconSize={8} />
                <Bar dataKey="predictedCount" name="CAT4 Target" fill="var(--brass)" fillOpacity={0.55} radius={[3, 3, 0, 0]} />
                <Bar dataKey="actualCount" radius={[3, 3, 0, 0]} name="Actual">
                  {gradeDistribution.map((d, i) => (
                    <Cell key={i} fill={d.pass ? "var(--sage)" : "var(--clay)"} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* VA by class + subgroup */}
          <div className="eax-section-title">
            <svg className="diag" viewBox="0 0 18 18"><path d="M2 15 L15 3" stroke="var(--brass)" strokeWidth="2" strokeLinecap="round" /></svg>
            <h2 className="eax-serif">Value added</h2>
            <div className="eax-rule" />
          </div>

          <div className={mapping.cls ? "eax-two-col" : ""}>
            {mapping.cls && (
              <div className="eax-chart-card">
                <h4>VA by class</h4>
                <p className="eax-chart-note">Actual grade minus CAT4 target, averaged per class</p>
                <ResponsiveContainer width="100%" height={230}>
                  <BarChart data={byClass} margin={{ top: 4, right: 10, left: -10, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" vertical={false} />
                    <XAxis dataKey="cls" tick={{ fontSize: 12, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false} />
                    <YAxis tick={{ fontSize: 12, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false} />
                    <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 4, border: "1px solid var(--line)" }} />
                    <ReferenceLine y={0} stroke="var(--ink)" />
                    <Bar dataKey="va" radius={[3, 3, 0, 0]} name="VA">
                      {byClass.map((d, i) => (
                        <Cell key={i} fill={d.va !== null && d.va < 0 ? "var(--clay)" : "var(--sage)"} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
                <div className="eax-table-scroll" style={{ marginTop: 6 }}>
                  <table className="eax-table">
                    <thead>
                      <tr><th>Class</th><th>Staff</th><th>Students</th><th>VA</th></tr>
                    </thead>
                    <tbody>
                      {byClass.map((c) => (
                        <tr key={c.cls} className={c.va !== null && c.va < 0 ? "eax-row-warn" : ""}>
                          <td>{c.cls}</td>
                          <td>{c.staff || "—"}</td>
                          <td className="eax-mono">{c.n}</td>
                          <td className={c.va !== null && c.va < 0 ? "eax-va-neg" : "eax-va-pos"}>
                            {fmtVA(c.va)}
                            {c.va !== null && c.va < 0 && <span className="eax-pill eax-pill-warn" style={{ marginLeft: 8 }}>below 0</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="eax-chart-card">
              <h4>VA by subgroup</h4>
              <p className="eax-chart-note">Cohort compared with available subgroups</p>
              <ResponsiveContainer width="100%" height={230}>
                <BarChart data={subgroupStats} margin={{ top: 4, right: 10, left: -10, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" vertical={false} />
                  <XAxis dataKey="name" tick={{ fontSize: 11.5, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 12, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false} />
                  <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 4, border: "1px solid var(--line)" }} />
                  <ReferenceLine y={0} stroke="var(--ink)" />
                  <Bar dataKey="va" radius={[3, 3, 0, 0]} name="VA">
                    {subgroupStats.map((d, i) => (
                      <Cell key={i} fill={d.va !== null && d.va < 0 ? "var(--clay)" : "var(--sage)"} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Scatter */}
          <div className="eax-chart-card">
            <h4>CAT4 target vs. actual</h4>
            <p className="eax-chart-note">
              Each dot is a CAT4 target/actual pair - bigger dots mean more students share that exact result. Above the line means they beat their target.
            </p>
            <ResponsiveContainer width="100%" height={300}>
              <ScatterChart margin={{ top: 8, right: 20, left: 12, bottom: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" />
                <XAxis type="number" dataKey="x" name="CAT4 Target" domain={SCATTER_DOMAIN} ticks={SCATTER_TICKS} allowDecimals={false}
                  tick={{ fontSize: 12, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false}
                  label={{ value: "CAT4 Target", position: "insideBottom", offset: -4, fontSize: 12, fill: "var(--slate)" }} />
                <YAxis type="number" dataKey="y" name="Actual" domain={SCATTER_DOMAIN} ticks={SCATTER_TICKS} allowDecimals={false}
                  tick={{ fontSize: 12, fill: "var(--slate)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false}
                  label={{ value: "Actual", angle: -90, position: "insideLeft", offset: -2, style: { textAnchor: "middle" }, fontSize: 12, fill: "var(--slate)" }} />
                <ZAxis dataKey="count" name="Students" range={[70, 560]} />
                <ReferenceLine segment={[{ x: SCATTER_DOMAIN[0], y: SCATTER_DOMAIN[0] }, { x: SCATTER_DOMAIN[1], y: SCATTER_DOMAIN[1] }]}
                  stroke="var(--brass)" strokeDasharray="5 4" strokeWidth={1.5} />
                <Tooltip
                  content={<ScatterTooltip />}
                  cursor={{ strokeDasharray: "3 3" }}
                />
                <Scatter data={scatterData} fill="var(--brass)" fillOpacity={0.7} />
              </ScatterChart>
            </ResponsiveContainer>
          </div>

          {/* Negative VA students */}
          <div className="eax-section-title">
            <TrendingDown size={16} color="var(--clay)" />
            <h2 className="eax-serif">Students below CAT4 (negative VA)</h2>
            <div className="eax-rule" />
          </div>
          <div className="eax-chart-card">
            {negativeStudents.length ? (
              <div className="eax-table-scroll">
                <table className="eax-table">
                  <thead>
                    <tr>
                      <th>#</th><th>Student</th>{mapping.cls && <th>Class</th>}<th>CAT4 Target</th><th>Actual</th><th>VA</th>
                      <th>Intervention notes</th><th>Possible reasons</th><th>Steps taken</th>
                    </tr>
                  </thead>
                  <tbody>
                    {negativeStudents.map((s, i) => (
                      <tr key={s.idx} className="eax-row-warn">
                        <td className="eax-mono">{i + 1}</td>
                        <td>{s.name}</td>
                        {mapping.cls && <td>{s.cls}</td>}
                        <td className="eax-mono">{rawGradeLabel(s.predictedRaw)}</td>
                        <td className="eax-mono">{rawGradeLabel(s.actualRaw)}</td>
                        <td className="eax-va-neg eax-mono">{fmtVA(s.va)}</td>
                        <td>
                          <input className="eax-note-input" placeholder="e.g. attended 3/6 sessions"
                            value={notes[s.idx]?.intervention || ""}
                            onChange={(e) => updateNote(s.idx, "intervention", e.target.value)} />
                        </td>
                        <td>
                          <input className="eax-note-input" placeholder="e.g. prolonged absence"
                            value={notes[s.idx]?.reasons || ""}
                            onChange={(e) => updateNote(s.idx, "reasons", e.target.value)} />
                        </td>
                        <td>
                          <input className="eax-note-input" placeholder="e.g. 1:1 revision plan"
                            value={notes[s.idx]?.steps || ""}
                            onChange={(e) => updateNote(s.idx, "steps", e.target.value)} />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="eax-empty">
                <CheckCircle2 size={20} style={{ marginBottom: 6, color: "var(--sage)" }} /><br />
                No students below their CAT4 prediction in this file.
              </div>
            )}
          </div>

          <div className="eax-footer-actions">
            <div className="eax-footer-note">
              <Info size={13} style={{ verticalAlign: -2, marginRight: 4 }} />
              Notes typed above are included in the Word export but aren't saved anywhere else — export before closing this tab if you want to keep them.
            </div>
            <button className="eax-btn eax-btn-brass" onClick={exportWord}>
              <Download size={15} /> Export to Word
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
