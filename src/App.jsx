import * as XLSX from "xlsx";
import { useState, useCallback, useMemo, useRef } from "react";
import {
  LineChart, Line, AreaChart, Area, BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, ReferenceLine
} from "recharts";

/* ═══════════════════════════════════════════════════════════════════════════
   XER PARSER + DATA EXTRACTOR
═══════════════════════════════════════════════════════════════════════════ */
function parseXER(text) {
  const tables = {};
  let tbl = null, hdrs = [];
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trimEnd();
    if (line.startsWith("%T\t")) { tbl = line.slice(3).trim(); tables[tbl] = []; hdrs = []; }
    else if (line.startsWith("%F\t")) { hdrs = line.slice(3).trim().split("\t"); }
    else if (line.startsWith("%R\t") && tbl) {
      const vals = line.slice(3).split("\t");
      const row = {}; hdrs.forEach((h, i) => row[h] = vals[i] ?? "");
      tables[tbl].push(row);
    }
  }
  return tables;
}

const pd = (v) => { if (!v?.trim()) return null; try { return new Date(v.trim().slice(0, 16)); } catch { return null; } };
const fd = (d) => d ? d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "—";
const fds = (d) => d ? d.toLocaleDateString("en-GB", { day: "2-digit", month: "short" }) : "—";
const num = (v, fb = 0) => isNaN(parseFloat(v)) ? fb : parseFloat(v);

function extractData(tables, label) {
  const proj    = (tables["PROJECT"]  || [])[0] || {};
  const tasks   = tables["TASK"]      || [];
  const blines  = tables["TASKBASELN"]|| [];
  const preds   = tables["TASKPRED"]  || [];
  const wbsRows = tables["PROJWBS"]   || [];
  const rsrcRows= tables["RSRC"]      || [];
  const taskRsrc= tables["TASKRSRC"]  || [];

  const wbsMap = {}; wbsRows.forEach(r => wbsMap[r.wbs_id] = r.wbs_name || "");
  const rsrcMap = {}; rsrcRows.forEach(r => rsrcMap[r.rsrc_id] = { name: r.rsrc_name || r.rsrc_short_name || r.rsrc_id, type: r.rsrc_type || "" });

  const byId = {}, byCode = {};
  tasks.forEach(r => {
    const o = {
      id: r.task_id, code: r.task_code || "", name: r.task_name || "",
      wbs_id: r.wbs_id || "", wbs: wbsMap[r.wbs_id] || "",
      status: r.status_code || "",
      act_start:    pd(r.act_start_date),
      act_finish:   pd(r.act_end_date),
      early_start:  pd(r.early_start_date),
      early_finish: pd(r.early_end_date),
      late_start:   pd(r.late_start_date),
      late_finish:  pd(r.late_end_date),
      target_start: pd(r.target_start_date),
      target_finish:pd(r.target_end_date),
      phys_pct:     num(r.phys_complete_pct),
      dur_pct:      num(r.drtn_complete_pct),
      total_float:  num(r.total_float_hr_cnt) / 8,
      free_float:   num(r.free_float_hr_cnt) / 8,
      orig_dur:     num(r.orig_drtn_hr_cnt) / 8,
      remain_dur:   num(r.remain_drtn_hr_cnt) / 8,
      budg_cost:    num(r.target_cost),
      act_cost:     num(r.act_reg_cost) + num(r.act_ot_cost),
      budg_qty:     num(r.target_qty),
      act_qty:      num(r.act_reg_qty) + num(r.act_ot_qty),
    };
    byId[o.id] = o; byCode[o.code] = o;
  });

  const blMap = {};
  blines.forEach(r => {
    if (!blMap[r.task_id] || r.base_type_id < blMap[r.task_id].base_type_id)
      blMap[r.task_id] = { bl_start: pd(r.target_start_date), bl_finish: pd(r.target_end_date), bl_pct: num(r.phys_complete_pct), bl_cost: num(r.target_cost), bl_qty: num(r.target_qty) };
  });

  const rels = [];
  preds.forEach(r => {
    const s = byId[r.task_id], p = byId[r.pred_task_id];
    if (s && p) rels.push({ succ_code: s.code, pred_code: p.code, type: r.pred_type || "PR_FS", lag: num(r.lag_hr_cnt) / 8 });
  });

  // Resource assignments
  const rsrcAssign = [];
  taskRsrc.forEach(r => {
    const task = byId[r.task_id];
    if (!task) return;
    rsrcAssign.push({
      task_id: r.task_id, task_code: task.code, task_name: task.name,
      rsrc_id: r.rsrc_id,
      rsrc_name: rsrcMap[r.rsrc_id]?.name || r.rsrc_id,
      rsrc_type: rsrcMap[r.rsrc_id]?.type || "",
      target_qty: num(r.target_qty),
      act_qty:    num(r.act_reg_qty) + num(r.act_ot_qty),
      remain_qty: num(r.remain_qty),
      target_cost:num(r.target_cost),
      act_cost:   num(r.act_reg_cost) + num(r.act_ot_cost),
      remain_cost:num(r.remain_cost),
      act_start:  pd(r.act_start_date) || task.act_start,
      act_finish: pd(r.act_end_date)   || task.act_finish,
      early_start: task.early_start,
      early_finish:task.early_finish,
    });
  });

  // Derive project date range
  const allStarts  = Object.values(byId).map(t => t.act_start || t.early_start).filter(Boolean);
  const allFinishes= Object.values(byId).map(t => t.act_finish || t.early_finish).filter(Boolean);
  const projStart  = allStarts.length  ? new Date(Math.min(...allStarts))  : new Date();
  const projFinish = allFinishes.length? new Date(Math.max(...allFinishes)): new Date();

  return { label, proj, byId, byCode, blMap, rels, rsrcAssign, rsrcMap, wbsMap, projStart, projFinish };
}

/* ═══════════════════════════════════════════════════════════════════════════
   METRICS ENGINE
═══════════════════════════════════════════════════════════════════════════ */
function computeMetrics(baseline, updated, dataDate) {
  const DD = dataDate || new Date();

  const acts = Object.values(updated.byId);
  const totalActs = acts.length;
  const completed = acts.filter(a => a.status === "TK_Complete").length;
  const inProgress= acts.filter(a => a.status === "TK_Active").length;
  const notStarted= acts.filter(a => a.status === "TK_NotStart").length;

  // ── EV Calculation ──────────────────────────────────────────────────────
  // BAC = sum of all baseline budgeted costs (from baseline file)
  let BAC = 0, PV = 0, EV = 0, AC = 0;

  Object.values(updated.byId).forEach(task => {
    const bl = baseline.blMap[task.id] || updated.blMap[task.id];
    const budgCost = bl?.bl_cost || task.budg_cost || 0;
    BAC += budgCost;
    AC  += task.act_cost;

    // EV = BAC * actual % complete
    EV  += budgCost * (task.phys_pct / 100);

    // PV = BAC * planned % (how much should be done by data date based on baseline schedule)
    const blStart  = bl?.bl_start  || task.target_start;
    const blFinish = bl?.bl_finish || task.target_finish;
    if (blStart && blFinish) {
      const totalDur = blFinish - blStart;
      const elapsed  = Math.min(Math.max(DD - blStart, 0), totalDur);
      const plannedPct = totalDur > 0 ? elapsed / totalDur : 0;
      PV += budgCost * plannedPct;
    }
  });

  const SV  = EV - PV;
  const CV  = EV - AC;
  const SPI = PV > 0 ? EV / PV : 1;
  const CPI = AC > 0 ? EV / AC : 1;
  const ETC = CPI > 0 ? (BAC - EV) / CPI : (BAC - EV);
  const EAC = AC + ETC;
  const VAC = BAC - EAC;
  const PCT_COMPLETE = BAC > 0 ? (EV / BAC) * 100 : 0;

  // ── Float / Critical ────────────────────────────────────────────────────
  const criticalActs = acts.filter(a => a.total_float <= 0 && a.status !== "TK_Complete");
  const nearCritical = acts.filter(a => a.total_float > 0 && a.total_float <= 5 && a.status !== "TK_Complete");

  // ── Schedule variance in days ───────────────────────────────────────────
  let totalSlipDays = 0, slipCount = 0;
  acts.forEach(task => {
    const bl = baseline.blMap[task.id] || updated.blMap[task.id];
    const blFinish = bl?.bl_finish || task.target_finish;
    const curFinish = task.act_finish || task.early_finish;
    if (blFinish && curFinish) {
      const slip = Math.round((curFinish - blFinish) / 86400000);
      if (slip > 0) { totalSlipDays += slip; slipCount++; }
    }
  });

  // ── Project finish variance ─────────────────────────────────────────────
  const blActsArr = Object.values(baseline.byId);
  const updActsArr= Object.values(updated.byId);
  const blProjectFinish  = blActsArr.length  ? new Date(Math.max(...blActsArr.map(t=>t.act_finish||t.early_finish||t.target_finish).filter(Boolean))) : null;
  const updProjectFinish = updActsArr.length ? new Date(Math.max(...updActsArr.map(t=>t.act_finish||t.early_finish||t.target_finish).filter(Boolean))) : null;
  const projectSlip = blProjectFinish && updProjectFinish ? Math.round((updProjectFinish - blProjectFinish) / 86400000) : 0;

  // ── Lookahead (next 28 days from data date) ──────────────────────────────
  const lookaheadEnd = new Date(DD); lookaheadEnd.setDate(lookaheadEnd.getDate() + 28);
  const lookahead = acts.filter(a => {
    const s = a.act_start || a.early_start;
    const f = a.act_finish || a.early_finish;
    if (a.status === "TK_Complete") return false;
    return (s && s >= DD && s <= lookaheadEnd) || (f && f >= DD && f <= lookaheadEnd) || (s && f && s <= DD && f >= DD);
  }).sort((a, b) => {
    const sa = a.act_start || a.early_start || new Date(9e15);
    const sb = b.act_start || b.early_start || new Date(9e15);
    return sa - sb;
  });

  // ── WBS progress ───────────────────────────────────────────────────────
  const wbsProgress = {};
  acts.forEach(a => {
    const w = a.wbs || "Unassigned";
    if (!wbsProgress[w]) wbsProgress[w] = { total: 0, pctSum: 0, completed: 0, budg: 0, ev: 0 };
    wbsProgress[w].total++;
    wbsProgress[w].pctSum += a.phys_pct;
    if (a.status === "TK_Complete") wbsProgress[w].completed++;
    const bl = baseline.blMap[a.id] || updated.blMap[a.id];
    const bc = bl?.bl_cost || a.budg_cost || 0;
    wbsProgress[w].budg += bc;
    wbsProgress[w].ev   += bc * (a.phys_pct / 100);
  });

  // ── S-Curve data (weekly buckets from project start to finish) ──────────
  const sCurveData = buildSCurve(baseline, updated, DD);

  // ── Resource S-Curves ───────────────────────────────────────────────────
  const resourceCurves = buildResourceCurves(updated);

  return {
    totalActs, completed, inProgress, notStarted,
    BAC, PV, EV, AC, SV, CV, SPI, CPI, ETC, EAC, VAC, PCT_COMPLETE,
    criticalActs, nearCritical,
    totalSlipDays, slipCount, projectSlip,
    blProjectFinish, updProjectFinish,
    lookahead, wbsProgress, sCurveData, resourceCurves,
    dataDate: DD,
  };
}

function buildSCurve(baseline, updated, dataDate) {
  // Build weekly time buckets spanning both files
  const allDates = [
    ...Object.values(baseline.byId).flatMap(t => [t.act_start||t.early_start||t.target_start, t.act_finish||t.early_finish||t.target_finish]),
    ...Object.values(updated.byId).flatMap(t => [t.act_start||t.early_start||t.target_start, t.act_finish||t.early_finish||t.target_finish]),
  ].filter(Boolean);
  if (!allDates.length) return [];
  const minD = new Date(Math.min(...allDates));
  const maxD = new Date(Math.max(...allDates));
  minD.setDate(minD.getDate() - minD.getDay()); // align to week start
  const weeks = [];
  const cur = new Date(minD);
  while (cur <= maxD) { weeks.push(new Date(cur)); cur.setDate(cur.getDate() + 7); }
  if (weeks.length > 104) weeks.splice(104); // cap at 2 years

  // For each week, accumulate planned % (PV) and actual % (EV)
  const blActs = Object.values(baseline.byId);
  const updActs= Object.values(updated.byId);

  return weeks.map(weekEnd => {
    // Planned value (baseline file progress by this date)
    let pvPct = 0, evPct = 0;
    blActs.forEach(t => {
      const s = t.act_start || t.early_start || t.target_start;
      const f = t.act_finish|| t.early_finish|| t.target_finish;
      if (!s || !f) return;
      const dur = f - s;
      if (dur <= 0) return;
      const elapsed = Math.min(Math.max(weekEnd - s, 0), dur);
      pvPct += (elapsed / dur) / blActs.length * 100;
    });
    // EV (actual progress up to this date from updated)
    updActs.forEach(t => {
      const s = t.act_start || t.early_start || t.target_start;
      const f = t.act_finish|| t.early_finish|| t.target_finish;
      if (!s || !f) return;
      const dur = f - s;
      if (dur <= 0) return;
      if (weekEnd < s) return; // not started yet at this point
      const elapsed = Math.min(Math.max(weekEnd - s, 0), dur);
      // If task is complete, count full pct; otherwise proportional
      const taskPct = t.status === "TK_Complete" ? 1 : Math.min(elapsed / dur, t.phys_pct / 100);
      evPct += taskPct / updActs.length * 100;
    });
    return {
      week: fds(weekEnd),
      weekDate: weekEnd,
      PV: Math.min(pvPct, 100),
      EV: Math.min(evPct, 100),
      isDataDate: weekEnd >= dataDate && weekEnd < new Date(dataDate.getTime() + 7 * 86400000),
    };
  });
}

function buildResourceCurves(data) {
  if (!data.rsrcAssign.length) return {};
  // Get unique resource names
  const rsrcNames = [...new Set(data.rsrcAssign.map(r => r.rsrc_name))].filter(Boolean);
  const curves = {};

  rsrcNames.forEach(rsrcName => {
    const assigns = data.rsrcAssign.filter(r => r.rsrc_name === rsrcName);
    // Get date range for this resource
    const dates = assigns.flatMap(a => [a.act_start||a.early_start, a.act_finish||a.early_finish]).filter(Boolean);
    if (!dates.length) return;
    const minD = new Date(Math.min(...dates));
    const maxD = new Date(Math.max(...dates));
    minD.setDate(minD.getDate() - minD.getDay());

    const weeks = [];
    const cur = new Date(minD);
    while (cur <= maxD) { weeks.push(new Date(cur)); cur.setDate(cur.getDate() + 7); }
    if (weeks.length > 104) weeks.splice(104);

    // Cumulative planned qty and actual qty per week
    let cumPlan = 0, cumAct = 0;
    const points = weeks.map(weekEnd => {
      assigns.forEach(a => {
        const s = a.act_start || a.early_start;
        const f = a.act_finish|| a.early_finish;
        if (!s || !f) return;
        const dur = Math.max(f - s, 1);
        // Plan: distribute target_qty linearly over duration
        const planElapsed = Math.min(Math.max(weekEnd - s, 0), dur);
        const planThisWeek= (planElapsed / dur) * a.target_qty;
        // Actual: distribute act_qty up to act_finish
        const actEnd = a.act_finish || weekEnd;
        const actElapsed = Math.min(Math.max(weekEnd - s, 0), Math.min(actEnd - s, dur));
        const actThisWeek = a.target_qty > 0 ? (actElapsed / dur) * a.act_qty : 0;
        cumPlan += planThisWeek / (weeks.length || 1);
        cumAct  += actThisWeek  / (weeks.length || 1);
      });
      return { week: fds(weekEnd), Planned: Math.round(cumPlan * 10) / 10, Actual: Math.round(cumAct * 10) / 10 };
    });

    // Recalculate properly as cumulative
    let cp = 0, ca = 0;
    const totalPlan = assigns.reduce((s, a) => s + a.target_qty, 0);
    const totalAct  = assigns.reduce((s, a) => s + a.act_qty, 0);
    const corrected = weeks.map(weekEnd => {
      let weekPlan = 0, weekAct = 0;
      assigns.forEach(a => {
        const s = a.act_start || a.early_start;
        const f = a.act_finish|| a.early_finish;
        if (!s || !f || a.target_qty === 0) return;
        const dur = Math.max(f - s, 1);
        const pPrev = weekEnd.getTime() - 7*86400000;
        const p0 = Math.min(Math.max(pPrev - s.getTime(), 0), dur);
        const p1 = Math.min(Math.max(weekEnd - s, 0), dur);
        weekPlan += ((p1 - p0) / dur) * a.target_qty;
        const actEnd = f;
        const a0 = Math.min(Math.max(pPrev - s.getTime(), 0), Math.min(actEnd - s, dur));
        const a1 = Math.min(Math.max(weekEnd - s, 0), Math.min(actEnd - s, dur));
        weekAct += ((a1 - a0) / dur) * a.act_qty;
      });
      cp += weekPlan; ca += weekAct;
      return { week: fds(weekEnd), Planned: Math.round(cp * 10) / 10, Actual: Math.round(ca * 10) / 10 };
    });

    curves[rsrcName] = { points: corrected, totalPlan, totalAct, unit: assigns[0]?.rsrc_type || "qty" };
  });

  return curves;
}

/* ═══════════════════════════════════════════════════════════════════════════
   EXCEL EXPORT (via SheetJS)
═══════════════════════════════════════════════════════════════════════════ */
function exportToExcel(metrics, projectName, dataDate) {



  const wb = XLSX.utils.book_new();
  const dd = dataDate.toLocaleDateString("en-GB");

  // ── Sheet 1: Executive Summary ───────────────────────────────────────────
  const sumData = [
    [`WEEKLY PROGRESS DASHBOARD — ${projectName}`],
    [`Data Date: ${dd}`],
    [],
    ["SCHEDULE PERFORMANCE"],
    ["Overall % Complete", `${metrics.PCT_COMPLETE.toFixed(1)}%`],
    ["Schedule Performance Index (SPI)", metrics.SPI.toFixed(3)],
    ["Schedule Variance (SV)", `$${(metrics.SV/1000).toFixed(0)}k`],
    ["Project Slip vs Baseline", `${metrics.projectSlip} days`],
    ["Critical Activities", metrics.criticalActs.length],
    ["Near-Critical Activities (≤5d float)", metrics.nearCritical.length],
    [],
    ["EARNED VALUE MANAGEMENT"],
    ["Budget at Completion (BAC)", metrics.BAC.toLocaleString()],
    ["Planned Value (PV)", metrics.PV.toFixed(0)],
    ["Earned Value (EV)", metrics.EV.toFixed(0)],
    ["Actual Cost (AC)", metrics.AC.toFixed(0)],
    ["Cost Performance Index (CPI)", metrics.CPI.toFixed(3)],
    ["Cost Variance (CV)", metrics.CV.toFixed(0)],
    ["Estimate at Completion (EAC)", metrics.EAC.toFixed(0)],
    ["Variance at Completion (VAC)", metrics.VAC.toFixed(0)],
    [],
    ["ACTIVITY STATUS"],
    ["Total Activities", metrics.totalActs],
    ["Completed", metrics.completed],
    ["In Progress", metrics.inProgress],
    ["Not Started", metrics.notStarted],
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(sumData);
  ws1["!cols"] = [{ wch: 38 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Executive Summary");

  // ── Sheet 2: Lookahead ───────────────────────────────────────────────────
  const laHeaders = ["Activity ID", "Activity Name", "WBS", "Status", "Start", "Finish", "% Complete", "Float (days)"];
  const laRows = metrics.lookahead.map(a => [
    a.code, a.name, a.wbs,
    a.status === "TK_Active" ? "In Progress" : a.status === "TK_NotStart" ? "Not Started" : a.status,
    fd(a.act_start || a.early_start), fd(a.act_finish || a.early_finish),
    `${a.phys_pct}%`, a.total_float.toFixed(1)
  ]);
  const ws2 = XLSX.utils.aoa_to_sheet([
    [`28-DAY LOOKAHEAD — Data Date: ${dd}`], [], laHeaders, ...laRows
  ]);
  ws2["!cols"] = [12, 40, 25, 14, 14, 14, 12, 12].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, "28-Day Lookahead");

  // ── Sheet 3: Critical Path ───────────────────────────────────────────────
  const cpHeaders = ["Activity ID", "Activity Name", "WBS", "Early Start", "Early Finish", "Total Float", "Free Float"];
  const cpRows = metrics.criticalActs.slice(0, 100).map(a => [
    a.code, a.name, a.wbs,
    fd(a.early_start), fd(a.early_finish),
    a.total_float.toFixed(1), a.free_float.toFixed(1)
  ]);
  const ws3 = XLSX.utils.aoa_to_sheet([
    ["CRITICAL PATH ACTIVITIES"], [], cpHeaders, ...cpRows
  ]);
  ws3["!cols"] = [12, 40, 25, 14, 14, 12, 12].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws3, "Critical Path");

  // ── Sheet 4: WBS Progress ────────────────────────────────────────────────
  const wbsHeaders = ["WBS Name", "Activities", "Completed", "% Progress", "Budget", "EV"];
  const wbsRows = Object.entries(metrics.wbsProgress).map(([w, v]) => [
    w, v.total, v.completed,
    `${(v.pctSum / v.total).toFixed(1)}%`,
    v.budg.toFixed(0), v.ev.toFixed(0)
  ]);
  const ws4 = XLSX.utils.aoa_to_sheet([["WBS PROGRESS SUMMARY"], [], wbsHeaders, ...wbsRows]);
  ws4["!cols"] = [35, 12, 12, 12, 16, 16].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws4, "WBS Progress");

  XLSX.writeFile(wb, `P6_Weekly_Dashboard_${dd.replace(/\//g, "-")}.xlsx`);
}

/* ═══════════════════════════════════════════════════════════════════════════
   UI HELPERS
═══════════════════════════════════════════════════════════════════════════ */
const C = {
  bg:      "#07090f",
  surface: "#0d1117",
  card:    "#111827",
  border:  "#1f2937",
  border2: "#374151",
  text:    "#f9fafb",
  muted:   "#9ca3af",
  dim:     "#4b5563",
  gold:    "#f59e0b",
  blue:    "#3b82f6",
  sky:     "#0ea5e9",
  green:   "#10b981",
  red:     "#ef4444",
  purple:  "#8b5cf6",
  orange:  "#f97316",
};

const pill = (label, color) => (
  <span style={{ background: color + "22", color, border: `1px solid ${color}44`, borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 700, letterSpacing: "0.05em" }}>{label}</span>
);

function KpiCard({ label, value, sub, color = C.sky, icon, trend, small }) {
  return (
    <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: small ? "14px 18px" : "18px 22px", display: "flex", flexDirection: "column", gap: 6, position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 2, background: `linear-gradient(90deg, ${color}, transparent)` }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ color: C.muted, fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.08em" }}>{label}</div>
        {icon && <div style={{ fontSize: 16, opacity: 0.7 }}>{icon}</div>}
      </div>
      <div style={{ color, fontSize: small ? 24 : 30, fontWeight: 800, fontFamily: "'Courier New', monospace", lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ color: C.muted, fontSize: 12 }}>{sub}</div>}
      {trend !== undefined && (
        <div style={{ color: trend >= 1 ? C.green : trend >= 0.8 ? C.gold : C.red, fontSize: 12, fontWeight: 600 }}>
          {trend >= 1 ? "▲ On Track" : trend >= 0.8 ? "⚠ Caution" : "▼ Behind"}
        </div>
      )}
    </div>
  );
}

function SectionHeader({ title, sub }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ fontSize: 16, fontWeight: 700, color: C.text, letterSpacing: "-0.02em" }}>{title}</div>
      {sub && <div style={{ color: C.muted, fontSize: 12, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function DropZone({ label, file, onFile, color }) {
  const [drag, setDrag] = useState(false);
  const onDrop = useCallback(e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }, [onFile]);
  return (
    <label onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={onDrop}
      style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, minHeight: 130, borderRadius: 14, border: `2px dashed ${drag ? color : file ? C.green : C.border2}`, background: drag ? "#0d1a0d" : file ? "#0a1a0a" : C.card, cursor: "pointer", padding: 20, transition: "all .2s" }}>
      <input type="file" accept=".xer" style={{ display: "none" }} onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
      <div style={{ fontSize: 28 }}>{file ? "✅" : "📂"}</div>
      <div style={{ color: file ? C.green : C.muted, fontSize: 13, fontWeight: 600, textAlign: "center" }}>{file ? file.name : label}</div>
      <div style={{ color: C.dim, fontSize: 11 }}>.XER files only</div>
    </label>
  );
}

const TABS = ["Overview", "Earned Value", "S-Curves", "Resource Curves", "Critical Path", "Lookahead", "WBS Progress"];

/* ═══════════════════════════════════════════════════════════════════════════
   CUSTOM TOOLTIP
═══════════════════════════════════════════════════════════════════════════ */
function ChartTip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#1f2937", border: `1px solid ${C.border2}`, borderRadius: 8, padding: "10px 14px", fontSize: 12 }}>
      <div style={{ color: C.muted, marginBottom: 6 }}>{label}</div>
      {payload.map((p, i) => <div key={i} style={{ color: p.color, fontWeight: 600 }}>{p.name}: {typeof p.value === "number" ? p.value.toFixed(1) : p.value}</div>)}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════
   MAIN APP
═══════════════════════════════════════════════════════════════════════════ */
export default function App() {
  const [blFile, setBlFile]     = useState(null);
  const [updFile, setUpdFile]   = useState(null);
  const [dataDate, setDataDate] = useState(() => { const d = new Date(); return d.toISOString().slice(0,10); });
  const [projName, setProjName] = useState("Project");
  const [loading, setLoading]   = useState(false);
  const [error, setError]       = useState("");
  const [metrics, setMetrics]   = useState(null);
  const [blData, setBlData]     = useState(null);
  const [updData, setUpdData]   = useState(null);
  const [tab, setTab]           = useState("Overview");
  const [selRsrc, setSelRsrc]   = useState("");

  const readFile = f => new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result);
    r.onerror = rej;
    r.readAsText(f, "latin1");
  });

  const handleRun = async () => {
    if (!blFile || !updFile) return;
    setLoading(true); setError("");
    try {
      const [tA, tB] = await Promise.all([readFile(blFile).then(parseXER), readFile(updFile).then(parseXER)]);
      const dA = extractData(tA, "Baseline");
      const dB = extractData(tB, "Updated");
      const DD = new Date(dataDate);
      // Auto-fill project name
      const pn = tB["PROJECT"]?.[0]?.proj_name || tB["PROJECT"]?.[0]?.proj_short_name || projName;
      setProjName(pn);
      const m = computeMetrics(dA, dB, DD);
      setBlData(dA); setUpdData(dB); setMetrics(m);
      // Default resource selection
      const rsrcNames = Object.keys(m.resourceCurves);
      if (rsrcNames.length) setSelRsrc(rsrcNames[0]);
      setTab("Overview");
    } catch(e) { setError("Error: " + e.message); console.error(e); }
    finally { setLoading(false); }
  };

  const rsrcNames = metrics ? Object.keys(metrics.resourceCurves) : [];
  const selCurve  = metrics?.resourceCurves[selRsrc];

  const thS = { padding: "9px 14px", textAlign: "left", color: C.dim, fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", borderBottom: `1px solid ${C.border}` };
  const tdS = { padding: "9px 14px", color: C.muted, fontSize: 13, borderBottom: `1px solid ${C.border}` };

  // ── RENDER ─────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'Georgia', 'Times New Roman', serif" }}>
      {/* Load SheetJS */}


      {/* Header */}
      <div style={{ background: `linear-gradient(135deg, #0d1117 0%, #111827 100%)`, borderBottom: `1px solid ${C.border}`, padding: "18px 32px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{ width: 40, height: 40, borderRadius: 10, background: `linear-gradient(135deg, ${C.gold}, ${C.orange})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>📈</div>
          <div>
            <div style={{ fontSize: 20, fontWeight: 700, letterSpacing: "-0.03em", color: C.text }}>{projName}</div>
            <div style={{ color: C.muted, fontSize: 12 }}>Weekly Progress Dashboard · Client Report</div>
          </div>
        </div>
        {metrics && (
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <div style={{ color: C.muted, fontSize: 12 }}>Data Date: <span style={{ color: C.gold, fontWeight: 700 }}>{new Date(dataDate).toLocaleDateString("en-GB", { day:"2-digit",month:"short",year:"numeric" })}</span></div>
            <button onClick={() => exportToExcel(metrics, projName, new Date(dataDate))}
              style={{ background: `linear-gradient(135deg, #166534, #15803d)`, color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
              ⬇ Export Excel
            </button>
            <button onClick={() => { setMetrics(null); setBlData(null); setUpdData(null); setBlFile(null); setUpdFile(null); }}
              style={{ background: C.border, color: C.muted, border: "none", borderRadius: 8, padding: "8px 16px", fontSize: 12, cursor: "pointer" }}>
              ↩ Reset
            </button>
          </div>
        )}
      </div>

      <div style={{ maxWidth: 1440, margin: "0 auto", padding: "28px 24px" }}>

        {/* ── UPLOAD SCREEN ─────────────────────────────────────────────── */}
        {!metrics && (
          <div style={{ maxWidth: 700, margin: "40px auto" }}>
            <div style={{ textAlign: "center", marginBottom: 32 }}>
              <div style={{ fontSize: 28, fontWeight: 700, marginBottom: 8 }}>Weekly Progress Dashboard</div>
              <div style={{ color: C.muted, fontSize: 15 }}>Upload your Baseline and Updated XER files to generate the full client report</div>
            </div>
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 18, padding: 28 }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
                <div>
                  <div style={{ color: C.sky, fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", marginBottom: 10 }}>📁 Baseline XER</div>
                  <DropZone label="Drop Baseline XER" file={blFile} onFile={setBlFile} color={C.sky} />
                </div>
                <div>
                  <div style={{ color: C.gold, fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", marginBottom: 10 }}>📁 Updated XER</div>
                  <DropZone label="Drop Updated XER" file={updFile} onFile={setUpdFile} color={C.gold} />
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
                <div>
                  <div style={{ color: C.muted, fontSize: 12, marginBottom: 6 }}>Project Name (optional)</div>
                  <input value={projName} onChange={e => setProjName(e.target.value)} placeholder="My Project" style={{ width: "100%", background: C.surface, border: `1px solid ${C.border2}`, borderRadius: 8, padding: "9px 12px", color: C.text, fontSize: 13, boxSizing: "border-box" }} />
                </div>
                <div>
                  <div style={{ color: C.muted, fontSize: 12, marginBottom: 6 }}>Data Date</div>
                  <input type="date" value={dataDate} onChange={e => setDataDate(e.target.value)} style={{ width: "100%", background: C.surface, border: `1px solid ${C.border2}`, borderRadius: 8, padding: "9px 12px", color: C.text, fontSize: 13, boxSizing: "border-box" }} />
                </div>
              </div>
              {error && <div style={{ background: "#1a0606", border: `1px solid ${C.red}44`, borderRadius: 8, padding: "10px 14px", color: C.red, fontSize: 13, marginBottom: 14 }}>⚠ {error}</div>}
              <button onClick={handleRun} disabled={!blFile || !updFile || loading}
                style={{ width: "100%", background: blFile && updFile ? `linear-gradient(135deg, ${C.gold}, ${C.orange})` : C.border, color: blFile && updFile ? "#000" : C.dim, border: "none", borderRadius: 10, padding: "14px 0", fontSize: 15, fontWeight: 700, cursor: blFile && updFile ? "pointer" : "default", letterSpacing: "-.01em" }}>
                {loading ? "⏳ Processing XER files..." : "🚀 Generate Dashboard"}
              </button>
            </div>
          </div>
        )}

        {/* ── DASHBOARD ─────────────────────────────────────────────────── */}
        {metrics && (
          <>
            {/* Tabs */}
            <div style={{ display: "flex", gap: 2, marginBottom: 24, borderBottom: `1px solid ${C.border}`, overflowX: "auto" }}>
              {TABS.map(t => (
                <button key={t} onClick={() => setTab(t)} style={{ background: "none", border: "none", padding: "10px 16px", fontSize: 13, fontWeight: 600, color: tab === t ? C.gold : C.dim, borderBottom: `2px solid ${tab === t ? C.gold : "transparent"}`, cursor: "pointer", whiteSpace: "nowrap", marginBottom: -1 }}>
                  {t}
                </button>
              ))}
            </div>

            {/* ── OVERVIEW ──────────────────────────────────────────────── */}
            {tab === "Overview" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
                {/* Top KPIs */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12 }}>
                  <KpiCard label="Overall Progress" value={`${metrics.PCT_COMPLETE.toFixed(1)}%`} color={C.gold} icon="📊" sub="Earned vs Budget" />
                  <KpiCard label="SPI" value={metrics.SPI.toFixed(2)} color={metrics.SPI >= 1 ? C.green : metrics.SPI >= 0.8 ? C.gold : C.red} icon="⏱" trend={metrics.SPI} sub="Schedule Performance" />
                  <KpiCard label="CPI" value={metrics.CPI.toFixed(2)} color={metrics.CPI >= 1 ? C.green : metrics.CPI >= 0.8 ? C.gold : C.red} icon="💰" trend={metrics.CPI} sub="Cost Performance" />
                  <KpiCard label="Project Slip" value={`${metrics.projectSlip}d`} color={metrics.projectSlip > 0 ? C.red : metrics.projectSlip < 0 ? C.green : C.sky} icon="📅" sub={metrics.updProjectFinish ? `Forecast: ${fd(metrics.updProjectFinish)}` : ""} />
                  <KpiCard label="Critical Acts" value={metrics.criticalActs.length} color={metrics.criticalActs.length > 10 ? C.red : C.gold} icon="🔴" sub={`${metrics.nearCritical.length} near-critical`} />
                  <KpiCard label="Lookahead" value={metrics.lookahead.length} color={C.purple} icon="🔭" sub="Activities next 28 days" />
                </div>

                {/* Activity status + EV summary row */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                  {/* Activity Donut replacement: bar chart */}
                  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                    <SectionHeader title="Activity Status" sub={`${metrics.totalActs} total activities`} />
                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {[
                        ["Completed", metrics.completed, C.green],
                        ["In Progress", metrics.inProgress, C.gold],
                        ["Not Started", metrics.notStarted, C.dim],
                      ].map(([lbl, cnt, col]) => (
                        <div key={lbl}>
                          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                            <span style={{ color: C.muted, fontSize: 13 }}>{lbl}</span>
                            <span style={{ color: col, fontWeight: 700, fontFamily: "monospace", fontSize: 13 }}>{cnt} <span style={{ color: C.dim, fontWeight: 400 }}>({metrics.totalActs > 0 ? (cnt/metrics.totalActs*100).toFixed(0) : 0}%)</span></span>
                          </div>
                          <div style={{ height: 8, background: C.border, borderRadius: 99 }}>
                            <div style={{ height: "100%", width: `${metrics.totalActs > 0 ? cnt/metrics.totalActs*100 : 0}%`, background: col, borderRadius: 99, transition: "width .8s" }} />
                          </div>
                        </div>
                      ))}
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 18 }}>
                      <KpiCard label="BAC" value={`$${(metrics.BAC/1000).toFixed(0)}k`} color={C.sky} small />
                      <KpiCard label="EAC" value={`$${(metrics.EAC/1000).toFixed(0)}k`} color={metrics.EAC > metrics.BAC ? C.red : C.green} small />
                    </div>
                  </div>

                  {/* S-Curve mini */}
                  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                    <SectionHeader title="S-Curve (Schedule)" sub="Planned vs Earned % complete over time" />
                    <ResponsiveContainer width="100%" height={220}>
                      <AreaChart data={metrics.sCurveData} margin={{ top: 4, right: 8, left: -20, bottom: 0 }}>
                        <defs>
                          <linearGradient id="pvGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={C.sky} stopOpacity={0.3}/><stop offset="95%" stopColor={C.sky} stopOpacity={0}/></linearGradient>
                          <linearGradient id="evGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={C.gold} stopOpacity={0.3}/><stop offset="95%" stopColor={C.gold} stopOpacity={0}/></linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                        <XAxis dataKey="week" tick={{ fill: C.dim, fontSize: 10 }} interval="preserveStartEnd" />
                        <YAxis tick={{ fill: C.dim, fontSize: 10 }} tickFormatter={v => `${v.toFixed(0)}%`} />
                        <Tooltip content={<ChartTip />} />
                        <Legend wrapperStyle={{ fontSize: 12 }} />
                        <Area type="monotone" dataKey="PV" name="Planned" stroke={C.sky} fill="url(#pvGrad)" strokeWidth={2} dot={false} />
                        <Area type="monotone" dataKey="EV" name="Earned"  stroke={C.gold} fill="url(#evGrad)" strokeWidth={2} dot={false} />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            )}

            {/* ── EARNED VALUE ──────────────────────────────────────────── */}
            {tab === "Earned Value" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12 }}>
                  {[
                    ["BAC", `$${(metrics.BAC/1000).toFixed(0)}k`, C.sky, "Budget at Completion"],
                    ["PV",  `$${(metrics.PV/1000).toFixed(0)}k`,  C.blue, "Planned Value"],
                    ["EV",  `$${(metrics.EV/1000).toFixed(0)}k`,  C.gold, "Earned Value"],
                    ["AC",  `$${(metrics.AC/1000).toFixed(0)}k`,  C.orange, "Actual Cost"],
                    ["SPI", metrics.SPI.toFixed(3), metrics.SPI >= 1 ? C.green : C.red, "Schedule Perf. Index"],
                    ["CPI", metrics.CPI.toFixed(3), metrics.CPI >= 1 ? C.green : C.red, "Cost Perf. Index"],
                    ["SV",  `$${(metrics.SV/1000).toFixed(0)}k`,  metrics.SV >= 0 ? C.green : C.red, "Schedule Variance"],
                    ["CV",  `$${(metrics.CV/1000).toFixed(0)}k`,  metrics.CV >= 0 ? C.green : C.red, "Cost Variance"],
                    ["EAC", `$${(metrics.EAC/1000).toFixed(0)}k`, metrics.EAC > metrics.BAC ? C.red : C.green, "Estimate at Completion"],
                    ["VAC", `$${(metrics.VAC/1000).toFixed(0)}k`, metrics.VAC >= 0 ? C.green : C.red, "Variance at Completion"],
                  ].map(([l, v, c, s]) => <KpiCard key={l} label={l} value={v} color={c} sub={s} />)}
                </div>

                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                  <SectionHeader title="EV Performance Bars" />
                  <ResponsiveContainer width="100%" height={200}>
                    <BarChart data={[{ name: "Value ($k)", BAC: metrics.BAC/1000, PV: metrics.PV/1000, EV: metrics.EV/1000, AC: metrics.AC/1000 }]} margin={{ top: 4, right: 16, left: 0, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis dataKey="name" tick={{ fill: C.dim, fontSize: 12 }} />
                      <YAxis tick={{ fill: C.dim, fontSize: 11 }} tickFormatter={v => `$${v.toFixed(0)}k`} />
                      <Tooltip content={<ChartTip />} />
                      <Legend wrapperStyle={{ fontSize: 12 }} />
                      <Bar dataKey="BAC" fill={C.sky}    radius={[4,4,0,0]} />
                      <Bar dataKey="PV"  fill={C.blue}   radius={[4,4,0,0]} />
                      <Bar dataKey="EV"  fill={C.gold}   radius={[4,4,0,0]} />
                      <Bar dataKey="AC"  fill={C.orange} radius={[4,4,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20 }}>
                    <div style={{ color: C.muted, fontSize: 12, marginBottom: 12 }}>EVM INTERPRETATION</div>
                    {[
                      ["SPI = " + metrics.SPI.toFixed(2), metrics.SPI >= 1 ? "Schedule is ahead of plan" : metrics.SPI >= 0.9 ? "Slight schedule slippage" : "Significant schedule delay", metrics.SPI >= 1 ? C.green : metrics.SPI >= 0.9 ? C.gold : C.red],
                      ["CPI = " + metrics.CPI.toFixed(2), metrics.CPI >= 1 ? "Under budget — cost efficient" : metrics.CPI >= 0.9 ? "Slight cost overrun" : "Significant cost overrun", metrics.CPI >= 1 ? C.green : metrics.CPI >= 0.9 ? C.gold : C.red],
                      ["VAC = $" + (metrics.VAC/1000).toFixed(0)+"k", metrics.VAC >= 0 ? "Project expected to finish under budget" : "Project expected to finish over budget", metrics.VAC >= 0 ? C.green : C.red],
                    ].map(([title, desc, col]) => (
                      <div key={title} style={{ borderLeft: `3px solid ${col}`, paddingLeft: 12, marginBottom: 14 }}>
                        <div style={{ color: col, fontWeight: 700, fontSize: 14 }}>{title}</div>
                        <div style={{ color: C.muted, fontSize: 12 }}>{desc}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20 }}>
                    <div style={{ color: C.muted, fontSize: 12, marginBottom: 12 }}>COMPLETION FORECAST</div>
                    {[
                      ["Budget at Completion (BAC)", `$${(metrics.BAC/1000).toFixed(0)}k`],
                      ["Estimate to Complete (ETC)", `$${(metrics.ETC/1000).toFixed(0)}k`],
                      ["Estimate at Completion (EAC)", `$${(metrics.EAC/1000).toFixed(0)}k`],
                      ["Overrun / Saving (VAC)", `$${(metrics.VAC/1000).toFixed(0)}k`],
                      ["% Complete", `${metrics.PCT_COMPLETE.toFixed(1)}%`],
                    ].map(([l, v]) => (
                      <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.border}` }}>
                        <span style={{ color: C.muted, fontSize: 13 }}>{l}</span>
                        <span style={{ color: C.text, fontWeight: 700, fontFamily: "monospace", fontSize: 13 }}>{v}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* ── S-CURVES ──────────────────────────────────────────────── */}
            {tab === "S-Curves" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24 }}>
                  <SectionHeader title="Schedule S-Curve" sub="Cumulative Planned Value (PV) vs Earned Value (EV) — % complete over time" />
                  <ResponsiveContainer width="100%" height={360}>
                    <AreaChart data={metrics.sCurveData} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
                      <defs>
                        <linearGradient id="pvG2" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={C.sky} stopOpacity={0.4}/><stop offset="95%" stopColor={C.sky} stopOpacity={0}/></linearGradient>
                        <linearGradient id="evG2" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={C.gold} stopOpacity={0.4}/><stop offset="95%" stopColor={C.gold} stopOpacity={0}/></linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis dataKey="week" tick={{ fill: C.dim, fontSize: 10 }} interval={Math.floor(metrics.sCurveData.length / 10)} />
                      <YAxis tick={{ fill: C.dim, fontSize: 11 }} tickFormatter={v => `${v.toFixed(0)}%`} domain={[0, 100]} />
                      <Tooltip content={<ChartTip />} />
                      <Legend wrapperStyle={{ fontSize: 13 }} />
                      <Area type="monotone" dataKey="PV" name="Planned %" stroke={C.sky}  fill="url(#pvG2)" strokeWidth={2.5} dot={false} />
                      <Area type="monotone" dataKey="EV" name="Earned %"  stroke={C.gold} fill="url(#evG2)" strokeWidth={2.5} dot={false} />
                    </AreaChart>
                  </ResponsiveContainer>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginTop: 16 }}>
                    <div style={{ background: C.surface, borderRadius: 10, padding: "12px 16px" }}>
                      <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase" }}>Planned at Data Date</div>
                      <div style={{ color: C.sky, fontSize: 22, fontWeight: 800, fontFamily: "monospace" }}>{metrics.PV > 0 && metrics.BAC > 0 ? (metrics.PV/metrics.BAC*100).toFixed(1) : "0.0"}%</div>
                    </div>
                    <div style={{ background: C.surface, borderRadius: 10, padding: "12px 16px" }}>
                      <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase" }}>Earned at Data Date</div>
                      <div style={{ color: C.gold, fontSize: 22, fontWeight: 800, fontFamily: "monospace" }}>{metrics.PCT_COMPLETE.toFixed(1)}%</div>
                    </div>
                    <div style={{ background: C.surface, borderRadius: 10, padding: "12px 16px" }}>
                      <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase" }}>Variance</div>
                      <div style={{ color: metrics.PCT_COMPLETE >= metrics.PV/metrics.BAC*100 ? C.green : C.red, fontSize: 22, fontWeight: 800, fontFamily: "monospace" }}>
                        {metrics.BAC > 0 ? ((metrics.PCT_COMPLETE - metrics.PV/metrics.BAC*100)).toFixed(1) : "0.0"}%
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* ── RESOURCE S-CURVES ─────────────────────────────────────── */}
            {tab === "Resource Curves" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20, flexWrap: "wrap", gap: 12 }}>
                    <SectionHeader title="Resource S-Curve" sub="Cumulative planned vs actual resource quantities" />
                    {rsrcNames.length > 0 ? (
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <label style={{ color: C.muted, fontSize: 12 }}>Select Resource:</label>
                        <select value={selRsrc} onChange={e => setSelRsrc(e.target.value)}
                          style={{ background: C.surface, border: `1px solid ${C.border2}`, borderRadius: 8, padding: "7px 12px", color: C.text, fontSize: 13, cursor: "pointer", minWidth: 200 }}>
                          {rsrcNames.map(n => <option key={n} value={n}>{n}</option>)}
                        </select>
                      </div>
                    ) : (
                      <div style={{ color: C.dim, fontSize: 13 }}>No resource assignments found in this XER</div>
                    )}
                  </div>

                  {selCurve ? (
                    <>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 20 }}>
                        <div style={{ background: C.surface, borderRadius: 10, padding: "12px 16px" }}>
                          <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase" }}>Planned Total</div>
                          <div style={{ color: C.sky, fontSize: 22, fontWeight: 800, fontFamily: "monospace" }}>{selCurve.totalPlan.toFixed(0)}</div>
                        </div>
                        <div style={{ background: C.surface, borderRadius: 10, padding: "12px 16px" }}>
                          <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase" }}>Actual Total</div>
                          <div style={{ color: C.gold, fontSize: 22, fontWeight: 800, fontFamily: "monospace" }}>{selCurve.totalAct.toFixed(0)}</div>
                        </div>
                        <div style={{ background: C.surface, borderRadius: 10, padding: "12px 16px" }}>
                          <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase" }}>Variance</div>
                          <div style={{ color: selCurve.totalAct <= selCurve.totalPlan ? C.green : C.red, fontSize: 22, fontWeight: 800, fontFamily: "monospace" }}>
                            {(selCurve.totalAct - selCurve.totalPlan).toFixed(0)}
                          </div>
                        </div>
                      </div>
                      <ResponsiveContainer width="100%" height={340}>
                        <LineChart data={selCurve.points} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                          <XAxis dataKey="week" tick={{ fill: C.dim, fontSize: 10 }} interval={Math.max(1, Math.floor(selCurve.points.length / 12))} />
                          <YAxis tick={{ fill: C.dim, fontSize: 11 }} />
                          <Tooltip content={<ChartTip />} />
                          <Legend wrapperStyle={{ fontSize: 13 }} />
                          <Line type="monotone" dataKey="Planned" stroke={C.sky}  strokeWidth={2.5} dot={false} strokeDasharray="6 3" />
                          <Line type="monotone" dataKey="Actual"  stroke={C.gold} strokeWidth={2.5} dot={false} />
                        </LineChart>
                      </ResponsiveContainer>
                    </>
                  ) : (
                    <div style={{ textAlign: "center", padding: 60, color: C.dim }}>
                      <div style={{ fontSize: 36, marginBottom: 12 }}>📊</div>
                      <div>No resource data found in this XER file.</div>
                      <div style={{ fontSize: 12, marginTop: 8 }}>Resource assignments require TASKRSRC and RSRC tables in the XER export.</div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* ── CRITICAL PATH ─────────────────────────────────────────── */}
            {tab === "Critical Path" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
                  <KpiCard label="Critical Activities" value={metrics.criticalActs.length} color={C.red} icon="🔴" sub="Total Float ≤ 0 days" />
                  <KpiCard label="Near-Critical" value={metrics.nearCritical.length} color={C.gold} icon="🟡" sub="Total Float 1–5 days" />
                  <KpiCard label="Project Slip" value={`${metrics.projectSlip > 0 ? "+" : ""}${metrics.projectSlip}d`} color={metrics.projectSlip > 0 ? C.red : C.green} icon="📅" sub={`Forecast: ${fd(metrics.updProjectFinish)}`} />
                </div>

                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                  <SectionHeader title="Critical Activities (Float ≤ 0)" sub={`${metrics.criticalActs.length} activities on critical path`} />
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse" }}>
                      <thead><tr style={{ background: C.surface }}>
                        {["Activity ID","Activity Name","WBS","Early Start","Early Finish","Total Float","Free Float"].map((h, i) => (
                          <th key={i} style={thS}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {metrics.criticalActs.length === 0 && <tr><td colSpan={7} style={{ ...tdS, textAlign: "center", padding: 32 }}>No critical activities found</td></tr>}
                        {metrics.criticalActs.slice(0, 60).map((a, i) => (
                          <tr key={i} style={{ background: i % 2 === 0 ? "transparent" : C.surface }}>
                            <td style={{ ...tdS, color: C.red, fontFamily: "monospace", fontWeight: 700 }}>{a.code}</td>
                            <td style={{ ...tdS, maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</td>
                            <td style={{ ...tdS, color: C.dim }}>{a.wbs}</td>
                            <td style={{ ...tdS, textAlign: "center" }}>{fd(a.early_start)}</td>
                            <td style={{ ...tdS, textAlign: "center" }}>{fd(a.early_finish)}</td>
                            <td style={{ ...tdS, textAlign: "center", color: C.red, fontWeight: 700, fontFamily: "monospace" }}>{a.total_float.toFixed(1)}</td>
                            <td style={{ ...tdS, textAlign: "center", fontFamily: "monospace" }}>{a.free_float.toFixed(1)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {metrics.nearCritical.length > 0 && (
                  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                    <SectionHeader title="Near-Critical Activities (Float 1–5 days)" sub="Monitor closely — at risk of becoming critical" />
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <thead><tr style={{ background: C.surface }}>
                          {["Activity ID","Activity Name","WBS","Early Finish","Total Float"].map((h, i) => <th key={i} style={thS}>{h}</th>)}
                        </tr></thead>
                        <tbody>
                          {metrics.nearCritical.slice(0, 30).map((a, i) => (
                            <tr key={i} style={{ background: i % 2 === 0 ? "transparent" : C.surface }}>
                              <td style={{ ...tdS, color: C.gold, fontFamily: "monospace", fontWeight: 700 }}>{a.code}</td>
                              <td style={{ ...tdS }}>{a.name}</td>
                              <td style={{ ...tdS, color: C.dim }}>{a.wbs}</td>
                              <td style={{ ...tdS, textAlign: "center" }}>{fd(a.early_finish)}</td>
                              <td style={{ ...tdS, textAlign: "center", color: C.gold, fontWeight: 700, fontFamily: "monospace" }}>{a.total_float.toFixed(1)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ── LOOKAHEAD ─────────────────────────────────────────────── */}
            {tab === "Lookahead" && (
              <div>
                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                  <SectionHeader title={`28-Day Lookahead`} sub={`${metrics.lookahead.length} activities from ${fd(metrics.dataDate)} to ${fd(new Date(metrics.dataDate.getTime() + 28*86400000))}`} />
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse" }}>
                      <thead><tr style={{ background: C.surface }}>
                        {["Activity ID","Activity Name","WBS","Status","Start","Finish","% Done","Float (d)"].map((h, i) => (
                          <th key={i} style={thS}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {metrics.lookahead.length === 0 && <tr><td colSpan={8} style={{ ...tdS, textAlign: "center", padding: 32 }}>No upcoming activities in the next 28 days</td></tr>}
                        {metrics.lookahead.map((a, i) => {
                          const isCrit = a.total_float <= 0;
                          const isNear = a.total_float > 0 && a.total_float <= 5;
                          return (
                            <tr key={i} style={{ background: i % 2 === 0 ? "transparent" : C.surface }}>
                              <td style={{ ...tdS, color: isCrit ? C.red : isNear ? C.gold : C.sky, fontFamily: "monospace", fontWeight: 700 }}>{a.code}</td>
                              <td style={{ ...tdS, maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis" }}>{a.name}</td>
                              <td style={{ ...tdS, color: C.dim, fontSize: 12 }}>{a.wbs}</td>
                              <td style={{ ...tdS }}>
                                {a.status === "TK_Active"   ? pill("In Progress", C.gold) :
                                 a.status === "TK_NotStart" ? pill("Not Started", C.dim) :
                                 pill(a.status, C.muted)}
                              </td>
                              <td style={{ ...tdS, textAlign: "center" }}>{fd(a.act_start || a.early_start)}</td>
                              <td style={{ ...tdS, textAlign: "center" }}>{fd(a.act_finish || a.early_finish)}</td>
                              <td style={{ ...tdS, textAlign: "center" }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                  <div style={{ flex: 1, height: 6, background: C.border, borderRadius: 99 }}>
                                    <div style={{ width: `${a.phys_pct}%`, height: "100%", background: a.phys_pct >= 80 ? C.green : a.phys_pct >= 40 ? C.gold : C.red, borderRadius: 99 }} />
                                  </div>
                                  <span style={{ fontSize: 12, color: C.muted, minWidth: 32 }}>{a.phys_pct}%</span>
                                </div>
                              </td>
                              <td style={{ ...tdS, textAlign: "center", color: isCrit ? C.red : isNear ? C.gold : C.muted, fontFamily: "monospace", fontWeight: isCrit || isNear ? 700 : 400 }}>{a.total_float.toFixed(1)}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* ── WBS PROGRESS ──────────────────────────────────────────── */}
            {tab === "WBS Progress" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                  <SectionHeader title="WBS Progress Overview" />
                  <ResponsiveContainer width="100%" height={Math.max(200, Object.keys(metrics.wbsProgress).length * 44)}>
                    <BarChart layout="vertical" data={Object.entries(metrics.wbsProgress).map(([w, v]) => ({ name: w.length > 30 ? w.slice(0,30)+"…" : w, pct: parseFloat((v.pctSum / v.total).toFixed(1)), acts: v.total }))} margin={{ top: 4, right: 40, left: 20, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} horizontal={false} />
                      <XAxis type="number" domain={[0, 100]} tick={{ fill: C.dim, fontSize: 11 }} tickFormatter={v => `${v}%`} />
                      <YAxis type="category" dataKey="name" tick={{ fill: C.muted, fontSize: 12 }} width={160} />
                      <Tooltip content={<ChartTip />} />
                      <Bar dataKey="pct" name="% Complete" fill={C.gold} radius={[0, 4, 4, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
                  <SectionHeader title="WBS Detail Table" />
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead><tr style={{ background: C.surface }}>
                      {["WBS Name","Activities","Completed","% Progress","Budget","Earned Value","SPI (EV/Budget)"].map((h, i) => <th key={i} style={thS}>{h}</th>)}
                    </tr></thead>
                    <tbody>
                      {Object.entries(metrics.wbsProgress).map(([w, v], i) => {
                        const pct = v.pctSum / v.total;
                        const wSPI = v.budg > 0 ? v.ev / v.budg : 1;
                        return (
                          <tr key={i} style={{ background: i % 2 === 0 ? "transparent" : C.surface }}>
                            <td style={{ ...tdS, maxWidth: 200 }}>{w}</td>
                            <td style={{ ...tdS, textAlign: "center", fontFamily: "monospace" }}>{v.total}</td>
                            <td style={{ ...tdS, textAlign: "center", fontFamily: "monospace", color: C.green }}>{v.completed}</td>
                            <td style={{ ...tdS, textAlign: "center" }}>
                              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                <div style={{ flex: 1, height: 8, background: C.border, borderRadius: 99 }}>
                                  <div style={{ width: `${pct}%`, height: "100%", background: pct >= 80 ? C.green : pct >= 50 ? C.gold : C.red, borderRadius: 99 }} />
                                </div>
                                <span style={{ fontSize: 12, color: C.muted, minWidth: 36 }}>{pct.toFixed(0)}%</span>
                              </div>
                            </td>
                            <td style={{ ...tdS, textAlign: "right", fontFamily: "monospace" }}>${(v.budg/1000).toFixed(0)}k</td>
                            <td style={{ ...tdS, textAlign: "right", fontFamily: "monospace", color: C.gold }}>${(v.ev/1000).toFixed(0)}k</td>
                            <td style={{ ...tdS, textAlign: "center", color: wSPI >= 1 ? C.green : wSPI >= 0.8 ? C.gold : C.red, fontFamily: "monospace", fontWeight: 700 }}>{wSPI.toFixed(2)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
