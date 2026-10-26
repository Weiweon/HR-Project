"use strict";

function renderNearestSurgeryReference(text) {
  const match = String(text).match(/^(當日最接近手術時間(?:點)?)(\s+)(.+)$/);
  if (!match) return escapeHtml(text);
  return `<span class="procedure-nearest-label">${escapeHtml(match[1])}</span>${match[2]}${escapeHtml(match[3])}`;
}

const surgeryAudit = {
  rows: [],
  loaded: false,
  fileName: "",
  issueFilter: "procedure-review",
  postoperativeMinutes: 30,

  parseDate(value) {
    if (value instanceof Date) return Number.isNaN(+value) ? null : new Date(+value);
    const text = cleanText(value);
    const compact = text.match(/^(\d{3}|\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})?$/);
    const separated = text.match(/^(\d{3,4})[/-](\d{1,2})[/-](\d{1,2})[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    const parts = compact || separated;
    if (!parts) return null;
    let year = Number(parts[1]);
    if (year < 1911) year += 1911;
    const [, , m, d, h, min, sec = "0"] = parts;
    const time = new Date(year, Number(m) - 1, Number(d), Number(h), Number(min), Number(sec));
    return time.getFullYear() === year && time.getMonth() === Number(m) - 1 && time.getDate() === Number(d)
      && time.getHours() === Number(h) && time.getMinutes() === Number(min) && time.getSeconds() === Number(sec) ? time : null;
  },

  parse(workbook) {
    const parsed = [];
    let found = false;
    let invalid = 0;
    let serial = 0;
    const caseNumbers = new Map();
    workbook.SheetNames.forEach((sheetName) => {
      const grid = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: "" });
      const header = grid.slice(0, 10).findIndex((row) => /執行起日|起始時間/.test(normalizeHeader(row[19])) && /執行迄日|終止時間/.test(normalizeHeader(row[20])));
      if (header < 0) return;
      found = true;
      grid.slice(header + 1).forEach((row, index) => {
        if (!row.some((value) => cleanText(value))) return;
        const start = this.parseDate(row[19]);
        const end = this.parseDate(row[20]);
        if (!start || !end || end < start) { invalid++; return; }
        const patientId = cleanText(row[5]).toUpperCase();
        const patientName = cleanText(row[6]).replace(/\s/g, "");
        const caseKey = patientId || (patientName ? `name:${patientName}` : `row:${sheetName}:${index}`);
        if (!caseNumbers.has(caseKey)) caseNumbers.set(caseKey, ++serial);
        parsed.push({
          caseKey, caseNumber: caseNumbers.get(caseKey), start, end,
          procedure: cleanText(row[18]) || cleanText(row[17]), sourceSheet: sheetName, sourceRows: [header + index + 2],
        });
      });
    });
    if (!found) throw new Error("手術明細須在 T／U 欄提供「執行起日／執行迄日」。");
    if (!parsed.length) throw new Error("未找到有效的手術起訖時間，請檢查 T／U 欄日期格式。");
    const deduped = new Map();
    parsed.forEach((row) => {
      const key = `${row.caseKey}|${+row.start}|${+row.end}`;
      const prior = deduped.get(key);
      if (prior) {
        prior.sourceRefs.push({ sheet: row.sourceSheet, row: row.sourceRows[0] });
        prior.procedures = unique([...prior.procedures, row.procedure]);
      } else deduped.set(key, { ...row, sourceRefs: [{ sheet: row.sourceSheet, row: row.sourceRows[0] }], procedures: [row.procedure] });
    });
    return { rows: [...deduped.values()].sort((a, b) => a.start - b.start), invalid, rawCount: parsed.length };
  },

  loadWorkbook(workbook, fileName) {
    const result = this.parse(workbook);
    this.rows = result.rows;
    this.loaded = true;
    this.fileName = fileName;
    if (state.workbookName) this.rebuild();
    const status = document.getElementById("surgeryImportState");
    if (status) status.textContent = `手術／ECMO：已匯入（${result.rows.length} 組起訖）`;
    this.message(result.invalid ? `${result.invalid} 筆手術明細的 T／U 起訖無效，未納入比對。` : "", result.invalid > 0);
  },

  message(message, warning = false) {
    const element = document.getElementById("surgeryStatus");
    if (!element) return;
    element.textContent = message;
    element.hidden = !message;
    element.classList.toggle("warning", warning);
  },

  rebuild() {
    const previous = { month: state.selectedMonth, person: state.selectedPerson, risk: state.selectedRisk, date: state.selectedDateKey };
    buildAnalysis();
    if (state.months.includes(previous.month)) state.selectedMonth = previous.month;
    if (previous.person === "all" || state.people.includes(previous.person)) state.selectedPerson = previous.person;
    state.selectedRisk = previous.risk;
    state.selectedDateKey = previous.date;
    setupControls();
    dom.monthSelect.value = state.selectedMonth;
    dom.personSelect.value = state.selectedPerson;
    dom.riskSelect.value = state.selectedRisk;
    render();
  },

  candidates(record) {
    if (!record.start || !record.end) return [];
    return this.rows.filter((row) => row.end >= addMinutes(record.start, -TOLERANCE_MINUTES) && row.start <= addMinutes(record.end, TOLERANCE_MINUTES));
  },

  nearestSameDay(record) {
    if (!record.start || !record.end) return [];
    const dayStart = new Date(record.start.getFullYear(), record.start.getMonth(), record.start.getDate());
    const dayEnd = new Date(dayStart.getFullYear(), dayStart.getMonth(), dayStart.getDate() + 1);
    const rows = this.rows.filter((row) => row.start < dayEnd && (row.end > dayStart || row.start >= dayStart));
    const gap = (row) => Math.max(row.start - record.end, record.start - row.end, 0);
    const nearest = Math.min(...rows.map(gap));
    return rows.filter((row) => gap(row) === nearest);
  },

  attendanceCorroboration(record) {
    const entry = record.matchedEntry;
    const exit = record.matchedExit;
    const blockedStatuses = [
      ATTENDANCE_STATUS.LATE_OVERTIME.short,
      ATTENDANCE_STATUS.EARLY_LEAVE.short,
      ATTENDANCE_STATUS.MISSING_SIGN_IN.short,
      ATTENDANCE_STATUS.MISSING_SIGN_OUT.short,
    ];
    const hasConflict = (record.attendanceStatuses || []).some((status) => blockedStatuses.includes(status));
    const hasProxyFlag = (record.attendanceMarkers || []).some((marker) => marker.key === ATTENDANCE_MARKER.PROXY_CARD.key);
    const supported = Boolean(entry && exit
      && entry.time <= record.start
      && exit.time >= record.end
      && entry.time <= exit.time
      && !hasConflict
      && !hasProxyFlag);
    return { supported, entry, exit };
  },

  attendanceRangeNote(record, fullDate = false) {
    const attendance = this.attendanceCorroboration(record);
    if (!attendance.supported) return "";
    const format = (time) => fullDate ? formatDateTime(time) : formatCalendarTime(time, record.date);
    return `出勤簽到退佐證 ${format(attendance.entry.time)}-${format(attendance.exit.time)}`;
  },

  evaluate(record) {
    if (!this.loaded) return null;
    const base = { evidence: [], nearbyEvidence: [], risk: false, pending: true, mode: "T／U 執行時段比對" };
    if (!record.start || !record.end) return { ...base, label: "手術起訖待覆核", detail: "加班單缺少有效起訖，無法比對。" };
    const evidence = this.candidates(record);
    const first = this.rows[0]?.start;
    const last = latestTime(this.rows.map((row) => row.end));
    if (!evidence.length) {
      // Nearby times are references only; they must not contribute to coverage or clear a risk.
      const nearbyEvidence = this.nearestSameDay(record);
      if (record.end < first || record.start > last) return { ...base, nearbyEvidence, label: "超出手術資料期間", detail: "申請區間不在本次手術資料期間內，尚無法覆核。" };
      return { ...base, nearbyEvidence, risk: true, label: "無對應手術時段", detail: "未找到與加班起訖相符之執行時段，須補其他佐證。" };
    }
    const intervals = evidence.filter((row) => row.end > row.start).sort((a, b) => a.start - b.start);
    if (!intervals.length) return { ...base, evidence, label: "手術僅時間點", detail: "T／U 時間相同，僅能佐證執行時間點，無法證明整段加班。" };
    const tolerance = TOLERANCE_MINUTES * 60000;
    const merged = [];
    intervals.forEach((row) => {
      const last = merged.at(-1);
      if (last && +row.start <= last.end) last.end = Math.max(last.end, +row.end);
      else merged.push({ start: +row.start, end: +row.end });
    });
    let coveredUntil = +record.start;
    let hasCoverage = false;
    let interrupted = false;
    for (const row of merged) {
      if (row.end <= +record.start) continue;
      if (row.start >= +record.end) break;
      if (row.start > coveredUntil + tolerance) { interrupted = true; break; }
      hasCoverage = true;
      coveredUntil = Math.max(coveredUntil, row.end);
    }
    const covered = hasCoverage && !interrupted && coveredUntil >= +record.end - tolerance;
    const surgeryReachesEnd = hasCoverage && !interrupted && coveredUntil >= +record.end;
    const endGapMinutes = Math.max(0, (+record.end - coveredUntil) / 60000);
    const overlapHours = sum(merged.map((row) => Math.max(0, Math.min(row.end, +record.end) - Math.max(row.start, +record.start)))) / 36e5;
    const returnedAfterSignOut = (record.attendanceMarkers || []).some((marker) => (
      marker.key === ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.key
      && marker.overtimeStartsNearScheduledEnd
    ));
    const disconnectedAttendanceReference = !returnedAfterSignOut && interrupted
      && Boolean(record.matchedEntry && record.matchedExit);
    if ((returnedAfterSignOut || disconnectedAttendanceReference) && overlapHours > 0 && !covered) {
      const nearestEvidence = this.nearestSameDay(record);
      return {
        ...base,
        evidence: [],
        nearbyEvidence: nearestEvidence.length ? nearestEvidence : evidence,
        risk: true,
        pending: true,
        overlapHours,
        returnAfterSignOutReference: returnedAfterSignOut,
        disconnectedAttendanceReference,
        label: "無對應手術時段",
        detail: returnedAfterSignOut
          ? "本筆有簽退後返回加班紀錄，但手術時段未完整涵蓋加班申請起訖，故不列為本筆對應手術佐證；當日最接近手術時間僅供參考，工作內容仍須依其他紀錄覆核。"
          : "手術起始或中段與加班申請之間有超過 10 分鐘的佐證缺口，且另有簽到退紀錄可供核對；手術不列為本筆連續對應佐證，當日最接近手術時間僅供參考，仍須依其他紀錄覆核。",
      };
    }
    const attendance = this.attendanceCorroboration(record);
    const attendanceCorroborated = hasCoverage && !interrupted && overlapHours > 0 && attendance.supported;
    const postoperativeAssumed = !surgeryReachesEnd && attendanceCorroborated
      && endGapMinutes <= this.postoperativeMinutes;
    const clockSupportedTail = !surgeryReachesEnd && attendanceCorroborated && !postoperativeAssumed;
    const supported = surgeryReachesEnd || attendanceCorroborated;
    const exact = covered && merged.some((row) => Math.abs(row.start - record.start) <= tolerance && Math.abs(row.end - record.end) <= tolerance);
    const spans = merged.map((row) => shortRangeLabel(new Date(row.start), new Date(row.end))).join("、");
    const attendanceRange = attendance.supported
      ? `簽到 ${formatDateTime(attendance.entry.time)}（${attendance.entry.sourceLabel}），簽退 ${formatDateTime(attendance.exit.time)}（${attendance.exit.sourceLabel}）`
      : "";
    const needsAttendance = !surgeryReachesEnd && !attendanceCorroborated && hasCoverage && !interrupted;
    const postoperativeNeedsAttendance = needsAttendance && endGapMinutes <= this.postoperativeMinutes;
    const detail = postoperativeAssumed
      ? `手術連續佐證至 ${formatDateTime(new Date(coveredUntil))}；加班申請至 ${formatDateTime(record.end)}，相差 ${formatNumber(endGapMinutes, 2)} 分鐘。${attendanceRange}，簽到退均涵蓋申請起訖，故僅作術後整理推定，並非實際整理紀錄；工作內容仍須依其他紀錄確認。`
      : clockSupportedTail
        ? `手術時段 ${spans}，重疊 ${formatNumber(overlapHours, 2)} h，手術結束早於申請終點 ${formatNumber(endGapMinutes, 2)} 分鐘；${attendanceRange}，簽到退均涵蓋申請起訖，因此不單憑手術結束時間判定未工作。簽到退僅佐證出勤時間，實際工作內容仍須依其他紀錄確認。`
        : interrupted && attendance.supported
          ? `手術時段 ${spans}，實際重疊 ${formatNumber(overlapHours, 2)} h；手術起始或中段有超過 ${TOLERANCE_MINUTES} 分鐘的佐證缺口。${attendanceRange}，雖簽到退涵蓋申請起訖，仍須補足手術缺口期間的工作佐證，不以簽到退單獨認定工作內容。`
        : needsAttendance
          ? `手術連續佐證至 ${formatDateTime(new Date(coveredUntil))}；加班申請至 ${formatDateTime(record.end)}，相差 ${formatNumber(endGapMinutes, 2)} 分鐘，但缺少可涵蓋申請起訖的有效簽到退佐證；不單憑手術結束時間推定術後整理，須補其他佐證。`
          : covered ? `T／U 時段 ${spans}${exact ? "，起訖差異在 10 分鐘內。" : "，涵蓋申請；非起訖完全相同。"}執行人員仍需依其他佐證確認。`
            : `T／U 時段 ${spans}，重疊 ${formatNumber(overlapHours, 2)} h；其餘申請區間需補其他工作或術前術後作業佐證。`;
    return {
      ...base, evidence, pending: !supported, risk: !supported, overlapHours,
      postoperativeAssumed,
      postoperativeMinutes: postoperativeAssumed ? endGapMinutes : 0,
      surgeryReachesEnd,
      surgeryInterrupted: interrupted,
      attendanceCorroborated,
      attendanceClockBounded: attendance.supported,
      clockSupportedTail,
      attendanceRequired: needsAttendance,
      postoperativeNeedsAttendance,
      attendanceEntry: attendanceCorroborated ? attendance.entry : null,
      attendanceExit: attendanceCorroborated ? attendance.exit : null,
      label: postoperativeAssumed ? "術後整理推定"
        : clockSupportedTail ? "出勤簽到退佐證"
          : covered ? exact ? "手術起訖相符" : "手術時段涵蓋"
            : overlapHours > 0 ? "手術部分重疊" : "無對應手術時段",
      detail,
    };
  },

  items(record) {
    const result = record.procedureAudit;
    if (!result || (!result.risk && !result.pending)) return [];
    const times = this.timeLabels(record, true);
    const notes = times.slice(0, 2);
    if (times.length > 2) notes.push(`另 ${times.length - 2} 組執行時段`);
    const label = this.displayLabel(result);
    const tone = label === "與加班起訖時間不符" || result.risk ? "procedure-risk" : "procedure-pending";
    return [{ key: "procedure", status: tone, label, tone, text: result.detail, calendarNotes: notes }];
  },

  displayLabel(result) {
    if (!result) return "未匯入手術明細";
    return result.evidence.length && (result.risk || result.pending) ? "與加班起訖時間不符" : result.label;
  },

  acceptedCalendarNotes(record) {
    const result = record.procedureAudit;
    if (!result || result.risk || result.pending) return [];
    const times = this.timeLabels(record, true);
    const notes = times.slice(0, 2);
    if (times.length > 2) notes.push(`另 ${times.length - 2} 組執行時段`);
    const attendanceNote = this.attendanceRangeNote(record);
    if (attendanceNote) notes.push(attendanceNote);
    return notes;
  },

  timeLabels(record, fullDate = false) {
    const format = (time) => fullDate ? formatDateTime(time) : formatCalendarTime(time, record.date);
    const result = record.procedureAudit;
    const matched = result?.evidence || [];
    const rows = matched.length ? matched : result?.nearbyEvidence || [];
    const prefix = matched.length ? "手術執行" : "當日最接近手術";
    return unique(rows.map((row) => +row.start === +row.end
      ? `${prefix}時間點 ${format(row.start)}`
      : `${matched.length ? prefix : `${prefix}時間`} ${format(row.start)}-${format(row.end)}`));
  },

  renderTimeList(times) {
    if (!times.length) return "";
    return `<ul class="record-list-compact">${times.map((time) => `<li${time.startsWith("當日最接近手術") ? ' class="procedure-nearest-text"' : ""}>${time.startsWith("當日最接近手術") ? renderNearestSurgeryReference(time) : escapeHtml(time)}</li>`).join("")}</ul>`;
  },

  summaryItems(record) {
    const result = record.procedureAudit;
    if (!result) return [];
    const missingClock = record.attendanceStatuses.some((status) =>
      [ATTENDANCE_STATUS.MISSING_SIGN_IN.short, ATTENDANCE_STATUS.MISSING_SIGN_OUT.short].includes(status));
    const lines = this.timeLabels(record, true);
    if (!lines.length && missingClock) return [];
    return this.items(record).map((item) => ({ ...item, text: lines.join("；"), lines }));
  },

  detailHtml(record) {
    if (!this.loaded) return "";
    const result = record.procedureAudit;
    if (!result) return "";
    const times = this.timeLabels(record, true);
    return `<section class="record-mini-section procedure-detail"><h4>手術／ECMO 佐證</h4>${this.items(record).map((item) => statusBadgeHtml(item.label, `attendance ${item.tone}`)).join("")}
      ${!result.evidence.length || result.postoperativeAssumed || result.attendanceCorroborated || result.attendanceRequired ? `<p>${escapeHtml(result.detail)}</p>` : ""}
      ${this.renderTimeList(times)}</section>`;
  },

  exportEvidence(record) {
    const result = record.procedureAudit;
    if (!result) return "";
    const reference = (row) => `${formatDateTime(row.start)}～${formatDateTime(row.end)} ${row.sourceRefs.map((ref) => `${ref.sheet}第${ref.row}列`).join("、")}`;
    return [result.mode, result.detail, ...result.evidence.map(reference),
      ...(result.nearbyEvidence || []).map((row) => `當日最接近（非本筆對應佐證）：${reference(row)}`)].join("；");
  },
};

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("surgeryInput").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    surgeryAudit.message("正在讀取手術／ECMO 明細…");
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
      surgeryAudit.loadWorkbook(workbook, file.name);
    } catch (error) {
      surgeryAudit.message(error.message || "無法讀取手術明細，請確認格式。", true);
    } finally { event.target.value = ""; }
  });
});
