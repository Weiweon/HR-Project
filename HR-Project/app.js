"use strict";

const dom = {};
const APPROVAL_MODE = document.body?.dataset.mode === "approval";

const TOLERANCE_MINUTES = 10;
const ENTRY_LOOKBACK_HOURS = 18;
const EXIT_LOOKAHEAD_HOURS = 18;
const ATTENDANCE_SOURCE_PRIORITY = ["car", "schedule", "application"];
const CAR_AUDIT_CUTOFF_DATEKEY = "2026-05-20";

const ATTENDANCE_STATUS = {
  NORMAL: { short: "正常", long: "正常" },
  EARLY_LEAVE: { short: "提早離院", long: "提早離院" },
  LATE_OVERTIME: { short: "晚於加班", long: "簽到晚於加班申請" },
  MISSING_SIGN_IN: { short: "簽到無紀錄", long: "簽到無紀錄" },
  MISSING_SIGN_OUT: { short: "簽退無紀錄", long: "簽退無紀錄" },
  PENDING_REVIEW: { short: "待覆核", long: "差勤待覆核" },
};

const ATTENDANCE_MARKER = {
  RETURN_AFTER_SIGNOUT: {
    key: "return-after-signout",
    label: "簽退後返回加班",
  },
  PROXY_CARD: {
    key: "proxy-card",
    label: "疑似代刷卡",
  },
};

const WORK_RISK = {
  HOLIDAY_WORK: "例假日出勤",
  OVER_12_HOURS: "日總工時12小時以上",
  REST_LESS_11: "間隔未達11小時",
  OVER_46_HOURS: "月延展工時46小時以上",
  PENDING_REVIEW: "差勤待覆核",
};

const riskOptions = [
  ["all", "全部"],
  ["hasRisk", "僅看異常"],
  [ATTENDANCE_STATUS.EARLY_LEAVE.short, ATTENDANCE_STATUS.EARLY_LEAVE.short],
  [ATTENDANCE_STATUS.LATE_OVERTIME.short, ATTENDANCE_STATUS.LATE_OVERTIME.short],
  [ATTENDANCE_STATUS.MISSING_SIGN_IN.short, ATTENDANCE_STATUS.MISSING_SIGN_IN.short],
  [ATTENDANCE_STATUS.MISSING_SIGN_OUT.short, ATTENDANCE_STATUS.MISSING_SIGN_OUT.short],
  [ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.label, ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.label],
  [ATTENDANCE_MARKER.PROXY_CARD.label, ATTENDANCE_MARKER.PROXY_CARD.label],
  [WORK_RISK.HOLIDAY_WORK, WORK_RISK.HOLIDAY_WORK],
  [WORK_RISK.OVER_12_HOURS, WORK_RISK.OVER_12_HOURS],
  [WORK_RISK.REST_LESS_11, WORK_RISK.REST_LESS_11],
  [WORK_RISK.OVER_46_HOURS, WORK_RISK.OVER_46_HOURS],
];

const DAY_TYPE = {
  WORKDAY: "平日",
  RESTDAY: "休息日",
  HOLIDAY: "例假日",
  FIXED_HOLIDAY: "固定假日",
  SPECIAL: "天災選舉日",
};

const state = {
  workbookName: "",
  scheduleRows: [],
  wageRows: [],
  applicationRows: [],
  carRows: [],
  attendanceIndex: new Map(),
  dailyRows: [],
  overtimeRecords: [],
  months: [],
  people: [],
  selectedDateKey: "",
  selectedMonth: "",
  selectedPerson: "all",
  selectedRisk: "all",
  dataSourceMode: "none",
};

document.addEventListener("DOMContentLoaded", () => {
  cacheDom();
  bindEvents();
  renderEmptyCalendar();
});

function cacheDom() {
  [
    "fileInput",
    "monthSelect",
    "personSelect",
    "riskSelect",
    "exportSummaryBtn",
    "exportDetailBtn",
    "fileStatus",
    "metricHours",
    "metricApprovedHours",
    "metricPay",
    "metricAttendanceRiskDays",
    "metricWorkRiskDays",
    "calendarTitle",
    "calendarSubtitle",
    "calendarGrid",
    "detailTitle",
    "detailSubtitle",
    "detailBody",
    "riskSubtitle",
    "riskTableBody",
  ].forEach((id) => {
    dom[id] = document.getElementById(id);
  });
}

function bindEvents() {
  dom.fileInput.addEventListener("change", handleFile);
  dom.monthSelect.addEventListener("change", () => {
    state.selectedMonth = dom.monthSelect.value;
    state.selectedDateKey = "";
    render();
  });
  dom.personSelect.addEventListener("change", () => {
    state.selectedPerson = dom.personSelect.value;
    state.selectedDateKey = "";
    render();
  });
  dom.riskSelect.addEventListener("change", () => {
    state.selectedRisk = dom.riskSelect.value;
    state.selectedDateKey = "";
    render();
  });
  dom.exportSummaryBtn.addEventListener("click", exportSummaryCsv);
  dom.exportDetailBtn.addEventListener("click", exportDetailCsv);
}

async function handleFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  setStatus(`正在讀取 ${file.name}...`, false);
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    loadWorkbook(workbook, file.name);
  } catch (error) {
    console.error(error);
    setStatus("讀取 Excel 失敗，請確認檔案格式是否正確。", true);
  }
}

function loadWorkbook(workbook, fileName) {
  if (APPROVAL_MODE) return approvalView.loadWorkbook(workbook, fileName);
  const scheduleSheetName = findSheetByHeaders(workbook, ["日期", "班別"], ["排班"]);
  const wageSheetName = findSheetByHeaders(workbook, ["姓名", "年月", "時薪"], ["時薪"]);
  const carSheetName = findSheetByHeaders(workbook, ["狀態", "日期", "紀錄"], ["汽車進出紀錄"]);

  if (!scheduleSheetName) {
    setStatus("找不到排班資料，請確認 Excel 內含「排班」分頁或相同欄位。", true);
    return;
  }

  state.workbookName = fileName;
  state.scheduleRows = parseScheduleSheet(workbook.Sheets[scheduleSheetName]);
  state.wageRows = wageSheetName ? parseWageSheet(workbook.Sheets[wageSheetName]) : [];
  state.applicationRows = dedupeApplicationRows(
    parseApplicationSheets(workbook, new Set([scheduleSheetName, wageSheetName, carSheetName].filter(Boolean))),
  );
  state.carRows = carSheetName ? parseCarSheet(workbook.Sheets[carSheetName]) : [];
  assignCarPeople(state.carRows, state.scheduleRows, state.applicationRows);

  buildAnalysis();
  setupControls();
  render();

  const modeText = state.dataSourceMode === "application"
    ? "已依加班單申請重算額外加班費並重整月曆稽核。"
    : "未讀到加班單申請，先以排班加班時數提供待覆核版本。";
  setStatus(
    `已匯入 ${fileName}。排班 ${state.scheduleRows.length} 筆、有效加班單 ${state.applicationRows.length} 筆、汽車進出 ${state.carRows.length} 筆、時薪 ${state.wageRows.length} 筆。${modeText}`,
    state.dataSourceMode !== "application",
  );
}

function findSheetByHeaders(workbook, requiredHeaders, preferredNameParts = []) {
  const normalizedRequired = requiredHeaders.map(normalizeHeader);
  const candidates = workbook.SheetNames.map((name) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "", raw: false });
    const headerScore = rows.slice(0, 10).reduce((score, row) => {
      const headers = row.map(normalizeHeader);
      const matches = normalizedRequired.filter((header) => headers.includes(header)).length;
      return Math.max(score, matches);
    }, 0);
    const nameScore = preferredNameParts.some((part) => normalizeHeader(name).includes(normalizeHeader(part))) ? 2 : 0;
    return { name, score: headerScore * 10 + nameScore };
  });
  candidates.sort((a, b) => b.score - a.score);
  return candidates[0]?.score >= normalizedRequired.length * 10 ? candidates[0].name : "";
}

function parseScheduleSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  return rows
    .map((row, index) => {
      const date = parseDateValue(getField(row, ["日期"]));
      const person = cleanText(getField(row, ["電子簽章", "姓名", "出勤者"]));
      if (!date || !person) return null;

      const signInRaw = getField(row, ["簽到", "正常簽到"]);
      const signOutRaw = getField(row, ["簽退", "正常簽退"]);
      const expectedInRaw = getField(row, ["應簽到"]);
      const expectedOutRaw = getField(row, ["應簽退"]);
      const signIn = combineDateTime(date, signInRaw);
      const signOut = normalizeEndTime(signIn || combineDateTime(date, expectedInRaw), combineDateTime(date, signOutRaw));
      const expectedIn = combineDateTime(date, expectedInRaw);
      const expectedOut = normalizeEndTime(expectedIn, combineDateTime(date, expectedOutRaw));
      const otherSignIns = parseSeriesDateTimes(date, getField(row, ["其他簽到"]));
      const otherSignOuts = parseSeriesDateTimes(date, getField(row, ["其他簽退"])).map((value) => normalizeEndTime(signIn || expectedIn || date, value));
      const shift = cleanText(getField(row, ["班別"]));
      const nationalHoliday = cleanText(getField(row, ["國定假日"]));

      return {
        rowIndex: index + 2,
        person,
        date,
        dateKey: formatDateKey(date),
        month: formatMonth(date),
        weekday: cleanText(getField(row, ["星期"])),
        shift,
        type: classifyDay(shift, nationalHoliday),
        scheduledHours: parseNumber(getField(row, ["班別時數", "班工時"])),
        dailyWorkHours: parseNumber(getField(row, ["當日工時"])),
        overtimeHours: APPROVAL_MODE ? 0 : parseNumber(getField(row, ["加班時數", "延展工時"])),
        signInRaw: cleanText(signInRaw),
        signOutRaw: cleanText(signOutRaw),
        expectedInRaw: cleanText(expectedInRaw),
        expectedOutRaw: cleanText(expectedOutRaw),
        otherSignInRaw: cleanText(getField(row, ["其他簽到"])),
        otherSignOutRaw: cleanText(getField(row, ["其他簽退"])),
        signIn,
        signOut,
        expectedIn,
        expectedOut,
        otherSignIns,
        otherSignOuts,
        leaveRecord: cleanText(getField(row, ["請假紀錄"])),
        attendanceIssue: cleanText(getField(row, ["出勤異常"])),
        attendanceResult: cleanText(getField(row, ["出勤異常結果"])),
        department: cleanText(getField(row, ["部門"])),
        title: cleanText(getField(row, ["職稱"])),
        job: cleanText(getField(row, ["職務"])),
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.person.localeCompare(b.person, "zh-Hant") || a.date - b.date);
}

function parseWageSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  return rows
    .map((row) => {
      const person = cleanText(getField(row, ["姓名", "出勤者"]));
      const month = parseMonthValue(getField(row, ["年月", "月份"]));
      const hourly = parseNumber(getField(row, ["時薪"]));
      if (!person || !month) return null;
      return { person, month, hourly };
    })
    .filter(Boolean);
}

function parseApplicationSheets(workbook, ignoredSheets) {
  const applications = [];

  workbook.SheetNames.forEach((sheetName) => {
    if (ignoredSheets.has(sheetName)) return;
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
    const headerIndex = rows.findIndex((row) => {
      const headers = row.map(normalizeHeader);
      return headers.includes(normalizeHeader("出勤者")) && headers.includes(normalizeHeader("申請時數"));
    });
    if (headerIndex < 0) return;

    const nextHeaderRow = rows[headerIndex + 1] || [];
    const headerRow = combineApplicationHeaders(rows[headerIndex], nextHeaderRow);
    const indexes = getApplicationIndexes(headerRow);
    if (indexes.date < 0 || indexes.person < 0 || indexes.hours < 0) return;

    const hasSubHeader = nextHeaderRow.some((value) => {
      const normalized = normalizeHeader(value);
      return ["時薪", "金額", "起始時間", "終止時間", "進", "出"].includes(normalized);
    });

    rows.slice(headerIndex + (hasSubHeader ? 2 : 1)).forEach((row, offset) => {
      const date = parseDateValue(row[indexes.date]);
      const person = cleanText(row[indexes.person]);
      const hours = parseNumber(row[indexes.hours]);
      if (!date || !person || hours <= 0) return;

      const startRaw = indexes.startDisplay >= 0 ? row[indexes.startDisplay] : row[indexes.startNormalized];
      const endRaw = indexes.endDisplay >= 0 ? row[indexes.endDisplay] : row[indexes.endNormalized];
      const start = indexes.startNormalized >= 0
        ? combineDateTime(date, row[indexes.startNormalized])
        : combineDateTime(date, startRaw);
      const end = indexes.endNormalized >= 0
        ? combineDateTime(date, row[indexes.endNormalized])
        : combineDateTime(date, endRaw);

      const signInRaw = indexes.signIn >= 0 ? row[indexes.signIn] : "";
      const signOutRaw = indexes.signOut >= 0 ? row[indexes.signOut] : "";

      applications.push({
        sourceSheet: sheetName,
        sourceRow: headerIndex + offset + (hasSubHeader ? 3 : 2),
        date,
        dateKey: formatDateKey(date),
        month: formatMonth(date),
        person,
        department: cleanText(row[indexes.department]),
        shift: cleanText(row[indexes.shift]),
        reason: cleanText(row[indexes.reason]),
        hours,
        category: cleanText(row[indexes.category]),
        audit: cleanText(row[indexes.audit]),
        hourly: parseNumber(row[indexes.hourly]),
        amount: parseNumber(row[indexes.amount]),
        start,
        end: normalizeEndTime(start, end),
        startRaw: cleanText(startRaw),
        endRaw: cleanText(endRaw),
        normalAttendanceText: cleanText(row[indexes.normalAttendance]),
        signInRaw: cleanText(signInRaw),
        signOutRaw: cleanText(signOutRaw),
        signIns: parseSeriesDateTimes(date, signInRaw),
        signOuts: parseSeriesDateTimes(date, signOutRaw).map((value) => normalizeEndTime(start || date, value)),
        carInRaw: cleanText(row[indexes.carIn]),
        carOutRaw: cleanText(row[indexes.carOut]),
        helperCarIns: parseSeriesDateTimes(date, row[indexes.carIn]),
        helperCarOuts: parseSeriesDateTimes(date, row[indexes.carOut]).map((value) => normalizeEndTime(start || date, value)),
      });
    });
  });

  return applications.sort((a, b) => a.person.localeCompare(b.person, "zh-Hant") || a.date - b.date);
}

function combineApplicationHeaders(primaryRow, secondaryRow) {
  const groupHeaders = new Set(["加班費", "加班起訖", "汽車進出紀錄", "時數拆解"]);
  return primaryRow.map((value, index) => {
    const primary = cleanText(value);
    const secondary = cleanText(secondaryRow[index]);
    if (!primary) return secondary;
    if (groupHeaders.has(normalizeHeader(primary)) && secondary) return secondary;
    return primary;
  });
}

function getApplicationIndexes(headerRow) {
  const findHeaderIndex = (...labels) => headerRow.findIndex((value) => labels.includes(normalizeHeader(value)));

  const startLabels = headerRow
    .map((value, index) => ({ normalized: normalizeHeader(value), index }))
    .filter((item) => item.normalized === normalizeHeader("起始時間"));
  const endLabels = headerRow
    .map((value, index) => ({ normalized: normalizeHeader(value), index }))
    .filter((item) => item.normalized === normalizeHeader("終止時間"));

  const dateCandidates = headerRow
    .map((value, index) => ({ normalized: normalizeHeader(value), index }))
    .filter((item) => item.normalized === normalizeHeader("日期"));
  const personIndex = findHeaderIndex("出勤者", "姓名");
  const dateIndex = dateCandidates.filter((item) => item.index < personIndex).at(-1)?.index
    ?? findHeaderIndex("日期");

  return {
    date: dateIndex,
    department: findHeaderIndex("部門"),
    person: personIndex,
    shift: findHeaderIndex("班別"),
    hours: findHeaderIndex("申請時數"),
    reason: findHeaderIndex("事由"),
    category: findHeaderIndex("類別"),
    audit: findHeaderIndex("稽核"),
    hourly: findHeaderIndex("時薪"),
    amount: findHeaderIndex("金額"),
    normalAttendance: findHeaderIndex("正常簽到退"),
    signIn: findHeaderIndex("簽到"),
    signOut: findHeaderIndex("簽退"),
    startDisplay: startLabels[0]?.index ?? -1,
    startNormalized: startLabels[1]?.index ?? startLabels[0]?.index ?? -1,
    endDisplay: endLabels[0]?.index ?? -1,
    endNormalized: endLabels[1]?.index ?? endLabels[0]?.index ?? -1,
    carIn: findHeaderIndex("進"),
    carOut: findHeaderIndex("出"),
  };
}

function parseCarSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  return rows
    .map((row, index) => {
      const status = cleanText(getField(row, ["狀態"]));
      const date = parseDateValue(getField(row, ["日期"]));
      const person = cleanText(getField(row, ["姓名", "出勤者", "電子簽章"]));
      if (!status || !date) return null;

      const time = combineDateTime(date, getField(row, ["時間序列(供計算)", "紀錄"]));
      if (!time) return null;

      return {
        rowIndex: index + 2,
        person,
        date,
        time,
        status,
        direction: /進/.test(status) ? "entry" : /出/.test(status) ? "exit" : "",
        sourceLabel: /進/.test(status) ? "汽車進場" : /出/.test(status) ? "汽車離場" : "汽車紀錄",
        rawText: cleanText(getField(row, ["紀錄"])),
      };
    })
    .filter((row) => row && row.direction);
}

function dedupeApplicationRows(rows) {
  const map = new Map();

  rows.forEach((row) => {
    const key = [
      row.person,
      row.dateKey,
      row.shift,
      row.reason,
      row.start ? row.start.getTime() : row.startRaw,
      row.end ? row.end.getTime() : row.endRaw,
    ].join("|");

    if (!map.has(key)) {
      map.set(key, {
        ...row,
        sourceRows: [row.sourceRow],
        duplicateCount: 1,
      });
      return;
    }

    const existing = map.get(key);
    existing.duplicateCount += 1;
    existing.sourceRows.push(row.sourceRow);
    existing.hours = Math.max(existing.hours, row.hours);
    existing.hourly = existing.hourly || row.hourly;
    existing.amount = existing.amount || row.amount;
    existing.signIns = mergeDateArrays(existing.signIns, row.signIns);
    existing.signOuts = mergeDateArrays(existing.signOuts, row.signOuts);
    existing.helperCarIns = mergeDateArrays(existing.helperCarIns, row.helperCarIns);
    existing.helperCarOuts = mergeDateArrays(existing.helperCarOuts, row.helperCarOuts);
    existing.signInRaw = existing.signInRaw || row.signInRaw;
    existing.signOutRaw = existing.signOutRaw || row.signOutRaw;
    existing.carInRaw = existing.carInRaw || row.carInRaw;
    existing.carOutRaw = existing.carOutRaw || row.carOutRaw;
    existing.audit = existing.audit || row.audit;
  });

  return Array.from(map.values()).sort((a, b) => a.person.localeCompare(b.person, "zh-Hant") || a.date - b.date);
}

function assignCarPeople(carRows, scheduleRows, applicationRows) {
  const people = unique([
    ...scheduleRows.map((row) => row.person),
    ...applicationRows.map((row) => row.person),
  ]);
  if (people.length !== 1) return;
  carRows.forEach((row) => {
    if (!row.person) row.person = people[0];
  });
}

function buildAnalysis() {
  const wageMap = buildWageMap(state.wageRows);
  const scheduleByKey = new Map();
  const dailyMap = new Map();

  state.scheduleRows.forEach((row) => {
    const key = dailyKey(row.person, row.dateKey);
    if (!scheduleByKey.has(key)) scheduleByKey.set(key, row);
    if (!dailyMap.has(key)) {
      dailyMap.set(key, createDaily(row));
    } else {
      const daily = dailyMap.get(key);
      daily.scheduleRows.push(row);
      daily.scheduledHours += row.scheduledHours || row.dailyWorkHours || 0;
      daily.signIn = earliestTime([daily.signIn, row.signIn, ...row.otherSignIns]);
      daily.signOut = latestTime([daily.signOut, row.signOut, ...row.otherSignOuts]);
    }
  });

  state.attendanceIndex = buildAttendanceIndex(state.scheduleRows, state.applicationRows, state.carRows);

  const useApplications = APPROVAL_MODE || state.applicationRows.length > 0;
  state.dataSourceMode = useApplications ? "application" : "schedule";
  const records = useApplications
    ? state.applicationRows.map((app) => recordFromApplication(app, scheduleByKey, wageMap, state.attendanceIndex))
    : state.scheduleRows.filter((row) => row.overtimeHours > 0).map((row) => recordFromSchedule(row, wageMap));

  records.forEach((record) => {
    if (!record) return;
    const key = dailyKey(record.person, record.dateKey);
    if (!dailyMap.has(key)) dailyMap.set(key, createDailyFromRecord(record));
    const daily = dailyMap.get(key);
    daily.overtimeRecords.push(record);
    daily.overtimeHours += record.hours;
    daily.estimatedPay += record.estimatedPay || 0;
    daily.type = record.type || daily.type;
    daily.shift = daily.shift || record.shift;
  });

  attachDailyAttendance(dailyMap, state.attendanceIndex);
  applyScheduleRisks(dailyMap);
  applyMonthlyOvertimeRisk(records, dailyMap);
  finalizeDailyRows(dailyMap);

  state.overtimeRecords = records;
  state.dailyRows = Array.from(dailyMap.values()).sort((a, b) => a.date - b.date || a.person.localeCompare(b.person, "zh-Hant"));
  state.months = unique(state.dailyRows.map((item) => item.month)).sort();
  state.people = unique(state.dailyRows.map((item) => item.person)).sort((a, b) => a.localeCompare(b, "zh-Hant"));
  const overtimeMonths = unique(records.map((record) => record.month)).sort();
  state.selectedMonth = overtimeMonths.at(-1) || state.months.at(-1) || "";
  state.selectedPerson = "all";
  state.selectedRisk = "all";
  state.selectedDateKey = "";
}

function buildWageMap(rows) {
  const map = new Map();
  rows.forEach((row) => {
    if (row.hourly > 0) {
      map.set(`${row.person}|${row.month}`, row.hourly);
      if (!map.has(`*|${row.month}`)) map.set(`*|${row.month}`, row.hourly);
    }
  });
  return map;
}

function buildAttendanceIndex(scheduleRows, applicationRows, carRows) {
  const index = new Map();

  const pushEvent = (person, time, direction, sourceLabel, isClockRecord, note = "", sourceGroup = "", auditChannel = "clock") => {
    if (!person || !time || !direction) return;
    if (!index.has(person)) index.set(person, []);
    index.get(person).push({
      person,
      time,
      direction,
      sourceLabel,
      isClockRecord,
      note,
      sourceGroup,
      auditChannel,
    });
  };

  scheduleRows.forEach((row) => {
    pushEvent(row.person, row.signIn, "entry", "排班簽到", true, row.signInRaw, "schedule", "clock");
    pushEvent(row.person, row.signOut, "exit", "排班簽退", true, row.signOutRaw, "schedule", "clock");
    row.otherSignIns.forEach((time) => pushEvent(row.person, time, "entry", "排班其他簽到", true, row.otherSignInRaw, "schedule", "clock"));
    row.otherSignOuts.forEach((time) => pushEvent(row.person, time, "exit", "排班其他簽退", true, row.otherSignOutRaw, "schedule", "clock"));
  });

  applicationRows.forEach((row) => {
    row.signIns.forEach((time) => pushEvent(row.person, time, "entry", "加班單簽到", true, row.signInRaw, "application", "clock"));
    row.signOuts.forEach((time) => pushEvent(row.person, time, "exit", "加班單簽退", true, row.signOutRaw, "application", "clock"));
    row.helperCarIns.forEach((time) => pushEvent(row.person, time, "entry", "加班單汽車進場", false, row.carInRaw, "application", "car-helper"));
    row.helperCarOuts.forEach((time) => pushEvent(row.person, time, "exit", "加班單汽車離場", false, row.carOutRaw, "application", "car-helper"));
  });

  carRows.forEach((row) => {
    pushEvent(row.person, row.time, row.direction, row.sourceLabel, false, row.rawText, "car", "car");
  });

  index.forEach((events, person) => {
    const deduped = uniqueBy(
      events,
      (event) => `${person}|${event.direction}|${event.sourceLabel}|${event.time.getTime()}`,
    ).sort((a, b) => a.time - b.time);
    index.set(person, deduped);
  });

  return index;
}

function createDaily(row) {
  return {
    person: row.person,
    date: row.date,
    dateKey: row.dateKey,
    month: row.month,
    weekday: row.weekday,
    shift: row.shift,
    type: row.type,
    scheduledHours: row.scheduledHours || row.dailyWorkHours || 0,
    overtimeHours: 0,
    estimatedPay: 0,
    signInRaw: row.signInRaw,
    signOutRaw: row.signOutRaw,
    expectedInRaw: row.expectedInRaw,
    expectedOutRaw: row.expectedOutRaw,
    signIn: earliestTime([row.signIn, ...row.otherSignIns]),
    signOut: latestTime([row.signOut, ...row.otherSignOuts]),
    department: row.department,
    title: row.title,
    scheduleRows: [row],
    overtimeRecords: [],
    attendanceEvents: [],
    attendanceStatuses: [],
    attendanceDetails: [],
    attendanceMarkers: [],
    attendanceCalendarLines: [],
    workRisks: [],
    workRiskDetails: [],
    calendarNotes: [],
    risks: [],
    riskDetails: [],
  };
}

function createDailyFromRecord(record) {
  return {
    person: record.person,
    date: record.date,
    dateKey: record.dateKey,
    month: record.month,
    weekday: "",
    shift: record.shift,
    type: record.type,
    scheduledHours: 0,
    overtimeHours: 0,
    estimatedPay: 0,
    signInRaw: "",
    signOutRaw: "",
    expectedInRaw: "",
    expectedOutRaw: "",
    signIn: null,
    signOut: null,
    department: "",
    title: "",
    scheduleRows: [],
    overtimeRecords: [],
    attendanceEvents: [],
    attendanceStatuses: [],
    attendanceDetails: [],
    attendanceMarkers: [],
    attendanceCalendarLines: [],
    workRisks: [],
    workRiskDetails: [],
    calendarNotes: [],
    risks: [],
    riskDetails: [],
  };
}

function attachDailyAttendance(dailyMap, attendanceIndex) {
  dailyMap.forEach((daily) => {
    const events = attendanceIndex.get(daily.person) || [];
    const start = new Date(daily.date.getFullYear(), daily.date.getMonth(), daily.date.getDate());
    const end = addDays(start, 1);
    daily.attendanceEvents = events.filter((event) => event.time >= start && event.time < end);
  });
}

function recordFromApplication(app, scheduleByKey, wageMap, attendanceIndex) {
  const schedule = scheduleByKey.get(dailyKey(app.person, app.dateKey));
  const type = normalizeOvertimeType(app.category || schedule?.type || classifyDay(schedule?.shift || app.shift || "", ""));
  const hourly = app.hourly || wageMap.get(`${app.person}|${app.month}`) || (!APPROVAL_MODE && wageMap.get(`*|${app.month}`)) || 0;
  const payResult = calculateOvertimePay(app.hours, hourly, type);
  const personEvents = attendanceIndex.get(app.person) || [];
  const attendance = evaluateAttendanceRecord(app, personEvents);
  const attendanceMarkers = [
    ...detectReturnAfterSignOut(app, schedule),
    ...detectProxyCard(app, schedule, personEvents),
  ];
  const audit = auditTextToCategories(app.audit);

  const record = {
    id: `${app.person}-${app.dateKey}-${app.sourceSheet}-${app.sourceRow}`,
    source: "加班單申請",
    person: app.person,
    date: app.date,
    dateKey: app.dateKey,
    month: app.month,
    shift: schedule?.shift || app.shift || "",
    type,
    reason: app.reason,
    hours: app.hours,
    hourly,
    estimatedPay: payResult.total,
    segments: payResult.segments,
    payNote: payResult.note,
    audit: app.audit,
    start: app.start,
    end: app.end,
    startRaw: app.startRaw,
    endRaw: app.endRaw,
    normalAttendanceText: app.normalAttendanceText,
    attendanceStatuses: unique([...attendance.statuses, ...audit.attendanceStatuses]),
    attendanceDetails: unique([...attendance.details, ...audit.attendanceDetails]),
    attendanceMarkers,
    attendanceSections: attendance.sections,
    workRisks: unique(audit.workRisks),
    workRiskDetails: unique(audit.workRiskDetails),
    matchedEntry: attendance.entry,
    matchedStart: attendance.entry,
    matchedExit: attendance.exit,
    isHolidayApplication: [DAY_TYPE.HOLIDAY, DAY_TYPE.FIXED_HOLIDAY].includes(type),
  };

  if (!record.attendanceStatuses.length) {
    record.attendanceDetails = unique([
      `出勤狀況：${ATTENDANCE_STATUS.NORMAL.long}。`,
      ...record.attendanceDetails,
    ]);
  }

  if (APPROVAL_MODE) approvalView.enrichRecord(record, app, wageMap);
  refreshRecordDerived(record);
  return record;
}

function recordFromSchedule(row, wageMap) {
  const hourly = wageMap.get(`${row.person}|${row.month}`) || wageMap.get(`*|${row.month}`) || 0;
  const payResult = calculateOvertimePay(row.overtimeHours, hourly, row.type);
  const record = {
    id: `${row.person}-${row.dateKey}-schedule-${row.rowIndex}`,
    source: "排班加班時數",
    person: row.person,
    date: row.date,
    dateKey: row.dateKey,
    month: row.month,
    shift: row.shift,
    type: row.type,
    reason: row.attendanceIssue || row.attendanceResult || "排班加班時數",
    hours: row.overtimeHours,
    hourly,
    estimatedPay: payResult.total,
    segments: payResult.segments,
    payNote: payResult.note,
    audit: "待覆核",
    start: null,
    end: null,
    startRaw: "",
    endRaw: "",
    normalAttendanceText: "",
    attendanceStatuses: [],
    attendanceDetails: ["出勤狀況：待覆核。未讀到加班單申請，請再確認原始資料。"],
    attendanceMarkers: [],
    workRisks: [WORK_RISK.PENDING_REVIEW],
    workRiskDetails: ["差勤待覆核：未讀到加班單申請，先以排班加班時數呈現。"],
    matchedEntry: null,
    matchedClockIn: null,
    matchedExit: null,
    isHolidayApplication: false,
  };
  refreshRecordDerived(record);
  return record;
}

function detectReturnAfterSignOut(app, schedule) {
  if (!app?.start || !schedule?.signOut) return [];

  const candidateEvents = [
    ...((schedule.otherSignIns || []).map((time) => ({ time, sourceLabel: "排班其他簽到" }))),
    ...((app.signIns || []).map((time) => ({ time, sourceLabel: "加班單簽到" }))),
  ]
    .filter((event) => event.time > schedule.signOut)
    .sort((left, right) => left.time - right.time);

  const threshold = new Date(Math.max(addMinutes(schedule.signOut, 30).getTime(), addMinutes(app.start, TOLERANCE_MINUTES).getTime()));
  const returnSignIn = candidateEvents.find((event) => event.time <= threshold);
  if (!returnSignIn) return [];

  return [{
    key: ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.key,
    label: ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.label,
    originalSignOut: schedule.signOut,
    returnSignIn: returnSignIn.time,
    sourceLabel: returnSignIn.sourceLabel,
    countsAsRisk: false,
    note: `簽退 ${formatTimeWithSeconds(schedule.signOut)}，返回 ${formatTimeWithSeconds(returnSignIn.time)}`,
    detail: `${ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.label}：原簽退 ${formatDateTime(schedule.signOut)}，新簽到 ${formatDateTime(returnSignIn.time)}${returnSignIn.sourceLabel ? `（${returnSignIn.sourceLabel}）` : ""}`,
  }];
}

function detectProxyCard(app, schedule, personEvents) {
  if (!app?.start || !app?.end) return [];

  const windowStart = addHours(app.start, -ENTRY_LOOKBACK_HOURS);
  const windowEnd = addMinutes(app.end, TOLERANCE_MINUTES);
  const clockCandidates = [
    ...(schedule?.signIn ? [{ time: schedule.signIn, sourceLabel: "排班簽到" }] : []),
    ...((schedule?.otherSignIns || []).map((time) => ({ time, sourceLabel: "排班其他簽到" }))),
    ...((app.signIns || []).map((time) => ({ time, sourceLabel: "加班單簽到" }))),
  ]
    .filter((event) => event.time >= windowStart && event.time <= addMinutes(app.start, TOLERANCE_MINUTES))
    .sort((left, right) => right.time - left.time);
  const clockEntry = clockCandidates[0];
  if (!clockEntry) return [];

  const carEntry = personEvents
    .filter((event) => (
      event.sourceGroup === "car"
      && event.direction === "entry"
      && event.time >= clockEntry.time
      && event.time <= windowEnd
    ))
    .sort((left, right) => left.time - right.time)[0];

  if (!carEntry) return [];

  const gapMs = carEntry.time - clockEntry.time;
  if (gapMs < 30 * 60 * 1000) return [];

  return [{
    key: ATTENDANCE_MARKER.PROXY_CARD.key,
    label: ATTENDANCE_MARKER.PROXY_CARD.label,
    countsAsRisk: true,
    clockEntry: clockEntry.time,
    carEntry: carEntry.time,
    sourceLabel: clockEntry.sourceLabel,
    note: `簽到 ${formatTimeWithSeconds(clockEntry.time)}，汽車進場 ${formatTimeWithSeconds(carEntry.time)}`,
    detail: `${ATTENDANCE_MARKER.PROXY_CARD.label}：簽到時間 ${formatDateTime(clockEntry.time)}，汽車進場卻是 ${formatDateTime(carEntry.time)}`,
  }];
}

function refreshRecordDerived(record) {
  record.risks = unique([...record.attendanceStatuses, ...record.workRisks]);
  record.riskDetails = unique([...record.attendanceDetails, ...record.workRiskDetails]);
  record.attendanceLabel = record.attendanceStatuses.length
    ? record.attendanceStatuses.map(statusDisplayLabel).join("、")
    : ATTENDANCE_STATUS.NORMAL.short;
}

function evaluateAttendanceRecord(app, personEvents) {
  if (!app.start || !app.end) {
    return { statuses: [], details: [], entry: null, exit: null, sections: [] };
  }

  const section = evaluateAttendanceSection(app, personEvents);

  return {
    statuses: section.statuses,
    details: section.details,
    entry: section.entry || null,
    exit: section.exit || null,
    sections: [section],
  };
}

function evaluateAttendanceSection(app, personEvents) {
  const windowStart = addHours(app.start, -ENTRY_LOOKBACK_HOURS);
  const windowEnd = addHours(app.end, EXIT_LOOKAHEAD_HOURS);
  const scopedEvents = personEvents.filter((event) => event.time >= windowStart && event.time <= windowEnd);
  const statuses = [];
  const details = [];
  const entrySource = pickAttendanceSource(scopedEvents, "entry", app.start, app.end);
  const exitSource = pickAttendanceSource(scopedEvents, "exit", app.start, app.end);
  const entryEvents = scopedEvents.filter((event) => event.direction === "entry" && event.sourceGroup === entrySource);
  const exitEvents = scopedEvents.filter((event) => event.direction === "exit" && event.sourceGroup === exitSource);
  const entry = pickEntryCandidate(entryEvents, exitEvents, app.start, app.end, entrySource ? [entrySource] : []);
  const exitEvidence = pickExitEvidence(
    exitEvents,
    entryEvents,
    entry?.time || app.start,
    app.end,
    exitSource ? [exitSource] : [],
  );
  const exit = exitEvidence.outcome === "matched" ? exitEvidence.event : null;
  const earlyExit = exitEvidence.outcome === "early" ? exitEvidence.event : null;
  const mergedMissing = !entry && !exit && !earlyExit;

  if (mergedMissing) {
    statuses.push(ATTENDANCE_STATUS.MISSING_SIGN_IN.short, ATTENDANCE_STATUS.MISSING_SIGN_OUT.short);
    details.push(`${formatDateTime(app.start)} 起至 ${formatDateTime(app.end)} 前後未找到可佐證開始或結束加班的紀錄。`);
  } else if (!entry) {
    statuses.push(ATTENDANCE_STATUS.MISSING_SIGN_IN.short);
    details.push(`${ATTENDANCE_STATUS.MISSING_SIGN_IN.long}：${formatDateTime(app.start)} 起未找到可佐證開始加班的紀錄。`);
  }

  if (entry && entry.time > app.start) {
    statuses.push(ATTENDANCE_STATUS.LATE_OVERTIME.short);
    details.push(
      `${ATTENDANCE_STATUS.LATE_OVERTIME.long}：加班起始 ${formatDateTime(app.start)}，實際採認開始紀錄為 ${formatDateTime(entry.time)}（${entry.sourceLabel}）。`,
    );
  }

  if (exit) {
    details.push(`加班結束以 ${formatDateTime(exit.time)}（${exit.sourceLabel}）採認。`);
  } else if (earlyExit) {
    statuses.push(ATTENDANCE_STATUS.EARLY_LEAVE.short);
    details.push(
      `${ATTENDANCE_STATUS.EARLY_LEAVE.long}：加班預計至 ${formatDateTime(app.end)}，最晚離院/簽退為 ${formatDateTime(earlyExit.time)}（${earlyExit.sourceLabel}）。`,
    );
  } else if (!mergedMissing) {
    statuses.push(ATTENDANCE_STATUS.MISSING_SIGN_OUT.short);
    details.push(`${ATTENDANCE_STATUS.MISSING_SIGN_OUT.long}：${formatDateTime(app.end)} 前後未找到可佐證結束加班的紀錄。`);
  }

  if (!statuses.length) {
    details.unshift(`出勤狀況：${ATTENDANCE_STATUS.NORMAL.long}。`);
  }

  return {
    key: "final",
    title: "出勤佐證",
    statuses: unique(statuses),
    details: unique(details),
    entry,
    exit: exit || earlyExit || null,
    calendarText: buildAttendanceSectionCalendarText("出勤", unique(statuses), false),
    calendarTone: attendanceClass(unique(statuses)),
    hasEvidence: Boolean(entrySource || exitSource),
    entrySource,
    exitSource,
    mergedMissing,
  };
}

function pickAttendanceSource(events, direction, start, end) {
  for (const sourceGroup of ATTENDANCE_SOURCE_PRIORITY) {
    const scopedEvents = events.filter((event) => event.direction === direction && event.sourceGroup === sourceGroup);
    if (!scopedEvents.length) continue;

    if (direction === "entry") {
      const hasCandidate = scopedEvents.some((event) => event.time <= addMinutes(end, TOLERANCE_MINUTES));
      if (hasCandidate) return sourceGroup;
      continue;
    }

    if (direction === "exit") {
      const hasCandidate = scopedEvents.some((event) => event.time >= addMinutes(start, -TOLERANCE_MINUTES));
      if (hasCandidate) return sourceGroup;
    }
  }
  return "";
}

function pickEntryCandidate(events, exitEvents, start, end, priorityGroups = ATTENDANCE_SOURCE_PRIORITY) {
  for (const sourceGroup of priorityGroups) {
    const scopedEvents = events.filter((event) => event.sourceGroup === sourceGroup);
    if (!scopedEvents.length) continue;

    const onTime = scopedEvents
      .filter((event) => event.time <= addMinutes(start, TOLERANCE_MINUTES))
      .sort((a, b) => b.time - a.time)
      .find((event) => !hasExitBetween(exitEvents, event.time, start));
    if (onTime) return onTime;

    const lateEntry = scopedEvents
      .filter((event) => event.time > addMinutes(start, TOLERANCE_MINUTES) && event.time <= addMinutes(end, TOLERANCE_MINUTES))
      .sort((a, b) => a.time - b.time)[0] || null;
    if (lateEntry) return lateEntry;
  }

  return null;
}

function pickExitCandidate(events, siblingEntryEvents, end) {
  const nextEntryAfterEnd = siblingEntryEvents
    .filter((event) => event.time > addMinutes(end, TOLERANCE_MINUTES))
    .sort((a, b) => a.time - b.time)[0] || null;

  return events
    .filter((event) => (
      event.time >= addMinutes(end, -TOLERANCE_MINUTES)
      && (!nextEntryAfterEnd || event.time <= nextEntryAfterEnd.time)
    ))
    .sort((a, b) => a.time - b.time)[0] || null;
}

function pickEarlyExitCandidate(events, start, end) {
  return events
    .filter((event) => event.time >= addMinutes(start, -TOLERANCE_MINUTES) && event.time < addMinutes(end, -TOLERANCE_MINUTES))
    .sort((a, b) => b.time - a.time)[0] || null;
}

function pickExitEvidence(events, siblingEntryEvents, start, end, priorityGroups = ATTENDANCE_SOURCE_PRIORITY) {
  for (const sourceGroup of priorityGroups) {
    const scopedEvents = events.filter((event) => event.sourceGroup === sourceGroup);
    const scopedEntryEvents = siblingEntryEvents.filter((event) => event.sourceGroup === sourceGroup);
    if (!scopedEvents.length) continue;

    const matchedExit = pickExitCandidate(scopedEvents, scopedEntryEvents, end);
    if (matchedExit) return { outcome: "matched", event: matchedExit };

    const earlyExit = pickEarlyExitCandidate(scopedEvents, start, end);
    if (earlyExit) return { outcome: "early", event: earlyExit };
  }

  return { outcome: "missing", event: null };
}

function hasExitBetween(exitEvents, fromTime, toTime) {
  return exitEvents.some((event) => event.time > fromTime && event.time < addMinutes(toTime, TOLERANCE_MINUTES));
}

function auditTextToCategories(text) {
  const value = cleanText(text);
  if (!value || value === "正常" || value === "無") {
    return {
      attendanceStatuses: [],
      attendanceDetails: [],
      workRisks: [],
      workRiskDetails: [],
    };
  }

  const attendanceStatuses = [];
  const workRisks = [];

  if (/提早離院/.test(value)) attendanceStatuses.push(ATTENDANCE_STATUS.EARLY_LEAVE.short);
  if (/晚於加班/.test(value)) attendanceStatuses.push(ATTENDANCE_STATUS.LATE_OVERTIME.short);
  if (/簽到無紀錄|車進無紀錄/.test(value)) attendanceStatuses.push(ATTENDANCE_STATUS.MISSING_SIGN_IN.short);
  if (/簽退無紀錄|車出無符合紀錄|簽退未佐證加班/.test(value)) attendanceStatuses.push(ATTENDANCE_STATUS.MISSING_SIGN_OUT.short);
  if (value.includes("12小時") || (value.includes("12") && value.includes("小時"))) workRisks.push(WORK_RISK.OVER_12_HOURS);

  return {
    attendanceStatuses: unique(attendanceStatuses),
    attendanceDetails: unique(attendanceStatuses.map((status) => `${statusLongLabel(status)}：${value}`)),
    workRisks: unique(workRisks),
    workRiskDetails: unique(workRisks.map((risk) => `${risk}：${value}`)),
  };
}

function applyScheduleRisks(dailyMap) {
  dailyMap.forEach((daily) => {
    if ([DAY_TYPE.HOLIDAY, DAY_TYPE.FIXED_HOLIDAY].includes(daily.type) && daily.overtimeRecords.length > 0) {
      addWorkRisk(daily, WORK_RISK.HOLIDAY_WORK, `${daily.shift || daily.type} 有有效加班申請。`);
    }
  });

  const grouped = groupBy(state.scheduleRows, (row) => row.person);
  grouped.forEach((rows) => {
    rows.sort((a, b) => a.date - b.date);
    rows.forEach((row, index) => {
      const next = rows[index + 1];
      if (!next) return;
      const currentEnd = latestTime([row.signOut, ...row.otherSignOuts]);
      const nextStart = earliestTime([next.signIn, ...next.otherSignIns]);
      if (!currentEnd || !nextStart) return;

      const restHours = (nextStart - currentEnd) / 36e5;
      if (restHours >= 0 && restHours < 11) {
        const daily = dailyMap.get(dailyKey(row.person, row.dateKey));
        addWorkRisk(daily, WORK_RISK.REST_LESS_11, `次日上班前休息僅 ${formatNumber(restHours, 1)} 小時。`);
      }
    });
  });
}

function applyMonthlyOvertimeRisk(records, dailyMap) {
  const grouped = groupBy(records, (record) => `${record.person}|${record.month}`);
  grouped.forEach((items) => {
    items.sort((a, b) => a.date - b.date || (a.start?.getTime() || 0) - (b.start?.getTime() || 0));
    let cumulative = 0;
    items.forEach((record) => {
      cumulative += record.hours;
      record.monthlyCumulative = cumulative;
      if (cumulative > 46) {
        const cumulativeLabel = APPROVAL_MODE ? "申請累計" : "當月累計延展工時";
        record.workRisks.push(WORK_RISK.OVER_46_HOURS);
        record.workRiskDetails.push(`${WORK_RISK.OVER_46_HOURS}：${cumulativeLabel} ${formatNumber(cumulative, 1)} 小時。`);
        refreshRecordDerived(record);
        const daily = dailyMap.get(dailyKey(record.person, record.dateKey));
        addWorkRisk(daily, WORK_RISK.OVER_46_HOURS, `${cumulativeLabel} ${formatNumber(cumulative, 1)} 小時。`);
      }
    });
  });
}

function finalizeDailyRows(dailyMap) {
  dailyMap.forEach((daily) => {
    if (APPROVAL_MODE) approvalView.finalizeDaily(daily);
    const attendanceStatuses = [];
    const attendanceDetails = [];
    const attendanceMarkers = [];
    const attendanceCalendarLines = [];
    const workRisks = [];
    const workRiskDetails = [...daily.workRiskDetails];

    daily.overtimeRecords.forEach((record) => {
      attendanceStatuses.push(...record.attendanceStatuses);
      attendanceDetails.push(...record.attendanceDetails);
      attendanceMarkers.push(...(record.attendanceMarkers || []));
      attendanceCalendarLines.push(...buildRecordAttendanceCalendarLines(record));
      workRisks.push(...record.workRisks);
      workRiskDetails.push(...record.workRiskDetails);
    });

    const totalWorkHours = daily.scheduledHours + daily.overtimeHours;
    daily.totalWorkHours = totalWorkHours;
    if (totalWorkHours > 12) {
      addWorkRisk(daily, WORK_RISK.OVER_12_HOURS, `日排班加延展工時合計 ${formatNumber(totalWorkHours, 1)} 小時。`);
    }

    daily.attendanceStatuses = unique(attendanceStatuses);
    daily.attendanceDetails = unique(attendanceDetails);
    daily.attendanceMarkers = uniqueBy(attendanceMarkers, (marker) => `${marker.key}|${marker.note || ""}|${marker.detail || ""}`);
    daily.attendanceCalendarLines = uniqueBy(attendanceCalendarLines, (line) => `${line.text}|${line.tone}`);
    daily.workRisks = unique([...daily.workRisks, ...workRisks]);
    daily.workRiskDetails = unique([...daily.workRiskDetails, ...workRiskDetails]);
    daily.calendarNotes = buildCalendarNotes(daily);
    daily.hasNormalAttendance = daily.overtimeRecords.length > 0 && daily.attendanceStatuses.length === 0;
    daily.risks = unique([...daily.attendanceStatuses, ...daily.workRisks]);
    daily.riskDetails = unique([...daily.attendanceDetails, ...daily.workRiskDetails]);
  });
}

function buildCalendarNotes(daily) {
  const notes = [];
  if (daily.workRisks.includes(WORK_RISK.REST_LESS_11) && daily.overtimeRecords.length === 0) {
    const signIn = daily.signIn ? `簽到 ${formatCalendarTime(daily.signIn, daily.date, true)}` : "";
    const signOut = daily.signOut ? `簽退 ${formatCalendarTime(daily.signOut, daily.date, true)}` : "";
    if (signIn || signOut) {
      if (APPROVAL_MODE) notes.push(...[signIn, signOut].filter(Boolean));
      else notes.push([signIn, signOut].filter(Boolean).join(" / "));
    }
  }
  return notes;
}

function buildRecordAttendanceCalendarLines(record) {
  if (!record.attendanceSections?.length) return [];

  if (!record.attendanceStatuses.length) return [];
  return [{
    text: `簽到退：${record.attendanceStatuses.map(statusDisplayLabel).join("、")}`,
    tone: attendanceClass(record.attendanceStatuses),
  }];
}

function buildAttendanceSectionCalendarText(title, statuses, showNormalInCalendar = false) {
  if (statuses.length) return `${title}：${statuses.map(statusDisplayLabel).join("、")}`;
  if (showNormalInCalendar) return `${title}：${ATTENDANCE_STATUS.NORMAL.long}`;
  return "";
}

function addWorkRisk(target, risk, detail = "") {
  if (!target || !risk) return;
  if (!target.workRisks) target.workRisks = [];
  if (!target.workRiskDetails) target.workRiskDetails = [];
  target.workRisks.push(risk);
  if (detail) target.workRiskDetails.push(`${risk}：${detail}`);
}

function setupControls() {
  setSelectOptions(dom.monthSelect, state.months.map((month) => [month, month]), state.selectedMonth);
  setSelectOptions(dom.personSelect, [["all", "全部"], ...state.people.map((person) => [person, person])], "all");
  const displayRiskOptions = APPROVAL_MODE ? riskOptions.map(([value, label]) => [value, value === ATTENDANCE_STATUS.LATE_OVERTIME.short ? approvalView.lateSignInLabel : label]) : riskOptions;
  setSelectOptions(dom.riskSelect, APPROVAL_MODE ? [...displayRiskOptions, [approvalView.limitLabel, approvalView.limitLabel], ["unapproved", "有未核給時數"]] : riskOptions, "all");
  [dom.monthSelect, dom.personSelect, dom.riskSelect, dom.exportSummaryBtn, dom.exportDetailBtn].forEach((control) => {
    control.disabled = false;
  });
}

function setSelectOptions(select, options, selected) {
  select.innerHTML = "";
  options.forEach(([value, label]) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    if (value === selected) option.selected = true;
    select.appendChild(option);
  });
}

function render() {
  const filteredDaily = getFilteredDailyRows();
  const filteredRecords = getFilteredRecords();
  renderMetrics(filteredDaily, filteredRecords);
  renderCalendar(filteredDaily);
  renderDetail(filteredDaily);
  renderRiskTable(filteredDaily);
}

function getFilteredDailyRows() {
  return state.dailyRows.filter((daily) => {
    if (state.selectedMonth && daily.month !== state.selectedMonth) return false;
    if (state.selectedPerson !== "all" && daily.person !== state.selectedPerson) return false;
    if (APPROVAL_MODE && state.selectedRisk === "unapproved") return daily.unapprovedHours > 0;
    if (APPROVAL_MODE && state.selectedRisk === approvalView.limitLabel) return daily.overtimeRecords.some((record) => record.approvalAlerts.includes(approvalView.limitLabel));
    if (state.selectedRisk === "hasRisk" && !dailyHasAnyRisk(daily)) return false;
    if (!["all", "hasRisk"].includes(state.selectedRisk)) {
      const hasMarker = (daily.attendanceMarkers || []).some((marker) => marker.label === state.selectedRisk);
      if (!daily.risks.includes(state.selectedRisk) && !hasMarker) return false;
    }
    return true;
  });
}

function getFilteredRecords() {
  return state.overtimeRecords.filter((record) => {
    if (state.selectedMonth && record.month !== state.selectedMonth) return false;
    if (state.selectedPerson !== "all" && record.person !== state.selectedPerson) return false;
    if (APPROVAL_MODE && state.selectedRisk === "unapproved") return record.unapprovedHours > 0;
    if (APPROVAL_MODE && state.selectedRisk === approvalView.limitLabel) return record.approvalAlerts.includes(approvalView.limitLabel);
    if (state.selectedRisk === "hasRisk" && !recordHasAnyRisk(record)) return false;
    if (!["all", "hasRisk"].includes(state.selectedRisk)) {
      const hasMarker = (record.attendanceMarkers || []).some((marker) => marker.label === state.selectedRisk);
      if (!record.risks.includes(state.selectedRisk) && !hasMarker) return false;
    }
    return true;
  });
}

function renderMetrics(dailyRows, records) {
  if (APPROVAL_MODE) return approvalView.renderMetrics();
  const totalHours = sum(records.map((record) => record.hours));
  const totalPay = sum(records.map((record) => record.estimatedPay));
  const attendanceRiskDays = unique(
    dailyRows
      .filter((daily) => dailyHasAttendanceRisk(daily))
      .map((daily) => daily.dateKey),
  ).length;
  const workRiskDays = unique(
    dailyRows
      .filter((daily) => daily.workRisks.some((risk) => risk !== WORK_RISK.PENDING_REVIEW))
      .map((daily) => daily.dateKey),
  ).length;

  dom.metricHours.textContent = formatNumber(totalHours, 1);
  dom.metricPay.textContent = currency(totalPay);
  dom.metricAttendanceRiskDays.textContent = `${attendanceRiskDays}`;
  dom.metricWorkRiskDays.textContent = `${workRiskDays}`;
}

function recordHasAttendanceRisk(record) {
  return Boolean((record.attendanceStatuses || []).length)
    || (record.attendanceMarkers || []).some((marker) => marker.countsAsRisk);
}

function recordHasAnyRisk(record) {
  return Boolean((record.risks || []).length) || recordHasAttendanceRisk(record)
    || (APPROVAL_MODE && (record.unapprovedHours > 0 || record.approvalNotes.length > 0));
}

function dailyHasAttendanceRisk(daily) {
  return Boolean((daily.attendanceStatuses || []).length)
    || (daily.overtimeRecords || []).some((record) => recordHasAttendanceRisk(record));
}

function dailyHasAnyRisk(daily) {
  return Boolean((daily.risks || []).length) || dailyHasAttendanceRisk(daily)
    || (APPROVAL_MODE && daily.overtimeRecords.some(recordHasAnyRisk));
}

function renderCalendar(dailyRows) {
  dom.calendarGrid.innerHTML = "";
  if (!state.selectedMonth) {
    renderEmptyCalendar();
    return;
  }

  const monthDate = parseMonthToDate(state.selectedMonth);
  const start = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  const end = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  const gridStart = addDays(start, -start.getDay());
  const totalCells = Math.ceil((start.getDay() + end.getDate()) / 7) * 7;

  dom.calendarTitle.textContent = `${state.selectedMonth} 月曆`;
  dom.calendarSubtitle.textContent = `${state.selectedPerson === "all" ? "全部人員" : state.selectedPerson}，${riskLabel(state.selectedRisk)}`;

  const dailyByDate = groupBy(dailyRows, (daily) => daily.dateKey);
  for (let index = 0; index < totalCells; index += 1) {
    const current = addDays(gridStart, index);
    const dateKey = formatDateKey(current);
    const items = dailyByDate.get(dateKey) || [];
    const cell = document.createElement("button");
    cell.type = "button";
    cell.className = "day-cell";
    if (current.getMonth() !== monthDate.getMonth()) cell.classList.add("outside");
    if (state.selectedDateKey === dateKey) cell.classList.add("selected");
    if (items[0]?.type) cell.classList.add(dayTypeClass(items[0].type));
    cell.addEventListener("click", () => {
      state.selectedDateKey = dateKey;
      renderDetail(dailyRows);
      renderCalendar(dailyRows);
      if (APPROVAL_MODE && window.matchMedia("(max-width: 1100px)").matches) {
        document.getElementById("dateDetail")?.scrollIntoView({ behavior: "smooth", block: "start" });
      } else if (APPROVAL_MODE) {
        const panel = document.getElementById("dateDetail");
        if (panel) panel.scrollTop = 0;
      }
    });
    cell.innerHTML = dayCellHtml(current, items, current.getMonth() !== monthDate.getMonth());
    dom.calendarGrid.appendChild(cell);
  }
}

function dayCellHtml(date, items, isOutside = false) {
  const daily = items[0];
  const shift = isOutside ? "" : daily?.shift || "";
  const overtimeHours = sum(items.map((item) => item.overtimeHours));
  const totalPay = sum(items.map((item) => item.estimatedPay));
  const attendanceLines = uniqueBy(items.flatMap((item) => item.attendanceCalendarLines || []), (line) => `${line.text}|${line.tone}`);
  const anomalyGroups = buildDayRecordIssueGroups(items.flatMap((item) => item.overtimeRecords || []));
  const workRisks = unique(items.flatMap((item) => APPROVAL_MODE
    ? approvalView.calendarWorkRisks(item)
    : item.workRisks.filter((risk) => risk !== WORK_RISK.PENDING_REVIEW)));
  const notes = unique(items.flatMap((item) => item.calendarNotes));
  const rows = [];

  if (APPROVAL_MODE) rows.push(...approvalView.dayMetricRows(items));
  else {
    if (overtimeHours > 0) rows.push(dayLineHtml(`⏱️延展工時 ${formatNumber(overtimeHours, 1)} h`, "metric metric-hours"));
    if (totalPay > 0) rows.push(dayLineHtml(`💰加班費${currency(totalPay)}`, "metric metric-pay"));
  }
  if (anomalyGroups.length) {
    anomalyGroups.forEach((group) => rows.push(dayIssueHtml(group)));
  } else {
    attendanceLines.forEach((line) => rows.push(dayLineHtml(line.text, `attendance ${line.tone}`)));
  }
  if (workRisks.length) {
    rows.push(`<div class="day-pill-row">${workRisks.map((risk) => dayPillHtml(risk, riskPillClass(risk))).join("")}</div>`);
  }
  notes.forEach((note) => rows.push(dayLineHtml(note, "note")));

  return `
    <div class="day-top">
      <div class="day-meta">
        <span class="day-number">${isOutside ? "&nbsp;" : `${date.getDate()}日`}</span>
        ${APPROVAL_MODE && !isOutside ? `<span class="day-weekday">（${["日", "一", "二", "三", "四", "五", "六"][date.getDay()]}）</span>` : ""}
        ${shift ? `<span class="day-shift-inline">${escapeHtml(shift)}</span>` : ""}
      </div>
    </div>
    <div class="day-content">
      ${rows.join("")}
    </div>
  `;
}

function dayLineHtml(text, kind = "") {
  return `<div class="day-line ${kind}">${escapeHtml(text)}</div>`;
}

function dayIssueHtml(line) {
  const range = line.range ? `<span class="day-issue-range">${escapeHtml(line.range)}</span>` : "";
  const pills = (line.items || []).map((item) => dayPillHtml(item.label, attendancePillClass(item.status))).join("");
  const notes = (line.notes || []).map((note) => `<div class="day-issue-note">${escapeHtml(note)}</div>`).join("");
  return `
    <div class="day-issue-block">
      <div class="day-issue-line">${range}<span class="day-pill-row">${pills}</span></div>
      ${notes ? `<div class="day-issue-notes">${notes}</div>` : ""}
    </div>
  `;
}

function dayPillHtml(text, kind = "") {
  return `<span class="day-pill ${kind}">${escapeHtml(text)}</span>`;
}

function riskPillClass(risk) {
  return risk === WORK_RISK.HOLIDAY_WORK ? "work-risk holiday-work" : "work-risk";
}

function attendanceStatusClass(status) {
  if (status === ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.key) return "return";
  if (status === ATTENDANCE_MARKER.PROXY_CARD.key) return "proxy";
  if ([ATTENDANCE_STATUS.MISSING_SIGN_IN.short, ATTENDANCE_STATUS.MISSING_SIGN_OUT.short].includes(status)) return "missing";
  if (status === ATTENDANCE_STATUS.EARLY_LEAVE.short) return "early";
  if (status === ATTENDANCE_STATUS.LATE_OVERTIME.short) return "late";
  return "ok";
}

function attendancePillClass(status) {
  return `attendance ${attendanceStatusClass(status)}`;
}

function renderDetail(dailyRows) {
  const selectedKey = state.selectedDateKey || dailyRows[0]?.dateKey || "";
  state.selectedDateKey = selectedKey;
  const items = dailyRows.filter((daily) => daily.dateKey === selectedKey);

  if (!selectedKey || items.length === 0) {
    dom.detailTitle.textContent = "明細";
    dom.detailSubtitle.textContent = "選取日期查看稽核內容";
    dom.detailBody.className = "detail-body empty-state";
    dom.detailBody.textContent = "目前篩選沒有日期資料。";
    return;
  }

  const date = parseDateValue(selectedKey);
  const totalHours = sum(items.map((item) => item.overtimeHours));
  const totalPay = sum(items.map((item) => item.estimatedPay));
  const records = items.flatMap((item) => item.overtimeRecords);
  const attendanceStatuses = unique(items.flatMap((item) => item.attendanceStatuses));
  const workRisks = unique(items.flatMap((item) => APPROVAL_MODE
    ? approvalView.calendarWorkRisks(item)
    : item.workRisks.filter((risk) => risk !== WORK_RISK.PENDING_REVIEW)));
  const attendanceEvents = uniqueBy(
    items.flatMap((item) => item.attendanceEvents),
    (event) => `${event.person}|${event.direction}|${event.sourceLabel}|${event.time.getTime()}`,
  ).sort((a, b) => a.time - b.time);
  const attendanceSummaryItems = records.flatMap((record) => buildAttendanceSummaryItems(record));
  const attendanceBadges = attendanceSummaryItems.length
    ? uniqueBy(attendanceSummaryItems, (item) => item.label).map((item) => statusBadgeHtml(item.label, attendancePillClass(item.status))).join("")
    : records.length
      ? statusBadgeHtml(ATTENDANCE_STATUS.NORMAL.long, "attendance ok")
      : '<span class="badge">無加班申請</span>';
  const workBadges = workRisks.length
    ? workRisks.map((risk) => statusBadgeHtml(risk, risk === WORK_RISK.HOLIDAY_WORK ? "work holiday-work" : "work")).join("")
    : '<span class="badge">無</span>';

  dom.detailTitle.textContent = formatDisplayDate(date);
  dom.detailSubtitle.textContent = `${items.length} 筆日資料，${records.length} 筆加班單`;
  dom.detailBody.className = "detail-body";
  dom.detailBody.innerHTML = `
    <div class="detail-stack">
      ${APPROVAL_MODE ? approvalView.detailTotals(items) : `<div class="detail-row"><span>加班時數</span><strong>${formatNumber(totalHours, 1)} 小時</strong></div><div class="detail-row"><span>加班費</span><strong>${currency(totalPay)}</strong></div>`}
      <div class="detail-row"><span>當日紀錄</span><div class="timeline-text">${attendanceEvents.length ? attendanceEvents.map(formatAttendanceEvent).join("、") : "無"}</div></div>
      <section class="detail-section">
        <h3>出勤狀況</h3>
        <div class="risk-list">${attendanceBadges}</div>
        <div class="detail-notes">${renderNoteList(attendanceSummaryItems.length ? attendanceSummaryItems.map((item) => item.text) : ["正常"])}</div>
      </section>
      <section class="detail-section">
        <h3>延展工時異常</h3>
        <div class="risk-list">${workBadges}</div>
        <div class="detail-notes">${renderNoteList(unique(items.flatMap((item) => APPROVAL_MODE ? approvalView.visibleWorkRiskDetails(item) : item.workRiskDetails)))}</div>
      </section>
      <div class="record-list">${records.length ? records.map(recordHtml).join("") : '<div class="empty-state">目前沒有加班單明細。</div>'}</div>
    </div>
  `;
}

function recordHtml(record) {
  if (APPROVAL_MODE) return approvalView.recordHtml(record);
  const segments = formatRecordSegmentSummary(record);
  const attendanceSummaryItems = buildAttendanceSummaryItems(record);

  const attendanceLabel = attendanceSummaryItems.length
    ? attendanceSummaryItems.map((item) => item.label).join("、")
    : ATTENDANCE_STATUS.NORMAL.long;
  const workLabel = record.workRisks.filter((risk) => risk !== WORK_RISK.PENDING_REVIEW).join("、") || "無";
  const evidenceItems = buildRecordEvidenceItems(record);
  const issueItems = buildRecordIssueItems(record);
  const rangeText = formatRange(record.start, record.end, record.startRaw, record.endRaw);

  return `
    <article class="record-item">
      <div class="record-header">
        <strong>${escapeHtml(rangeText)}</strong>
        <span class="record-hours">${formatNumber(record.hours, 1)}h</span>
      </div>
      <p class="record-reason">${escapeHtml(record.reason || record.source)}</p>
      <ul class="record-bullets">
        <li>加班費 ${currency(record.estimatedPay)}${record.payNote ? `｜${escapeHtml(record.payNote)}` : ""}</li>
        <li>倍率 ${escapeHtml(segments || "無倍率拆段")}</li>
        <li>出勤狀況 ${escapeHtml(attendanceLabel)}</li>
        <li>延展工時異常 ${escapeHtml(workLabel)}</li>
      </ul>
      <div class="record-mini-section">
        <h4>出勤佐證</h4>
        ${renderCompactList(evidenceItems)}
      </div>
      <div class="record-mini-section">
        <h4>異常說明</h4>
        ${renderCompactList(issueItems.length ? issueItems : ["無"])}
      </div>
    </article>
  `;
}

function formatRecordSegmentSummary(record) {
  const displaySegments = buildDisplaySegments(record);
  if (!displaySegments.length) return "無倍率拆段";
  return displaySegments.map((segment) => {
    const multiplierText = formatDisplayMultiplier(segment.multiplier);
    const taxedText = segment.taxed ? "(含稅)" : "";
    return `${multiplierText}倍${taxedText} ${formatNumber(segment.hours, 1)}H`;
  }).join(" + ");
}

function buildDisplaySegments(record) {
  const hours = Math.max(0, Number(record.hours) || 0);
  if (!hours) return [];

  if (record.type === DAY_TYPE.WORKDAY) {
    const firstBlock = Math.min(hours, 2);
    const secondBlock = Math.min(Math.max(hours - 2, 0), 2);
    const thirdBlock = Math.max(hours - 4, 0);
    return [
      firstBlock > 0 ? { multiplier: 1.34, hours: firstBlock, taxed: false } : null,
      secondBlock > 0 ? { multiplier: 1.67, hours: secondBlock, taxed: false } : null,
      thirdBlock > 0 ? { multiplier: 1.67, hours: thirdBlock, taxed: true } : null,
    ].filter(Boolean);
  }

  return (record.segments || [])
    .filter((segment) => segment.hours > 0 || segment.paidHours > 0)
    .map((segment) => ({
      multiplier: segment.multiplier,
      hours: segment.paidHours && segment.paidHours !== segment.hours ? segment.paidHours : segment.hours,
      taxed: /13H\+/.test(segment.label || ""),
    }));
}

function formatDisplayMultiplier(multiplier) {
  const rounded = Number(multiplier) || 0;
  if (rounded === 1.34) return "1.334";
  if (rounded === 1.67) return "1.667";
  if (rounded === 2.67) return "2.667";
  return formatNumber(rounded, 3);
}

function renderRiskTable(dailyRows) {
  if (APPROVAL_MODE) return approvalView.renderRiskTable(dailyRows);
  const riskRows = dailyRows.filter((daily) => dailyHasAnyRisk(daily));
  dom.riskSubtitle.textContent = `${riskRows.length} 筆日期需覆核`;
  if (!riskRows.length) {
    dom.riskTableBody.innerHTML = '<tr><td colspan="6" class="table-empty">目前篩選無異常資料</td></tr>';
    return;
  }

  dom.riskTableBody.innerHTML = riskRows.map((daily) => `
    <tr>
      <td>${escapeHtml(formatDisplayDate(daily.date))}</td>
      <td>${escapeHtml(daily.person)}</td>
      <td>${escapeHtml(daily.shift || daily.type)}</td>
      <td>${formatNumber(daily.overtimeHours, 1)} 小時</td>
      <td>${currency(daily.estimatedPay)}</td>
      <td>${renderRiskSummaryCell(daily)}</td>
    </tr>
  `).join("");
}

function renderRiskSummaryCell(daily) {
  const attendanceBlocks = buildRiskSummaryAttendanceBlocks(daily.overtimeRecords || []);
  const workBlocks = buildRiskSummaryWorkBlocks(daily);

  return `
    <div class="risk-summary">
      ${attendanceBlocks.map((block) => `
        <div class="risk-summary-block">
          <div class="risk-block-title">●加班起訖 ${escapeHtml(block.range)}</div>
          ${block.items.map((item) => `
            <div class="risk-summary-line">
              <span class="risk-note">${escapeHtml(item.text)}</span>
              ${dayPillHtml(item.label, attendancePillClass(item.status))}
            </div>
          `).join("")}
        </div>
      `).join("")}
      ${workBlocks.map((item) => `
        <div class="risk-summary-block">
          <div class="risk-summary-line">
            ${dayPillHtml(item.label, riskPillClass(item.label))}
            ${item.note ? `<span class="risk-note">${escapeHtml(item.note)}</span>` : ""}
          </div>
        </div>
      `).join("")}
    </div>
  `;
}

function renderEmptyCalendar() {
  dom.calendarGrid.innerHTML = "";
  for (let index = 0; index < 35; index += 1) {
    const cell = document.createElement("div");
    cell.className = "day-cell outside";
    cell.innerHTML = '<div class="day-top"><span class="day-number">&nbsp;</span></div>';
    dom.calendarGrid.appendChild(cell);
  }
}

function exportSummaryCsv() {
  if (APPROVAL_MODE) return approvalView.exportSummaryCsv();
  const rows = getFilteredDailyRows();
  const grouped = groupBy(rows, (row) => `${row.month}|${row.person}`);
  const data = [["月份", "人員", "延展工時", "額外加班費", "異常天數", "異常類型"]];

  grouped.forEach((items, key) => {
    const [month, person] = key.split("|");
    data.push([
      month,
      person,
      formatNumber(sum(items.map((item) => item.overtimeHours)), 2),
      Math.round(sum(items.map((item) => item.estimatedPay))),
      items.filter((item) => item.risks.length > 0).length,
      unique(items.flatMap((item) => item.risks)).map(statusDisplayLabel).join("、"),
    ]);
  });

  downloadCsv(data, `異常工時摘要_${state.selectedMonth || "全部"}.csv`);
}

function exportDetailCsv() {
  if (APPROVAL_MODE) return approvalView.exportDetailCsv();
  const data = [[
    "日期",
    "月份",
    "人員",
    "班別",
    "類別",
    "延展工時",
    "時薪",
    "額外加班費",
    "起始時間",
    "終止時間",
    "出勤狀況",
    "延展工時異常",
    "稽核說明",
    "事由",
  ]];

  getFilteredRecords().forEach((record) => {
    data.push([
      record.dateKey,
      record.month,
      record.person,
      record.shift,
      record.type,
      formatNumber(record.hours, 2),
      record.hourly || "",
      Math.round(record.estimatedPay),
      formatDateTime(record.start) || record.startRaw,
      formatDateTime(record.end) || record.endRaw,
      record.attendanceStatuses.length ? record.attendanceStatuses.map(statusLongLabel).join("、") : ATTENDANCE_STATUS.NORMAL.long,
      record.workRisks.filter((risk) => risk !== WORK_RISK.PENDING_REVIEW).join("、"),
      record.riskDetails.join("；"),
      record.reason,
    ]);
  });

  downloadCsv(data, `異常工時明細_${state.selectedMonth || "全部"}.csv`);
}

function downloadCsv(rows, filename) {
  const csv = rows.map((row) => row.map(csvCell).join(",")).join("\r\n");
  const blob = new Blob([`\ufeff${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function csvCell(value) {
  const text = value == null ? "" : String(value);
  return `"${text.replace(/"/g, '""')}"`;
}

function setStatus(message, warning) {
  dom.fileStatus.textContent = message;
  if (APPROVAL_MODE) dom.fileStatus.hidden = !message;
  dom.fileStatus.classList.toggle("warning", Boolean(warning));
}

function calculateOvertimePay(hours, hourly, type) {
  const safeHours = Math.max(0, Number(hours) || 0);
  const safeHourly = Math.max(0, Number(hourly) || 0);
  const segments = [];
  const note = [];
  let remaining = safeHours;

  const addSegment = (label, actualHours, multiplier, options = {}) => {
    if (actualHours <= 0) return;
    const paidHours = options.paidHours ?? actualHours;
    segments.push({
      label,
      hours: actualHours,
      paidHours,
      multiplier,
      amount: Math.ceil(paidHours * safeHourly * multiplier),
    });
  };

  if (type === DAY_TYPE.FIXED_HOLIDAY) {
    if (safeHours > 0) {
      const firstBlock = Math.min(safeHours, 8);
      addSegment("固定假日基本加給", firstBlock, 1, { paidHours: 8 });
      note.push("未滿 8 小時仍以 8 小時計額外加給。");
      remaining = Math.max(0, safeHours - 8);
    }
    const secondBlock = Math.min(remaining, 2);
    addSegment("固定假日 9-10H", secondBlock, 1.34);
    remaining -= secondBlock;
    const thirdBlock = Math.min(remaining, 2);
    addSegment("固定假日 11-12H", thirdBlock, 1.67);
    remaining -= thirdBlock;
    addSegment("固定假日 13H+", remaining, 1.67);
  } else if (type === DAY_TYPE.HOLIDAY) {
    const firstBlock = Math.min(remaining, 8);
    addSegment("例假日基本加給", firstBlock, 1);
    if (firstBlock > 0) note.push("另有一日補休。");
    remaining -= firstBlock;
    const secondBlock = Math.min(remaining, 2);
    addSegment("例假日 9-10H", secondBlock, 1.34);
    remaining -= secondBlock;
    const thirdBlock = Math.min(remaining, 2);
    addSegment("例假日 11-12H", thirdBlock, 1.67);
    remaining -= thirdBlock;
    addSegment("例假日 13H+", remaining, 1.67);
  } else if (type === DAY_TYPE.SPECIAL) {
    const firstBlock = Math.min(remaining, 8);
    addSegment("天災/選舉日 1-8H", firstBlock, 1);
    remaining -= firstBlock;
    const secondBlock = Math.min(remaining, 2);
    addSegment("天災/選舉日 9-10H", secondBlock, 1.34);
    remaining -= secondBlock;
    const thirdBlock = Math.min(remaining, 2);
    addSegment("天災/選舉日 11-12H", thirdBlock, 1.67);
    remaining -= thirdBlock;
    addSegment("天災/選舉日 13H+", remaining, 1.67);
  } else if (type === DAY_TYPE.RESTDAY) {
    const firstBlock = Math.min(remaining, 2);
    addSegment("休息日 1-2H", firstBlock, 1.34);
    remaining -= firstBlock;
    const secondBlock = Math.min(remaining, 6);
    addSegment("休息日 3-8H", secondBlock, 1.67);
    remaining -= secondBlock;
    addSegment("休息日 9H+", remaining, 2.67);
  } else {
    const firstBlock = Math.min(remaining, 2);
    addSegment("平日 9-10H", firstBlock, 1.34);
    remaining -= firstBlock;
    addSegment("平日 11H+", remaining, 1.67);
  }

  return {
    total: sum(segments.map((segment) => segment.amount)),
    segments,
    note: note.join(" "),
  };
}

function normalizeOvertimeType(type) {
  const value = cleanText(type);
  if (!value) return DAY_TYPE.WORKDAY;
  if ([DAY_TYPE.WORKDAY, DAY_TYPE.RESTDAY, DAY_TYPE.HOLIDAY, DAY_TYPE.FIXED_HOLIDAY, DAY_TYPE.SPECIAL].includes(value)) return value;
  return classifyDay(value, value);
}

function classifyDay(shift, nationalHoliday) {
  const value = `${shift || ""} ${nationalHoliday || ""}`;
  if (/天災|颱風|選舉/.test(value)) return DAY_TYPE.SPECIAL;
  if (/國定|固定假/.test(value)) return DAY_TYPE.FIXED_HOLIDAY;
  if (/例假/.test(value)) return DAY_TYPE.HOLIDAY;
  if (/休息/.test(value)) return DAY_TYPE.RESTDAY;
  return DAY_TYPE.WORKDAY;
}

function dayTypeClass(type) {
  switch (type) {
    case DAY_TYPE.RESTDAY:
      return "type-restday";
    case DAY_TYPE.HOLIDAY:
      return "type-holiday";
    case DAY_TYPE.FIXED_HOLIDAY:
      return "type-fixed-holiday";
    case DAY_TYPE.SPECIAL:
      return "type-special";
    default:
      return "type-workday";
  }
}

function attendanceClass(statuses) {
  if (statuses.some((status) => [ATTENDANCE_STATUS.MISSING_SIGN_IN.short, ATTENDANCE_STATUS.MISSING_SIGN_OUT.short].includes(status))) {
    return "missing";
  }
  if (statuses.includes(ATTENDANCE_STATUS.EARLY_LEAVE.short)) return "early";
  if (statuses.includes(ATTENDANCE_STATUS.LATE_OVERTIME.short)) return "late";
  return "ok";
}

function statusLongLabel(shortLabel) {
  const matched = Object.values(ATTENDANCE_STATUS).find((item) => item.short === shortLabel);
  return matched?.long || shortLabel;
}

function statusDisplayLabel(value) {
  if (APPROVAL_MODE && value === ATTENDANCE_STATUS.LATE_OVERTIME.short) return approvalView.lateSignInLabel;
  return statusLongLabel(value);
}

function statusBadgeHtml(text, extraClass = "") {
  return `<span class="badge ${extraClass}">${escapeHtml(text)}</span>`;
}

function renderNoteList(items) {
  if (!items.length) return '<div class="empty-note">無</div>';
  return `<ul class="note-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
}

function renderCompactList(items) {
  if (!items.length) return '<div class="empty-note">無</div>';
  return `<ul class="record-list-compact">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
}

function buildDayRecordIssueLines(records) {
  return buildDayRecordIssueGroups(records).flatMap((group) => (
    group.items.map((item) => ({
      range: group.range,
      label: item.label,
      status: item.status,
      tone: item.tone,
    }))
  ));
}

function buildDayRecordIssueGroups(records) {
  return dedupeDuplicateProxyMarkers(
    records
    .slice()
    .sort((left, right) => (left.start?.getTime() || 0) - (right.start?.getTime() || 0))
    .map((record) => {
      const items = [
        ...buildAttendanceMarkerItems(record),
        ...buildAttendanceDisplayItems(record.attendanceStatuses || []),
      ];
      if (!items.length) return null;
      return {
        range: shortRangeLabel(record.start, record.end, record.startRaw, record.endRaw),
        items,
        notes: unique(items.flatMap((item) => item.calendarNotes || [item.note]).filter(Boolean)),
      };
    })
    .filter(Boolean),
  );
}

function buildRiskSummaryAttendanceBlocks(records) {
  return dedupeDuplicateProxyMarkers(
    records
    .slice()
    .sort((left, right) => (left.start?.getTime() || 0) - (right.start?.getTime() || 0))
    .map((record) => {
      const items = buildAttendanceSummaryItems(record);
      if (!items.length) return null;
      return {
        range: shortRangeLabel(record.start, record.end, record.startRaw, record.endRaw),
        items,
      };
    })
    .filter(Boolean),
  );
}

function dedupeDuplicateProxyMarkers(groups) {
  const seenProxyMarkers = new Set();
  return groups
    .map((group) => {
      const items = (group.items || []).filter((item) => {
        if (item.status !== ATTENDANCE_MARKER.PROXY_CARD.key) return true;
        const dedupeKey = `${item.status}|${item.note || ""}|${item.text || ""}`;
        if (seenProxyMarkers.has(dedupeKey)) return false;
        seenProxyMarkers.add(dedupeKey);
        return true;
      });
      if (!items.length) return null;
      return {
        ...group,
        items,
        notes: unique(items.flatMap((item) => item.calendarNotes || [item.note]).filter(Boolean)),
      };
    })
    .filter(Boolean);
}

function buildAttendanceDisplayItems(statuses) {
  const uniqueStatuses = unique(statuses || []);
  const hasMissingIn = uniqueStatuses.includes(ATTENDANCE_STATUS.MISSING_SIGN_IN.short);
  const hasMissingOut = uniqueStatuses.includes(ATTENDANCE_STATUS.MISSING_SIGN_OUT.short);
  const items = [];

  if (hasMissingIn && hasMissingOut) {
    items.push({
      key: "missing-both",
      label: "簽到、簽退無紀錄",
      status: ATTENDANCE_STATUS.MISSING_SIGN_IN.short,
      tone: "missing",
      merged: true,
    });
  } else {
    if (hasMissingIn) {
      items.push({
        key: ATTENDANCE_STATUS.MISSING_SIGN_IN.short,
        label: statusDisplayLabel(ATTENDANCE_STATUS.MISSING_SIGN_IN.short),
        status: ATTENDANCE_STATUS.MISSING_SIGN_IN.short,
        tone: "missing",
      });
    }
    if (hasMissingOut) {
      items.push({
        key: ATTENDANCE_STATUS.MISSING_SIGN_OUT.short,
        label: statusDisplayLabel(ATTENDANCE_STATUS.MISSING_SIGN_OUT.short),
        status: ATTENDANCE_STATUS.MISSING_SIGN_OUT.short,
        tone: "missing",
      });
    }
  }

  if (uniqueStatuses.includes(ATTENDANCE_STATUS.LATE_OVERTIME.short)) {
    items.push({
      key: ATTENDANCE_STATUS.LATE_OVERTIME.short,
      label: statusDisplayLabel(ATTENDANCE_STATUS.LATE_OVERTIME.short),
      status: ATTENDANCE_STATUS.LATE_OVERTIME.short,
      tone: "late",
    });
  }

  if (uniqueStatuses.includes(ATTENDANCE_STATUS.EARLY_LEAVE.short)) {
    items.push({
      key: ATTENDANCE_STATUS.EARLY_LEAVE.short,
      label: statusDisplayLabel(ATTENDANCE_STATUS.EARLY_LEAVE.short),
      status: ATTENDANCE_STATUS.EARLY_LEAVE.short,
      tone: "early",
    });
  }

  return items;
}

function buildAttendanceMarkerItems(record) {
  return (record.attendanceMarkers || []).map((marker) => {
    const tone = marker.key === ATTENDANCE_MARKER.PROXY_CARD.key ? "proxy" : "return";
    return {
      key: marker.key,
      label: marker.label,
      status: marker.key,
      tone,
      marker: true,
      note: marker.note || "",
      calendarNotes: APPROVAL_MODE && marker.key === ATTENDANCE_MARKER.PROXY_CARD.key
        ? [`簽到 ${formatCalendarTime(marker.clockEntry, record.date)}`, `汽車進場 ${formatCalendarTime(marker.carEntry, record.date)}`]
        : APPROVAL_MODE && marker.key === ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.key
          ? [`原簽退 ${formatCalendarTime(marker.originalSignOut, record.date)}`, `返回簽到 ${formatCalendarTime(marker.returnSignIn, record.date)}`]
          : undefined,
      detail: marker.detail,
    };
  });
}

function buildAttendanceSummaryItems(record) {
  const section = record.attendanceSections?.[0];
  return [
    ...buildAttendanceMarkerItems(record).map((item) => ({
      ...item,
      text: buildAttendanceSummaryText(record, section, item),
    })),
    ...buildAttendanceDisplayItems(record.attendanceStatuses || []).map((item) => ({
      ...item,
      text: buildAttendanceSummaryText(record, section, item),
    })),
  ];
}

function buildAttendanceSummaryText(record, section, item) {
  if (item.marker) return item.detail || item.label;

  if (item.merged) {
    return `簽到、簽退無紀錄：${formatDateTime(record.start)} 起至 ${formatDateTime(record.end)} 前後未找到可佐證開始或結束加班的紀錄`;
  }

  if (item.status === ATTENDANCE_STATUS.LATE_OVERTIME.short) {
    return `${APPROVAL_MODE ? "簽到晚" : ATTENDANCE_STATUS.LATE_OVERTIME.long}：實際簽到採認紀錄為 ${formatDateTime(section?.entry?.time)}${section?.entry?.sourceLabel ? `（${section.entry.sourceLabel}）` : ""}`;
  }

  if (item.status === ATTENDANCE_STATUS.EARLY_LEAVE.short) {
    return `${ATTENDANCE_STATUS.EARLY_LEAVE.long}：實際簽退採認紀錄為 ${formatDateTime(section?.exit?.time)}${section?.exit?.sourceLabel ? `（${section.exit.sourceLabel}）` : ""}`;
  }

  if (item.status === ATTENDANCE_STATUS.MISSING_SIGN_IN.short) {
    return `${ATTENDANCE_STATUS.MISSING_SIGN_IN.long}：${formatDateTime(record.start)} 起未找到可佐證開始加班的紀錄`;
  }

  if (item.status === ATTENDANCE_STATUS.MISSING_SIGN_OUT.short) {
    return `${ATTENDANCE_STATUS.MISSING_SIGN_OUT.long}：${formatDateTime(record.end)} 前後未找到可佐證結束加班的紀錄`;
  }

  return item.label;
}

function buildRiskSummaryWorkBlocks(daily) {
  const noteMap = new Map();
  (daily.calendarNotes || []).forEach((note) => {
    if (daily.workRisks.includes(WORK_RISK.REST_LESS_11)) {
      noteMap.set(WORK_RISK.REST_LESS_11, note);
    }
  });

  return unique(daily.workRisks.filter((risk) => risk !== WORK_RISK.PENDING_REVIEW)).map((risk) => ({
    label: risk,
    note: noteMap.get(risk) || "",
  }));
}

function buildRecordEvidenceItems(record) {
  const items = [];
  (record.attendanceMarkers || []).forEach((marker) => {
    if (marker.key === ATTENDANCE_MARKER.RETURN_AFTER_SIGNOUT.key) {
      items.push(`${marker.label}：原簽退 ${formatDateTime(marker.originalSignOut)}，新簽到 ${formatDateTime(marker.returnSignIn)}`);
      return;
    }
    if (marker.key === ATTENDANCE_MARKER.PROXY_CARD.key) {
      items.push(`${marker.label}：簽到時間 ${formatDateTime(marker.clockEntry)}，汽車進場卻是 ${formatDateTime(marker.carEntry)}`);
    }
  });
  (record.attendanceSections || []).forEach((section) => {
    if (section.entry) items.push(`${section.title}起始 ${formatDateTime(section.entry.time)}（${section.entry.sourceLabel}）`);
    if (section.exit) items.push(`${section.title}結束 ${formatDateTime(section.exit.time)}（${section.exit.sourceLabel}）`);
    if (!section.entry && !section.exit && section.calendarText?.includes("無汽車進出紀錄")) {
      items.push(`${section.title}：無汽車進出紀錄`);
    }
  });
  if (!items.length && record.normalAttendanceText) items.push(record.normalAttendanceText);
  return items.length ? items : ["無額外佐證"];
}

function buildRecordIssueItems(record) {
  const details = buildAttendanceSummaryItems(record).map((item) => `${item.text} (${item.label})`);
  const workDetails = (record.workRiskDetails || []).filter(Boolean);
  return [...details, ...workDetails];
}

function shortRangeLabel(start, end, startRaw = "", endRaw = "") {
  if (start && end) return `${formatCalendarTime(start, start, false)}-${formatCalendarTime(end, start, false)}`;
  if (startRaw || endRaw) return `${cleanText(startRaw)}-${cleanText(endRaw)}`.replace(/^-|-$/g, "");
  return "本筆";
}

function getField(row, names) {
  const entries = Object.entries(row);
  for (const name of names) {
    if (!name) continue;
    if (Object.prototype.hasOwnProperty.call(row, name) && row[name] !== "") return row[name];
    const normalized = normalizeHeader(name);
    const found = entries.find(([key, value]) => normalizeHeader(key) === normalized && value !== "");
    if (found) return found[1];
  }
  return "";
}

function normalizeHeader(value) {
  return cleanText(value).replace(/\s+/g, "").replace(/[｜|_（）()]/g, "");
}

function cleanText(value) {
  if (value == null) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = cleanText(value).replace(/,/g, "");
  if (!text || text === "--") return 0;
  const match = text.match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const text = cleanText(value);
  if (!text) return null;

  const full = text.match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (full) return new Date(Number(full[1]), Number(full[2]) - 1, Number(full[3]));

  const serial = Number(text);
  if (Number.isFinite(serial) && serial > 20000) {
    const parsed = XLSX.SSF.parse_date_code(serial);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  return null;
}

function parseMonthValue(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return formatMonth(value);
  const text = cleanText(value);
  const match = text.match(/(\d{4})[/-](\d{1,2})/);
  if (match) return `${match[1]}/${String(Number(match[2])).padStart(2, "0")}`;
  const date = parseDateValue(value);
  return date ? formatMonth(date) : "";
}

function parseMonthToDate(month) {
  const match = cleanText(month).match(/(\d{4})[/-](\d{1,2})/);
  return match ? new Date(Number(match[1]), Number(match[2]) - 1, 1) : new Date();
}

function parseSeriesDateTimes(baseDate, value) {
  const text = cleanText(value);
  if (!baseDate || !text) return [];
  if (typeof value === "number" || value instanceof Date) {
    const single = combineDateTime(baseDate, value);
    return single ? [single] : [];
  }

  const segments = text
    .split(/[|、；;\n\r]+/)
    .map((item) => cleanText(item))
    .filter(Boolean);

  return segments
    .map((item) => combineDateTime(baseDate, item))
    .filter(Boolean);
}

function combineDateTime(baseDate, value) {
  if (!baseDate || value == null) return null;
  const text = cleanText(value);
  if (!text || text === "--") return null;

  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    if (value.getFullYear() > 1905) return new Date(value);
    return new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), value.getHours(), value.getMinutes(), value.getSeconds());
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    if (value > 1) {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, Math.floor(parsed.S || 0));
    }
    const totalSeconds = Math.round(value * 86400);
    return addSeconds(new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate()), totalSeconds);
  }

  const fullDateTime = parseDateTimeText(text, baseDate);
  if (fullDateTime) return fullDateTime;
  const time = parseTimeText(text);
  if (!time) return null;
  return new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), time.hour, time.minute, time.second);
}

function parseDateTimeText(text, baseDate) {
  const ymd = text.match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2}).*?(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if (ymd) {
    const hour = normalizeHour(Number(ymd[5]), ymd[4]);
    return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]), hour, Number(ymd[6]), Number(ymd[7] || 0));
  }

  const md = text.match(/(\d{1,2})[/-](\d{1,2}).*?(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if (md) {
    const hour = normalizeHour(Number(md[4]), md[3]);
    let result = new Date(baseDate.getFullYear(), Number(md[1]) - 1, Number(md[2]), hour, Number(md[5]), Number(md[6] || 0));
    if (result < addDays(baseDate, -10)) result = new Date(baseDate.getFullYear() + 1, Number(md[1]) - 1, Number(md[2]), hour, Number(md[5]), Number(md[6] || 0));
    return result;
  }
  return null;
}

function parseTimeText(text) {
  const match = text.match(/(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if (!match) return null;
  return {
    hour: normalizeHour(Number(match[2]), match[1]),
    minute: Number(match[3]),
    second: Number(match[4] || 0),
  };
}

function normalizeHour(hour, meridiem) {
  const marker = cleanText(meridiem).toUpperCase();
  if ((marker === "下午" || marker === "PM") && hour < 12) return hour + 12;
  if ((marker === "上午" || marker === "AM") && hour === 12) return 0;
  return hour;
}

function normalizeEndTime(start, end) {
  if (!start || !end) return end || null;
  let normalized = new Date(end);
  while (normalized < addHours(start, -1)) {
    normalized = addDays(normalized, 1);
  }
  return normalized;
}

function formatDateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatMonth(date) {
  return `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function formatDisplayDate(date) {
  if (!date) return "";
  return `${formatDateKey(date)}（${"日一二三四五六"[date.getDay()]}）`;
}

function formatDateTime(date) {
  if (!date) return "";
  return `${formatDateKey(date)} ${formatTimeWithSeconds(date)}`;
}

function formatTimeWithSeconds(date) {
  if (!date) return "";
  return [
    String(date.getHours()).padStart(2, "0"),
    String(date.getMinutes()).padStart(2, "0"),
    String(date.getSeconds()).padStart(2, "0"),
  ].join(":");
}

function formatCalendarTime(date, baseDate, includeSeconds = true) {
  if (!date) return "";
  const timeText = includeSeconds
    ? formatTimeWithSeconds(date)
    : `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
  const sameDay = baseDate && formatDateKey(date) === formatDateKey(baseDate);
  if (sameDay) return timeText;
  return `${String(date.getMonth() + 1).padStart(2, "0")}/${String(date.getDate()).padStart(2, "0")} ${timeText}`;
}

function formatRange(start, end, startRaw, endRaw) {
  if (start && end) return `${formatDateTime(start)} - ${formatDateTime(end)}`;
  return [startRaw, endRaw].filter(Boolean).join(" - ");
}

function formatAttendanceEvent(event) {
  const label = event.direction === "entry" ? "進" : "出";
  return `${label} ${formatDateTime(event.time)} ${event.sourceLabel}`;
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function addHours(date, hours) {
  return new Date(date.getTime() + hours * 36e5);
}

function addMinutes(date, minutes) {
  return new Date(date.getTime() + minutes * 60000);
}

function addSeconds(date, seconds) {
  return new Date(date.getTime() + seconds * 1000);
}

function earliestTime(values) {
  const items = values.filter(Boolean);
  return items.length ? items.sort((a, b) => a - b)[0] : null;
}

function latestTime(values) {
  const items = values.filter(Boolean);
  return items.length ? items.sort((a, b) => b - a)[0] : null;
}

function mergeDateArrays(left, right) {
  return uniqueBy([...(left || []), ...(right || [])], (item) => item.getTime()).sort((a, b) => a - b);
}

function dailyKey(person, dateKey) {
  return `${person}|${dateKey}`;
}

function groupBy(items, keyFn) {
  const map = new Map();
  items.forEach((item) => {
    const key = keyFn(item);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(item);
  });
  return map;
}

function unique(items) {
  return Array.from(new Set(items.filter((item) => item !== "" && item != null)));
}

function uniqueBy(items, keyFn) {
  const seen = new Set();
  return items.filter((item) => {
    const key = keyFn(item);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function sum(values) {
  return values.reduce((total, value) => total + (Number(value) || 0), 0);
}

function formatNumber(value, digits) {
  return Number(value || 0).toLocaleString("zh-TW", {
    minimumFractionDigits: 0,
    maximumFractionDigits: digits,
  });
}

function currency(value) {
  return Math.round(Number(value) || 0).toLocaleString("zh-TW", {
    style: "currency",
    currency: "TWD",
    maximumFractionDigits: 0,
  });
}

function riskLabel(value) {
  if (APPROVAL_MODE && value === ATTENDANCE_STATUS.LATE_OVERTIME.short) return approvalView.lateSignInLabel;
  if (APPROVAL_MODE && value === "unapproved") return "有未核給時數";
  if (APPROVAL_MODE && value === approvalView.limitLabel) return approvalView.limitLabel;
  return riskOptions.find((item) => item[0] === value)?.[1] || "全部";
}

function escapeHtml(value) {
  return cleanText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
