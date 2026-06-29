"use strict";

const dom = {};

const state = {
  workbookName: "",
  scheduleRows: [],
  wageRows: [],
  applicationRows: [],
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

const riskOptions = [
  ["all", "全部"],
  ["hasRisk", "僅異常"],
  ["例假日出勤", "例假日出勤"],
  ["12小時以上", "12小時以上"],
  ["間隔未達11小時", "間隔未達11小時"],
  ["46小時以上", "46小時以上"],
  ["差勤待覆核", "差勤待覆核"],
];

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
    "metricPay",
    "metricRiskDays",
    "metricMonthRisk",
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

  setStatus(`正在讀取：${file.name}`, false);
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    loadWorkbook(workbook, file.name);
  } catch (error) {
    console.error(error);
    setStatus("讀取失敗，請確認檔案為 Excel 格式且未受密碼保護。", true);
  }
}

function loadWorkbook(workbook, fileName) {
  const scheduleSheetName = findSheetByHeaders(workbook, ["日期", "班別"], ["排班"]);
  const wageSheetName = findSheetByHeaders(workbook, ["姓名", "年月", "時薪"], ["時薪"]);

  if (!scheduleSheetName) {
    setStatus("找不到可辨識的排班資料。請確認工作表含「日期」與「班別」欄位。", true);
    return;
  }

  state.workbookName = fileName;
  state.scheduleRows = parseScheduleSheet(workbook.Sheets[scheduleSheetName]);
  state.wageRows = wageSheetName ? parseWageSheet(workbook.Sheets[wageSheetName]) : [];
  state.applicationRows = parseApplicationSheets(workbook, scheduleSheetName, wageSheetName);
  buildAnalysis();
  setupControls();
  render();

  const modeText = state.dataSourceMode === "application"
    ? "已使用加班申請資料計算。"
    : "未找到完整加班申請表，已依排班加班時數推估，請人工覆核。";
  setStatus(`已匯入 ${fileName}。排班 ${state.scheduleRows.length} 筆、時薪 ${state.wageRows.length} 筆、加班明細 ${state.overtimeRecords.length} 筆。${modeText}`, state.dataSourceMode !== "application");
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

      const shift = cleanText(getField(row, ["班別"]));
      const scheduledHours = parseNumber(getField(row, ["班別時數", "班工時"]));
      const dailyWorkHours = parseNumber(getField(row, ["當日工時"]));
      const overtimeHours = parseNumber(getField(row, ["加班時數", "延展工時"]));
      const signInRaw = getField(row, ["簽到", "正常簽到"]);
      const signOutRaw = getField(row, ["簽退", "正常簽退"]);
      const expectedInRaw = getField(row, ["應簽到"]);
      const expectedOutRaw = getField(row, ["應簽退"]);
      const signIn = combineDateTime(date, signInRaw);
      const signOut = normalizeEndTime(signIn, combineDateTime(date, signOutRaw));
      const expectedIn = combineDateTime(date, expectedInRaw);
      const expectedOut = normalizeEndTime(expectedIn, combineDateTime(date, expectedOutRaw));
      const nationalHoliday = cleanText(getField(row, ["國定假日"]));
      const type = classifyDay(shift, nationalHoliday);

      return {
        rowIndex: index + 2,
        person,
        date,
        dateKey: formatDateKey(date),
        month: formatMonth(date),
        weekday: cleanText(getField(row, ["星期"])),
        shift,
        type,
        scheduledHours,
        dailyWorkHours,
        overtimeHours,
        signInRaw: cleanText(signInRaw),
        signOutRaw: cleanText(signOutRaw),
        expectedInRaw: cleanText(expectedInRaw),
        expectedOutRaw: cleanText(expectedOutRaw),
        signIn,
        signOut,
        expectedIn,
        expectedOut,
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

function parseApplicationSheets(workbook, scheduleSheetName, wageSheetName) {
  const applications = [];
  workbook.SheetNames.forEach((sheetName) => {
    if (sheetName === scheduleSheetName || sheetName === wageSheetName) return;
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    const headerIndex = rows.findIndex((row) => {
      const headers = row.map(normalizeHeader);
      return headers.includes("出勤者") && headers.includes("申請時數");
    });
    if (headerIndex < 0) return;

    const nextHeaderRow = rows[headerIndex + 1] || [];
    const headerRow = combineApplicationHeaders(rows[headerIndex], nextHeaderRow);
    const indexes = getApplicationIndexes(headerRow);
    if (indexes.date < 0 || indexes.person < 0 || indexes.hours < 0) return;
    const hasSubHeader = nextHeaderRow.some((value) => ["時薪", "金額", "起始時間", "終止時間", "進", "出"].includes(normalizeHeader(value)));

    rows.slice(headerIndex + (hasSubHeader ? 2 : 1)).forEach((row, offset) => {
      const date = parseDateValue(row[indexes.date]);
      const person = cleanText(row[indexes.person]);
      const hours = parseNumber(row[indexes.hours]);
      if (!date || !person || hours <= 0) return;
      const start = indexes.startNormalized >= 0
        ? combineDateTime(date, row[indexes.startNormalized])
        : combineDateTime(date, row[indexes.startDisplay]);
      const end = indexes.endNormalized >= 0
        ? combineDateTime(date, row[indexes.endNormalized])
        : combineDateTime(date, row[indexes.endDisplay]);

      applications.push({
        sourceSheet: sheetName,
        sourceRow: headerIndex + offset + 2,
        date,
        dateKey: formatDateKey(date),
        month: formatMonth(date),
        person,
        reason: cleanText(row[indexes.reason]),
        hours,
        category: cleanText(row[indexes.category]),
        audit: cleanText(row[indexes.audit]),
        shift: cleanText(row[indexes.shift]),
        hourly: parseNumber(row[indexes.hourly]),
        amount: parseNumber(row[indexes.amount]),
        start,
        end: normalizeEndTime(start, end),
        startRaw: cleanText(row[indexes.startDisplay]),
        endRaw: cleanText(row[indexes.endDisplay]),
        carIn: cleanText(row[indexes.carIn]),
        carOut: cleanText(row[indexes.carOut]),
      });
    });
  });
  return applications.sort((a, b) => a.person.localeCompare(b.person, "zh-Hant") || a.date - b.date);
}

function combineApplicationHeaders(primaryRow, secondaryRow) {
  const groupHeaders = new Set(["時數拆解", "加班費", "加班起訖", "汽車進出紀錄"]);
  return primaryRow.map((value, index) => {
    const primary = cleanText(value);
    const secondary = cleanText(secondaryRow[index]);
    if (!primary) return secondary;
    if (groupHeaders.has(normalizeHeader(primary)) && secondary) return secondary;
    return primary;
  });
}

function getApplicationIndexes(headerRow) {
  const person = headerRow.findIndex((value) => normalizeHeader(value) === "出勤者");
  const hours = headerRow.findIndex((value) => normalizeHeader(value) === "申請時數");
  const reason = headerRow.findIndex((value) => normalizeHeader(value) === "事由");
  const category = headerRow.findIndex((value) => normalizeHeader(value) === "類別");
  const audit = headerRow.findIndex((value) => normalizeHeader(value) === "稽核");
  const shift = headerRow.findIndex((value) => normalizeHeader(value) === "班別");
  const hourly = headerRow.findIndex((value) => normalizeHeader(value) === "時薪");
  const amount = headerRow.findIndex((value) => normalizeHeader(value) === "金額");
  const carIn = headerRow.findIndex((value) => normalizeHeader(value) === "進");
  const carOut = headerRow.findIndex((value) => normalizeHeader(value) === "出");
  const dateCandidates = headerRow
    .map((value, index) => ({ value: normalizeHeader(value), index }))
    .filter((item) => item.value === "日期" && item.index < person);
  const date = dateCandidates.length ? dateCandidates[dateCandidates.length - 1].index : headerRow.findIndex((value) => normalizeHeader(value) === "日期");

  const startLabels = headerRow
    .map((value, index) => ({ value: normalizeHeader(value), index }))
    .filter((item) => item.value === "起始時間");
  const endLabels = headerRow
    .map((value, index) => ({ value: normalizeHeader(value), index }))
    .filter((item) => item.value === "終止時間");

  return {
    date,
    person,
    hours,
    reason,
    category,
    audit,
    shift,
    hourly,
    amount,
    startDisplay: startLabels[0]?.index ?? -1,
    startNormalized: startLabels[1]?.index ?? startLabels[0]?.index ?? -1,
    endDisplay: endLabels[0]?.index ?? -1,
    endNormalized: endLabels[1]?.index ?? endLabels[0]?.index ?? -1,
    carIn,
    carOut,
  };
}

function buildAnalysis() {
  const wageMap = buildWageMap(state.wageRows);
  const scheduleByKey = new Map();
  const dailyMap = new Map();

  state.scheduleRows.forEach((row) => {
    const key = dailyKey(row.person, row.dateKey);
    scheduleByKey.set(key, row);
    if (!dailyMap.has(key)) {
      dailyMap.set(key, createDaily(row));
    } else {
      const daily = dailyMap.get(key);
      daily.scheduleRows.push(row);
      daily.scheduledHours += row.scheduledHours || row.dailyWorkHours || 0;
    }
  });

  const useApplications = state.applicationRows.length > 0;
  state.dataSourceMode = useApplications ? "application" : "schedule";
  const records = useApplications
    ? state.applicationRows.map((app) => recordFromApplication(app, scheduleByKey, wageMap))
    : state.scheduleRows.filter((row) => row.overtimeHours > 0).map((row) => recordFromSchedule(row, wageMap));

  records.forEach((record) => {
    if (!record) return;
    if (!dailyMap.has(dailyKey(record.person, record.dateKey))) {
      dailyMap.set(dailyKey(record.person, record.dateKey), createDailyFromRecord(record));
    }
    const daily = dailyMap.get(dailyKey(record.person, record.dateKey));
    daily.overtimeRecords.push(record);
    daily.overtimeHours += record.hours;
    daily.estimatedPay += record.estimatedPay;
    daily.type = record.type || daily.type;
    daily.shift = daily.shift || record.shift;
  });

  applyScheduleRisks(dailyMap);
  applyMonthlyOvertimeRisk(records, dailyMap);
  finalizeDailyRisks(dailyMap);

  state.overtimeRecords = records;
  state.dailyRows = Array.from(dailyMap.values()).sort((a, b) => a.date - b.date || a.person.localeCompare(b.person, "zh-Hant"));
  state.months = unique(state.dailyRows.map((item) => item.month)).sort();
  state.people = unique(state.dailyRows.map((item) => item.person)).sort((a, b) => a.localeCompare(b, "zh-Hant"));
  const overtimeMonths = unique(records.map((record) => record.month)).sort();
  state.selectedMonth = overtimeMonths.at(-1) || state.months.at(-1) || "";
  state.selectedPerson = "all";
  state.selectedRisk = "all";
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
    signIn: row.signIn,
    signOut: row.signOut,
    department: row.department,
    title: row.title,
    scheduleRows: [row],
    overtimeRecords: [],
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
    risks: [],
    riskDetails: [],
  };
}

function recordFromApplication(app, scheduleByKey, wageMap) {
  const schedule = scheduleByKey.get(dailyKey(app.person, app.dateKey));
  const type = app.category || schedule?.type || classifyDay(app.shift || schedule?.shift || "", "");
  const hourly = app.hourly || wageMap.get(`${app.person}|${app.month}`) || wageMap.get(`*|${app.month}`) || 0;
  const fee = app.amount || calculateOvertimePay(app.hours, hourly, type).total;
  const shift = app.shift || schedule?.shift || "";
  const auditRisks = auditTextToRisks(app.audit);

  return {
    id: `${app.person}-${app.dateKey}-${app.sourceSheet}-${app.sourceRow}`,
    source: "加班申請",
    person: app.person,
    date: app.date,
    dateKey: app.dateKey,
    month: app.month,
    shift,
    type,
    reason: app.reason,
    hours: app.hours,
    hourly,
    estimatedPay: fee,
    segments: calculateOvertimePay(app.hours, hourly, type).segments,
    audit: app.audit,
    start: app.start,
    end: app.end,
    startRaw: app.startRaw,
    endRaw: app.endRaw,
    carIn: app.carIn,
    carOut: app.carOut,
    risks: auditRisks,
  };
}

function recordFromSchedule(row, wageMap) {
  const hourly = wageMap.get(`${row.person}|${row.month}`) || wageMap.get(`*|${row.month}`) || 0;
  const pay = calculateOvertimePay(row.overtimeHours, hourly, row.type);
  return {
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
    estimatedPay: pay.total,
    segments: pay.segments,
    audit: "待覆核",
    start: null,
    end: null,
    startRaw: "",
    endRaw: "",
    carIn: "",
    carOut: "",
    risks: ["差勤待覆核"],
  };
}

function applyScheduleRisks(dailyMap) {
  state.scheduleRows.forEach((row) => {
    const daily = dailyMap.get(dailyKey(row.person, row.dateKey));
    if (!daily) return;
    if (row.type === "例假日" && (row.overtimeHours > 0 || row.signIn || row.signOut)) {
      addRisk(daily, "例假日出勤", "例假日有出勤或加班紀錄，需確認調班、補休或主管核准紀錄。");
    }
    if (row.overtimeHours > 0) {
      if (!row.signIn || !row.signOut) {
        addRisk(daily, "差勤待覆核", "加班時數有值，但簽到或簽退資料不完整。");
      }
      if (row.expectedOut && row.signOut) {
        const estimatedEnd = addHours(row.expectedOut, row.overtimeHours);
        if (row.signOut < addMinutes(estimatedEnd, -10)) {
          addRisk(daily, "差勤待覆核", "簽退時間可能不足以佐證排班加班時數。");
        }
      }
    }
  });

  const grouped = groupBy(state.scheduleRows, (row) => row.person);
  grouped.forEach((rows) => {
    rows.sort((a, b) => a.date - b.date);
    rows.forEach((row, index) => {
      const next = rows[index + 1];
      if (!next || !row.signOut || !next.signIn) return;
      const restHours = (next.signIn - row.signOut) / 36e5;
      if (restHours >= 0 && restHours < 11) {
        const daily = dailyMap.get(dailyKey(row.person, row.dateKey));
        addRisk(daily, "間隔未達11小時", `至次日出勤間隔約 ${formatNumber(restHours, 1)} 小時。`);
      }
    });
  });
}

function applyMonthlyOvertimeRisk(records, dailyMap) {
  const grouped = groupBy(records, (record) => `${record.person}|${record.month}`);
  grouped.forEach((items) => {
    items.sort((a, b) => a.date - b.date);
    let cumulative = 0;
    items.forEach((record) => {
      cumulative += record.hours;
      record.monthlyCumulative = cumulative;
      if (cumulative > 46) {
        record.risks.push("46小時以上");
        const daily = dailyMap.get(dailyKey(record.person, record.dateKey));
        addRisk(daily, "46小時以上", `當月累計加班約 ${formatNumber(cumulative, 1)} 小時。`);
      }
    });
  });
}

function finalizeDailyRisks(dailyMap) {
  dailyMap.forEach((daily) => {
    daily.overtimeRecords.forEach((record) => {
      record.risks.forEach((risk) => addRisk(daily, risk, record.audit || record.reason || risk));
    });
    const totalWorkHours = daily.scheduledHours + daily.overtimeHours;
    daily.totalWorkHours = totalWorkHours;
    if (totalWorkHours > 12) {
      addRisk(daily, "12小時以上", `當日排班與加班合計約 ${formatNumber(totalWorkHours, 1)} 小時。`);
    }
    daily.risks = unique(daily.risks);
    daily.riskDetails = unique(daily.riskDetails);
  });
}

function setupControls() {
  setSelectOptions(dom.monthSelect, state.months.map((month) => [month, month]), state.selectedMonth);
  setSelectOptions(dom.personSelect, [["all", "全部"], ...state.people.map((person) => [person, person])], "all");
  setSelectOptions(dom.riskSelect, riskOptions, "all");
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
    if (state.selectedRisk === "hasRisk" && daily.risks.length === 0) return false;
    if (!["all", "hasRisk"].includes(state.selectedRisk) && !daily.risks.includes(state.selectedRisk)) return false;
    return true;
  });
}

function getFilteredRecords() {
  return state.overtimeRecords.filter((record) => {
    if (state.selectedMonth && record.month !== state.selectedMonth) return false;
    if (state.selectedPerson !== "all" && record.person !== state.selectedPerson) return false;
    if (state.selectedRisk === "hasRisk" && record.risks.length === 0) return false;
    if (!["all", "hasRisk"].includes(state.selectedRisk) && !record.risks.includes(state.selectedRisk)) return false;
    return true;
  });
}

function renderMetrics(dailyRows, records) {
  const totalHours = sum(records.map((record) => record.hours));
  const totalPay = sum(records.map((record) => record.estimatedPay));
  const riskDays = dailyRows.filter((daily) => daily.risks.length > 0).length;
  const over46People = unique(records.filter((record) => (record.monthlyCumulative || 0) > 46).map((record) => record.person)).length;
  dom.metricHours.textContent = formatNumber(totalHours, 1);
  dom.metricPay.textContent = currency(totalPay);
  dom.metricRiskDays.textContent = `${riskDays}`;
  dom.metricMonthRisk.textContent = `${over46People} 人`;
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
    cell.addEventListener("click", () => {
      state.selectedDateKey = dateKey;
      renderDetail(dailyRows);
      renderCalendar(dailyRows);
    });
    cell.innerHTML = dayCellHtml(current, items);
    dom.calendarGrid.appendChild(cell);
  }
}

function dayCellHtml(date, items) {
  const totalHours = sum(items.map((item) => item.overtimeHours));
  const totalPay = sum(items.map((item) => item.estimatedPay));
  const risks = unique(items.flatMap((item) => item.risks));
  const first = items[0];
  const badge = risks.length ? `<span class="badge danger">${risks.length}</span>` : totalHours ? `<span class="badge warn">${formatNumber(totalHours, 1)}h</span>` : "";
  const personLine = items.length > 1 ? `${items.length} 筆 / ${unique(items.map((item) => item.person)).length} 人` : first?.person || "";
  const shiftLine = first?.shift || "";
  const riskLine = risks.slice(0, 2).join("、");
  const payLine = totalPay ? currency(totalPay) : "";

  return `
    <div class="day-top">
      <span class="day-number">${date.getDate()}</span>
      ${badge}
    </div>
    <div class="day-lines">
      <span>${escapeHtml(personLine)}</span>
      <span>${escapeHtml(shiftLine)}</span>
      <span>${totalHours ? `加班 ${formatNumber(totalHours, 1)}h` : ""}</span>
      <span>${escapeHtml(payLine)}</span>
      <span>${escapeHtml(riskLine)}</span>
    </div>
  `;
}

function renderDetail(dailyRows) {
  const selectedKey = state.selectedDateKey || dailyRows[0]?.dateKey || "";
  state.selectedDateKey = selectedKey;
  const items = dailyRows.filter((daily) => daily.dateKey === selectedKey);
  if (!selectedKey || items.length === 0) {
    dom.detailTitle.textContent = "明細";
    dom.detailSubtitle.textContent = "選取日期查看稽核內容";
    dom.detailBody.className = "detail-body empty-state";
    dom.detailBody.textContent = "此日期無符合目前篩選的資料。";
    return;
  }

  const date = parseDateValue(selectedKey);
  const totalHours = sum(items.map((item) => item.overtimeHours));
  const totalPay = sum(items.map((item) => item.estimatedPay));
  const risks = unique(items.flatMap((item) => item.risks));
  const records = items.flatMap((item) => item.overtimeRecords);

  dom.detailTitle.textContent = formatDisplayDate(date);
  dom.detailSubtitle.textContent = `${items.length} 筆日資料，${records.length} 筆加班明細`;
  dom.detailBody.className = "detail-body";
  dom.detailBody.innerHTML = `
    <div class="detail-stack">
      <div class="detail-row"><span>人員</span><strong>${escapeHtml(unique(items.map((item) => item.person)).join("、"))}</strong></div>
      <div class="detail-row"><span>加班時數</span><strong>${formatNumber(totalHours, 1)} 小時</strong></div>
      <div class="detail-row"><span>估算金額</span><strong>${currency(totalPay)}</strong></div>
      <div class="detail-row"><span>風險</span><div class="risk-list">${risks.length ? risks.map((risk) => `<span class="badge danger">${escapeHtml(risk)}</span>`).join("") : '<span class="badge">無</span>'}</div></div>
      <div class="record-list">${records.length ? records.map(recordHtml).join("") : '<div class="empty-state">無加班明細</div>'}</div>
    </div>
  `;
}

function recordHtml(record) {
  const segments = record.segments
    .filter((segment) => segment.hours > 0)
    .map((segment) => `${segment.label} ${formatNumber(segment.hours, 1)}h`)
    .join("、");
  return `
    <article class="record-item">
      <strong>${escapeHtml(record.person)}｜${escapeHtml(record.type)}｜${formatNumber(record.hours, 1)}h</strong>
      <p>${escapeHtml(record.reason || record.source)}</p>
      <p>時薪 ${record.hourly ? currency(record.hourly) : "未提供"}，估算 ${currency(record.estimatedPay)}</p>
      <p>${escapeHtml(segments || "未拆分")}</p>
      <p>${escapeHtml(record.audit || "待覆核")}</p>
    </article>
  `;
}

function renderRiskTable(dailyRows) {
  const riskRows = dailyRows.filter((daily) => daily.risks.length > 0);
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
      <td>${escapeHtml(daily.riskDetails.join("；"))}</td>
    </tr>
  `).join("");
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
  const rows = getFilteredDailyRows();
  const grouped = groupBy(rows, (row) => `${row.month}|${row.person}`);
  const data = [["月份", "人員", "加班時數", "估算加班費", "異常天數", "異常類型"]];
  grouped.forEach((items, key) => {
    const [month, person] = key.split("|");
    data.push([
      month,
      person,
      formatNumber(sum(items.map((item) => item.overtimeHours)), 2),
      Math.round(sum(items.map((item) => item.estimatedPay))),
      items.filter((item) => item.risks.length > 0).length,
      unique(items.flatMap((item) => item.risks)).join("、"),
    ]);
  });
  downloadCsv(data, `異常工時摘要_${state.selectedMonth || "全部"}.csv`);
}

function exportDetailCsv() {
  const data = [["日期", "月份", "人員", "班別", "類別", "加班時數", "時薪", "估算加班費", "來源", "稽核", "異常", "事由"]];
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
      record.source,
      record.audit,
      unique(record.risks).join("、"),
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
  dom.fileStatus.classList.toggle("warning", Boolean(warning));
}

function calculateOvertimePay(hours, hourly, type) {
  const safeHours = Math.max(0, Number(hours) || 0);
  const safeHourly = Math.max(0, Number(hourly) || 0);
  const segments = [];
  let remaining = safeHours;

  const addSegment = (label, takeHours, multiplier) => {
    const used = Math.min(remaining, takeHours);
    if (used > 0) {
      segments.push({
        label,
        hours: used,
        multiplier,
        amount: Math.ceil(used * safeHourly * multiplier),
      });
      remaining -= used;
    }
  };

  if (["例假日", "固定假日"].includes(type)) {
    addSegment("當日工資/補休", 8, 1);
    addSegment("延長1-2H", 2, 1.34);
    addSegment("延長3-4H", 2, 1.67);
    addSegment("延長應稅", Number.POSITIVE_INFINITY, 1.67);
  } else if (type === "休息日") {
    addSegment("休息日1-2H", 2, 1.34);
    addSegment("休息日3-8H", 6, 1.67);
    addSegment("休息日9H以上", Number.POSITIVE_INFINITY, 2.67);
  } else {
    addSegment("平日1-2H", 2, 1.34);
    addSegment("平日3H以上", Number.POSITIVE_INFINITY, 1.67);
  }

  return {
    total: sum(segments.map((segment) => segment.amount)),
    segments,
  };
}

function auditTextToRisks(text) {
  const value = cleanText(text);
  if (!value || value === "正常") return [];
  const risks = ["差勤待覆核"];
  if (value.includes("12") || value.includes("小時")) risks.push("12小時以上");
  return risks;
}

function addRisk(daily, risk, detail) {
  if (!daily || !risk) return;
  daily.risks.push(risk);
  if (detail) daily.riskDetails.push(`${risk}：${detail}`);
}

function classifyDay(shift, nationalHoliday) {
  const value = `${shift || ""} ${nationalHoliday || ""}`;
  if (/國定|固定/.test(value)) return "固定假日";
  if (/例假/.test(value)) return "例假日";
  if (/休息/.test(value)) return "休息日";
  return "平日";
}

function getField(row, names) {
  const entries = Object.entries(row);
  for (const name of names) {
    if (Object.prototype.hasOwnProperty.call(row, name) && row[name] !== "") return row[name];
    const normalized = normalizeHeader(name);
    const found = entries.find(([key, value]) => normalizeHeader(key) === normalized && value !== "");
    if (found) return found[1];
  }
  return "";
}

function normalizeHeader(value) {
  return cleanText(value).replace(/\s+/g, "").replace(/[_｜|]/g, "");
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

function combineDateTime(baseDate, value) {
  if (!baseDate || value == null || cleanText(value) === "" || cleanText(value) === "--") return null;
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

  const text = cleanText(value);
  const parsedFull = parseDateTimeText(text, baseDate);
  if (parsedFull) return parsedFull;
  const time = parseTimeText(text);
  if (!time) return null;
  return new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), time.hour, time.minute, time.second);
}

function parseDateTimeText(text, baseDate) {
  const ymd = text.match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2}).*?(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if (ymd) {
    const time = normalizeHour(Number(ymd[5]), ymd[4]);
    return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]), time, Number(ymd[6]), Number(ymd[7] || 0));
  }
  const md = text.match(/(\d{1,2})[/-](\d{1,2}).*?(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if (md) {
    const time = normalizeHour(Number(md[4]), md[3]);
    let result = new Date(baseDate.getFullYear(), Number(md[1]) - 1, Number(md[2]), time, Number(md[5]), Number(md[6] || 0));
    if (result < addDays(baseDate, -10)) result = new Date(baseDate.getFullYear() + 1, Number(md[1]) - 1, Number(md[2]), time, Number(md[5]), Number(md[6] || 0));
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
