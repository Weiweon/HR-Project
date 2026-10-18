"use strict";

// This entry shares attendance rules with the original view, but owns approval accounting.
const approvalView = {
  limitLabel: "超出日/月延展工時無核給",
  lateSignInNote: "起始時間往前推估半小時",
  get lateSignInLabel() { return `簽到晚，${this.lateSignInNote}`; },
  confirmedCarOwners: new Map([["【AI作業檔】心臟血管外科異常工時1151018.xlsx", "姚曉霈"]]),
  monthlyApproved: new Map(),
  importNotes: [],
  unnamedCarRows: [],
  sourceCounts: {},

  loadWorkbook(workbook, fileName) {
    const findName = (pattern) => workbook.SheetNames.find((name) => pattern.test(normalizeHeader(name)));
    const approvedName = findName(/^核可加班單(?:申請)?$/);
    const requestedName = findName(/^未核可加班單(?:申請)?$/);
    const scheduleName = findName(/^排班$/);
    if (!scheduleName || !approvedName || !requestedName) {
      setStatus("請匯入含「排班」、「核可｜加班單申請」及「未核可｜加班單申請」的 Excel。", true);
      return;
    }
    const approvedGrid = XLSX.utils.sheet_to_json(workbook.Sheets[approvedName], { header: 1, raw: true, defval: "" });
    const requestedGrid = XLSX.utils.sheet_to_json(workbook.Sheets[requestedName], { header: 1, raw: true, defval: "" });
    const approvedHeader = approvedGrid.slice(0, 10).find((row) => normalizeHeader(row[11]) === "申請時數" && normalizeHeader(row[12]) === "金額");
    const requestedHeader = requestedGrid.slice(0, 10).find((row) => normalizeHeader(row[4]) === "申請時數");
    if (!approvedHeader || !requestedHeader) {
      setStatus("分頁欄位不符：未核可 E 欄須為申請時數，核可 L／M 欄須為申請時數／金額。請確認後重新匯入。", true);
      return;
    }
    const wageName = findName(/^時薪$/);
    const carName = findName(/^汽車進出紀錄$/);
    const parseSheet = (name) => dedupeApplicationRows(parseApplicationSheets(
      workbook, new Set(workbook.SheetNames.filter((item) => item !== name)),
    ));
    const approved = parseSheet(approvedName);
    const requests = parseSheet(requestedName);
    const grid = approvedGrid;
    this.importNotes = [];
    approved.forEach((row) => {
      const amounts = unique(row.sourceRows.map((number) => this.numericAmount(grid[number - 1]?.[12])).filter((amount) => amount !== null));
      row.reportedAmount = amounts[0] ?? null;
      row.hours = parseNumber(grid[row.sourceRow - 1]?.[11]);
      row.approvalNotes = amounts.length > 1 ? ["重複核可單的 M 欄金額不同，請覆核。"] : [];
      if (unique(row.sourceRows.map((number) => parseNumber(grid[number - 1]?.[11]))).length > 1) row.approvalNotes.push("重複核可單的 L 欄時數不同，暫採第一列，請覆核。");
      this.importNotes.push(...row.approvalNotes);
    });
    this.monthlyApproved = new Map();
    approved.forEach((row) => {
      const key = `${row.person}|${row.month}`;
      this.monthlyApproved.set(key, (this.monthlyApproved.get(key) || 0) + row.hours);
    });

    state.workbookName = fileName;
    state.scheduleRows = parseScheduleSheet(workbook.Sheets[scheduleName]);
    state.wageRows = wageName ? parseWageSheet(workbook.Sheets[wageName]) : [];
    state.applicationRows = this.reconcile(requests, approved);
    state.carRows = carName ? parseCarSheet(workbook.Sheets[carName]) : [];
    this.assignConfirmedCarOwner(state.carRows, unique([...state.scheduleRows, ...state.applicationRows].map((row) => row.person)), fileName);
    assignCarPeople(state.carRows, state.scheduleRows, state.applicationRows);
    this.unnamedCarRows = state.carRows.filter((row) => !row.person);
    this.sourceCounts = { requests: requests.length, approved: approved.length };
    buildAnalysis();
    setupControls();
    render();
    this.showImportStatus();
  },

  assignConfirmedCarOwner(carRows, people, fileName) {
    const owner = this.confirmedCarOwners.get(cleanText(fileName));
    if (!owner || !people.includes(owner)) return;
    carRows.forEach((row) => { if (!row.person) row.person = owner; });
  },

  selectCarPerson(person) {
    if (person && !state.people.includes(person)) return;
    const previous = { month: state.selectedMonth, person: state.selectedPerson, risk: state.selectedRisk, date: state.selectedDateKey };
    this.unnamedCarRows.forEach((row) => { row.person = person; });
    buildAnalysis();
    state.selectedMonth = previous.month;
    state.selectedPerson = previous.person;
    state.selectedRisk = previous.risk;
    state.selectedDateKey = previous.date;
    render();
    this.showImportStatus();
  },

  showImportStatus() {
    setStatus("", false);
  },

  numericAmount(value) {
    if (typeof value === "number") return Number.isFinite(value) && value >= 0 ? value : null;
    const text = cleanText(value).replace(/[,\s$＄]/g, "");
    return /^\d+(?:\.\d+)?$/.test(text) ? Number(text) : null;
  },

  reconcile(requests, approved) {
    const rows = requests.map((row) => ({ ...row, approvals: [], requestOrigin: "未核可 E 欄", approvalNotes: [] }));
    const originalRequests = [...rows];
    approved.forEach((grant) => {
      const candidates = originalRequests.filter((row) => row.person === grant.person && row.dateKey === grant.dateKey)
        .map((row) => ({ row, score: this.matchScore(row, grant) }))
        .filter((item) => item.score > 0).sort((a, b) => b.score - a.score);
      if (candidates.length && (candidates.length === 1 || candidates[0].score > candidates[1].score)) {
        candidates[0].row.approvals.push(grant);
      } else {
        const ambiguous = candidates.length > 1;
        rows.push({
          ...grant, approvals: [grant], requestOrigin: "僅核可表 L 欄", approvalNotes: ambiguous ? ["同日有多筆可能對應申請，核給配對待確認。"] : [],
        });
        if (ambiguous) this.importNotes.push("有核給單無法唯一配對申請，已獨立列示，請覆核。");
      }
    });
    return rows.sort((a, b) => a.person.localeCompare(b.person, "zh-Hant") || a.date - b.date || (a.start || 0) - (b.start || 0));
  },

  matchScore(request, grant) {
    if (!request.start || !request.end || !grant.start || !grant.end) return 0;
    const sameReason = request.reason === grant.reason;
    if (+request.start === +grant.start && +request.end === +grant.end) return 100 + Number(sameReason);
    if (!sameReason) return 0;
    if (+request.start === +grant.start) return 60;
    return Math.min(+request.end, +grant.end) > Math.max(+request.start, +grant.start) ? 20 : 0;
  },

  resolveHourly(person, month, wageMap) {
    const exact = wageMap.get(`${person}|${month}`);
    if (exact > 0) return exact;
    const baseName = cleanText(person).replace(/\s*[（(]?大林(?:分院)?[）)]?$/, "").trim();
    if (!baseName || baseName === person) return 0;
    return wageMap.get(`${baseName}|${month}`) || 0;
  },

  estimateGrant(grant, wageMap, type) {
    const hourly = this.resolveHourly(grant.person, grant.month, wageMap);
    const result = calculateOvertimePay(grant.hours, hourly, type);
    return {
      ...grant,
      hourly,
      estimatedPay: grant.reportedAmount ?? (hourly > 0 ? result.total : null),
      paySource: grant.reportedAmount !== null ? "核可 M 欄" : hourly > 0 ? "依時薪估算" : "缺少當月時薪",
      segments: result.segments,
      payNote: result.note,
    };
  },

  enrichRecord(record, app, wageMap) {
    record.hourly = this.resolveHourly(record.person, record.month, wageMap);
    record.requestOrigin = app.requestOrigin;
    record.sourceReference = `${app.sourceSheet} 第 ${app.sourceRow} 列`;
    record.approvals = app.approvals.map((grant) => this.estimateGrant(grant, wageMap, record.type));
    record.approvedHours = sum(record.approvals.map((grant) => grant.hours));
    record.unapprovedHours = Math.max(0, record.hours - record.approvedHours);
    record.approvalNotes = unique([...app.approvalNotes, ...record.approvals.flatMap((grant) => grant.approvalNotes || [])]);
    record.payMissing = record.approvals.some((grant) => grant.estimatedPay === null);
    record.estimatedPay = sum(record.approvals.map((grant) => grant.estimatedPay || 0));
    record.monthApprovedHours = this.monthlyApproved.get(`${record.person}|${record.month}`) || 0;
    record.approvalAlerts = [];
    if (record.approvedHours > record.hours) record.approvalNotes.push("核給時數大於對應申請時數，請覆核來源。");
    if (record.payMissing) record.approvalNotes.push("核可 M 欄空白且缺少本人當月時薪，加班費待補。");
    const dayRows = state.applicationRows.filter((row) => row.person === record.person && row.dateKey === record.dateKey);
    const dayRequestedHours = sum(dayRows.map((row) => row.hours));
    const dailyLimitExceeded = record.type === DAY_TYPE.WORKDAY ? dayRequestedHours > 4 : dayRequestedHours > 12;
    if (record.unapprovedHours > 0 && (record.monthApprovedHours >= 46 || dailyLimitExceeded)) {
      record.approvalAlerts.push(this.limitLabel);
      record.approvalNotes.push("日／月額度未核給原因依申請與核給資料推估，需以簽核紀錄確認。");
    }
    if (record.unapprovedHours > 0 && !record.approvalAlerts.length) {
      record.approvalNotes.push(record.type === DAY_TYPE.WORKDAY && record.hours > 4
        ? "平日申請超過 4 小時；未核給部分請依退單及簽核紀錄確認。"
        : "未核給原因待確認，請依退單及簽核紀錄確認。");
    }
    record.payNote = "僅計核給部分";
  },

  finalizeDaily(daily) {
    daily.approvedHours = sum(daily.overtimeRecords.map((record) => record.approvedHours));
    daily.unapprovedHours = sum(daily.overtimeRecords.map((record) => record.unapprovedHours));
    daily.payMissing = daily.overtimeRecords.some((record) => record.payMissing);
    // Retain monthly approved totals even on dates with no application.
    daily.monthApprovedHours = this.monthlyApproved.get(`${daily.person}|${daily.month}`) || 0;
    if (daily.monthApprovedHours > 46 && daily.overtimeRecords.length) {
      addWorkRisk(daily, WORK_RISK.OVER_46_HOURS, `當月核給共 ${formatNumber(daily.monthApprovedHours, 1)} 小時，超過 46 小時。`);
    }
  },

  monthRows() {
    return state.dailyRows.filter((row) => (!state.selectedMonth || row.month === state.selectedMonth)
      && (state.selectedPerson === "all" || row.person === state.selectedPerson));
  },

  payText(rows) {
    const total = currency(sum(rows.map((row) => row.estimatedPay || 0)));
    return rows.some((row) => row.payMissing) ? `${total}＋待補` : total;
  },

  renderMetrics() {
    const rows = this.monthRows();
    const approved = sum(rows.map((row) => row.approvedHours));
    dom.metricHours.textContent = formatNumber(sum(rows.map((row) => row.overtimeHours)), 1);
    dom.metricApprovedHours.textContent = formatNumber(approved, 1);
    dom.metricApprovedHours.closest(".metric").classList.toggle("approval-over-limit", rows.some((row) => row.monthApprovedHours > 46));
    dom.metricPay.textContent = this.payText(rows);
    dom.metricAttendanceRiskDays.textContent = unique(rows.filter(dailyHasAttendanceRisk).map((row) => row.dateKey)).length;
    dom.metricWorkRiskDays.textContent = unique(rows.filter((row) => row.workRisks.some((risk) => risk !== WORK_RISK.PENDING_REVIEW)).map((row) => row.dateKey)).length;
  },

  dayMetricRows(items) {
    const records = items.flatMap((item) => item.overtimeRecords);
    if (!records.length) return [];
    const rows = [
      dayLineHtml(`加班時數申請 ${formatNumber(sum(items.map((item) => item.overtimeHours)), 1)} h`, "metric metric-hours"),
      dayLineHtml(`⏱️延展工時核給 ${formatNumber(sum(items.map((item) => item.approvedHours)), 1)} h`, "metric metric-approved"),
      dayLineHtml(`💰加班費${this.payText(items)}`, "metric metric-pay"),
    ];
    const alerts = unique(records.flatMap((record) => record.approvalAlerts));
    if (alerts.length) rows.push(`<div class="day-pill-row">${alerts.map((label) => dayPillHtml(label, "approval-denied")).join("")}</div>`);
    return rows;
  },

  calendarWorkRisks(daily) {
    return daily.workRisks.filter((risk) => risk !== WORK_RISK.PENDING_REVIEW
      && !(risk === WORK_RISK.OVER_46_HOURS && daily.approvedHours > 0));
  },

  visibleWorkRiskDetails(daily) {
    return daily.workRiskDetails.filter((detail) => !(daily.approvedHours > 0 && detail.includes(WORK_RISK.OVER_46_HOURS)));
  },

  visibleApprovalNotes(record) {
    return record.approvalNotes.filter((note) => !/原因依.*推估/.test(note));
  },

  detailTotals(items) {
    return [
      ["加班時數申請", `${formatNumber(sum(items.map((item) => item.overtimeHours)), 1)} 小時`],
      ["延展工時核給", `${formatNumber(sum(items.map((item) => item.approvedHours)), 1)} 小時`],
      ["未核給時數", `${formatNumber(sum(items.map((item) => item.unapprovedHours)), 1)} 小時`],
      ["加班費（核給部分）", this.payText(items)],
    ].map(([label, value]) => `<div class="detail-row"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join("");
  },

  grantHtml(grant, type) {
    const range = formatRange(grant.start, grant.end, grant.startRaw, grant.endRaw);
    const amount = grant.estimatedPay === null ? "待補金額" : currency(grant.estimatedPay);
    const segments = grant.reportedAmount === null && grant.hourly > 0
      ? `<div>估算倍率 ${escapeHtml(formatRecordSegmentSummary({ hours: grant.hours, type, segments: grant.segments }))}</div>` : "";
    return `<li><div>${escapeHtml(range)}</div><div>核給 ${formatNumber(grant.hours, 1)} h · ${escapeHtml(amount)}</div>${segments}</li>`;
  },

  recordHtml(record) {
    const issues = buildAttendanceSummaryItems(record);
    const attendance = issues.length ? issues.map((item) => statusBadgeHtml(item.label, attendancePillClass(item.status))).join("") : statusBadgeHtml("正常", "attendance ok");
    const visibleIssues = buildRecordIssueItems(record).filter((detail) => !(record.approvedHours > 0 && detail.includes(WORK_RISK.OVER_46_HOURS)));
    return `<article class="record-item">
      <div class="record-header"><strong>${escapeHtml(formatRange(record.start, record.end, record.startRaw, record.endRaw))}</strong></div>
      <p class="record-reason">${escapeHtml(record.reason || record.source)}</p>
      <ul class="record-bullets">
        <li>加班申請 ${formatNumber(record.hours, 1)} h；核給 ${formatNumber(record.approvedHours, 1)} h；未核給 ${formatNumber(record.unapprovedHours, 1)} h</li>
        <li>加班費 ${escapeHtml(this.payText([record]))}（僅核給部分）</li>
      </ul>
      <div class="risk-list">${record.approvalAlerts.map((label) => dayPillHtml(label, "approval-denied")).join("")}${attendance}</div>
      ${record.approvals.length ? `<section class="record-mini-section"><h4>核給明細</h4><ul class="record-list-compact">${record.approvals.map((grant) => this.grantHtml(grant, record.type)).join("")}</ul></section>` : ""}
      ${this.visibleApprovalNotes(record).length ? `<section class="record-mini-section"><h4>核給說明</h4>${renderCompactList(this.visibleApprovalNotes(record))}</section>` : ""}
      <section class="record-mini-section"><h4>出勤佐證</h4>${renderCompactList(buildRecordEvidenceItems(record))}</section>
      <section class="record-mini-section"><h4>異常說明</h4>${renderCompactList(visibleIssues.length ? visibleIssues : ["無"])}</section>
    </article>`;
  },

  renderRiskTable(dailyRows) {
    const rows = dailyRows.filter(dailyHasAnyRisk);
    dom.riskSubtitle.textContent = `${rows.length} 筆日期需覆核`;
    dom.riskTableBody.innerHTML = rows.length ? rows.map((daily) => `<tr>
      <td>${escapeHtml(formatDisplayDate(daily.date))}</td><td>${escapeHtml(daily.person)}</td><td>${escapeHtml(daily.shift || daily.type)}</td>
      <td>${formatNumber(daily.overtimeHours, 1)} 小時</td><td>${formatNumber(daily.approvedHours, 1)} 小時</td><td>${formatNumber(daily.unapprovedHours, 1)} 小時</td>
      <td>${escapeHtml(this.payText([daily]))}</td>
      <td>${renderRiskSummaryCell(daily)}${daily.overtimeRecords.filter((record) => record.unapprovedHours > 0 || record.approvalNotes.length).map((record) => `<div class="approval-summary"><strong>${escapeHtml(shortRangeLabel(record.start, record.end, record.startRaw, record.endRaw))}</strong><span>核給 ${formatNumber(record.approvedHours, 1)} h／未核給 ${formatNumber(record.unapprovedHours, 1)} h</span>${record.approvalAlerts.map((label) => dayPillHtml(label, "approval-denied")).join("")}${this.visibleApprovalNotes(record).length ? renderCompactList(this.visibleApprovalNotes(record)) : ""}</div>`).join("")}</td>
    </tr>`).join("") : '<tr><td colspan="8" class="table-empty">目前篩選無異常資料</td></tr>';
  },

  exportSummaryCsv() {
    const data = [["月份", "人員", "加班時數申請", "月延展工時核給", "未核給時數", "估算金額（核給部分）", "金額狀態", "出勤異常天數", "延展工時異常天數"]];
    groupBy(this.monthRows(), (row) => `${row.month}|${row.person}`).forEach((rows) => {
      data.push([rows[0].month, rows[0].person, sum(rows.map((row) => row.overtimeHours)), sum(rows.map((row) => row.approvedHours)), sum(rows.map((row) => row.unapprovedHours)),
        sum(rows.map((row) => row.estimatedPay)), rows.some((row) => row.payMissing) ? "金額未齊" : "完整", unique(rows.filter(dailyHasAttendanceRisk).map((row) => row.dateKey)).length,
        unique(rows.filter((row) => row.workRisks.length).map((row) => row.dateKey)).length]);
    });
    downloadCsv(data, `申請核給摘要_${state.selectedMonth || "全部"}.csv`);
  },

  exportDetailCsv() {
    const data = [["日期", "月份", "人員", "班別", "加班申請", "核給時數", "未核給時數", "估算金額（核給部分）", "金額狀態", "申請起始", "申請終止", "申請來源", "核給來源", "金額來源", "核給說明", "出勤狀況", "延展工時異常"]];
    getFilteredRecords().forEach((record) => data.push([
      record.dateKey, record.month, record.person, record.shift, record.hours, record.approvedHours, record.unapprovedHours, record.estimatedPay,
      record.payMissing ? "金額未齊" : "完整", formatDateTime(record.start), formatDateTime(record.end), record.sourceReference,
      record.approvals.map((grant) => `${grant.sourceSheet} 第${grant.sourceRow}列 L/M`).join("；"), unique(record.approvals.map((grant) => grant.paySource)).join("；"),
      [...record.approvalAlerts, ...record.approvalNotes].join("；"), record.attendanceLabel, record.workRisks.join("；"),
    ]));
    downloadCsv(data, `申請核給明細_${state.selectedMonth || "全部"}.csv`);
  },
};
