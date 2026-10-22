"use strict";

const payroll = {
  rates: [1, 1.334, 1.667, 2.667, 1.667],
  labels: ["1倍", "1.334倍", "1.667倍", "2.667倍", "1.667倍(含稅)"],

  parseApproved(workbook, name) {
    const grid = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: "" });
    const header = grid.findIndex((row) => normalizeHeader(row[0]) === "日期" && normalizeHeader(row[5]) === "加班總時數");
    if (header < 0 || normalizeHeader(grid[header][6]) !== "加班起始時間"
      || normalizeHeader(grid[header][7]) !== "加班終止時間"
      || normalizeHeader(grid[header][10]) !== "加班費倍率") {
      throw new Error("核可表須含 F 欄「加班總時數」、G／H 欄加班起訖及 K～O 欄倍率時數。");
    }
    const bands = grid[header + 1] || [];
    if (!["1", "1.334", "1.667", "2.667", "1.667(應稅)"].every((value, i) => normalizeHeader(bands[10 + i]) === normalizeHeader(value))) {
      throw new Error("核可表第二列表頭須依序為 1、1.334、1.667、2.667、1.667(應稅)。");
    }
    const rows = [];
    grid.slice(header + 2).forEach((row, index) => {
      const sourcePerson = cleanText(row[1]);
      const person = personSettings.canonicalName(sourcePerson);
      if (!person) return;
      const sourceRow = header + index + 3;
      const date = parseDateValue(row[0]);
      const hours = this.number(row[5]);
      if (!date || hours === null) throw new Error(`核可表第 ${sourceRow} 列日期或 F 欄時數無效。`);
      const start = combineDateTime(date, row[6]);
      const end = normalizeEndTime(start, combineDateTime(date, row[7]));
      if (!start || !end || end <= start) throw new Error(`核可表第 ${sourceRow} 列加班起訖無效。`);
      const bandHours = row.slice(10, 15).map((value) => cleanText(value) === "" ? null : this.number(value));
      if (row.slice(10, 15).some((value, i) => cleanText(value) !== "" && bandHours[i] === null)) {
        throw new Error(`核可表第 ${sourceRow} 列倍率時數須為非負數。`);
      }
      if (hours === 0) {
        if (bandHours.some((value) => value > 0)) throw new Error(`核可表第 ${sourceRow} 列 F 欄為零但有核給時數，請覆核。`);
        return;
      }
      const signIns = parseSeriesDateTimes(date, row[8]);
      const signOuts = parseSeriesDateTimes(date, row[9]).map((time) => normalizeEndTime(start, time));
      rows.push({
        sourceSheet: name, sourceRow, date, dateKey: formatDateKey(date), month: formatMonth(date), person, sourcePerson,
        department: cleanText(row[2]), shift: cleanText(row[4]), reason: "", hours, category: "", audit: "", hourly: 0, amount: 0,
        start, end, startRaw: cleanText(row[6]), endRaw: cleanText(row[7]), normalAttendanceText: "",
        signInRaw: cleanText(row[8]), signOutRaw: cleanText(row[9]), signIns, signOuts,
        carInRaw: "", carOutRaw: "", helperCarIns: [], helperCarOuts: [], bandHours,
        grantedHours: sum(bandHours.map((value) => value || 0)), approvalNotes: [],
      });
    });
    const deduped = dedupeApplicationRows(rows);
    deduped.forEach((row) => {
      const duplicates = rows.filter((item) => row.sourceRows.includes(item.sourceRow));
      if (duplicates.some((item) => item.hours !== duplicates[0].hours || JSON.stringify(item.bandHours) !== JSON.stringify(duplicates[0].bandHours))) {
        row.approvalNotes.push("重複核可單的時數或倍率分段不同，暫採第一列，請覆核。");
      }
      row.hours = duplicates[0].hours;
    });
    return deduped;
  },

  number(value) {
    if (typeof value === "number") return Number.isFinite(value) && value >= 0 ? value : null;
    const text = cleanText(value).replace(/,/g, "");
    return /^\d+(?:\.\d+)?$/.test(text) ? Number(text) : null;
  },

  calculate(grant, hourly, type) {
    const bands = grant.bandHours.map((hours) => hours || 0);
    const grantedHours = sum(bands);
    const notes = [...(grant.approvalNotes || [])];
    if (grantedHours > grant.hours + 0.001) notes.push("核給時數大於申請時數，保留原資料，請覆核。");
    if (grantedHours > 0 && hourly <= 0) notes.push("缺少本人當月時薪，加班費待補。");
    const segments = bands.map((hours, index) => {
      const paidHours = hours;
      return { label: this.labels[index], hours, paidHours, multiplier: this.rates[index], taxable: index === 4, amount: Math.ceil(paidHours * hourly * this.rates[index]) };
    }).filter((segment) => segment.hours > 0);
    return {
      segments, total: grantedHours === 0 ? 0 : hourly > 0 ? sum(segments.map((segment) => segment.amount)) : null,
      paySource: "核可 K～O 倍率時數 × 本人當月時薪",
      approvalNotes: notes,
      note: type === DAY_TYPE.HOLIDAY && grantedHours > 0 ? "另有一日補休。" : "",
    };
  },

  segmentText(segments) {
    return segments.map((segment) => `${segment.label} ${formatNumber(segment.paidHours, 2)}H`).join(" + ");
  },
};
