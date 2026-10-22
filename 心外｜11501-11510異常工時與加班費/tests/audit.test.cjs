"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const root = path.resolve(__dirname, "..");
const sources = ["vendor/xlsx.full.min.js", "person-settings.js", "payroll.js", "approval.js", "surgery.js", "app.js"].map((name) => fs.readFileSync(path.join(root, name), "utf8"));

function context() {
  const ctx = vm.createContext({ console, document: { body: { dataset: { mode: "approval" } }, addEventListener() {}, getElementById() { return null; } } });
  sources.forEach((source) => vm.runInContext(source, ctx));
  vm.runInContext(`setupControls=()=>{};render=()=>{};setStatus=(message,warning)=>{globalThis.statusResult={message,warning};};
    downloadCsv=(rows)=>{globalThis.csvResult=rows;};dom.monthSelect={};dom.personSelect={};dom.riskSelect={};`, ctx);
  return ctx;
}
const run = (ctx, code) => vm.runInContext(code, ctx);
const json = (ctx, code) => JSON.parse(run(ctx, `JSON.stringify(${code})`));

function fixture(ctx, { requests = [], grants = [], wage = 100 } = {}) {
  ctx.input = { requests, grants, wage };
  run(ctx, `var book=XLSX.utils.book_new();
    var add=(name,rows)=>XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(rows),name);
    add('排班',[['電子簽章','日期','班別','班別時數','簽到','簽退','應簽到','應簽退','其他簽到','其他簽退','加班時數'],
      ['Test A','2026/09/15','D8',8,'06:00:03','23:59:50','08:00','16:30','','',999]]);
    add('未核可｜加班單申請',[['日期','部門','出勤者','班別','申請時數','事由','起始時間','終止時間','正常簽到退','簽到','簽退'],
      ...input.requests.map(r=>[r.date||'2026/09/15','','Test A','D8',r.hours,r.reason||'Test',r.start||'16:30',r.end||'20:30','','',''])]);
    add('核可｜加班單申請',[['日期','出勤者','部門','計薪年月','班別','加班總時數','加班起始時間','加班終止時間','簽到','簽退','加班費倍率','','','',''],
      ['','','','','','','','','','',1,1.334,1.667,2.667,'1.667(應稅)'],
      ...input.grants.map(r=>[r.date||'2026/09/15',r.person||'Test A','','2026/09',r.shift||'D8',r.hours,r.start||'16:30',r.end||'20:30',r.signIn||'',r.signOut||'',...(r.bands||[0,2,2,0,''])])]);
    add('時薪',[['姓名','年月','時薪'],['Test A','2026/09',input.wage]]);
    loadWorkbook(book,'fixture.xlsx');`);
}

function surgeryFixture(ctx, ranges) {
  ctx.ranges = ranges;
  run(ctx, `var sb=XLSX.utils.book_new();var header=Array(21).fill('');header[19]='執行起日';header[20]='執行迄日';
    XLSX.utils.book_append_sheet(sb,XLSX.utils.aoa_to_sheet([header,...ranges.map((r,i)=>{var row=Array(21).fill('');row[5]=r[2]||'Case A';row[6]='Patient';row[19]=r[0];row[20]=r[1];return row;})]),'Procedures');
    surgeryAudit.loadWorkbook(sb,'procedure.xlsx');`);
}

let count = 0;
function check(name, fn) { fn(); count++; console.log(`PASS ${name}`); }

check("entry shows its release and versions every local script and stylesheet", () => {
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const version = html.match(/<meta name="app-version" content="([\d.]+)">/)?.[1];
  assert.ok(version);
  assert.ok(html.includes(`本機處理 · v${version}`));
  const resources = [...html.matchAll(/(?:src|href)="([^"]+\.(?:js|css)(?:\?[^"]*)?)"/g)].map(match => match[1]);
  assert.equal(resources.length, 8);
  for (const resource of resources) {
    const url = new URL(resource, 'file:///app/index.html');
    assert.equal(url.searchParams.get('v'), version);
    const name = resource.split('?')[0];
    if (name !== 'private-settings.js') assert.ok(fs.existsSync(path.join(root, name)));
  }
});

check("F is requested; K:O is granted; no schedule AH totals", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 9 }] });
  assert.deepEqual(json(ctx, 'state.overtimeRecords.map(r=>[r.hours,r.approvedHours,r.unapprovedHours,r.estimatedPay])'), [[9, 4, 5, 601]]);
  assert.equal(run(ctx, 'state.scheduleRows.every(r=>r.overtimeHours===0)'), true);
  assert.equal(run(ctx, 'approvalView.monthlyApproved.get("Test A|2026/09")'), 4);
});

check("all five multiplier columns including taxable O are priced independently", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 16, bands: [1, 2, 3, 4, 6] }] });
  assert.equal(run(ctx, 'state.overtimeRecords[0].estimatedPay'), [100, 267, 501, 1067, 1001].reduce((a,b)=>a+b,0));
  assert.equal(run(ctx, 'payroll.segmentText(state.overtimeRecords[0].approvals[0].segments).includes("1.667倍(含稅) 6H")'), true);
});

check("M is a multiplier hour column, never imported as currency", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 2, bands: [0, 0, 2, 0, ""] }] });
  assert.equal(run(ctx, 'state.overtimeRecords[0].estimatedPay'), 334);
});

check("zero or blank K:O means no grant, not a missing wage or auto-filled grant", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 8, bands: [0, 0, 0, 0, ""] }], wage: "" });
  assert.deepEqual(json(ctx, 'state.overtimeRecords.map(r=>[r.approvedHours,r.unapprovedHours,r.estimatedPay,r.payMissing])'), [[0, 8, 0, false]]);
});

check("unapproved E plus nonduplicated approved F form all requests", () => {
  const ctx = context(); fixture(ctx, { requests: [{ hours: 9, end: "2026/09/16 01:30" }], grants: [{ hours: 9 }] });
  assert.deepEqual(json(ctx, 'state.overtimeRecords.map(r=>[r.hours,r.approvedHours,r.unapprovedHours])'), [[9, 4, 5]]);
});

check("duplicates are counted once and conflicting duplicate bands require review", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }, { hours: 4 }] });
  assert.equal(run(ctx, 'sum(state.overtimeRecords.map(r=>r.approvedHours))'), 4);
  fixture(ctx, { grants: [{ hours: 4 }, { hours: 4, bands: [0, 1, 3, 0, ""] }] });
  assert.equal(run(ctx, 'state.overtimeRecords[0].approvalNotes.some(n=>n.includes("重複"))'), true);
});

check("missing wages are not taken from another person or month", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4, person: "Test B" }] });
  assert.equal(run(ctx, 'state.overtimeRecords[0].payMissing'), true);
  assert.equal(run(ctx, 'approvalView.resolveHourly("Test A大林","2026/09",buildWageMap(state.wageRows))'), 100);
  assert.equal(run(ctx, 'approvalView.resolveHourly("Test A","2026/08",buildWageMap(state.wageRows))'), 0);
});

check("fixed and regular holidays use recorded 1x hours", () => {
  const ctx = context();
  fixture(ctx, { grants: [{ hours: 8, shift: "國休(8)", bands: [8, 0, 0, 0, ""] }], wage: 175 });
  assert.equal(run(ctx, 'state.overtimeRecords[0].estimatedPay'), 1400);
  fixture(ctx, { grants: [{ hours: 6, shift: "OFF(例假)", bands: [6, 0, 0, 0, ""] }], wage: 175 });
  assert.equal(run(ctx, 'state.overtimeRecords[0].estimatedPay'), 1050);
  assert.equal(run(ctx, 'state.overtimeRecords[0].approvals[0].payNote.includes("補休")'), false); // Schedule shift wins.
  assert.equal(run(ctx, 'payroll.calculate({hours:6,bandHours:[6,0,0,0,0]},175,DAY_TYPE.HOLIDAY).note.includes("補休")'), true);
});

check("monthly grants are K:O, never capped or filled from F", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 60, bands: [0, 2, 2, 47, ""] }] });
  assert.equal(run(ctx, 'approvalView.monthlyApproved.get("Test A|2026/09")'), 51);
  assert.equal(run(ctx, 'state.overtimeRecords[0].unapprovedHours'), 9);
  assert.equal(run(ctx, 'approvalView.calendarWorkRisks(state.dailyRows[0]).includes(WORK_RISK.OVER_46_HOURS)'), false);
  assert.equal(run(ctx, 'renderRiskSummaryCell(state.dailyRows[0]).includes(WORK_RISK.OVER_46_HOURS)'), true);
});

check("wrong F or K:O headers leave the existing import intact", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }] });
  run(ctx, 'book.Sheets["核可｜加班單申請"].F1.v="Wrong";loadWorkbook(book,"wrong.xlsx")');
  assert.equal(run(ctx, 'statusResult.warning'), true);
  assert.equal(run(ctx, 'state.workbookName'), "fixture.xlsx");
});

check("ROC timestamps, seconds, midnight and invalid dates are parsed strictly", () => {
  const ctx = context();
  assert.equal(run(ctx, 'formatDateTime(surgeryAudit.parseDate("1150915163001"))'), "2026-09-15 16:30:01");
  assert.equal(run(ctx, 'formatDateTime(surgeryAudit.parseDate("11509151630"))'), "2026-09-15 16:30:00");
  assert.equal(run(ctx, 'formatDateTime(surgeryAudit.parseDate("115/09/16 00:00:02"))'), "2026-09-16 00:00:02");
  assert.equal(run(ctx, 'surgeryAudit.parseDate("11502301630")'), null);
});

check("T/U time comparison does not require a staff or patient match", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }] });
  surgeryFixture(ctx, [["11509151630", "11509152030", "Unrelated case"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), "手術起訖相符");
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.risk'), false);
});

check("partial overlap, no overlap, and point-only evidence are distinct", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }] });
  surgeryFixture(ctx, [["11509151700", "11509151900"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), "手術部分重疊");
  surgeryFixture(ctx, [["11509150800", "11509151000"], ["11509160800", "11509161000"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), "無對應手術時段");
  surgeryFixture(ctx, [["11509151800", "11509151800"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), "手術僅時間點");
});

check("cross-midnight procedures and gaps between cases are respected", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 8, start: "20:30", end: "2026/09/16 04:30" }] });
  surgeryFixture(ctx, [["11509152030", "11509160430"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), "手術起訖相符");
  surgeryFixture(ctx, [["11509152030", "11509152300"], ["11509160100", "11509160430"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.risk'), true);
});

check("duplicates do not inflate surgery duration, invalid data is counted", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }] });
  surgeryFixture(ctx, [["11509151630", "11509152030"], ["11509151630", "11509152030"], ["bad", "bad"]]);
  assert.equal(run(ctx, 'surgeryAudit.rows.length'), 1);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.overlapHours'), 4);
  assert.equal(run(ctx, 'surgeryAudit.parse(sb).invalid'), 1);
});

check("either import order works and replacing data recomputes results", () => {
  const ctx = context(); surgeryFixture(ctx, [["11509151630", "11509152030"]]);
  fixture(ctx, { grants: [{ hours: 4 }] });
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), "手術起訖相符");
  surgeryFixture(ctx, [["11509151800", "11509152000"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), "手術部分重疊");
});

check("month/day summaries, filter and exports share the computed result", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 9 }] });
  surgeryFixture(ctx, [["11509151800", "11509152000"]]);
  run(ctx, 'state.selectedRisk=surgeryAudit.issueFilter');
  assert.equal(run(ctx, 'getFilteredRecords().length'), 1);
  assert.equal(run(ctx, 'dayCellHtml(state.dailyRows[0].date,state.dailyRows).includes("與加班起訖時間不符")'), true);
  run(ctx, 'exportDetailCsv()');
  assert.equal(run(ctx, 'csvResult.every(r=>r.length===csvResult[0].length)'), true);
  assert.equal(run(ctx, 'csvResult[1][18]'), "與加班起訖時間不符");
  assert.equal(run(ctx, 'csvResult[1][19].includes("第2列")'), true);
});

check("actual attendance remains separate and 30-minute note is display-only", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }] });
  const before = run(ctx, 'JSON.stringify(state.overtimeRecords.map(r=>[r.start,r.end,r.attendanceStatuses,r.estimatedPay]))');
  surgeryFixture(ctx, [["11509151800", "11509152000"]]);
  assert.equal(run(ctx, 'JSON.stringify(state.overtimeRecords.map(r=>[r.start,r.end,r.attendanceStatuses,r.estimatedPay]))'), before);
  assert.equal(run(ctx, 'approvalView.lateSignInLabel'), "簽到晚，起始時間往前推估半小時");
});

check("March 15 continuous scheduled attendance is normal and keeps second precision", () => {
  const ctx = context();
  run(ctx, `var a={start:new Date(2026,2,15,7),end:new Date(2026,2,15,21)};
    var e=[{time:new Date(2026,2,15,6,33,50),direction:'entry',sourceGroup:'schedule',sourceLabel:'排班簽到'},
      {time:new Date(2026,2,15,21,38,32),direction:'exit',sourceGroup:'schedule',sourceLabel:'排班簽退'}];
    var result=evaluateAttendanceRecord(a,e);`);
  assert.deepEqual(json(ctx, 'result.statuses'), []);
  assert.equal(run(ctx, 'formatDateTime(result.exit.time)'), '2026-03-15 21:38:32');
});

check("car evidence takes priority, while missing car falls back to schedule", () => {
  const ctx = context();
  run(ctx, `var a={start:new Date(2026,3,16,18,30),end:new Date(2026,3,17,0)};
    var e=[{time:new Date(2026,3,16,19,42),direction:'entry',sourceGroup:'car',sourceLabel:'汽車進場'},
      {time:new Date(2026,3,16,20,59),direction:'exit',sourceGroup:'car',sourceLabel:'汽車離場'},
      {time:new Date(2026,3,16,17),direction:'entry',sourceGroup:'schedule',sourceLabel:'排班簽到'},
      {time:new Date(2026,3,17,1),direction:'exit',sourceGroup:'schedule',sourceLabel:'排班簽退'}];
    var result=evaluateAttendanceRecord(a,e);`);
  assert.equal(run(ctx, 'result.sections[0].entrySource'), 'car');
  assert.equal(run(ctx, 'result.statuses.includes(ATTENDANCE_STATUS.EARLY_LEAVE.short)'), true);
  assert.equal(run(ctx, 'evaluateAttendanceRecord(a,e.filter(r=>r.sourceGroup==="schedule")).sections[0].entrySource'), 'schedule');
});

check("same record has one time header in the risk table with grant and surgery notes", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 9 }] });
  surgeryFixture(ctx, [["11509151800", "11509152000"]]);
  assert.equal(run(ctx, '(renderRiskSummaryCell(state.dailyRows[0]).match(/●加班起訖/g)||[]).length'), 1);
  assert.equal(run(ctx, 'renderRiskSummaryCell(state.dailyRows[0]).includes("核給 4 h／未核給 5 h")'), true);
});

check("small gaps use tolerance without inflating actual overlap hours", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }] });
  surgeryFixture(ctx, [["11509151630", "11509151800"], ["11509151805", "11509152030"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.risk'), false);
  assert.ok(Math.abs(run(ctx, 'state.overtimeRecords[0].procedureAudit.overlapHours') - 235/60) < 0.0001);
});

check("time-only evidence outside source coverage is not a fabricated attendance violation", () => {
  const ctx = context(); fixture(ctx, { grants: [{ hours: 4 }] });
  surgeryFixture(ctx, [["11501010800", "11501011000"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.label'), '超出手術資料期間');
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.risk'), false);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.pending'), true);
});

check("confirmed aliases join requests, grants, attendance and the people selector", () => {
  const ctx = context();
  run(ctx, 'globalThis.HR_LOCAL_SETTINGS={personAliases:{"Test A Return":"Test A"},hourlyOverrides:[]}');
  fixture(ctx, { requests: [{ hours: 9 }], grants: [{ hours: 9, person: "Test A Return" }] });
  assert.deepEqual(json(ctx, 'state.people'), ['Test A']);
  assert.deepEqual(json(ctx, 'state.overtimeRecords.map(r=>[r.hours,r.approvedHours,r.unapprovedHours,r.person])'), [[9,4,5,'Test A']]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].approvals[0].sourcePerson'), 'Test A Return');
  assert.equal(run(ctx, 'state.overtimeRecords[0].attendanceStatuses.length'), 0);
  assert.equal(run(ctx, 'approvalView.monthlyApproved.get("Test A|2026/09")'), 4);
  assert.equal(run(ctx, 'personSettings.canonicalName("Test B Return")'), 'Test B Return');
});

check("confirmed hourly values apply to the specified person and month only", () => {
  const ctx = context();
  run(ctx, 'globalThis.HR_LOCAL_SETTINGS={personAliases:{"Test A Return":"Test A"},hourlyOverrides:[{person:"Test A",month:"2026/09",hourly:123}]}');
  fixture(ctx, { grants: [{ hours:4, person:'Test A Return' }], wage:0 });
  assert.equal(run(ctx,'state.overtimeRecords[0].hourly'),123);
  assert.equal(run(ctx,'state.overtimeRecords[0].payMissing'),false);
  assert.equal(run(ctx,'state.overtimeRecords[0].estimatedPay'),Math.ceil(2*123*1.334)+Math.ceil(2*123*1.667));
  assert.equal(run(ctx,'approvalView.resolveHourly("Test A Return","2026/10",new Map())'),0);
  assert.equal(run(ctx,'approvalView.resolveHourly("Test B","2026/09",new Map())'),0);
  assert.equal(run(ctx,'approvalView.resolveHourly("Test A","2026/09",new Map([["Test A|2026/09",99]]))'),123);
});

check("missing clocks show tags only and absent surgery does not add long list text", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:4}] });
  surgeryFixture(ctx, [["11509140800","11509140900"],["11509160800","11509160900"]]);
  run(ctx, 'var record=state.overtimeRecords[0];record.attendanceStatuses=[ATTENDANCE_STATUS.MISSING_SIGN_IN.short,ATTENDANCE_STATUS.MISSING_SIGN_OUT.short]');
  assert.deepEqual(json(ctx,'surgeryAudit.summaryItems(record)'),[]);
  assert.equal(run(ctx,'buildRiskSummaryAttendanceBlocks([record])[0].items[0].label'),'簽到、簽退無紀錄');
  assert.equal(run(ctx,'buildRiskSummaryAttendanceBlocks([record])[0].items[0].text'),'');
  assert.equal(run(ctx,'surgeryAudit.detailHtml(record).includes("未找到與加班起訖相符之執行時段，須補其他佐證。")'),true);
});

check("partial surgery shows exact execution times with mismatch tag in all views", () => {
  const ctx = context();fixture(ctx,{grants:[{hours:7.5,start:'16:30',end:'23:59'}]});
  surgeryFixture(ctx,[["11509152046","11509152145"]]);
  assert.equal(run(ctx,'surgeryAudit.items(state.overtimeRecords[0])[0].label'),'與加班起訖時間不符');
  const table=run(ctx,'renderRiskSummaryCell(state.dailyRows[0])');
  for(const expected of ['20:46:00-21:45:00','與加班起訖時間不符'])assert.ok(table.includes(expected));
  assert.ok(!table.includes('其餘申請區間需補'));
  const detail=run(ctx,'surgeryAudit.detailHtml(state.overtimeRecords[0])');
  assert.ok(detail.includes('2026-09-15 20:46:00-2026-09-15 21:45:00'));
  assert.ok(!detail.includes('手術部分重疊'));
});

check("point-only and partial surgery use the same mismatch style without changing risk counts", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:4}] });
  surgeryFixture(ctx, [["11509151800","11509151800"]]);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.risk'), false);
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.pending'), true);
  assert.deepEqual(json(ctx, 'surgeryAudit.items(state.overtimeRecords[0]).map(i=>[i.label,i.status,i.tone])'),
    [['與加班起訖時間不符','procedure-risk','procedure-risk']]);
  for (const code of ['dayCellHtml(state.dailyRows[0].date,state.dailyRows)', 'renderRiskSummaryCell(state.dailyRows[0])', 'surgeryAudit.detailHtml(state.overtimeRecords[0])']) {
    const html = run(ctx, code);
    assert.match(html, /procedure-risk[^>]*>與加班起訖時間不符/);
    assert.ok(!html.includes('procedure-pending'));
    assert.ok(html.includes('18:00:00'));
  }
  surgeryFixture(ctx, [["11509151800","11509151900"]]);
  assert.equal(run(ctx, 'surgeryAudit.items(state.overtimeRecords[0])[0].tone'), 'procedure-risk');
  assert.equal(run(ctx, 'state.overtimeRecords[0].procedureAudit.risk'), true);
});

check("no overlap lists the nearest same-day point as reference, including when clocks are missing", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:2.5,start:'05:38',end:'08:00'}] });
  surgeryFixture(ctx, [["11509140800","11509140900"], ["11509150836","11509150836"],
    ["11509150900","11509151350"], ["11509151613","11509151613"]]);
  run(ctx, 'var record=state.overtimeRecords[0];record.attendanceStatuses=[ATTENDANCE_STATUS.MISSING_SIGN_IN.short,ATTENDANCE_STATUS.MISSING_SIGN_OUT.short]');
  assert.equal(run(ctx, 'record.procedureAudit.evidence.length'), 0);
  assert.equal(run(ctx, 'record.procedureAudit.risk'), true);
  assert.equal(run(ctx, 'record.procedureAudit.label'), '無對應手術時段');
  assert.deepEqual(json(ctx, 'surgeryAudit.timeLabels(record)'), ['當日最接近手術時間點 08:36:00']);
  for (const code of ['dayCellHtml(state.dailyRows[0].date,state.dailyRows)', 'renderRiskSummaryCell(state.dailyRows[0])', 'surgeryAudit.detailHtml(record)']) {
    const html = run(ctx, code);
    assert.ok(html.includes('當日最接近手術時間點'));
    assert.ok(html.includes('08:36:00'));
    assert.ok(html.includes('無對應手術時段'));
    assert.ok(!html.includes('13:50:00'));
  }
  assert.ok(run(ctx,'surgeryAudit.exportEvidence(record)').includes('當日最接近（未重疊）：2026-09-15 08:36:00'));
});

check("nearest same-day ranges use interval distance, keep ties and do not borrow another date", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:4,start:'16:30',end:'20:30'}] });
  surgeryFixture(ctx, [["11509140900","11509141000"], ["11509151400","11509151600"],
    ["11509152100","11509152300"], ["11509150800","11509151000"], ["11509160800","11509161000"]]);
  assert.deepEqual(json(ctx, 'surgeryAudit.timeLabels(state.overtimeRecords[0])'),
    ['當日最接近手術 14:00:00-16:00:00','當日最接近手術 21:00:00-23:00:00']);
  surgeryFixture(ctx, [["11509140900","11509141000"], ["11509160800","11509161000"]]);
  assert.deepEqual(json(ctx, 'state.overtimeRecords[0].procedureAudit.nearbyEvidence'), []);
  assert.deepEqual(json(ctx, 'surgeryAudit.timeLabels(state.overtimeRecords[0])'), []);
});

check("same-day nearest references handle midnight and never replace overlapping evidence", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:4,start:'16:30',end:'20:30'}] });
  surgeryFixture(ctx, [["11509142300","11509150000"], ["11509152330","11509160030"]]);
  assert.deepEqual(json(ctx, 'surgeryAudit.timeLabels(state.overtimeRecords[0])'), ['當日最接近手術 23:30:00-09/16 00:30:00']);
  surgeryFixture(ctx, [["11509151400","11509151600"], ["11509151800","11509151900"]]);
  assert.deepEqual(json(ctx, 'state.overtimeRecords[0].procedureAudit.nearbyEvidence'), []);
  assert.deepEqual(json(ctx, 'surgeryAudit.timeLabels(state.overtimeRecords[0])'), ['手術執行 18:00:00-19:00:00']);
});

check("matching and covering procedures have no UI badge but retain detail times and export results", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:2.5,start:'16:30',end:'19:00'}] });
  for (const ranges of [[["11509151355","11509151900"]], [["11509151630","11509151900"]]]) {
    surgeryFixture(ctx, ranges);
    const record = 'state.overtimeRecords[0]';
    assert.deepEqual(json(ctx, `surgeryAudit.items(${record})`), []);
    assert.deepEqual(json(ctx, `surgeryAudit.summaryItems(${record})`), []);
    for (const code of ['dayCellHtml(state.dailyRows[0].date,state.dailyRows)', 'renderRiskSummaryCell(state.dailyRows[0])', `surgeryAudit.detailHtml(${record})`]) {
      const html = run(ctx, code);
      assert.ok(!html.includes('手術時段涵蓋'));
      assert.ok(!html.includes('手術起訖相符'));
      assert.ok(!html.includes('procedure-ok'));
    }
    assert.ok(run(ctx, `surgeryAudit.detailHtml(${record})`).includes('2026-09-15 19:00:00'));
    assert.equal(run(ctx, `${record}.procedureAudit.risk || ${record}.procedureAudit.pending`), false);
    assert.ok(['手術時段涵蓋','手術起訖相符'].includes(run(ctx, `surgeryAudit.displayLabel(${record}.procedureAudit)`)));
  }
});

check("risk list hides generic unapproved notices while retaining hours, filters and audit notes", () => {
  const ctx = context(); fixture(ctx, { requests:[{hours:2.5,start:'16:30',end:'19:00'}] });
  run(ctx, 'var record=state.overtimeRecords[0]');
  assert.equal(run(ctx, 'record.approvedHours'), 0);
  assert.equal(run(ctx, 'record.unapprovedHours'), 2.5);
  assert.deepEqual(json(ctx, 'approvalView.riskSummaryItems(record)'), []);
  const table = run(ctx, 'renderRiskSummaryCell(state.dailyRows[0])');
  for (const text of ['有未核給時數','未核給原因待確認','核給 0 h／未核給 2.5 h']) assert.ok(!table.includes(text));
  assert.equal(run(ctx, 'record.approvalNotes.some(n=>n.includes("未核給原因待確認"))'), true);
  run(ctx, 'state.selectedRisk="unapproved"');
  assert.equal(run(ctx, 'getFilteredRecords().length'), 1);
  run(ctx, 'record.approvalNotes.push("重複核可單的時數或倍率分段不同，暫採第一列，請覆核。")');
  assert.equal(run(ctx, 'approvalView.riskSummaryItems(record)[0].label'), '核給資料待覆核');
  assert.ok(!run(ctx, 'approvalView.riskSummaryItems(record)[0].text').includes('未核給原因待確認'));
  assert.ok(run(ctx, 'approvalView.riskSummaryItems(record)[0].text').includes('重複核可單'));
});

check("specific limit and missing-pay alerts remain in the risk list", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:9}] });
  assert.equal(run(ctx, 'approvalView.riskSummaryItems(state.overtimeRecords[0])[0].label'), '超出日/月延展工時無核給');
  fixture(ctx, { grants:[{hours:4}], wage:0 });
  assert.equal(run(ctx, 'approvalView.riskSummaryItems(state.overtimeRecords[0])[0].label'), '加班費待補');
});

check("October 1 request keeps surgery and hours without generic unapproved table text", () => {
  const ctx = context(); fixture(ctx, { requests:[{date:'2026/10/01',hours:4,start:'16:30',end:'20:30'}] });
  surgeryFixture(ctx, [["11510011625","11510011645"]]);
  run(ctx, 'var daily=state.dailyRows.find(d=>d.dateKey==="2026-10-01");var record=daily.overtimeRecords[0]');
  const html = run(ctx, 'renderRiskSummaryCell(daily)');
  for (const text of ['有未核給時數','未核給原因待確認','請依退單及簽核紀錄確認','核給 0 h／未核給 4 h','手術時段涵蓋']) assert.ok(!html.includes(text));
  for (const text of ['16:30-20:30','16:25:00-16:45:00','與加班起訖時間不符']) assert.ok(html.includes(text));
  assert.deepEqual(json(ctx,'[record.hours,record.approvedHours,record.unapprovedHours]'),[4,0,4]);
});

check("September 29 coverage retains missing clocks and actual procedure without normal badge", () => {
  const ctx = context(); fixture(ctx, { requests:[{date:'2026/09/29',hours:2.5,start:'16:30',end:'19:00'}] });
  surgeryFixture(ctx, [["11509291355","11509291900"]]);
  run(ctx, 'var daily=state.dailyRows.find(d=>d.dateKey==="2026-09-29");var record=daily.overtimeRecords[0]');
  assert.equal(run(ctx, 'record.procedureAudit.label'), '手術時段涵蓋');
  for (const code of ['dayCellHtml(daily.date,[daily])','renderRiskSummaryCell(daily)','surgeryAudit.detailHtml(record)']) {
    const html = run(ctx, code);
    assert.ok(!html.includes('手術時段涵蓋'));
    assert.ok(!html.includes('手術起訖相符'));
  }
  assert.ok(run(ctx, 'renderRiskSummaryCell(daily)').includes('簽到、簽退無紀錄'));
  assert.ok(run(ctx, 'surgeryAudit.detailHtml(record)').includes('2026-09-29 13:55:00-2026-09-29 19:00:00'));
});

function detailSummarySections(ctx, date) {
  ctx.selectedDate = date;
  run(ctx, 'dom.detailTitle={};dom.detailSubtitle={};dom.detailBody={};state.selectedDateKey=selectedDate;renderDetail(state.dailyRows)');
  return [...run(ctx, 'dom.detailBody.innerHTML').matchAll(/<section class="detail-section">([\s\S]*?)<\/section>/g)].map(match => match[1]);
}

check("normal attendance and no work risks show one badge without duplicate notes", () => {
  const ctx = context(); fixture(ctx, { grants:[{hours:4}] });
  const [attendance, work] = detailSummarySections(ctx, '2026-09-15');
  assert.equal((attendance.match(/>正常</g) || []).length, 1);
  assert.equal((work.match(/>無</g) || []).length, 1);
  for (const section of [attendance, work]) {
    assert.ok(!section.includes('detail-notes'));
    assert.ok(!section.includes('<li>'));
    assert.ok(!section.includes('empty-note'));
  }
});

check("real attendance and work risk explanations remain in the right detail", () => {
  const ctx = context(); fixture(ctx, { grants:[{date:'2026/09/16',hours:14}] });
  const [attendance, work] = detailSummarySections(ctx, '2026-09-16');
  assert.ok(attendance.includes('簽到、簽退無紀錄'));
  assert.ok(attendance.includes('detail-notes'));
  assert.ok(attendance.includes('未找到'));
  assert.ok(work.includes('日總工時12小時以上'));
  assert.ok(work.includes('detail-notes'));
  assert.ok(work.includes('14'));
});

check("a schedule-only day does not get a contradictory normal attendance note", () => {
  const ctx = context(); fixture(ctx);
  const [attendance, work] = detailSummarySections(ctx, '2026-09-15');
  assert.ok(attendance.includes('無加班申請'));
  assert.ok(!attendance.includes('正常'));
  assert.ok(!attendance.includes('detail-notes'));
  assert.ok(!work.includes('detail-notes'));
});

if (process.argv[2] && process.argv[3]) {
  check("supplied F and all K:O hours reconcile without sensitive output", () => {
    const ctx = context(); ctx.hrBytes = fs.readFileSync(process.argv[2]); ctx.surgeryBytes = fs.readFileSync(process.argv[3]); ctx.hrName = path.basename(process.argv[2]);
    const localSettings = path.join(root, 'private-settings.js');
    if (fs.existsSync(localSettings)) run(ctx, fs.readFileSync(localSettings, 'utf8'));
    run(ctx, 'var actualBook=XLSX.read(hrBytes,{type:"buffer",cellDates:true});loadWorkbook(actualBook,hrName);surgeryAudit.loadWorkbook(XLSX.read(surgeryBytes,{type:"buffer",cellDates:true}),"actual-surgery.xlsx")');
    const result = json(ctx, `(()=>{const rows=payroll.parseApproved(actualBook,'核可｜加班單申請');const raw=XLSX.utils.sheet_to_json(actualBook.Sheets['核可｜加班單申請'],{header:1,raw:true,defval:''}).slice(2).filter(r=>r[1]&&Number(r[5])>0);const wages=buildWageMap(state.wageRows);return {sourceHours:raw.reduce((s,r)=>s+r.slice(10,15).reduce((a,v)=>a+(Number(v)||0),0),0),resultHours:sum(state.overtimeRecords.map(r=>r.approvedHours)),sourceCount:rows.length,resultCount:state.overtimeRecords.flatMap(r=>r.approvals).length,allAudited:state.overtimeRecords.every(r=>r.procedureAudit),sourcePay:sum(rows.map(g=>{const wage=approvalView.resolveHourly(g.person,g.month,wages);return g.bandHours.reduce((s,h,i)=>s+Math.ceil((h||0)*wage*[1,1.334,1.667,2.667,1.667][i]),0);})),resultPay:sum(state.overtimeRecords.map(r=>r.estimatedPay)),pending:state.overtimeRecords.filter(r=>r.payMissing).length};})()`);
    assert.equal(result.sourceHours, result.resultHours);
    assert.equal(result.sourceCount, result.resultCount);
    assert.equal(result.sourcePay, result.resultPay);
    assert.equal(result.allAudited, true);
    console.log(JSON.stringify(result));
  });
}
console.log(`${count} checks passed`);
