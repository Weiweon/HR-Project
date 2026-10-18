"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { execFileSync } = require("node:child_process");
const root = path.resolve(__dirname, "..");
const vendor = fs.readFileSync(path.join(root, "vendor/xlsx.full.min.js"), "utf8");
const app = fs.readFileSync(path.join(root, "app.js"), "utf8");
const approval = fs.readFileSync(path.join(root, "approval.js"), "utf8");

function context(mode = "approval", source = app) {
  const ctx = vm.createContext({ console, document: { body: { dataset: { mode } }, addEventListener() {} } });
  vm.runInContext(vendor, ctx);
  if (mode === "approval") vm.runInContext(approval, ctx);
  vm.runInContext(source, ctx);
  vm.runInContext(`setupControls = () => {}; render = () => {};
    setStatus = (message, warning) => { globalThis.statusResult = {message, warning}; };
    downloadCsv = (rows) => { globalThis.csvResult = rows; };`, ctx);
  return ctx;
}

function evaluate(ctx, source) {
  return vm.runInContext(source, ctx);
}

function fixture(ctx, { requests = [], grants = [], wage = 100, scheduleOvertime = 999 } = {}) {
  ctx.testData = { requests, grants, wage, scheduleOvertime };
  evaluate(ctx, `
    var workbook = XLSX.utils.book_new();
    var addSheet = (name, rows) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
    addSheet("排班", [["電子簽章","日期","班別","班別時數","簽到","簽退","應簽到","應簽退","其他簽到","其他簽退","加班時數"],
      ["Test A","2026/09/15","D8",8,"06:00:03","23:59:50","08:00","16:30","","",testData.scheduleOvertime],
      ["Test A","2026/09/01","D8",8,"06:00:03","23:59:50","08:00","16:30","","",testData.scheduleOvertime]]);
    addSheet("未核可｜加班單申請", [["日期","部門","出勤者","班別","申請時數","事由","起始時間","終止時間","正常簽到退","簽到","簽退"],
      ...testData.requests.map(r => [r.date||"2026/09/15","","Test A","D8",r.hours,r.reason||"Test",r.start||"16:30",r.end||"20:30","","",""])]);
    addSheet("核可｜加班單申請", [["日期","出勤者","部門","類別","事由","加班費倍率","","","","","班別","申請時數","金額","起始時間","終止時間"],
      ["","","","","",1,1.334,1.667,2.667,"1.667(應稅)","","","","",""],
      ...testData.grants.map(r=>[r.date||"2026/09/15",r.person||"Test A","","平日",r.reason||"Test",0,0,0,0,0,"D8",r.hours,r.amount??"",r.start||"16:30",r.end||"20:30"])]);
    addSheet("時薪", [["姓名","年月","時薪"],["Test A","2026/09",testData.wage]]);
    loadWorkbook(workbook,"test.xlsx");
  `);
}

function snapshot(ctx) {
  return JSON.parse(evaluate(ctx, `JSON.stringify(state.overtimeRecords.map(r=>({hours:r.hours,approved:r.approvedHours,unapproved:r.unapprovedHours,pay:r.estimatedPay,missing:r.payMissing,alerts:r.approvalAlerts})))`));
}

let checks = 0;
function check(name, fn) { fn(); checks++; console.log(`PASS ${name}`); }

check("original request and HR replacement: 9 requested, 4 granted, 5 ungranted", () => {
  const ctx = context();
  fixture(ctx, { requests: [{hours:5,end:"21:30"},{hours:4}], grants:[{hours:4,amount:999}] });
  assert.deepEqual(snapshot(ctx).map(r=>[r.hours,r.approved,r.unapproved,r.pay]), [[5,0,5,0],[4,4,0,999]]);
  assert.equal(evaluate(ctx, "state.scheduleRows.every(r=>r.overtimeHours===0)"), true);
});

check("partial grant attaches once and missing M uses granted hours only", () => {
  const ctx = context();
  fixture(ctx, { requests: [{hours:9,end:"2026/09/16 01:30"}], grants:[{hours:4}] });
  assert.deepEqual(snapshot(ctx).map(r=>[r.hours,r.approved,r.unapproved,r.pay]), [[9,4,5,602]]);
});

check("M=0 is a real zero; M amount takes priority over wages", () => {
  const ctx = context();
  fixture(ctx, { grants:[{hours:4,amount:0},{hours:1,start:"06:30",end:"07:30",amount:12345}] });
  assert.deepEqual(snapshot(ctx).map(r=>r.pay), [12345,0]);
});

check("missing personal wage is shown as incomplete, never borrowed from another employee", () => {
  const ctx = context();
  fixture(ctx, { grants:[{hours:4,person:"Test B"}] });
  assert.equal(snapshot(ctx)[0].missing, true);
  assert.equal(evaluate(ctx, 'approvalView.payText(state.dailyRows).includes("待補")'), true);
});

check("duplicate rows within each sheet do not double totals or money", () => {
  const ctx = context();
  fixture(ctx, { requests:[{hours:4},{hours:4}],grants:[{hours:4,amount:4321},{hours:4,amount:4321}] });
  assert.deepEqual(snapshot(ctx).map(r=>[r.hours,r.approved,r.pay]), [[4,4,4321]]);
});

check("monthly total is source-controlled, never capped or chronologically reassigned", () => {
  const ctx = context();
  fixture(ctx, { requests:[{hours:2,date:"2026/09/01",start:"18:00",end:"20:00"}], grants:[{hours:46,amount:4600}] });
  const records = snapshot(ctx);
  assert.equal(records[0].approved, 0);
  assert.equal(records[0].alerts[0], "超出日/月延展工時無核給");
  assert.equal(evaluate(ctx,"sum(state.dailyRows.map(r=>r.approvedHours))"),46);
  fixture(ctx, { grants:[{hours:51.5,amount:5150}] });
  assert.equal(evaluate(ctx,"sum(state.dailyRows.map(r=>r.approvedHours))"),51.5);
  assert.equal(evaluate(ctx,'state.dailyRows.some(r=>r.workRisks.includes(WORK_RISK.OVER_46_HOURS))'),true);
});

check("missing grant below monthly limit is not assigned a fabricated cap reason", () => {
  const ctx = context();
  fixture(ctx, { requests:[{hours:2}],grants:[] });
  assert.equal(snapshot(ctx)[0].pay,0);
  assert.deepEqual(snapshot(ctx)[0].alerts,[]);
  assert.equal(evaluate(ctx,'riskLabel("unapproved")'),"有未核給時數");
  assert.equal(evaluate(ctx,'riskLabel(approvalView.limitLabel)'),"超出日/月延展工時無核給");
});

check("different employees, days and reasons cannot consume the same grant", () => {
  const ctx = context();
  fixture(ctx, { requests:[{hours:9,end:"2026/09/16 01:30",reason:"Other"}],grants:[{hours:4}] });
  assert.equal(snapshot(ctx).length,2);
  assert.equal(snapshot(ctx).reduce((s,r)=>s+r.approved,0),4);
});

check("summary and detail exports have approval columns and consistent totals", () => {
  const ctx = context();
  fixture(ctx, { requests:[{hours:9,end:"2026/09/16 01:30"}],grants:[{hours:4,amount:1234}] });
  evaluate(ctx,'exportSummaryCsv()');
  assert.equal(evaluate(ctx,'csvResult[1][2]'),9);
  assert.equal(evaluate(ctx,'csvResult[1][3]'),4);
  assert.equal(evaluate(ctx,'csvResult[1][4]'),5);
  evaluate(ctx,'exportDetailCsv()');
  assert.equal(evaluate(ctx,'csvResult.every(r=>r.length===csvResult[0].length)'),true);
  assert.equal(evaluate(ctx,'csvResult[1][7]'),1234);
});

check("unnamed car assignment preserves selection and does not alter named records", () => {
  const ctx = context();
  fixture(ctx, { requests:[{hours:4}],grants:[{hours:4,amount:1234}] });
  evaluate(ctx, `
    var unnamed = { person:"", time:new Date(2026,8,15,16,40), date:new Date(2026,8,15), direction:"entry", sourceLabel:"汽車進場" };
    var named = { ...unnamed, person:"Test B" };
    state.carRows = [unnamed,named]; approvalView.unnamedCarRows = [unnamed];
    state.selectedMonth="2026/09";state.selectedRisk="unapproved";state.selectedDateKey="2026-09-15";
    approvalView.selectCarPerson("Test A");
  `);
  assert.equal(evaluate(ctx,'state.carRows[0].person'),"Test A");
  assert.equal(evaluate(ctx,'state.carRows[1].person'),"Test B");
  assert.equal(evaluate(ctx,'state.selectedRisk'),"unapproved");
  assert.equal(evaluate(ctx,'state.selectedDateKey'),"2026-09-15");
  evaluate(ctx,'approvalView.selectCarPerson("")');
  assert.equal(evaluate(ctx,'state.carRows[0].person'),"");
});

check("invalid source columns cannot silently become zero approved totals", () => {
  const ctx = context();
  fixture(ctx, { grants:[{hours:4,amount:1234}] });
  evaluate(ctx,'workbook.Sheets["核可｜加班單申請"].L1.v="Wrong column";loadWorkbook(workbook,"invalid.xlsx")');
  assert.equal(evaluate(ctx,'statusResult.warning'),true);
  assert.equal(evaluate(ctx,'state.workbookName'),"test.xlsx");
});

check("campus suffix wage matching uses the same person's month and exact match wins", () => {
  const ctx = context();
  fixture(ctx,{grants:[{hours:4,person:"Test A大林"}]});
  assert.equal(snapshot(ctx)[0].missing,false);
  assert.equal(snapshot(ctx)[0].pay,602);
  assert.equal(evaluate(ctx,'approvalView.resolveHourly("Test A大林","2026/10",buildWageMap(state.wageRows))'),0);
  assert.equal(evaluate(ctx,'approvalView.resolveHourly("Test B大林","2026/09",buildWageMap(state.wageRows))'),0);
  assert.equal(evaluate(ctx,'approvalView.resolveHourly("Test A大林","2026/09",new Map([["Test A大林|2026/09",200],["Test A|2026/09",100]]))'),200);
});

check("approved days keep monthly risks for the list while omitting them from day cells", () => {
  const ctx = context();
  fixture(ctx,{requests:[{hours:9,end:"2026/09/16 01:30"}],grants:[{hours:4}]});
  evaluate(ctx,'var day=state.dailyRows.find(r=>r.overtimeRecords.length);day.workRisks.push(WORK_RISK.OVER_46_HOURS)');
  assert.equal(evaluate(ctx,'approvalView.calendarWorkRisks(day).includes(WORK_RISK.OVER_46_HOURS)'),false);
  assert.equal(evaluate(ctx,'renderRiskSummaryCell(day).includes(WORK_RISK.OVER_46_HOURS)'),true);
  assert.equal(evaluate(ctx,'dayCellHtml(day.date,[day]).includes(WORK_RISK.OVER_46_HOURS)'),false);
  assert.equal(evaluate(ctx,'dayCellHtml(day.date,[day]).includes("⏱️延展工時核給")'),true);
  assert.equal(evaluate(ctx,'dayCellHtml(day.date,[day]).includes("💰加班費")'),true);
  assert.equal(snapshot(ctx)[0].alerts[0],"超出日/月延展工時無核給");
});

check("detail display omits source coordinates and estimate metadata, exports retain them", () => {
  const ctx = context();
  fixture(ctx,{requests:[{hours:9,end:"2026/09/16 01:30"}],grants:[{hours:4}]});
  const html=evaluate(ctx,'approvalView.recordHtml(state.overtimeRecords[0])');
  for(const unwanted of ["未核可 E 欄","未核可｜加班單申請 第","依時薪估算","L／M 欄","原因依申請與核給資料推估"]) assert.equal(html.includes(unwanted),false);
  evaluate(ctx,'exportDetailCsv()');
  assert.equal(evaluate(ctx,'csvResult[1][11].includes("第")'),true);
  assert.equal(evaluate(ctx,'csvResult[1][13]'),"依時薪估算");
  evaluate(ctx,'dom.riskSubtitle={};dom.riskTableBody={};approvalView.renderRiskTable(state.dailyRows)');
  assert.equal(evaluate(ctx,'dom.riskTableBody.innerHTML.includes("日／月額度未核給原因依申請與核給資料推估，需以簽核紀錄確認。")'),false);
  assert.equal(evaluate(ctx,'csvResult[1][14].includes("依申請與核給資料推估")'),true);
});

check("late sign-in annotation is display-only across calendar, detail and list", () => {
  const ctx = context();
  fixture(ctx,{requests:[{hours:4}],grants:[{hours:4,amount:1234}]});
  evaluate(ctx,`var record=state.overtimeRecords[0];
    var appCase={start:new Date(2026,8,15,16,30),end:new Date(2026,8,15,20,30)};
    var events=[{time:new Date(2026,8,15,16,40),direction:"entry",sourceGroup:"schedule",sourceLabel:"排班簽到"},
      {time:new Date(2026,8,15,20,30),direction:"exit",sourceGroup:"schedule",sourceLabel:"排班簽退"}];
    var audit=evaluateAttendanceRecord(appCase,events);
    record.attendanceStatuses=audit.statuses;record.attendanceSections=audit.sections;
    var before=JSON.stringify([record.start,record.end,record.attendanceStatuses,record.approvedHours,record.estimatedPay,events]);
  `);
  assert.equal(evaluate(ctx,'record.attendanceStatuses.includes(ATTENDANCE_STATUS.LATE_OVERTIME.short)'),true);
  assert.equal(evaluate(ctx,'buildDayRecordIssueGroups([record])[0].items.some(item=>item.label===approvalView.lateSignInLabel)'),true);
  assert.equal(evaluate(ctx,'buildDayRecordIssueGroups([record])[0].notes.includes(approvalView.lateSignInNote)'),false);
  assert.equal(evaluate(ctx,'approvalView.recordHtml(record).includes(approvalView.lateSignInLabel)'),true);
  assert.equal(evaluate(ctx,'buildRiskSummaryAttendanceBlocks([record])[0].items.some(item=>item.label===approvalView.lateSignInLabel)'),true);
  assert.equal(evaluate(ctx,'JSON.stringify([record.start,record.end,record.attendanceStatuses,record.approvedHours,record.estimatedPay,events])===before'),true);
  const legacy=context("");
  assert.equal(evaluate(legacy,'buildAttendanceDisplayItems([ATTENDANCE_STATUS.LATE_OVERTIME.short])[0].note')??"","");
});

check("confirmed car owner applies only to this workbook's unnamed rows", () => {
  const ctx=context();
  evaluate(ctx,`approvalView.confirmedCarOwners=new Map([["confirmed.xlsx","Test A"]]);
    var cars=[{person:""},{person:"Test B"}];
    approvalView.assignConfirmedCarOwner(cars,["Test A","Test B"],"other.xlsx");`);
  assert.equal(evaluate(ctx,'cars[0].person'),"");
  evaluate(ctx,'approvalView.assignConfirmedCarOwner(cars,["Test B"],"confirmed.xlsx")');
  assert.equal(evaluate(ctx,'cars[0].person'),"");
  evaluate(ctx,'approvalView.assignConfirmedCarOwner(cars,["Test A","Test B"],"confirmed.xlsx")');
  assert.equal(evaluate(ctx,'cars[0].person'),"Test A");
  assert.equal(evaluate(ctx,'cars[1].person'),"Test B");
});

const actualPath = process.argv[2];
if (actualPath) {
  check("supplied workbook reconciles every approved L and M row exactly once", () => {
    const ctx = context();
    ctx.bytes = fs.readFileSync(actualPath);
    ctx.actualName=path.basename(actualPath);
    evaluate(ctx,'var realBook=XLSX.read(bytes,{type:"buffer",cellDates:true});loadWorkbook(realBook,actualName)');
    const controls = JSON.parse(evaluate(ctx, `JSON.stringify((()=>{
      const raw=XLSX.utils.sheet_to_json(realBook.Sheets["核可｜加班單申請"],{header:1,raw:true,defval:""}).slice(2).filter(r=>parseDateValue(r[0])&&r[1]&&Number(r[11])>0);
      return {sourceHours:sum(raw.map(r=>r[11])),resultHours:sum(state.overtimeRecords.map(r=>r.approvedHours)),sourcePay:sum(raw.map(r=>approvalView.numericAmount(r[12])||0)),resultPay:sum(state.overtimeRecords.flatMap(r=>r.approvals).map(g=>g.reportedAmount||0)),count:raw.length,linked:state.overtimeRecords.flatMap(r=>r.approvals).length};
    })())`));
    assert.equal(controls.sourceHours,controls.resultHours);
    assert.equal(controls.sourcePay,controls.resultPay);
    assert.equal(controls.count,controls.linked);
    const example = JSON.parse(evaluate(ctx,'JSON.stringify(state.dailyRows.filter(r=>r.dateKey==="2026-09-15"&&r.overtimeHours>0).map(r=>[r.overtimeHours,r.approvedHours,r.unapprovedHours]))'));
    assert.deepEqual(example,[[9,4,5]]);
    const grantCase=JSON.parse(evaluate(ctx,'JSON.stringify(state.overtimeRecords.filter(r=>r.person.startsWith("蘇哲民")&&r.dateKey==="2026-08-06"&&r.approvedHours>0).map(r=>({hours:r.approvedHours,pay:r.estimatedPay,missing:r.payMissing,hourly:r.hourly})))'));
    assert.deepEqual(grantCase,[{hours:4,pay:2114,missing:false,hourly:351}]);
    assert.equal(evaluate(ctx,'state.carRows.length>0&&state.carRows.every(row=>row.person===approvalView.confirmedCarOwners.get(actualName))'),true);
  });

  const oldPath = process.argv[3];
  if (oldPath) check("original entry matches the committed version on the original workbook", () => {
    const baseline = execFileSync("git",["show","HEAD:HR-Project/app.js"],{cwd:root,encoding:"utf8"});
    const currentContext=context("",app), oldContext=context("",baseline);
    const getResult=(ctx)=>{
      ctx.bytes=fs.readFileSync(oldPath);
      evaluate(ctx,'loadWorkbook(XLSX.read(bytes,{type:"buffer",cellDates:true}),"old.xlsx")');
      return evaluate(ctx,'JSON.stringify(state.overtimeRecords.map(r=>[r.person,r.dateKey,r.hours,r.estimatedPay,r.attendanceStatuses,r.workRisks]))');
    };
    assert.equal(getResult(currentContext),getResult(oldContext));
  });
}
console.log(`${checks} checks passed`);
