// =====================================================================
// ===      CODE.GS - NOVASOL ERP V9.0 (MULTI-BRANCH EDITION)        ===
// ===  เพิ่ม: Supabase Sync ทุก save/delete function                 ===
// ===  หมายเหตุ: โค้ดเดิมไม่มีการเปลี่ยนแปลง                       ===
// ===  มีเพิ่มเฉพาะบรรทัดที่มี comment // ✅ SUPABASE               ===
// =====================================================================

const DB_CONFIG = {
  supplies: {
    name: 'Supplies',
    headers: ['ID', 'รายการ', 'รูปภาพ', 'ราคาต่อหน่วย', 'คงเหลือ', 'จำนวนที่ทำรายการ', 'ประเภท', 'สถานที่/โครงการ', 'วันที่', 'ผู้บันทึก'],
  },
  contracts: {
    name: 'Contracts',
    headers: ['ID', 'ชื่อโครงการ/หน่วยงาน', 'วันเริ่มสัญญา', 'วันสิ้นสุดสัญญา', 'มูลค่าสัญญารวม', 'งบค่าแรง(ประมาณการ)', 'สถานะ'],
  },
  labor: {
    name: 'LaborCosts',
    headers: ['ID', 'โครงการ', 'วันที่จ่าย', 'จำนวนเงิน', 'รายละเอียด/งวดงาน', 'ผู้บันทึก']
  },
  users: {
    name: 'Users',
    headers: ['Username', 'Password', 'Name', 'Role', 'Branch'],
    defaultAdmin: ['admin', '1234', 'Admin User', 'Admin', 'ทุกสาขา']
  },
  requests: {
    name: 'Requests',
    headers: ['ID', 'รายการ', 'จำนวน', 'สถานที่/โครงการ', 'ผู้ขอเบิก', 'วันที่ขอ', 'สถานะ', 'วันที่อนุมัติ'],
  },
  workers: {
    name: 'Workers',
    headers: ['ID', 'ชื่อ-สกุล', 'บัตรประชาชน', 'วันเกิด', 'เบอร์โทร', 'ตำแหน่ง', 'วันเริ่มงาน', 'วันสิ้นสุด', 'ธนาคาร', 'เลขบัญชี', 'ฐานเงินเดือน', 'สถานะ', 'สาขา'],
  },
  payroll: {
    name: 'Payroll',
    headers: ['ID', 'วันที่จ่าย', 'ชื่อพนักงาน', 'ค่าตำแหน่ง', 'รายได้อื่น', 'รวมรายรับ', 'หักภาษี', 'หักประกันสังคม', 'หักอื่นๆ', 'รับสุทธิ', 'หมายเหตุ', 'สาขา'],
  },
  capital: {
    name: 'CapitalInjection',
    headers: ['ID', 'วันที่', 'จำนวนเงิน', 'แหล่งที่มา', 'โครงการ', 'หมายเหตุ', 'ผู้บันทึก'],
  },
  branches: {
    name: 'Branches',
    headers: ['ID', 'ชื่อสาขา', 'ที่อยู่', 'ผู้จัดการสาขา', 'เบอร์โทร', 'วันเปิด', 'สถานะ'],
    defaultBranch: ['BR-001', 'สำนักงานใหญ่', '-', 'Admin User', '-', '', 'เปิดใช้งาน']
  }
};

function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    supplies:  ensureSheet(ss, DB_CONFIG.supplies),
    contracts: ensureSheet(ss, DB_CONFIG.contracts),
    labor:     ensureSheet(ss, DB_CONFIG.labor),
    users:     ensureSheet(ss, DB_CONFIG.users, true),
    requests:  ensureSheet(ss, DB_CONFIG.requests),
    workers:   ensureSheet(ss, DB_CONFIG.workers),
    payroll:   ensureSheet(ss, DB_CONFIG.payroll),
    capital:   ensureSheet(ss, DB_CONFIG.capital),
    branches:  ensureSheet(ss, DB_CONFIG.branches, false, true),
  };
  return sheets;
}

function ensureSheet(ss, config, isUserSheet = false, isBranchSheet = false) {
  let sheet = ss.getSheetByName(config.name);
  if (!sheet) {
    sheet = ss.insertSheet(config.name);
    sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    sheet.setFrozenRows(1);
    if (isUserSheet) sheet.appendRow(DB_CONFIG.users.defaultAdmin);
    if (isBranchSheet) sheet.appendRow(DB_CONFIG.branches.defaultBranch);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (config.name === 'Workers' && currentHeaders.length < 13) {
      sheet.getRange(1, 13).setValue('สาขา');
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const range = sheet.getRange(2, 13, lastRow - 1, 1);
        range.setValue('สำนักงานใหญ่');
      }
    }
    if (config.name === 'Payroll' && currentHeaders.length < 12) {
      sheet.getRange(1, 12).setValue('สาขา');
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 12, lastRow - 1, 1).setValue('สำนักงานใหญ่');
      }
    }
    if (config.name === 'Users' && currentHeaders.length < 5) {
      sheet.getRange(1, 5).setValue('Branch');
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 5, lastRow - 1, 1).setValue('ทุกสาขา');
      }
    }
  }
  return sheet;
}

// ==========================================
// 🌐 HTTP HANDLERS
// ==========================================
function doGet(request) {
  if (request && request.parameter && request.parameter.action) {
    return handleApiRequest(request.parameter, null);
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('NovaSol ERP V9.0')
    .setFaviconUrl('https://i.postimg.cc/FHGkmGKB/NOVASOL-1/logo.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(request) {
  try {
    const body = JSON.parse(request.postData.contents);
    const action = body.action;
    const params = body.params || {};
    const result = dispatchAction(action, params);
    return buildCorsResponse(result);
  } catch (e) {
    return buildCorsResponse({ success: false, error: e.toString() });
  }
}

function handleApiRequest(params, body) {
  try {
    const result = dispatchAction(params.action, params);
    return buildCorsResponse(result);
  } catch (e) {
    return buildCorsResponse({ success: false, error: e.toString() });
  }
}

function buildCorsResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function dispatchAction(action, params) {
  switch (action) {
    case 'loginSystem':       return loginSystem(params.username, params.password);
    case 'getUsersList':      return getUsersList();
    case 'saveUser':          return saveUser(params);
    case 'deleteUser':        return deleteUser(params.row);
    case 'getBranchesData':   return getBranchesData();
    case 'saveBranch':        return saveBranch(params);
    case 'deleteBranch':      return deleteBranch(params.id);
    case 'getBranchSummary':  return getBranchSummary(params.branchName);
    case 'getSuppliesData':   return getSuppliesData();
    case 'saveSupply':        return saveSupply(params);
    case 'deleteSupply':      return deleteSupply(params.id);
    case 'getMaterialList':   return getMaterialList();
    case 'getContractsData':  return getContractsData();
    case 'saveContract':      return saveContract(params);
    case 'deleteContract':    return deleteContract(params.id);
    case 'getLaborData':      return getLaborData();
    case 'saveLabor':         return saveLabor(params);
    case 'deleteLabor':       return deleteLabor(params.id);
    case 'getRequestsData':   return getRequestsData();
    case 'saveRequest':       return saveRequest(params);
    case 'updateRequestStatus': return updateRequestStatus(params.id, params.status, params.approver);
    case 'deleteRequest':     return deleteRequest(params.id);
    case 'getWorkersData':    return getWorkersData(params.branch);
    case 'saveWorker':        return saveWorker(params);
    case 'deleteWorker':      return deleteWorker(params.id);
    case 'getActiveWorkersList': return getActiveWorkersList(params.branch);
    case 'transferWorker':    return transferWorker(params.id, params.newBranch);
    case 'getPayrollData':    return getPayrollData(params.branch);
    case 'savePayroll':       return savePayroll(params);
    case 'deletePayroll':     return deletePayroll(params.id);
    case 'getPayslipData':    return getPayslipData(params.payrollId);
    case 'getPayrollSummaryByBranch': return getPayrollSummaryByBranch(params.monthStr);
    case 'getCapitalData':    return getCapitalData();
    case 'saveCapital':       return saveCapital(params);
    case 'deleteCapital':     return deleteCapital(params.id);
    case 'getCapitalSummary': return getCapitalSummary();
    case 'getDashboardStats': return getDashboardStats();
    case 'getInventoryStatus': return getInventoryStatus();
    case 'getExpenseReport':  return getExpenseReport(params.monthStr);
    case 'getProjectExpenseReport': return getProjectExpenseReport(params.monthStr);
    case 'evaluateMaterial':  return evaluateMaterial(params.itemName, params.workDays);
    default:
      return { success: false, error: 'Unknown action: ' + action };
  }
}

// ==========================================
// 🔐 AUTH
// ==========================================
function loginSystem(u, p) {
  const { users } = initDatabase();
  const data = users.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === u && String(data[i][1]) === p) {
      return {
        success: true,
        name: data[i][2],
        role: data[i][3],
        branch: data[i][4] || 'ทุกสาขา'
      };
    }
  }
  return { success: false, message: 'Login Failed' };
}

// ==========================================
// 👥 USERS
// ==========================================
function getUsersList() {
  const { users } = initDatabase();
  const data = users.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map((r, i) => ({
    row: i + 2,
    username: r[0],
    password: r[1],
    name: r[2],
    role: r[3],
    branch: r[4] || 'ทุกสาขา'
  }));
}

function saveUser(form) {
  const { users } = initDatabase();
  const username = form.user || form.username || '';
  const password = form.pass || form.password || '';
  const name     = form.name   || '';
  const role     = form.role   || '';
  const branch   = form.branch || 'ทุกสาขา';

  if (form.row) {
    users.getRange(form.row, 1, 1, 5).setValues([[username, password, name, role, branch]]);
  } else {
    users.appendRow([username, password, name, role, branch]);
  }

  // ✅ SUPABASE: sync user
  try {
    syncUserToSupabase({
      gas_row:  String(form.row || (users.getLastRow())),
      username: username,
      password: password,
      name:     name,
      role:     role,
      branch:   branch
    });
  } catch(e) { Logger.log('⚠️ Supabase syncUser failed: ' + e.toString()); }

  return { success: true };
}

function deleteUser(row) {
  // ✅ SUPABASE: ดึง gas_row ก่อนลบ แล้วลบใน Supabase
  try {
    deleteUserFromSupabase(String(row));
  } catch(e) { Logger.log('⚠️ Supabase deleteUser failed: ' + e.toString()); }

  initDatabase().users.deleteRow(row);
  return { success: true };
}

// ==========================================
// 🏢 BRANCHES
// ==========================================
function getBranchesData() {
  const { branches } = initDatabase();
  const lastRow = branches.getLastRow();
  if (lastRow <= 1) return [];
  return branches.getRange(2, 1, lastRow - 1, DB_CONFIG.branches.headers.length).getValues()
    .filter(r => r[0] !== '')
    .map(row => ({
      id:       row[0],
      name:     row[1],
      address:  row[2] || '',
      manager:  row[3] || '',
      phone:    row[4] || '',
      openDate: row[5] ? Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      rawOpen:  row[5] ? Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      status:   row[6] || 'เปิดใช้งาน'
    }));
}

function saveBranch(form) {
  const { branches } = initDatabase();
  if (!form.name) return { success: false, message: 'กรุณาระบุชื่อสาขา' };

  const rowData = [
    form.name,
    form.address  || '',
    form.manager  || '',
    form.phone    || '',
    form.openDate ? new Date(form.openDate) : '',
    form.status   || 'เปิดใช้งาน'
  ];

  let savedId = form.id;

  if (form.id) {
    const data = branches.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(form.id)) {
        branches.getRange(i + 1, 2, 1, 6).setValues([rowData]);

        // ✅ SUPABASE: update branch
        try {
          syncBranchToSupabase({
            gas_id:    String(form.id),
            name:      form.name,
            address:   form.address  || '',
            manager:   form.manager  || '',
            phone:     form.phone    || '',
            open_date: form.openDate || null,
            status:    form.status   || 'เปิดใช้งาน'
          });
        } catch(e) { Logger.log('⚠️ Supabase syncBranch failed: ' + e.toString()); }

        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบสาขาที่ต้องการแก้ไข' };
  } else {
    savedId = 'BR-' + new Date().getTime();
    branches.appendRow([savedId, ...rowData]);

    // ✅ SUPABASE: insert branch
    try {
      syncBranchToSupabase({
        gas_id:    savedId,
        name:      form.name,
        address:   form.address  || '',
        manager:   form.manager  || '',
        phone:     form.phone    || '',
        open_date: form.openDate || null,
        status:    form.status   || 'เปิดใช้งาน'
      });
    } catch(e) { Logger.log('⚠️ Supabase syncBranch failed: ' + e.toString()); }

    return { success: true };
  }
}

function deleteBranch(id) {
  try {
    if (id === 'BR-001') return { success: false, message: 'ไม่สามารถลบสำนักงานใหญ่ได้' };
    const { branches } = initDatabase();
    const data = branches.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        branches.deleteRow(i + 1);

        // ✅ SUPABASE: delete branch
        try {
          deleteBranchFromSupabase(String(id));
        } catch(e) { Logger.log('⚠️ Supabase deleteBranch failed: ' + e.toString()); }

        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบสาขาที่ต้องการลบ' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getBranchSummary(branchName) {
  const workers  = getWorkersData(branchName);
  const payrolls = getPayrollData(branchName);
  const activeWorkers = workers.filter(w => w.status === 'ทำงานอยู่').length;
  const totalWorkers  = workers.length;
  let totalPayroll = 0;
  payrolls.forEach(p => { totalPayroll += parseFloat(p.netPay || 0); });
  return { branchName, totalWorkers, activeWorkers, totalPayroll };
}

// ==========================================
// 📝 REQUESTS
// ==========================================
function getRequestsData() {
  const { requests } = initDatabase();
  const lastRow = requests.getLastRow();
  if (lastRow <= 1) return [];
  return requests.getRange(2, 1, lastRow - 1, DB_CONFIG.requests.headers.length).getValues().map((r) => ({
    id: r[0], itemName: r[1], qty: r[2], location: r[3], requester: r[4],
    reqDate: formatDate(r[5]), status: r[6], approvedDate: formatDate(r[7])
  })).sort((a, b) => b.id - a.id);
}

function saveRequest(form) {
  const { requests } = initDatabase();

  if (form.id) {
    const data = requests.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] == form.id);
    if (idx !== -1) {
      requests.getRange(idx + 1, 2, 1, 3).setValues([[form.itemName, form.qty, form.location]]);

      // ✅ SUPABASE: update request
      try {
        syncRequestToSupabase({
          gas_id:    String(form.id),
          req_date:  null,
          requester: form.requester || '',
          item_name: form.itemName,
          qty:       form.qty,
          location:  form.location,
          status:    'รออนุมัติ',
          approver:  ''
        });
      } catch(e) { Logger.log('⚠️ Supabase syncRequest failed: ' + e.toString()); }

      return { success: true, message: 'Updated' };
    }
    return { success: false, message: 'ID Not Found' };
  } else {
    const id = requests.getLastRow() === 1 ? 1 : Number(requests.getRange(requests.getLastRow(), 1).getValue()) + 1;
    const today = new Date();
    requests.appendRow([id, form.itemName, form.qty, form.location, form.requester, today, 'รออนุมัติ', '']);

    // ✅ SUPABASE: insert request
    try {
      syncRequestToSupabase({
        gas_id:    String(id),
        req_date:  Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        requester: form.requester || '',
        item_name: form.itemName,
        qty:       form.qty,
        location:  form.location,
        status:    'รออนุมัติ',
        approver:  ''
      });
    } catch(e) { Logger.log('⚠️ Supabase syncRequest failed: ' + e.toString()); }

    return { success: true, message: 'Created' };
  }
}

function updateRequestStatus(id, status, approver) {
  const { requests } = initDatabase();
  const data = requests.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == id);
  if (idx !== -1) {
    requests.getRange(idx + 1, 7).setValue(status);
    if (status === 'อนุมัติ/จ่ายแล้ว') requests.getRange(idx + 1, 8).setValue(new Date());

    // ✅ SUPABASE: update request status
    try {
      const row = data[idx];
      syncRequestToSupabase({
        gas_id:    String(id),
        req_date:  row[5] ? Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null,
        requester: row[4] || '',
        item_name: row[1] || '',
        qty:       row[2] || 0,
        location:  row[3] || '',
        status:    status,
        approver:  approver || ''
      });
    } catch(e) { Logger.log('⚠️ Supabase updateRequestStatus failed: ' + e.toString()); }
  }
  return { success: true };
}

function deleteRequest(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
    if (!sheet) throw new Error("ไม่พบชีต Requests");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);

        // ✅ SUPABASE: delete request
        try {
          deleteRequestFromSupabase(String(id));
        } catch(e) { Logger.log('⚠️ Supabase deleteRequest failed: ' + e.toString()); }

        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบ ID ที่ต้องการลบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 📦 SUPPLIES
// ==========================================
function getMaterialList() {
  const { supplies } = initDatabase();
  const data = supplies.getDataRange().getValues();
  const materials = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'รับเข้า') materials.add(data[i][1]);
  }
  return Array.from(materials).sort();
}

function getSuppliesData() {
  const { supplies } = initDatabase();
  const lastRow = supplies.getLastRow();
  if (lastRow <= 1) return [];
  return supplies.getRange(2, 1, lastRow - 1, DB_CONFIG.supplies.headers.length).getValues()
    .filter(r => r[0] !== '')
    .map(row => ({
      id:               row[0],
      itemName:         row[1],
      imageUrl: (() => {
        const raw = String(row[2] || '').trim();
        if (raw.startsWith('https://')) return raw;
        const m = raw.match(/src=["']?(https?:\/\/[^"'\s>]+)/i);
        return m ? m[1] : '';
      })(),
      unitPrice:        parseFloat(row[3]) || 0,
      itemQuantity:     parseFloat(row[5]) || 0,
      actionType:       row[6],
      deliveryLocation: row[7],
      requestDate:      formatDate(row[8]),
      requester:        row[9],
      rawDate:          row[8] ? Utilities.formatDate(new Date(row[8]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ""
    })).sort((a, b) => b.rawDate.localeCompare(a.rawDate));
}

function saveSupply(form) {
  const { supplies } = initDatabase();
  let finalImgUrl = '';
  const rawExisting = (form.imageUrl || '').trim();
  if (rawExisting.startsWith('https://')) {
    finalImgUrl = rawExisting;
  } else if (rawExisting.includes('src=')) {
    const m = rawExisting.match(/src=["']?(https?:\/\/[^"'\s>]+)/i);
    finalImgUrl = m ? m[1] : '';
  }
  if (form.imageBase64 && form.imageBase64.includes("base64,")) {
    try {
      finalImgUrl = uploadImageToDrive(form.imageBase64, form.imageName || 'image_' + new Date().getTime());
    } catch (e) { Logger.log("Upload Failed: " + e.message); }
  } else if (finalImgUrl.length > 2000) {
    finalImgUrl = "";
  }

  let unitPrice = parseFloat(form.unitPrice) || 0;
  if (form.actionType === 'จ่ายออก' && unitPrice === 0) {
    const status = getInventoryStatus();
    const item = status.find(x => x.name === form.itemName);
    if (item && item.countIn > 0) {
      unitPrice = item.totalValue / item.balance;
      if (!isFinite(unitPrice)) unitPrice = 0;
    }
  }

  const rowData = [
    form.itemName, finalImgUrl, unitPrice, 0,
    form.itemQuantity, form.actionType, form.deliveryLocation,
    new Date(form.requestDate), form.requester
  ];

  let savedId = form.id;

  if (form.id) {
    const data = supplies.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] == form.id);
    if (idx === -1) return { success: false, message: "ID Not Found" };
    supplies.getRange(idx + 1, 2, 1, 9).setValues([rowData]);
  } else {
    savedId = supplies.getLastRow() === 1 ? 1 : Number(supplies.getRange(supplies.getLastRow(), 1).getValue()) + 1;
    supplies.appendRow([savedId, ...rowData]);
  }

  // ✅ SUPABASE: sync supply
  try {
    syncSupplyToSupabase({
      gas_id:            String(savedId),
      request_date:      form.requestDate || null,
      item_name:         form.itemName    || '',
      action_type:       form.actionType  || '',
      item_quantity:     form.itemQuantity || 0,
      unit_price:        unitPrice,
      delivery_location: form.deliveryLocation || '',
      requester:         form.requester   || '',
      image_url:         finalImgUrl
    });
  } catch(e) { Logger.log('⚠️ Supabase syncSupply failed: ' + e.toString()); }

  return { success: true };
}

function deleteSupply(id) {
  const { supplies } = initDatabase();
  const idx = supplies.getDataRange().getValues().findIndex(r => r[0] == id);
  if (idx !== -1) {
    supplies.deleteRow(idx + 1);

    // ✅ SUPABASE: delete supply
    try {
      deleteSupplyFromSupabase(String(id));
    } catch(e) { Logger.log('⚠️ Supabase deleteSupply failed: ' + e.toString()); }
  }
  return { success: true };
}

// ==========================================
// 📁 CONTRACTS
// ==========================================
function getContractsData() {
  const { contracts } = initDatabase();
  const lastRow = contracts.getLastRow();
  if (lastRow <= 1) return [];
  return contracts.getRange(2, 1, lastRow - 1, DB_CONFIG.contracts.headers.length).getValues()
    .filter(r => r[0] !== '')
    .map(row => ({
      id:          row[0],
      name:        String(row[1]),
      startDate:   formatDate(row[2]),
      endDate:     formatDate(row[3]),
      totalValue:  parseFloat(row[4]) || 0,
      laborBudget: parseFloat(row[5]) || 0,
      status:      row[6],
      rawStart:    row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      rawEnd:      row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    }));
}

function saveContract(form) {
  const { contracts } = initDatabase();
  const start  = new Date(form.startDate);
  const end    = new Date(form.endDate);
  const status = new Date() < start ? 'Pending' : (new Date() > end ? 'Expired' : 'Active');
  const rowData = [form.name, start, end, form.totalBudget, form.laborBudget, status];

  let savedId = form.id;

  if (form.id) {
    const idx = contracts.getDataRange().getValues().findIndex(r => r[0] == form.id);
    if (idx === -1) return { success: false };
    contracts.getRange(idx + 1, 2, 1, 6).setValues([rowData]);
  } else {
    savedId = contracts.getLastRow() === 1 ? 1 : Number(contracts.getRange(contracts.getLastRow(), 1).getValue()) + 1;
    contracts.appendRow([savedId, ...rowData]);
  }

  // ✅ SUPABASE: sync contract
  try {
    syncContractToSupabase({
      gas_id:       String(savedId),
      name:         form.name,
      start_date:   form.startDate || null,
      end_date:     form.endDate   || null,
      total_budget: form.totalBudget || 0,
      labor_budget: form.laborBudget || 0,
      status:       status
    });
  } catch(e) { Logger.log('⚠️ Supabase syncContract failed: ' + e.toString()); }

  return { success: true };
}

function deleteContract(id) {
  const { contracts } = initDatabase();
  const idx = contracts.getDataRange().getValues().findIndex(r => r[0] == id);
  if (idx !== -1) {
    contracts.deleteRow(idx + 1);

    // ✅ SUPABASE: delete contract
    try {
      deleteContractFromSupabase(String(id));
    } catch(e) { Logger.log('⚠️ Supabase deleteContract failed: ' + e.toString()); }
  }
  return { success: true };
}

// ==========================================
// 👷 LABOR COSTS
// ==========================================
function getLaborData() {
  const { labor } = initDatabase();
  const lastRow = labor.getLastRow();
  if (lastRow <= 1) return [];
  return labor.getRange(2, 1, lastRow - 1, DB_CONFIG.labor.headers.length).getValues()
    .map(r => ({
      id: r[0], project: r[1], date: formatDate(r[2]), amount: parseFloat(r[3]) || 0,
      note: r[4], recorder: r[5],
      rawDate: r[2] ? Utilities.formatDate(new Date(r[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
    })).sort((a, b) => b.id - a.id);
}

function saveLabor(form) {
  const { labor } = initDatabase();
  const rowData = [form.project, new Date(form.date), form.amount, form.note, form.recorder];
  if (form.id) {
    const data = labor.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] == form.id);
    if (idx === -1) return { success: false };
    labor.getRange(idx + 1, 2, 1, 5).setValues([rowData]);
  } else {
    const id = labor.getLastRow() === 1 ? 1 : Number(labor.getRange(labor.getLastRow(), 1).getValue()) + 1;
    labor.appendRow([id, ...rowData]);
  }
  return { success: true };
}

function deleteLabor(id) {
  const { labor } = initDatabase();
  const idx = labor.getDataRange().getValues().findIndex(r => r[0] == id);
  if (idx !== -1) labor.deleteRow(idx + 1);
  return { success: true };
}

// ==========================================
// 📊 INVENTORY
// ==========================================
function getInventoryStatus() {
  const supplies = getSuppliesData();
  const inventory = {};
  supplies.forEach(item => {
    const name = (item.itemName || '').toString();
    if (name.includes('เพิ่มทุน') || name.includes('ค่างวด') || name.includes('รายรับ')) return;
    if (!inventory[name]) {
      inventory[name] = { in: 0, out: 0, balance: 0, totalPriceIn: 0, countIn: 0, lastImage: item.imageUrl };
    }
    if (item.imageUrl) inventory[name].lastImage = item.imageUrl;
    if (item.actionType === 'รับเข้า') {
      inventory[name].in += item.itemQuantity;
      inventory[name].totalPriceIn += (item.unitPrice * item.itemQuantity);
      inventory[name].countIn += item.itemQuantity;
    } else if (item.actionType === 'จ่ายออก') {
      inventory[name].out += item.itemQuantity;
    }
  });
  return Object.keys(inventory).map(name => {
    const item = inventory[name];
    const balance = item.in - item.out;
    const avgPrice = item.countIn > 0 ? (item.totalPriceIn / item.countIn) : 0;
    return {
      name, in: item.in, out: item.out, balance,
      totalValue: balance * avgPrice, imageUrl: item.lastImage,
      status: balance <= 0 ? 'สินค้าหมด' : (balance < 5 ? 'ใกล้หมด' : 'ปกติ'),
      countIn: item.countIn
    };
  });
}

// ==========================================
// 📈 DASHBOARD ANALYTICS
// ==========================================
function getDashboardStats() {
  const supplies  = getSuppliesData();
  const contracts = getContractsData();
  const laborData = getLaborData();
  const materialCostByProj = {};
  const laborCostByProj    = {};

  supplies.forEach(s => {
    const name = (s.itemName || '').toString();
    if (name.includes('เพิ่มทุน') || name.includes('ค่างวด') || name.includes('รายรับ')) return;
    if (s.actionType === 'จ่ายออก') {
      const p    = s.deliveryLocation;
      const cost = s.unitPrice * s.itemQuantity;
      materialCostByProj[p] = (materialCostByProj[p] || 0) + cost;
    }
  });

  laborData.forEach(l => {
    laborCostByProj[l.project] = (laborCostByProj[l.project] || 0) + l.amount;
  });

  let totalRevenue = 0, totalExpense = 0;
  const labels = [], budgets = [], expenses = [], projects = [];

  contracts.forEach(c => {
    const matCost   = materialCostByProj[c.name] || 0;
    const laborCost = laborCostByProj[c.name]    || 0;
    const totalCost = matCost + laborCost;
    const progress  = c.totalValue > 0 ? (totalCost / c.totalValue * 100) : 0;
    totalRevenue += c.totalValue;
    totalExpense += totalCost;
    labels.push(c.name);
    budgets.push(c.totalValue);
    expenses.push(totalCost);
    projects.push({
      name: c.name, budget: c.totalValue, expense: totalCost,
      profit: c.totalValue - totalCost,
      isLoss: (c.totalValue - totalCost) < 0,
      costProgress: progress
    });
  });

  return { revenue: totalRevenue, expense: totalExpense, labels, budgets, expenses, projects };
}

// ==========================================
// 💜 CAPITAL INJECTION
// ==========================================
function getCapitalData() {
  const { capital } = initDatabase();
  const lastRow = capital.getLastRow();
  if (lastRow <= 1) return [];
  return capital.getRange(2, 1, lastRow - 1, DB_CONFIG.capital.headers.length).getValues()
    .filter(r => r[0] !== '')
    .map(row => ({
      id:       row[0],
      date:     row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      rawDate:  row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      amount:   parseFloat(row[2]) || 0,
      source:   row[3] || '',
      project:  row[4] || '',
      note:     row[5] || '',
      recorder: row[6] || ''
    })).sort((a, b) => b.rawDate.localeCompare(a.rawDate));
}

function saveCapital(form) {
  const { capital } = initDatabase();
  if (!form.date || !form.amount || !form.source)
    return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
  const amount = parseFloat(form.amount);
  if (isNaN(amount) || amount <= 0)
    return { success: false, message: 'จำนวนเงินต้องมากกว่า 0' };

  const rowData = [new Date(form.date), amount, form.source, form.project || '', form.note || '', form.recorder || ''];
  let savedId = form.id;

  if (form.id) {
    const data = capital.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(form.id)) {
        capital.getRange(i + 1, 2, 1, 6).setValues([rowData]);

        // ✅ SUPABASE: update capital
        try {
          syncCapitalToSupabase({
            gas_id:   String(form.id),
            cap_date: form.date    || null,
            amount:   amount,
            source:   form.source  || '',
            project:  form.project || '',
            note:     form.note    || '',
            recorder: form.recorder || ''
          });
        } catch(e) { Logger.log('⚠️ Supabase syncCapital failed: ' + e.toString()); }

        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบรายการ' };
  } else {
    savedId = 'CAP-' + new Date().getTime();
    capital.appendRow([savedId, ...rowData]);

    // ✅ SUPABASE: insert capital
    try {
      syncCapitalToSupabase({
        gas_id:   savedId,
        cap_date: form.date    || null,
        amount:   amount,
        source:   form.source  || '',
        project:  form.project || '',
        note:     form.note    || '',
        recorder: form.recorder || ''
      });
    } catch(e) { Logger.log('⚠️ Supabase syncCapital failed: ' + e.toString()); }

    return { success: true };
  }
}

function deleteCapital(id) {
  try {
    const { capital } = initDatabase();
    const data = capital.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        capital.deleteRow(i + 1);

        // ✅ SUPABASE: delete capital
        try {
          deleteCapitalFromSupabase(String(id));
        } catch(e) { Logger.log('⚠️ Supabase deleteCapital failed: ' + e.toString()); }

        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบรายการ' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getCapitalSummary() {
  const capitals = getCapitalData();
  const bySource = {};
  let total = 0;
  capitals.forEach(c => {
    bySource[c.source] = (bySource[c.source] || 0) + c.amount;
    total += c.amount;
  });
  return {
    total,
    bySource: Object.keys(bySource).map(k => ({ source: k, amount: bySource[k] })),
    count: capitals.length
  };
}

// ==========================================
// 📋 EXPENSE REPORT
// ==========================================
function getExpenseReport(monthStr) {
  if (!monthStr) return [];
  const [year, month] = monthStr.split('-').map(Number);
  const supplies = getSuppliesData();
  const result = [];
  supplies.forEach(item => {
    if (!item.rawDate) return;
    const [y, m] = item.rawDate.split('-').map(Number);
    if (item.actionType === 'จ่ายออก' && y === year && m === month) {
      result.push({ date: item.requestDate, item: item.itemName, loc: item.deliveryLocation, amount: item.itemQuantity * item.unitPrice });
    }
  });
  return result.sort((a, b) => a.date.localeCompare(b.date));
}

function getProjectExpenseReport(monthStr) {
  if (!monthStr) return [];
  const [year, month] = monthStr.split('-').map(Number);
  const supplies = getSuppliesData();
  const projectExpenses = {};
  supplies.forEach(item => {
    if (!item.rawDate) return;
    const [y, m] = item.rawDate.split('-').map(Number);
    if (item.actionType === 'จ่ายออก' && y === year && m === month) {
      const proj = item.deliveryLocation;
      if (!projectExpenses[proj]) projectExpenses[proj] = { total: 0, items: [] };
      projectExpenses[proj].total += (item.itemQuantity * item.unitPrice);
      projectExpenses[proj].items.push({ name: item.itemName, qty: item.itemQuantity, price: item.unitPrice, cost: (item.itemQuantity * item.unitPrice), date: item.requestDate });
    }
  });
  return Object.keys(projectExpenses).map(proj => ({ projectName: proj, totalExpense: projectExpenses[proj].total, details: projectExpenses[proj].items }));
}

// ==========================================
// 🚀 WORKERS
// ==========================================
function getWorkersData(branchFilter) {
  const { workers } = initDatabase();
  const lastRow = workers.getLastRow();
  if (lastRow <= 1) return [];
  const numCols = Math.max(workers.getLastColumn(), 13);
  const data = workers.getRange(2, 1, lastRow - 1, numCols).getValues();
  return data
    .filter(r => {
      if (!branchFilter || branchFilter === 'ทุกสาขา' || branchFilter === '') return true;
      return String(r[12] || 'สำนักงานใหญ่') === branchFilter;
    })
    .map(r => ({
      id:        r[0],
      name:      r[1],
      idCard:    r[2] ? String(r[2]).replace(/^'/, '') : '',
      dob:       r[3] ? Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      phone:     r[4] ? String(r[4]).replace(/^'/, '') : '',
      role:      r[5],
      startDate: r[6] ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      endDate:   r[7] ? Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      bankName:  r[8],
      bankAcc:   r[9] ? String(r[9]).replace(/^'/, '') : '',
      baseWage:  parseFloat(r[10]) || 0,
      status:    r[11] || 'ทำงานอยู่',
      branch:    r[12] || 'สำนักงานใหญ่',
      rawStart:  r[6] ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      rawEnd:    r[7] ? Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    }));
}

function saveWorker(form) {
  const { workers } = initDatabase();
  const name      = form.name     || "";
  const idCard    = form.idCard   ? "'" + form.idCard : "";
  const dob       = form.dob      || "";
  const phone     = form.phone    ? "'" + form.phone : "";
  const position  = form.role     || form.position || "";
  const startDate = form.startDate || "";
  const endDate   = form.endDate  || "";
  const bank      = form.bankName || form.bank     || "";
  const account   = (form.bankAcc || form.account) ? "'" + (form.bankAcc || form.account) : "";
  const salary    = parseFloat(form.baseWage || form.salary || 0);
  const status    = form.status   || 'ทำงานอยู่';
  const branch    = form.branch   || 'สำนักงานใหญ่';
  const rowData   = [name, idCard, dob, phone, position, startDate, endDate, bank, account, salary, status, branch];

  let savedId = form.id;

  if (form.id) {
    const data = workers.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(form.id)) {
        workers.getRange(i + 1, 2, 1, 12).setValues([rowData]);

        // ✅ SUPABASE: update worker
        try {
          syncWorkerToSupabase({
            gas_id:     String(form.id),
            name:       name,
            id_card:    form.idCard    || '',
            dob:        dob,
            phone:      form.phone     || '',
            role:       position,
            start_date: startDate || null,
            end_date:   endDate   || null,
            bank_name:  bank,
            bank_acc:   form.bankAcc   || '',
            base_wage:  salary,
            branch:     branch,
            status:     status
          });
        } catch(e) { Logger.log('⚠️ Supabase syncWorker failed: ' + e.toString()); }

        return { success: true };
      }
    }
    return { success: false, message: 'Worker ID not found' };
  } else {
    savedId = new Date().getTime().toString();
    workers.appendRow([savedId, ...rowData]);

    // ✅ SUPABASE: insert worker
    try {
      syncWorkerToSupabase({
        gas_id:     savedId,
        name:       name,
        id_card:    form.idCard    || '',
        dob:        dob,
        phone:      form.phone     || '',
        role:       position,
        start_date: startDate || null,
        end_date:   endDate   || null,
        bank_name:  bank,
        bank_acc:   form.bankAcc   || '',
        base_wage:  salary,
        branch:     branch,
        status:     status
      });
    } catch(e) { Logger.log('⚠️ Supabase syncWorker failed: ' + e.toString()); }

    return { success: true };
  }
}

function deleteWorker(id) {
  const { workers } = initDatabase();
  const data = workers.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      workers.deleteRow(i + 1);

      // ✅ SUPABASE: delete worker
      try {
        deleteWorkerFromSupabase(String(id));
      } catch(e) { Logger.log('⚠️ Supabase deleteWorker failed: ' + e.toString()); }

      return { success: true };
    }
  }
  return { success: false };
}

function getActiveWorkersList(branchFilter) {
  return getWorkersData(branchFilter).filter(w => w.status === 'ทำงานอยู่');
}

function transferWorker(id, newBranch) {
  const { workers } = initDatabase();
  const data = workers.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      workers.getRange(i + 1, 13).setValue(newBranch);

      // ✅ SUPABASE: update worker branch
      try {
        // ดึงข้อมูล worker เต็มแล้ว sync ใหม่ทั้งแถว
        const r = data[i];
        syncWorkerToSupabase({
          gas_id:     String(id),
          name:       r[1]  || '',
          id_card:    r[2]  ? String(r[2]).replace(/^'/, '') : '',
          dob:        r[3]  || '',
          phone:      r[4]  ? String(r[4]).replace(/^'/, '') : '',
          role:       r[5]  || '',
          start_date: r[6]  ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null,
          end_date:   r[7]  ? Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null,
          bank_name:  r[8]  || '',
          bank_acc:   r[9]  ? String(r[9]).replace(/^'/, '') : '',
          base_wage:  parseFloat(r[10]) || 0,
          branch:     newBranch,
          status:     r[11] || 'ทำงานอยู่'
        });
      } catch(e) { Logger.log('⚠️ Supabase transferWorker failed: ' + e.toString()); }

      return { success: true, message: `โอนย้ายไปสาขา ${newBranch} เรียบร้อย` };
    }
  }
  return { success: false, message: 'ไม่พบคนงาน' };
}

// ==========================================
// 💰 PAYROLL
// ==========================================
function getPayrollData(branchFilter) {
  const { payroll } = initDatabase();
  const lastRow = payroll.getLastRow();
  if (lastRow <= 1) return [];
  const numCols = Math.max(payroll.getLastColumn(), 12);
  const data = payroll.getRange(2, 1, lastRow - 1, numCols).getValues();
  return data
    .filter(r => {
      if (!branchFilter || branchFilter === 'ทุกสาขา' || branchFilter === '') return true;
      return String(r[11] || 'สำนักงานใหญ่') === branchFilter;
    })
    .map(r => ({
      id:           r[0],
      rawDate:      r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      workerName:   r[2],
      posAllowance: parseFloat(r[3])  || 0,
      otherIncome:  parseFloat(r[4])  || 0,
      grossIncome:  parseFloat(r[5])  || 0,
      taxDeduct:    parseFloat(r[6])  || 0,
      ssoDeduct:    parseFloat(r[7])  || 0,
      otherDeduct:  parseFloat(r[8])  || 0,
      netPay:       parseFloat(r[9])  || 0,
      note:         r[10] || '',
      branch:       r[11] || 'สำนักงานใหญ่'
    })).sort((a, b) => b.id > a.id ? 1 : -1);
}

function savePayroll(form) {
  const { payroll } = initDatabase();
  const branch  = form.branch || 'สำนักงานใหญ่';
  const rowData = [
    form.date, form.workerName, form.posAllowance, form.otherIncome,
    form.grossIncome, form.taxDeduct, form.ssoDeduct, form.otherDeduct,
    form.netPay, form.note, branch
  ];

  let savedId = form.id;

  if (form.id) {
    const data = payroll.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(form.id)) {
        payroll.getRange(i + 1, 2, 1, 11).setValues([rowData]);

        // ✅ SUPABASE: update payroll
        try {
          syncPayrollToSupabase({
            gas_id:        String(form.id),
            pay_date:      form.date         || null,
            worker_name:   form.workerName   || '',
            branch:        branch,
            pos_allowance: form.posAllowance || 0,
            other_income:  form.otherIncome  || 0,
            gross_income:  form.grossIncome  || 0,
            tax_deduct:    form.taxDeduct    || 0,
            sso_deduct:    form.ssoDeduct    || 0,
            other_deduct:  form.otherDeduct  || 0,
            net_pay:       form.netPay       || 0,
            note:          form.note         || ''
          });
        } catch(e) { Logger.log('⚠️ Supabase syncPayroll failed: ' + e.toString()); }

        return { success: true };
      }
    }
  }

  savedId = 'PR-' + new Date().getTime();
  payroll.appendRow([savedId, ...rowData]);

  // ✅ SUPABASE: insert payroll
  try {
    syncPayrollToSupabase({
      gas_id:        savedId,
      pay_date:      form.date         || null,
      worker_name:   form.workerName   || '',
      branch:        branch,
      pos_allowance: form.posAllowance || 0,
      other_income:  form.otherIncome  || 0,
      gross_income:  form.grossIncome  || 0,
      tax_deduct:    form.taxDeduct    || 0,
      sso_deduct:    form.ssoDeduct    || 0,
      other_deduct:  form.otherDeduct  || 0,
      net_pay:       form.netPay       || 0,
      note:          form.note         || ''
    });
  } catch(e) { Logger.log('⚠️ Supabase syncPayroll failed: ' + e.toString()); }

  return { success: true };
}

function deletePayroll(id) {
  const { payroll } = initDatabase();
  const data = payroll.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      payroll.deleteRow(i + 1);

      // ✅ SUPABASE: delete payroll
      try {
        deletePayrollFromSupabase(String(id));
      } catch(e) { Logger.log('⚠️ Supabase deletePayroll failed: ' + e.toString()); }

      return { success: true };
    }
  }
  return { success: false };
}

function getPayslipData(payrollId) {
  const payrolls = getPayrollData();
  const pr = payrolls.find(p => p.id === payrollId);
  if (!pr) return null;
  const workers = getWorkersData();
  const worker  = workers.find(w => w.name === pr.workerName);
  return { payroll: pr, worker: worker || {} };
}

function getPayrollSummaryByBranch(monthStr) {
  const payrolls = getPayrollData();
  const result = {};
  payrolls.forEach(p => {
    const branch = p.branch || 'สำนักงานใหญ่';
    if (!result[branch]) result[branch] = { totalNet: 0, totalGross: 0, count: 0 };
    if (monthStr) {
      const [tY, tM] = monthStr.split('-');
      let rowY = '', rowM = '';
      if (p.rawDate) {
        if (p.rawDate.includes('/')) { const pts = p.rawDate.split('/'); rowY = pts[2]; rowM = pts[1]; }
        else if (p.rawDate.includes('-')) { const pts = p.rawDate.split('-'); rowY = pts[0]; rowM = pts[1]; }
      }
      if (rowY !== tY || rowM !== tM) return;
    }
    result[branch].totalNet   += parseFloat(p.netPay      || 0);
    result[branch].totalGross += parseFloat(p.grossIncome  || 0);
    result[branch].count++;
  });
  return Object.keys(result).map(b => ({ branch: b, ...result[b] }));
}

// ==========================================
// 🔍 MATERIAL EVALUATION
// ==========================================
function evaluateMaterial(itemName, workDays) {
  const supplies = getSuppliesData();
  const dispatches = supplies
    .filter(s => s.itemName === itemName && s.actionType === 'จ่ายออก')
    .sort((a, b) => b.rawDate.localeCompare(a.rawDate));

  if (dispatches.length === 0)
    return { success: false, message: 'ไม่พบประวัติการจ่ายวัสดุชิ้นนี้มาก่อน' };

  const last      = dispatches[0];
  const lastDate  = new Date(last.rawDate);
  const today     = new Date();
  const daysDiff  = Math.max(1, Math.floor((today - lastDate) / (1000 * 60 * 60 * 24)));
  const unitPrice = last.unitPrice > 0 ? last.unitPrice :
    (() => { const inv = getInventoryStatus().find(x => x.name === itemName); return inv ? inv.totalValue / Math.max(1, inv.balance) : 0; })();

  const burnRate      = last.itemQuantity / daysDiff;
  const estimatedQty  = burnRate * workDays;
  const estimatedCost = estimatedQty * unitPrice;
  const monthlyCost   = burnRate * 30 * unitPrice;

  return {
    success: true, lastDate: last.requestDate, lastQty: last.itemQuantity,
    daysSinceLast: daysDiff, burnRateDaily: burnRate.toFixed(2),
    estimatedQty: estimatedQty.toFixed(2), estimatedCost: estimatedCost.toFixed(2),
    monthlyCost: monthlyCost.toFixed(2)
  };
}

// ==========================================
// 🖼️ IMAGE UPLOAD
// ==========================================
function uploadImageToDrive(base64Data, fileName) {
  try {
    const folderName = "NovaSol_Images";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    if (!base64Data.includes("base64,")) return "";
    const splitBase   = base64Data.split(',');
    const contentType = splitBase[0].split(':')[1].split(';')[0];
    const decoded = Utilities.base64Decode(splitBase[1]);
    const blob    = Utilities.newBlob(decoded, contentType, fileName);
    const file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
  } catch (e) {
    Logger.log("Upload Error: " + e.toString());
    throw new Error("Cannot upload image: " + e.message);
  }
}

// ==========================================
// 🛠️ UTILITIES
// ==========================================
function formatDate(d) {
  try {
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (e) { return ""; }
}
