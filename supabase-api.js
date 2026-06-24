// =====================================================================
// supabase-api.js — NovaSol ERP backend bridge (Supabase edition)
//
// This file REPLACES the old `runGoogle()` function and the GAS_URL
// fetch logic in index.html. Function names/arguments match the old
// Apps Script actions 1:1, so the rest of your UI code (openWorkerModal,
// loadData, etc.) does NOT need to change.
//
// SETUP:
// 1. In index.html <head>, add (before this script):
//      <script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
// 2. Fill in your project URL + anon key below.
// 3. Replace the old <script> block that defines GAS_URL and runGoogle()
//    with: <script src="supabase-api.js"></script>
// =====================================================================

const SUPABASE_URL = 'https://fivytiaqjetnpeatszqm.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_0aiDU-AkVPyGG-SJ72biyg_un6oQJnD';

const sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const fmtDate = d => d ? new Date(d).toLocaleDateString('th-TH', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\d{4}$/, y => y) : '';
function toDDMMYYYY(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return '';
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${d.getFullYear()}`;
}
function toYYYYMMDD(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return '';
  return d.toISOString().split('T')[0];
}

// =====================================================================
// MAIN ENTRY POINT — same signature as before: runGoogle(funcName, ...args)
// =====================================================================
async function runGoogle(funcName, ...args) {
  try {
    const handler = ACTIONS[funcName];
    if (!handler) throw new Error('Unknown action: ' + funcName);
    return await handler(...args);
  } catch (err) {
    console.error('Supabase action error:', funcName, err);
    return { success: false, error: err.message || String(err) };
  }
}

const ACTIONS = {

  // ---------------- AUTH ----------------
  async loginSystem(username, password) {
    const { data, error } = await sb.from('app_users')
      .select('*').eq('username', username).eq('password', password).maybeSingle();
    if (error || !data) return { success: false, message: 'Login Failed' };
    return { success: true, name: data.name, role: data.role, branch: data.branch || 'ทุกสาขา' };
  },

  // ---------------- USERS ----------------
  async getUsersList() {
    const { data, error } = await sb.from('app_users').select('*').order('username');
    if (error) throw error;
    return (data || []).map(u => ({
      row: u.username, username: u.username, password: u.password,
      name: u.name, role: u.role, branch: u.branch || 'ทุกสาขา'
    }));
  },
  async saveUser(form) {
    const payload = {
      username: form.user || form.username,
      password: form.pass || form.password,
      name: form.name, role: form.role, branch: form.branch || 'ทุกสาขา'
    };
    const { error } = await sb.from('app_users').upsert(payload, { onConflict: 'username' });
    if (error) throw error;
    return { success: true };
  },
  async deleteUser(row) {
    const { error } = await sb.from('app_users').delete().eq('username', row);
    if (error) throw error;
    return { success: true };
  },

  // ---------------- BRANCHES ----------------
  async getBranchesData() {
    const { data, error } = await sb.from('branches').select('*').order('created_at');
    if (error) throw error;
    return (data || []).map(b => ({
      id: b.id, name: b.name, address: b.address || '', manager: b.manager || '',
      phone: b.phone || '', openDate: toDDMMYYYY(b.open_date), rawOpen: b.open_date || '',
      status: b.status || 'เปิดใช้งาน'
    }));
  },
  async saveBranch(form) {
    if (!form.name) return { success: false, message: 'กรุณาระบุชื่อสาขา' };
    const payload = {
      name: form.name, address: form.address || '', manager: form.manager || '',
      phone: form.phone || '', open_date: form.openDate || null, status: form.status || 'เปิดใช้งาน'
    };
    if (form.id) {
      const { error } = await sb.from('branches').update(payload).eq('id', form.id);
      if (error) throw error;
    } else {
      const { error } = await sb.from('branches').insert(payload);
      if (error) throw error;
    }
    return { success: true };
  },
  async deleteBranch(id) {
    if (id === 'BR-001') return { success: false, message: 'ไม่สามารถลบสำนักงานใหญ่ได้' };
    const { error } = await sb.from('branches').delete().eq('id', id);
    if (error) return { success: false, message: error.message };
    return { success: true };
  },
  async getBranchSummary(branchName) {
    const workers = await ACTIONS.getWorkersData(branchName);
    const payrolls = await ACTIONS.getPayrollData(branchName);
    const activeWorkers = workers.filter(w => w.status === 'ทำงานอยู่').length;
    let totalPayroll = 0;
    payrolls.forEach(p => totalPayroll += parseFloat(p.netPay || 0));
    return { branchName, totalWorkers: workers.length, activeWorkers, totalPayroll };
  },

  // ---------------- REQUESTS ----------------
  async getRequestsData() {
    const { data, error } = await sb.from('requests').select('*').order('id', { ascending: false });
    if (error) throw error;
    return (data || []).map(r => ({
      id: r.id, itemName: r.item_name, qty: r.qty, location: r.location, requester: r.requester,
      reqDate: toDDMMYYYY(r.req_date), status: r.status, approvedDate: toDDMMYYYY(r.approved_date)
    }));
  },
  async saveRequest(form) {
    if (form.id) {
      const { error } = await sb.from('requests')
        .update({ item_name: form.itemName, qty: form.qty, location: form.location })
        .eq('id', form.id);
      if (error) throw error;
      return { success: true, message: 'Updated' };
    }
    const { error } = await sb.from('requests').insert({
      item_name: form.itemName, qty: form.qty, location: form.location,
      requester: form.requester, status: 'รออนุมัติ'
    });
    if (error) throw error;
    return { success: true, message: 'Created' };
  },
  async updateRequestStatus(id, status, approver) {
    const payload = { status };
    if (status === 'อนุมัติ/จ่ายแล้ว') payload.approved_date = new Date().toISOString();
    const { error } = await sb.from('requests').update(payload).eq('id', id);
    if (error) throw error;
    return { success: true };
  },
  async deleteRequest(id) {
    const { error } = await sb.from('requests').delete().eq('id', id);
    if (error) return { success: false, message: error.message };
    return { success: true };
  },

  // ---------------- SUPPLIES ----------------
  async getMaterialList() {
    const { data, error } = await sb.from('supplies').select('item_name').eq('action_type', 'รับเข้า');
    if (error) throw error;
    return Array.from(new Set((data || []).map(d => d.item_name))).sort();
  },
  async getSuppliesData() {
    const { data, error } = await sb.from('supplies').select('*').order('request_date', { ascending: false });
    if (error) throw error;
    return (data || []).map(row => ({
      id: row.id, itemName: row.item_name, imageUrl: row.image_url || '',
      unitPrice: parseFloat(row.unit_price) || 0, itemQuantity: parseFloat(row.item_quantity) || 0,
      actionType: row.action_type, deliveryLocation: row.location,
      requestDate: toDDMMYYYY(row.request_date), requester: row.requester,
      rawDate: row.request_date || ''
    }));
  },
  async saveSupply(form) {
    let finalImgUrl = (form.imageUrl || '').trim();
    if (form.imageBase64 && form.imageBase64.includes('base64,')) {
      try {
        finalImgUrl = await uploadImageToSupabase(form.imageBase64, form.imageName || ('image_' + Date.now()));
      } catch (e) { console.error('Image upload failed', e); }
    }

    let unitPrice = parseFloat(form.unitPrice) || 0;
    if (form.actionType === 'จ่ายออก' && unitPrice === 0) {
      const status = await ACTIONS.getInventoryStatus();
      const item = status.find(x => x.name === form.itemName);
      if (item && item.countIn > 0) {
        unitPrice = item.totalValue / item.balance;
        if (!isFinite(unitPrice)) unitPrice = 0;
      }
    }

    const payload = {
      item_name: form.itemName, image_url: finalImgUrl, unit_price: unitPrice,
      item_quantity: form.itemQuantity, action_type: form.actionType,
      location: form.deliveryLocation, request_date: form.requestDate, requester: form.requester
    };

    if (form.id) {
      const { error } = await sb.from('supplies').update(payload).eq('id', form.id);
      if (error) throw error;
    } else {
      const { error } = await sb.from('supplies').insert(payload);
      if (error) throw error;
    }
    return { success: true };
  },
  async deleteSupply(id) {
    const { error } = await sb.from('supplies').delete().eq('id', id);
    if (error) throw error;
    return { success: true };
  },

  // ---------------- CONTRACTS ----------------
  async getContractsData() {
    const { data, error } = await sb.from('contracts').select('*').order('id');
    if (error) throw error;
    return (data || []).map(row => ({
      id: row.id, name: row.name, startDate: toDDMMYYYY(row.start_date), endDate: toDDMMYYYY(row.end_date),
      totalValue: parseFloat(row.total_value) || 0, laborBudget: parseFloat(row.labor_budget) || 0,
      status: row.status, rawStart: row.start_date || '', rawEnd: row.end_date || ''
    }));
  },
  async saveContract(form) {
    const start = form.startDate, end = form.endDate;
    const today = new Date();
    const status = today < new Date(start) ? 'Pending' : (today > new Date(end) ? 'Expired' : 'Active');
    const payload = { name: form.name, start_date: start, end_date: end, total_value: form.totalBudget, labor_budget: form.laborBudget, status };
    if (form.id) {
      const { error } = await sb.from('contracts').update(payload).eq('id', form.id);
      if (error) throw error;
    } else {
      const { error } = await sb.from('contracts').insert(payload);
      if (error) throw error;
    }
    return { success: true };
  },
  async deleteContract(id) {
    const { error } = await sb.from('contracts').delete().eq('id', id);
    if (error) throw error;
    return { success: true };
  },

  // ---------------- LABOR COSTS ----------------
  async getLaborData() {
    const { data, error } = await sb.from('labor_costs').select('*').order('id', { ascending: false });
    if (error) throw error;
    return (data || []).map(r => ({
      id: r.id, project: r.project, date: toDDMMYYYY(r.pay_date), amount: parseFloat(r.amount) || 0,
      note: r.note, recorder: r.recorder, rawDate: r.pay_date || ''
    }));
  },
  async saveLabor(form) {
    const payload = { project: form.project, pay_date: form.date, amount: form.amount, note: form.note, recorder: form.recorder };
    if (form.id) {
      const { error } = await sb.from('labor_costs').update(payload).eq('id', form.id);
      if (error) throw error;
    } else {
      const { error } = await sb.from('labor_costs').insert(payload);
      if (error) throw error;
    }
    return { success: true };
  },
  async deleteLabor(id) {
    const { error } = await sb.from('labor_costs').delete().eq('id', id);
    if (error) throw error;
    return { success: true };
  },

  // ---------------- INVENTORY (derived) ----------------
  async getInventoryStatus() {
    const supplies = await ACTIONS.getSuppliesData();
    const inventory = {};
    supplies.forEach(item => {
      const name = (item.itemName || '').toString();
      if (name.includes('เพิ่มทุน') || name.includes('ค่างวด') || name.includes('รายรับ')) return;
      if (!inventory[name]) inventory[name] = { in: 0, out: 0, totalPriceIn: 0, countIn: 0, lastImage: item.imageUrl };
      if (item.imageUrl) inventory[name].lastImage = item.imageUrl;
      if (item.actionType === 'รับเข้า') {
        inventory[name].in += item.itemQuantity;
        inventory[name].totalPriceIn += item.unitPrice * item.itemQuantity;
        inventory[name].countIn += item.itemQuantity;
      } else if (item.actionType === 'จ่ายออก') {
        inventory[name].out += item.itemQuantity;
      }
    });
    return Object.keys(inventory).map(name => {
      const item = inventory[name];
      const balance = item.in - item.out;
      const avgPrice = item.countIn > 0 ? item.totalPriceIn / item.countIn : 0;
      return {
        name, in: item.in, out: item.out, balance, totalValue: balance * avgPrice,
        imageUrl: item.lastImage, countIn: item.countIn,
        status: balance <= 0 ? 'สินค้าหมด' : (balance < 5 ? 'ใกล้หมด' : 'ปกติ')
      };
    });
  },

  // ---------------- WORKERS ----------------
  async getWorkersData(branchFilter) {
    let q = sb.from('workers').select('*');
    if (branchFilter && branchFilter !== 'ทุกสาขา') q = q.eq('branch', branchFilter);
    const { data, error } = await q;
    if (error) throw error;
    return (data || []).map(r => ({
      id: r.id, name: r.name, idCard: r.id_card || '', dob: r.dob || '', phone: r.phone || '',
      role: r.role, startDate: toDDMMYYYY(r.start_date), endDate: toDDMMYYYY(r.end_date),
      bankName: r.bank_name, bankAcc: r.bank_account, baseWage: parseFloat(r.base_wage) || 0,
      status: r.status || 'ทำงานอยู่', branch: r.branch || 'สำนักงานใหญ่',
      rawStart: r.start_date || '', rawEnd: r.end_date || ''
    }));
  },
  async saveWorker(form) {
    const payload = {
      name: form.name, id_card: form.idCard || '', dob: form.dob || '', phone: form.phone || '',
      role: form.role || form.position || '', start_date: form.startDate || null, end_date: form.endDate || null,
      bank_name: form.bankName || form.bank || '', bank_account: form.bankAcc || form.account || '',
      base_wage: parseFloat(form.baseWage || form.salary || 0), status: form.status || 'ทำงานอยู่',
      branch: form.branch || 'สำนักงานใหญ่'
    };
    if (form.id) {
      const { error } = await sb.from('workers').update(payload).eq('id', form.id);
      if (error) throw error;
    } else {
      payload.id = Date.now().toString();
      const { error } = await sb.from('workers').insert(payload);
      if (error) throw error;
    }
    return { success: true };
  },
  async deleteWorker(id) {
    const { error } = await sb.from('workers').delete().eq('id', id);
    if (error) throw error;
    return { success: true };
  },
  async getActiveWorkersList(branchFilter) {
    const list = await ACTIONS.getWorkersData(branchFilter);
    return list.filter(w => w.status === 'ทำงานอยู่');
  },
  async transferWorker(id, newBranch) {
    const { error } = await sb.from('workers').update({ branch: newBranch }).eq('id', id);
    if (error) return { success: false, message: error.message };
    return { success: true, message: `โอนย้ายไปสาขา ${newBranch} เรียบร้อย` };
  },

  // ---------------- PAYROLL ----------------
  async getPayrollData(branchFilter) {
    let q = sb.from('payroll').select('*');
    if (branchFilter && branchFilter !== 'ทุกสาขา') q = q.eq('branch', branchFilter);
    const { data, error } = await q.order('id', { ascending: false });
    if (error) throw error;
    return (data || []).map(r => ({
      id: r.id, rawDate: toDDMMYYYY(r.pay_date), workerName: r.worker_name,
      posAllowance: parseFloat(r.pos_allowance) || 0, otherIncome: parseFloat(r.other_income) || 0,
      grossIncome: parseFloat(r.gross_income) || 0, taxDeduct: parseFloat(r.tax_deduct) || 0,
      ssoDeduct: parseFloat(r.sso_deduct) || 0, otherDeduct: parseFloat(r.other_deduct) || 0,
      netPay: parseFloat(r.net_pay) || 0, note: r.note || '', branch: r.branch || 'สำนักงานใหญ่'
    }));
  },
  async savePayroll(form) {
    const payload = {
      pay_date: form.date, worker_name: form.workerName, pos_allowance: form.posAllowance,
      other_income: form.otherIncome, gross_income: form.grossIncome, tax_deduct: form.taxDeduct,
      sso_deduct: form.ssoDeduct, other_deduct: form.otherDeduct, net_pay: form.netPay,
      note: form.note, branch: form.branch || 'สำนักงานใหญ่'
    };
    if (form.id) {
      const { error } = await sb.from('payroll').update(payload).eq('id', form.id);
      if (error) throw error;
    } else {
      payload.id = 'PR-' + Date.now();
      const { error } = await sb.from('payroll').insert(payload);
      if (error) throw error;
    }
    return { success: true };
  },
  async deletePayroll(id) {
    const { error } = await sb.from('payroll').delete().eq('id', id);
    if (error) throw error;
    return { success: true };
  },
  async getPayslipData(payrollId) {
    const { data: pr, error: e1 } = await sb.from('payroll').select('*').eq('id', payrollId).maybeSingle();
    if (e1 || !pr) return null;
    const { data: w } = await sb.from('workers').select('*').eq('name', pr.worker_name).maybeSingle();
    return {
      payroll: {
        id: pr.id, rawDate: toDDMMYYYY(pr.pay_date), workerName: pr.worker_name,
        posAllowance: pr.pos_allowance, otherIncome: pr.other_income, grossIncome: pr.gross_income,
        taxDeduct: pr.tax_deduct, ssoDeduct: pr.sso_deduct, otherDeduct: pr.other_deduct,
        netPay: pr.net_pay, note: pr.note, branch: pr.branch
      },
      worker: w ? { role: w.role, bankName: w.bank_name, bankAcc: w.bank_account, baseWage: w.base_wage } : {}
    };
  },
  async getPayrollSummaryByBranch(monthStr) {
    const { data, error } = await sb.from('payroll').select('*');
    if (error) throw error;
    const result = {};
    (data || []).forEach(p => {
      const branch = p.branch || 'สำนักงานใหญ่';
      if (!result[branch]) result[branch] = { totalNet: 0, totalGross: 0, count: 0 };
      if (monthStr) {
        const m = p.pay_date ? p.pay_date.slice(0, 7) : '';
        if (m !== monthStr) return;
      }
      result[branch].totalNet += parseFloat(p.net_pay || 0);
      result[branch].totalGross += parseFloat(p.gross_income || 0);
      result[branch].count++;
    });
    return Object.keys(result).map(b => ({ branch: b, ...result[b] }));
  },

  // ---------------- CAPITAL ----------------
  async getCapitalData() {
    const { data, error } = await sb.from('capital_injection').select('*').order('date', { ascending: false });
    if (error) throw error;
    return (data || []).map(row => ({
      id: row.id, date: toDDMMYYYY(row.date), rawDate: row.date || '',
      amount: parseFloat(row.amount) || 0, source: row.source || '',
      project: row.project || '', note: row.note || '', recorder: row.recorder || ''
    }));
  },
  async saveCapital(form) {
    if (!form.date || !form.amount || !form.source) return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
    const amount = parseFloat(form.amount);
    if (isNaN(amount) || amount <= 0) return { success: false, message: 'จำนวนเงินต้องมากกว่า 0' };
    const payload = { date: form.date, amount, source: form.source, project: form.project || '', note: form.note || '', recorder: form.recorder || '' };
    if (form.id) {
      const { error } = await sb.from('capital_injection').update(payload).eq('id', form.id);
      if (error) throw error;
    } else {
      payload.id = 'CAP-' + Date.now();
      const { error } = await sb.from('capital_injection').insert(payload);
      if (error) throw error;
    }
    return { success: true };
  },
  async deleteCapital(id) {
    const { error } = await sb.from('capital_injection').delete().eq('id', id);
    if (error) return { success: false, message: error.message };
    return { success: true };
  },
  async getCapitalSummary() {
    const capitals = await ACTIONS.getCapitalData();
    const bySource = {};
    let total = 0;
    capitals.forEach(c => { bySource[c.source] = (bySource[c.source] || 0) + c.amount; total += c.amount; });
    return { total, bySource: Object.keys(bySource).map(k => ({ source: k, amount: bySource[k] })), count: capitals.length };
  },

  // ---------------- DASHBOARD / REPORTS ----------------
  async getDashboardStats() {
    const supplies = await ACTIONS.getSuppliesData();
    const contracts = await ACTIONS.getContractsData();
    const laborData = await ACTIONS.getLaborData();
    const materialCostByProj = {}, laborCostByProj = {};

    supplies.forEach(s => {
      const name = (s.itemName || '').toString();
      if (name.includes('เพิ่มทุน') || name.includes('ค่างวด') || name.includes('รายรับ')) return;
      if (s.actionType === 'จ่ายออก') {
        materialCostByProj[s.deliveryLocation] = (materialCostByProj[s.deliveryLocation] || 0) + (s.unitPrice * s.itemQuantity);
      }
    });
    laborData.forEach(l => { laborCostByProj[l.project] = (laborCostByProj[l.project] || 0) + l.amount; });

    let totalRevenue = 0, totalExpense = 0;
    const labels = [], budgets = [], expenses = [], projects = [];
    contracts.forEach(c => {
      const matCost = materialCostByProj[c.name] || 0;
      const laborCost = laborCostByProj[c.name] || 0;
      const totalCost = matCost + laborCost;
      totalRevenue += c.totalValue; totalExpense += totalCost;
      labels.push(c.name); budgets.push(c.totalValue); expenses.push(totalCost);
      projects.push({
        name: c.name, budget: c.totalValue, expense: totalCost, profit: c.totalValue - totalCost,
        isLoss: (c.totalValue - totalCost) < 0,
        costProgress: c.totalValue > 0 ? (totalCost / c.totalValue * 100) : 0
      });
    });
    return { revenue: totalRevenue, expense: totalExpense, labels, budgets, expenses, projects };
  },
  async getExpenseReport(monthStr) {
    if (!monthStr) return [];
    const supplies = await ACTIONS.getSuppliesData();
    return supplies
      .filter(i => i.actionType === 'จ่ายออก' && i.rawDate && i.rawDate.startsWith(monthStr))
      .map(i => ({ date: i.requestDate, item: i.itemName, loc: i.deliveryLocation, amount: i.itemQuantity * i.unitPrice }))
      .sort((a, b) => a.date.localeCompare(b.date));
  },
  async getProjectExpenseReport(monthStr) {
    if (!monthStr) return [];
    const supplies = await ACTIONS.getSuppliesData();
    const projectExpenses = {};
    supplies.forEach(item => {
      if (!item.rawDate || !item.rawDate.startsWith(monthStr) || item.actionType !== 'จ่ายออก') return;
      const proj = item.deliveryLocation;
      if (!projectExpenses[proj]) projectExpenses[proj] = { total: 0, items: [] };
      projectExpenses[proj].total += item.itemQuantity * item.unitPrice;
      projectExpenses[proj].items.push({ name: item.itemName, qty: item.itemQuantity, price: item.unitPrice, cost: item.itemQuantity * item.unitPrice, date: item.requestDate });
    });
    return Object.keys(projectExpenses).map(proj => ({ projectName: proj, totalExpense: projectExpenses[proj].total, details: projectExpenses[proj].items }));
  },
  async evaluateMaterial(itemName, workDays) {
    const supplies = await ACTIONS.getSuppliesData();
    const dispatches = supplies.filter(s => s.itemName === itemName && s.actionType === 'จ่ายออก')
      .sort((a, b) => b.rawDate.localeCompare(a.rawDate));
    if (dispatches.length === 0) return { success: false, message: 'ไม่พบประวัติการจ่ายวัสดุชิ้นนี้มาก่อน' };

    const last = dispatches[0];
    const lastDate = new Date(last.rawDate);
    const daysDiff = Math.max(1, Math.floor((new Date() - lastDate) / 86400000));
    let unitPrice = last.unitPrice;
    if (!unitPrice) {
      const inv = (await ACTIONS.getInventoryStatus()).find(x => x.name === itemName);
      unitPrice = inv ? inv.totalValue / Math.max(1, inv.balance) : 0;
    }
    const burnRate = last.itemQuantity / daysDiff;
    const estimatedQty = burnRate * workDays;
    const estimatedCost = estimatedQty * unitPrice;
    const monthlyCost = burnRate * 30 * unitPrice;
    return {
      success: true, lastDate: last.requestDate, lastQty: last.itemQuantity, daysSinceLast: daysDiff,
      burnRateDaily: burnRate.toFixed(2), estimatedQty: estimatedQty.toFixed(2),
      estimatedCost: estimatedCost.toFixed(2), monthlyCost: monthlyCost.toFixed(2)
    };
  },
};

// =====================================================================
// IMAGE UPLOAD — replaces uploadImageToDrive() (now uses Supabase Storage)
// =====================================================================
async function uploadImageToSupabase(base64Data, fileName) {
  const [meta, b64] = base64Data.split(',');
  const contentType = meta.split(':')[1].split(';')[0];
  const byteChars = atob(b64);
  const bytes = new Uint8Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) bytes[i] = byteChars.charCodeAt(i);

  const path = `${Date.now()}_${fileName}`;
  const { error } = await sb.storage.from('novasol-images').upload(path, bytes, { contentType, upsert: true });
  if (error) throw error;
  const { data } = sb.storage.from('novasol-images').getPublicUrl(path);
  return data.publicUrl;
}