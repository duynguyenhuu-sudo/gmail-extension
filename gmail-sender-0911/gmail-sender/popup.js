// popup.js - GITS Mailer Pro v2
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ========== DATA STORES ==========
let companyInsight = [];  // File 1: {Company_Name, Customer_Name, Email, Customer_Domain}
let domainMaster = {};    // File 2: {Domain: {Title_Mail, CaseStudies[]}}
let mailTemplate = '';    // File 3: Email template
let attachments = [];     // Attachments
let isDataValidated = false;

const DAILY_LIMIT = 100;
const MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024;

const DEFAULT_TEMPLATE = `初めまして。どうぞ良い新しい一週間が元気いっぱいでありますように。
GITS 株式会社の営業部のアインと申します。

開発実績(一部):
{{CaseStudy_List}}

何卒よろしくお願い申し上げます。`;

// ========== INITIALIZATION ==========
document.addEventListener('DOMContentLoaded', async () => {
  $('#bodyTemplate').value = DEFAULT_TEMPLATE;
  $('#logDate').valueAsDate = new Date();
  
  await loadStats();
  setupTabs();
  setupFileHandlers();
  setupEventListeners();
  await loadLogs();
});

// ========== TAB NAVIGATION ==========
function setupTabs() {
  $$('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      $$('.tab').forEach(t => t.classList.remove('active'));
      $$('.tab-content').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      $(`#tab-${tab.dataset.tab}`).classList.add('active');
    });
  });
}

// ========== FILE HANDLERS ==========
function setupFileHandlers() {
  // File 1: Company Insight
  $('#fileCustomer').addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const label = $('label[for="fileCustomer"]');
    label.textContent = '⏳ Reading...';
    try {
      const data = await readExcelFile(file);
      companyInsight = parseCompanyInsight(data);
      renderCompanyPreview();
      label.textContent = `✅ ${companyInsight.length} customers`;
      isDataValidated = false;
    } catch (err) {
      label.textContent = '❌ Error';
      showStatus('#validateStatus', err.message, 'err');
    }
  });

  // File 2: Domain Master
  $('#fileDomain').addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const label = $('label[for="fileDomain"]');
    label.textContent = '⏳ Reading...';
    try {
      const data = await readExcelFile(file);
      domainMaster = parseDomainMaster(data);
      renderDomainPreview();
      label.textContent = `✅ ${Object.keys(domainMaster).length} domains`;
      isDataValidated = false;
    } catch (err) {
      label.textContent = '❌ Error';
      showStatus('#validateStatus', err.message, 'err');
    }
  });

  // File 3: Mail Template
  $('#fileTemplate').addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      mailTemplate = await file.text();
      $('#bodyTemplate').value = mailTemplate;
      $('label[for="fileTemplate"]').textContent = `✅ ${file.name}`;
    } catch (err) {
      showStatus('#validateStatus', err.message, 'err');
    }
  });

  // Attachments
  $('#fileAttach').addEventListener('change', handleAttachments);
}


// ========== EVENT LISTENERS ==========
function setupEventListeners() {
  // Validate button
  $('#btnValidate').addEventListener('click', validateAndLoadData);

  // Preview customer select
  $('#previewCustomer').addEventListener('change', renderEmailPreview);

  // Confirm preview
  $('#btnConfirmPreview').addEventListener('click', () => {
    if (!isDataValidated) return showStatus('#validateStatus', 'Please validate data first', 'err');
    switchTab('send');
  });

  // Auth
  $('#btnAuth').addEventListener('click', async () => {
    showStatus('#authStatus', 'Logging in...', 'ok');
    const { ok, error } = await sendMsg({ type: 'AUTH' });
    showStatus('#authStatus', ok ? '✅ Ready!' : `❌ ${error}`, ok ? 'ok' : 'err');
  });

  // Delay mode
  $('#delayMode').addEventListener('change', () => {
    const isRandom = $('#delayMode').value === 'random';
    $('#randomDelayWrapper').style.display = isRandom ? 'block' : 'none';
    $('#fixedDelayWrapper').style.display = isRandom ? 'none' : 'block';
  });

  // Queue & Send
  $('#btnAddToQueue').addEventListener('click', addToQueue);
  $('#btnStartWorker').addEventListener('click', startWorker);
  $('#btnStopWorker').addEventListener('click', stopWorker);

  // Logs
  $('#btnFilterLogs').addEventListener('click', loadLogs);
  $('#btnExportToday').addEventListener('click', () => exportLogs('today'));
  $('#btnExportFiltered').addEventListener('click', () => exportLogs('filtered'));
  $('#btnClearOldLogs').addEventListener('click', clearOldLogs);
}

function switchTab(tabName) {
  $$('.tab').forEach(t => t.classList.remove('active'));
  $$('.tab-content').forEach(c => c.classList.remove('active'));
  $(`.tab[data-tab="${tabName}"]`).classList.add('active');
  $(`#tab-${tabName}`).classList.add('active');
}

// ========== VALIDATION ==========
async function validateAndLoadData() {
  const errors = [];
  const warnings = [];

  // Check File 1
  if (!companyInsight.length) {
    errors.push('❌ File 1 (Company Insight) is required');
  } else {
    // Validate emails
    const invalidEmails = companyInsight.filter(c => !c.Email || !c.Email.includes('@'));
    if (invalidEmails.length) {
      errors.push(`❌ ${invalidEmails.length} invalid emails found`);
    }

    // Validate domains exist in Domain Master
    if (Object.keys(domainMaster).length) {
      const unknownDomains = new Set();
      companyInsight.forEach(c => {
        const domains = c.Customer_Domain.split(/[,\s]+/).filter(d => d);
        domains.forEach(d => {
          if (!domainMaster[d.toUpperCase()]) unknownDomains.add(d);
        });
      });
      if (unknownDomains.size) {
        warnings.push(`⚠️ Unknown domains: ${[...unknownDomains].join(', ')}`);
      }
    }
  }

  // Check File 2
  if (!Object.keys(domainMaster).length) {
    errors.push('❌ File 2 (Domain Master) is required');
  }

  // Get template
  mailTemplate = $('#bodyTemplate').value || DEFAULT_TEMPLATE;

  // Show results
  if (errors.length) {
    showStatus('#validateStatus', errors.join('<br>'), 'err');
    isDataValidated = false;
    return;
  }

  if (warnings.length) {
    showStatus('#validateStatus', warnings.join('<br>') + '<br>✅ Data loaded with warnings', 'warn');
  } else {
    showStatus('#validateStatus', '✅ All data validated successfully!', 'ok');
  }

  isDataValidated = true;

  // Populate preview
  populatePreviewCustomers();
  renderCustomerList();
  switchTab('preview');
}


// ========== PREVIEW ==========
function populatePreviewCustomers() {
  const select = $('#previewCustomer');
  select.innerHTML = '<option value="">-- Select customer to preview --</option>';
  companyInsight.forEach((c, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `${c.Company_Name} - ${c.Customer_Name} (${c.Email})`;
    select.appendChild(opt);
  });
}

function renderEmailPreview() {
  const idx = $('#previewCustomer').value;
  if (idx === '') {
    $('#previewDomain').value = '';
    $('#previewSubject').textContent = '';
    $('#previewBody').textContent = '';
    return;
  }

  const customer = companyInsight[parseInt(idx)];
  const { subject, body } = buildEmailContent(customer);

  $('#previewDomain').value = customer.Customer_Domain || '(none)';
  $('#previewSubject').textContent = subject;
  $('#previewBody').textContent = body;
}

function renderCustomerList() {
  const list = $('#customerList');
  list.innerHTML = companyInsight.map((c, i) => 
    `<div style="padding:4px 0; border-bottom:1px solid #eee;">
      ${i + 1}. ${c.Company_Name} | ${c.Customer_Name} | ${c.Email} | <span class="pill">${c.Customer_Domain || 'N/A'}</span>
    </div>`
  ).join('');
  $('#totalCustomers').textContent = companyInsight.length;
}

// ========== EMAIL CONTENT BUILDER ==========
function buildEmailContent(customer) {
  const { Company_Name, Customer_Name, Customer_Domain } = customer;
  const { caseStudies, title } = buildCaseStudyList(Customer_Domain);

  // Build case study text
  let caseStudyText = '';
  if (caseStudies.length > 0) {
    caseStudyText = caseStudies.map(cs => `・${cs}`).join('\n') + ' など';
  }

  // Build header
  const nameWithSama = Customer_Name ? `${Customer_Name} 様` : '様';
  const header = Company_Name ? `${Company_Name}\n${nameWithSama}` : nameWithSama;

  // Subject from domain
  const subject = title || '【GITS 株式会社】アプリ・システム開発で貴社の DX 推進を支援いたします';

  // Replace placeholders in template
  let body = (mailTemplate || DEFAULT_TEMPLATE)
    .replace(/\{\{Company_Name\}\}/g, Company_Name || '')
    .replace(/\{\{Customer_Name\}\}/g, Customer_Name || '')
    .replace(/\{\{Title_Mail\}\}/g, title || '')
    .replace(/\{\{CaseStudy_List\}\}/g, caseStudyText);

  body = `${header}\n\n${body}`;

  return { subject, body };
}

function buildCaseStudyList(customerDomains) {
  if (!customerDomains) return { caseStudies: [], title: '' };
  
  const domains = customerDomains.split(/[,\s]+/).filter(d => d);
  if (domains.length === 0) return { caseStudies: [], title: '' };

  let caseStudies = [];
  let title = '';

  if (domains.length === 1) {
    // Single domain: 4 case studies
    const domain = domains[0].toUpperCase();
    const mapping = domainMaster[domain];
    if (mapping) {
      title = mapping.Title_Mail || '';
      caseStudies = (mapping.CaseStudies || []).slice(0, 4);
    }
  } else {
    // Multiple domains: 2 case studies per domain
    const firstDomain = domains[0].toUpperCase();
    title = domainMaster[firstDomain]?.Title_Mail || '';

    domains.forEach(d => {
      const mapping = domainMaster[d.toUpperCase()];
      if (mapping && mapping.CaseStudies) {
        caseStudies.push(...mapping.CaseStudies.slice(0, 2));
      }
    });
  }

  return { caseStudies, title };
}


// ========== QUEUE & WORKER ==========
async function addToQueue() {
  if (!isDataValidated) return showStatus('#sendStatus', '❌ Please validate data first', 'err');
  if (!companyInsight.length) return showStatus('#sendStatus', '❌ No customers to send', 'err');

  // Check daily limit
  const stats = await chrome.storage.local.get(['sentToday', 'sentDate']);
  const today = new Date().toISOString().split('T')[0];
  const sentCount = (stats.sentDate === today && Array.isArray(stats.sentToday)) ? stats.sentToday.length : 0;
  const remaining = DAILY_LIMIT - sentCount;

  if (remaining <= 0) {
    return showStatus('#sendStatus', '❌ Daily limit reached (100 emails)', 'err');
  }

  const toQueue = companyInsight.slice(0, remaining);
  if (toQueue.length < companyInsight.length) {
    showStatus('#sendStatus', `⚠️ Only ${toQueue.length}/${companyInsight.length} will be queued (daily limit)`, 'warn');
  }

  // Build jobs
  const jobs = toQueue.map((customer, i) => {
    const { subject, body } = buildEmailContent(customer);
    return {
      id: Date.now() + i,
      customer,
      subject,
      body,
      attachments: attachments.map(a => ({ name: a.name, mimeType: a.mimeType, base64: a.base64, size: a.size })),
      status: 'PENDING',
      createdAt: new Date().toISOString()
    };
  });

  // Save to queue
  const { jobQueue = [] } = await chrome.storage.local.get('jobQueue');
  await chrome.storage.local.set({ jobQueue: [...jobQueue, ...jobs] });

  showStatus('#sendStatus', `✅ Added ${jobs.length} jobs to queue`, 'ok');
  await updateQueueCount();
}

async function startWorker() {
  const delayMode = $('#delayMode').value;
  const delayConfig = delayMode === 'random' 
    ? { mode: 'random', min: parseInt($('#delayMin').value) * 1000, max: parseInt($('#delayMax').value) * 1000 }
    : { mode: 'fixed', delay: parseInt($('#fixedDelay').value) * 1000 };

  const { ok, error } = await sendMsg({ type: 'START_WORKER', delayConfig });
  
  if (ok) {
    showStatus('#sendStatus', '▶️ Worker started', 'ok');
    startProgressPolling();
  } else {
    showStatus('#sendStatus', `❌ ${error}`, 'err');
  }
}

async function stopWorker() {
  const { ok } = await sendMsg({ type: 'STOP_WORKER' });
  showStatus('#sendStatus', ok ? '⏹️ Worker stopped' : '❌ Failed to stop', ok ? 'ok' : 'err');
  stopProgressPolling();
}

let progressInterval = null;
function startProgressPolling() {
  if (progressInterval) return;
  progressInterval = setInterval(updateProgress, 1000);
}

function stopProgressPolling() {
  if (progressInterval) {
    clearInterval(progressInterval);
    progressInterval = null;
  }
}

async function updateProgress() {
  const { workerStatus } = await chrome.storage.local.get('workerStatus');
  if (!workerStatus) return;

  const { total, sent, failed, nextSendIn, isRunning, lastLog } = workerStatus;
  const done = sent + failed;
  const percent = total > 0 ? (done / total * 100) : 0;

  $('#progressBar').style.width = `${percent}%`;
  $('#progressText').textContent = `${done}/${total} (${sent} sent, ${failed} failed)`;
  $('#nextSendIn').textContent = isRunning ? Math.max(0, Math.ceil(nextSendIn / 1000)) : '--';

  if (lastLog) {
    appendLog(lastLog);
  }

  await loadStats();
  await updateQueueCount();

  if (!isRunning && done >= total) {
    stopProgressPolling();
    showStatus('#sendStatus', '✅ All jobs completed!', 'ok');
  }
}

function appendLog(msg) {
  const log = $('#liveLog');
  const time = new Date().toLocaleTimeString();
  log.innerHTML += `[${time}] ${msg}\n`;
  log.scrollTop = log.scrollHeight;
}


// ========== LOGS ==========
async function loadLogs() {
  const dateFilter = $('#logDate').value;
  const statusFilter = $('#logStatus').value;

  const { sendLogs = [] } = await chrome.storage.local.get('sendLogs');
  
  let filtered = sendLogs;
  if (dateFilter) {
    filtered = filtered.filter(l => l.date && l.date.startsWith(dateFilter));
  }
  if (statusFilter) {
    filtered = filtered.filter(l => l.status === statusFilter);
  }

  // Summary
  const success = filtered.filter(l => l.status === 'Success').length;
  const failed = filtered.filter(l => l.status === 'Failed').length;
  $('#logTotal').textContent = filtered.length;
  $('#logSuccess').textContent = success;
  $('#logFailed').textContent = failed;

  // Table
  const tbody = $('#logTableBody');
  tbody.innerHTML = filtered.map(l => `
    <tr>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.date ? new Date(l.date).toLocaleString('ja-JP') : ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.Company_Name || ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.Customer_Name || ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.Email || ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee;">
        <span style="color:${l.status === 'Success' ? '#27ae60' : l.status === 'Failed' ? '#e74c3c' : '#f39c12'};">
          ${l.status || 'Pending'}
        </span>
      </td>
    </tr>
  `).join('');
}

async function exportLogs(type) {
  const { sendLogs = [] } = await chrome.storage.local.get('sendLogs');
  const today = new Date().toISOString().split('T')[0];
  
  let data = sendLogs;
  if (type === 'today') {
    data = sendLogs.filter(l => l.date && l.date.startsWith(today));
  } else if (type === 'filtered') {
    const dateFilter = $('#logDate').value;
    const statusFilter = $('#logStatus').value;
    if (dateFilter) data = data.filter(l => l.date && l.date.startsWith(dateFilter));
    if (statusFilter) data = data.filter(l => l.status === statusFilter);
  }

  if (!data.length) return alert('No data to export');

  // Group by domain for summary
  const domainStats = {};
  data.forEach(l => {
    const domain = l.Customer_Domain || 'Unknown';
    if (!domainStats[domain]) domainStats[domain] = { total: 0, success: 0, failed: 0 };
    domainStats[domain].total++;
    if (l.status === 'Success') domainStats[domain].success++;
    if (l.status === 'Failed') domainStats[domain].failed++;
  });

  const exportData = [
    ['GITS Mailer - Send Report'],
    ['Export Date:', new Date().toLocaleString('ja-JP')],
    ['Total Sent:', data.length],
    [],
    ['Domain Summary:'],
    ['Domain', 'Total', 'Success', 'Failed'],
    ...Object.entries(domainStats).map(([d, s]) => [d, s.total, s.success, s.failed]),
    [],
    ['Detail:'],
    ['Date', 'Company Name', 'Customer Name', 'Email', 'Domain', 'Status', 'Error'],
    ...data.map(l => [
      l.date ? new Date(l.date).toLocaleString('ja-JP') : '',
      l.Company_Name || '', l.Customer_Name || '', l.Email || '',
      l.Customer_Domain || '', l.status || '', l.error || ''
    ])
  ];

  const ws = XLSX.utils.aoa_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Send Report');
  XLSX.writeFile(wb, `GITS_SendReport_${today}.xlsx`);
}

async function clearOldLogs() {
  if (!confirm('Delete logs older than 7 days?')) return;
  
  const { sendLogs = [] } = await chrome.storage.local.get('sendLogs');
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
  const filtered = sendLogs.filter(l => l.date && new Date(l.date).getTime() > cutoff);
  
  await chrome.storage.local.set({ sendLogs: filtered });
  await loadLogs();
  alert(`Deleted ${sendLogs.length - filtered.length} old logs`);
}


// ========== ATTACHMENTS ==========
async function handleAttachments(e) {
  const files = Array.from(e.target.files || []);
  if (!files.length) return;
  
  if (files.length > 3) alert('Only first 3 files will be kept');
  const picked = files.slice(0, 3);
  
  const totalSize = picked.reduce((sum, f) => sum + f.size, 0);
  if (totalSize > MAX_ATTACHMENT_SIZE) {
    return alert(`Total size exceeds 25MB (${(totalSize/1024/1024).toFixed(1)}MB)`);
  }

  const label = $('label[for="fileAttach"]');
  label.textContent = '⏳ Processing...';

  try {
    attachments = await Promise.all(picked.map(async f => ({
      name: f.name,
      mimeType: f.type || 'application/octet-stream',
      base64: await fileToBase64(f),
      size: f.size
    })));
    
    renderAttachments();
    label.textContent = `✅ ${attachments.length} files`;
    e.target.value = '';
  } catch (err) {
    label.textContent = '❌ Error';
    alert(err.message);
  }
}

function renderAttachments() {
  const container = $('#attachedFiles');
  if (!attachments.length) {
    container.innerHTML = '';
    $('#attachInfo').textContent = '0 files.';
    return;
  }

  const totalMB = (attachments.reduce((s, f) => s + f.size, 0) / 1024 / 1024).toFixed(1);
  $('#attachInfo').innerHTML = `<strong>${attachments.length} files</strong> (${totalMB} MB)`;

  container.innerHTML = attachments.map((f, i) => `
    <div style="display:flex; align-items:center; gap:10px; padding:8px; background:#f8f9fa; border-radius:4px; margin-top:5px;">
      <span>📎 ${f.name} (${(f.size/1024).toFixed(1)} KB)</span>
      <button onclick="removeAttachment(${i})" style="background:#e74c3c; color:white; border:none; border-radius:50%; width:20px; height:20px; cursor:pointer;">×</button>
    </div>
  `).join('');
}

window.removeAttachment = (i) => {
  attachments.splice(i, 1);
  renderAttachments();
};

// ========== STATS ==========
async function loadStats() {
  const today = new Date().toISOString().split('T')[0];
  const { sentToday = [], sentDate } = await chrome.storage.local.get(['sentToday', 'sentDate']);
  
  const count = sentDate === today ? sentToday.length : 0;
  $('#dailyCount').textContent = count;
  $('#limitWarning').style.display = count >= DAILY_LIMIT ? 'block' : 'none';
}

async function updateQueueCount() {
  const { jobQueue = [] } = await chrome.storage.local.get('jobQueue');
  const pending = jobQueue.filter(j => j.status === 'PENDING').length;
  $('#queueCount').textContent = pending;
}


// ========== PARSERS ==========
async function readExcelFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
}

function parseCompanyInsight(data) {
  if (!data.length) return [];
  const headers = data[0].map(h => (h || '').toString().trim().toLowerCase());
  
  const cIdx = findColumnIndex(headers, ['company_name', 'company', '会社名']);
  const nIdx = findColumnIndex(headers, ['customer_name', 'name', '名前']);
  const eIdx = findColumnIndex(headers, ['email', 'mail', 'メール']);
  const dIdx = findColumnIndex(headers, ['customer_domain', 'domain']);

  return data.slice(1)
    .filter(row => row && row.length > Math.max(cIdx, nIdx, eIdx))
    .map(row => ({
      Company_Name: (row[cIdx] || '').toString().trim(),
      Customer_Name: (row[nIdx] || '').toString().trim(),
      Email: (row[eIdx] || '').toString().trim(),
      Customer_Domain: (row[dIdx] || '').toString().trim().toUpperCase()
    }))
    .filter(r => r.Email && r.Email.includes('@'));
}

function parseDomainMaster(data) {
  if (!data.length) return {};
  const headers = data[0].map(h => (h || '').toString().trim().toLowerCase());
  
  const dIdx = findColumnIndex(headers, ['domain', 'ドメイン']);
  const tIdx = findColumnIndex(headers, ['title_mail', 'title']);
  
  // Find CaseStudy columns
  const csIndices = [];
  headers.forEach((h, i) => {
    if (/casestudy|case_study|cs_/i.test(h)) csIndices.push(i);
  });

  const mapping = {};
  data.slice(1).forEach(row => {
    if (!row || !row[dIdx]) return;
    const domain = row[dIdx].toString().trim().toUpperCase();
    const caseStudies = csIndices
      .map(i => (row[i] || '').toString().trim())
      .filter(cs => cs && !/^#?null$/i.test(cs));
    
    mapping[domain] = {
      Title_Mail: (row[tIdx] || '').toString().trim(),
      CaseStudies: caseStudies
    };
  });
  return mapping;
}

function findColumnIndex(headers, names) {
  const idx = headers.findIndex(h => names.includes(h));
  return idx >= 0 ? idx : names.indexOf(names[0]);
}

// ========== RENDERERS ==========
function renderCompanyPreview() {
  const el = $('#customerPreview');
  if (!companyInsight.length) {
    el.textContent = 'No data.';
    el.classList.add('muted');
    return;
  }
  el.classList.remove('muted');
  el.textContent = companyInsight.slice(0, 5)
    .map(c => `${c.Company_Name} | ${c.Customer_Name} | ${c.Email} | ${c.Customer_Domain}`)
    .join('\n') + (companyInsight.length > 5 ? `\n... (+${companyInsight.length - 5} more)` : '');
}

function renderDomainPreview() {
  const el = $('#domainPreview');
  const domains = Object.keys(domainMaster);
  if (!domains.length) {
    el.textContent = 'No data.';
    el.classList.add('muted');
    return;
  }
  el.classList.remove('muted');
  el.textContent = domains.slice(0, 5)
    .map(d => `${d}: ${domainMaster[d].Title_Mail} | ${domainMaster[d].CaseStudies.length} CS`)
    .join('\n') + (domains.length > 5 ? `\n... (+${domains.length - 5} more)` : '');
}

// ========== HELPERS ==========
function showStatus(selector, msg, type) {
  const el = $(selector);
  el.innerHTML = msg;
  el.className = 'small ' + (type || '');
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function sendMsg(msg) {
  return new Promise(resolve => chrome.runtime.sendMessage(msg, resolve));
}
