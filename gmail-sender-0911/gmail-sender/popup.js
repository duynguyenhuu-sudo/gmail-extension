// popup.js - GITS Mailer Pro v2
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ========== DATA STORES ==========
let companyInsight = [];  // File 1
let domainMaster = {};    // File 2
let mailTemplate = '';    // File 3
let attachments = [];
let isDataValidated = false;

const DAILY_LIMIT = 100;
const MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024;

const DEFAULT_TEMPLATE = `初めまして。どうぞ良い新しい一週間が元気いっぱいでありますように。
GITS株式会社の営業部のアインと申します。

御社のウェブサイトを拝見し、ITを活用した事業展開に深く感銘を受けました。
弊社はベトナムを拠点に、日本子会社を通じて日系企業様向けにアプリ・システム開発や業務DXを支援しており、貴社の今後のIT推進にお役立てできるのではと考え、ご連絡差し上げました。

現時点では、新しいアプリ、システム開発、ウェブサイト構築や業務ソリューション、IT推進、DX導入をご検討されていますでしょうか。

【GITS株式会社について】
ベトナム本社・日本拠点を持ち、製造・物流・医療・IoTなど幅広い分野で開発実績を積んでおります。

主なサービス：
・ITコンサルティング／システム・アプリ開発
・保守・運用サポート
・ITエンジニア派遣（オンサイト・ラボ型対応）

開発実績（一部）：
{{CaseStudy_List}}

GITSの特徴：
・CMMI 2.0 Level 3準拠の品質管理体制
・日本語堪能なPMとコミュニケーターによる円滑な対応
・コスト最適化と柔軟な開発体制（小規模～大規模案件まで対応）
・情報セキュリティ管理（ISO準拠）

以上です。参考資料として会社案内と開発の流れを添付しております。
もしご関心をお持ちいただけましたら、簡単にオンラインでご紹介させていただきます。
また、日本在留のGITS部長が直接お伺いしての対面でも、ご希望に合わせて調整させていただきます。

貴社のDX推進やIT課題の解決に、GITSを選択の一つを考慮していただけますと幸いです。
これまでお時間いただきありがとうございます。

何卒よろしくお願い申し上げます。`;

// ========== INITIALIZATION ==========
document.addEventListener('DOMContentLoaded', async () => {
  await loadSavedData(); // Load saved data first
  $('#logDate').valueAsDate = new Date();
  await loadStats();
  setupTabs();
  setupFileHandlers();
  setupEventListeners();
  await loadLogs();
  await updateQueueCount();
  
  // Check if worker is running
  const { workerStatus } = await chrome.storage.local.get('workerStatus');
  if (workerStatus?.isRunning) {
    showStatus('#sendStatus', '▶️ Worker đang chạy nền...', 'ok');
    startProgressPolling();
  }
  
  // Open in new tab button
  $('#btnOpenTab')?.addEventListener('click', () => {
    chrome.tabs.create({ url: chrome.runtime.getURL('popup.html') });
    window.close();
  });
});

// ========== SAVE/LOAD DATA ==========
async function saveDataToStorage() {
  await chrome.storage.local.set({
    savedCompanyInsight: companyInsight,
    savedDomainMaster: domainMaster,
    savedMailTemplate: mailTemplate,
    savedAttachments: attachments,
    savedIsValidated: isDataValidated
  });
  console.log('💾 Data saved to storage');
}

async function loadSavedData() {
  const data = await chrome.storage.local.get([
    'savedCompanyInsight', 
    'savedDomainMaster', 
    'savedMailTemplate', 
    'savedAttachments',
    'savedIsValidated'
  ]);
  
  if (data.savedCompanyInsight?.length) {
    companyInsight = data.savedCompanyInsight;
    renderCompanyPreview();
    $('label[for="fileCustomer"]').textContent = `✅ ${companyInsight.length} customers (saved)`;
  }
  
  if (data.savedDomainMaster && Object.keys(data.savedDomainMaster).length) {
    domainMaster = data.savedDomainMaster;
    renderDomainPreview();
    $('label[for="fileDomain"]').textContent = `✅ ${Object.keys(domainMaster).length} domains (saved)`;
  }
  
  if (data.savedMailTemplate) {
    mailTemplate = data.savedMailTemplate;
    $('#bodyTemplate').value = mailTemplate;
  } else {
    $('#bodyTemplate').value = DEFAULT_TEMPLATE;
  }
  
  if (data.savedAttachments?.length) {
    attachments = data.savedAttachments;
    renderAttachments();
  }
  
  if (data.savedIsValidated) {
    isDataValidated = true;
    populatePreviewCustomers();
    renderCustomerList();
  }
  
  console.log('📂 Data loaded from storage');
}


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

function switchTab(tabName) {
  $$('.tab').forEach(t => t.classList.remove('active'));
  $$('.tab-content').forEach(c => c.classList.remove('active'));
  $(`.tab[data-tab="${tabName}"]`).classList.add('active');
  $(`#tab-${tabName}`).classList.add('active');
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
      await saveDataToStorage();
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
      await saveDataToStorage();
    } catch (err) {
      label.textContent = '❌ Error';
      showStatus('#validateStatus', err.message, 'err');
    }
  });

  // File 3: Mail Template (Word/TXT/HTML)
  $('#fileTemplate').addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const label = $('label[for="fileTemplate"]');
    const ext = file.name.split('.').pop().toLowerCase();
    label.textContent = '⏳ Reading...';
    
    try {
      if (ext === 'docx') {
        // DOCX file - need JSZip
        if (typeof JSZip === 'undefined') {
          throw new Error('JSZip not loaded. Please use .txt file instead.');
        }
        const arrayBuffer = await file.arrayBuffer();
        const text = await extractTextFromDocx(arrayBuffer);
        if (text && text.trim()) {
          mailTemplate = text;
          $('#bodyTemplate').value = mailTemplate;
          label.textContent = `✅ ${file.name}`;
          showStatus('#validateStatus', `✅ Template loaded: ${text.length} characters`, 'ok');
          await saveDataToStorage();
        } else {
          throw new Error('Could not extract text from DOCX');
        }
      } else {
        // TXT or HTML - read as text
        const text = await file.text();
        if (text && text.trim()) {
          mailTemplate = text;
          $('#bodyTemplate').value = mailTemplate;
          label.textContent = `✅ ${file.name}`;
          showStatus('#validateStatus', `✅ Template loaded: ${text.length} characters`, 'ok');
          await saveDataToStorage();
        } else {
          throw new Error('Empty file');
        }
      }
    } catch (err) {
      console.error('Template read error:', err);
      label.textContent = `❌ ${file.name}`;
      showStatus('#validateStatus', `❌ Cannot read DOCX. Using existing template. Try .txt file or paste content.`, 'err');
      // Don't change textarea content on error - keep existing template
    }
  });

  // Template textarea change
  $('#bodyTemplate').addEventListener('change', async () => {
    mailTemplate = $('#bodyTemplate').value;
    await saveDataToStorage();
  });

  // Attachments
  $('#fileAttach').addEventListener('change', handleAttachments);
}

// Read Word (.docx) file - extract text from XML
async function extractTextFromDocx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file('word/document.xml')?.async('string');
  
  if (!docXml) {
    throw new Error('Invalid DOCX: no document.xml');
  }
  
  // Parse XML and extract text
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(docXml, 'text/xml');
  
  // Get all paragraphs
  const paragraphs = xmlDoc.getElementsByTagName('w:p');
  let result = [];
  
  for (let p of paragraphs) {
    const texts = p.getElementsByTagName('w:t');
    let paraText = '';
    for (let t of texts) {
      paraText += t.textContent || '';
    }
    if (paraText) {
      result.push(paraText);
    }
  }
  
  const finalText = result.join('\n').trim();
  if (!finalText) {
    throw new Error('No text content found in DOCX');
  }
  
  return finalText;
}


// ========== EVENT LISTENERS ==========
function setupEventListeners() {
  $('#btnValidate').addEventListener('click', validateAndLoadData);
  $('#previewCustomer').addEventListener('change', renderEmailPreview);
  $('#btnConfirmPreview').addEventListener('click', () => {
    if (!isDataValidated) return showStatus('#validateStatus', 'Please validate data first', 'err');
    switchTab('send');
  });

  $('#btnAuth').addEventListener('click', async () => {
    showStatus('#authStatus', 'Logging in...', 'ok');
    const { ok, error } = await sendMsg({ type: 'AUTH' });
    showStatus('#authStatus', ok ? '✅ Ready!' : `❌ ${error}`, ok ? 'ok' : 'err');
  });

  $('#delayMode').addEventListener('change', () => {
    const isRandom = $('#delayMode').value === 'random';
    $('#randomDelayWrapper').style.display = isRandom ? 'block' : 'none';
    $('#fixedDelayWrapper').style.display = isRandom ? 'none' : 'block';
  });

  $('#btnSendNow').addEventListener('click', sendNow);
  $('#btnAddToQueue').addEventListener('click', addToQueue);
  $('#btnStartWorker').addEventListener('click', startWorker);
  $('#btnStopWorker').addEventListener('click', stopWorker);

  $('#btnFilterLogs').addEventListener('click', loadLogs);
  $('#btnExportToday').addEventListener('click', () => exportLogs('today'));
  $('#btnExportFiltered').addEventListener('click', () => exportLogs('filtered'));
  $('#btnClearOldLogs').addEventListener('click', clearOldLogs);
  $('#btnClearAllData').addEventListener('click', clearAllData);
}

// ========== VALIDATION ==========
async function validateAndLoadData() {
  const errors = [];
  const warnings = [];

  if (!companyInsight.length) {
    errors.push('❌ File 1 (Company Insight) is required');
  } else {
    const invalidEmails = companyInsight.filter(c => !c.Email || !c.Email.includes('@'));
    if (invalidEmails.length) errors.push(`❌ ${invalidEmails.length} invalid emails`);

    if (Object.keys(domainMaster).length) {
      const unknownDomains = new Set();
      companyInsight.forEach(c => {
        const domains = c.Customer_Domain.split(/[,\s]+/).filter(d => d);
        domains.forEach(d => { if (!domainMaster[d.toUpperCase()]) unknownDomains.add(d); });
      });
      if (unknownDomains.size) warnings.push(`⚠️ Unknown domains: ${[...unknownDomains].join(', ')}`);
    }
  }

  if (!Object.keys(domainMaster).length) errors.push('❌ File 2 (Domain Master) is required');

  mailTemplate = $('#bodyTemplate').value || DEFAULT_TEMPLATE;

  if (errors.length) {
    showStatus('#validateStatus', errors.join('<br>'), 'err');
    isDataValidated = false;
    return;
  }

  showStatus('#validateStatus', warnings.length 
    ? warnings.join('<br>') + '<br>✅ Data loaded with warnings' 
    : '✅ All data validated!', warnings.length ? 'warn' : 'ok');

  isDataValidated = true;
  await saveDataToStorage();
  populatePreviewCustomers();
  renderCustomerList();
  switchTab('preview');
}


// ========== PREVIEW ==========
function populatePreviewCustomers() {
  const select = $('#previewCustomer');
  select.innerHTML = '<option value="">-- Select customer --</option>';
  companyInsight.forEach((c, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `${c.Company_Name} - ${c.Customer_Name} (${c.Customer_Domain})`;
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
  const { subject, body, caseStudies } = buildEmailContent(customer);

  $('#previewDomain').value = `${customer.Customer_Domain || '(none)'} → ${caseStudies?.length || 0} CaseStudies`;
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

  // Build header: Company + Name + 様
  const nameWithSama = Customer_Name ? `${Customer_Name} 様` : '様';
  const header = Company_Name ? `${Company_Name}\n${nameWithSama}` : nameWithSama;

  // Subject from domain
  const subject = title || '【GITS 株式会社】アプリ・システム開発で貴社の DX 推進を支援いたします';

  // Build case study text with など at end - luôn đủ 4 CS
  let caseStudyText = '';
  if (caseStudies.length > 0) {
    // Loại bỏ dấu ・ ở đầu nếu đã có sẵn trong data
    caseStudyText = caseStudies.map(cs => {
      const cleanCS = cs.replace(/^[・･●▪▸►•\-\*]\s*/, '').trim();
      return `・${cleanCS}`;
    }).join('\n') + ' など';
  }

  // Replace placeholders
  let body = (mailTemplate || DEFAULT_TEMPLATE)
    .replace(/\{\{Company_Name\}\}/g, Company_Name || '')
    .replace(/\{\{Customer_Name\}\}/g, Customer_Name || '')
    .replace(/\{\{Title_Mail\}\}/g, title || '')
    .replace(/\{\{CaseStudy_List\}\}/g, caseStudyText);

  // Support individual CaseStudy placeholders: {{CaseStudy_1}}, {{CaseStudy_2}}, etc.
  for (let i = 0; i < 10; i++) {
    const cs = caseStudies[i] || '';
    body = body.replace(new RegExp(`\\{\\{CaseStudy_${i + 1}\\}\\}`, 'g'), cs);
    body = body.replace(new RegExp(`\\{\\{READ_Casestudy_${i + 1}\\}\\}`, 'gi'), cs);
    body = body.replace(new RegExp(`\\{\\{READ_CaseStudy_${i + 1}\\}\\}`, 'gi'), cs);
  }

  body = `${header}\n\n${body}`;

  return { subject, body, caseStudies }; // Return caseStudies for debugging
}

function buildCaseStudyList(customerDomains) {
  if (!customerDomains) return { caseStudies: [], title: '' };
  
  const domains = customerDomains.split(/[,\s]+/).filter(d => d).map(d => d.toUpperCase());
  if (domains.length === 0) return { caseStudies: [], title: '' };

  const TARGET_CS = 4; // Luôn đủ 4 CaseStudy
  let caseStudies = [];
  let title = '';

  // Lấy title từ domain đầu tiên
  const firstMapping = domainMaster[domains[0]];
  title = firstMapping?.Title_Mail || '';

  if (domains.length === 1) {
    // 1 Domain: Random 4 CS từ danh sách
    const mapping = domainMaster[domains[0]];
    if (mapping && mapping.CaseStudies && mapping.CaseStudies.length > 0) {
      const allCS = [...mapping.CaseStudies];
      const shuffled = shuffleArray(allCS);
      // Nếu không đủ 4, lặp lại để đủ
      while (shuffled.length < TARGET_CS && allCS.length > 0) {
        shuffled.push(...shuffleArray([...allCS]));
      }
      caseStudies = shuffled.slice(0, TARGET_CS);
    }
  } else {
    // 2+ Domains: Mỗi domain ít nhất 1 CS, bổ sung để đủ 4
    const usedCS = new Set();
    const allAvailableCS = []; // Pool để bổ sung

    // Bước 1: Lấy 1 CS random từ mỗi domain
    domains.forEach(d => {
      const mapping = domainMaster[d];
      if (mapping && mapping.CaseStudies && mapping.CaseStudies.length > 0) {
        const domainCS = [...mapping.CaseStudies];
        const shuffled = shuffleArray(domainCS);
        
        // Lấy 1 CS đầu tiên cho domain này
        if (shuffled.length > 0 && caseStudies.length < TARGET_CS) {
          caseStudies.push(shuffled[0]);
          usedCS.add(shuffled[0]);
        }
        
        // Thêm các CS còn lại vào pool
        shuffled.forEach(cs => {
          if (!usedCS.has(cs)) {
            allAvailableCS.push(cs);
          }
        });
      }
    });

    // Bước 2: Bổ sung random từ pool để đủ 4 CS
    if (caseStudies.length < TARGET_CS && allAvailableCS.length > 0) {
      const shuffledPool = shuffleArray(allAvailableCS);
      for (const cs of shuffledPool) {
        if (caseStudies.length >= TARGET_CS) break;
        if (!usedCS.has(cs)) {
          caseStudies.push(cs);
          usedCS.add(cs);
        }
      }
    }

    // Bước 3: Nếu vẫn chưa đủ 4, lặp lại từ pool
    if (caseStudies.length < TARGET_CS && allAvailableCS.length > 0) {
      let idx = 0;
      while (caseStudies.length < TARGET_CS) {
        caseStudies.push(allAvailableCS[idx % allAvailableCS.length]);
        idx++;
      }
    }
  }

  console.log(`📋 Built ${caseStudies.length} CaseStudies for domains [${domains.join(',')}]:`, caseStudies);
  return { caseStudies: caseStudies.slice(0, TARGET_CS), title };
}

// Fisher-Yates shuffle
function shuffleArray(array) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}


// ========== SEND NOW (1 click) ==========
async function sendNow() {
  if (!isDataValidated) return showStatus('#sendStatus', '❌ Please validate data first', 'err');
  if (!companyInsight.length) return showStatus('#sendStatus', '❌ No customers', 'err');

  // Check daily limit
  const stats = await chrome.storage.local.get(['sentToday', 'sentDate']);
  const today = new Date().toISOString().split('T')[0];
  const sentCount = (stats.sentDate === today && Array.isArray(stats.sentToday)) ? stats.sentToday.length : 0;
  const remaining = DAILY_LIMIT - sentCount;

  if (remaining <= 0) return showStatus('#sendStatus', '❌ Daily limit reached (100)', 'err');

  const toSend = companyInsight.slice(0, remaining);
  if (toSend.length < companyInsight.length) {
    showStatus('#sendStatus', `⚠️ Only ${toSend.length}/${companyInsight.length} will be sent (limit)`, 'warn');
  }

  // Build jobs
  const jobs = toSend.map((customer, i) => {
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

  // Clear old queue and add new jobs
  await chrome.storage.local.set({ jobQueue: jobs });

  // Get delay config
  const delayMode = $('#delayMode').value;
  const delayConfig = delayMode === 'random' 
    ? { mode: 'random', min: parseInt($('#delayMin').value) * 1000, max: parseInt($('#delayMax').value) * 1000 }
    : { mode: 'fixed', delay: parseInt($('#fixedDelay').value) * 1000 };

  showStatus('#sendStatus', `🚀 Sending ${jobs.length} emails...`, 'ok');
  $('#liveLog').innerHTML = '';
  await updateQueueCount();

  // Start worker immediately
  const { ok, error } = await sendMsg({ type: 'START_WORKER', delayConfig });
  
  if (ok) {
    startProgressPolling();
  } else {
    showStatus('#sendStatus', `❌ ${error}`, 'err');
  }
}

// ========== QUEUE MODE (2 steps) ==========
async function addToQueue() {
  if (!isDataValidated) return showStatus('#sendStatus', '❌ Please validate data first', 'err');
  if (!companyInsight.length) return showStatus('#sendStatus', '❌ No customers', 'err');

  const stats = await chrome.storage.local.get(['sentToday', 'sentDate']);
  const today = new Date().toISOString().split('T')[0];
  const sentCount = (stats.sentDate === today && Array.isArray(stats.sentToday)) ? stats.sentToday.length : 0;
  const remaining = DAILY_LIMIT - sentCount;

  if (remaining <= 0) return showStatus('#sendStatus', '❌ Daily limit reached (100)', 'err');

  const toQueue = companyInsight.slice(0, remaining);

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

  const { jobQueue = [] } = await chrome.storage.local.get('jobQueue');
  await chrome.storage.local.set({ jobQueue: [...jobQueue, ...jobs] });

  showStatus('#sendStatus', `✅ Added ${jobs.length} to queue. Click "Start Queue" to send.`, 'ok');
  await updateQueueCount();
}

async function startWorker() {
  const delayMode = $('#delayMode').value;
  const delayConfig = delayMode === 'random' 
    ? { mode: 'random', min: parseInt($('#delayMin').value) * 1000, max: parseInt($('#delayMax').value) * 1000 }
    : { mode: 'fixed', delay: parseInt($('#fixedDelay').value) * 1000 };

  const { ok, error } = await sendMsg({ type: 'START_WORKER', delayConfig });
  
  if (ok) {
    showStatus('#sendStatus', '▶️ Queue started', 'ok');
    $('#liveLog').innerHTML = '';
    startProgressPolling();
  } else {
    showStatus('#sendStatus', `❌ ${error}`, 'err');
  }
}

async function stopWorker() {
  const { ok } = await sendMsg({ type: 'STOP_WORKER' });
  showStatus('#sendStatus', ok ? '⏹️ Stopped' : '❌ Failed', ok ? 'ok' : 'err');
  stopProgressPolling();
}

let progressInterval = null;
function startProgressPolling() {
  if (progressInterval) return;
  progressInterval = setInterval(updateProgress, 1000);
}

function stopProgressPolling() {
  if (progressInterval) { clearInterval(progressInterval); progressInterval = null; }
}

async function updateProgress() {
  const { workerStatus } = await chrome.storage.local.get('workerStatus');
  if (!workerStatus) return;

  const { total = 0, sent = 0, failed = 0, nextSendIn = 0, isRunning, lastLog } = workerStatus;
  const done = sent + failed;
  const percent = total > 0 ? (done / total * 100) : 0;

  $('#progressBar').style.width = `${percent}%`;
  $('#progressText').textContent = `${done}/${total} (${sent} ✓, ${failed} ✗)`;
  $('#nextSendIn').textContent = isRunning ? Math.max(0, Math.ceil(nextSendIn / 1000)) : '--';

  if (lastLog) appendLog(lastLog);

  await loadStats();
  await updateQueueCount();

  if (!isRunning && done >= total && total > 0) {
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
  if (dateFilter) filtered = filtered.filter(l => l.date && l.date.startsWith(dateFilter));
  if (statusFilter) filtered = filtered.filter(l => l.status === statusFilter);

  const success = filtered.filter(l => l.status === 'Success').length;
  const failed = filtered.filter(l => l.status === 'Failed').length;
  $('#logTotal').textContent = filtered.length;
  $('#logSuccess').textContent = success;
  $('#logFailed').textContent = failed;

  $('#logTableBody').innerHTML = filtered.map(l => `
    <tr>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.date ? new Date(l.date).toLocaleString('ja-JP') : ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.Company_Name || ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.Customer_Name || ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee;">${l.Email || ''}</td>
      <td style="padding:6px; border-bottom:1px solid #eee; color:${l.status === 'Success' ? '#27ae60' : '#e74c3c'};">${l.status || ''}</td>
    </tr>
  `).join('');
}

async function exportLogs(type) {
  const { sendLogs = [] } = await chrome.storage.local.get('sendLogs');
  const today = new Date().toISOString().split('T')[0];
  
  let data = sendLogs;
  if (type === 'today') data = sendLogs.filter(l => l.date && l.date.startsWith(today));
  else if (type === 'filtered') {
    const dateFilter = $('#logDate').value;
    const statusFilter = $('#logStatus').value;
    if (dateFilter) data = data.filter(l => l.date && l.date.startsWith(dateFilter));
    if (statusFilter) data = data.filter(l => l.status === statusFilter);
  }

  if (!data.length) return alert('No data to export');

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

async function clearAllData() {
  if (!confirm('Clear ALL saved data? (Customers, Domains, Template, Attachments)')) return;
  
  companyInsight = [];
  domainMaster = {};
  mailTemplate = '';
  attachments = [];
  isDataValidated = false;
  
  await chrome.storage.local.remove([
    'savedCompanyInsight', 
    'savedDomainMaster', 
    'savedMailTemplate', 
    'savedAttachments',
    'savedIsValidated',
    'jobQueue'
  ]);
  
  // Reset UI
  $('#bodyTemplate').value = DEFAULT_TEMPLATE;
  renderCompanyPreview();
  renderDomainPreview();
  renderAttachments();
  $('label[for="fileCustomer"]').textContent = '📊 Chọn file Company Insight';
  $('label[for="fileDomain"]').textContent = '📊 Chọn file Domain Master';
  $('label[for="fileTemplate"]').textContent = '📄 Chọn file Template (Word/TXT)';
  
  showStatus('#validateStatus', '✅ All data cleared', 'ok');
}


// ========== ATTACHMENTS ==========
async function handleAttachments(e) {
  const files = Array.from(e.target.files || []);
  if (!files.length) return;
  
  if (files.length > 3) alert('Only first 3 files kept');
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
    await saveDataToStorage();
  } catch (err) {
    label.textContent = '❌ Error';
    alert(err.message);
  }
}

function renderAttachments() {
  const container = $('#attachedFiles');
  if (!attachments.length) { container.innerHTML = ''; $('#attachInfo').textContent = '0 files.'; return; }

  const totalMB = (attachments.reduce((s, f) => s + f.size, 0) / 1024 / 1024).toFixed(1);
  $('#attachInfo').innerHTML = `<strong>${attachments.length} files</strong> (${totalMB} MB)`;

  container.innerHTML = attachments.map((f, i) => `
    <div style="display:flex; align-items:center; gap:10px; padding:8px; background:#f8f9fa; border-radius:4px; margin-top:5px;">
      <span>📎 ${f.name} (${(f.size/1024).toFixed(1)} KB)</span>
      <button onclick="removeAttachment(${i})" style="background:#e74c3c; color:white; border:none; border-radius:50%; width:20px; height:20px; cursor:pointer;">×</button>
    </div>
  `).join('');
}

window.removeAttachment = async (i) => { 
  attachments.splice(i, 1); 
  renderAttachments(); 
  await saveDataToStorage();
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
  $('#queueCount').textContent = jobQueue.filter(j => j.status === 'PENDING').length;
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
  const cIdx = findCol(headers, ['company_name', 'company', '会社名']);
  const nIdx = findCol(headers, ['customer_name', 'name', '名前']);
  const eIdx = findCol(headers, ['email', 'mail', 'メール']);
  const dIdx = findCol(headers, ['customer_domain', 'domain']);

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
  const dIdx = findCol(headers, ['domain', 'ドメイン']);
  const tIdx = findCol(headers, ['title_mail', 'title']);
  const csIndices = [];
  headers.forEach((h, i) => { if (/casestudy|case_study|cs_/i.test(h)) csIndices.push(i); });

  const mapping = {};
  data.slice(1).forEach(row => {
    if (!row || !row[dIdx]) return;
    const domain = row[dIdx].toString().trim().toUpperCase();
    const caseStudies = csIndices.map(i => (row[i] || '').toString().trim()).filter(cs => cs && !/^#?null$/i.test(cs));
    mapping[domain] = { Title_Mail: (row[tIdx] || '').toString().trim(), CaseStudies: caseStudies };
  });
  return mapping;
}

function findCol(headers, names) {
  const idx = headers.findIndex(h => names.includes(h));
  return idx >= 0 ? idx : 0;
}

// ========== RENDERERS ==========
function renderCompanyPreview() {
  const el = $('#customerPreview');
  if (!companyInsight.length) { el.textContent = 'No data.'; el.classList.add('muted'); return; }
  el.classList.remove('muted');
  el.textContent = companyInsight.slice(0, 5).map(c => `${c.Company_Name} | ${c.Customer_Name} | ${c.Email} | ${c.Customer_Domain}`).join('\n') 
    + (companyInsight.length > 5 ? `\n... (+${companyInsight.length - 5})` : '');
}

function renderDomainPreview() {
  const el = $('#domainPreview');
  const domains = Object.keys(domainMaster);
  if (!domains.length) { el.textContent = 'No data.'; el.classList.add('muted'); return; }
  el.classList.remove('muted');
  el.textContent = domains.slice(0, 5).map(d => `${d}: ${domainMaster[d].Title_Mail} | ${domainMaster[d].CaseStudies.length} CS`).join('\n')
    + (domains.length > 5 ? `\n... (+${domains.length - 5})` : '');
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
