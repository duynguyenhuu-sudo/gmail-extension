// service_worker.js - GITS Mailer Pro v2 (MV3, module)
import { buildRawEmail, cleanSignatureHtml } from './utils/mime.js';

const CLIENT_ID = '385178155043-j6t92gukuvh47oihp3ml794h9rhuqkhg.apps.googleusercontent.com';
const SCOPES = [
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/spreadsheets.readonly',
  'https://www.googleapis.com/auth/gmail.settings.basic'
];

// ========== OAUTH (giữ nguyên code cũ) ==========
async function getAccessTokenInteractive() {
  const redirectUri = chrome.identity.getRedirectURL('oauth2');
  const authUrl =
    'https://accounts.google.com/o/oauth2/v2/auth' +
    `?client_id=${encodeURIComponent(CLIENT_ID)}` +
    `&response_type=token` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&scope=${encodeURIComponent(SCOPES.join(' '))}` +
    `&prompt=consent&access_type=online&include_granted_scopes=true`;

  return new Promise((resolve, reject) => {
    chrome.identity.launchWebAuthFlow({ url: authUrl, interactive: true }, async (redirectedTo) => {
      if (chrome.runtime.lastError) return reject(chrome.runtime.lastError);
      if (!redirectedTo) return reject(new Error('No redirect URL'));
      try {
        const hash = new URL(redirectedTo).hash.substring(1);
        const params = new URLSearchParams(hash);
        const accessToken = params.get('access_token');
        const expiresIn = parseInt(params.get('expires_in') || '0', 10);
        if (params.get('error')) return reject(new Error(`OAuth error: ${params.get('error')}`));
        if (!accessToken) return reject(new Error('No access token'));
        const expiry = Date.now() + (expiresIn - 60) * 1000;
        await chrome.storage.session.set({ accessToken, expiry });
        resolve(accessToken);
      } catch (e) { reject(e); }
    });
  });
}

async function getAccessTokenSimple() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) return reject(chrome.runtime.lastError);
      if (!token) return reject(new Error('No token received'));
      resolve(token);
    });
  });
}

async function ensureToken() {
  const { accessToken, expiry } = await chrome.storage.session.get(['accessToken', 'expiry']);
  if (accessToken && expiry && Date.now() < expiry) return accessToken;
  try {
    const token = await getAccessTokenSimple();
    const expiryTime = Date.now() + 3600 * 1000;
    await chrome.storage.session.set({ accessToken: token, expiry: expiryTime });
    return token;
  } catch (e) {
    console.warn('Simple OAuth failed, trying interactive...', e);
    return getAccessTokenInteractive();
  }
}

// ========== GMAIL API (giữ nguyên code cũ) ==========
async function gmailSendRaw(raw) {
  const token = await ensureToken();
  const res = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ raw })
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(`Gmail send failed: ${res.status} ${res.statusText} — ${JSON.stringify(err)}`);
  }
  return res.json();
}

async function getDefaultSignature() {
  try {
    const token = await ensureToken();
    const res = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs', {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) throw new Error(`Failed fetch sendAs: ${res.status}`);
    const data = await res.json();
    const primary = data.sendAs?.find(sa => sa.isPrimary);
    return cleanSignatureHtml(primary?.signature || '');
  } catch (e) {
    console.error("getDefaultSignature error:", e);
    return '';
  }
}


// ========== WORKER STATE ==========
let workerRunning = false;
let workerTimeout = null;
let delayConfig = { mode: 'random', min: 15000, max: 45000 };

function getDelay() {
  if (delayConfig.mode === 'fixed') return delayConfig.delay || 20000;
  const min = delayConfig.min || 15000;
  const max = delayConfig.max || 45000;
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// ========== WORKER ==========
async function startWorker(config) {
  // Stop any existing worker first
  if (workerRunning) {
    console.log('⚠️ Worker already running, stopping first...');
    stopWorker();
    await new Promise(r => setTimeout(r, 100)); // Wait a bit
  }
  
  delayConfig = config || delayConfig;
  workerRunning = true;
  console.log('🚀 Worker started with config:', delayConfig);
  processNextJob();
  return { ok: true };
}

function stopWorker() {
  workerRunning = false;
  if (workerTimeout) { 
    clearTimeout(workerTimeout); 
    workerTimeout = null; 
  }
  updateWorkerStatus({ isRunning: false });
  console.log('⏹️ Worker stopped');
  return { ok: true };
}

async function processNextJob() {
  if (!workerRunning) return;

  const { jobQueue = [] } = await chrome.storage.local.get('jobQueue');
  const pendingJobs = jobQueue.filter(j => j.status === 'PENDING');
  
  if (pendingJobs.length === 0) {
    console.log('✅ No more pending jobs');
    workerRunning = false;
    updateWorkerStatus({ isRunning: false, lastLog: '✅ All jobs completed' });
    return;
  }

  const job = pendingJobs[0];
  const total = jobQueue.length;
  const sent = jobQueue.filter(j => j.status === 'DONE').length;
  const failed = jobQueue.filter(j => j.status === 'FAILED').length;

  console.log(`📧 Processing: ${job.customer.Email}`);
  await updateWorkerStatus({ total, sent, failed, isRunning: true, nextSendIn: 0, lastLog: `📧 Sending to ${job.customer.Email}...` });

  try {
    // Ensure we have auth token first
    console.log('🔐 Getting auth token...');
    const token = await ensureToken();
    if (!token) throw new Error('No auth token');
    console.log('✅ Token OK');

    const signature = await getDefaultSignature();
    console.log('📝 Signature loaded');
    
    // Convert body to HTML
    const htmlBody = job.body.replace(/\r?\n/g, '<br>');
    const fullBody = `${htmlBody}<br><br>${signature}`;

    console.log('📧 Building email for:', job.customer.Email);
    console.log('📧 Subject:', job.subject);

    // Build raw email
    const raw = await buildRawEmail({
      to: job.customer.Email,
      subject: job.subject,
      body: fullBody,
      attachments: job.attachments || []
    });

    console.log('📤 Sending via Gmail API...');
    await gmailSendRaw(raw);

    job.status = 'DONE';
    job.sentAt = new Date().toISOString();
    
    await addSendLog({
      date: job.sentAt,
      Company_Name: job.customer.Company_Name,
      Customer_Name: job.customer.Customer_Name,
      Email: job.customer.Email,
      Customer_Domain: job.customer.Customer_Domain,
      status: 'Success'
    });

    console.log(`✅ Sent to ${job.customer.Email}`);
    await updateWorkerStatus({ total, sent: sent + 1, failed, isRunning: true, lastLog: `✅ Sent to ${job.customer.Email}` });

  } catch (err) {
    console.error(`❌ Failed: ${job.customer.Email}`, err);
    job.status = 'FAILED';
    job.error = err.message;

    await addSendLog({
      date: new Date().toISOString(),
      Company_Name: job.customer.Company_Name,
      Customer_Name: job.customer.Customer_Name,
      Email: job.customer.Email,
      Customer_Domain: job.customer.Customer_Domain,
      status: 'Failed',
      error: err.message
    });

    await updateWorkerStatus({ total, sent, failed: failed + 1, isRunning: true, lastLog: `❌ Failed: ${job.customer.Email} - ${err.message}` });
  }

  // Save updated queue
  await chrome.storage.local.set({ jobQueue });
  await updateDailyCount();

  // Schedule next job
  if (workerRunning) {
    const delay = getDelay();
    console.log(`⏳ Next in ${delay}ms`);
    await updateWorkerStatus({ nextSendIn: delay });
    workerTimeout = setTimeout(processNextJob, delay);
  }
}


// ========== HELPERS ==========
async function updateWorkerStatus(update) {
  const { workerStatus = {} } = await chrome.storage.local.get('workerStatus');
  await chrome.storage.local.set({ workerStatus: { ...workerStatus, ...update } });
}

async function addSendLog(log) {
  const { sendLogs = [] } = await chrome.storage.local.get('sendLogs');
  sendLogs.push(log);
  await chrome.storage.local.set({ sendLogs });
}

async function updateDailyCount() {
  const today = new Date().toISOString().split('T')[0];
  const { sentToday = [], sentDate } = await chrome.storage.local.get(['sentToday', 'sentDate']);
  if (sentDate !== today) {
    await chrome.storage.local.set({ sentToday: [1], sentDate: today });
  } else {
    sentToday.push(1);
    await chrome.storage.local.set({ sentToday });
  }
}

// ========== MESSAGE LISTENER ==========
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    try {
      console.log('📨 Message:', msg.type);

      switch (msg.type) {
        case 'AUTH': {
          const token = await ensureToken();
          sendResponse({ ok: !!token, token });
          break;
        }
        case 'START_WORKER': {
          const result = await startWorker(msg.delayConfig);
          sendResponse(result);
          break;
        }
        case 'STOP_WORKER': {
          sendResponse(stopWorker());
          break;
        }
        case 'GET_WORKER_STATUS': {
          const { workerStatus } = await chrome.storage.local.get('workerStatus');
          sendResponse({ ok: true, status: workerStatus });
          break;
        }
        case 'CLEAR_QUEUE': {
          await chrome.storage.local.set({ jobQueue: [] });
          sendResponse({ ok: true });
          break;
        }
        default:
          sendResponse({ ok: false, error: 'Unknown message type' });
      }
    } catch (err) {
      console.error('💥 Error:', err);
      sendResponse({ ok: false, error: err.message });
    }
  })();
  return true;
});

// Keep service worker alive
chrome.alarms.create('keepAlive', { periodInMinutes: 0.4 });
chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name === 'keepAlive') {
    console.log('💓 Heartbeat');
    
    // Check if worker should be running but stopped
    const { workerStatus, jobQueue = [] } = await chrome.storage.local.get(['workerStatus', 'jobQueue']);
    const pendingJobs = jobQueue.filter(j => j.status === 'PENDING');
    
    if (workerStatus?.isRunning && !workerRunning && pendingJobs.length > 0) {
      console.log('🔄 Resuming worker...');
      workerRunning = true;
      processNextJob();
    }
  }
});

// Resume worker on startup if needed
chrome.runtime.onStartup.addListener(async () => {
  console.log('🚀 Extension started');
  const { workerStatus, jobQueue = [] } = await chrome.storage.local.get(['workerStatus', 'jobQueue']);
  const pendingJobs = jobQueue.filter(j => j.status === 'PENDING');
  
  if (pendingJobs.length > 0 && workerStatus?.isRunning) {
    console.log(`📋 Resuming ${pendingJobs.length} pending jobs`);
    workerRunning = true;
    processNextJob();
  }
});

// Also check on install/update
chrome.runtime.onInstalled.addListener(async () => {
  console.log('📦 Extension installed/updated');
  chrome.alarms.create('keepAlive', { periodInMinutes: 0.4 });
});
