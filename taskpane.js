// ─────────────────────────────────────────────
//  DAILA Word Add-in — taskpane.js
//  Paste your Make webhook URLs below
// ─────────────────────────────────────────────

const WEBHOOKS = {
  summary:  'https://hook.eu2.make.com/l2x4komxsl8viga28kjp1tkgt8oj53hv',
  contract: 'https://hook.eu2.make.com/hqsp163dh21x55jhohkries202okqkfn',
  issues:   'https://hook.eu2.make.com/mnoi4ynf3nwiigmq1fnpoj2jfnn2wj5q',
};

const ACTION_LABELS = {
  summary:  'Document summary',
  contract: 'Contract analysis',
  issues:   'Issue list generation',
};

// ─── Office initialisation ───────────────────

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    loadDocumentName();
  }
});

function loadDocumentName() {
  Office.context.document.getFilePropertiesAsync((result) => {
    const nameEl = document.getElementById('docName');
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const url = result.value.url;
      // Extract filename from full path/URL
      const parts = url.split(/[\\/]/);
      nameEl.textContent = parts[parts.length - 1] || 'Active document';
    } else {
      nameEl.textContent = 'Active document';
    }
  });
}

// ─── Core action runner ───────────────────────

async function runAction(action) {
  const webhookUrl = WEBHOOKS[action];

  if (!webhookUrl || webhookUrl.startsWith('PASTE_')) {
    showResult('error', 'Webhook not configured', 'Check webhook URLs.');
    return;
  }

  setRunning(action, true);
  clearResult();

  try {
    console.log('Step 1: starting');
    showResult('info', 'Debug', 'Step 1: starting...');

    const base64Doc = await getDocumentAsBase64();
    console.log('Step 2: got document, length:', base64Doc.length);
    showResult('info', 'Debug', 'Step 2: document read OK, length: ' + base64Doc.length);

    const docName = document.getElementById('docName').textContent || 'document.docx';
    console.log('Step 3: calling webhook:', webhookUrl);
    showResult('info', 'Debug', 'Step 3: calling Make webhook...');

    const response = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: action,
        filename: docName,
        document: base64Doc,
        timestamp: new Date().toISOString(),
      }),
    });

    console.log('Step 4: response status:', response.status);
    showResult('info', 'Debug', 'Step 4: Make responded with status ' + response.status);

    if (!response.ok) {
      throw new Error(`Make returned ${response.status}: ${response.statusText}`);
    }

    const result = await response.json();
    console.log('Step 5: result:', result);
    await handleResult(action, result, docName);

  } catch (err) {
    console.error('DAILA error:', err);
    showResult('error', 'Error: ' + err.message, err.stack || '');
  } finally {
    setRunning(action, false);
  }
}

// ─── Result handler ───────────────────────────
//
//  Make should return ONE of the following JSON shapes:
//
//  Shape A — base64 docx:
//    { "type": "base64", "data": "<base64 string>", "filename": "summary.docx" }
//
//  Shape B — download URL:
//    { "type": "url", "url": "https://...", "filename": "summary.docx" }
//
//  Both shapes may also include an optional "message" string.
//
async function handleResult(action, result, originalName) {
  const label = ACTION_LABELS[action];

  // Optional message from Make
  const message = result.message || `${label} complete.`;

  if (result.type === 'base64' && result.data) {
    // Open as new Word document directly
    await openBase64AsNewDoc(result.data, result.filename || buildOutputName(action, originalName));
    showResult('success', `${label} ready`,
      message + '\n\nThe result has been opened as a new document.');
    showStatus('Done');

  } else if (result.type === 'url' && result.url) {
    // Show a download / open button
    showResult('success', `${label} ready`, message);
    showDownloadButton(result.url, result.filename || buildOutputName(action, originalName));
    showStatus('Done');

  } else {
    // Unexpected shape — show raw for debugging
    showResult('info', 'Unexpected response format',
      'Make returned an unrecognised response. Check the browser console for details.\n\n'
      + JSON.stringify(result, null, 2).slice(0, 300));
    console.warn('DAILA unexpected result:', result);
    showStatus('Check response format', true);
  }
}

// ─── Read document as base64 ──────────────────

function getDocumentAsBase64() {
  return new Promise((resolve) => {
    // TEMPORARY: return dummy base64 to test webhook connectivity
    console.log('getDocumentAsBase64: returning dummy data');
    setTimeout(() => {
      resolve('DUMMYDOCUMENTBASE64TESTSTRING');
    }, 500);
  });
}
// ─── Open base64 docx as new Word document ────

function openBase64AsNewDoc(base64, filename) {
  return new Promise((resolve, reject) => {
    // Decode base64 → byte array
    const binary = atob(base64);
    const bytes  = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }

    // Create a Blob and object URL, then use Word API to open
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    const url = URL.createObjectURL(blob);

    // Word.run to open the document
    Word.run(async (context) => {
      try {
        // openDocument opens a new document from a base64 string (Word 2016+)
        context.application.openDocument(base64);
        await context.sync();
        URL.revokeObjectURL(url);
        resolve();
      } catch (e) {
        // Fallback: trigger download if openDocument not available
        URL.revokeObjectURL(url);
        const blobUrl = URL.createObjectURL(blob);
        triggerDownload(blobUrl, filename);
        resolve();
      }
    }).catch((err) => {
      // If Word.run fails entirely, fall back to download
      triggerDownload(url, filename);
      resolve();
    });
  });
}

// ─── Helpers ─────────────────────────────────

function buildOutputName(action, original) {
  const base = original.replace(/\.docx$/i, '');
  const suffix = { summary: 'Summary', contract: 'Contract-Analysis', issues: 'Issue-List' };
  return `${base}_${suffix[action] || action}.docx`;
}

function concatUint8Arrays(arrays) {
  const totalLen = arrays.reduce((acc, a) => acc + a.byteLength, 0);
  const result   = new Uint8Array(totalLen);
  let   offset   = 0;
  for (const a of arrays) {
    result.set(new Uint8Array(a), offset);
    offset += a.byteLength;
  }
  return result;
}

function uint8ToBase64(bytes) {
  let binary = '';
  const chunkSize = 8192;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function triggerDownload(url, filename) {
  const a = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// ─── UI state helpers ─────────────────────────

function setRunning(action, running) {
  const actionMap = { summary: 'btn-summary', contract: 'btn-contract', issues: 'btn-issues' };
  const allBtns = Object.values(actionMap).map(id => document.getElementById(id));

  // Disable / enable all buttons
  allBtns.forEach(btn => { btn.disabled = running; });

  // Toggle running class on the active button
  const activeBtn = document.getElementById(actionMap[action]);
  if (running) {
    activeBtn.classList.add('running');
  } else {
    activeBtn.classList.remove('running');
  }

  // Progress bar
  const wrap = document.getElementById('progressWrap');
  wrap.classList.toggle('visible', running);
}

function showResult(type, header, body) {
  const area    = document.getElementById('resultArea');
  const headerEl = document.getElementById('resultHeader');
  const bodyEl   = document.getElementById('resultBody');
  const actionsEl = document.getElementById('resultActions');

  const icons = {
    success: '✓',
    error:   '✕',
    info:    'i',
  };

  area.className = `result-area visible ${type}`;
  headerEl.textContent = `${icons[type] || ''} ${header}`;
  bodyEl.textContent   = body || '';
  actionsEl.innerHTML  = '';
}

function showDownloadButton(url, filename) {
  const actionsEl = document.getElementById('resultActions');
  actionsEl.innerHTML = `
    <button class="download-btn" onclick="triggerDownload('${url}', '${filename}')">
      <svg viewBox="0 0 13 13" fill="none" stroke="white" stroke-width="1.5">
        <path d="M6.5 1v8M3.5 6l3 3 3-3"/>
        <path d="M1 10h11v1.5H1z"/>
      </svg>
      Download ${filename}
    </button>`;
}

function clearResult() {
  const area = document.getElementById('resultArea');
  area.className = 'result-area';
}

function showStatus(text, isError = false) {
  document.getElementById('statusText').textContent = text;
  const dot = document.getElementById('statusDot');
  dot.className = isError ? 'status-dot error' : 'status-dot';
}
