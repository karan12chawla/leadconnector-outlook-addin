// ── GHL Contact Capture – taskpane.js ─────────────────────────────────────
// Uses Office.js to read sender info directly from Outlook's API.
// No DOM scraping. Works on Outlook Web, Desktop (Win/Mac), and Mobile.

'use strict';

// ── Storage helpers (localStorage — this is our own hosted page) ───────────
const Store = {
  get: key         => localStorage.getItem(key) || '',
  set: (key, val)  => localStorage.setItem(key, val),
};

// ── DOM helpers ───────────────────────────────────────────────────────────
const el = id => document.getElementById(id);

function show(id) { el(id).classList.remove('hidden'); }
function hide(id) { el(id).classList.add('hidden');    }

function setStatus(msg, ok) {
  const s = el('status');
  s.textContent = msg;
  s.className   = 'status ' + (ok ? 'ok' : 'err');
}
function clearStatus() {
  const s = el('status');
  s.className   = 'status hidden';
  s.textContent = '';
}

function setLoading(on) {
  el('btn-label').textContent = on ? 'Adding…' : 'Add to GHL';
  el('spinner').classList.toggle('hidden', !on);
  el('btn-add').disabled = on;
}

// ── Office.js initialisation ──────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) {
    hide('loading');
    setStatus('This add-in only works in Outlook.', false);
    show('main');
    return;
  }

  bindEvents();
  loadSettings();
  loadContact();
});

// ── Load contact from Office.js API ──────────────────────────────────────
function loadContact() {
  const item = Office.context.mailbox.item;
  if (!item) {
    hide('loading');
    setStatus('Open an email to use this add-in.', false);
    show('main');
    return;
  }

  // item.from has getAsync in newer Outlook builds; fall back to sync access
  if (item.from && typeof item.from.getAsync === 'function') {
    item.from.getAsync(result => {
      const sender = result.status === Office.AsyncResultStatus.Succeeded ? result.value : null;
      renderContact(
        sender?.displayName  || '',
        sender?.emailAddress || '',
        typeof item.subject === 'string' ? item.subject : ''
      );
    });
  } else {
    renderContact(
      item.from?.displayName  || '',
      item.from?.emailAddress || '',
      typeof item.subject === 'string' ? item.subject : ''
    );
  }
}

function renderContact(name, email, subject) {
  hide('loading');
  checkConfig();

  const parts    = name.trim().split(' ');
  const fname    = parts[0] || '';
  const lname    = parts.slice(1).join(' ') || '';
  const initials = ((fname[0] || '') + (lname[0] || '')).toUpperCase() || '?';

  el('f-fname').value  = fname;
  el('f-lname').value  = lname;
  el('f-email').value  = email;
  el('f-note').value   = subject ? `Re: ${subject}` : '';

  el('avatar').textContent  = initials;
  el('c-name').textContent  = name  || 'Unknown sender';
  el('c-email').textContent = email || 'No email detected';

  show('main');
}

// ── Check if API credentials are saved ───────────────────────────────────
function checkConfig() {
  const hasKey = !!Store.get('ghlApiKey');
  const hasLoc = !!Store.get('ghlLocationId');
  if (!hasKey || !hasLoc) {
    show('no-config');
  } else {
    hide('no-config');
  }
}

// ── Load saved settings into fields ──────────────────────────────────────
function loadSettings() {
  const key = Store.get('ghlApiKey');
  const loc = Store.get('ghlLocationId');
  if (key) el('s-apikey').value = key;
  if (loc) el('s-locid').value  = loc;
}

// ── Bind all UI events ────────────────────────────────────────────────────
function bindEvents() {

  // Save settings
  el('btn-save').addEventListener('click', () => {
    const key = el('s-apikey').value.trim();
    const loc = el('s-locid').value.trim();
    if (!key || !loc) {
      el('save-status').textContent = '✗ Both fields required';
      el('save-status').style.color = 'var(--error)';
      show('save-status');
      setTimeout(() => hide('save-status'), 2500);
      return;
    }
    Store.set('ghlApiKey',    key);
    Store.set('ghlLocationId', loc);
    el('save-status').textContent = '✓ Saved';
    el('save-status').style.color = 'var(--success)';
    show('save-status');
    setTimeout(() => hide('save-status'), 2000);
    checkConfig();
  });

  // Toggle API key visibility
  el('toggle-pw').addEventListener('click', () => {
    const input   = el('s-apikey');
    const showing = input.type === 'text';
    input.type    = showing ? 'password' : 'text';
    el('eye-icon').innerHTML = showing
      ? '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>'
      : '<line x1="1" y1="1" x2="23" y2="23"/><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="17.94" y1="17.94" x2="23" y2="23"/>';
  });

  // Submit contact to GHL
  el('btn-add').addEventListener('click', submitContact);

  // Enter key on fields triggers submit
  ['f-fname','f-lname','f-email','f-phone','f-company','f-tags'].forEach(id => {
    el(id).addEventListener('keydown', e => { if (e.key === 'Enter') submitContact(); });
  });
}

// ── Submit contact to GoHighLevel ─────────────────────────────────────────
async function submitContact() {
  clearStatus();

  const fname   = el('f-fname').value.trim();
  const lname   = el('f-lname').value.trim();
  const email   = el('f-email').value.trim();
  const phone   = el('f-phone').value.trim();
  const company = el('f-company').value.trim();
  const tagsRaw = el('f-tags').value.trim();
  const note    = el('f-note').value.trim();

  if (!email) { setStatus('Email is required.', false); return; }

  const apiKey     = Store.get('ghlApiKey');
  const locationId = Store.get('ghlLocationId');

  if (!apiKey || !locationId) {
    setStatus('API key or Location ID missing — fill in Settings below.', false);
    show('no-config');
    return;
  }

  const tags = tagsRaw ? tagsRaw.split(',').map(t => t.trim()).filter(Boolean) : [];

  const payload = {
    firstName:  fname,
    lastName:   lname,
    email,
    locationId,
    ...(phone   && { phone }),
    ...(company && { companyName: company }),
    ...(tags.length && { tags }),
    source: 'Outlook Add-in – GHL Contact Capture',
  };

  setLoading(true);

  try {
    // ── Try GHL v2 (LeadConnector) first ─────────────────────────────────
    const res = await fetch('https://services.leadconnectorhq.com/contacts/', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'Version': '2021-07-28',
      },
      body: JSON.stringify(payload),
    });

    const data = await res.json();

    if (res.ok && data.contact) {
      // Attach note if provided
      if (note) await addNote(data.contact.id, note, apiKey);
      setStatus('✓ Contact added to GoHighLevel!', true);
      el('btn-add').disabled = true;
      setLoading(false);
      return;
    }

    // Surface GHL error message if available
    if (data.message) throw new Error(data.message);

    // ── Fallback: GHL v1 ──────────────────────────────────────────────────
    const res1 = await fetch('https://rest.gohighlevel.com/v1/contacts/', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify(payload),
    });

    const data1 = await res1.json();

    if (res1.ok && data1.contact) {
      if (note) await addNote(data1.contact.id, note, apiKey);
      setStatus('✓ Contact added to GoHighLevel!', true);
      el('btn-add').disabled = true;
    } else {
      throw new Error(data1.message || `HTTP ${res1.status}`);
    }

  } catch (err) {
    setStatus(`✗ ${err.message || 'Something went wrong.'}`, false);
  }

  setLoading(false);
}

// ── Attach a note to a GHL contact ───────────────────────────────────────
async function addNote(contactId, body, apiKey) {
  try {
    await fetch(`https://services.leadconnectorhq.com/contacts/${contactId}/notes`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'Version': '2021-07-28',
      },
      body: JSON.stringify({ body }),
    });
  } catch (_) {
    // Note failure is non-critical — don't surface to user
  }
}
