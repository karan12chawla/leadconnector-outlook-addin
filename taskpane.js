// ── GHL Contact Capture – taskpane.js ─────────────────────────────────────
'use strict';

const Store = {
  get: key        => localStorage.getItem(key) || '',
  set: (key, val) => localStorage.setItem(key, val),
};

const el = id => document.getElementById(id);
function show(id) { el(id).classList.remove('hidden'); }
function hide(id) { el(id).classList.add('hidden'); }

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

// ── Screen navigation ─────────────────────────────────────────────────────
function showScreen(name) {
  ['screen-contact', 'screen-settings'].forEach(id => hide(id));
  show('screen-' + name);
}

// ── Office.js initialisation ──────────────────────────────────────────────
Office.onReady(() => {
  bindEvents();
  loadSettings();

  const hasCredentials = Store.get('ghlApiKey') && Store.get('ghlLocationId');
  if (!hasCredentials) {
    showScreen('settings');
  } else {
    showScreen('contact');
    loadContact();
  }
});

// ── Load contact from email ───────────────────────────────────────────────
function loadContact() {
  const item = Office.context.mailbox && Office.context.mailbox.item;

  // Safety timeout — show empty form after 4s if nothing resolves
  const timeout = setTimeout(() => renderContact('', '', ''), 4000);

  if (!item) {
    clearTimeout(timeout);
    renderContact('', '', '');
    return;
  }

  try {
    // New Outlook: item.from is a From object with getAsync
    if (item.from && typeof item.from.getAsync === 'function') {
      item.from.getAsync(result => {
        clearTimeout(timeout);
        const sender = result.status === Office.AsyncResultStatus.Succeeded ? result.value : null;
        renderContact(
          sender?.displayName  || '',
          sender?.emailAddress || '',
          typeof item.subject === 'string' ? item.subject : ''
        );
      });
    } else {
      // Classic Outlook: item.from is a plain EmailAddressDetails object
      clearTimeout(timeout);
      renderContact(
        item.from?.displayName  || '',
        item.from?.emailAddress || '',
        typeof item.subject === 'string' ? item.subject : ''
      );
    }
  } catch (e) {
    clearTimeout(timeout);
    renderContact('', '', '');
  }
}

function renderContact(name, email, subject) {
  existingContactId = null;
  el('btn-label').textContent = 'Add to GHL';
  el('btn-add').disabled = false;
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

  hide('loading');
  show('contact-body');
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

  // Navigate to settings
  el('btn-go-settings').addEventListener('click', () => showScreen('settings'));

  // Navigate back to contact
  el('btn-go-contact').addEventListener('click', () => {
    showScreen('contact');
    // If contact body isn't loaded yet, trigger load
    if (el('contact-body').classList.contains('hidden') &&
        el('loading').classList.contains('hidden')) {
      show('loading');
      loadContact();
    }
  });

  // Save settings
  el('btn-save').addEventListener('click', () => {
    const key = el('s-apikey').value.trim();
    const loc = el('s-locid').value.trim();
    if (!key || !loc) {
      el('save-status').textContent = '✗ Both fields are required';
      el('save-status').style.color = 'var(--error)';
      show('save-status');
      setTimeout(() => hide('save-status'), 2500);
      return;
    }
    Store.set('ghlApiKey',     key);
    Store.set('ghlLocationId', loc);
    el('save-status').textContent = '✓ Saved';
    el('save-status').style.color = 'var(--success)';
    show('save-status');
    setTimeout(() => {
      hide('save-status');
      showScreen('contact');
      if (el('contact-body').classList.contains('hidden')) {
        loadContact();
      }
    }, 1000);
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

  // Submit contact
  el('btn-add').addEventListener('click', submitContact);

  ['f-fname','f-lname','f-email','f-phone','f-company','f-tags'].forEach(id => {
    el(id).addEventListener('keydown', e => { if (e.key === 'Enter') submitContact(); });
  });
}

// ── Status helpers ────────────────────────────────────────────────────────
function setWarn(msg) {
  const s = el('status');
  s.textContent = msg;
  s.className   = 'status warn';
}

// ── Track existing contact for update mode ────────────────────────────────
let existingContactId = null;

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
    setStatus('API credentials missing — go to Settings.', false);
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

  // ── Update mode (contact already existed) ────────────────────────────────
  if (existingContactId) {
    try {
      const res = await fetch(`https://services.leadconnectorhq.com/contacts/${existingContactId}`, {
        method: 'PUT',
        headers: {
          'Content-Type':  'application/json',
          'Authorization': `Bearer ${apiKey}`,
          'Version':       '2021-07-28',
        },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (res.ok) {
        if (note) await addNote(existingContactId, note, apiKey);
        setStatus('✓ Contact updated in GoHighLevel!', true);
        el('btn-add').disabled = true;
      } else {
        throw new Error(data.message || `HTTP ${res.status}`);
      }
    } catch (err) {
      setStatus(`✗ ${err.message || 'Something went wrong.'}`, false);
    }
    setLoading(false);
    return;
  }

  // ── Create mode ───────────────────────────────────────────────────────────
  try {
    const res  = await fetch('https://services.leadconnectorhq.com/contacts/', {
      method: 'POST',
      headers: {
        'Content-Type':  'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'Version':       '2021-07-28',
      },
      body: JSON.stringify(payload),
    });

    const data = await res.json();

    if (res.ok && data.contact) {
      if (note) await addNote(data.contact.id, note, apiKey);
      setStatus('✓ Contact added to GoHighLevel!', true);
      el('btn-add').disabled = true;
      setLoading(false);
      return;
    }

    // Detect duplicate
    const isDuplicate = !res.ok && /exist|duplicate/i.test(data.message || '');
    if (isDuplicate) {
      setLoading(false);
      await handleExistingContact(email, apiKey, locationId);
      return;
    }

    if (data.message) throw new Error(data.message);

    // Fallback: GHL v1
    const res1  = await fetch('https://rest.gohighlevel.com/v1/contacts/', {
      method: 'POST',
      headers: {
        'Content-Type':  'application/json',
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
      const isDup1 = /exist|duplicate/i.test(data1.message || '');
      if (isDup1) {
        setLoading(false);
        await handleExistingContact(email, apiKey, locationId);
        return;
      }
      throw new Error(data1.message || `HTTP ${res1.status}`);
    }

  } catch (err) {
    setStatus(`✗ ${err.message || 'Something went wrong.'}`, false);
  }

  setLoading(false);
}

// ── Fetch existing contact and pre-fill form ──────────────────────────────
async function handleExistingContact(email, apiKey, locationId) {
  try {
    const res  = await fetch(
      `https://services.leadconnectorhq.com/contacts/search?locationId=${encodeURIComponent(locationId)}&query=${encodeURIComponent(email)}&limit=1`,
      { headers: { 'Authorization': `Bearer ${apiKey}`, 'Version': '2021-07-28' } }
    );
    const data = await res.json();
    const c    = data.contacts?.[0];

    if (c) {
      existingContactId = c.id;
      el('f-fname').value   = c.firstName   || '';
      el('f-lname').value   = c.lastName    || '';
      el('f-email').value   = c.email       || '';
      el('f-phone').value   = c.phone       || '';
      el('f-company').value = c.companyName || '';
      el('f-tags').value    = (c.tags || []).join(', ');
      el('btn-label').textContent = 'Update in GHL';
      setWarn('Contact already exists — fields pre-filled. Edit and click Update.');
    } else {
      setWarn('Contact already exists in GHL.');
    }
  } catch (err) {
    setWarn('Contact already exists in GHL.');
  }
}

// ── Attach a note to a GHL contact ───────────────────────────────────────
async function addNote(contactId, body, apiKey) {
  try {
    await fetch(`https://services.leadconnectorhq.com/contacts/${contactId}/notes`, {
      method: 'POST',
      headers: {
        'Content-Type':  'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'Version':       '2021-07-28',
      },
      body: JSON.stringify({ body }),
    });
  } catch (_) {
    // Non-critical — don't surface to user
  }
}
