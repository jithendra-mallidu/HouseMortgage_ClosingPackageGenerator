/**
 * Closing Package Generator - Frontend JavaScript
 * Calls the backend API on Cloud Run to generate closing PDFs.
 */

// ── Configuration ──────────────────────────────────────────────────────────
// Set this to your Cloud Run URL after deployment.
// During local dev, point to http://localhost:8080
const API_BASE_URL = window.CLOSING_API_URL || '';

// ── State ──────────────────────────────────────────────────────────────────
let dotStates = new Set();
let availableAdditionalDocs = [];

// ── DOM Elements ───────────────────────────────────────────────────────────
const stateSelect = document.getElementById('state');
const modeInput = document.getElementById('modeInput');
const fillSection = document.getElementById('fillSection');
const trusteeSection = document.getElementById('trusteeSection');
const modeBtns = document.querySelectorAll('.mode-btn');
const form = document.getElementById('form');
const submitBtn = document.getElementById('submitBtn');
const loadingOverlay = document.getElementById('loadingOverlay');
const toast = document.getElementById('toast');

// ── Initialize: Load States from API ───────────────────────────────────────
async function loadStates() {
  try {
    const resp = await fetch(`${API_BASE_URL}/api/states`);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();

    // Populate state dropdown
    data.states.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      stateSelect.appendChild(opt);
    });

    // Store DOT states
    dotStates = new Set(data.dot_states);
  } catch (err) {
    console.error('Failed to load states:', err);
    showToast('Failed to connect to the server. Please try again later.');
  }
}

loadStates();

// ── Mode Toggle ────────────────────────────────────────────────────────────
modeBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    modeBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    modeInput.value = btn.dataset.mode;
    if (btn.dataset.mode === 'empty') {
      fillSection.classList.add('hidden');
    } else {
      fillSection.classList.remove('hidden');
    }
  });
});

// ── Additional Documents Section ──────────────────────────────────────────
const additionalDocsSection = document.getElementById('additionalDocsSection');
const additionalDocsList = document.getElementById('additionalDocsList');
const additionalDocsLoading = document.getElementById('additionalDocsLoading');

async function loadAdditionalDocs(state) {
  if (!state) {
    additionalDocsSection.style.display = 'none';
    return;
  }
  additionalDocsSection.style.display = '';
  additionalDocsLoading.style.display = '';
  additionalDocsList.innerHTML = '';

  try {
    const resp = await fetch(`${API_BASE_URL}/api/additional-documents?state=${encodeURIComponent(state)}`);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    availableAdditionalDocs = data.documents || [];
    renderAdditionalDocs(availableAdditionalDocs);
  } catch (err) {
    console.error('Failed to load additional documents:', err);
    additionalDocsList.innerHTML = '<div class="no-docs-message">Failed to load additional documents.</div>';
  } finally {
    additionalDocsLoading.style.display = 'none';
  }
}

function renderAdditionalDocs(docs) {
  additionalDocsList.innerHTML = '';
  if (!docs.length) {
    additionalDocsList.innerHTML = '<div class="no-docs-message">No additional documents available for this state.</div>';
    return;
  }

  // Group by category
  const grouped = {};
  docs.forEach(doc => {
    if (!grouped[doc.category]) grouped[doc.category] = [];
    grouped[doc.category].push(doc);
  });

  for (const [category, catDocs] of Object.entries(grouped)) {
    const title = document.createElement('div');
    title.className = 'doc-category-title';
    title.textContent = category;
    additionalDocsList.appendChild(title);

    const list = document.createElement('div');
    list.className = 'doc-checkbox-list';
    catDocs.forEach(doc => {
      const item = document.createElement('div');
      item.className = 'doc-checkbox-item';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.id = `doc_${doc.id}`;
      cb.value = doc.id;
      cb.name = 'additional_doc';
      const lbl = document.createElement('label');
      lbl.htmlFor = `doc_${doc.id}`;
      lbl.textContent = doc.name;
      item.appendChild(cb);
      item.appendChild(lbl);
      list.appendChild(item);
    });
    additionalDocsList.appendChild(list);
  }
}

// ── Trustee Section Visibility ─────────────────────────────────────────────
stateSelect.addEventListener('change', () => {
  if (dotStates.has(stateSelect.value)) {
    trusteeSection.classList.remove('hidden');
  } else {
    trusteeSection.classList.add('hidden');
  }
  // Load additional documents for this state
  loadAdditionalDocs(stateSelect.value);
});

// ── Auto-fill Sample Data ──────────────────────────────────────────────────
document.getElementById('autofillBtn').addEventListener('click', async () => {
  const state = stateSelect.value;
  if (!state) {
    showToast('Please select a state first.');
    return;
  }
  try {
    const resp = await fetch(`${API_BASE_URL}/api/sample-data?state=${encodeURIComponent(state)}`);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const d = await resp.json();

    function toDateInput(dateStr) {
      try {
        const d = new Date(dateStr);
        if (isNaN(d)) return '';
        return d.toISOString().split('T')[0];
      } catch { return ''; }
    }

    document.getElementById('borrower_name').value = d.borrower_name || '';
    document.getElementById('co_borrower_name').value = d.co_borrower_name || '';
    document.getElementById('borrower_address').value = d.borrower_address || '';
    document.getElementById('loan_date').value = toDateInput(d.loan_date);
    document.getElementById('interest_rate').value = d.interest_rate || '';
    document.getElementById('loan_amount_number').value = d.loan_amount_number || '';
    document.getElementById('loan_amount_words').value = d.loan_amount_words || '';
    document.getElementById('monthly_payment').value = d.monthly_payment || '';
    document.getElementById('first_payment_date').value = toDateInput(d.first_payment_date);
    document.getElementById('maturity_date').value = toDateInput(d.maturity_date);
    document.getElementById('late_charge_days').value = d.late_charge_days || '15';
    document.getElementById('late_charge_percent').value = d.late_charge_percent || '5';
    document.getElementById('lender_name').value = d.lender_name || '';
    document.getElementById('lender_org_type').value = d.lender_org_type || '';
    document.getElementById('lender_org_state').value = d.lender_org_state || '';
    document.getElementById('lender_address').value = d.lender_address || '';
    document.getElementById('trustee_name').value = d.trustee_name || '';
    document.getElementById('trustee_address').value = d.trustee_address || '';
    document.getElementById('property_street').value = d.property_street || '';
    document.getElementById('property_city').value = d.property_city || '';
    document.getElementById('property_zip').value = d.property_zip || '';
    document.getElementById('property_county').value = d.property_county || '';
    document.getElementById('recording_jurisdiction_name').value = d.recording_jurisdiction_name || '';

    // Cancel deadline: 3 days from loan date
    const loanDate = new Date(d.loan_date);
    if (!isNaN(loanDate)) {
      loanDate.setDate(loanDate.getDate() + 3);
      document.getElementById('cancel_deadline').value = loanDate.toISOString().split('T')[0];
    }
  } catch (e) {
    console.error('Auto-fill failed:', e);
    showToast('Failed to generate sample data.');
  }
});

// ── Form Submission (API call) ─────────────────────────────────────────────
form.addEventListener('submit', async (e) => {
  e.preventDefault();

  const state = stateSelect.value;
  if (!state) {
    showToast('Please select a state.');
    return;
  }

  // Show loading
  submitBtn.classList.add('loading');
  submitBtn.disabled = true;
  loadingOverlay.classList.add('active');

  try {
    // Build form data to send to API
    const formData = new FormData(form);
    const body = {};
    formData.forEach((value, key) => {
      if (key === 'additional_doc') return; // handled below
      body[key] = value;
    });

    // Collect selected additional documents
    const selectedDocs = [];
    document.querySelectorAll('input[name="additional_doc"]:checked').forEach(cb => {
      selectedDocs.push(cb.value);
    });
    body.additional_documents = selectedDocs;

    const resp = await fetch(`${API_BASE_URL}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ error: 'Generation failed' }));
      throw new Error(err.error || `HTTP ${resp.status}`);
    }

    // Download the PDF
    const blob = await resp.blob();
    const filename = resp.headers.get('Content-Disposition')
      ?.match(/filename="?(.+?)"?$/)?.[1]
      || `ClosingPackage_${state}.pdf`;

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

  } catch (err) {
    console.error('Generation failed:', err);
    showToast(err.message || 'Failed to generate closing package.');
  } finally {
    submitBtn.classList.remove('loading');
    submitBtn.disabled = false;
    loadingOverlay.classList.remove('active');
  }
});

// ── Toast Notification ─────────────────────────────────────────────────────
function showToast(message) {
  toast.textContent = message;
  toast.classList.add('active');
  setTimeout(() => toast.classList.remove('active'), 4000);
}
