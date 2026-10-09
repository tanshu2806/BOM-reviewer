document.addEventListener('DOMContentLoaded', () => {
  // DOM Elements
  const themeToggle = document.getElementById('themeToggle');
  const body = document.body;

  const drawingDropzone = document.getElementById('drawingDropzone');
  const drawingInput = document.getElementById('drawingInput');
  const drawingDetails = document.getElementById('drawingDetails');
  const removeDrawingBtn = document.getElementById('removeDrawing');

  const bomDropzone = document.getElementById('bomDropzone');
  const bomInput = document.getElementById('bomInput');
  const bomDetails = document.getElementById('bomDetails');
  const removeBomBtn = document.getElementById('removeBom');

  const configInput = document.getElementById('configInput');
  const configFileName = document.getElementById('configFileName');

  const runCheckBtn = document.getElementById('runCheckBtn');
  const resetBtn = document.getElementById('resetBtn');
  const sampleBtn = document.getElementById('sampleBtn');

  const loadingOverlay = document.getElementById('loadingOverlay');
  const errorAlert = document.getElementById('errorAlert');
  const errorMessage = document.getElementById('errorMessage');

  const resultsSection = document.getElementById('resultsSection');
  const completenessVal = document.getElementById('completenessVal');
  const circleFill = document.getElementById('circleFill');
  const drawingItemsVal = document.getElementById('drawingItemsVal');
  const drawingTypeMeta = document.getElementById('drawingTypeMeta');
  const bomRowsVal = document.getElementById('bomRowsVal');
  const bomSheetMeta = document.getElementById('bomSheetMeta');
  const reliabilityBadge = document.getElementById('reliabilityBadge');
  const reliabilityReasons = document.getElementById('reliabilityReasons');

  const issueCountersStrip = document.getElementById('issueCountersStrip');
  const totalIssuesCount = document.getElementById('totalIssuesCount');
  const issuesContainer = document.getElementById('issuesContainer');
  const issuesSearch = document.getElementById('issuesSearch');

  const downloadPdfBtn = document.getElementById('downloadPdfBtn');
  const downloadXlsxBtn = document.getElementById('downloadXlsxBtn');
  const pdfFrame = document.getElementById('pdfFrame');
  const pdfOpenTab = document.getElementById('pdfOpenTab');
  const bomTableBody = document.getElementById('bomTableBody');
  const jsonOutput = document.getElementById('jsonOutput');
  const copyJsonBtn = document.getElementById('copyJsonBtn');
  const runIdVal = document.getElementById('runIdVal');

  const configBtn = document.getElementById('configBtn');
  const configModal = document.getElementById('configModal');
  const closeConfigModal = document.getElementById('closeConfigModal');
  const configJsonEditor = document.getElementById('configJsonEditor');
  const saveConfigBtn = document.getElementById('saveConfigBtn');
  const resetConfigBtn = document.getElementById('resetConfigBtn');

  let activeReportData = null;
  let currentActiveSeverity = 'all';

  // --- Theme Toggle ---
  const savedTheme = localStorage.getItem('bom_checker_theme') || 'light-theme';
  body.className = savedTheme;
  updateThemeIcon();

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('light-theme')) {
      body.classList.replace('light-theme', 'dark-theme');
      localStorage.setItem('bom_checker_theme', 'dark-theme');
    } else {
      body.classList.replace('dark-theme', 'light-theme');
      localStorage.setItem('bom_checker_theme', 'light-theme');
    }
    updateThemeIcon();
  });

  function updateThemeIcon() {
    const isDark = body.classList.contains('dark-theme');
    themeToggle.innerHTML = isDark ? '<i class="fa-solid fa-sun"></i>' : '<i class="fa-solid fa-moon"></i>';
  }

  // --- Dropzone Event Listeners ---
  function setupDropzone(dropzone, input, detailsEl, removeBtn) {
    ['dragenter', 'dragover'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
      }, false);
    });

    dropzone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files.length > 0) {
        input.files = files;
        updateFileDetails(files[0], detailsEl);
      }
    });

    input.addEventListener('change', () => {
      if (input.files.length > 0) {
        updateFileDetails(input.files[0], detailsEl);
      }
    });

    removeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      input.value = '';
      detailsEl.classList.remove('active');
    });
  }

  function updateFileDetails(file, detailsEl) {
    const fileNameSpan = detailsEl.querySelector('.file-name');
    fileNameSpan.textContent = `${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
    detailsEl.classList.add('active');
  }

  setupDropzone(drawingDropzone, drawingInput, drawingDetails, removeDrawingBtn);
  setupDropzone(bomDropzone, bomInput, bomDetails, removeBomBtn);

  configInput.addEventListener('change', () => {
    if (configInput.files.length > 0) {
      configFileName.textContent = configInput.files[0].name;
    } else {
      configFileName.textContent = 'No file selected';
    }
  });

  resetBtn.addEventListener('click', () => {
    drawingInput.value = '';
    bomInput.value = '';
    configInput.value = '';
    drawingDetails.classList.remove('active');
    bomDetails.classList.remove('active');
    configFileName.textContent = 'No file selected';
    resultsSection.classList.add('hidden');
    errorAlert.classList.add('hidden');
  });

  // --- Run Check Action ---
  runCheckBtn.addEventListener('click', () => {
    if (!drawingInput.files[0] || !bomInput.files[0]) {
      showError('Please upload both an Engineering Drawing PDF and a BOM Excel file.');
      return;
    }

    const formData = new FormData();
    formData.append('drawing', drawingInput.files[0]);
    formData.append('bom', bomInput.files[0]);
    if (configInput.files[0]) {
      formData.append('config', configInput.files[0]);
    }

    executeCheck(formData);
  });

  sampleBtn.addEventListener('click', () => {
    const formData = new FormData();
    formData.append('use_sample', 'true');
    executeCheck(formData);
  });

  function executeCheck(formData) {
    hideError();
    loadingOverlay.classList.remove('hidden');
    resultsSection.classList.add('hidden');

    fetch('/api/check', {
      method: 'POST',
      body: formData
    })
    .then(async res => {
      const contentType = res.headers.get('content-type') || '';
      if (contentType.includes('application/json')) {
        return res.json();
      } else {
        const text = await res.text();
        throw new Error(`Server response error (${res.status}): ${text.slice(0, 180)}`);
      }
    })
    .then(data => {
      loadingOverlay.classList.add('hidden');
      if (data.status === 'success') {
        renderResults(data);
      } else {
        showError(data.message || 'Verification failed. Please check inputs.');
      }
    })
    .catch(err => {
      loadingOverlay.classList.add('hidden');
      showError(`Verification Error: ${err.message}`);
    });
  }

  function showError(msg) {
    errorMessage.textContent = msg;
    errorAlert.classList.remove('hidden');
    errorAlert.scrollIntoView({ behavior: 'smooth' });
  }

  function hideError() {
    errorAlert.classList.add('hidden');
  }

  // --- Render Results & Dashboard ---
  function renderResults(data) {
    const report = data.report || {};
    activeReportData = report;
    runIdVal.textContent = data.run_id;

    // Completeness Gauge
    const compVal = report.completeness || 0;
    completenessVal.textContent = compVal.toFixed(1);
    circleFill.setAttribute('stroke-dasharray', `${compVal}, 100`);

    // Counts
    drawingItemsVal.textContent = report.n_drawing_items || 0;
    bomRowsVal.textContent = report.n_bom_rows || 0;

    // Reliability
    const rel = report.reliability || 'HIGH';
    reliabilityBadge.textContent = rel;
    reliabilityBadge.className = `status-badge ${rel}`;
    reliabilityReasons.textContent = (report.reliability_reasons && report.reliability_reasons.length > 0)
      ? report.reliability_reasons.join('; ')
      : 'All extraction metrics optimal';

    // File Downloads & Viewing
    if (data.annotated_pdf_filename) {
      const pdfUrl = `/api/view/${data.run_id}/${data.annotated_pdf_filename}`;
      const downloadPdfUrl = `/api/download/${data.run_id}/${data.annotated_pdf_filename}`;
      pdfFrame.src = pdfUrl;
      pdfOpenTab.href = pdfUrl;
      downloadPdfBtn.href = downloadPdfUrl;
      downloadPdfBtn.style.display = 'inline-flex';
    } else {
      downloadPdfBtn.style.display = 'none';
    }

    if (data.checked_xlsx_filename) {
      downloadXlsxBtn.href = `/api/download/${data.run_id}/${data.checked_xlsx_filename}`;
      downloadXlsxBtn.style.display = 'inline-flex';
    } else {
      downloadXlsxBtn.style.display = 'none';
    }

    // Issues list & counter strip
    const issues = report.issues || [];
    totalIssuesCount.textContent = issues.length;
    renderCounterStrip(issues);
    renderIssuesList(issues, 'all', '');

    // BOM Table
    renderBomTable(report);

    // JSON Viewer
    jsonOutput.querySelector('code').textContent = JSON.stringify(report, null, 2);

    resultsSection.classList.remove('hidden');
    resultsSection.scrollIntoView({ behavior: 'smooth' });
  }

  function renderCounterStrip(issues) {
    const counts = { red: 0, orange: 0, yellow: 0, blue: 0 };
    issues.forEach(i => {
      const sev = (i.severity || '').toLowerCase();
      if (counts[sev] !== undefined) counts[sev]++;
    });

    issueCountersStrip.innerHTML = `
      <div class="counter-pill pill-red">Omissions <span>${counts.red}</span></div>
      <div class="counter-pill pill-orange">Mismatches <span>${counts.orange}</span></div>
      <div class="counter-pill pill-yellow">Confirm <span>${counts.yellow}</span></div>
      <div class="counter-pill pill-blue">Implied <span>${counts.blue}</span></div>
    `;
  }

  function renderIssuesList(issues, severityFilter, searchQuery) {
    issuesContainer.innerHTML = '';

    const filtered = issues.filter(issue => {
      const matchesSeverity = (severityFilter === 'all') || (issue.severity.toLowerCase() === severityFilter.toLowerCase());
      const query = searchQuery.toLowerCase();
      const matchesSearch = !query || 
        (issue.message && issue.message.toLowerCase().includes(query)) ||
        (issue.type && issue.type.toLowerCase().includes(query)) ||
        (issue.item_id && issue.item_id.toString().toLowerCase().includes(query));
      return matchesSeverity && matchesSearch;
    });

    if (filtered.length === 0) {
      issuesContainer.innerHTML = `
        <div style="padding: 2rem; text-align: center; color: var(--text-muted);">
          <i class="fa-solid fa-circle-check" style="font-size: 2.5rem; margin-bottom: 0.5rem; color: var(--color-emerald);"></i>
          <p>No issues match the selected filter query.</p>
        </div>
      `;
      return;
    }

    filtered.forEach(issue => {
      const sevClass = (issue.severity || 'green').toLowerCase();
      const card = document.createElement('div');
      card.className = `issue-card severity-${sevClass}`;
      card.innerHTML = `
        <div class="issue-header">
          <span class="severity-tag ${issue.severity}">${issue.severity}</span>
          <span class="issue-type">${issue.type || 'Finding'}</span>
          ${issue.item_id ? `<span style="margin-left:auto; font-size:0.78rem; opacity:0.8;">Item #${issue.item_id}</span>` : ''}
        </div>
        <div class="issue-msg">${issue.message || ''}</div>
      `;
      issuesContainer.appendChild(card);
    });
  }

  // --- Filter Chips & Search ---
  document.querySelectorAll('.filter-chips .chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('.filter-chips .chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      currentActiveSeverity = chip.dataset.severity;
      if (activeReportData) {
        renderIssuesList(activeReportData.issues || [], currentActiveSeverity, issuesSearch.value);
      }
    });
  });

  issuesSearch.addEventListener('input', (e) => {
    if (activeReportData) {
      renderIssuesList(activeReportData.issues || [], currentActiveSeverity, e.target.value);
    }
  });

  // --- Render BOM Items Table ---
  function renderBomTable(report) {
    bomTableBody.innerHTML = '';
    const bomRows = report.bom_rows || [];
    
    if (bomRows.length === 0) {
      bomTableBody.innerHTML = `<tr><td colspan="6" style="text-align:center; color: var(--text-muted);">No BOM rows extracted.</td></tr>`;
      return;
    }

    bomRows.forEach(row => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${row.item || row.item_no || '-'}</strong></td>
        <td>${row.partno || row.part_no || '-'}</td>
        <td>${row.description || '-'}</td>
        <td>${row.material || '-'}</td>
        <td>${row.qty || '-'}</td>
        <td><span class="badge-tag">${row.status || 'OK'}</span></td>
      `;
      bomTableBody.appendChild(tr);
    });
  }

  // --- Tabs Navigation ---
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      
      btn.classList.add('active');
      const targetTab = document.getElementById(btn.dataset.tab);
      if (targetTab) targetTab.classList.add('active');
    });
  });

  // --- Copy JSON ---
  copyJsonBtn.addEventListener('click', () => {
    const text = jsonOutput.querySelector('code').textContent;
    navigator.clipboard.writeText(text).then(() => {
      copyJsonBtn.innerHTML = '<i class="fa-solid fa-check"></i> Copied!';
      setTimeout(() => {
        copyJsonBtn.innerHTML = '<i class="fa-solid fa-copy"></i> Copy JSON';
      }, 2000);
    });
  });

  // --- Configuration Modal ---
  configBtn.addEventListener('click', () => {
    fetch('/api/config')
      .then(res => res.json())
      .then(data => {
        if (data.status === 'success') {
          configJsonEditor.value = JSON.stringify(data.config, null, 2);
          configModal.classList.remove('hidden');
        }
      });
  });

  closeConfigModal.addEventListener('click', () => configModal.classList.add('hidden'));

  saveConfigBtn.addEventListener('click', () => {
    try {
      const parsed = JSON.parse(configJsonEditor.value);
      fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(parsed)
      })
      .then(res => res.json())
      .then(data => {
        if (data.status === 'success') {
          alert('Configuration saved successfully!');
          configModal.classList.add('hidden');
        } else {
          alert(`Error saving configuration: ${data.message}`);
        }
      });
    } catch (e) {
      alert(`Invalid JSON format: ${e.message}`);
    }
  });

  resetConfigBtn.addEventListener('click', () => {
    fetch('/api/config')
      .then(res => res.json())
      .then(data => {
        if (data.status === 'success') {
          configJsonEditor.value = JSON.stringify(data.config, null, 2);
        }
      });
  });
});
