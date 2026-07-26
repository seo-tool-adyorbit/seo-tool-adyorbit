document.addEventListener('DOMContentLoaded', () => {
    const BACKEND_URL = 'http://localhost:8000';

    // DOM Elements
    const userDigitsInput       = document.getElementById('userDigitsInput');
    const sourceAirlineInput    = document.getElementById('sourceAirlineInput');
    const originalPreviewVal    = document.getElementById('originalPreviewVal');
    const outputPreviewVal      = document.getElementById('outputPreviewVal');
    const applySmartTfnBtn      = document.getElementById('applySmartTfnBtn');
    const exportVectorPdfBtn    = document.getElementById('exportVectorPdfBtn');
    const exportHtmlBtn         = document.getElementById('exportHtmlBtn');
    const docTitleInput         = document.getElementById('docTitleInput');
    const statusBadge           = document.getElementById('statusBadge');
    const themeToggleBtn        = document.getElementById('themeToggleBtn');
    const googleSheetsCsvInput  = document.getElementById('googleSheetsCsvInput');
    const generateBulkZipBtn    = document.getElementById('generateBulkZipBtn');
    const appModalOverlay       = document.getElementById('appModalOverlay');
    const closeModalBtn         = document.getElementById('closeModalBtn');
    const modalBody             = document.getElementById('modalBody');

    const openModal  = (html) => { if (modalBody && appModalOverlay) { modalBody.innerHTML = html; appModalOverlay.style.display = 'flex'; } };
    const closeModal = ()     => { if (appModalOverlay) appModalOverlay.style.display = 'none'; };
    if (closeModalBtn)   closeModalBtn.addEventListener('click', closeModal);
    if (appModalOverlay) appModalOverlay.addEventListener('click', e => { if (e.target === appModalOverlay) closeModal(); });

    // ── Quill Editor ──────────────────────────────────────────────────────────
    const sourceQuill = new Quill('#sourceEditor', {
        theme: 'snow',
        placeholder: 'Type or paste document content here...',
        modules: {
            toolbar: [
                [{ header: [1, 2, 3, 4, 5, 6, false] }],
                [{ font: [] }, { size: ['small', false, 'large', 'huge'] }],
                ['bold', 'italic', 'underline', 'strike'],
                [{ color: [] }, { background: [] }],
                [{ list: 'ordered' }, { list: 'bullet' }],
                [{ align: [] }],
                ['link', 'clean']
            ],
            clipboard: { matchVisual: false }
        }
    });

    sourceQuill.root.innerHTML = `
        <h1>Delta Airlines Flight Cancellation &amp; Refund Guide</h1>
        <p>Need urgent Delta Airlines flight booking assistance? Call our dedicated line at <strong>+51*23*456(7890)</strong> available 24/7.</p>
        <h2>Delta Airlines Customer Support Helpline</h2>
        <p>Alternatively, reach Delta Airlines customer support at <strong>+52 (55) 1234-5678</strong> or via international helpline at <strong>+44 20 7946 0958</strong>.</p>
        <p>For instant cancellation refund inquiries, dial <strong>+51*23*456(7890)</strong> again to speak with our representative.</p>
    `;

    // ── Helpers ───────────────────────────────────────────────────────────────
    function escapeRegExp(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

    function convertGoogleSheetUrlToCsvUrl(inputUrl) {
        const trimmed = (inputUrl || '').trim();
        const m = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (m && m[1]) {
            const gid = trimmed.match(/[?&#]gid=([0-9]+)/);
            return `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=csv${gid ? '&gid=' + gid[1] : ''}`;
        }
        return trimmed;
    }

    function formatDigitsToCompetitorPattern(competitorStr, userInputStr) {
        if (!competitorStr || !userInputStr) return competitorStr;
        const targetDigits = userInputStr.replace(/\D/g, '');
        if (!targetDigits) return competitorStr;
        let result = '', idx = 0;
        for (const ch of competitorStr) {
            result += /\d/.test(ch.normalize('NFKC'))
                ? (idx < targetDigits.length ? targetDigits[idx++] : '')
                : ch;
        }
        while (idx < targetDigits.length) result += targetDigits[idx++];
        return result;
    }

    function updateLivePreview() {
        const text     = sourceQuill.getText();
        const tfnRegex = /(?:\+?[^\p{L}\p{N}/\n]{0,4})?(?:\p{N}[^\p{L}\p{N}/\n]{0,4}){9,14}\p{N}/gu;
        const matches  = [...new Set((text.match(tfnRegex) || []).map(m => m.trim()))];
        const rawInput = userDigitsInput?.value.trim() || '';
        const targets  = rawInput.split('\n').map(l => l.trim()).filter(Boolean);
        const primary  = targets[0] || '+1 (888) 830-8592';
        const comp     = matches[0] || '+51*23*456(7890)';
        if (originalPreviewVal) originalPreviewVal.innerText = comp;
        if (outputPreviewVal) {
            const out = formatDigitsToCompetitorPattern(comp, primary);
            if (targets.length > 1) {
                const out2 = formatDigitsToCompetitorPattern(matches[1] || comp, targets[1]);
                outputPreviewVal.innerHTML = `${out} <span style="font-size:.72rem;color:var(--text-muted)">(TFN 1)</span><br><span style="color:#60a5fa">${out2}</span> <span style="font-size:.72rem;color:var(--text-muted)">(TFN 2)</span>`;
            } else {
                outputPreviewVal.innerText = out;
            }
        }
    }

    if (userDigitsInput) userDigitsInput.addEventListener('input', updateLivePreview);
    updateLivePreview();

    // ── Apply TFN Button ──────────────────────────────────────────────────────
    if (applySmartTfnBtn) {
        applySmartTfnBtn.addEventListener('click', () => {
            const rawInput = userDigitsInput.value.trim();
            if (!rawInput) return alert('Please enter target replacement number(s) first!');
            const targets  = rawInput.split('\n').map(l => l.trim()).filter(Boolean);
            const text     = sourceQuill.getText();
            const tfnRegex = /(?:\+?[^\p{L}\p{N}/\n]{0,4})?(?:\p{N}[^\p{L}\p{N}/\n]{0,4}){9,14}\p{N}/gu;
            let m, found = [];
            while ((m = tfnRegex.exec(text)) !== null)
                found.push({ index: m.index, length: m[0].length, text: m[0] });
            if (!found.length) return alert('No competitor numbers found.');
            for (let i = found.length - 1; i >= 0; i--) {
                const mv          = found[i];
                const replacement = formatDigitsToCompetitorPattern(mv.text, targets[i % targets.length]);
                const fmt         = sourceQuill.getFormat(mv.index, 1);
                sourceQuill.deleteText(mv.index, mv.length);
                if (replacement) sourceQuill.insertText(mv.index, replacement, fmt);
            }
            if (statusBadge) statusBadge.innerHTML = `<i class="fa-solid fa-check"></i> ${found.length} Replaced`;
            updateLivePreview();
        });
    }

    function applyTfnReplacementToHtml(htmlString, tfnLines) {
        const div      = document.createElement('div');
        div.innerHTML  = htmlString;
        const tfnRegex = /(?:\+?[^\p{L}\p{N}/\n]{0,4})?(?:\p{N}[^\p{L}\p{N}/\n]{0,4}){9,14}\p{N}/gu;
        let counter    = 0;
        function walk(node) {
            if (node.nodeType === Node.TEXT_NODE) {
                if (tfnRegex.test(node.textContent)) {
                    tfnRegex.lastIndex = 0;
                    node.textContent = node.textContent.replace(tfnRegex, match =>
                        formatDigitsToCompetitorPattern(match, tfnLines[counter++ % tfnLines.length])
                    );
                }
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                [...node.childNodes].forEach(walk);
            }
        }
        walk(div);
        return div.innerHTML;
    }

    function generateUniqueFilename(context) {
        const slug = (context || 'SEO_Document')
            .replace(/<[^>]*>/g, '').trim()
            .replace(/[^a-zA-Z0-9\s\-_]/g, '').replace(/\s+/g, '_').substring(0, 45) || 'SEO_Document';
        const d   = new Date();
        const ts  = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
        const rnd = Math.random().toString(36).substring(2, 6).toUpperCase();
        return `${slug}_${ts}_${rnd}.pdf`;
    }

    // ── BACKEND PDF API (LibreOffice pipeline) ────────────────────────────────
    async function backendAvailable() {
        try {
            const r = await fetch(`${BACKEND_URL}/api/health`, { signal: AbortSignal.timeout(2000) });
            if (r.ok) {
                const d = await r.json();
                return d.status === 'ok';
            }
        } catch { /* backend not running */ }
        return false;
    }

    async function generatePdfViaBackend(htmlContent, title) {
        const response = await fetch(`${BACKEND_URL}/api/v1/pdf/generate`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ html_content: htmlContent, title: title || 'SEO PDF Document' })
        });
        if (!response.ok) {
            const err = await response.json().catch(() => ({ detail: 'Unknown error' }));
            throw new Error(err.detail || `Backend error ${response.status}`);
        }
        return await response.blob();
    }

    async function generateBulkZipViaBackend(sourceAirline, templateHtml, tfnLines, csvRows) {
        const response = await fetch(`${BACKEND_URL}/api/v1/pdf/bulk-zip`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                source_airline:    sourceAirline,
                template_html:     templateHtml,
                target_tfn_lines:  tfnLines,
                csv_rows:          csvRows
            })
        });
        if (!response.ok) {
            const err = await response.json().catch(() => ({ detail: 'Unknown error' }));
            throw new Error(err.detail || `Backend error ${response.status}`);
        }
        return await response.blob();
    }

    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a   = document.createElement('a');
        a.href    = url;
        a.download = filename;
        a.click();
        setTimeout(() => URL.revokeObjectURL(url), 3000);
    }

    // ── Export PDF Button ─────────────────────────────────────────────────────
    if (exportVectorPdfBtn) {
        exportVectorPdfBtn.addEventListener('click', async () => {
            const html = sourceQuill.root.innerHTML;
            if (!html || html === '<p><br></p>') return alert('Document is empty!');

            const docTitle = docTitleInput?.value.trim() || 'SEO PDF Document';
            const origHtml = exportVectorPdfBtn.innerHTML;
            exportVectorPdfBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating PDF...';
            exportVectorPdfBtn.disabled  = true;

            try {
                const isAvailable = await backendAvailable();

                if (isAvailable) {
                    // ✅ Backend running: use LibreOffice pipeline
                    if (statusBadge) statusBadge.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> LibreOffice converting...';
                    const blob     = await generatePdfViaBackend(html, docTitle);
                    const filename = generateUniqueFilename(docTitle);
                    downloadBlob(blob, filename);
                    if (statusBadge) statusBadge.innerHTML = '<i class="fa-solid fa-circle-check"></i> PDF Exported — Google Docs Quality!';
                } else {
                    // ⚠️ Backend not running: show setup instructions
                    if (statusBadge) statusBadge.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Backend not running';
                    openModal(`
                        <h3 style="color:#ef4444;margin-bottom:12px"><i class="fa-solid fa-triangle-exclamation"></i> Python Backend Required</h3>
                        <p style="margin-bottom:12px">For Google Docs quality PDF (LibreOffice pipeline), you need to start the Python backend server.</p>
                        <div style="background:#0f172a;border-radius:8px;padding:12px;font-family:monospace;font-size:0.85rem;margin-bottom:12px">
                            <p style="color:#94a3b8;margin:0 0 6px"># Step 1: Install dependencies</p>
                            <p style="color:#34d399;margin:0 0 10px">cd "${window.location.pathname.replace('/index.html','')}/pdf-engine" && pip install -r requirements.txt</p>
                            <p style="color:#94a3b8;margin:0 0 6px"># Step 2: Install LibreOffice (macOS)</p>
                            <p style="color:#34d399;margin:0 0 10px">brew install libreoffice</p>
                            <p style="color:#94a3b8;margin:0 0 6px"># Step 3: Start backend server</p>
                            <p style="color:#34d399;margin:0">uvicorn app.main:app --reload --port 8000</p>
                        </div>
                        <p style="color:#94a3b8;font-size:0.85rem">Once backend is running at <strong>http://localhost:8000</strong>, click Download PDF again.</p>
                    `);
                }
            } catch (err) {
                console.error('PDF Export Error:', err);
                if (statusBadge) statusBadge.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Export failed';
                alert(`PDF Export Error:\n${err.message}`);
            }

            exportVectorPdfBtn.innerHTML = origHtml;
            exportVectorPdfBtn.disabled  = false;
        });
    }

    // ── Export HTML Button ────────────────────────────────────────────────────
    if (exportHtmlBtn) {
        exportHtmlBtn.addEventListener('click', () => {
            const html = sourceQuill.root.innerHTML;
            if (!html || html === '<p><br></p>') return alert('Document is empty!');
            downloadBlob(new Blob([html], { type: 'text/html;charset=utf-8;' }),
                generateUniqueFilename('HTML_Export').replace('.pdf', '.html'));
        });
    }

    // ── Bulk ZIP Button ───────────────────────────────────────────────────────
    if (generateBulkZipBtn) {
        generateBulkZipBtn.addEventListener('click', async () => {
            let csvRaw = googleSheetsCsvInput?.value.trim() || '';
            if (!csvRaw) return alert('Please enter Google Sheets CSV URL or CSV data!');

            const sourceAirline = sourceAirlineInput?.value.trim() || '';
            const tfnLines      = (userDigitsInput?.value.trim() || '').split('\n').map(l => l.trim()).filter(Boolean);
            const origBtnHtml   = generateBulkZipBtn.innerHTML;
            generateBulkZipBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processing...';
            generateBulkZipBtn.disabled  = true;

            try {
                // Fetch CSV if URL provided
                if (csvRaw.startsWith('http')) {
                    if (statusBadge) statusBadge.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Fetching CSV...';
                    const res = await fetch(convertGoogleSheetUrlToCsvUrl(csvRaw));
                    csvRaw    = await res.text();
                    if (csvRaw.toLowerCase().startsWith('<!doctype') || csvRaw.includes('<html')) {
                        alert('Could not fetch CSV. Make sure Google Sheet is publicly shared.');
                        generateBulkZipBtn.innerHTML = origBtnHtml;
                        generateBulkZipBtn.disabled  = false;
                        return;
                    }
                }

                const parsed = Papa.parse(csvRaw, { header: false, skipEmptyLines: true });
                let rows     = parsed.data.filter(r => r.length >= 3 && r[0]?.trim());
                if (!rows.length) {
                    alert('No valid CSV rows. Expected: Col A=Airline, Col B=Headline, Col C=PDF Name');
                    generateBulkZipBtn.innerHTML = origBtnHtml;
                    generateBulkZipBtn.disabled  = false;
                    return;
                }
                if (rows[0][0].toLowerCase().includes('airline')) rows = rows.slice(1);

                const isAvailable = await backendAvailable();
                const templateHtml = sourceQuill.root.innerHTML;

                if (isAvailable) {
                    // ✅ Backend: send all rows at once
                    if (statusBadge) statusBadge.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Backend generating ${rows.length} PDFs...`;
                    const csvRows = rows.map(([airline, headline, pdfName]) => ({
                        airline_name: airline.trim(),
                        headline:     headline.trim(),
                        pdf_name:     pdfName.trim()
                    }));
                    const zipBlob = await generateBulkZipViaBackend(sourceAirline, templateHtml, tfnLines, csvRows);
                    downloadBlob(zipBlob, `Bulk_SEO_PDFs_${new Date().toISOString().slice(0,10)}.zip`);
                    if (statusBadge) statusBadge.innerHTML = `<i class="fa-solid fa-circle-check"></i> ${rows.length} PDFs Downloaded!`;
                } else {
                    alert('Backend server not running. Please start the Python server first.\nSee Download PDF button for setup instructions.');
                }
            } catch (err) {
                console.error('Bulk ZIP Error:', err);
                alert(`Bulk ZIP Error:\n${err.message}`);
            }

            generateBulkZipBtn.innerHTML = origBtnHtml;
            generateBulkZipBtn.disabled  = false;
        });
    }

    // ── Modal Bindings ────────────────────────────────────────────────────────
    const bindNavModal = (id, html) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => openModal(html));
    };
    bindNavModal('navProjects', `
        <h3 style="margin-bottom:12px;color:var(--accent-color)"><i class="fa-solid fa-folder-open"></i> Projects Overview</h3>
        <div style="background:var(--bg-main);border:1px solid var(--border-color);padding:12px;border-radius:8px;font-size:0.85rem">
            <p><strong>Delta Airlines Flight Guide</strong></p>
            <p style="color:var(--success-color);margin-top:4px">Engine: LibreOffice → Google Docs Quality PDF</p>
        </div>
    `);
    bindNavModal('navVariables', `
        <h3 style="margin-bottom:12px;color:var(--accent-color)"><i class="fa-solid fa-code"></i> Document Variables</h3>
        <div style="display:flex;flex-wrap:wrap;gap:8px">
            ${['{{AIRLINE}}','{{HEADLINE}}','{{TFN}}','{{KEYWORD}}','{{CITY}}']
                .map(v => `<span class="tag-pill accent" style="cursor:pointer" onclick="navigator.clipboard.writeText('${v}');alert('Copied ${v}!')">${v}</span>`)
                .join('')}
        </div>
    `);

    if (themeToggleBtn) themeToggleBtn.addEventListener('click', () => {
        document.body.classList.toggle('dark-theme');
        document.body.classList.toggle('light-theme');
    });

    // ── Check backend on load ─────────────────────────────────────────────────
    backendAvailable().then(ok => {
        if (statusBadge) {
            statusBadge.innerHTML = ok
                ? '<i class="fa-solid fa-circle-check"></i> Backend Ready — LibreOffice PDF'
                : '<i class="fa-solid fa-circle-xmark" style="color:#ef4444"></i> Backend Offline — Start server';
        }
    });
});
