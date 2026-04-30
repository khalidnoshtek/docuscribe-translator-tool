document.addEventListener('DOMContentLoaded', () => {
    const textArea = document.getElementById('text-input');
    const wordCount = document.getElementById('word-count');
    const docStatus = document.getElementById('doc-status');
    const btnAiGenerate = document.getElementById('btn-ai-generate');
    const downloadSection = document.getElementById('download-section');
    const btnDownloadWord = document.getElementById('btn-download-word');
    const btnDownloadPdf = document.getElementById('btn-download-pdf');
    const loader = document.getElementById('loader');

    let currentDocData = null;

    const getDocxLib = () => window.docx || (typeof docx !== 'undefined' ? docx : null);

    // Theme palettes — AI picks one based on doc type
    const THEMES = {
        legal:    { primary: '1A2332', accent: '8B6F47', rule: '2C3E50', name: 'Legal Navy' },
        property: { primary: '2D4A2B', accent: 'B8860B', rule: '3D5A3B', name: 'Property Forest' },
        nikah:    { primary: '4A1942', accent: 'C9A961', rule: '6B2960', name: 'Nikah Plum & Gold' },
        notary:   { primary: '1F1F1F', accent: '8B0000', rule: '333333', name: 'Notary Classic' },
        academic: { primary: '0F3057', accent: '988558', rule: '1F4E79', name: 'Academic Royal' },
        medical:  { primary: '0E5C7A', accent: '5B7C6A', rule: '17708F', name: 'Medical Teal' },
        affidavit:{ primary: '2B2B2B', accent: '7B341E', rule: '4A4A4A', name: 'Affidavit Sepia' },
    };

    const getTheme = (key) => THEMES[key] || THEMES.legal;

    textArea.addEventListener('input', () => {
        const text = textArea.value;
        const words = text.trim() ? text.trim().split(/\s+/).length : 0;
        wordCount.textContent = `${words} word${words !== 1 ? 's' : ''}`;
    });

    btnAiGenerate.addEventListener('click', async () => {
        const text = textArea.value;
        if (!text.trim()) return;

        let apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) {
            const key = prompt('Please enter your Gemini API Key:');
            if (key) { localStorage.setItem('gemini_api_key', key); apiKey = key; }
            else return;
        }

        loader.classList.remove('hidden');
        downloadSection.classList.add('hidden');

        try {
            const promptText = `You are an elite legal document architect.
Reconstruct the following translated text into a prestigious, official document in elite legal English.
Choose the BEST visual style based on the document type. Available themes:
- "legal" (general legal, contracts)
- "property" (property deeds, sale agreements, land documents)
- "nikah" (nikah-nama, marriage certificates, Islamic documents)
- "notary" (notarized statements, attestations)
- "academic" (degrees, transcripts, certificates)
- "medical" (medical records, certificates)
- "affidavit" (affidavits, sworn statements)

Choose layout that best fits:
- "classic" (centered, formal, single column)
- "two-column-sig" (signatures side-by-side at bottom)
- "decorative" (with ornamental dividers, suitable for nikah/certificates)
- "minimal-elite" (sparse, premium, lots of whitespace)

Return ONLY a valid JSON object — no markdown fences.

{
  "docType": "Short Category (e.g. 'Sale Deed', 'Nikah Nama', 'Affidavit')",
  "theme": "one of: legal, property, nikah, notary, academic, medical, affidavit",
  "layout": "one of: classic, two-column-sig, decorative, minimal-elite",
  "title": "PRESTIGIOUS OFFICIAL TITLE",
  "subtitle": "Optional Latin/formal subtitle or empty string",
  "reference": "Ref: DOC-${Math.floor(Math.random()*9000)+1000}",
  "header": "OFFICIAL CERTIFIED TRANSLATION",
  "preamble": "1-2 line opening statement in formal legal prose, or empty string",
  "contentSections": [{ "title": "Heading", "text": "Detailed elite legal content" }],
  "signatures": [{"label": "Role/Designation", "name": "Full Name"}],
  "footer": "Certification / authentication text"
}

Text to process:
${text}`;

            const models = [
                'gemini-2.5-flash',
                'gemini-2.5-pro',
                'gemini-flash-latest',
                'gemini-pro-latest',
                'gemini-2.0-flash',
                'gemini-2.0-flash-001'
            ];

            let jsonResponse = null;
            let success = false;
            let lastError = '';

            for (const model of models) {
                try {
                    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] })
                    });

                    if (response.ok) {
                        const data = await response.json();
                        let rawText = data.candidates[0].content.parts[0].text;
                        const jsonMatch = rawText.match(/\{[\s\S]*\}/);
                        if (jsonMatch) {
                            jsonResponse = JSON.parse(jsonMatch[0]);
                            success = true;
                            break;
                        }
                    } else {
                        const err = await response.json();
                        lastError = err.error?.message || response.statusText;
                    }
                } catch (e) { lastError = e.message; }
            }

            if (!success) throw new Error(lastError || 'API Service Error');

            currentDocData = jsonResponse;
            docStatus.textContent = `Draft Created: ${jsonResponse.docType}`;
            downloadSection.classList.remove('hidden');

        } catch (error) {
            alert('Drafting Failed: ' + error.message);
        } finally {
            loader.classList.add('hidden');
        }
    });

    // ===== WORD DOWNLOAD =====
    btnDownloadWord.addEventListener('click', () => {
        if (!currentDocData) return;
        try {
            const lib = getDocxLib();
            if (!lib) throw new Error("Library 'docx' not detected. Please reload.");

            const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType } = lib;
            const theme = getTheme(currentDocData.theme);
            const layout = currentDocData.layout || 'classic';
            const children = [];

            // Header band
            children.push(new Paragraph({
                children: [new TextRun({ text: currentDocData.header || '', bold: true, size: 18, color: theme.accent, characterSpacing: 80 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 80 }
            }));

            // Decorative divider for "decorative" layout
            if (layout === 'decorative') {
                children.push(new Paragraph({
                    children: [new TextRun({ text: '❖ ⸻⸻⸻⸻ ❖ ⸻⸻⸻⸻ ❖', size: 20, color: theme.accent })],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 }
                }));
            }

            // Reference (right-aligned)
            children.push(new Paragraph({
                children: [new TextRun({ text: currentDocData.reference || '', size: 18, color: '888888', italics: true })],
                alignment: AlignmentType.RIGHT,
                spacing: { after: layout === 'minimal-elite' ? 800 : 400 }
            }));

            // Title
            children.push(new Paragraph({
                children: [new TextRun({
                    text: currentDocData.title || '',
                    bold: true,
                    size: layout === 'minimal-elite' ? 44 : 38,
                    color: theme.primary,
                    allCaps: true
                })],
                alignment: AlignmentType.CENTER,
                spacing: { after: currentDocData.subtitle ? 100 : 400 }
            }));

            if (currentDocData.subtitle) {
                children.push(new Paragraph({
                    children: [new TextRun({ text: currentDocData.subtitle, italics: true, size: 22, color: theme.accent })],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 }
                }));
            }

            // Decorative under-title rule
            if (layout === 'decorative') {
                children.push(new Paragraph({
                    children: [new TextRun({ text: '◈ ◈ ◈', size: 18, color: theme.accent })],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 500 }
                }));
            } else {
                children.push(new Paragraph({
                    children: [new TextRun({ text: '', size: 2 })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: theme.accent, space: 1 } },
                    spacing: { after: 500 }
                }));
            }

            // Preamble
            if (currentDocData.preamble) {
                children.push(new Paragraph({
                    children: [new TextRun({ text: currentDocData.preamble, italics: true, size: 22, color: '444444' })],
                    alignment: AlignmentType.JUSTIFIED,
                    indent: { left: 600, right: 600 },
                    spacing: { after: 400 }
                }));
            }

            // Content sections
            (currentDocData.contentSections || []).forEach((s, idx) => {
                children.push(new Paragraph({
                    children: [
                        new TextRun({ text: `${String(idx+1).padStart(2,'0')}. `, bold: true, size: 22, color: theme.accent }),
                        new TextRun({ text: (s.title || '').toUpperCase(), bold: true, size: 22, color: theme.primary, characterSpacing: 40 })
                    ],
                    spacing: { before: 280, after: 140 },
                    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: theme.rule } }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: s.text || '', size: 22, color: '222222' })],
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { after: 220, line: 340 }
                }));
            });

            // Signatures
            const sigs = currentDocData.signatures || [];
            if (sigs.length) {
                children.push(new Paragraph({ children: [new TextRun({ text: '', size: 2 })], spacing: { before: 600 } }));

                if (layout === 'two-column-sig' || sigs.length > 1) {
                    const rows = [];
                    for (let i = 0; i < sigs.length; i += 2) {
                        const makeCell = (sig) => sig ? new TableCell({
                            children: [
                                new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 800 } }),
                                new Paragraph({
                                    children: [new TextRun({ text: '_______________________________', color: theme.rule })],
                                    spacing: { after: 60 }
                                }),
                                new Paragraph({
                                    children: [new TextRun({ text: sig.name || '', bold: true, size: 20, color: theme.primary })]
                                }),
                                new Paragraph({
                                    children: [new TextRun({ text: sig.label || '', size: 18, color: theme.accent, italics: true })]
                                })
                            ],
                            margins: { top: 200, bottom: 200, left: 200, right: 200 },
                            borders: { top: {style:BorderStyle.NONE}, bottom: {style:BorderStyle.NONE}, left: {style:BorderStyle.NONE}, right: {style:BorderStyle.NONE} }
                        }) : new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: '' })] })],
                            borders: { top: {style:BorderStyle.NONE}, bottom: {style:BorderStyle.NONE}, left: {style:BorderStyle.NONE}, right: {style:BorderStyle.NONE} }
                        });
                        rows.push(new TableRow({ children: [makeCell(sigs[i]), makeCell(sigs[i+1])] }));
                    }
                    children.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
                } else {
                    const sig = sigs[0];
                    children.push(new Paragraph({
                        children: [new TextRun({ text: '_______________________________', color: theme.rule })],
                        alignment: AlignmentType.RIGHT,
                        spacing: { before: 800, after: 60 }
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: sig.name || '', bold: true, size: 22, color: theme.primary })],
                        alignment: AlignmentType.RIGHT
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: sig.label || '', size: 18, color: theme.accent, italics: true })],
                        alignment: AlignmentType.RIGHT
                    }));
                }
            }

            // Footer
            if (currentDocData.footer) {
                children.push(new Paragraph({
                    children: [new TextRun({ text: '', size: 2 })],
                    border: { top: { style: BorderStyle.SINGLE, size: 4, color: theme.accent, space: 1 } },
                    spacing: { before: 1200, after: 200 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: currentDocData.footer, size: 16, color: '666666', italics: true })],
                    alignment: AlignmentType.CENTER
                }));
            }

            const doc = new Document({
                creator: 'DocuScribe',
                title: currentDocData.title,
                styles: {
                    default: {
                        document: { run: { font: 'Garamond' } }
                    }
                },
                sections: [{
                    properties: {
                        page: {
                            margin: { top: 1700, right: 1500, bottom: 1700, left: 1500 },
                            borders: {
                                pageBorderLeft:   { style: BorderStyle.SINGLE, size: 8, color: theme.primary, space: 24 },
                                pageBorderRight:  { style: BorderStyle.SINGLE, size: 8, color: theme.primary, space: 24 },
                                pageBorderTop:    { style: BorderStyle.SINGLE, size: 8, color: theme.primary, space: 24 },
                                pageBorderBottom: { style: BorderStyle.SINGLE, size: 8, color: theme.primary, space: 24 },
                            }
                        }
                    },
                    children
                }]
            });

            Packer.toBlob(doc).then(blob => {
                saveAs(blob, `${(currentDocData.docType || 'Document').replace(/\s+/g, '_')}_Official.docx`);
            });
        } catch (e) { alert('Word Error: ' + e.message); }
    });

    // ===== PDF DOWNLOAD =====
    btnDownloadPdf.addEventListener('click', async () => {
        if (!currentDocData) return;
        const theme = getTheme(currentDocData.theme);
        const layout = currentDocData.layout || 'classic';

        const wrap = document.createElement('div');
        wrap.id = 'pdf-render-target';
        // Position offscreen but in normal flow so html2canvas reliably renders
        wrap.style.cssText = `
            position: absolute;
            top: 0;
            left: 0;
            width: 794px;
            background: #ffffff;
            color: #1a1a1a;
            font-family: 'Garamond', 'Times New Roman', Georgia, serif;
            box-sizing: border-box;
            z-index: -1;
            opacity: 1;
        `;

        const decorTop = layout === 'decorative'
            ? `<div style="text-align:center; color:#${theme.accent}; font-size:18px; letter-spacing:8px; margin:8px 0 24px;">❖ &nbsp; ⸻⸻⸻⸻ &nbsp; ❖ &nbsp; ⸻⸻⸻⸻ &nbsp; ❖</div>`
            : '';

        const subtitleHtml = currentDocData.subtitle
            ? `<div style="text-align:center; font-style:italic; color:#${theme.accent}; font-size:18px; margin-top:6px; letter-spacing:1px;">${currentDocData.subtitle}</div>`
            : '';

        const preambleHtml = currentDocData.preamble
            ? `<p style="font-style:italic; color:#444; text-align:justify; padding:0 32px; margin:24px 0 32px; font-size:16px; line-height:1.7;">${currentDocData.preamble}</p>`
            : '';

        const sectionsHtml = (currentDocData.contentSections || []).map((s, i) => `
            <div style="margin-bottom: 28px;">
                <div style="border-bottom:1.5px solid #${theme.rule}; padding-bottom:6px; margin-bottom:12px;">
                    <span style="color:#${theme.accent}; font-weight:700; font-size:15px; letter-spacing:2px;">${String(i+1).padStart(2,'0')}.</span>
                    <span style="color:#${theme.primary}; font-weight:700; font-size:15px; letter-spacing:2px; text-transform:uppercase; margin-left:8px;">${s.title || ''}</span>
                </div>
                <p style="text-align:justify; margin:0; font-size:15.5px; line-height:1.75; color:#222;">${s.text || ''}</p>
            </div>
        `).join('');

        const sigs = currentDocData.signatures || [];
        let sigHtml = '';
        if (sigs.length) {
            const useTwoCol = layout === 'two-column-sig' || sigs.length > 1;
            if (useTwoCol) {
                const rows = [];
                for (let i = 0; i < sigs.length; i += 2) {
                    rows.push(`<tr>${sigs.slice(i, i+2).map(s => `
                        <td style="width:50%; padding:24px 16px; vertical-align:top;">
                            <div style="margin-top:60px; border-top:1px solid #${theme.rule}; padding-top:8px;">
                                <div style="font-weight:700; color:#${theme.primary}; font-size:15px;">${s.name || ''}</div>
                                <div style="font-style:italic; color:#${theme.accent}; font-size:13px; margin-top:2px;">${s.label || ''}</div>
                            </div>
                        </td>`).join('')}${sigs.length % 2 && i + 1 >= sigs.length ? '<td></td>' : ''}</tr>`);
                }
                sigHtml = `<table style="width:100%; margin-top:40px; border-collapse:collapse;">${rows.join('')}</table>`;
            } else {
                const s = sigs[0];
                sigHtml = `
                    <div style="margin-top:80px; text-align:right;">
                        <div style="display:inline-block; min-width:240px; border-top:1px solid #${theme.rule}; padding-top:8px;">
                            <div style="font-weight:700; color:#${theme.primary}; font-size:15px;">${s.name || ''}</div>
                            <div style="font-style:italic; color:#${theme.accent}; font-size:13px;">${s.label || ''}</div>
                        </div>
                    </div>`;
            }
        }

        wrap.innerHTML = `
            <div style="border: 3px double #${theme.primary}; padding: 56px 60px; min-height: 1050px; position: relative;">
                <div style="position:absolute; top:18px; left:24px; right:24px; border-top:1px solid #${theme.accent}33;"></div>
                <div style="text-align:center; color:#${theme.accent}; font-size:11px; letter-spacing:6px; font-weight:700; margin-bottom:6px;">${currentDocData.header || ''}</div>
                ${decorTop}
                <div style="text-align:right; color:#888; font-size:11px; font-style:italic; margin-bottom:32px;">${currentDocData.reference || ''}</div>
                <h1 style="text-align:center; font-size:30px; font-weight:700; color:#${theme.primary}; margin:0; letter-spacing:3px; text-transform:uppercase;">${currentDocData.title || ''}</h1>
                ${subtitleHtml}
                ${layout === 'decorative'
                    ? `<div style="text-align:center; color:#${theme.accent}; margin:18px 0 28px; font-size:14px; letter-spacing:6px;">◈ &nbsp; ◈ &nbsp; ◈</div>`
                    : `<div style="height:2px; background:#${theme.accent}; width:120px; margin:18px auto 32px;"></div>`}
                ${preambleHtml}
                ${sectionsHtml}
                ${sigHtml}
                ${currentDocData.footer ? `
                    <div style="margin-top:60px; border-top:1px solid #${theme.accent}; padding-top:14px; text-align:center;">
                        <em style="font-size:11px; color:#666;">${currentDocData.footer}</em>
                    </div>` : ''}
                <div style="position:absolute; bottom:18px; left:24px; right:24px; border-bottom:1px solid #${theme.accent}33;"></div>
            </div>
        `;

        document.body.appendChild(wrap);

        // Wait one frame for layout
        await new Promise(r => requestAnimationFrame(() => r()));
        await new Promise(r => setTimeout(r, 200));

        const opt = {
            margin: 0,
            filename: `${(currentDocData.docType || 'Document').replace(/\s+/g, '_')}_Official.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
            pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
        };

        try {
            await html2pdf().set(opt).from(wrap).save();
        } catch (err) {
            alert('PDF Error: ' + err.message);
        } finally {
            if (wrap.parentNode) wrap.parentNode.removeChild(wrap);
        }
    });
});
