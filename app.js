document.addEventListener('DOMContentLoaded', () => {
    const textArea = document.getElementById('text-input');
    const wordCount = document.getElementById('word-count');
    const docStatus = document.getElementById('doc-status');
    const btnAiGenerate = document.getElementById('btn-ai-generate');
    const downloadSection = document.getElementById('download-section');
    const btnDownloadWord = document.getElementById('btn-download-word');
    const btnShareWhatsapp = document.getElementById('btn-share-whatsapp');
    const btnShareEmail = document.getElementById('btn-share-email');
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

    // Toast popup helper
    function showToast(message, type = 'info') {
        const existing = document.querySelector('.toast');
        if (existing) existing.remove();
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.textContent = message;
        document.body.appendChild(toast);
        // Trigger animation
        requestAnimationFrame(() => toast.classList.add('toast-show'));
        setTimeout(() => {
            toast.classList.remove('toast-show');
            setTimeout(() => toast.remove(), 300);
        }, 3500);
    }

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
Reconstruct the following rough text into a prestigious, official document in elite legal English.
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
  "header": "OFFICIAL CERTIFIED DOCUMENT",
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

            // Smooth scroll to share/download section + show success popup
            setTimeout(() => {
                downloadSection.scrollIntoView({ behavior: 'smooth', block: 'center' });
                showToast(`✓ ${jsonResponse.docType} drafted — choose Word, PDF or share below.`, 'success');
            }, 250);

        } catch (error) {
            alert('Drafting Failed: ' + error.message);
        } finally {
            loader.classList.add('hidden');
        }
    });

    // ===== WORD BLOB BUILDER =====
    async function buildWordBlob() {
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
                creator: 'Sanad',
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

        return await Packer.toBlob(doc);
    }

    function wordFilename() {
        return `${(currentDocData?.docType || 'Document').replace(/\s+/g, '_')}_Official.docx`;
    }

    btnDownloadWord.addEventListener('click', async () => {
        if (!currentDocData) return;
        try {
            const blob = await buildWordBlob();
            saveAs(blob, wordFilename());
            showToast('✓ Word document downloaded', 'success');
        } catch (e) { alert('Word Error: ' + e.message); }
    });

    // Build a clean, multi-page PDF using jsPDF text rendering (no html2canvas).
    // This is reliable, small, and shareable on WhatsApp/Email.
    function buildSharePdfBlob() {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait', compress: true });
        const d = currentDocData;
        const theme = getTheme(d.theme);
        const primaryRgb = hexToRgb(theme.primary);
        const accentRgb = hexToRgb(theme.accent);

        const pageW = 210, pageH = 297;
        const margin = 18;
        const usableW = pageW - margin * 2;
        let y = margin;

        const ensureSpace = (need) => {
            if (y + need > pageH - margin) {
                pdf.addPage();
                y = margin;
            }
        };

        const writeText = (text, opts = {}) => {
            const {
                size = 11, font = 'times', style = 'normal',
                color = [40, 40, 40], align = 'left', leading = 1.45,
                spaceAfter = 3
            } = opts;
            pdf.setFont(font, style);
            pdf.setFontSize(size);
            pdf.setTextColor(color[0], color[1], color[2]);
            const lines = pdf.splitTextToSize(text, usableW);
            const lineH = (size * 0.3528) * leading; // pt → mm
            for (const line of lines) {
                ensureSpace(lineH);
                let x = margin;
                if (align === 'center') x = pageW / 2;
                else if (align === 'right') x = pageW - margin;
                pdf.text(line, x, y + lineH * 0.75, { align });
                y += lineH;
            }
            y += spaceAfter;
        };

        const drawRule = (color, weight = 0.4, width = usableW) => {
            ensureSpace(2);
            pdf.setDrawColor(color[0], color[1], color[2]);
            pdf.setLineWidth(weight);
            const x = (pageW - width) / 2;
            pdf.line(x, y, x + width, y);
            y += 4;
        };

        // Header
        if (d.header) {
            writeText(d.header.toUpperCase(), {
                size: 8, style: 'bold', color: accentRgb, align: 'center', spaceAfter: 2
            });
        }
        drawRule(accentRgb, 0.3);

        // Reference (right)
        if (d.reference) {
            writeText(d.reference, {
                size: 8, style: 'italic', color: [120, 120, 120], align: 'right', spaceAfter: 6
            });
        }

        // Title
        y += 4;
        writeText((d.title || '').toUpperCase(), {
            size: 18, style: 'bold', color: primaryRgb, align: 'center', leading: 1.2, spaceAfter: 2
        });
        if (d.subtitle) {
            writeText(d.subtitle, {
                size: 11, style: 'italic', color: accentRgb, align: 'center', spaceAfter: 4
            });
        }

        // Decorative rule under title
        ensureSpace(6);
        pdf.setDrawColor(accentRgb[0], accentRgb[1], accentRgb[2]);
        pdf.setLineWidth(0.6);
        pdf.line(pageW/2 - 20, y, pageW/2 + 20, y);
        y += 8;

        // Preamble (italic)
        if (d.preamble) {
            writeText(d.preamble, {
                size: 11, style: 'italic', color: [80, 80, 80], align: 'left', leading: 1.55, spaceAfter: 5
            });
        }

        // Sections
        (d.contentSections || []).forEach((s, i) => {
            ensureSpace(12);
            writeText(`${String(i+1).padStart(2,'0')}.  ${(s.title || '').toUpperCase()}`, {
                size: 11, style: 'bold', color: primaryRgb, leading: 1.3, spaceAfter: 1
            });
            // Underline rule
            pdf.setDrawColor(220, 215, 200);
            pdf.setLineWidth(0.3);
            pdf.line(margin, y, margin + usableW, y);
            y += 3;
            writeText(s.text || '', {
                size: 11, color: [40, 40, 40], leading: 1.55, spaceAfter: 5
            });
        });

        // Signatures
        const sigs = d.signatures || [];
        if (sigs.length) {
            y += 8;
            ensureSpace(28);
            const colWidth = usableW / 2 - 6;
            sigs.forEach((s, i) => {
                const col = i % 2;
                const row = Math.floor(i / 2);
                if (col === 0 && row > 0) {
                    y += 22;
                    ensureSpace(28);
                }
                const x = margin + col * (colWidth + 12);
                const sigY = y;
                pdf.setDrawColor(150, 150, 150);
                pdf.setLineWidth(0.3);
                pdf.line(x, sigY + 14, x + colWidth, sigY + 14);
                pdf.setFont('times', 'bold');
                pdf.setFontSize(11);
                pdf.setTextColor(primaryRgb[0], primaryRgb[1], primaryRgb[2]);
                pdf.text(s.name || '', x, sigY + 18);
                pdf.setFont('times', 'italic');
                pdf.setFontSize(9);
                pdf.setTextColor(accentRgb[0], accentRgb[1], accentRgb[2]);
                pdf.text(s.label || '', x, sigY + 22);
            });
            y += 26;
        }

        // Footer
        if (d.footer) {
            ensureSpace(12);
            y += 4;
            drawRule(accentRgb, 0.3, 80);
            writeText(d.footer, {
                size: 8, style: 'italic', color: [110, 110, 110], align: 'center', leading: 1.4, spaceAfter: 0
            });
        }

        return pdf.output('blob');
    }

    function hexToRgb(hex) {
        const h = hex.replace('#', '');
        return [
            parseInt(h.slice(0, 2), 16),
            parseInt(h.slice(2, 4), 16),
            parseInt(h.slice(4, 6), 16)
        ];
    }

    function pdfShareFilename() {
        return `${(currentDocData?.docType || 'Document').replace(/\s+/g, '_')}_Official.pdf`;
    }

    // Build a plain-text version of the document for sharing as message body
    function buildPlainText() {
        const d = currentDocData;
        const lines = [];
        if (d.header) lines.push(d.header.toUpperCase());
        if (d.reference) lines.push(d.reference);
        lines.push('');
        if (d.title) lines.push(d.title.toUpperCase());
        if (d.subtitle) lines.push(d.subtitle);
        lines.push('');
        if (d.preamble) { lines.push(d.preamble); lines.push(''); }
        (d.contentSections || []).forEach((s, i) => {
            lines.push(`${String(i+1).padStart(2,'0')}. ${(s.title || '').toUpperCase()}`);
            lines.push(s.text || '');
            lines.push('');
        });
        (d.signatures || []).forEach(s => {
            lines.push(`${s.label || ''}: ${s.name || ''}`);
        });
        if (d.footer) { lines.push(''); lines.push(d.footer); }
        lines.push('');
        lines.push('— Drafted with Sanad');
        return lines.join('\n');
    }

    // ===== SHARE =====
    async function shareDocument(target) {
        if (!currentDocData) return;

        const shareTitle = currentDocData.title || 'Official Document';
        const fullText = buildPlainText();

        // STEP 1: Build a PDF (WhatsApp accepts PDFs natively via Web Share API)
        try {
            const pdfBlob = buildSharePdfBlob();
            const pdfName = pdfShareFilename();
            const pdfFile = new File([pdfBlob], pdfName, { type: 'application/pdf' });

            if (navigator.canShare && navigator.canShare({ files: [pdfFile] })) {
                try {
                    await navigator.share({ files: [pdfFile], title: shareTitle });
                    showToast('✓ Shared successfully', 'success');
                    return;
                } catch (err) {
                    if (err.name === 'AbortError') return;
                }
            }

            // STEP 2: Final fallback — download PDF + deep-link into the installed app
            saveAs(pdfBlob, pdfName);
            const encoded = encodeURIComponent(fullText);
            const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);

            if (target === 'whatsapp') {
                showToast('✓ Document downloaded — opening WhatsApp. Attach file from your downloads.', 'success');
                // whatsapp:// scheme opens the installed app directly on iOS & Android
                const waApp = `whatsapp://send?text=${encoded}`;
                const waWeb = `https://wa.me/?text=${encoded}`;
                setTimeout(() => {
                    if (isMobile) {
                        // Try the app first; if not installed, fallback to web after a short delay
                        window.location.href = waApp;
                        setTimeout(() => { window.location.href = waWeb; }, 1500);
                    } else {
                        window.open(waWeb, '_blank');
                    }
                }, 600);
            } else {
                showToast('✓ Document downloaded — opening Email. Attach file from your downloads.', 'success');
                setTimeout(() => {
                    window.location.href = `mailto:?subject=${encodeURIComponent(shareTitle)}&body=${encoded}`;
                }, 600);
            }
        } catch (e) {
            alert('Share Error: ' + e.message);
        }
    }

    btnShareWhatsapp.addEventListener('click', () => shareDocument('whatsapp'));
    btnShareEmail.addEventListener('click', () => shareDocument('email'));
});
