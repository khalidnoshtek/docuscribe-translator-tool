document.addEventListener('DOMContentLoaded', () => {
    const textArea = document.getElementById('text-input');
    const wordCount = document.getElementById('word-count');
    const docStatus = document.getElementById('doc-status');
    const btnAiGenerate = document.getElementById('btn-ai-generate');
    const previewContent = document.getElementById('preview-content');
    const downloadActions = document.getElementById('download-actions');
    const btnDownloadWord = document.getElementById('btn-download-word');
    const btnDownloadPdf = document.getElementById('btn-download-pdf');
    const loader = document.getElementById('loader');

    let currentDocData = null;

    // Update stats
    textArea.addEventListener('input', () => {
        const text = textArea.value;
        const words = text.trim() ? text.trim().split(/\s+/).length : 0;
        wordCount.textContent = `${words} word${words !== 1 ? 's' : ''}`;
    });

    // AI Generate Official Document
    btnAiGenerate.addEventListener('click', async () => {
        const text = textArea.value;
        if (!text.trim()) return;
        
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) {
            alert('Please enter and save a Gemini API Key in the settings first. (You can do this in the browser console by running: localStorage.setItem("gemini_api_key", "YOUR_KEY"))');
            return;
        }

        loader.querySelector('p').textContent = 'AI is drafting your official document...';
        loader.classList.remove('hidden');

        try {
            const prompt = `You are a master legal and official document drafter.
I will provide you with a translated text of an old document (possibly a Property Deed, Nikah Nama/Marriage Contract, or Official Agreement).
Your task:
1. Identify the document type and a professional official title for it.
2. Draft a beautifully structured, formal English version of this document.
3. Use formal legal language where appropriate.
4. Return ONLY a JSON object with this exact structure:
{
  "docType": "Short Type",
  "title": "FULL OFFICIAL TITLE IN CAPS",
  "header": "Header info (like dates, reference numbers, or parties)",
  "sections": [
    { "heading": "Heading Title or Section Number", "content": "The actual text content for this section." }
  ],
  "signatures": ["Party 1 Name/Title", "Party 2 Name/Title", "Witness 1", "Witness 2"],
  "footer": "Official footer or certification text"
}

Text to process:
${text}`;

            const models = ['gemini-3-flash', 'gemini-3-pro', 'gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-1.5-flash'];
            let jsonResponse = null;
            let success = false;
            let lastError = '';

            for (const model of models) {
                try {
                    const response = await fetch(`https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${apiKey}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            contents: [{ parts: [{ text: prompt }] }],
                            generationConfig: { response_mime_type: "application/json" }
                        })
                    });

                    if (response.ok) {
                        const data = await response.json();
                        const rawJson = data.candidates[0].content.parts[0].text;
                        jsonResponse = JSON.parse(rawJson);
                        success = true;
                        break;
                    } else {
                        const err = await response.json();
                        lastError = err.error?.message || response.statusText;
                    }
                } catch (e) {
                    lastError = e.message;
                }
            }

            if (!success) throw new Error(lastError || 'AI failed to generate document.');
            
            currentDocData = jsonResponse;
            renderPreview(jsonResponse);
            downloadActions.classList.remove('hidden');
            docStatus.textContent = `Generated: ${jsonResponse.docType}`;
            
        } catch (error) {
            console.error('Generation Error:', error);
            alert('Failed to generate official document: ' + error.message);
        } finally {
            loader.classList.add('hidden');
        }
    });

    const renderPreview = (data) => {
        let sectionsHtml = data.sections.map(s => `
            <div class="preview-section">
                ${s.heading ? `<h4>${s.heading}</h4>` : ''}
                <p>${s.content}</p>
            </div>
        `).join('');

        let sigHtml = data.signatures.map(s => `
            <div class="signature-line">${s}</div>
        `).join('');

        previewContent.innerHTML = `
            <div class="paper-preview">
                <h2>${data.title}</h2>
                <p style="text-align: center; font-style: italic; font-size: 0.85rem;">${data.header}</p>
                <hr style="margin: 1rem 0; border: none; border-top: 1px double #333;">
                ${sectionsHtml}
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-top: 2rem;">
                    ${sigHtml}
                </div>
                <div style="margin-top: 3rem; text-align: center; font-size: 0.75rem; border-top: 1px solid #eee; padding-top: 1rem;">
                    ${data.footer}
                </div>
            </div>
        `;
    };

    // Download Word
    btnDownloadWord.addEventListener('click', () => {
        if (!currentDocData) return;

        const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType } = docx;
        
        const children = [];

        // Title
        children.push(new Paragraph({
            children: [new TextRun({ text: currentDocData.title, bold: true, size: 32, font: "Playfair Display" })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));

        // Header
        children.push(new Paragraph({
            children: [new TextRun({ text: currentDocData.header, italic: true, size: 20 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 }
        }));

        // Sections
        currentDocData.sections.forEach(s => {
            if (s.heading) {
                children.push(new Paragraph({
                    children: [new TextRun({ text: s.heading, bold: true, size: 24, underline: {} })],
                    spacing: { before: 200, after: 120 }
                }));
            }
            children.push(new Paragraph({
                children: [new TextRun({ text: s.content, size: 24 })],
                alignment: AlignmentType.JUSTIFIED,
                spacing: { after: 200 }
            }));
        });

        // Signatures Table
        const rows = [];
        for (let i = 0; i < currentDocData.signatures.length; i += 2) {
            const cells = [
                new TableCell({
                    children: [
                        new Paragraph({ spacing: { before: 800 } }),
                        new Paragraph({
                            border: { top: { style: BorderStyle.SINGLE, size: 1 } },
                            children: [new TextRun({ text: currentDocData.signatures[i], size: 20 })],
                            spacing: { before: 100 }
                        })
                    ],
                    borders: { top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL }, left: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL } }
                })
            ];

            if (currentDocData.signatures[i + 1]) {
                cells.push(new TableCell({
                    children: [
                        new Paragraph({ spacing: { before: 800 } }),
                        new Paragraph({
                            border: { top: { style: BorderStyle.SINGLE, size: 1 } },
                            children: [new TextRun({ text: currentDocData.signatures[i + 1], size: 20 })],
                            spacing: { before: 100 }
                        })
                    ],
                    borders: { top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL }, left: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL } }
                }));
            } else {
                cells.push(new TableCell({ children: [] }));
            }

            rows.push(new TableRow({ children: cells }));
        }

        const sigTable = new Table({
            rows: rows,
            width: { size: 100, type: WidthType.PERCENTAGE },
        });

        children.push(sigTable);

        // Footer
        children.push(new Paragraph({
            children: [new TextRun({ text: currentDocData.footer, size: 16, color: "666666" })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 1000 }
        }));

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, // 1 inch
                        borders: {
                            pageBorderLeft: { style: BorderStyle.SINGLE, size: 12, space: 24, color: "000000" },
                            pageBorderRight: { style: BorderStyle.SINGLE, size: 12, space: 24, color: "000000" },
                            pageBorderTop: { style: BorderStyle.SINGLE, size: 12, space: 24, color: "000000" },
                            pageBorderBottom: { style: BorderStyle.SINGLE, size: 12, space: 24, color: "000000" },
                        }
                    }
                },
                children: children
            }]
        });

        Packer.toBlob(doc).then(blob => {
            saveAs(blob, `${currentDocData.docType.replace(/\s+/g, '_')}_Official.docx`);
        });
    });

    // Download PDF
    btnDownloadPdf.addEventListener('click', () => {
        if (!currentDocData) return;
        const element = previewContent.querySelector('.paper-preview');
        const opt = {
            margin:       0.5,
            filename:     `${currentDocData.docType.replace(/\s+/g, '_')}_Official.pdf`,
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2 },
            jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
        };
        html2pdf().set(opt).from(element).save();
    });
});
