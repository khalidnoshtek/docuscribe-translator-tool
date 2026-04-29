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

    // Update word count
    textArea.addEventListener('input', () => {
        const text = textArea.value;
        const words = text.trim() ? text.trim().split(/\s+/).length : 0;
        wordCount.textContent = `${words} word${words !== 1 ? 's' : ''}`;
    });

    // AI Generate Document Structure
    btnAiGenerate.addEventListener('click', async () => {
        const text = textArea.value;
        if (!text.trim()) return;
        
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) {
            alert('API Key missing. Run: localStorage.setItem("gemini_api_key", "YOUR_KEY") in console.');
            return;
        }

        loader.querySelector('p').textContent = 'AI is formatting your official document...';
        loader.classList.remove('hidden');
        downloadSection.classList.add('hidden');

        try {
            const prompt = `You are a master legal and official document drafter.
I will provide you with a translated text of an old document.
Your task:
1. Identify the document type and a professional official title.
2. Draft a beautifully structured, formal English version.
3. Return ONLY a JSON object.
Structure:
{
  "docType": "Short Type",
  "title": "FULL OFFICIAL TITLE IN CAPS",
  "header": "Header info",
  "sections": [{ "heading": "...", "content": "..." }],
  "signatures": ["Party 1", "Party 2", "Witness 1", "Witness 2"],
  "footer": "Official footer"
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
                            contents: [{ parts: [{ text: prompt }] }]
                        })
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
                } catch (e) {
                    lastError = e.message;
                }
            }

            if (!success) throw new Error(lastError || 'AI failed to generate document.');
            
            currentDocData = jsonResponse;
            docStatus.textContent = `Generated: ${jsonResponse.docType}`;
            downloadSection.classList.remove('hidden');
            
        } catch (error) {
            console.error('Generation Error:', error);
            alert('Error: ' + error.message);
        } finally {
            loader.classList.add('hidden');
        }
    });

    // Word Download
    btnDownloadWord.addEventListener('click', () => {
        if (!currentDocData) return;
        generateWordFile(currentDocData);
    });

    // PDF Download
    btnDownloadPdf.addEventListener('click', () => {
        if (!currentDocData) return;
        generatePdfFile(currentDocData);
    });

    const generateWordFile = (data) => {
        const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType } = window.docx;
        const children = [];

        children.push(new Paragraph({
            children: [new TextRun({ text: data.title, bold: true, size: 32 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));

        children.push(new Paragraph({
            children: [new TextRun({ text: data.header, italic: true, size: 20 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 }
        }));

        data.sections.forEach(s => {
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

        const rows = [];
        for (let i = 0; i < data.signatures.length; i += 2) {
            const cells = [
                new TableCell({
                    children: [
                        new Paragraph({ spacing: { before: 800 } }),
                        new Paragraph({
                            border: { top: { style: BorderStyle.SINGLE, size: 1 } },
                            children: [new TextRun({ text: data.signatures[i], size: 20 })],
                            spacing: { before: 100 }
                        })
                    ],
                    borders: { top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL }, left: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL } }
                })
            ];
            if (data.signatures[i + 1]) {
                cells.push(new TableCell({
                    children: [
                        new Paragraph({ spacing: { before: 800 } }),
                        new Paragraph({
                            border: { top: { style: BorderStyle.SINGLE, size: 1 } },
                            children: [new TextRun({ text: data.signatures[i + 1], size: 20 })],
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

        children.push(new Table({
            rows: rows,
            width: { size: 100, type: WidthType.PERCENTAGE },
        }));

        children.push(new Paragraph({
            children: [new TextRun({ text: data.footer, size: 16, color: "666666" })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 1000 }
        }));

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
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
            saveAs(blob, `${data.docType.replace(/\s+/g, '_')}_Official.docx`);
        });
    };

    const generatePdfFile = (data) => {
        // Create in-memory element for PDF generation
        const temp = document.createElement('div');
        temp.style.width = '8.5in';
        temp.style.padding = '0.75in';
        temp.style.background = 'white';
        temp.style.color = 'black';
        temp.style.fontFamily = 'serif';
        temp.style.lineHeight = '1.5';
        temp.style.border = '2px solid black';
        
        let sectionsHtml = data.sections.map(s => `
            <div style="margin-bottom: 1.5rem;">
                ${s.heading ? `<h4 style="margin-bottom: 0.5rem; text-decoration: underline;">${s.heading}</h4>` : ''}
                <p style="text-align: justify; margin: 0;">${s.content}</p>
            </div>
        `).join('');

        let sigHtml = data.signatures.map(s => `
            <div style="margin-top: 3rem; border-top: 1px solid black; width: 45%; padding-top: 0.5rem; font-size: 0.9rem;">${s}</div>
        `).join('');

        temp.innerHTML = `
            <h2 style="text-align: center; margin-bottom: 0.5rem; font-size: 1.6rem;">${data.title}</h2>
            <p style="text-align: center; font-style: italic; font-size: 0.9rem; margin-bottom: 2rem;">${data.header}</p>
            ${sectionsHtml}
            <div style="display: flex; flex-wrap: wrap; gap: 10%; margin-top: 2rem;">
                ${sigHtml}
            </div>
            <div style="margin-top: 4rem; text-align: center; font-size: 0.8rem; color: #555;">
                ${data.footer}
            </div>
        `;

        const opt = {
            margin:       0.5,
            filename:     `${data.docType.replace(/\s+/g, '_')}_Official.pdf`,
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2 },
            jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
        };
        html2pdf().set(opt).from(temp).save();
    };
});
