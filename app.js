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

    // Library Check
    const getDocxLib = () => {
        // Try multiple global names used by different CDN versions
        return window.docx || (typeof docx !== 'undefined' ? docx : null);
    };

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
            const key = prompt('Please enter your Gemini API Key:');
            if (key) localStorage.setItem('gemini_api_key', key);
            else return;
        }

        loader.querySelector('p').textContent = 'AI is architecting your official document...';
        loader.classList.remove('hidden');
        downloadSection.classList.add('hidden');

        try {
            const promptText = `You are an elite legal document architect.
Reconstruct the following translated text into a prestigious, official document.
Maintain all original information but use elite legal English.
Return ONLY a valid JSON object.

JSON Structure:
{
  "docType": "Short Category",
  "title": "PRESTIGIOUS OFFICIAL TITLE",
  "reference": "Ref: DOC-${Math.floor(Math.random()*9000)+1000}",
  "header": "OFFICIAL CERTIFIED TRANSLATION",
  "contentSections": [{ "title": "Heading", "text": "Detailed content" }],
  "signatures": [{"label": "Role", "name": "Name"}],
  "footer": "Certification text"
}

Text to process:
${text}`;

            const models = ['gemini-3-flash', 'gemini-1.5-flash', 'gemini-1.5-pro'];
            let jsonResponse = null;
            let success = false;
            let lastError = '';

            for (const model of models) {
                try {
                    const response = await fetch(`https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${apiKey}`, {
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
                } catch (e) {
                    lastError = e.message;
                }
            }

            if (!success) throw new Error(lastError || 'AI Service Error');
            
            currentDocData = jsonResponse;
            docStatus.textContent = `Document Ready: ${jsonResponse.docType}`;
            downloadSection.classList.remove('hidden');
            
        } catch (error) {
            alert('Failed to generate document: ' + error.message);
        } finally {
            loader.classList.add('hidden');
        }
    });

    // Word Download
    btnDownloadWord.addEventListener('click', () => {
        if (!currentDocData) return;
        
        try {
            const lib = getDocxLib();
            if (!lib) throw new Error("Word generation library (docx) failed to load from CDN. Please refresh.");

            const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType } = lib;
            const children = [];

            // Professional Styles
            children.push(new Paragraph({
                children: [new TextRun({ text: currentDocData.header, bold: true, size: 20, color: "555555" })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }));

            children.push(new Paragraph({
                children: [new TextRun({ text: currentDocData.reference, size: 18, color: "888888" })],
                alignment: AlignmentType.RIGHT,
                spacing: { after: 300 }
            }));

            children.push(new Paragraph({
                children: [new TextRun({ text: currentDocData.title, bold: true, size: 36, underline: { type: "double" } })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 600 }
            }));

            currentDocData.contentSections.forEach(s => {
                children.push(new Paragraph({
                    children: [new TextRun({ text: s.title, bold: true, size: 24 })],
                    spacing: { before: 300, after: 150 },
                    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: s.text, size: 24 })],
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { after: 200 }
                }));
            });

            const rows = [];
            for (let i = 0; i < currentDocData.signatures.length; i += 2) {
                const cells = [
                    new TableCell({
                        children: [
                            new Paragraph({ spacing: { before: 1200 } }),
                            new Paragraph({
                                border: { top: { style: BorderStyle.SINGLE, size: 1 } },
                                children: [
                                    new TextRun({ text: currentDocData.signatures[i].label + ": ", bold: true, size: 18 }),
                                    new TextRun({ text: currentDocData.signatures[i].name, size: 18 })
                                ]
                            })
                        ],
                        borders: { top: BorderStyle.NIL, bottom: BorderStyle.NIL, left: BorderStyle.NIL, right: BorderStyle.NIL }
                    })
                ];
                if (currentDocData.signatures[i + 1]) {
                    cells.push(new TableCell({
                        children: [
                            new Paragraph({ spacing: { before: 1200 } }),
                            new Paragraph({
                                border: { top: { style: BorderStyle.SINGLE, size: 1 } },
                                children: [
                                    new TextRun({ text: currentDocData.signatures[i+1].label + ": ", bold: true, size: 18 }),
                                    new TextRun({ text: currentDocData.signatures[i+1].name, size: 18 })
                                ]
                            })
                        ],
                        borders: { top: BorderStyle.NIL, bottom: BorderStyle.NIL, left: BorderStyle.NIL, right: BorderStyle.NIL }
                    }));
                } else {
                    cells.push(new TableCell({ children: [], borders: { top: BorderStyle.NIL, bottom: BorderStyle.NIL, left: BorderStyle.NIL, right: BorderStyle.NIL } }));
                }
                rows.push(new TableRow({ children: cells }));
            }

            children.push(new Table({ rows: rows, width: { size: 100, type: WidthType.PERCENTAGE } }));

            children.push(new Paragraph({
                children: [new TextRun({ text: currentDocData.footer, size: 16, color: "999999", italic: true })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 1500 }
            }));

            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                            borders: {
                                pageBorderLeft: { style: BorderStyle.SINGLE, size: 18, color: "222222" },
                                pageBorderRight: { style: BorderStyle.SINGLE, size: 18, color: "222222" },
                                pageBorderTop: { style: BorderStyle.SINGLE, size: 18, color: "222222" },
                                pageBorderBottom: { style: BorderStyle.SINGLE, size: 18, color: "222222" },
                            }
                        }
                    },
                    children: children
                }]
            });

            Packer.toBlob(doc).then(blob => {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${currentDocData.docType.replace(/\s+/g, '_')}_Official.docx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            });
        } catch (e) {
            alert(e.message);
        }
    });

    // PDF Download
    btnDownloadPdf.addEventListener('click', () => {
        if (!currentDocData) return;
        
        const temp = document.createElement('div');
        temp.style.width = '794px'; 
        temp.style.padding = '60px';
        temp.style.background = 'white';
        temp.style.color = '#111';
        temp.style.fontFamily = "'Times New Roman', serif";
        temp.style.position = 'fixed';
        temp.style.top = '0';
        temp.style.left = '0';
        temp.style.zIndex = '-1000'; // Under current UI but in view
        temp.style.boxSizing = 'border-box';
        temp.style.border = '4px double #333';
        
        let sectionsHtml = currentDocData.contentSections.map(s => `
            <div style="margin-bottom: 2rem;">
                <h4 style="margin-bottom: 0.5rem; border-bottom: 1px solid #ddd;">${s.title}</h4>
                <p style="text-align: justify; margin: 0; padding-left: 20px;">${s.text}</p>
            </div>
        `).join('');

        let sigHtml = currentDocData.signatures.map(s => `
            <div style="margin-top: 4rem; border-top: 1px solid #000; width: 42%; padding-top: 0.5rem; font-size: 0.95rem;">
                <strong>${s.label}:</strong> ${s.name}
            </div>
        `).join('');

        temp.innerHTML = `
            <div style="text-align: center; color: #666; font-size: 0.8rem; margin-bottom: 0.5rem;">${currentDocData.header}</div>
            <div style="text-align: right; color: #888; font-size: 0.75rem; margin-bottom: 1rem;">${currentDocData.reference}</div>
            <h1 style="text-align: center; margin-bottom: 3rem; font-size: 2.2rem; text-decoration: underline;">${currentDocData.title}</h1>
            ${sectionsHtml}
            <div style="display: flex; flex-wrap: wrap; justify-content: space-between; margin-top: 2rem;">${sigHtml}</div>
            <div style="margin-top: 5rem; text-align: center; font-size: 0.8rem; color: #777; border-top: 1px solid #eee; padding-top: 1rem;">
                <em>${currentDocData.footer}</em>
            </div>
        `;

        document.body.appendChild(temp);

        setTimeout(() => {
            const opt = {
                margin: 0,
                filename: `${currentDocData.docType.replace(/\s+/g, '_')}_Official.pdf`,
                image: { type: 'jpeg', quality: 1.0 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'px', format: [794, 1123], orientation: 'portrait' }
            };

            html2pdf().set(opt).from(temp).save().then(() => {
                document.body.removeChild(temp);
            }).catch(err => {
                alert("PDF Error: " + err.message);
                document.body.removeChild(temp);
            });
        }, 1000); // 1 second delay for rendering
    });
});
