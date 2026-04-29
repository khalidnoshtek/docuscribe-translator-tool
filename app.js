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

        loader.querySelector('p').textContent = 'AI is drafting your professional document...';
        loader.classList.remove('hidden');
        downloadSection.classList.add('hidden');

        try {
            const prompt = `You are an elite legal document architect.
I will provide you with a rough translation of an old document (e.g., Nikah Nama, Property Deed, Legal Agreement).
Your task:
1. Reconstruct it into a prestigious, official document with a sophisticated layout.
2. Use high-level legal English.
3. Structure it into logical sections with clear headings.
4. Return ONLY a JSON object.

JSON Structure:
{
  "docType": "Short Category",
  "title": "PRESTIGIOUS OFFICIAL TITLE",
  "reference": "Ref: DOC-${Math.floor(Math.random()*9000)+1000}",
  "header": "OFFICIAL CERTIFIED TRANSLATION",
  "contentSections": [
    { "title": "I. PREAMBLE / PARTIES", "text": "..." },
    { "title": "II. SUBJECT MATTER", "text": "..." },
    { "title": "III. TERMS AND CONDITIONS", "text": "..." }
  ],
  "signatures": [
    {"label": "Executing Party", "name": "..."},
    {"label": "Recipient/Second Party", "name": "..."},
    {"label": "Witness 1", "name": "..."},
    {"label": "Witness 2", "name": "..."}
  ],
  "footer": "This document is a certified translation of the original record."
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
        try {
            generateWordFile(currentDocData);
        } catch (e) {
            console.error("Word Export Error:", e);
            alert("Word export failed. Error: " + e.message);
        }
    });

    // PDF Download
    btnDownloadPdf.addEventListener('click', () => {
        if (!currentDocData) return;
        generatePdfFile(currentDocData);
    });

    const generateWordFile = (data) => {
        // Use window.docx explicitly to avoid scope issues
        const lib = window.docx;
        if (!lib) throw new Error("docx library not loaded correctly.");

        const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType, Border } = lib;
        const children = [];

        // Header Label
        children.push(new Paragraph({
            children: [new TextRun({ text: data.header, bold: true, size: 20, color: "555555" })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 }
        }));

        // Reference
        children.push(new Paragraph({
            children: [new TextRun({ text: data.reference, size: 18, color: "888888" })],
            alignment: AlignmentType.RIGHT,
            spacing: { after: 300 }
        }));

        // Main Title
        children.push(new Paragraph({
            children: [new TextRun({ text: data.title, bold: true, size: 36, underline: { type: "double" } })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 600 }
        }));

        // Content Sections
        data.contentSections.forEach(s => {
            children.push(new Paragraph({
                children: [new TextRun({ text: s.title, bold: true, size: 24 })],
                spacing: { before: 300, after: 150 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }
            }));
            children.push(new Paragraph({
                children: [new TextRun({ text: s.text, size: 24 })],
                alignment: AlignmentType.JUSTIFIED,
                spacing: { after: 200 },
                indent: { left: 400 }
            }));
        });

        // Signatures Grid
        const rows = [];
        for (let i = 0; i < data.signatures.length; i += 2) {
            const cells = [
                new TableCell({
                    children: [
                        new Paragraph({ spacing: { before: 1200 } }),
                        new Paragraph({
                            border: { top: { style: BorderStyle.SINGLE, size: 1, color: "000000" } },
                            children: [
                                new TextRun({ text: data.signatures[i].label + ": ", bold: true, size: 18 }),
                                new TextRun({ text: data.signatures[i].name, size: 18 })
                            ],
                            spacing: { before: 100 }
                        })
                    ],
                    borders: { top: BorderStyle.NIL, bottom: BorderStyle.NIL, left: BorderStyle.NIL, right: BorderStyle.NIL }
                })
            ];

            if (data.signatures[i + 1]) {
                cells.push(new TableCell({
                    children: [
                        new Paragraph({ spacing: { before: 1200 } }),
                        new Paragraph({
                            border: { top: { style: BorderStyle.SINGLE, size: 1, color: "000000" } },
                            children: [
                                new TextRun({ text: data.signatures[i+1].label + ": ", bold: true, size: 18 }),
                                new TextRun({ text: data.signatures[i+1].name, size: 18 })
                            ],
                            spacing: { before: 100 }
                        })
                    ],
                    borders: { top: BorderStyle.NIL, bottom: BorderStyle.NIL, left: BorderStyle.NIL, right: BorderStyle.NIL }
                }));
            } else {
                cells.push(new TableCell({ children: [], borders: { top: BorderStyle.NIL, bottom: BorderStyle.NIL, left: BorderStyle.NIL, right: BorderStyle.NIL } }));
            }
            rows.push(new TableRow({ children: cells }));
        }

        children.push(new Table({
            rows: rows,
            width: { size: 100, type: WidthType.PERCENTAGE },
            margins: { top: 400 }
        }));

        // Footer
        children.push(new Paragraph({
            children: [new TextRun({ text: data.footer, size: 16, color: "999999", italic: true })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 1500 }
        }));

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                        borders: {
                            pageBorderLeft: { style: BorderStyle.SINGLE, size: 18, space: 24, color: "222222" },
                            pageBorderRight: { style: BorderStyle.SINGLE, size: 18, space: 24, color: "222222" },
                            pageBorderTop: { style: BorderStyle.SINGLE, size: 18, space: 24, color: "222222" },
                            pageBorderBottom: { style: BorderStyle.SINGLE, size: 18, space: 24, color: "222222" },
                        }
                    }
                },
                children: children
            }]
        });

        Packer.toBlob(doc).then(blob => {
            if (window.saveAs) {
                window.saveAs(blob, `${data.docType.replace(/\s+/g, '_')}_Official.docx`);
            } else {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${data.docType.replace(/\s+/g, '_')}_Official.docx`;
                a.click();
            }
        });
    };

    const generatePdfFile = (data) => {
        // Create in-memory element for PDF generation with fixed A4 dimensions to prevent cropping
        const temp = document.createElement('div');
        temp.style.width = '794px'; // A4 width at 96 DPI
        temp.style.padding = '60px';
        temp.style.background = 'white';
        temp.style.color = '#111';
        temp.style.fontFamily = "'Times New Roman', serif";
        temp.style.lineHeight = '1.6';
        temp.style.border = '4px double #333';
        temp.style.boxSizing = 'border-box';
        temp.style.position = 'absolute';
        temp.style.left = '-10000px'; // Hide off-screen
        document.body.appendChild(temp);
        
        let sectionsHtml = data.contentSections.map(s => `
            <div style="margin-bottom: 2rem;">
                <h4 style="margin-bottom: 0.5rem; font-size: 1.2rem; border-bottom: 1px solid #ddd;">${s.title}</h4>
                <p style="text-align: justify; margin: 0; padding-left: 20px;">${s.text}</p>
            </div>
        `).join('');

        let sigHtml = data.signatures.map(s => `
            <div style="margin-top: 4rem; border-top: 1px solid #000; width: 42%; padding-top: 0.5rem; font-size: 0.95rem;">
                <strong>${s.label}:</strong> ${s.name}
            </div>
        `).join('');

        temp.innerHTML = `
            <div style="text-align: center; color: #666; font-size: 0.8rem; font-weight: bold; margin-bottom: 0.5rem;">${data.header}</div>
            <div style="text-align: right; color: #888; font-size: 0.75rem; margin-bottom: 1rem;">${data.reference}</div>
            <h1 style="text-align: center; margin-bottom: 3rem; font-size: 2rem; text-decoration: underline;">${data.title}</h1>
            ${sectionsHtml}
            <div style="display: flex; flex-wrap: wrap; justify-content: space-between; margin-top: 2rem;">
                ${sigHtml}
            </div>
            <div style="margin-top: 5rem; text-align: center; font-size: 0.8rem; color: #777; border-top: 1px solid #eee; padding-top: 1rem;">
                <em>${data.footer}</em>
            </div>
        `;

        const opt = {
            margin:       0,
            filename:     `${data.docType.replace(/\s+/g, '_')}_Official.pdf`,
            image:        { type: 'jpeg', quality: 1.0 },
            html2canvas:  { scale: 2, useCORS: true, letterRendering: true },
            jsPDF:        { unit: 'px', format: [794, 1123], orientation: 'portrait' }
        };

        html2pdf().set(opt).from(temp).toPdf().get('pdf').then(pdf => {
            document.body.removeChild(temp);
        }).save();
    };
});
