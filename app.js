document.addEventListener('DOMContentLoaded', () => {
    const textArea = document.getElementById('text-input');
    const wordCount = document.getElementById('word-count');
    const docStatus = document.getElementById('doc-status');
    const btnAiGenerate = document.getElementById('btn-ai-generate');
    const loader = document.getElementById('loader');

    // Update word count
    textArea.addEventListener('input', () => {
        const text = textArea.value;
        const words = text.trim() ? text.trim().split(/\s+/).length : 0;
        wordCount.textContent = `${words} word${words !== 1 ? 's' : ''}`;
    });

    // AI Generate & Download Official Word Document
    btnAiGenerate.addEventListener('click', async () => {
        const text = textArea.value;
        if (!text.trim()) return;
        
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) {
            alert('API Key missing. Run: localStorage.setItem("gemini_api_key", "YOUR_KEY") in console.');
            return;
        }

        loader.querySelector('p').textContent = 'AI is formatting and generating your official Word document...';
        loader.classList.remove('hidden');

        try {
            const prompt = `You are a master legal and official document drafter.
I will provide you with a translated text of an old document (Property Deed, Nikah Nama, or Agreement).
Your task:
1. Identify the document type and a professional official title.
2. Draft a beautifully structured, formal English version of this document.
3. Return ONLY a JSON object. If you add any conversational text, put the JSON inside a code block.
JSON structure:
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
                            // REMOVED responseMimeType to fix compatibility issues
                        })
                    });

                    if (response.ok) {
                        const data = await response.json();
                        let rawText = data.candidates[0].content.parts[0].text;
                        
                        // Extract JSON if it's wrapped in markdown code blocks
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
            
            docStatus.textContent = `Generating Word File...`;
            generateWordFile(jsonResponse);
            docStatus.textContent = `Ready`;
            
        } catch (error) {
            console.error('Generation Error:', error);
            alert('Error: ' + error.message);
        } finally {
            loader.classList.add('hidden');
        }
    });

    const generateWordFile = (data) => {
        const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType } = window.docx;
        
        const children = [];

        // Title
        children.push(new Paragraph({
            children: [new TextRun({ text: data.title, bold: true, size: 32 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));

        // Header
        children.push(new Paragraph({
            children: [new TextRun({ text: data.header, italic: true, size: 20 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 }
        }));

        // Sections
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

        // Signatures Table
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

        const sigTable = new Table({
            rows: rows,
            width: { size: 100, type: WidthType.PERCENTAGE },
        });

        children.push(sigTable);

        // Footer
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
});
