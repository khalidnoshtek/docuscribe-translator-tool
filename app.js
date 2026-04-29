document.addEventListener('DOMContentLoaded', () => {
    const textArea = document.getElementById('text-input');
    const wordCount = document.getElementById('word-count');
    const charCount = document.getElementById('char-count');
    const btnAiFormat = document.getElementById('btn-ai-format');
    const btnClean = document.getElementById('btn-clean');
    const btnGrammar = document.getElementById('btn-grammar');
    const btnWord = document.getElementById('btn-word');
    const btnPdf = document.getElementById('btn-pdf');
    const suggestionsContainer = document.getElementById('suggestions-container');
    const loader = document.getElementById('loader');

    // Update stats
    const updateStats = () => {
        const text = textArea.value;
        const words = text.trim() ? text.trim().split(/\s+/).length : 0;
        wordCount.textContent = `${words} word${words !== 1 ? 's' : ''}`;
        charCount.textContent = `${text.length} character${text.length !== 1 ? 's' : ''}`;
    };

    textArea.addEventListener('input', updateStats);

    // AI Smart Format
    btnAiFormat.addEventListener('click', async () => {
        const text = textArea.value;
        if (!text.trim()) return;
        
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) {
            alert('Please enter and save a Gemini API Key in the settings panel first.');
            return;
        }

        loader.querySelector('p').textContent = 'AI is formatting your document...';
        loader.classList.remove('hidden');

        try {
            const prompt = `You are a professional document formatting assistant.
I will provide you with a roughly translated text from Arabic/Urdu into English.
Your task is to:
1. Fix any grammar, spelling, or punctuation errors.
2. Structure the text logically into proper paragraphs.
3. Apply a formal, official document tone.
4. If there are clear headings, format them in ALL CAPS.
5. Return ONLY the formatted text. Do not add any conversational responses, markdown formatting (like *, #, or \`\`\`), or extra comments. Just return plain text formatted with newlines.

Text to format:
${text}`;

            // DISCOVERY: Find which models are actually available for this API Key
            let availableModels = [];
            try {
                const listResp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
                if (listResp.ok) {
                    const listData = await listResp.json();
                    availableModels = listData.models
                        .filter(m => m.supportedGenerationMethods.includes('generateContent'))
                        .map(m => m.name.split('/').pop());
                    console.log('Discovered models:', availableModels);
                }
            } catch (e) {
                console.warn('Discovery failed:', e);
            }

            // Priority order for fallback if discovery fails or to sort discovery results
            const priorityModels = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-1.0-pro', 'gemini-pro'];
            
            // Final list: discovered models first, then fallbacks
            const modelsToTry = availableModels.length > 0 
                ? [...new Set([...availableModels.filter(m => priorityModels.includes(m)), ...availableModels, ...priorityModels])]
                : priorityModels;

            let resultText = '';
            let success = false;
            let lastError = '';

            for (const model of modelsToTry) {
                try {
                    console.log(`Trying model: ${model}`);
                    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            contents: [{ parts: [{ text: prompt }] }]
                        })
                    });

                    if (response.ok) {
                        const data = await response.json();
                        if (data.candidates && data.candidates[0].content) {
                            resultText = data.candidates[0].content.parts[0].text;
                            success = true;
                            break;
                        }
                    } else {
                        const err = await response.json();
                        lastError = err.error?.message || response.statusText;
                        console.warn(`Model ${model} failed:`, lastError);
                    }
                } catch (e) {
                    console.warn(`Fetch error for ${model}:`, e);
                    lastError = e.message;
                }
            }

            if (!success) throw new Error(lastError || 'All models failed to respond.');
            
            // Cleanup any stray markdown
            resultText = resultText.replace(/```[a-z]*\n?/g, '').replace(/```/g, '').trim();
            
            textArea.value = resultText;
            updateStats();
            
            const originalText = btnAiFormat.innerHTML;
            btnAiFormat.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"></path></svg> Formatted!';
            setTimeout(() => btnAiFormat.innerHTML = originalText, 2000);
            
        } catch (error) {
            console.error('AI Formatting Error:', error);
            alert('Failed to format document: ' + error.message);
        } finally {
            loader.classList.add('hidden');
        }
    });

    // Clean text
    btnClean.addEventListener('click', () => {
        let text = textArea.value;
        if (!text) return;

        // Clean extra spaces
        text = text.replace(/ {2,}/g, ' ');
        // Fix space before punctuation
        text = text.replace(/ ([.,?!;])/g, '$1');
        // Ensure space after punctuation (except in numbers/decimals)
        text = text.replace(/([.,?!;])(?=[a-zA-Z])/g, '$1 ');
        // Capitalize first letter of each sentence
        text = text.replace(/(^\s*|[.!?]\s+)([a-z])/g, (m, p1, p2) => p1 + p2.toUpperCase());
        // Remove multiple empty lines
        text = text.replace(/\n{3,}/g, '\n\n');

        textArea.value = text;
        updateStats();
        
        // Show subtle feedback
        const originalText = btnClean.innerHTML;
        btnClean.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"></path></svg> Cleaned!';
        setTimeout(() => {
            btnClean.innerHTML = originalText;
        }, 2000);
    });

    // Check grammar via LanguageTool API
    btnGrammar.addEventListener('click', async () => {
        const text = textArea.value;
        if (!text.trim()) {
            suggestionsContainer.innerHTML = '<div class="empty-state"><p>Please enter some text to check.</p></div>';
            return;
        }

        loader.classList.remove('hidden');
        suggestionsContainer.innerHTML = '';

        try {
            const response = await fetch('https://api.languagetool.org/v2/check', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: `text=${encodeURIComponent(text)}&language=en-US`
            });

            if (!response.ok) throw new Error('API request failed');

            const data = await response.json();
            
            if (data.matches.length === 0) {
                suggestionsContainer.innerHTML = '<div class="empty-state"><p>No issues found! Your text looks great.</p></div>';
            } else {
                renderSuggestions(data.matches);
            }
        } catch (error) {
            console.error('Error checking grammar:', error);
            suggestionsContainer.innerHTML = '<div class="empty-state"><p style="color: var(--danger)">Failed to check grammar. Please try again later.</p></div>';
        } finally {
            loader.classList.add('hidden');
        }
    });

    const renderSuggestions = (matches) => {
        matches.forEach((match, index) => {
            const item = document.createElement('div');
            item.className = 'suggestion-item';
            
            // Extract context and highlight error
            const ctx = match.context;
            const before = ctx.text.substring(0, ctx.offset);
            const errorText = ctx.text.substring(ctx.offset, ctx.offset + ctx.length);
            const after = ctx.text.substring(ctx.offset + ctx.length);
            
            let replacementsHtml = '';
            match.replacements.slice(0, 5).forEach(rep => {
                replacementsHtml += `<button class="replacement-btn" data-offset="${match.offset}" data-length="${match.length}" data-value="${rep.value.replace(/"/g, '&quot;')}">${rep.value}</button>`;
            });

            item.innerHTML = `
                <div class="suggestion-message"><strong>${match.rule.issueType || 'Issue'}:</strong> ${match.message}</div>
                <div class="suggestion-context">
                    ${before}<span class="error">${errorText}</span>${after}
                </div>
                <div class="replacements">
                    ${replacementsHtml}
                    <button class="ignore-btn">Ignore</button>
                </div>
            `;

            suggestionsContainer.appendChild(item);

            // Add event listeners for replacements
            const repBtns = item.querySelectorAll('.replacement-btn');
            repBtns.forEach(btn => {
                btn.addEventListener('click', function() {
                    applyReplacement(parseInt(this.dataset.offset), parseInt(this.dataset.length), this.dataset.value);
                    item.remove();
                    if (suggestionsContainer.children.length === 0) {
                        suggestionsContainer.innerHTML = '<div class="empty-state"><p>All issues resolved!</p></div>';
                    }
                });
            });

            // Ignore button
            item.querySelector('.ignore-btn').addEventListener('click', () => {
                item.remove();
                if (suggestionsContainer.children.length === 0) {
                    suggestionsContainer.innerHTML = '<div class="empty-state"><p>All issues resolved!</p></div>';
                }
            });
        });
    };

    const applyReplacement = (offset, length, value) => {
        const text = textArea.value;
        // Apply replacement directly
        textArea.value = text.substring(0, offset) + value + text.substring(offset + length);
        updateStats();
        
        // Clear remaining suggestions because offsets are now broken
        suggestionsContainer.innerHTML = '<div class="empty-state"><p>Text modified. Please <a href="#" id="recheck-link" style="color:var(--primary); text-decoration: underline;">re-check</a> for further issues.</p></div>';
        document.getElementById('recheck-link').addEventListener('click', (e) => {
            e.preventDefault();
            btnGrammar.click();
        });
    };

    // Export Word
    btnWord.addEventListener('click', () => {
        const text = textArea.value;
        if (!text) return;

        const { Document, Packer, Paragraph, TextRun } = docx;
        
        const paragraphs = text.split('\n').map(line => {
            const trimmedLine = line.trim();
            const isHeading = trimmedLine.length > 0 && trimmedLine === trimmedLine.toUpperCase() && trimmedLine.length < 100;
            
            return new Paragraph({
                children: [
                    new TextRun({ 
                        text: line || " ",
                        bold: isHeading,
                        size: isHeading ? 28 : 24 // 14pt and 12pt (half-points in docx)
                    })
                ],
                spacing: { after: isHeading ? 100 : 200, before: isHeading ? 200 : 0 }
            });
        });

        const doc = new Document({
            sections: [{
                properties: {},
                children: paragraphs
            }]
        });

        Packer.toBlob(doc).then(blob => {
            window.saveAs(blob, "Translated_Document.docx");
        });
    });

    // Export PDF
    btnPdf.addEventListener('click', () => {
        const text = textArea.value;
        if (!text) return;

        const element = document.createElement('div');
        element.innerHTML = text.split('\n').map(line => `<p style="margin-bottom: 12px; line-height: 1.6;">${line}</p>`).join('');
        element.style.padding = '40px';
        element.style.fontFamily = 'Georgia, serif';
        element.style.fontSize = '12pt';
        element.style.color = 'black';
        element.style.background = 'white';

        const opt = {
            margin:       1,
            filename:     'Translated_Document.pdf',
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2 },
            jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
        };

        html2pdf().set(opt).from(element).save();
    });
});
