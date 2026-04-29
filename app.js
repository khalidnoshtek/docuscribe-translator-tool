document.addEventListener('DOMContentLoaded', () => {
    const textArea = document.getElementById('text-input');
    const wordCount = document.getElementById('word-count');
    const charCount = document.getElementById('char-count');
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
            return new Paragraph({
                children: [new TextRun(line)],
                spacing: { after: 200 }
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
