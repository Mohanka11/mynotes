(function () {
    // 1. Dynamically load dependencies (html-docx-js and FileSaver.js)
    function loadScript(src) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    const loadDependencies = Promise.all([
        loadScript('https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.min.js'),
        loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
    ]);

    // 2. Inject CSS Styles for the Action Buttons
    const style = document.createElement('style');
    style.textContent = `
        .action-bar {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .action-btn {
            border: none;
            padding: 10px 18px;
            font-size: 14px;
            font-weight: bold;
            border-radius: 6px;
            cursor: pointer;
            transition: background 0.2s ease, transform 0.1s ease;
        }
        .text-btn {
            background: #4b5563;
            color: #ffffff;
        }
        .text-btn:hover {
            background: #374151;
        }
        .doc-btn {
            background: #2b579a;
            color: #ffffff;
        }
        .doc-btn:hover {
            background: #1e3e6d;
        }
        .action-btn:active {
            transform: scale(0.98);
        }
    `;
    document.head.appendChild(style);

    // 3. Inject download buttons automatically on DOM load
    document.addEventListener('DOMContentLoaded', function () {
        const container = document.querySelector('.container');
        if (!container) return;

        const actionBar = document.createElement('div');
        actionBar.className = 'action-bar';
        actionBar.innerHTML = `
            <button class="action-btn text-btn" id="btnExportTxt">📄 Save as Plain Text (.txt)</button>
            <button class="action-btn doc-btn" id="btnExportDocx">📝 Export to MS Word (.docx)</button>
        `;

        container.insertBefore(actionBar, container.firstChild);

        document.getElementById('btnExportTxt').addEventListener('click', downloadAsText);
        document.getElementById('btnExportDocx').addEventListener('click', () => {
            loadDependencies
                .then(exportToDocx)
                .catch(err => {
                    console.error('Failed to load docx export libraries:', err);
                    alert('Could not load docx converter libraries. Check your internet connection.');
                });
        });
    });

    // 4. Export to real .docx format
    function exportToDocx() {
        const title = document.querySelector('h1')?.innerText || 'Study_Notes';
        const subTitle = document.querySelector('header h3')?.innerText || '';

        // Clone container to clean up buttons and navigation links
        const containerClone = document.querySelector('.container').cloneNode(true);
        containerClone.querySelectorAll('.action-bar, .toc, .back').forEach(el => el.remove());

        // Prepare clean HTML with inline styling for docx conversion
        const docxHtml = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body { font-family: Arial, sans-serif; line-height: 1.5; font-size: 11pt; color: #111111; }
                    h1 { color: #0066cc; font-size: 18pt; text-align: center; margin-bottom: 4pt; }
                    h2 { color: #0066cc; font-size: 13pt; border-bottom: 1pt solid #0066cc; margin-top: 16pt; margin-bottom: 6pt; }
                    h3 { color: #333333; font-size: 11pt; margin-top: 10pt; }
                    table { width: 100%; border-collapse: collapse; margin-top: 8pt; margin-bottom: 8pt; }
                    table, th, td { border: 1px solid #999999; }
                    th { background-color: #0066cc; color: #ffffff; padding: 6px; font-weight: bold; }
                    td { padding: 6px; }
                    .note { background-color: #fff8d8; border-left: 4pt solid #ff9900; padding: 8px; margin-top: 10pt; }
                    .example-box { background-color: #f4f6f8; border-left: 3pt solid #0066cc; padding: 8px; margin-top: 8pt; font-family: Consolas, Courier, monospace; }
                    ul, ol { margin-left: 15pt; }
                </style>
            </head>
            <body>
                <h1>${title}</h1>
                <p style="text-align: center; font-weight: bold; color: #555555;">${subTitle}</p>
                <hr/>
                ${containerClone.innerHTML}
            </body>
            </html>
        `;

        // Generate .docx Blob using html-docx-js
        const convertedBlob = htmlDocx.asBlob(docxHtml, {
            orientation: 'portrait',
            margins: { top: 720, right: 720, bottom: 720, left: 720 } // 0.5 inch margins (twips)
        });

        const filename = (subTitle || title).replace(/[^a-zA-Z0-9_-]/g, '_') + '.docx';
        saveAs(convertedBlob, filename);
    }

    // 5. Plain Text (.txt) Download Handler
    function downloadAsText() {
        const title = document.querySelector('h1')?.innerText || 'Study Notes';
        const subTitle = document.querySelector('header h3')?.innerText || '';

        let textContent = `${title.toUpperCase()}\n${subTitle}\n${'='.repeat(50)}\n\n`;

        document.querySelectorAll('.container section').forEach(sec => {
            const secTitle = sec.querySelector('h2')?.innerText || '';
            textContent += `\n\n[ ${secTitle.toUpperCase()} ]\n${'-'.repeat(40)}\n`;

            sec.querySelectorAll('p, h3, li, .note, .example-box, table').forEach(el => {
                if (el.tagName === 'H3') {
                    textContent += `\n### ${el.innerText.trim()}\n`;
                } else if (el.tagName === 'LI') {
                    textContent += `  • ${el.innerText.trim()}\n`;
                } else if (el.tagName === 'TABLE') {
                    textContent += `\n--- Table Data ---\n`;
                    el.querySelectorAll('tr').forEach(tr => {
                        const rowData = Array.from(tr.querySelectorAll('th, td')).map(cell => cell.innerText.trim()).join(' | ');
                        textContent += `${rowData}\n`;
                    });
                    textContent += `------------------\n`;
                } else if (el.classList.contains('example-box') || el.classList.contains('note')) {
                    textContent += `\n>> ${el.innerText.trim()}\n`;
                } else if (el.tagName === 'P') {
                    textContent += `\n${el.innerText.trim()}\n`;
                }
            });
        });

        const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8' });
        const filename = (subTitle || title).replace(/[^a-zA-Z0-9_-]/g, '_') + '.txt';

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
    }
})();