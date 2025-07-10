console.log('Main script starting...');

(function() {
    'use strict';

    // Get configuration from loader
    const nazimName = window.scriptConfig.nazimName;
    const styles = window.scriptConfig.styles;

    // Helper function to convert German number format to standard decimal
    function parseGermanNumber(value) {
        if (!value) return 0;
        let cleanValue = value.toString().replace(/\s/g, '').replace(/\./g, '');
        cleanValue = cleanValue.replace(',', '.');
        const number = parseFloat(cleanValue);
        return isNaN(number) ? 0 : number;
    }

    // Helper function to format number to German format
    function formatGermanNumber(number) {
        return number.toFixed(2).replace('.', ',');
    }

    function cleanText(text) {
        return text.toString()
            .trim()
            .replace(/\s+/g, '_')
            .replace(/[^\w\-\.]/g, '')
            .replace(/_+/g, '_');
    }

    function generateReport(row) {
        const chandaYear = document.querySelector('select[name="iYear"]')?.value || new Date().getFullYear();
        const cells = row.cells;

        // Convert values to numbers using German number format
        const budget = {
            majlis: parseGermanNumber(cells[5].textContent),
            ijtema: parseGermanNumber(cells[8].textContent),
            ishaat: parseGermanNumber(cells[11].textContent)
        };

        const paid = {
            majlis: parseGermanNumber(cells[6].textContent),
            ijtema: parseGermanNumber(cells[9].textContent),
            ishaat: parseGermanNumber(cells[12].textContent)
        };

        // Calculate totals
        const totalBudget = budget.majlis + budget.ijtema + budget.ishaat;
        const totalPaid = paid.majlis + paid.ijtema + paid.ishaat;
        const totalRest = totalBudget - totalPaid;

        // Create report container
        const reportContainer = document.createElement('div');
        reportContainer.style.position = 'fixed';
        reportContainer.style.left = '-9999px';
        reportContainer.style.top = '0';
        reportContainer.style.background = 'white';
        reportContainer.style.width = '297mm';
        reportContainer.style.height = '105mm';

        // Generate report HTML
        reportContainer.innerHTML = `
            <table style="margin:2% 3%; width: 94%; border-collapse: collapse; border: 1px solid black; background-color: ${styles.backgroundColor}; font-size:${styles.fontSize}; font-family:${styles.fontFamily}; color:${styles.color}; letter-spacing:${styles.letterSpacing};">
                <!-- Your existing table HTML -->
                <tr>
                    <td colspan="5" style="border: 1px solid black; padding: 3px; text-align:center; background-color:${styles.highlightColor};"><h2>Chanda Bericht Majlis Khuddam-ul-Ahmadiyya ${cells[2].textContent}</h2></td>
                </tr>
                <!-- ... rest of your table HTML ... -->
            </table>
        `;

        document.body.appendChild(reportContainer);

        // Generate PDF
        html2canvas(reportContainer, {
            scale: 2,
            backgroundColor: '#ffffff',
            logging: false,
            allowTaint: true,
            useCORS: true
        }).then(canvas => {
            const imgData = canvas.toDataURL('image/jpeg', 1.0);
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF({
                orientation: 'landscape',
                format: 'a6',
                unit: 'mm'
            });

            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = pdf.internal.pageSize.getHeight();

            pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
            pdf.save(`report_${cleanText(cells[4].textContent)}.pdf`);

            reportContainer.remove();
        }).catch(error => {
            console.error('Error generating PDF:', error);
            reportContainer.remove();
        });
    }

    function initializeScript(table) {
        // Add Download column header
        const thead = table.querySelector('thead');
        const headerRow = thead ? thead.querySelector('tr') : table.rows[0];
        const th = document.createElement('th');
        th.textContent = 'Download';
        headerRow.appendChild(th);

        // Add download buttons to rows
        const rows = thead ? 
                    table.querySelectorAll('tbody tr') : 
                    table.querySelectorAll('tr');

        rows.forEach(row => {
            const td = document.createElement('td');
            const span = document.createElement('span');
            span.textContent = 'Download';
            span.style.cursor = 'pointer';
            span.style.color = 'blue';
            span.style.textDecoration = 'underline';
            span.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                generateReport(row);
            });
            td.appendChild(span);
            row.appendChild(td);
        });
    }

    // Initialize when table is available
    const table = document.getElementById('memberBudgetList');
    if (table) {
        initializeScript(table);
    } else {
        const waitForTable = setInterval(() => {
            const table = document.getElementById('memberBudgetList');
            if (table) {
                clearInterval(waitForTable);
                initializeScript(table);
            }
        }, 1000);
    }
})();

console.log('Main script loaded completely');
