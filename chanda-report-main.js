window.initializeScript = function(table) {
    console.log('Initializing script for table:', table);
    // First, find the header row (might be thead or first tr)
    const thead = table.querySelector('thead');
    const headerRow = thead ? thead.querySelector('tr') : table.rows[0];

    // Remove existing Download column if it exists
    const headerCells = headerRow.cells;
    if (headerCells.length > 0 && headerCells[headerCells.length - 1].textContent === 'Download') {
        headerRow.deleteCell(-1);
    }

    // Add header for new column
    const th = document.createElement('th');
    th.textContent = 'Download';
    headerRow.appendChild(th);

    // Add download span to ALL data rows
    const rows = thead ? 
                table.querySelectorAll('tbody tr') : 
                table.querySelectorAll('tr');

    console.log('Found rows:', rows.length);

    rows.forEach(row => {
        // Remove existing Download cell if it exists
        const cells = row.cells;
        if (cells.length > 0 && cells[cells.length - 1].textContent === 'Download') {
            row.deleteCell(-1);
        }

        const td = document.createElement('td');
        const span = document.createElement('span');
        span.textContent = 'Download';
        span.style.cursor = 'pointer';
        span.style.color = 'blue';
        span.style.textDecoration = 'underline';
        span.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            window.generateReport(row);
        });
        td.appendChild(span);
        row.appendChild(td);
    });
};

(function() {
    'use strict';

    // Get configuration from loader
    const nazimName = window.scriptConfig.nazimName;
    const styles = window.scriptConfig.styles;

    console.log('Config loaded:', { nazimName, styles });

    // Helper function for safe element access
    window.safelyAccessElement = function(selector) {
        try {
            return document.querySelector(selector);
        } catch (e) {
            console.warn('Could not access element:', selector);
            return null;
        }
    };

    // Helper function to convert German number format to standard decimal
    window.parseGermanNumber = function(value) {
        if (!value) return 0;
        let cleanValue = value.toString().replace(/\s/g, '').replace(/\./g, '');
        cleanValue = cleanValue.replace(',', '.');
        const number = parseFloat(cleanValue);
        return isNaN(number) ? 0 : number;
    };

    // Helper function to format number to German format
    window.formatGermanNumber = function(number) {
        return number.toFixed(2).replace('.', ',');
    };

    window.cleanText = function(text) {
        return text.toString()
            .trim()
            .replace(/\s+/g, '_')
            .replace(/[^\w\-\.]/g, '')
            .replace(/_+/g, '_');
    };

    window.generateReport = function(row) {
        //getting chanda year safely
        const yearSelect = window.safelyAccessElement('select[name="iYear"]');
        const chandaYear = yearSelect ? yearSelect.value : new Date().getFullYear();

        const cells = row.cells;

        // Convert values to numbers using German number format
        const budget = {
            majlis: window.parseGermanNumber(cells[5].textContent),
            ijtema: window.parseGermanNumber(cells[8].textContent),
            ishaat: window.parseGermanNumber(cells[11].textContent)
        };

        const paid = {
            majlis: window.parseGermanNumber(cells[6].textContent),
            ijtema: window.parseGermanNumber(cells[9].textContent),
            ishaat: window.parseGermanNumber(cells[12].textContent)
        };

        // Calculate totals
        const totalBudget = budget.majlis + budget.ijtema + budget.ishaat;
        const totalPaid = paid.majlis + paid.ijtema + paid.ishaat;
        const totalRest = totalBudget - totalPaid;

        // Create report container with specific styling
        const reportContainer = document.createElement('div');
        reportContainer.style.position = 'fixed';
        reportContainer.style.left = '-9999px';
        reportContainer.style.top = '0';
        reportContainer.style.background = 'white';
        reportContainer.style.width = '297mm'; // A6 landscape width
        reportContainer.style.height = '105mm'; // A6 landscape height

        // Generate report HTML with German number format
        const reportHTML = `
            <table style="margin:2% 3%; width: 94%; border-collapse: collapse; border: 1px solid black; background-color: ${styles.backgroundColor}; font-size:${styles.fontSize}; font-family:${styles.fontFamily}; color:${styles.color}; letter-spacing:${styles.letterSpacing};">
                <tr>
                    <td colspan="5" style="border: 1px solid black; padding: 3px; text-align:center; background-color:${styles.highlightColor};"><h2>Chanda Bericht Majlis Khuddam-ul-Ahmadiyya ${cells[2].textContent}</h2></td>
                </tr>
                <tr>
                    <td colspan="5" style="text-align:center;"><h4>Jahr ${chandaYear}</h4></td>
                </tr>
                <tr>
                    <td colspan="5" style="border: 1px solid black; padding: 3px;"><b>ID:</b> ${cells[3].textContent}</td>
                </tr>
                <tr>
                    <td colspan="5" style="border: 1px solid black; padding: 3px;"><b>Name:</b> ${cells[4].textContent}</td>
                </tr>
                <tr style="background-color:${styles.highlightColor};">
                    <td style="border: 1px solid black; padding: 3px;"></td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">Majlis</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">Ijtema</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">Ishaat</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">Total</td>
                </tr>
                <tr>
                    <td style="border: 1px solid black; padding: 3px; background-color:${styles.highlightColor}">Budget</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(budget.majlis)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(budget.ijtema)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(budget.ishaat)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(totalBudget)}</td>
                </tr>
                <tr>
                    <td style="border: 1px solid black; padding: 3px; background-color:${styles.highlightColor};">Paid</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(paid.majlis)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(paid.ijtema)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(paid.ishaat)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(totalPaid)}</td>
                </tr>
                <tr>
                    <td style="border: 1px solid black; padding: 3px; background-color:${styles.highlightColor};">Rest</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(budget.majlis - paid.majlis)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(budget.ijtema - paid.ijtema)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(budget.ishaat - paid.ishaat)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${window.formatGermanNumber(totalRest)}</td>
                </tr>
                <tr>
                    <td colspan="5" style="border: 1px solid black; padding: 3px;">Nazim Maal / Qaid Majlis: ${nazimName}</td>
                </tr>
                <tr>
                    <td colspan="5" style="border: 1px solid black; padding: 3px;">Link: <u><a href="https://www.software.khuddam.de/maalonline">https://www.software.khuddam.de/maalonline</a></u> </td>
                </tr>
            </table>
        `;

        reportContainer.innerHTML = reportHTML;
        document.body.appendChild(reportContainer);

        // Generate PDF using html2canvas and jsPDF with updated configuration
        window.html2canvas(reportContainer, {
            scale: 2,
            backgroundColor: '#ffffff',
            logging: false,
            allowTaint: true,
            useCORS: true,
            ignoreElements: (element) => {
                return element.tagName.toLowerCase() === 'iframe' ||
                       element.hasAttribute('crossorigin');
            }
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
            pdf.save(`report_${window.cleanText(cells[4].textContent)}.pdf`);

            // Clean up
            reportContainer.remove();
        }).catch(error => {
            console.error('Error generating PDF:', error);
            reportContainer.remove();
        });
    };

    // Start initialization immediately
    const table = document.getElementById('memberBudgetList');
    if (table) {
        console.log('Table found, initializing...');
        window.initializeScript(table);
    } else {
        console.log('Table not found, waiting...');
        const waitForTable = setInterval(() => {
            const table = document.getElementById('memberBudgetList');
            if (table) {
                clearInterval(waitForTable);
                console.log('Table found after waiting, initializing...');
                window.initializeScript(table);
            }
        }, 1000);
    }
})();

console.log('Main script loaded completely');
