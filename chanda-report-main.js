console.log('Main script starting...');

(function() {
    'use strict';

    // Get configuration from loader
    const nazimName = window.scriptConfig.nazimName;
    const styles = window.scriptConfig.styles;

    console.log('Config loaded:', { nazimName, styles });

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

    function initializeScript(table) {
        console.log('Initializing script for table:', table);
        // First, find the header row (might be thead or first tr)
        const thead = table.querySelector('thead');
        const headerRow = thead ? thead.querySelector('tr') : table.rows[0];

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

    function generateReport(row) {
        //getting chanda year
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
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(budget.majlis)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(budget.ijtema)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(budget.ishaat)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(totalBudget)}</td>
                </tr>
                <tr>
                    <td style="border: 1px solid black; padding: 3px; background-color:${styles.highlightColor};">Paid</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(paid.majlis)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(paid.ijtema)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(paid.ishaat)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(totalPaid)}</td>
                </tr>
                <tr>
                    <td style="border: 1px solid black; padding: 3px; background-color:${styles.highlightColor};">Rest</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(budget.majlis - paid.majlis)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(budget.ijtema - paid.ijtema)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(budget.ishaat - paid.ishaat)}</td>
                    <td style="border: 1px solid black; padding: 3px; text-align:center;">${formatGermanNumber(totalRest)}</td>
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

        // Generate PDF using html2canvas and jsPDF
        html2canvas(reportContainer, {
            scale: 2,
            backgroundColor: '#ffffff',
            logging: false
        }).then(canvas => {
            const imgData = canvas.toDataURL('image/jpeg', 1.0);
            const pdf = new window.jspdf.jsPDF({
    orientation: 'landscape',
    format: 'a6',
    unit: 'mm'
});
                orientation: 'landscape',
                format: 'a6',
                unit: 'mm'
            });

            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = pdf.internal.pageSize.getHeight();

            pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
            pdf.save(`report_${cleanText(cells[4].textContent)}.pdf`);

            // Clean up
            reportContainer.remove();
        }).catch(error => {
            console.error('Error generating PDF:', error);
            reportContainer.remove();
        });
    }

    // Start initialization immediately
    const table = document.getElementById('memberBudgetList');
    if (table) {
        console.log('Table found, initializing...');
        initializeScript(table);
    } else {
        console.log('Table not found, waiting...');
        const waitForTable = setInterval(() => {
            const table = document.getElementById('memberBudgetList');
            if (table) {
                clearInterval(waitForTable);
                console.log('Table found after waiting, initializing...');
                initializeScript(table);
            }
        }, 1000);
    }
})();

console.log('Main script loaded completely');
