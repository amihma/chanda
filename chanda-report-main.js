(function() {
    'use strict';

    // Get configuration from loader
    const nazimName = window.scriptConfig.nazimName;
    const styles = window.scriptConfig.styles;

    // Add the new helper function here
    function safelyAccessElement(selector) {
        try {
            return document.querySelector(selector);
        } catch (e) {
            console.warn('Could not access element:', selector);
            return null;
        }
    }

    // Your existing helper functions
    function parseGermanNumber(value) {
        if (!value) return 0;
        let cleanValue = value.toString().replace(/\s/g, '').replace(/\./g, '');
        cleanValue = cleanValue.replace(',', '.');
        const number = parseFloat(cleanValue);
        return isNaN(number) ? 0 : number;
    }

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
        // getting chanda year safely
        const yearSelect = safelyAccessElement('select[name="iYear"]');
        const chandaYear = yearSelect ? yearSelect.value : new Date().getFullYear();

        const cells = row.cells;
        // ... rest of your generateReport function ...

        // Update the html2canvas part with the new configuration
        html2canvas(reportContainer, {
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
            pdf.save(`report_${cleanText(cells[4].textContent)}.pdf`);

            // Clean up
            reportContainer.remove();
        }).catch(error => {
            console.error('Error generating PDF:', error);
            reportContainer.remove();
        });
    }

    // Your existing initialization code
    function initializeScript(table) {
        // ... your existing initializeScript code ...
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
