(function() {
    'use strict';

    // Make initializeScript global
    window.initializeScript = function() {
        console.log('Initializing Chanda Report Downloader...');
        
        // Check if jsPDF is loaded
        if (!window.jspdf) {
            console.error('jsPDF library not loaded!');
            return;
        }

        addDownloadButtons();
        console.log('Script initialized successfully!');
    };

    // Helper Functions
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

    async function generateReport(row) {
        try {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();

            // Get data from row
            const cells = row.cells;
            const name = cells[0].textContent.trim();
            const majlis = cells[1].textContent.trim();
            const ijtema = parseGermanNumber(cells[2].textContent);
            const chandaAam = parseGermanNumber(cells[3].textContent);
            const chandaWasiyyat = parseGermanNumber(cells[4].textContent);
            const total = ijtema + chandaAam + chandaWasiyyat;

            // Set up PDF
            doc.setFont('helvetica');
            doc.setFontSize(12);

            // Add content
            doc.text('Chanda Receipt', 105, 20, { align: 'center' });
            doc.text(`Name: ${name}`, 20, 40);
            doc.text(`Majlis: ${majlis}`, 20, 50);
            doc.text(`Date: ${new Date().toLocaleDateString()}`, 20, 60);
            
            // Add table
            doc.text('Chanda Type', 20, 80);
            doc.text('Amount (€)', 100, 80);
            
            doc.line(20, 85, 190, 85);
            
            let y = 95;
            if (ijtema > 0) {
                doc.text('Ijtema', 20, y);
                doc.text(formatGermanNumber(ijtema), 100, y);
                y += 10;
            }
            if (chandaAam > 0) {
                doc.text('Chanda Aam', 20, y);
                doc.text(formatGermanNumber(chandaAam), 100, y);
                y += 10;
            }
            if (chandaWasiyyat > 0) {
                doc.text('Chanda Wasiyyat', 20, y);
                doc.text(formatGermanNumber(chandaWasiyyat), 100, y);
                y += 10;
            }

            doc.line(20, y, 190, y);
            y += 10;
            
            doc.text('Total:', 20, y);
            doc.text(formatGermanNumber(total), 100, y);

            // Add signature section
            y += 30;
            doc.line(20, y, 80, y);
            doc.text('Nazim Mal', 20, y + 5);

            doc.line(120, y, 180, y);
            doc.text('Member', 120, y + 5);

            // Save PDF
            const filename = `chanda_receipt_${cleanText(name)}_${cleanText(new Date().toISOString())}.pdf`;
            doc.save(filename);

        } catch (error) {
            console.error('Error generating PDF:', error);
            alert('Error generating PDF. Please check console for details.');
        }
    }

    function addDownloadButtons() {
        const table = document.getElementById('memberBudgetList');
        if (!table) return;

        // Add button to table header
        const headerRow = table.querySelector('thead tr');
        const headerCell = document.createElement('th');
        headerCell.textContent = 'Download';
        headerRow.appendChild(headerCell);

        // Add buttons to each row
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const cell = document.createElement('td');
            const button = document.createElement('button');
            button.textContent = 'Download Receipt';
            button.style.padding = '5px 10px';
            button.style.backgroundColor = '#4CAF50';
            button.style.color = 'white';
            button.style.border = 'none';
            button.style.borderRadius = '3px';
            button.style.cursor = 'pointer';
            
            button.addEventListener('click', () => generateReport(row));
            
            cell.appendChild(button);
            row.appendChild(cell);
        });
    }

})();
