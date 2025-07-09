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

    // Wait for the table to be loaded
    const waitForTable = setInterval(() => {
        const table = document.getElementById('memberBudgetList');
        if (table) {
            clearInterval(waitForTable);
            initializeScript(table);
        }
    }, 1000);

    function initializeScript(table) {
        // [Your existing initializeScript function code]
    }

    function generateReport(row) {
        // [Your existing generateReport function code]
    }

    function cleanText(text) {
        return text.toString()
            .trim()
            .replace(/\s+/g, '_')
            .replace(/[^\w\-\.]/g, '')
            .replace(/_+/g, '_');
    }
})();
