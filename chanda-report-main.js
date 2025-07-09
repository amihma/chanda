// ==UserScript==
// @name         Chanda Report Downloader (Main)
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Main script for Chanda Report Downloader
// @author       Amir Ahmad
// ==/UserScript==

console.log('Main script starting...');

// Explicitly create a container for our functions
window.chandaReport = {};

// Initialize everything after config is loaded
(async function() {
    'use strict';

    try {
        // Wait for configuration
        const config = await new Promise((resolve) => {
            const checkConfig = setInterval(() => {
                if (window.scriptConfig) {
                    clearInterval(checkConfig);
                    resolve(window.scriptConfig);
                }
            }, 100);
        });

        const nazimName = config.nazimName;
        const styles = config.styles;

        console.log('Config loaded:', { nazimName, styles });

        // Explicitly attach all functions to window.chandaReport
        window.chandaReport.safelyAccessElement = function(selector) {
            try {
                return document.querySelector(selector);
            } catch (e) {
                console.warn('Could not access element:', selector);
                return null;
            }
        };

        window.chandaReport.parseGermanNumber = function(value) {
            if (!value) return 0;
            let cleanValue = value.toString().replace(/\s/g, '').replace(/\./g, '');
            cleanValue = cleanValue.replace(',', '.');
            const number = parseFloat(cleanValue);
            return isNaN(number) ? 0 : number;
        };

        window.chandaReport.formatGermanNumber = function(number) {
            return number.toFixed(2).replace('.', ',');
        };

        window.chandaReport.cleanText = function(text) {
            return text.toString()
                .trim()
                .replace(/\s+/g, '_')
                .replace(/[^\w\-\.]/g, '')
                .replace(/_+/g, '_');
        };

        // Rest of your functions...
        // [Previous code remains the same, just change function declarations to attach to window.chandaReport]

        // Initialize table
        const table = document.getElementById('memberBudgetList');
        if (table) {
            console.log('Table found, initializing...');
            window.chandaReport.initializeScript(table);
        } else {
            console.log('Table not found, waiting...');
            const waitForTable = setInterval(() => {
                const table = document.getElementById('memberBudgetList');
                if (table) {
                    clearInterval(waitForTable);
                    console.log('Table found after waiting, initializing...');
                    window.chandaReport.initializeScript(table);
                }
            }, 1000);
        }

        console.log('Main script initialization complete');
    } catch (error) {
        console.error('Error in main script:', error);
    }
})();

// Log availability of functions
setTimeout(() => {
    console.log('Checking chandaReport functions:');
    console.log('chandaReport object:', window.chandaReport);
    console.log('initializeScript:', typeof window.chandaReport?.initializeScript);
    console.log('generateReport:', typeof window.chandaReport?.generateReport);
}, 2000);

console.log('Main script loaded completely');
