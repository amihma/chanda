(function() {
    'use strict';

    const SLACK_WEBHOOK_URL = 'https://hooks.slack.com/triggers/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';

    // Storage helper functions
    function getStoredValue(key, defaultValue = '') {
        return localStorage.getItem(key) || defaultValue;
    }

    function setStoredValue(key, value) {
        localStorage.setItem(key, value);
    }

    // Initialize login from localStorage
    let login = getStoredValue('login');

    // Function to extract login from email
    function extractLogin() {
        const userInfoDiv = document.querySelector('.x-toolbar-text.dbtext');
        if (userInfoDiv) {
            const emailMatch = userInfoDiv.textContent.match(/\(([^@]+)@/);
            if (emailMatch && emailMatch[1]) {
                login = emailMatch[1].toUpperCase();
                setStoredValue('login', login);
            }
        }
    }

    // Function to create and inject the button
    function createSlackButton() {
        const button = document.createElement('a');
        button.className = 'x-btn x-unselectable x-btn-default-small';
        button.style.margin = '0px 0px 0px 20px';
        button.innerHTML = 'FWO->Slack';
        button.style.backgroundColor = '#4A154B';
        button.style.color = 'white';
        button.style.padding = '5px 10px';
        button.style.borderRadius = '4px';
        button.style.cursor = 'pointer';
        button.style.textDecoration = 'none';
        button.style.display = 'inline-block';

        button.addEventListener('click', sendToSlack);
        return button;
    }

    // Function to send data to Slack
    function sendToSlack() {
        const wo_description = document.querySelector('input[name="description"]')?.value || '';
        const wo_number = document.querySelector('input[name="workordernum"]')?.value || '';
        const equipment = document.querySelector('input[name="equipment"]')?.value || '';
        const equipmentDesc = document.querySelector('input[name="equipmentdesc"]')?.value || '';
        const wo_equipment = equipment && equipmentDesc ? `${equipment} (${equipmentDesc})` : '';
        const createdBy = login;

        const payload = {
            wo_number: wo_number,
            wo_description: wo_description,
            wo_equipment: wo_equipment,
            createdBy: createdBy
        };

        GM_xmlhttpRequest({
            method: 'POST',
            url: SLACK_WEBHOOK_URL,
            headers: {
                'Content-Type': 'application/json'
            },
            data: JSON.stringify(payload),
            onload: function(response) {
                if (response.status === 200) {
                    alert('Message has been sent to Slack!');
                } else {
                    alert('Error: Failed to send message to Slack');
                }
            },
            onerror: function(error) {
                alert('Error: Failed to send message to Slack');
            }
        });
    }

    // Function to check and inject button
// Function to check and inject button
function checkAndInjectButton() {
    const descriptionInput = document.querySelector('input[name="description"]');
    if (!descriptionInput) return;

    const descriptionValue = descriptionInput.value;
    // if (!descriptionValue.startsWith('FWO')) return;
    if (!descriptionValue.startsWith('')) return;

    // Find the row instead of td
    const row = descriptionInput.closest('tr');
    if (!row || row.querySelector('.slack-button')) return;

    // Create new td element
    const newTd = document.createElement('td');
    newTd.setAttribute('role', 'presentation');
    newTd.style.verticalAlign = 'top';
    newTd.className = 'x-table-layout-cell';

    // Create the button
    const button = createSlackButton();
    button.classList.add('slack-button');

    // Add button to new td
    newTd.appendChild(button);

    // Add new td to the row
    row.appendChild(newTd);
}

    // Initialize function
    function initialize() {
        extractLogin();
        checkAndInjectButton();
    }

    // Create and start observer
    function startObserving() {
        const observer = new MutationObserver((mutations) => {
            checkAndInjectButton();
            extractLogin();
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }

    // Start the script
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            initialize();
            startObserving();
        });
    } else {
        initialize();
        startObserving();
    }
})();
