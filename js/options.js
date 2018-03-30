// Saves options to chrome.storage.sync.
function save_options() {
    var owaUrl = document.getElementById('owaUrl').value;
    var timezoneId = document.getElementById('timezoneId').value;
    var finalReminder = document.getElementById('finalReminder').checked;

    chrome.storage.sync.set({
        owaUrl: owaUrl,
        timezoneId: timezoneId,
        finalReminder: finalReminder
    }, function() {
        // Update status to let user know options were saved.
        var status = document.getElementById('status');
        status.textContent = 'Options saved.';
        setTimeout(function() {
            status.textContent = '';
        }, 750);

        chrome.extension.getBackgroundPage().window.location.reload()
    });
}

// Restores select box and checkbox state using the preferences
// stored in chrome.storage.
function restore_options() {
    // Use default value color = 'red' and likesColor = true.
    chrome.storage.sync.get({
        owaUrl: '',
        timezoneId: '',
        finalReminder: true
    }, function(items) {
        document.getElementById('owaUrl').value = items.owaUrl;
        document.getElementById('timezoneId').value = items.timezoneId;
        document.getElementById('finalReminder').checked = items.finalReminder;
    });
}
document.addEventListener('DOMContentLoaded', restore_options);
document.getElementById('save').addEventListener('click',
    save_options);