// Saves options to chrome.storage.sync.
function save_options() {
    var owaUrl = document.getElementById('owaUrl').value;
    var calendarId = document.getElementById('calendarId').value;
    var reminderTime = parseInt(document.getElementById('reminderTime').value);
    var finalReminder = document.getElementById('finalReminder').checked;

    chrome.storage.sync.set({
        owaUrl: owaUrl,
        calendarId: calendarId,
        reminderTime: reminderTime,
        finalReminder: finalReminder
    }, function() {
        // Update status to let user know options were saved.
        var status = document.getElementById('status');
        status.textContent = 'Options saved.';
        setTimeout(function() {
            status.textContent = '';
        }, 750);
    });
}

// Restores select box and checkbox state using the preferences
// stored in chrome.storage.
function restore_options() {
    // Use default value color = 'red' and likesColor = true.
    chrome.storage.sync.get({
        owaUrl: '',
        calendarId: '',
        reminderTime: 15,
        finalReminder: true
    }, function(items) {
        document.getElementById('owaUrl').value = items.owaUrl;
        document.getElementById('calendarId').value = items.calendarId;
        document.getElementById('reminderTime').value = items.reminderTime;
        document.getElementById('finalReminder').checked = items.finalReminder;
    });
}
document.addEventListener('DOMContentLoaded', restore_options);
document.getElementById('save').addEventListener('click',
    save_options);