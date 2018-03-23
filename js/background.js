'use strict';

var calendarId = '';
var owaUrl = '';
var notifyBeforeEventMinutes = 15;
var finalReminder = true;


var notifyBeforeEventMilli = 1000 * 60 * notifyBeforeEventMinutes;
var registeredNotifications = [];

chrome.storage.sync.get({
    owaUrl: '',
    calendarId: '',
    reminderTime: 15,
    finalReminder: true
}, (items) => {
    owaUrl = items.owaUrl;
    calendarId = items.calendarId;
    notifyBeforeEventMinutes = items.reminderTime;
    finalReminder = items.finalReminder;
});

function registerNotification(item, millisBeforeShow) {
    var time = new Date(Date.parse(item.Start));
    var notificationConfig = {
        subject: item.Subject,
        location: item.Location.DisplayName,
        organizer: item.Organizer.Mailbox.Name,
        time: base12Hours(time) + ":" + time.getMinutes()
    };

    return setTimeout(() => notify(notificationConfig), millisBeforeShow);
}

function reRegisterNotifications() {
    var postData = {
        "__type": "GetCalendarViewJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
            "TimeZoneContext": {
                "__type": "TimeZoneContext:#Exchange",
                "TimeZoneDefinition": {"__type": "TimeZoneDefinitionType:#Exchange", "Id": "Mountain Standard Time"}
            }
        },
        "Body": {
            "__type": "GetCalendarViewRequest:#Exchange",
            "CalendarId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": {
                    "__type": "FolderId:#Exchange",
                    "Id": calendarId,
                    "ChangeKey": "AgAAAA=="
                }
            },
            "RangeStart": "2018-02-25T00:00:00.001",
            "RangeEnd": "2018-04-08T00:00:00.000"
        }
    };

    return fetch(owaUrl + 'service.svc', {
        credentials: 'include',
        headers: {
            'Content-Length': '0',
            'X-OWA-CANARY': '', //must exist
            'Action': 'GetCalendarView',
            'X-OWA-ActionName': 'GetCalendarViewAction_PrefetchMonth',
            'X-OWA-UrlPostData': JSON.stringify(postData)
        }
    }).then(value => {
        return value.json()
    }).then(json => {
        var newRegisteredNotifications = [];
        json.Body.Items.forEach((item) => {
            try {
                var timeTillEvent = Date.parse(item.Start) - Date.now();
                var timeTillNotify = timeTillEvent - notifyBeforeEventMilli;

                if (timeTillNotify > 0)
                    newRegisteredNotifications.push(registerNotification(item, timeTillNotify));

                if (finalReminder && timeTillEvent > 0)
                    newRegisteredNotifications.push(registerNotification(item, timeTillEvent));

            } catch (e) {
                console.error("Unable to add item", item, e)
            }

        });

        registeredNotifications.forEach((id) => clearTimeout(id));
        registeredNotifications = newRegisteredNotifications;
    }).catch(() => {
        chrome.notifications.create({
            type: "basic",
            iconUrl: "icons/128.png",
            title: "Calendar Error",
            message: "Calendar Authentication Error. Are you logged in?",
            requireInteraction: true
        });

        return Promise.reject();
    });
}

function notify(config) {
    var items = [];
    if (config.time)
        items.push({title: 'Time', message: config.time});
    if (config.location)
        items.push({title: 'Location', message: config.location});
    if (config.organizer)
        items.push({ title: 'Organizer', message: config.organizer});

    chrome.notifications.create({
        type: "list",
        iconUrl: "icons/128.png",
        title: "Event - " + config.subject,
        message: config.subject,
        items: items,
        requireInteraction: true
    });
}

var isLooping = false;
function loop() {
    isLooping = true;

    reRegisterNotifications().then(() => {
        setTimeout(loop, 300000);
    }).catch(() => {
        isLooping = false;
    })
}

chrome.runtime.onInstalled.addListener(() => {
    loop();
});

chrome.notifications.onClosed.addListener(() => {
    if (!isLooping)
        loop();
});

function base12Hours(time) {
    var hours = time.getHours();
    if (hours > 12) {
        return hours - 12;
    } else if (hours === 0) {
        return 12;
    }
}

