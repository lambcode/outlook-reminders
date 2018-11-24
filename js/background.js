'use strict';

let timezoneId = '';
let owaUrl = '';
let finalReminder = true;


let minuteInMilli = 1000 * 60;
let dayInMilli = minuteInMilli * 60 * 24;
let registeredNotifications = [];

chrome.storage.sync.get({
    owaUrl: '',
    timezoneId: '',
    finalReminder: true
}, (items) => {
    owaUrl = items.owaUrl;
    timezoneId = items.timezoneId;
    finalReminder = items.finalReminder;

    startLoop();
});

function twoDigitFormat(num) {
    return ("0" + num).slice(-2)
}

function registerNotification(item, millisBeforeShow) {
    let time = new Date(Date.parse(item.StartDate));
    let notificationConfig = {
        debug: { millisBeforeShow: millisBeforeShow },
        subject: item.Subject,
        location: item.Location,
        time: base12Hours(time) + ":" + twoDigitFormat(time.getMinutes())
    };

    return setTimeout(() => {
        notify(notificationConfig)
    }, millisBeforeShow);
}

function reRegisterNotifications() {

    let monthFromToday = new Date(Date.now() + (dayInMilli * 7));
    let monthFromTodayString = monthFromToday.getFullYear()
        + "-" + twoDigitFormat(monthFromToday.getMonth() + 1)
        + "-" + twoDigitFormat(monthFromToday.getDate());

    let postData = {
        "__type": "GetRemindersJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
            "TimeZoneContext": {
                "__type": "TimeZoneContext:#Exchange",
                "TimeZoneDefinition": {"__type": "TimeZoneDefinitionType:#Exchange", "Id": timezoneId}
            }
        },
        "Body": {"__type": "GetRemindersRequest:#Exchange", "EndTime": monthFromTodayString + "T00:00:00", "MaxItems": 0}
    };

    return new Promise((resolve, reject) => {
        chrome.cookies.get({
            url: owaUrl + 'service.svc',
            name: 'X-OWA-CANARY'
        }, (cookie) => {
            if (cookie)
                resolve(cookie);
            reject();
        });
    }).then(cookie => {

        return fetch(owaUrl + 'service.svc', {
            credentials: 'include',
            headers: {
                'Content-Length': '0',
                'X-OWA-CANARY': cookie.value,
                'Action': 'GetReminders',
                'X-OWA-ActionName': 'GetRemindersAction',
                'X-OWA-UrlPostData': JSON.stringify(postData)
            }
        }).then(value => {
            return value.json()
        }).then(json => {
            let newRegisteredNotifications = [];
            json.Body.Reminders.forEach((item) => {
                try {
                    let timeTillEvent = Date.parse(item.StartDate) - Date.now();
                    let timeTillNotify = Date.parse(item.ReminderTime) - Date.now();

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
        });
    })
    .catch(() => {
        chrome.notifications.create({
            type: "basic",
            iconUrl: "icons/128.png",
            title: "Calendar Error",
            message: "Calendar Authentication Error. Are you logged in?"
        });

        return Promise.reject();
    });

}

function notify(config) {
    let items = [];
    if (config.time)
        items.push({title: 'Time', message: config.time});
    if (config.location)
        items.push({title: 'Location', message: config.location});

    chrome.notifications.create({
        type: "list",
        iconUrl: "icons/128.png",
        title: "Event - " + config.subject,
        message: config.subject,
        items: items
    });
}

let isLooping = false;
function loop() {
    isLooping = true;

    reRegisterNotifications().then(() => {
        setTimeout(loop, minuteInMilli * 5);
    }).catch(() => {
        isLooping = false;
    })
}
function startLoop() {
    if (!isLooping)
        loop();
}

chrome.runtime.onStartup.addListener(startLoop);
chrome.runtime.onInstalled.addListener(startLoop);
chrome.notifications.onClosed.addListener(startLoop);

function base12Hours(time) {
    let hours = time.getHours();
    if (hours > 12)
        return hours - 12;
    else if (hours === 0)
        return 12;
    else
        return hours;
}

