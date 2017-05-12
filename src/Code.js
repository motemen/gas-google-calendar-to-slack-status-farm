"use strict";
// declare namespace GoogleAppsScript {
//   module Calendar {
//     interface CalendarApp {
//       GuestStatus: typeof GuestStatus;
//     }
//   }
// }
var OAuth2;
var SLACK_CLIENT_ID = '';
var SLACK_CLIENT_SECRET = '';
var SLACK_LOG_WEBHOOK_URL = '';
function doGet(e) {
    var slack = getSlackService();
    var template = HtmlService.createTemplate('<!DOCTYPE html>' +
        '<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/css/bootstrap.min.css" integrity="sha384-rwoIResjU2yc3z8GV/NPeZWAv56rSmLldC3R/AZzGRnGxQQKnKkoFVhFQhNUwEyJ" crossorigin="anonymous">' +
        '<div class="container" style="padding-top: 3rem; font-size: 1.25rem">' +
        '<h1 style="margin-bottom: 1rem">Google Calendar to Slack Status</h1>' +
        '<p>こんにちは <strong><?= userEmail ?></strong> さん</p>' +
        '<p><? if (slackUser) { ?> Slack連携済み (@<?= slackUser ?>)  <form action="<?= scriptUrl ?>" method="post"><input class="btn btn-danger" type="submit" name="revoke" value="Slack連携解除"><? } else { ?><a class="btn btn-primary" href="<?= authorizationUrl ?>" target="_top">Slack連携する</a><? } ?></p>');
    template.authorizationUrl = slack.getAuthorizationUrl();
    template.userEmail = Session.getActiveUser().getEmail();
    template.scriptUrl = getWebAppUrl();
    template.slackUser = null;
    if (slack.hasAccess()) {
        var auth = getSlackAuthInfo(slack);
        if ('user' in auth) {
            template.slackUser = auth['user'];
        }
    }
    return HtmlService.createHtmlOutput(template.evaluate());
}
function doPost(e) {
    if (e.parameters.revoke) {
        revoke();
        return HtmlService.createHtmlOutput("\u9023\u643A\u89E3\u9664\u3057\u307E\u3057\u305F <a target=\"_top\" href=\"" + getWebAppUrl() + "\">\u623B\u308B</a>");
    }
}
function revoke() {
    getSlackService().reset();
    ScriptApp.getProjectTriggers().forEach(function (trigger) { return ScriptApp.deleteTrigger(trigger); });
    log("revoke: " + Session.getActiveUser().getEmail());
}
function updateStatusTrigger() {
    var slack = getSlackService();
    var status = getCurrentStatusFromCalendarEvent();
    updateSlackStatus(slack, status);
}
function getSlackService() {
    return OAuth2.createService('slack')
        .setAuthorizationBaseUrl('https://slack.com/oauth/authorize')
        .setTokenUrl('https://slack.com/api/oauth.access')
        .setClientId(SLACK_CLIENT_ID)
        .setClientSecret(SLACK_CLIENT_SECRET)
        .setCallbackFunction('authCallback')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope('users.profile:write');
}
function authCallback(request) {
    var slack = getSlackService();
    var authorized = slack.handleCallback(request);
    log("authCallback: " + Session.getActiveUser().getEmail() + " authorized=" + authorized);
    if (authorized) {
        ScriptApp.newTrigger('updateStatusTrigger').timeBased().everyMinutes(1).create();
        return HtmlService.createHtmlOutput("\u9023\u643A\u5B8C\u4E86 <a target=\"_top\" href=\"" + getWebAppUrl() + "\">\u623B\u308B</a>");
    }
    else {
        return HtmlService.createHtmlOutput("\u8A8D\u8A3C\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F <a target=\"_top\" href=\"" + getWebAppUrl() + "\">\u623B\u308B</a>");
    }
}
function getWebAppUrl() {
    return ScriptApp.getService().getUrl();
}
function getSlackAuthInfo(slack) {
    var resp = UrlFetchApp.fetch("https://slack.com/api/auth.test?token=" + slack.getAccessToken());
    return JSON.parse(resp.getContentText());
}
function updateSlackStatus(slack, status) {
    var resp = UrlFetchApp.fetch("https://slack.com/api/users.profile.set?token=" + slack.getAccessToken() + "&profile=" + encodeURIComponent(JSON.stringify(status)));
    if (resp.getResponseCode() !== 200) {
        log("updateSlackStatus: " + Session.getActiveUser().getEmail() + " responseCode=" + resp.getResponseCode());
    }
}
function getCurrentStatusFromCalendarEvent() {
    var now = new Date();
    var events = CalendarApp.getEvents(now, new Date(now.getTime() + 5 * 60 * 1000));
    for (var _i = 0, events_1 = events; _i < events_1.length; _i++) {
        var event_1 = events_1[_i];
        if (event_1.getMyStatus() === CalendarApp.GuestStatus.NO) {
            continue;
        }
        if (event_1.isAllDayEvent()) {
            continue;
        }
        var title = event_1.getTitle();
        var emoji = ':spiral_calendar_pad:';
        if (/(:[^:]+:)/.test(title)) {
            emoji = RegExp.$1;
            title = title.replace(/\s*:[^:]+:\s*/, '');
        }
        else if (/休/.test(title)) {
            emoji = ':yasumi:';
        }
        else if (/早退|遅刻/.test(title)) {
            emoji = ':bus:';
        }
        var start = event_1.getStartTime();
        var end = event_1.getEndTime();
        title = pad0(start.getHours()) + ":" + pad0(start.getMinutes()) + "\u301C" + pad0(end.getHours()) + ":" + pad0(end.getMinutes()) + " " + title;
        return {
            status_text: title,
            status_emoji: emoji
        };
    }
    ;
    return {
        status_text: '',
        status_emoji: ''
    };
}
function pad0(s) {
    return ('0' + s).substr(-2);
}
function log(message) {
    if (SLACK_LOG_WEBHOOK_URL.length) {
        UrlFetchApp.fetch(SLACK_LOG_WEBHOOK_URL, {
            method: 'post',
            payload: {
                payload: JSON.stringify({
                    text: "" + message,
                    username: 'cal2slack',
                    icon_emoji: ':robot_face:'
                })
            }
        });
    }
}
