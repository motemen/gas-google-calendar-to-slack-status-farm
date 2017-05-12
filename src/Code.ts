// declare namespace GoogleAppsScript {
//   module Calendar {
//     interface CalendarApp {
//       GuestStatus: typeof GuestStatus;
//     }
//   }
// }

let OAuth2: any;

let SLACK_CLIENT_ID = '';
let SLACK_CLIENT_SECRET = '';
let SLACK_LOG_WEBHOOK_URL = '';

function doGet(e: any) {
  let slack = getSlackService();
  let template = HtmlService.createTemplate(
    '<!DOCTYPE html>' +
    '<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/css/bootstrap.min.css" integrity="sha384-rwoIResjU2yc3z8GV/NPeZWAv56rSmLldC3R/AZzGRnGxQQKnKkoFVhFQhNUwEyJ" crossorigin="anonymous">' +
    '<div class="container" style="padding-top: 3rem; font-size: 1.25rem">' +
    '<h1 style="margin-bottom: 1rem">Google Calendar to Slack Status</h1>' +
    '<p>こんにちは <strong><?= userEmail ?></strong> さん</p>' +
    '<p><? if (slackUser) { ?> Slack連携済み (@<?= slackUser ?>)  <form action="<?= scriptUrl ?>" method="post"><input class="btn btn-danger" type="submit" name="revoke" value="Slack連携解除"><? } else { ?><a class="btn btn-primary" href="<?= authorizationUrl ?>" target="_top">Slack連携する</a><? } ?></p>'
  );
  (template as any).authorizationUrl = slack.getAuthorizationUrl();
  (template as any).userEmail = Session.getActiveUser().getEmail();
  (template as any).scriptUrl = getWebAppUrl();
  (template as any).slackUser = null;

  if (slack.hasAccess()) {
    let auth = getSlackAuthInfo(slack);
    if ('user' in auth) {
      (template as any).slackUser = auth['user'];
    }
  }

  return HtmlService.createHtmlOutput(template.evaluate());
}

function doPost(e: any) {
  if (e.parameters.revoke) {
    revoke();
    return HtmlService.createHtmlOutput(`連携解除しました <a target="_top" href="${getWebAppUrl()}">戻る</a>`);
  }
}

function revoke() {
  getSlackService().reset();
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));

  log(`revoke: ${Session.getActiveUser().getEmail()}`);
}

function updateStatusTrigger() {
  let slack = getSlackService();
  let status = getCurrentStatusFromCalendarEvent();
  updateSlackStatus(slack, status)
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

function authCallback(request: any) {
  let slack = getSlackService();
  let authorized = slack.handleCallback(request);

  log(`authCallback: ${Session.getActiveUser().getEmail()} authorized=${authorized}`)

  if (authorized) {
    ScriptApp.newTrigger('updateStatusTrigger').timeBased().everyMinutes(1).create();
    return HtmlService.createHtmlOutput(`連携完了 <a target="_top" href="${getWebAppUrl()}">戻る</a>`);
  } else {
    return HtmlService.createHtmlOutput(`認証できませんでした <a target="_top" href="${getWebAppUrl()}">戻る</a>`);
  }
}

function getWebAppUrl(): string {
  return (ScriptApp.getService() as any).getUrl();
}

function getSlackAuthInfo(slack: any) {
  let resp = UrlFetchApp.fetch(`https://slack.com/api/auth.test?token=${slack.getAccessToken()}`);
  return JSON.parse(resp.getContentText());
}

function updateSlackStatus(slack: any, status: SlackProfile) {
  let resp = UrlFetchApp.fetch(`https://slack.com/api/users.profile.set?token=${slack.getAccessToken()}&profile=${encodeURIComponent(JSON.stringify(status))}`);
  if (resp.getResponseCode() !== 200) {
    log(`updateSlackStatus: ${Session.getActiveUser().getEmail()} responseCode=${resp.getResponseCode()}`)
  }
}

interface SlackProfile {
  status_text: string;
  status_emoji: string;
}

function getCurrentStatusFromCalendarEvent(): SlackProfile {
  let now = new Date();
  let events = CalendarApp.getEvents(now, new Date(now.getTime() + 5 * 60 * 1000));
  for (let event of events) {
    if (event.getMyStatus() === (CalendarApp.GuestStatus as any).NO) {
      continue;
    }

    if (event.isAllDayEvent()) {
      continue;
    }

    let title = event.getTitle();
    let emoji = ':spiral_calendar_pad:';
    if (/(:[^:]+:)/.test(title)) {
      emoji = RegExp.$1;
      title = title.replace(/\s*:[^:]+:\s*/, '');
    } else if (/休/.test(title)) {
      emoji = ':yasumi:';
    } else if (/早退|遅刻/.test(title)) {
      emoji = ':bus:';
    }

    let start = event.getStartTime();
    let end   = event.getEndTime();

    title = `${pad0(start.getHours())}:${pad0(start.getMinutes())}〜${pad0(end.getHours())}:${pad0(end.getMinutes())} ${title}`;

    return {
      status_text: title,
      status_emoji: emoji
    };
  };

  return {
    status_text: '',
    status_emoji: ''
  };
}

function pad0(s: any): string {
  return ('0' + s).substr(-2);
}

function log(message: any) {
  if (SLACK_LOG_WEBHOOK_URL.length) {
    UrlFetchApp.fetch(SLACK_LOG_WEBHOOK_URL, {
      method: 'post',
      payload: {
        payload: JSON.stringify({
          text: `${message}`,
          username: 'cal2slack',
          icon_emoji: ':robot_face:'
        })
      }
    })
  }
}
