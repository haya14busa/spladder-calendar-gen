// Spreadsheet ID of Spladder event.
var SPLADDER_SPREADSHEET_ID = getProperty('SPLADDER_SPREADSHEET_ID');
// Calenader ID of Spladder event. e.g. xxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com
var SPLADDER_CALENDAR_ID = getProperty('SPLADDER_CALENDAR_ID');
// If this value is 1, this script only updates latest round events.
var SPLADDER_CREATE_ONLY_LATEST = getProperty('SPLADDER_CREATE_ONLY_LATEST');

function main() {
  Logger.log('Start');

  const calendar = CalendarApp.getCalendarById(SPLADDER_CALENDAR_ID);
  const spreadsheet = SpreadsheetApp.openById(SPLADDER_SPREADSHEET_ID);
  const numbering = getSpladderEventNumbering(spreadsheet.getName());

  const sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();
    var matched = name.match(/^Challenges@R(\d+)/);
    if (matched == null || matched.length < 2) {
      continue; // Name did not match.
    }
    var round = matched[1];
    // Delete all events to avoid creating duplicated events.
    deletetEvents(calendar, numbering, round);
    createEventInSheet(calendar, sheet, numbering, round);
    if (SPLADDER_CREATE_ONLY_LATEST === "1") {
      // Update only latest round.
      break;
    }
  }
  Logger.log('Finished');
}

function getProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function getSpladderEventNumbering(sheet_name) {
  const matched = sheet_name.match(/Spladder #(\d+)/);
  if (matched.length < 2) {
    return -1;
  }
  return matched[1];
}

function deletetEvents(calendar, numbering, round) {
  // Search -90days ~ +30 days for existing events and delete them.
  const now = new Date();
  const start = new Date(now.getTime() - (3 * 30 * 60 * 60 * 24 * 1000));
  const end = new Date(now.getTime() + (30 * 60 * 60 * 24 * 1000));
  const tag = spladderTag(numbering, round);
  const events = calendar.getEvents(start, end, {'search': tag});
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (event.getTag('spladdercal') !== tag) {
      continue;
    }
    event.deleteEvent();
    Logger.log('Deleted event. id=' + event.getId() + ' title=' + event.getTitle());
    wait();
  }
}

function createEventInSheet(calendar, sheet, numbering, round) {
  Logger.log('Processing ' + sheet.getName() + ' ...');

  const values = sheet.getDataRange().getValues();

  if (values.length < 1) {
    Logger.log('Sheet ' + sheet.getName() + ' is empty.');
    return false;
  }

  const header = values[0];
  const idx = headerIndexMap(header);
  assertheaderIndexMap(idx);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var start = new Date(row[idx.scheduled_time]);
    var end = new Date(start.getTime() + (1 * 60 * 60 * 1000));
    var tag = spladderTag(numbering, round);
    var title = row[idx.alpha] + ' v.s. ' + row[idx.bravo] + ' ' + tag;
    var event = calendar.createEvent(title, start, end, {'description': tag});
    event.setTag('spladdercal', tag);
    event.setTag('round', round);
    event.setTag('numbering', numbering);
    Logger.log('Created event. id=' + event.getId() + ' title=' + event.getTitle());
    wait();
  }
  return true;
}

function spladderTag(numbering, round) {
  return '[SPLADDER#' + numbering + 'R' + round + ']';
}

function headerIndexMap(header) {
  var header_map = {
    'scheduled_time': -1,
    'alpha': -1,
    'bravo': -1
  };
  for (var i = 1; i < header.length; i++) {
    var name = header[i];
    if (name === '予定時間') {
      header_map.scheduled_time = i;
    } else if (name === '挑戦者チーム') {
      header_map.alpha = i;
    } else if (name === '防衛者チーム') {
      header_map.bravo = i;
    }
  }
  return header_map;
}

function assertheaderIndexMap(hmap) {
  for (key in hmap) {
    if (hmap[key] === -1) {
      throw key + ' index is not set';
    }
  }
}

// wait waits for short time to avoid API limit.
function wait() {
  Utilities.sleep(500);
}
