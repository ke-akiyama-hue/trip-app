// ========================================
// 📅 出張日程のGoogleカレンダー連携
// ========================================
//
// Webアプリは「次のユーザーとして実行: 自分（有料アカウント）」で公開する。
// そのため CalendarApp は操作ユーザーではなく、デプロイ実行ユーザーの権限で
// 設定済みカレンダーID（共通/事業所別）を直接操作する。

function authorizeTripCalendarAccess() {
  var results = [];
  checkTripCalendarAccess_(results, '共通', TRIP_SHARED_CALENDAR_ID);
  Object.keys(TRIP_OFFICE_CALENDAR_IDS || {}).forEach(function(office) {
    checkTripCalendarAccess_(results, office, TRIP_OFFICE_CALENDAR_IDS[office]);
  });

  var failed = results.filter(function(r) { return !r.success; });
  var lines = results.map(function(r) {
    return (r.success ? 'OK: ' : 'NG: ') + r.label + ' - ' + r.message;
  });
  var message = failed.length
    ? 'カレンダー接続に失敗した項目があります。\n\n' + lines.join('\n')
    : 'カレンダー権限と接続を確認できました。\n\n' + lines.join('\n');

  Logger.log(message);
  return message;
}

function authorizeTripCalendarWriteAccess() {
  var results = [];
  checkTripCalendarWriteAccess_(results, '共通', TRIP_SHARED_CALENDAR_ID);
  Object.keys(TRIP_OFFICE_CALENDAR_IDS || {}).forEach(function(office) {
    checkTripCalendarWriteAccess_(results, office, TRIP_OFFICE_CALENDAR_IDS[office]);
  });

  var failed = results.filter(function(r) { return !r.success; });
  var lines = results.map(function(r) {
    return (r.success ? 'OK: ' : 'NG: ') + r.label + ' - ' + r.message;
  });
  var message = failed.length
    ? 'カレンダー書き込み確認に失敗した項目があります。\n\n' + lines.join('\n')
    : 'カレンダー書き込み権限を確認できました。\n\n' + lines.join('\n');

  Logger.log(message);
  return message;
}

function checkTripCalendarAccess_(results, label, calendarId) {
  calendarId = String(calendarId || '').trim();
  if (!calendarId) {
    results.push({ success: true, label: label, message: '未設定のためスキップ' });
    return;
  }

  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      results.push({ success: false, label: label, message: 'カレンダーが見つからないか、閲覧権限がありません' });
      return;
    }
    results.push({ success: true, label: label, message: calendar.getName() });
  } catch (e) {
    results.push({ success: false, label: label, message: e.message });
  }
}

function checkTripCalendarWriteAccess_(results, label, calendarId) {
  calendarId = String(calendarId || '').trim();
  if (!calendarId) {
    results.push({ success: true, label: label, message: '未設定のためスキップ' });
    return;
  }

  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      results.push({ success: false, label: label, message: 'カレンダーが見つからないか、閲覧権限がありません' });
      return;
    }

    var start = new Date();
    start.setMinutes(start.getMinutes() + 5);
    var end = new Date(start.getTime() + 5 * 60 * 1000);
    var event = calendar.createEvent('【権限確認】出張申請カレンダー連携テスト', start, end, {
      description: '出張申請アプリのカレンダー書き込み権限確認用です。自動削除されます。'
    });
    event.deleteEvent();
    results.push({ success: true, label: label, message: calendar.getName() + ' に作成/削除できました' });
  } catch (e) {
    results.push({ success: false, label: label, message: e.message });
  }
}

function syncTripCalendar_(trip) {
  var result = { success: true, messages: [] };
  if (!trip) return result;

  try {
    if (trip.status !== TRIP_STATUS.SUBMITTED && trip.status !== TRIP_STATUS.APPROVED) {
      deleteTripCalendarEvents_(trip, result);
      return result;
    }

    syncSharedTripCalendarEvent_(trip, result);
    syncOfficeTripCalendarEvent_(trip, result);
  } catch (e) {
    result.success = false;
    result.messages.push('カレンダー同期エラー: ' + e.message);
    Logger.log('syncTripCalendar_: ' + e.message);
  }
  return result;
}

function syncSharedTripCalendarEvent_(trip, result) {
  var calendarId = String(TRIP_SHARED_CALENDAR_ID || '').trim();
  if (!calendarId) {
    if (trip.sharedCalendarEventId) trip.sharedCalendarEventId = '';
    return;
  }

  var synced = upsertTripCalendarEvent_(calendarId, trip, trip.sharedCalendarEventId, '共通');
  if (synced.success) {
    trip.sharedCalendarEventId = synced.eventId;
  } else {
    result.success = false;
    result.messages.push(synced.message);
  }
}

function syncOfficeTripCalendarEvent_(trip, result) {
  var employee = findEmployeeByEmail(trip.applicantEmail);
  var office = employee && employee.office ? String(employee.office).trim() : String(trip.calendarOffice || '').trim();
  var previousOffice = String(trip.calendarOffice || '').trim();

  if (previousOffice && previousOffice !== office && trip.officeCalendarEventId) {
    deleteCalendarEventById_(getOfficeTripCalendarId_(previousOffice), trip.officeCalendarEventId, result);
    trip.officeCalendarEventId = '';
  }

  trip.calendarOffice = office;
  var calendarId = getOfficeTripCalendarId_(office);
  if (!calendarId) {
    if (trip.officeCalendarEventId) trip.officeCalendarEventId = '';
    return;
  }

  var synced = upsertTripCalendarEvent_(calendarId, trip, trip.officeCalendarEventId, '事業所');
  if (synced.success) {
    trip.officeCalendarEventId = synced.eventId;
  } else {
    result.success = false;
    result.messages.push(synced.message);
  }
}

function getOfficeTripCalendarId_(office) {
  office = String(office || '').trim();
  if (!office || !TRIP_OFFICE_CALENDAR_IDS) return '';
  return String(TRIP_OFFICE_CALENDAR_IDS[office] || '').trim();
}

function upsertTripCalendarEvent_(calendarId, trip, eventId, label) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return { success: false, message: label + 'カレンダーが見つかりません: ' + calendarId };

    var start = parseTripCalendarDate_(trip.tripStart, false);
    var end = parseTripCalendarDate_(trip.tripEnd, true);
    if (!start || !end || end.getTime() <= start.getTime()) {
      return { success: false, message: 'カレンダー登録できない日程です: ' + trip.tripRequestId };
    }

    var title = buildTripCalendarTitle_(trip);
    var description = buildTripCalendarDescription_(trip);
    var event = eventId ? calendar.getEventById(eventId) : null;
    if (event) {
      event.setTitle(title);
      event.setTime(start, end);
      event.setDescription(description);
      event.setLocation(trip.destination || '');
    } else {
      event = calendar.createEvent(title, start, end, {
        description: description,
        location: trip.destination || ''
      });
    }
    setTripCalendarColor_(event, trip.status);
    return { success: true, eventId: event.getId() };
  } catch (e) {
    Logger.log('upsertTripCalendarEvent_(' + label + '): ' + e.message);
    return { success: false, message: label + 'カレンダー同期失敗: ' + e.message };
  }
}

function deleteTripCalendarEvents_(trip, result) {
  result = result || { success: true, messages: [] };
  deleteCalendarEventById_(TRIP_SHARED_CALENDAR_ID, trip.sharedCalendarEventId, result);
  deleteCalendarEventById_(getOfficeTripCalendarId_(trip.calendarOffice), trip.officeCalendarEventId, result);
  trip.sharedCalendarEventId = '';
  trip.officeCalendarEventId = '';
  trip.calendarOffice = '';
  return result;
}

function deleteCalendarEventById_(calendarId, eventId, result) {
  calendarId = String(calendarId || '').trim();
  eventId = String(eventId || '').trim();
  if (!calendarId || !eventId) return;

  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    var event = calendar ? calendar.getEventById(eventId) : null;
    if (event) event.deleteEvent();
  } catch (e) {
    if (result) {
      result.success = false;
      result.messages.push('カレンダー予定削除失敗: ' + e.message);
    }
    Logger.log('deleteCalendarEventById_: ' + e.message);
  }
}

function buildTripCalendarTitle_(trip) {
  return '【' + trip.status + '】出張: ' + (trip.applicantName || trip.applicantEmail) + ' / ' + (trip.destination || '行先未設定');
}

function buildTripCalendarDescription_(trip) {
  return [
    '出張申請ID: ' + trip.tripRequestId,
    '状態: ' + trip.status,
    '申請者: ' + (trip.applicantName || '') + ' (' + (trip.applicantEmail || '') + ')',
    '出張先: ' + (trip.destination || ''),
    '目的: ' + (trip.purpose || ''),
    '交通手段: ' + (trip.transport || ''),
    '宿泊先: ' + (trip.lodgingDestination || ''),
    '精算状況: ' + (trip.settlementStatus || ''),
    '更新日時: ' + (trip.updatedAt || '')
  ].join('\n');
}

function setTripCalendarColor_(event, status) {
  try {
    if (status === TRIP_STATUS.APPROVED) {
      event.setColor(CalendarApp.EventColor.GREEN);
    } else if (status === TRIP_STATUS.SUBMITTED) {
      event.setColor(CalendarApp.EventColor.YELLOW);
    }
  } catch (e) {
    Logger.log('setTripCalendarColor_: ' + e.message);
  }
}

function parseTripCalendarDate_(value, endOfDay) {
  var s = String(value || '').trim();
  if (!s) return null;

  var dateOnly = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateOnly) {
    return new Date(
      parseInt(dateOnly[1], 10),
      parseInt(dateOnly[2], 10) - 1,
      parseInt(dateOnly[3], 10),
      endOfDay ? 23 : 0,
      endOfDay ? 59 : 0,
      0
    );
  }

  var dateTime = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  if (dateTime) {
    return new Date(
      parseInt(dateTime[1], 10),
      parseInt(dateTime[2], 10) - 1,
      parseInt(dateTime[3], 10),
      parseInt(dateTime[4], 10),
      parseInt(dateTime[5], 10),
      0
    );
  }

  var parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function appendCalendarSyncMessage_(message, syncResult) {
  if (!syncResult || syncResult.success || !syncResult.messages.length) return message;
  return message + '\n\n※' + syncResult.messages.join('\n※');
}
