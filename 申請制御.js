// ========================================
// 📝 出張申請の作成・保存・提出
// ========================================

function validateTripPayload_(payload, isSubmit) {
  var errors = [];
  if (!payload.tripStart) errors.push('出張開始日を入力してください');
  if (!payload.tripEnd) errors.push('出張終了日を入力してください');
  if (payload.tripStart && payload.tripEnd && normalizeDate(payload.tripEnd) < normalizeDate(payload.tripStart)) {
    errors.push('出張終了日は開始日以降にしてください');
  }
  if (!String(payload.destination || '').trim()) errors.push('出張先を入力してください');
  if (!String(payload.purpose || '').trim()) errors.push('出張目的を入力してください');
  return errors;
}

function saveTripRequest(payload, submit) {
  payload = payload || {};
  var userEmail = getCurrentUserEmail_();
  if (!userEmail) {
    return { success: false, message: 'ログインユーザーを取得できません。Webアプリのデプロイ設定を確認してください。' };
  }

  var employee = findEmployeeByEmail(userEmail);
  var isSubmit = submit === true;
  if (isSubmit && !employee) {
    return { success: false, message: '共通マスタの社員マスタに登録されていません。' };
  }

  var errors = validateTripPayload_(payload, isSubmit);
  if (errors.length > 0) return { success: false, message: errors.join('\n') };

  var tripRequestId = String(payload.tripRequestId || '').trim() || generateTripRequestId_();
  var existing = buildTripRequest_(tripRequestId);

  if (existing) {
    if (existing.applicantEmail !== userEmail) {
      return { success: false, message: '他のユーザーの申請は編集できません。' };
    }
    if (existing.status !== TRIP_STATUS.DRAFT && existing.status !== TRIP_STATUS.REJECTED) {
      return { success: false, message: '現在のステータス（' + existing.status + '）では編集できません。' };
    }
    if (existing.settlementStatus !== SETTLEMENT_STATUS.NONE && existing.expenseClaimId) {
      return { success: false, message: '精算が開始されているため編集できません。' };
    }
  }

  var now = formatDateTime(new Date());
  var routeId = String(payload.routeId || '').trim() || (existing ? existing.routeId : '');
  if (isSubmit && isWorkflowLinked_() && !routeId) {
    routeId = getDefaultWorkflowRouteId_();
  }

  var wf = isSubmit
    ? resolveSubmitWorkflow_(routeId, userEmail, employee)
    : { success: true, routeId: existing ? existing.routeId : routeId, currentStep: 0, totalSteps: 0, currentStepName: '', approverEmail: '' };
  if (isSubmit && !wf.success) {
    return { success: false, message: wf.message };
  }

  var trip = {
    tripRequestId: tripRequestId,
    requestDate: existing ? existing.requestDate : normalizeDate(new Date()),
    applicantEmail: userEmail,
    applicantName: employee ? employee.name : (payload.applicantName || userEmail.split('@')[0]),
    tripStart: normalizeDate(payload.tripStart),
    tripEnd: normalizeDate(payload.tripEnd),
    destination: String(payload.destination || '').trim(),
    purpose: String(payload.purpose || '').trim(),
    transport: String(payload.transport || '').trim(),
    lodgingPlan: String(payload.lodgingPlan || '').trim(),
    estimatedCost: normalizeAmount(payload.estimatedCost),
    note: String(payload.note || ''),
    status: isSubmit ? TRIP_STATUS.SUBMITTED : TRIP_STATUS.DRAFT,
    settlementStatus: existing ? existing.settlementStatus : SETTLEMENT_STATUS.NONE,
    expenseClaimId: existing ? existing.expenseClaimId : '',
    approverEmail: isSubmit ? wf.approverEmail : (existing ? existing.approverEmail : ''),
    approvedAt: isSubmit ? '' : (existing ? existing.approvedAt : ''),
    rejectReason: isSubmit ? '' : (existing ? existing.rejectReason : ''),
    updatedAt: now,
    routeId: isSubmit ? wf.routeId : (existing ? existing.routeId : routeId),
    currentStep: isSubmit ? wf.currentStep : (existing ? existing.currentStep : 0),
    totalSteps: isSubmit ? wf.totalSteps : (existing ? existing.totalSteps : 0),
    currentStepName: isSubmit ? wf.currentStepName : (existing ? existing.currentStepName : '')
  };

  writeTripRow_(trip);
  appendHistory_(tripRequestId, isSubmit ? '申請' : '下書き保存', '');

  return {
    success: true,
    tripRequestId: tripRequestId,
    message: isSubmit ? '出張申請を提出しました。' : '下書きを保存しました。',
    trip: buildTripRequest_(tripRequestId)
  };
}

function deleteTripRequest(tripRequestId) {
  var userEmail = getCurrentUserEmail_();
  var trip = buildTripRequest_(tripRequestId);
  if (!trip) return { success: false, message: '申請が見つかりません。' };
  if (trip.applicantEmail !== userEmail) return { success: false, message: '削除権限がありません。' };
  if (trip.status !== TRIP_STATUS.DRAFT) return { success: false, message: '下書きのみ削除できます。' };

  var all = readTripRows_(function(r) { return r.tripRequestId !== tripRequestId; });
  writeAllTripRows_(getSpreadsheet_().getSheetByName(SHEET_TRIPS), all);
  appendHistory_(tripRequestId, '削除', '');
  return { success: true, message: '下書きを削除しました。' };
}

function cancelTripRequest(tripRequestId) {
  var userEmail = getCurrentUserEmail_();
  var trip = buildTripRequest_(tripRequestId);
  if (!trip) return { success: false, message: '申請が見つかりません。' };
  if (trip.applicantEmail !== userEmail) return { success: false, message: '取消権限がありません。' };
  if (trip.status !== TRIP_STATUS.APPROVED) return { success: false, message: '承認済の申請のみ取消できます。' };
  if (trip.settlementStatus !== SETTLEMENT_STATUS.NONE) {
    return { success: false, message: '精算が開始されているため取消できません。' };
  }

  trip.status = TRIP_STATUS.CANCELLED;
  trip.updatedAt = formatDateTime(new Date());
  writeTripRow_(trip);
  appendHistory_(tripRequestId, '取消', '');
  return { success: true, message: '出張申請を取り消しました。' };
}

function listMyTripRequests() {
  var userEmail = getCurrentUserEmail_();
  return readTripRows_(function(r) { return r.applicantEmail === userEmail; })
    .sort(function(a, b) { return (b.updatedAt || '').localeCompare(a.updatedAt || ''); });
}

function getTripRequestDetail(tripRequestId) {
  var trip = buildTripRequest_(tripRequestId);
  if (!trip) return { success: false, message: '申請が見つかりません。' };
  var userEmail = getCurrentUserEmail_();
  var canView = trip.applicantEmail === userEmail || trip.approverEmail === userEmail;
  if (!canView) return { success: false, message: '閲覧権限がありません。' };
  return { success: true, trip: enrichTripWithWorkflowStep_(trip) };
}
