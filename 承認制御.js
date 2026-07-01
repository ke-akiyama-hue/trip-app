// ========================================
// ✅ 出張申請の承認・差戻し
// ========================================

function listPendingTripApprovals() {
  var userEmail = getCurrentUserEmail_();
  if (!userEmail) return [];
  return readTripRows_(function(r) {
    return r.status === TRIP_STATUS.SUBMITTED && r.approverEmail === userEmail;
  }).sort(function(a, b) { return (a.requestDate || '').localeCompare(b.requestDate || ''); });
}

function approveTripRequest(tripRequestId, comment) {
  return updateTripStatus_(tripRequestId, TRIP_STATUS.APPROVED, comment || '承認しました');
}

function rejectTripRequest(tripRequestId, reason, rejectTargetChoice) {
  reason = String(reason || '').trim();
  if (!reason) return { success: false, message: '差戻し理由を入力してください。' };
  return updateTripStatus_(tripRequestId, TRIP_STATUS.REJECTED, reason, rejectTargetChoice);
}

function updateTripStatus_(tripRequestId, newStatus, comment, rejectTargetChoice) {
  var userEmail = getCurrentUserEmail_();
  var trip = buildTripRequest_(tripRequestId);
  if (!trip) return { success: false, message: '申請が見つかりません。' };
  if (trip.status !== TRIP_STATUS.SUBMITTED) {
    return { success: false, message: '申請中のデータのみ承認・差戻しできます。' };
  }
  if (trip.approverEmail !== userEmail) {
    return { success: false, message: '承認権限がありません。' };
  }

  if (newStatus === TRIP_STATUS.APPROVED) {
    var nextStep = getNextWorkflowStep_(trip.routeId, trip.applicantEmail, trip.currentStep);
    if (nextStep) {
      trip.currentStep = nextStep.stepNo;
      trip.currentStepName = nextStep.stepName;
      trip.approverEmail = nextStep.approverEmail;
      trip.updatedAt = formatDateTime(new Date());
      var stepCalendarSync = syncTripCalendar_(trip);
      writeTripRow_(trip);
      appendHistory_(tripRequestId, '承認（' + (trip.currentStep - 1) + '/' + trip.totalSteps + '）', comment || '承認しました');
      return {
        success: true,
        message: appendCalendarSyncMessage_('承認しました。次の承認者（' + nextStep.stepName + '）に回りました。', stepCalendarSync),
        trip: buildTripRequest_(tripRequestId)
      };
    }
  }

  if (newStatus === TRIP_STATUS.REJECTED) {
    var route = resolveRejectRoute_(trip, rejectTargetChoice);
    if (route.mode === 'previous_step') {
      trip.currentStep = route.stepNo;
      trip.currentStepName = route.stepName;
      trip.approverEmail = route.approverEmail;
      trip.rejectReason = comment;
      trip.updatedAt = formatDateTime(new Date());
      var rejectStepCalendarSync = syncTripCalendar_(trip);
      writeTripRow_(trip);
      appendHistory_(tripRequestId, '差戻し（前ステップへ）', comment || '');
      return {
        success: true,
        message: appendCalendarSyncMessage_('差戻しました。前の承認者（' + route.stepName + '）に戻しました。', rejectStepCalendarSync),
        trip: enrichTripWithWorkflowStep_(buildTripRequest_(tripRequestId))
      };
    }
  }

  trip.status = newStatus;
  trip.updatedAt = formatDateTime(new Date());
  if (newStatus === TRIP_STATUS.APPROVED) {
    trip.approvedAt = trip.updatedAt;
    trip.rejectReason = '';
  }
  if (newStatus === TRIP_STATUS.REJECTED) {
    trip.rejectReason = comment;
    trip.approvedAt = '';
    trip.currentStep = 0;
    trip.currentStepName = '';
  }
  var calendarSync = syncTripCalendar_(trip);
  writeTripRow_(trip);

  var actionLabel = newStatus === TRIP_STATUS.APPROVED ? '承認' : '差戻し';
  appendHistory_(tripRequestId, actionLabel, comment || '');
  return { success: true, message: appendCalendarSyncMessage_(actionLabel + 'しました。', calendarSync), trip: buildTripRequest_(tripRequestId) };
}
