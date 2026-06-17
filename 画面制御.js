function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('出張申請')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialAppData() {
  var userEmail = getCurrentUserEmail_();
  var employee = findEmployeeByEmail(userEmail);
  return {
    userEmail: userEmail,
    employee: employee,
    isApprover: listPendingTripApprovals().length > 0,
    myTrips: listMyTripRequests(),
    pendingApprovals: listPendingTripApprovals(),
    statusLabels: TRIP_STATUS,
    settlementLabels: SETTLEMENT_STATUS,
    workflowRoutes: getAvailableWorkflowRoutes(),
    workflowLinked: isWorkflowLinked_(),
    submitReadiness: getSubmitReadiness_(userEmail, employee)
  };
}

function previewTripWorkflowRouteApi(routeId, applicantEmail) {
  return previewWorkflowRoute_(routeId, applicantEmail || getCurrentUserEmail_());
}
