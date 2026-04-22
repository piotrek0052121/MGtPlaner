const ORDER_STATUSES = ["Dokumentacja", "Produkcja", "Kosmetyka", "Zakonczone"];
const POSITION_DEPARTMENT_STATUSES = [
  "Dokumentacja",
  "Maszynownia",
  "Lakiernia",
  "Kompletacja",
  "Kosmetyka",
  "Zakonczone",
];
const KPI_POSITION_STATUSES = POSITION_DEPARTMENT_STATUSES.slice();
const KPI_POSITION_MIXED_STATUS = "Mieszane statusy";
const KPI_POSITION_EMPTY_STATUS = "Brak pozycji";
const POSITION_WORKFLOW_STATUS_OPTIONS = [
  { value: "pending", label: "Planowana" },
  { value: "in_progress", label: "W toku" },
  { value: "done", label: "Zakonczona" },
];
const DEPARTMENTS = ["Maszynownia", "Lakiernia", "Kompletacja", "Kosmetyka"];
const PROCESS_FLOW = [
  { key: "machining", department: "Maszynownia" },
  { key: "painting", department: "Lakiernia" },
  { key: "assembly", department: "Kompletacja" },
];
const MATERIALS = [
  { key: "wood", label: "Drewno" },
  { key: "corpus", label: "Korpusy" },
  { key: "glass", label: "Szyby" },
  { key: "hardware", label: "Okucia" },
  { key: "accessories", label: "Akcesoria" },
];
const USER_VISIBLE_SECTIONS = [
  "orders",
  "kpi",
  "gantt",
  "reports",
  "execution",
  "skills",
  "feedback",
  "archive",
  "users",
  "settings",
];
const SECTION_LABELS = {
  orders: "Zamowienia",
  kpi: "KPI",
  gantt: "Gantt",
  reports: "Raporty",
  execution: "Wykonanie",
  skills: "Matryca umiejetnosci",
  feedback: "Informacja zwrotna",
  archive: "Archiwum",
  users: "Uzytkownicy",
  settings: "Ustawienia",
};

const state = {
  orders: [],
  feedbackEvents: [],
  users: [],
  settings: {
    minutesPerShift: 480,
    workingDays: [1, 2, 3, 4, 5],
    weekdayShifts: { 0: 0, 1: 2, 2: 2, 3: 2, 4: 2, 5: 2, 6: 0 },
    stationOvertime: {},
    calendarOverrides: {},
  },
  stations: [],
  stationSettings: {},
  technologies: {},
  materialRules: {},
  skillWorkers: [],
  skillAvailability: {},
  databases: [],
  activeDatabaseKey: "default",
  databaseAccessMap: {},
};

const ui = {
  selectedTechnology: "",
  selectedProcess: "machining",
  selectedOrderId: "",
  editingUserId: "",
  orderDetailsEditMode: false,
  editingPositionId: "",
  kpiFilter: null,
  ganttExpandedOrders: {},
  ganttCalendarScrollLeft: 0,
  ganttCalendarScrollTop: 0,
  ganttDragOrderId: "",
  feedbackSelectionByOrder: {},
  archiveSelectionOrderIds: [],
  ordersSortKey: "createdAt",
  ordersSortDir: "desc",
  currentView: "dashboard",
  currentUser: null,
  pendingAttachmentPositionId: "",
  ganttDaysBackToShow: 30,
  ganttDaysForwardToShow: 180,
  editingSkillWorkerId: "",
  selectedSkillWorkerIds: [],
};

const el = {
  navButtons: Array.from(document.querySelectorAll(".nav-btn")),
  views: Array.from(document.querySelectorAll(".view")),
  todayDate: document.querySelector("#todayDate"),
  activeOrdersCount: document.querySelector("#activeOrdersCount"),
  goButtons: Array.from(document.querySelectorAll("[data-go]")),
  loginForm: document.querySelector("#loginForm"),
  loginDatabaseSelect: document.querySelector("#loginDatabaseSelect"),
  loginInput: document.querySelector("#loginInput"),
  loginPasswordInput: document.querySelector("#loginPasswordInput"),
  loginBtn: document.querySelector("#loginBtn"),
  logoutBtn: document.querySelector("#logoutBtn"),
  loginStatusText: document.querySelector("#loginStatusText"),
  currentUserInfo: document.querySelector("#currentUserInfo"),
  postLoginLinks: document.querySelector("#postLoginLinks"),
  openOrderModalBtn: document.querySelector("#openOrderModalBtn"),
  openPositionModalBtn: document.querySelector("#openPositionModalBtn"),
  orderModal: document.querySelector("#orderModal"),
  orderDetailsModal: document.querySelector("#orderDetailsModal"),
  positionModal: document.querySelector("#positionModal"),
  positionModalTitle: document.querySelector("#positionModalTitle"),
  positionAttachmentHint: document.querySelector("#positionAttachmentHint"),
  openCurrentAttachmentBtn: document.querySelector("#openCurrentAttachmentBtn"),
  clearAttachmentWrap: document.querySelector("#clearAttachmentWrap"),
  clearAttachmentCheckbox: document.querySelector('input[name="clearAttachment"]'),
  orderForm: document.querySelector("#orderForm"),
  saveOrderBtn: document.querySelector("#saveOrderBtn"),
  orderStatusCreateSelect: document.querySelector("#orderStatusCreateSelect"),
  positionForm: document.querySelector("#positionForm"),
  savePositionBtn: document.querySelector("#savePositionBtn"),
  technologySelect: document.querySelector("#technologySelect"),
  ordersTable: document.querySelector(".orders-table"),
  ordersTableBody: document.querySelector("#ordersTableBody"),
  ordersSearchInput: document.querySelector("#ordersSearchInput"),
  ordersPositionSearchInput: document.querySelector("#ordersPositionSearchInput"),
  ordersStatusFilter: document.querySelector("#ordersStatusFilter"),
  ordersWeekFilterInput: document.querySelector("#ordersWeekFilterInput"),
  ordersImportFileInput: document.querySelector("#ordersImportFileInput"),
  importOrdersBtn: document.querySelector("#importOrdersBtn"),
  ordersImportResult: document.querySelector("#ordersImportResult"),
  selectCompletedOrdersBtn: document.querySelector("#selectCompletedOrdersBtn"),
  clearOrdersFiltersBtn: document.querySelector("#clearOrdersFiltersBtn"),
  archiveCompletedOrdersBtn: document.querySelector("#archiveCompletedOrdersBtn"),
  archiveSearchInput: document.querySelector("#archiveSearchInput"),
  archiveWeekFilterInput: document.querySelector("#archiveWeekFilterInput"),
  archiveDateFilterInput: document.querySelector("#archiveDateFilterInput"),
  clearArchiveFiltersBtn: document.querySelector("#clearArchiveFiltersBtn"),
  archiveTableBody: document.querySelector("#archiveTableBody"),
  orderDetailsTitle: document.querySelector("#orderDetailsTitle"),
  selectedOrderForm: document.querySelector("#selectedOrderForm"),
  selectedOrderNumberInput: document.querySelector("#selectedOrderNumberInput"),
  selectedOrderEntryDateInput: document.querySelector("#selectedOrderEntryDateInput"),
  selectedOrderPlannedDateInput: document.querySelector("#selectedOrderPlannedDateInput"),
  selectedOrderStatusSelect: document.querySelector("#selectedOrderStatusSelect"),
  selectedOrderOwnerInput: document.querySelector("#selectedOrderOwnerInput"),
  selectedOrderClientInput: document.querySelector("#selectedOrderClientInput"),
  selectedOrderColorInput: document.querySelector("#selectedOrderColorInput"),
  selectedOrderFramesInput: document.querySelector("#selectedOrderFramesInput"),
  selectedOrderSashesInput: document.querySelector("#selectedOrderSashesInput"),
  selectedOrderExtrasInput: document.querySelector("#selectedOrderExtrasInput"),
  editSelectedOrderBtn: document.querySelector("#editSelectedOrderBtn"),
  saveSelectedOrderBtn: document.querySelector("#saveSelectedOrderBtn"),
  selectedOrderPositionsBody: document.querySelector("#selectedOrderPositionsBody"),
  positionAttachmentUploadInput: document.querySelector("#positionAttachmentUploadInput"),
  kpiCards: document.querySelector("#kpiCards"),
  kpiDrilldownTitle: document.querySelector("#kpiDrilldownTitle"),
  kpiDrilldownBody: document.querySelector("#kpiDrilldownBody"),
  ganttBoard: document.querySelector("#ganttBoard"),
  ganttDaysBackInput: document.querySelector("#ganttDaysBackInput"),
  ganttDaysForwardInput: document.querySelector("#ganttDaysForwardInput"),
  applyGanttDaysBtn: document.querySelector("#applyGanttDaysBtn"),
  ganttSearchInput: document.querySelector("#ganttSearchInput"),
  ganttStatusFilter: document.querySelector("#ganttStatusFilter"),
  clearGanttFiltersBtn: document.querySelector("#clearGanttFiltersBtn"),
  reportDate: document.querySelector("#reportDate"),
  reportStationSelect: document.querySelector("#reportStationSelect"),
  runReportBtn: document.querySelector("#runReportBtn"),
  printReportBtn: document.querySelector("#printReportBtn"),
  reportSummary: document.querySelector("#reportSummary"),
  reportCockpit: document.querySelector("#reportCockpit"),
  executionPeriodMode: document.querySelector("#executionPeriodMode"),
  executionAnchorDate: document.querySelector("#executionAnchorDate"),
  runExecutionBtn: document.querySelector("#runExecutionBtn"),
  executionSummary: document.querySelector("#executionSummary"),
  executionKpiCards: document.querySelector("#executionKpiCards"),
  executionDepartmentBody: document.querySelector("#executionDepartmentBody"),
  executionOrdersBody: document.querySelector("#executionOrdersBody"),
  workloadTableBody: document.querySelector("#workloadTableBody"),
  skillWorkerForm: document.querySelector("#skillWorkerForm"),
  skillWorkerNameInput: document.querySelector("#skillWorkerNameInput"),
  skillWorkerDepartmentSelect: document.querySelector("#skillWorkerDepartmentSelect"),
  skillWorkerPrimaryStationSelect: document.querySelector("#skillWorkerPrimaryStationSelect"),
  skillWorkerAssignedShiftSelect: document.querySelector("#skillWorkerAssignedShiftSelect"),
  skillWorkerActiveInput: document.querySelector("#skillWorkerActiveInput"),
  skillWorkerSkillsWrap: document.querySelector("#skillWorkerSkillsWrap"),
  skillWorkerSubmitBtn: document.querySelector("#skillWorkerSubmitBtn"),
  cancelSkillWorkerEditBtn: document.querySelector("#cancelSkillWorkerEditBtn"),
  skillWorkersList: document.querySelector("#skillWorkersList"),
  selectAllSkillWorkersBtn: document.querySelector("#selectAllSkillWorkersBtn"),
  clearSkillWorkersSelectionBtn: document.querySelector("#clearSkillWorkersSelectionBtn"),
  skillWorkersBulkShiftSelect: document.querySelector("#skillWorkersBulkShiftSelect"),
  applySkillWorkersBulkShiftBtn: document.querySelector("#applySkillWorkersBulkShiftBtn"),
  skillWorkersBulkStatus: document.querySelector("#skillWorkersBulkStatus"),
  skillsAllocationMode: document.querySelector("#skillsAllocationMode"),
  skillsAllocationAnchorDate: document.querySelector("#skillsAllocationAnchorDate"),
  runSkillsAllocationBtn: document.querySelector("#runSkillsAllocationBtn"),
  skillsAllocationSummary: document.querySelector("#skillsAllocationSummary"),
  skillsAllocationTableBody: document.querySelector("#skillsAllocationTableBody"),
  skillAvailabilityWorkerSelect: document.querySelector("#skillAvailabilityWorkerSelect"),
  skillAvailabilityStartDate: document.querySelector("#skillAvailabilityStartDate"),
  skillAvailabilityDays: document.querySelector("#skillAvailabilityDays"),
  renderSkillAvailabilityBtn: document.querySelector("#renderSkillAvailabilityBtn"),
  saveSkillAvailabilityBtn: document.querySelector("#saveSkillAvailabilityBtn"),
  skillAvailabilityStatus: document.querySelector("#skillAvailabilityStatus"),
  skillAvailabilityGridWrap: document.querySelector("#skillAvailabilityGridWrap"),
  feedbackOrderSelect: document.querySelector("#feedbackOrderSelect"),
  feedbackOrderStatusFilter: document.querySelector("#feedbackOrderStatusFilter"),
  feedbackOrderPositionStatusFilter: document.querySelector("#feedbackOrderPositionStatusFilter"),
  feedbackDepartmentSelect: document.querySelector("#feedbackDepartmentSelect"),
  feedbackSearchInput: document.querySelector("#feedbackSearchInput"),
  feedbackWorkflowStatusSelect: document.querySelector("#feedbackWorkflowStatusSelect"),
  feedbackProgressPercentInput: document.querySelector("#feedbackProgressPercentInput"),
  feedbackClearFiltersBtn: document.querySelector("#feedbackClearFiltersBtn"),
  feedbackSelectionInfo: document.querySelector("#feedbackSelectionInfo"),
  feedbackPositions: document.querySelector("#feedbackPositions"),
  feedbackItemTemplate: document.querySelector("#feedbackItemTemplate"),
  selectAllPositionsBtn: document.querySelector("#selectAllPositionsBtn"),
  clearSelectedPositionsBtn: document.querySelector("#clearSelectedPositionsBtn"),
  startPositionsBtn: document.querySelector("#startPositionsBtn"),
  finishPositionsBtn: document.querySelector("#finishPositionsBtn"),
  setFeedbackDepartmentBtn: document.querySelector("#setFeedbackDepartmentBtn"),
  applyFeedbackChangesBtn: document.querySelector("#applyFeedbackChangesBtn"),
  userForm: document.querySelector("#userForm"),
  userCanCreateDatabasesInput: document.querySelector('#userForm input[name="canCreateDatabases"]'),
  userRoleSelect: document.querySelector("#userRoleSelect"),
  userSubmitBtn: document.querySelector("#userSubmitBtn"),
  cancelUserEditBtn: document.querySelector("#cancelUserEditBtn"),
  usersAdminHint: document.querySelector("#usersAdminHint"),
  userPermissionsGrid: document.querySelector("#userPermissionsGrid"),
  usersList: document.querySelector("#usersList"),
  databaseAccessMatrixWrap: document.querySelector("#databaseAccessMatrixWrap"),
  saveDatabaseAccessBtn: document.querySelector("#saveDatabaseAccessBtn"),
  databaseAccessStatusText: document.querySelector("#databaseAccessStatusText"),
  settingsForm: document.querySelector("#settingsForm"),
  databaseSelect: document.querySelector("#databaseSelect"),
  switchDatabaseBtn: document.querySelector("#switchDatabaseBtn"),
  newDatabaseNameInput: document.querySelector("#newDatabaseNameInput"),
  createDatabaseBtn: document.querySelector("#createDatabaseBtn"),
  databaseStatusText: document.querySelector("#databaseStatusText"),
  databaseAdminHint: document.querySelector("#databaseAdminHint"),
  overtimeDateInput: document.querySelector("#overtimeDateInput"),
  overtimeStationSelect: document.querySelector("#overtimeStationSelect"),
  overtimeMinutesInput: document.querySelector("#overtimeMinutesInput"),
  addStationOvertimeBtn: document.querySelector("#addStationOvertimeBtn"),
  stationOvertimeTableBody: document.querySelector("#stationOvertimeTableBody"),
  stationSettingsTableBody: document.querySelector("#stationSettingsTableBody"),
  addStationRowBtn: document.querySelector("#addStationRowBtn"),
  saveStationSettingsBtn: document.querySelector("#saveStationSettingsBtn"),
  materialRulesTableBody: document.querySelector("#materialRulesTableBody"),
  saveMaterialRulesBtn: document.querySelector("#saveMaterialRulesBtn"),
  technologyEditorSelect: document.querySelector("#technologyEditorSelect"),
  processEditorSelect: document.querySelector("#processEditorSelect"),
  technologyAllocationTableBody: document.querySelector("#technologyAllocationTableBody"),
  saveTechnologyAllocationBtn: document.querySelector("#saveTechnologyAllocationBtn"),
  addTechnologyBtn: document.querySelector("#addTechnologyBtn"),
};

init().catch((error) => {
  console.error(error);
  alert(`Nie udalo sie uruchomic aplikacji: ${error.message}`);
});

async function init() {
  hydrateDates();
  fillStatusSelectors();
  bindNav();
  bindActions();
  await loadPublicDatabases();
  await restoreSessionUser();
  syncUserFormMode();
  await reloadAndRender();
  const firstSection = firstVisibleSectionForCurrentUser();
  setView(firstSection || "dashboard", true);
}

function bindNav() {
  el.navButtons.forEach((btn) => {
    btn.addEventListener("click", () => setView(btn.dataset.section));
  });
  el.goButtons.forEach((btn) => {
    btn.addEventListener("click", () => setView(btn.dataset.go));
  });
}

function bindActions() {
  el.loginForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    safeAction(loginUser);
  });
  el.loginDatabaseSelect?.addEventListener("change", () => {
    state.activeDatabaseKey = String(el.loginDatabaseSelect.value || state.activeDatabaseKey || "default");
  });
  el.logoutBtn?.addEventListener("click", () => safeAction(logoutUser));
  el.postLoginLinks?.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-go]");
    if (!btn) {
      return;
    }
    setView(btn.dataset.go);
  });
  if (el.openOrderModalBtn) {
    el.openOrderModalBtn.addEventListener("click", () => openModal("orderModal"));
  }
  if (el.importOrdersBtn) {
    el.importOrdersBtn.addEventListener("click", () => safeAction(importOrdersFromExcel));
  }
  if (el.openPositionModalBtn) {
    el.openPositionModalBtn.addEventListener("click", () => openPositionModalForSelectedOrder());
  }
  if (el.openCurrentAttachmentBtn) {
    el.openCurrentAttachmentBtn.addEventListener("click", () => safeAction(openCurrentEditingAttachment));
  }
  document.addEventListener("click", (event) => {
    const closeBtn = event.target.closest("[data-close-modal]");
    if (!closeBtn) {
      return;
    }
    closeModal(closeBtn.dataset.closeModal);
  });
  if (el.saveOrderBtn) {
    el.saveOrderBtn.addEventListener("click", () => safeAction(createOrder));
  }
  if (el.savePositionBtn) {
    el.savePositionBtn.addEventListener("click", () => safeAction(createPosition));
  }
  if (el.ordersTableBody) {
    el.ordersTableBody.addEventListener("click", onOrderTableClick);
    el.ordersTableBody.addEventListener("change", onOrderTableChange);
  }
  el.ordersTable?.querySelector("thead")?.addEventListener("click", onOrdersHeaderClick);
  if (el.archiveTableBody) {
    el.archiveTableBody.addEventListener("click", onArchiveTableClick);
  }
  el.ordersSearchInput?.addEventListener("input", renderOrdersTable);
  el.ordersPositionSearchInput?.addEventListener("input", renderOrdersTable);
  el.ordersStatusFilter?.addEventListener("change", renderOrdersTable);
  el.ordersWeekFilterInput?.addEventListener("input", renderOrdersTable);
  el.selectCompletedOrdersBtn?.addEventListener("click", selectCompletedOrdersForArchive);
  el.clearOrdersFiltersBtn?.addEventListener("click", clearOrdersFilters);
  el.archiveCompletedOrdersBtn?.addEventListener("click", () => safeAction(archiveCompletedOrders));
  el.archiveSearchInput?.addEventListener("input", renderArchiveTable);
  el.archiveWeekFilterInput?.addEventListener("input", renderArchiveTable);
  el.archiveDateFilterInput?.addEventListener("change", renderArchiveTable);
  el.clearArchiveFiltersBtn?.addEventListener("click", clearArchiveFilters);
  if (el.editSelectedOrderBtn) {
    el.editSelectedOrderBtn.addEventListener("click", startSelectedOrderEditMode);
  }
  if (el.saveSelectedOrderBtn) {
    el.saveSelectedOrderBtn.addEventListener("click", () => safeAction(saveSelectedOrder));
  }
  if (el.selectedOrderPositionsBody) {
    el.selectedOrderPositionsBody.addEventListener("click", onOrderPositionsTableClick);
  }
  el.positionAttachmentUploadInput?.addEventListener("change", () => safeAction(uploadAttachmentFromPicker));
  el.kpiCards.addEventListener("click", onKpiCardClick);
  el.kpiDrilldownBody?.addEventListener("click", onKpiDrilldownClick);
  el.runReportBtn.addEventListener("click", renderReport);
  el.printReportBtn?.addEventListener("click", printReportCockpit);
  el.reportDate?.addEventListener("change", renderReport);
  el.reportStationSelect?.addEventListener("change", renderReport);
  el.runExecutionBtn?.addEventListener("click", renderExecution);
  el.executionPeriodMode?.addEventListener("change", renderExecution);
  el.executionAnchorDate?.addEventListener("change", renderExecution);
  el.skillWorkerForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    safeAction(saveSkillWorker);
  });
  el.skillWorkerDepartmentSelect?.addEventListener("change", renderSkillWorkerForm);
  el.cancelSkillWorkerEditBtn?.addEventListener("click", cancelSkillWorkerEdit);
  el.skillWorkersList?.addEventListener("click", onSkillWorkersListClick);
  el.skillWorkersList?.addEventListener("change", onSkillWorkersListChange);
  el.selectAllSkillWorkersBtn?.addEventListener("click", selectAllSkillWorkers);
  el.clearSkillWorkersSelectionBtn?.addEventListener("click", clearSkillWorkerSelection);
  el.applySkillWorkersBulkShiftBtn?.addEventListener("click", () => safeAction(applyBulkShiftForSkillWorkers));
  el.runSkillsAllocationBtn?.addEventListener("click", renderSkillsAllocation);
  el.skillsAllocationMode?.addEventListener("change", renderSkillsAllocation);
  el.skillsAllocationAnchorDate?.addEventListener("change", renderSkillsAllocation);
  el.renderSkillAvailabilityBtn?.addEventListener("click", renderSkillAvailabilityCalendar);
  el.saveSkillAvailabilityBtn?.addEventListener("click", () => safeAction(saveSkillAvailabilityCalendar));
  el.skillAvailabilityWorkerSelect?.addEventListener("change", renderSkillAvailabilityCalendar);
  el.skillAvailabilityStartDate?.addEventListener("change", renderSkillAvailabilityCalendar);
  el.skillAvailabilityDays?.addEventListener("change", renderSkillAvailabilityCalendar);
  el.feedbackOrderSelect.addEventListener("change", renderFeedbackPositions);
  el.feedbackOrderStatusFilter?.addEventListener("change", () => {
    renderOrderSelectors();
    renderFeedbackPositions();
  });
  el.feedbackOrderPositionStatusFilter?.addEventListener("change", () => {
    renderOrderSelectors();
    renderFeedbackPositions();
  });
  el.feedbackSearchInput?.addEventListener("input", renderFeedbackPositions);
  el.feedbackClearFiltersBtn?.addEventListener("click", clearFeedbackFilters);
  el.feedbackPositions?.addEventListener("click", onFeedbackPositionRowClick);
  el.feedbackPositions?.addEventListener("change", onFeedbackPositionCheckboxChange);
  el.feedbackPositions?.addEventListener("input", onFeedbackPositionInput);
  el.selectAllPositionsBtn.addEventListener("click", selectAllFeedbackPositions);
  el.clearSelectedPositionsBtn?.addEventListener("click", clearFeedbackSelection);
  el.startPositionsBtn.addEventListener("click", () => safeAction(() => updateSelectedPositions("start")));
  el.finishPositionsBtn.addEventListener("click", () => safeAction(() => updateSelectedPositions("finish")));
  el.setFeedbackDepartmentBtn.addEventListener("click", () => safeAction(setDepartmentForSelectedPositions));
  el.applyFeedbackChangesBtn?.addEventListener("click", () => safeAction(applyFeedbackChangesForSelected));
  el.userForm.addEventListener("submit", (event) => {
    event.preventDefault();
    safeAction(createUser);
  });
  el.userRoleSelect?.addEventListener("change", syncUserPermissionsEnabledState);
  el.cancelUserEditBtn?.addEventListener("click", cancelUserEdit);
  el.usersList?.addEventListener("click", onUsersListClick);
  el.databaseAccessMatrixWrap?.addEventListener("change", onDatabaseAccessMatrixChange);
  el.saveDatabaseAccessBtn?.addEventListener("click", () => safeAction(saveDatabaseAccessMatrix));
  el.settingsForm.addEventListener("submit", (event) => {
    event.preventDefault();
    safeAction(saveSettings);
  });
  el.switchDatabaseBtn?.addEventListener("click", () => safeAction(switchActiveDatabase));
  el.createDatabaseBtn?.addEventListener("click", () => safeAction(createDatabaseVariant));
  el.addStationOvertimeBtn?.addEventListener("click", () => safeAction(upsertStationOvertime));
  el.stationOvertimeTableBody?.addEventListener("click", onStationOvertimeTableClick);
  el.addStationRowBtn.addEventListener("click", addStationEditorRow);
  el.saveStationSettingsBtn.addEventListener("click", () => safeAction(saveStationSettings));
  el.stationSettingsTableBody.addEventListener("click", onStationSettingsTableClick);
  el.saveMaterialRulesBtn.addEventListener("click", () => safeAction(saveMaterialRules));
  el.processEditorSelect.addEventListener("change", () => {
    ui.selectedProcess = el.processEditorSelect.value;
    renderTechnologyAllocationEditor();
  });
  el.technologyEditorSelect.addEventListener("change", () => {
    ui.selectedTechnology = el.technologyEditorSelect.value;
    renderTechnologyAllocationEditor();
  });
  el.saveTechnologyAllocationBtn.addEventListener("click", () => safeAction(saveTechnologyAllocations));
  el.addTechnologyBtn.addEventListener("click", () => safeAction(addTechnology));
  el.ganttBoard.addEventListener("click", onGanttOrderClick);
  el.ganttBoard.addEventListener("change", onGanttCalendarToggleChange);
  el.ganttBoard.addEventListener("dragstart", onGanttDragStart);
  el.ganttBoard.addEventListener("dragover", onGanttDragOver);
  el.ganttBoard.addEventListener("drop", onGanttDrop);
  el.ganttBoard.addEventListener("dragend", onGanttDragEnd);
  el.applyGanttDaysBtn.addEventListener("click", applyGanttDaysCount);
  el.ganttSearchInput?.addEventListener("input", renderGantt);
  el.ganttStatusFilter?.addEventListener("change", renderGantt);
  el.clearGanttFiltersBtn?.addEventListener("click", clearGanttFilters);
  el.ganttDaysBackInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    applyGanttDaysCount();
  });
  el.ganttDaysForwardInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    applyGanttDaysCount();
  });
  syncUserPermissionsEnabledState();
}

function openModal(id) {
  const modal = el[id];
  if (!modal) {
    return;
  }
  modal.classList.remove("hidden");
  if (id === "orderModal") {
    el.orderForm.querySelector("input[name='orderNumber']")?.focus();
  }
  if (id === "positionModal") {
    el.positionForm.querySelector("input[name='positionNumber']")?.focus();
  }
  if (id === "orderDetailsModal") {
    if (ui.orderDetailsEditMode) {
      el.selectedOrderNumberInput?.focus();
    } else {
      el.editSelectedOrderBtn?.focus();
    }
  }
}

function closeModal(id) {
  const modal = el[id];
  if (!modal) {
    return;
  }
  modal.classList.add("hidden");
  if (id === "orderDetailsModal") {
    ui.orderDetailsEditMode = false;
    renderSelectedOrderDetails();
  }
  if (id === "positionModal") {
    ui.editingPositionId = "";
    resetPositionFormState();
    if (el.positionModalTitle) {
      el.positionModalTitle.textContent = "Nowa pozycja";
    }
    if (el.savePositionBtn) {
      el.savePositionBtn.textContent = "Dodaj pozycje";
    }
  }
}

function openPositionModalForSelectedOrder() {
  const order = state.orders.find((item) => item.id === ui.selectedOrderId);
  if (!order) {
    alert("Najpierw dodaj zamowienie.");
    return;
  }
  ui.editingPositionId = "";
  resetPositionFormState();
  if (el.positionModalTitle) {
    el.positionModalTitle.textContent = `Nowa pozycja - ${order.orderNumber}`;
  }
  if (el.savePositionBtn) {
    el.savePositionBtn.textContent = "Dodaj pozycje";
  }
  openModal("positionModal");
}

function openPositionModalForEdit(positionId) {
  const order = state.orders.find((item) => item.id === ui.selectedOrderId);
  const position = order?.positions?.find((item) => item.id === positionId);
  if (!order || !position) {
    alert("Nie znaleziono pozycji do edycji.");
    return;
  }
  ui.editingPositionId = position.id;
  fillPositionForm(position);
  if (el.positionModalTitle) {
    el.positionModalTitle.textContent = `Edycja pozycji ${position.positionNumber} - ${order.orderNumber}`;
  }
  if (el.savePositionBtn) {
    el.savePositionBtn.textContent = "Zapisz pozycje";
  }
  if (el.positionAttachmentHint) {
    el.positionAttachmentHint.textContent = position.attachmentName
      ? `Aktualny zalacznik: ${position.attachmentName} (zostaw puste, aby nie zmieniac)`
      : "Brak zalacznika";
  }
  if (el.clearAttachmentWrap) {
    el.clearAttachmentWrap.classList.toggle("panel-hidden", !position.attachmentName);
  }
  if (el.openCurrentAttachmentBtn) {
    el.openCurrentAttachmentBtn.classList.toggle("panel-hidden", !position.attachmentName);
  }
  openModal("positionModal");
}

function resetPositionFormState() {
  if (!el.positionForm) {
    return;
  }
  el.positionForm.reset();
  const defaults = [
    ['input[name="positionFramesCount"]', "1"],
    ['input[name="positionSashesCount"]', "1"],
    ['input[name="slemieCount"]', "0"],
    ['input[name="slupekStalyCount"]', "0"],
    ['input[name="przymykCount"]', "0"],
    ['input[name="niskiProgCount"]', "0"],
  ];
  defaults.forEach(([selector, value]) => {
    const node = el.positionForm.querySelector(selector);
    if (node) {
      node.value = value;
    }
  });
  const shapeRect = el.positionForm.querySelector('input[name="shapeRect"]');
  if (shapeRect) {
    shapeRect.checked = true;
  }
  const shapeSkos = el.positionForm.querySelector('input[name="shapeSkos"]');
  if (shapeSkos) {
    shapeSkos.checked = false;
  }
  const shapeLuk = el.positionForm.querySelector('input[name="shapeLuk"]');
  if (shapeLuk) {
    shapeLuk.checked = false;
  }
  if (el.positionAttachmentHint) {
    el.positionAttachmentHint.textContent = "";
  }
  if (el.clearAttachmentCheckbox) {
    el.clearAttachmentCheckbox.checked = false;
  }
  if (el.clearAttachmentWrap) {
    el.clearAttachmentWrap.classList.add("panel-hidden");
  }
  if (el.openCurrentAttachmentBtn) {
    el.openCurrentAttachmentBtn.classList.add("panel-hidden");
  }
}

function fillPositionForm(position) {
  resetPositionFormState();
  const setValue = (selector, value) => {
    const node = el.positionForm.querySelector(selector);
    if (node) {
      node.value = value;
    }
  };
  const setChecked = (selector, checked) => {
    const node = el.positionForm.querySelector(selector);
    if (node) {
      node.checked = Boolean(checked);
    }
  };
  setValue('input[name="positionNumber"]', position.positionNumber || "");
  setValue('input[name="width"]', String(Math.max(0, toInt(position.width))));
  setValue('input[name="height"]', String(Math.max(0, toInt(position.height))));
  setValue('select[name="technology"]', position.technology || "");
  setValue('input[name="line"]', position.line || "");
  setValue('input[name="positionFramesCount"]', String(Math.max(0, toInt(position.framesCount))));
  setValue('input[name="positionSashesCount"]', String(Math.max(0, toInt(position.sashesCount))));
  setChecked('input[name="shapeRect"]', position.shapeRect);
  setChecked('input[name="shapeSkos"]', position.shapeSkos);
  setChecked('input[name="shapeLuk"]', position.shapeLuk);
  setValue('input[name="slemieCount"]', String(Math.max(0, toInt(position.slemieCount))));
  setValue('input[name="slupekStalyCount"]', String(Math.max(0, toInt(position.slupekStalyCount))));
  setValue('input[name="przymykCount"]', String(Math.max(0, toInt(position.przymykCount))));
  setValue('input[name="niskiProgCount"]', String(Math.max(0, toInt(position.niskiProgCount))));
  setValue('input[name="machiningTime"]', String(Math.max(0, toFloat(position.times?.machining))));
  setValue('input[name="paintingTime"]', String(Math.max(0, toFloat(position.times?.painting))));
  setValue('input[name="assemblyTime"]', String(Math.max(0, toFloat(position.times?.assembly))));
  setValue('input[name="materialWoodDate"]', normalizeOptionalDate(position.materials?.wood?.date) || "");
  setValue('input[name="materialCorpusDate"]', normalizeOptionalDate(position.materials?.corpus?.date) || "");
  setValue('input[name="materialGlassDate"]', normalizeOptionalDate(position.materials?.glass?.date) || "");
  setValue('input[name="materialHardwareDate"]', normalizeOptionalDate(position.materials?.hardware?.date) || "");
  setValue('input[name="materialAccessoriesDate"]', normalizeOptionalDate(position.materials?.accessories?.date) || "");
  setChecked('input[name="materialWoodToOrder"]', Boolean(position.materials?.wood?.toOrder));
  setChecked('input[name="materialCorpusToOrder"]', Boolean(position.materials?.corpus?.toOrder));
  setChecked('input[name="materialGlassToOrder"]', Boolean(position.materials?.glass?.toOrder));
  setChecked('input[name="materialHardwareToOrder"]', Boolean(position.materials?.hardware?.toOrder));
  setChecked('input[name="materialAccessoriesToOrder"]', Boolean(position.materials?.accessories?.toOrder));
  setValue('textarea[name="notes"]', position.notes || "");
}

function openOrderDetails(orderId, options = {}) {
  if (!orderId) {
    return;
  }
  ui.selectedOrderId = String(orderId);
  ui.orderDetailsEditMode = false;
  if (options.switchToOrders) {
    setView("orders");
  }
  renderOrdersTable();
  renderSelectedOrderDetails();
  openModal("orderDetailsModal");
}

function startSelectedOrderEditMode() {
  if (!ui.selectedOrderId) {
    return;
  }
  ui.orderDetailsEditMode = true;
  renderSelectedOrderDetails();
  el.selectedOrderNumberInput?.focus();
}

function fillStatusSelectors() {
  fillSelect(el.orderStatusCreateSelect, ORDER_STATUSES);
  fillSelect(el.selectedOrderStatusSelect, ORDER_STATUSES);
  fillSelect(el.feedbackDepartmentSelect, POSITION_DEPARTMENT_STATUSES);
  fillSelect(el.feedbackOrderStatusFilter, ["Wszystkie", ...ORDER_STATUSES], "Wszystkie");
  fillSelect(
    el.feedbackOrderPositionStatusFilter,
    ["Wszystkie", ...KPI_POSITION_STATUSES, KPI_POSITION_MIXED_STATUS, KPI_POSITION_EMPTY_STATUS],
    "Wszystkie",
  );
  fillSelect(
    el.feedbackWorkflowStatusSelect,
    [{ value: "", label: "Bez zmiany" }, ...POSITION_WORKFLOW_STATUS_OPTIONS],
    "",
  );
  fillSelect(el.ordersStatusFilter, ["Wszystkie", ...ORDER_STATUSES], "Wszystkie");
  fillSelect(el.ganttStatusFilter, ["Wszystkie", ...ORDER_STATUSES], "Wszystkie");
}

function fillSelect(selectNode, options, currentValue) {
  if (!selectNode) {
    return;
  }
  const normalized = (options || []).map((item) => {
    if (item && typeof item === "object") {
      const value = String(item.value ?? "");
      const label = String(item.label ?? value);
      return { value, label };
    }
    const value = String(item ?? "");
    return { value, label: value };
  });
  selectNode.innerHTML = normalized
    .map((item) => `<option value="${escapeHtml(item.value)}">${escapeHtml(item.label)}</option>`)
    .join("");
  const desired = currentValue === undefined || currentValue === null ? "" : String(currentValue);
  const values = normalized.map((item) => item.value);
  if (values.includes(desired)) {
    selectNode.value = desired;
  }
}

async function safeAction(fn) {
  try {
    await fn();
  } catch (error) {
    console.error(error);
    alert(`Blad: ${error.message}`);
  }
}

async function restoreSessionUser() {
  try {
    const payload = await api("/api/auth/session");
    ui.currentUser = normalizeCurrentUser(payload.user || {});
    if (Array.isArray(payload.databases)) {
      state.databases = payload.databases;
    }
    state.activeDatabaseKey = String(payload.activeDatabase || state.activeDatabaseKey || "default");
  } catch (_error) {
    ui.currentUser = null;
  }
  syncLoginDatabaseSelect();
}

async function loadPublicDatabases() {
  try {
    const payload = await api("/api/public/databases");
    state.databases = Array.isArray(payload.databases) ? payload.databases : [];
    state.activeDatabaseKey = String(payload.activeDatabase || state.activeDatabaseKey || "default");
  } catch (_error) {
    if (!Array.isArray(state.databases) || state.databases.length === 0) {
      state.databases = [{ key: "default", name: "Domyslna baza", fileName: "planner.db", active: true, variant: false }];
      state.activeDatabaseKey = "default";
    }
  }
  syncLoginDatabaseSelect();
}

async function loginUser() {
  if (!el.loginForm?.reportValidity()) {
    return;
  }
  const login = String(el.loginInput?.value || "").trim();
  const password = String(el.loginPasswordInput?.value || "");
  const databaseKey = String(el.loginDatabaseSelect?.value || state.activeDatabaseKey || "default").trim() || "default";
  const payload = await api("/api/auth/login", {
    method: "POST",
    body: { login, password, databaseKey },
  });
  ui.currentUser = normalizeCurrentUser(payload.user || {});
  state.databases = Array.isArray(payload.databases) ? payload.databases : state.databases;
  state.activeDatabaseKey = String(payload.activeDatabase || databaseKey);
  if (el.loginDatabaseSelect) {
    el.loginDatabaseSelect.value = state.activeDatabaseKey;
  }
  await reloadAndRender();
  const firstSection = firstVisibleSectionForCurrentUser();
  setView(firstSection || "dashboard", true);
}

function resetLocalSessionState() {
  ui.currentUser = null;
  state.orders = [];
  state.feedbackEvents = [];
  state.users = [];
  state.skillWorkers = [];
  state.skillAvailability = {};
  state.stations = [];
  state.settings = normalizeSettings({});
  state.stationSettings = {};
  state.technologies = {};
  state.materialRules = {};
  state.databaseAccessMap = {};
  if (!Array.isArray(state.databases) || state.databases.length === 0) {
    state.databases = [{ key: "default", name: "Domyslna baza", fileName: "planner.db", active: true, variant: false }];
  }
  state.activeDatabaseKey =
    String(el.loginDatabaseSelect?.value || state.activeDatabaseKey || "default").trim() || "default";
  if (el.loginInput) {
    el.loginInput.value = "";
  }
  if (el.loginPasswordInput) {
    el.loginPasswordInput.value = "";
  }
  ui.selectedOrderId = "";
  ui.editingUserId = "";
  ui.orderDetailsEditMode = false;
  ui.pendingAttachmentPositionId = "";
  ui.editingSkillWorkerId = "";
  ui.kpiFilter = null;
  closeModal("orderDetailsModal");
  closeModal("positionModal");
  closeModal("orderModal");
  setView("dashboard", true);
  syncLoginDatabaseSelect();
}

async function logoutUser() {
  try {
    await api("/api/auth/logout", { method: "POST" });
  } catch (_error) {
    // Ignorujemy blad sieci - lokalnie i tak zamykamy sesje.
  }
  resetLocalSessionState();
  renderAll();
}

async function createOrder() {
  if (!el.orderForm.reportValidity()) {
    return;
  }
  const fd = new FormData(el.orderForm);
  await api("/api/orders", {
    method: "POST",
    body: {
      orderNumber: String(fd.get("orderNumber") || "").trim(),
      entryDate: String(fd.get("entryDate") || ""),
      orderStatus: String(fd.get("orderStatus") || "Dokumentacja"),
      owner: String(fd.get("owner") || "").trim(),
      client: String(fd.get("client") || "").trim(),
      color: String(fd.get("color") || "").trim(),
      framesCount: toInt(fd.get("framesCount")),
      sashesCount: toInt(fd.get("sashesCount")),
      extras: String(fd.get("extras") || "").trim(),
    },
  });
  el.orderForm.reset();
  hydrateDates();
  el.orderStatusCreateSelect.value = ORDER_STATUSES[0];
  closeModal("orderModal");
  await reloadAndRender();
}

async function importOrdersFromExcel() {
  const file = el.ordersImportFileInput?.files?.[0];
  if (!file) {
    throw new Error("Wybierz plik .xlsx do importu.");
  }
  const formData = new FormData();
  formData.append("file", file, file.name || "import.xlsx");
  const payload = await apiForm("/api/orders/import-excel", formData);
  const summary = payload.summary || {};
  const message = `Import zakonczony. Zamowienia +${toInt(summary.createdOrders)} / zaktualizowane ${toInt(
    summary.updatedOrders,
  )}, pozycje +${toInt(summary.createdPositions)} / zaktualizowane ${toInt(summary.updatedPositions)}, pominiete wiersze: ${toInt(
    summary.skippedRows,
  )}.`;
  if (el.ordersImportResult) {
    const errorCount = toInt(summary.errorCount);
    const firstError = Array.isArray(summary.errors) && summary.errors.length > 0 ? ` Pierwszy blad: ${summary.errors[0]}` : "";
    el.ordersImportResult.textContent = `${message}${errorCount > 0 ? ` Bledy: ${errorCount}.` : ""}${firstError}`;
  }
  alert(message + (toInt(summary.errorCount) > 0 ? `\nBledy: ${toInt(summary.errorCount)} (szczegoly pod formularzem).` : ""));
  if (el.ordersImportFileInput) {
    el.ordersImportFileInput.value = "";
  }
  await reloadAndRender();
}

async function createPosition() {
  if (!el.positionForm.reportValidity()) {
    return;
  }
  const orderId = ui.selectedOrderId;
  if (!orderId) {
    alert("Najpierw otworz szczegoly zamowienia.");
    return;
  }
  const fd = new FormData(el.positionForm);
  const attachmentFile = fd.get("attachment");
  const clearAttachment = toBoolean(fd.get("clearAttachment"));
  const hasAttachmentUpload = attachmentFile instanceof File && attachmentFile.size > 0;

  const payload = {
    positionNumber: String(fd.get("positionNumber") || "").trim(),
    width: toInt(fd.get("width")),
    height: toInt(fd.get("height")),
    framesCount: toInt(fd.get("positionFramesCount")),
    sashesCount: toInt(fd.get("positionSashesCount")),
    shapeRect: toBoolean(fd.get("shapeRect")),
    shapeSkos: toBoolean(fd.get("shapeSkos")),
    shapeLuk: toBoolean(fd.get("shapeLuk")),
    slemieCount: toInt(fd.get("slemieCount")),
    slupekStalyCount: toInt(fd.get("slupekStalyCount")),
    przymykCount: toInt(fd.get("przymykCount")),
    niskiProgCount: toInt(fd.get("niskiProgCount")),
    technology: String(fd.get("technology") || "").trim(),
    line: String(fd.get("line") || "").trim(),
    notes: String(fd.get("notes") || "").trim(),
    times: {
      machining: toFloat(fd.get("machiningTime")),
      painting: toFloat(fd.get("paintingTime")),
      assembly: toFloat(fd.get("assemblyTime")),
    },
    materials: {
      wood: {
        date: normalizeOptionalDate(fd.get("materialWoodDate")),
        toOrder: toBoolean(fd.get("materialWoodToOrder")),
      },
      corpus: {
        date: normalizeOptionalDate(fd.get("materialCorpusDate")),
        toOrder: toBoolean(fd.get("materialCorpusToOrder")),
      },
      glass: {
        date: normalizeOptionalDate(fd.get("materialGlassDate")),
        toOrder: toBoolean(fd.get("materialGlassToOrder")),
      },
      hardware: {
        date: normalizeOptionalDate(fd.get("materialHardwareDate")),
        toOrder: toBoolean(fd.get("materialHardwareToOrder")),
      },
      accessories: {
        date: normalizeOptionalDate(fd.get("materialAccessoriesDate")),
        toOrder: toBoolean(fd.get("materialAccessoriesToOrder")),
      },
    },
  };
  if (!payload.shapeRect && !payload.shapeSkos && !payload.shapeLuk) {
    payload.shapeRect = true;
  }
  let targetPositionId = "";

  if (ui.editingPositionId) {
    targetPositionId = String(ui.editingPositionId || "");
    await api(`/api/positions/${encodeURIComponent(targetPositionId)}`, {
      method: "PUT",
      body: payload,
    });
  } else {
    const created = await api(`/api/orders/${orderId}/positions`, {
      method: "POST",
      body: {
        ...payload,
        currentDepartmentStatus: "Dokumentacja",
      },
    });
    targetPositionId = String(created.id || "");
  }

  if (targetPositionId) {
    if (hasAttachmentUpload) {
      const formData = new FormData();
      formData.append("file", attachmentFile, attachmentFile.name || "zalacznik.bin");
      await apiForm(`/api/positions/${encodeURIComponent(targetPositionId)}/attachment`, formData);
    } else if (ui.editingPositionId && clearAttachment) {
      await api(`/api/positions/${encodeURIComponent(targetPositionId)}/attachment`, { method: "DELETE" });
    }
  }

  ui.editingPositionId = "";
  resetPositionFormState();
  closeModal("positionModal");
  await reloadAndRender();
}

function collectSkillLevelsFromForm() {
  const output = {};
  const inputs = Array.from(el.skillWorkerSkillsWrap?.querySelectorAll("input[data-skill-level='1']") || []);
  inputs.forEach((input) => {
    const stationId = String(input.dataset.stationId || "").trim();
    if (!stationId) {
      return;
    }
    const level = clamp(toInt(input.value), 0, 3);
    input.value = String(level);
    if (level > 0) {
      output[stationId] = level;
    }
  });
  return output;
}

function availableSkillDepartments() {
  return DEPARTMENTS.slice();
}

async function saveSkillWorker() {
  if (!isLoggedIn()) {
    throw new Error("Zaloguj sie, aby zapisac pracownika.");
  }
  if (!el.skillWorkerForm?.reportValidity()) {
    return;
  }
  const departments = availableSkillDepartments();
  const payload = {
    name: String(el.skillWorkerNameInput?.value || "").trim(),
    department: String(el.skillWorkerDepartmentSelect?.value || departments[0] || DEPARTMENTS[0]),
    primaryStationId: String(el.skillWorkerPrimaryStationSelect?.value || "").trim(),
    assignedShift: clamp(toInt(el.skillWorkerAssignedShiftSelect?.value || 1), 1, 3),
    active: Boolean(el.skillWorkerActiveInput?.checked),
    skills: collectSkillLevelsFromForm(),
  };
  const editingId = String(ui.editingSkillWorkerId || "").trim();
  if (editingId) {
    await api(`/api/skills/workers/${encodeURIComponent(editingId)}`, {
      method: "PUT",
      body: payload,
    });
  } else {
    await api("/api/skills/workers", {
      method: "POST",
      body: payload,
    });
  }
  resetSkillWorkerFormState();
  await reloadAndRender();
}

function getSkillWorkerSelectionSet() {
  const validIds = new Set((state.skillWorkers || []).map((worker) => String(worker.id || "").trim()).filter(Boolean));
  const selected = new Set((ui.selectedSkillWorkerIds || []).map((id) => String(id || "").trim()).filter(Boolean));
  Array.from(selected).forEach((id) => {
    if (!validIds.has(id)) {
      selected.delete(id);
    }
  });
  return selected;
}

function setSkillWorkerSelectionSet(selection) {
  ui.selectedSkillWorkerIds = Array.from(selection || []).map((id) => String(id || "").trim()).filter(Boolean);
}

function updateSkillWorkerSelectionStatus() {
  if (!el.skillWorkersBulkStatus) {
    return;
  }
  const count = getSkillWorkerSelectionSet().size;
  el.skillWorkersBulkStatus.textContent = count > 0 ? `Zaznaczono pracownikow: ${count}` : "Brak zaznaczonych pracownikow.";
}

function onSkillWorkersListChange(event) {
  const checkbox = event.target.closest('input[data-action="select-skill-worker"]');
  if (!checkbox) {
    return;
  }
  const workerId = String(checkbox.dataset.workerId || "").trim();
  if (!workerId) {
    return;
  }
  const selected = getSkillWorkerSelectionSet();
  if (checkbox.checked) {
    selected.add(workerId);
  } else {
    selected.delete(workerId);
  }
  setSkillWorkerSelectionSet(selected);
  updateSkillWorkerSelectionStatus();
}

function selectAllSkillWorkers() {
  const selected = new Set((state.skillWorkers || []).map((worker) => String(worker.id || "").trim()).filter(Boolean));
  setSkillWorkerSelectionSet(selected);
  renderSkillWorkersList();
}

function clearSkillWorkerSelection() {
  setSkillWorkerSelectionSet(new Set());
  renderSkillWorkersList();
}

async function applyBulkShiftForSkillWorkers() {
  const selected = getSkillWorkerSelectionSet();
  if (selected.size === 0) {
    throw new Error("Zaznacz pracownikow do przypisania zmiany.");
  }
  const assignedShift = clamp(toInt(el.skillWorkersBulkShiftSelect?.value || 1), 1, 3);
  await api("/api/skills/workers/bulk-shift", {
    method: "PUT",
    body: {
      workerIds: Array.from(selected),
      assignedShift,
    },
  });
  if (el.skillWorkersBulkStatus) {
    el.skillWorkersBulkStatus.textContent = `Przypisano zmiane ${assignedShift} dla ${selected.size} pracownikow.`;
  }
  await reloadAndRender();
}

function onSkillWorkersListClick(event) {
  const editBtn = event.target.closest("button[data-action='edit-skill-worker']");
  if (editBtn) {
    startSkillWorkerEdit(editBtn.dataset.workerId);
    return;
  }
  const deleteBtn = event.target.closest("button[data-action='delete-skill-worker']");
  if (!deleteBtn) {
    return;
  }
  const workerId = String(deleteBtn.dataset.workerId || "").trim();
  if (!workerId) {
    return;
  }
  const worker = state.skillWorkers.find((item) => item.id === workerId);
  const label = worker?.name || "pracownika";
  if (!window.confirm(`Czy na pewno usunac ${label}?`)) {
    return;
  }
  safeAction(() => deleteSkillWorker(workerId));
}

function startSkillWorkerEdit(workerId) {
  const targetId = String(workerId || "").trim();
  const worker = state.skillWorkers.find((item) => item.id === targetId);
  if (!worker) {
    return;
  }
  ui.editingSkillWorkerId = targetId;
  if (el.skillWorkerDepartmentSelect) {
    const departments = availableSkillDepartments();
    const validDepartment = departments.includes(worker.department) ? worker.department : departments[0] || DEPARTMENTS[0];
    el.skillWorkerDepartmentSelect.value = validDepartment;
  }
  if (el.skillWorkerAssignedShiftSelect) {
    el.skillWorkerAssignedShiftSelect.value = String(clamp(toInt(worker.assignedShift || 1), 1, 3));
  }
  renderSkillWorkerForm();
  el.skillWorkerNameInput?.focus();
}

function cancelSkillWorkerEdit() {
  resetSkillWorkerFormState();
}

function resetSkillWorkerFormState() {
  ui.editingSkillWorkerId = "";
  el.skillWorkerForm?.reset();
  if (el.skillWorkerDepartmentSelect) {
    const departments = availableSkillDepartments();
    el.skillWorkerDepartmentSelect.value = departments.includes(el.skillWorkerDepartmentSelect.value)
      ? el.skillWorkerDepartmentSelect.value
      : departments[0] || DEPARTMENTS[0];
  }
  if (el.skillWorkerActiveInput) {
    el.skillWorkerActiveInput.checked = true;
  }
  if (el.skillWorkerAssignedShiftSelect) {
    el.skillWorkerAssignedShiftSelect.value = "1";
  }
  Array.from(el.skillWorkerSkillsWrap?.querySelectorAll("input[data-skill-level='1']") || []).forEach((input) => {
    input.value = "0";
  });
  if (el.skillWorkerSubmitBtn) {
    el.skillWorkerSubmitBtn.textContent = "Dodaj pracownika";
  }
  el.cancelSkillWorkerEditBtn?.classList.add("panel-hidden");
}

async function deleteSkillWorker(workerId) {
  const targetId = String(workerId || "").trim();
  if (!targetId) {
    return;
  }
  const selected = getSkillWorkerSelectionSet();
  selected.delete(targetId);
  setSkillWorkerSelectionSet(selected);
  await api(`/api/skills/workers/${encodeURIComponent(targetId)}`, { method: "DELETE" });
  if (ui.editingSkillWorkerId === targetId) {
    resetSkillWorkerFormState();
  }
  await reloadAndRender();
}

async function createUser() {
  if (!isAdminUser()) {
    throw new Error("Tylko administrator moze zarzadzac uzytkownikami.");
  }
  if (!el.userForm.reportValidity()) {
    return;
  }
  const fd = new FormData(el.userForm);
  const role = String(fd.get("role") || "user").trim().toLowerCase() === "admin" ? "admin" : "user";
  const visibleSections = role === "admin" ? USER_VISIBLE_SECTIONS.slice() : normalizeVisibleSections(fd.getAll("visibleSection"));
  if (role !== "admin" && visibleSections.length === 0) {
    throw new Error("Wybierz przynajmniej jedna sekcje widoczna dla uzytkownika.");
  }
  const payload = {
    name: String(fd.get("name") || "").trim(),
    login: String(fd.get("login") || "").trim(),
    password: String(fd.get("password") || ""),
    department: String(fd.get("department") || DEPARTMENTS[0]),
    role,
    canCreateDatabases: role === "admin" ? true : toBoolean(fd.get("canCreateDatabases")),
    visibleSections,
  };
  if (ui.editingUserId) {
    await api(`/api/users/${encodeURIComponent(ui.editingUserId)}`, {
      method: "PUT",
      body: payload,
    });
  } else {
    await api("/api/users", {
      method: "POST",
      body: payload,
    });
  }
  resetUserFormState();
  await reloadAndRender();
}

function onUsersListClick(event) {
  const editBtn = event.target.closest("button[data-action='edit-user']");
  if (editBtn) {
    startUserEdit(editBtn.dataset.userId);
    return;
  }
  const deleteBtn = event.target.closest("button[data-action='delete-user']");
  if (!deleteBtn) {
    return;
  }
  const userId = String(deleteBtn.dataset.userId || "");
  if (!userId) {
    return;
  }
  const user = state.users.find((item) => item.id === userId);
  const label = user?.name || user?.login || "uzytkownika";
  if (!window.confirm(`Czy na pewno usunac ${label}?`)) {
    return;
  }
  safeAction(() => deleteUser(userId));
}

function startUserEdit(userId) {
  if (!isAdminUser()) {
    return;
  }
  const user = state.users.find((item) => item.id === String(userId || ""));
  if (!user) {
    return;
  }
  ui.editingUserId = user.id;
  const nameInput = el.userForm?.querySelector('input[name="name"]');
  const loginInput = el.userForm?.querySelector('input[name="login"]');
  const passwordInput = el.userForm?.querySelector('input[name="password"]');
  const departmentSelect = el.userForm?.querySelector('select[name="department"]');
  const roleSelect = el.userForm?.querySelector('select[name="role"]');
  const canCreateDbInput = el.userForm?.querySelector('input[name="canCreateDatabases"]');

  if (nameInput) {
    nameInput.value = user.name || "";
  }
  if (loginInput) {
    loginInput.value = user.login || "";
  }
  if (passwordInput) {
    passwordInput.value = "";
  }
  if (departmentSelect) {
    departmentSelect.value = DEPARTMENTS.includes(String(user.department || "").trim())
      ? String(user.department || "").trim()
      : DEPARTMENTS[0];
  }
  if (roleSelect) {
    roleSelect.value = String(user.role || "user").toLowerCase() === "admin" ? "admin" : "user";
  }
  if (canCreateDbInput) {
    canCreateDbInput.checked = Boolean(user.canCreateDatabases);
  }
  setUserPermissionsSelection(user.visibleSections || []);
  syncUserPermissionsEnabledState();
  syncUserFormMode();
  nameInput?.focus();
}

function cancelUserEdit() {
  resetUserFormState();
  renderUsers();
}

async function deleteUser(userId) {
  if (!isAdminUser()) {
    throw new Error("Tylko administrator moze zarzadzac uzytkownikami.");
  }
  await api(`/api/users/${encodeURIComponent(userId)}`, { method: "DELETE" });
  if (ui.editingUserId === userId) {
    resetUserFormState();
  }
  await reloadAndRender();
}

function resetUserFormState() {
  ui.editingUserId = "";
  el.userForm?.reset();
  const roleNode = el.userForm?.querySelector('select[name="role"]');
  const canCreateDbInput = el.userForm?.querySelector('input[name="canCreateDatabases"]');
  if (roleNode) {
    roleNode.value = "user";
  }
  if (canCreateDbInput) {
    canCreateDbInput.checked = false;
  }
  setUserPermissionsSelection(["orders", "gantt", "reports", "execution", "feedback"]);
  syncUserPermissionsEnabledState();
  syncUserFormMode();
}

function syncUserFormMode() {
  const editing = Boolean(ui.editingUserId);
  const passwordInput = el.userForm?.querySelector('input[name="password"]');
  if (el.userSubmitBtn) {
    el.userSubmitBtn.textContent = editing ? "Zapisz zmiany" : "Dodaj uzytkownika";
  }
  if (el.cancelUserEditBtn) {
    el.cancelUserEditBtn.classList.toggle("panel-hidden", !editing);
  }
  if (passwordInput) {
    passwordInput.required = !editing;
    passwordInput.placeholder = editing ? "Zostaw puste, aby nie zmieniac hasla" : "";
  }
}

async function saveSettings() {
  if (!el.settingsForm.reportValidity()) {
    return;
  }
  const fd = new FormData(el.settingsForm);
  const weekdayShifts = {};
  Array.from(el.settingsForm.querySelectorAll('input[name="weekdayShift"]')).forEach((input) => {
    const day = clamp(toInt(input.dataset.day), 0, 6);
    weekdayShifts[String(day)] = clamp(toInt(input.value), 0, 3);
  });
  const workingDays = Object.entries(weekdayShifts)
    .filter(([, shifts]) => toInt(shifts) > 0)
    .map(([day]) => toInt(day))
    .sort((a, b) => a - b);
  if (workingDays.length === 0) {
    throw new Error("Przynajmniej jeden dzien musi miec minimum 1 zmiane.");
  }
  await api("/api/settings", {
    method: "PUT",
    body: { minutesPerShift: toInt(fd.get("minutesPerShift")), workingDays, weekdayShifts },
  });
  await reloadAndRender();
}

async function createDatabaseVariant() {
  if (!canCreateDatabaseVariants()) {
    throw new Error("Brak uprawnien do tworzenia nowych baz.");
  }
  const name = String(el.newDatabaseNameInput?.value || "").trim();
  if (!name) {
    throw new Error("Podaj nazwe nowej bazy.");
  }
  const payload = await api("/api/databases", {
    method: "POST",
    body: { name, activate: false },
  });
  state.databases = Array.isArray(payload.databases) ? payload.databases : [];
  state.activeDatabaseKey = String(payload.activeDatabase || state.activeDatabaseKey || "default");
  if (el.newDatabaseNameInput) {
    el.newDatabaseNameInput.value = "";
  }
  if (el.databaseStatusText) {
    el.databaseStatusText.textContent = "Nowa baza zostala utworzona.";
  }
  syncLoginDatabaseSelect();
  renderDatabaseManager();
}

async function switchActiveDatabase() {
  if (!isLoggedIn()) {
    throw new Error("Zaloguj sie, aby przelaczyc baze.");
  }
  const key = String(el.databaseSelect?.value || "").trim();
  if (!key) {
    throw new Error("Wybierz baze do przelaczenia.");
  }
  const payload = await api("/api/databases/active", {
    method: "PUT",
    body: { key },
  });
  state.databases = Array.isArray(payload.databases) ? payload.databases : [];
  state.activeDatabaseKey = String(payload.activeDatabase || key);
  syncLoginDatabaseSelect();
  if (el.databaseStatusText) {
    el.databaseStatusText.textContent = "Baza zostala przelaczona.";
  }
  if (payload.requiresRelogin) {
    resetLocalSessionState();
    renderAll();
    alert("Przelaczono baze. Ta sesja nie istnieje w nowej bazie - zaloguj sie ponownie.");
    return;
  }
  await reloadAndRender();
}

async function saveCalendarDayWorking(date, working) {
  await api("/api/settings/calendar-day", {
    method: "PUT",
    body: { date, working: Boolean(working) },
  });
  await reloadAndRender();
}

async function upsertStationOvertime() {
  const date = normalizeOptionalDate(el.overtimeDateInput?.value);
  const stationId = String(el.overtimeStationSelect?.value || "").trim();
  const minutes = clamp(toInt(el.overtimeMinutesInput?.value), 0, 1440);
  if (!date) {
    throw new Error("Wybierz date nadgodzin.");
  }
  if (!stationId) {
    throw new Error("Wybierz stanowisko dla nadgodzin.");
  }
  if (minutes <= 0) {
    throw new Error("Minuty nadgodzin musza byc wieksze od 0.");
  }
  await api("/api/settings/station-overtime", {
    method: "PUT",
    body: { date, stationId, minutes },
  });
  await reloadAndRender();
}

async function removeStationOvertime(date, stationId) {
  const normalizedDate = normalizeOptionalDate(date);
  const normalizedStation = String(stationId || "").trim();
  if (!normalizedDate || !normalizedStation) {
    throw new Error("Brak danych wpisu nadgodzin do usuniecia.");
  }
  await api("/api/settings/station-overtime", {
    method: "DELETE",
    body: { date: normalizedDate, stationId: normalizedStation },
  });
  await reloadAndRender();
}

function onStationOvertimeTableClick(event) {
  const removeBtn = event.target.closest("button[data-overtime-remove='1']");
  if (!removeBtn) {
    return;
  }
  const date = String(removeBtn.dataset.date || "");
  const stationId = String(removeBtn.dataset.stationId || "");
  safeAction(() => removeStationOvertime(date, stationId));
}

async function updateOrderManualStart(orderId, targetDate) {
  const order = state.orders.find((item) => item.id === orderId);
  if (!order) {
    throw new Error("Nie znaleziono zamowienia do przesuniecia.");
  }
  let manualStartDate = null;
  if (targetDate) {
    const normalized = normalizeOptionalDate(targetDate);
    if (!normalized) {
      throw new Error("Nieprawidlowa data upuszczenia.");
    }
    const snapped = isoDate(normalizeWorkday(normalized));
    manualStartDate = maxDate([order.entryDate, snapped]);
  }
  await api(`/api/orders/${encodeURIComponent(orderId)}/manual-start`, {
    method: "PUT",
    body: { manualStartDate },
  });
  await reloadAndRender();
}

async function saveStationSettings() {
  const rows = Array.from(el.stationSettingsTableBody.querySelectorAll("tr[data-station-row='1']"));
  if (rows.length === 0) {
    throw new Error("Lista stanowisk nie moze byc pusta. Dodaj przynajmniej jedno stanowisko.");
  }
  const stations = rows.map((row, index) => {
    const stationIdInput = row.querySelector('input[data-kind="stationId"]');
    const nameInput = row.querySelector('input[data-kind="name"]');
    const departmentSelect = row.querySelector('select[data-kind="department"]');
    const activeInput = row.querySelector('input[data-kind="active"]');
    const shiftInput = row.querySelector('input[data-kind="shift"]');
    const peopleInput = row.querySelector('input[data-kind="people"]');
    const name = String(nameInput?.value || "").trim();
    if (!name) {
      throw new Error("Kazde stanowisko musi miec nazwe.");
    }
    return {
      id: String(stationIdInput?.value || "").trim(),
      name,
      department: String(departmentSelect?.value || "Maszynownia"),
      active: Boolean(activeInput?.checked),
      shiftCount: clamp(toInt(shiftInput?.value), 1, 3),
      peopleCount: clamp(toInt(peopleInput?.value), 1, 200),
      sortOrder: index + 1,
    };
  });
  await api("/api/stations-config", {
    method: "PUT",
    body: { stations },
  });
  await reloadAndRender();
}

async function saveMaterialRules() {
  const rules = {};
  DEPARTMENTS.forEach((dept) => {
    rules[dept] = [];
  });
  Array.from(el.materialRulesTableBody.querySelectorAll("input[type='checkbox']")).forEach((input) => {
    if (!input.checked) {
      return;
    }
    const material = input.dataset.material;
    const department = input.dataset.department;
    rules[department].push(material);
  });
  await api("/api/material-rules", {
    method: "PUT",
    body: { materialRules: rules },
  });
  await reloadAndRender();
}

async function saveTechnologyAllocations() {
  const technology = el.technologyEditorSelect.value;
  const process = el.processEditorSelect.value;
  if (!technology || !process) {
    return;
  }
  const inputs = Array.from(el.technologyAllocationTableBody.querySelectorAll("input[data-station-id]"));
  const percentages = {};
  inputs.forEach((input) => {
    percentages[input.dataset.stationId] = Math.max(0, toFloat(input.value));
  });
  await api(`/api/technologies/${encodeURIComponent(technology)}/process/${encodeURIComponent(process)}`, {
    method: "PUT",
    body: { percentages },
  });
  await reloadAndRender();
}

async function addTechnology() {
  const name = window.prompt("Podaj nazwe nowej technologii.");
  if (!name || !name.trim()) {
    return;
  }
  await api("/api/technologies", {
    method: "POST",
    body: { name: name.trim() },
  });
  ui.selectedTechnology = name.trim();
  await reloadAndRender();
}

async function saveSelectedOrder() {
  if (!ui.selectedOrderId) {
    return;
  }
  if (!ui.orderDetailsEditMode) {
    alert('Kliknij "Edycja", aby zmienic dane zamowienia.');
    return;
  }
  if (!el.selectedOrderForm.reportValidity()) {
    return;
  }
  const order = state.orders.find((item) => item.id === ui.selectedOrderId);
  const totals = sumOrderFramesAndSashes(order || { positions: [] });
  const editedPlannedDate = normalizeOptionalDate(el.selectedOrderPlannedDateInput?.value);
  const autoPlannedDate = normalizeOptionalDate(order?.calculation?.calculatedPlannedDate);
  const manualPlannedDate = editedPlannedDate && editedPlannedDate !== autoPlannedDate ? editedPlannedDate : null;
  await api(`/api/orders/${ui.selectedOrderId}`, {
    method: "PUT",
    body: {
      orderNumber: String(el.selectedOrderNumberInput.value || "").trim(),
      entryDate: String(el.selectedOrderEntryDateInput.value || ""),
      orderStatus: String(el.selectedOrderStatusSelect.value || "Dokumentacja"),
      owner: String(el.selectedOrderOwnerInput.value || "").trim(),
      client: String(el.selectedOrderClientInput.value || "").trim(),
      color: String(el.selectedOrderColorInput.value || "").trim(),
      framesCount: totals.framesCount,
      sashesCount: totals.sashesCount,
      extras: String(el.selectedOrderExtrasInput.value || "").trim(),
      manualPlannedDate,
    },
  });
  ui.orderDetailsEditMode = false;
  await reloadAndRender();
}

async function openCurrentEditingAttachment() {
  if (!ui.selectedOrderId || !ui.editingPositionId) {
    return;
  }
  const order = state.orders.find((item) => item.id === ui.selectedOrderId);
  const position = order?.positions?.find((item) => item.id === ui.editingPositionId);
  if (!position?.attachmentName) {
    throw new Error("Ta pozycja nie ma zalacznika.");
  }
  await openPositionAttachment(position.id, position.attachmentName);
}

async function chooseAttachmentForPosition(positionId) {
  const id = String(positionId || "").trim();
  if (!id) {
    return;
  }
  if (!el.positionAttachmentUploadInput) {
    throw new Error("Brak kontrolki wyboru pliku.");
  }
  ui.pendingAttachmentPositionId = id;
  el.positionAttachmentUploadInput.value = "";
  el.positionAttachmentUploadInput.click();
}

async function uploadAttachmentFromPicker() {
  if (!el.positionAttachmentUploadInput) {
    return;
  }
  const positionId = String(ui.pendingAttachmentPositionId || "").trim();
  const file = el.positionAttachmentUploadInput.files?.[0];
  ui.pendingAttachmentPositionId = "";
  if (!positionId || !file) {
    return;
  }
  const formData = new FormData();
  formData.append("file", file, file.name || "zalacznik.bin");
  await apiForm(`/api/positions/${encodeURIComponent(positionId)}/attachment`, formData);
  await reloadAndRender();
}

async function deletePositionAttachment(positionId) {
  const id = String(positionId || "").trim();
  if (!id) {
    return;
  }
  if (!window.confirm("Usunac zalacznik z tej pozycji?")) {
    return;
  }
  await api(`/api/positions/${encodeURIComponent(id)}/attachment`, { method: "DELETE" });
  await reloadAndRender();
}

async function openPositionAttachment(positionId, fileName) {
  const response = await fetch(`/api/positions/${encodeURIComponent(positionId)}/attachment`, {
    method: "GET",
    credentials: "same-origin",
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    if (response.status === 401 && isLoggedIn()) {
      resetLocalSessionState();
      renderAll();
    }
    throw new Error(payload.error || "Nie mozna otworzyc zalacznika.");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const tab = window.open(url, "_blank");
  if (!tab) {
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName || "zalacznik";
    link.click();
  }
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

async function updateSelectedPositions(action) {
  const orderId = el.feedbackOrderSelect.value;
  if (!orderId) {
    return;
  }
  const positionIds = getSelectedFeedbackPositionIds();
  if (positionIds.length === 0) {
    return;
  }
  await api("/api/positions/bulk-status", {
    method: "POST",
    body: {
      orderId,
      positionIds,
      action,
      actor: "Brygadzista",
    },
  });
  ui.feedbackSelectionByOrder[orderId] = [];
  await reloadAndRender();
}

async function setDepartmentForSelectedPositions() {
  const orderId = el.feedbackOrderSelect.value;
  if (!orderId) {
    return;
  }
  const positionIds = getSelectedFeedbackPositionIds();
  if (positionIds.length === 0) {
    return;
  }
  await api("/api/positions/bulk-department", {
    method: "PUT",
    body: {
      positionIds,
      currentDepartmentStatus: el.feedbackDepartmentSelect.value,
    },
  });
  ui.feedbackSelectionByOrder[orderId] = [];
  await reloadAndRender();
}

async function applyFeedbackChangesForSelected() {
  const orderId = String(el.feedbackOrderSelect?.value || "");
  if (!orderId) {
    return;
  }
  const positionIds = getSelectedFeedbackPositionIds();
  if (positionIds.length === 0) {
    alert("Zaznacz pozycje do aktualizacji.");
    return;
  }
  const workflowStatus = normalizeWorkflowStatusValue(el.feedbackWorkflowStatusSelect?.value, "");
  const progressPercent = parseFeedbackProgressPercent(el.feedbackProgressPercentInput?.value);
  if (!workflowStatus && progressPercent === null) {
    alert("Wybierz status realizacji lub wpisz % realizacji.");
    return;
  }
  await api("/api/positions/bulk-feedback", {
    method: "PUT",
    body: {
      positionIds,
      workflowStatus: workflowStatus || null,
      progressPercent,
      actor: "Brygadzista",
    },
  });
  ui.feedbackSelectionByOrder[orderId] = [];
  await reloadAndRender();
}

async function saveSingleFeedbackPosition(positionId) {
  const id = String(positionId || "").trim();
  if (!id) {
    return;
  }
  const row = el.feedbackPositions?.querySelector(`[data-position-id="${cssEscape(id)}"]`);
  if (!row) {
    return;
  }
  const department = String(row.querySelector('[data-feedback-field="department"]')?.value || "Dokumentacja");
  const workflowStatus = normalizeWorkflowStatusValue(row.querySelector('[data-feedback-field="workflow"]')?.value, "pending");
  const progressPercent = parseFeedbackProgressPercent(row.querySelector('[data-feedback-field="progress"]')?.value);
  await api(`/api/positions/${encodeURIComponent(id)}/feedback`, {
    method: "PUT",
    body: {
      currentDepartmentStatus: department,
      workflowStatus,
      progressPercent,
      actor: "Brygadzista",
    },
  });
  await reloadAndRender();
}

function getSelectedFeedbackPositionIds() {
  const orderId = String(el.feedbackOrderSelect?.value || "");
  if (!orderId) {
    return [];
  }
  return Array.from(getFeedbackSelectionSet(orderId));
}

function selectAllFeedbackPositions() {
  const orderId = String(el.feedbackOrderSelect?.value || "");
  if (!orderId) {
    return;
  }
  const selected = getFeedbackSelectionSet(orderId);
  Array.from(el.feedbackPositions.querySelectorAll('input[data-feedback-position-id]')).forEach((input) => {
    selected.add(String(input.value));
  });
  setFeedbackSelectionSet(orderId, selected);
  renderFeedbackPositions();
}

function clearFeedbackSelection() {
  const orderId = String(el.feedbackOrderSelect?.value || "");
  if (!orderId) {
    return;
  }
  setFeedbackSelectionSet(orderId, new Set());
  renderFeedbackPositions();
}

function onOrdersHeaderClick(event) {
  const header = event.target.closest("th[data-order-sort]");
  if (!header) {
    return;
  }
  const sortKey = String(header.dataset.orderSort || "").trim();
  if (!sortKey) {
    return;
  }
  if (ui.ordersSortKey === sortKey) {
    ui.ordersSortDir = ui.ordersSortDir === "asc" ? "desc" : "asc";
  } else {
    ui.ordersSortKey = sortKey;
    ui.ordersSortDir = sortKey === "blockedMaterials" || sortKey === "positionsCount" ? "desc" : "asc";
  }
  renderOrdersTable();
}

function onFeedbackPositionRowClick(event) {
  const quickStatusBtn = event.target.closest("button[data-action='quick-feedback-status']");
  if (quickStatusBtn) {
    const row = event.target.closest("[data-position-id]");
    if (!row) {
      return;
    }
    const workflow = normalizeWorkflowStatusValue(quickStatusBtn.dataset.value, "pending");
    applyFeedbackRowWorkflow(row, workflow);
    refreshFeedbackQuickButtons(row);
    return;
  }

  const quickProgressBtn = event.target.closest("button[data-action='quick-feedback-progress']");
  if (quickProgressBtn) {
    const row = event.target.closest("[data-position-id]");
    if (!row) {
      return;
    }
    const progress = clamp(toFloat(quickProgressBtn.dataset.value), 0, 100);
    applyFeedbackRowProgress(row, progress);
    refreshFeedbackQuickButtons(row);
    return;
  }

  const quickProgressStepBtn = event.target.closest("button[data-action='quick-feedback-progress-step']");
  if (quickProgressStepBtn) {
    const row = event.target.closest("[data-position-id]");
    if (!row) {
      return;
    }
    const progressInput = row.querySelector('[data-feedback-field="progress"]');
    const current = parseFeedbackProgressPercent(progressInput?.value);
    const step = toFloat(quickProgressStepBtn.dataset.step || 0);
    const next = clamp((current ?? 0) + step, 0, 100);
    applyFeedbackRowProgress(row, next);
    refreshFeedbackQuickButtons(row);
    return;
  }

  const saveBtn = event.target.closest("button[data-action='save-feedback-position']");
  if (saveBtn) {
    const positionId = String(saveBtn.dataset.positionId || "");
    if (!positionId) {
      return;
    }
    safeAction(() => saveSingleFeedbackPosition(positionId));
    return;
  }

  const row = event.target.closest("[data-feedback-position-row='1']");
  if (!row) {
    return;
  }
  if (event.target.closest("input, button, a, select, textarea")) {
    return;
  }
  const checkbox = row.querySelector("input[data-feedback-position-id]");
  if (!checkbox) {
    return;
  }
  checkbox.checked = !checkbox.checked;
  checkbox.dispatchEvent(new Event("change", { bubbles: true }));
}

function onFeedbackPositionCheckboxChange(event) {
  const checkbox = event.target.closest("input[data-feedback-position-id]");
  if (!checkbox) {
    const field = event.target.closest("[data-feedback-field]");
    if (!field) {
      return;
    }
    const row = event.target.closest("[data-position-id]");
    if (!row) {
      return;
    }
    const fieldName = String(field.getAttribute("data-feedback-field") || "");
    if (fieldName === "workflow") {
      applyFeedbackRowWorkflow(row, field.value);
    } else if (fieldName === "progress") {
      const value = parseFeedbackProgressPercent(field.value);
      if (value !== null) {
        applyFeedbackRowProgress(row, value);
      }
    }
    refreshFeedbackQuickButtons(row);
    return;
  }
  const orderId = String(el.feedbackOrderSelect?.value || "");
  if (!orderId) {
    return;
  }
  const selected = getFeedbackSelectionSet(orderId);
  const positionId = String(checkbox.value || "");
  if (!positionId) {
    return;
  }
  if (checkbox.checked) {
    selected.add(positionId);
  } else {
    selected.delete(positionId);
  }
  setFeedbackSelectionSet(orderId, selected);
  applyFeedbackActionState();
}

function onFeedbackPositionInput(event) {
  const field = event.target.closest('[data-feedback-field="progress"]');
  if (!field) {
    return;
  }
  const row = event.target.closest("[data-position-id]");
  if (!row) {
    return;
  }
  const value = parseFeedbackProgressPercent(field.value);
  if (value !== null) {
    applyFeedbackRowProgress(row, value);
  }
  refreshFeedbackQuickButtons(row);
}

function applyFeedbackRowWorkflow(row, workflow) {
  const workflowSelect = row.querySelector('[data-feedback-field="workflow"]');
  const progressInput = row.querySelector('[data-feedback-field="progress"]');
  if (workflowSelect) {
    workflowSelect.value = normalizeWorkflowStatusValue(workflow, "pending");
  }
  const normalized = normalizeWorkflowStatusValue(workflowSelect?.value, "pending");
  const current = parseFeedbackProgressPercent(progressInput?.value);
  if (normalized === "done") {
    if (progressInput) {
      progressInput.value = "100";
    }
    return;
  }
  if (normalized === "in_progress") {
    if (progressInput && (current === null || current <= 0)) {
      progressInput.value = "1";
    }
    return;
  }
  if (progressInput) {
    progressInput.value = "0";
  }
}

function applyFeedbackRowProgress(row, progress) {
  const workflowSelect = row.querySelector('[data-feedback-field="workflow"]');
  const progressInput = row.querySelector('[data-feedback-field="progress"]');
  const bounded = clamp(progress, 0, 100);
  if (progressInput) {
    progressInput.value = String(Math.round(bounded));
  }
  if (!workflowSelect) {
    return;
  }
  if (bounded >= 100) {
    workflowSelect.value = "done";
    return;
  }
  if (bounded <= 0) {
    workflowSelect.value = "pending";
    return;
  }
  if (String(workflowSelect.value || "") === "pending") {
    workflowSelect.value = "in_progress";
  }
}

function refreshFeedbackQuickButtons(row) {
  const workflowSelect = row.querySelector('[data-feedback-field="workflow"]');
  const progressInput = row.querySelector('[data-feedback-field="progress"]');
  const workflow = normalizeWorkflowStatusValue(workflowSelect?.value, "pending");
  const progress = parseFeedbackProgressPercent(progressInput?.value) ?? 0;

  row.querySelectorAll("button[data-action='quick-feedback-status']").forEach((btn) => {
    const active = String(btn.dataset.value || "") === workflow;
    btn.classList.toggle("is-active", active);
    btn.setAttribute("aria-pressed", active ? "true" : "false");
  });
  row.querySelectorAll("button[data-action='quick-feedback-progress']").forEach((btn) => {
    const target = toFloat(btn.dataset.value || 0);
    const active = Math.abs(target - progress) < 0.1;
    btn.classList.toggle("is-active", active);
    btn.setAttribute("aria-pressed", active ? "true" : "false");
  });
}

function getFeedbackSelectionSet(orderId) {
  return new Set(ui.feedbackSelectionByOrder?.[orderId] || []);
}

function setFeedbackSelectionSet(orderId, selectedSet) {
  ui.feedbackSelectionByOrder[orderId] = Array.from(selectedSet);
}

function onOrderTableClick(event) {
  const detailsBtn = event.target.closest("button[data-action='open-order-details']");
  if (!detailsBtn) {
    return;
  }
  openOrderDetails(detailsBtn.dataset.orderId);
}

function onOrderTableChange(event) {
  const checkbox = event.target.closest('input[data-action="toggle-archive-order"]');
  if (!checkbox) {
    return;
  }
  const orderId = String(checkbox.dataset.orderId || "");
  if (!orderId) {
    return;
  }
  const selected = new Set(ui.archiveSelectionOrderIds || []);
  if (checkbox.checked) {
    selected.add(orderId);
  } else {
    selected.delete(orderId);
  }
  ui.archiveSelectionOrderIds = Array.from(selected);
  if (el.archiveCompletedOrdersBtn) {
    el.archiveCompletedOrdersBtn.disabled = ui.archiveSelectionOrderIds.length === 0;
  }
}

function onArchiveTableClick(event) {
  const detailsBtn = event.target.closest("button[data-action='open-order-details']");
  if (!detailsBtn) {
    return;
  }
  openOrderDetails(detailsBtn.dataset.orderId, { switchToOrders: false });
}

function onOrderPositionsTableClick(event) {
  const uploadBtn = event.target.closest("button[data-action='upload-attachment']");
  if (uploadBtn) {
    safeAction(() => chooseAttachmentForPosition(uploadBtn.dataset.positionId));
    return;
  }
  const deleteBtn = event.target.closest("button[data-action='delete-attachment']");
  if (deleteBtn) {
    safeAction(() => deletePositionAttachment(deleteBtn.dataset.positionId));
    return;
  }
  const attachmentBtn = event.target.closest("button[data-action='open-attachment']");
  if (attachmentBtn) {
    safeAction(() => openPositionAttachment(attachmentBtn.dataset.positionId, attachmentBtn.dataset.fileName || "zalacznik"));
    return;
  }
  const btn = event.target.closest("button[data-action='edit-position']");
  if (!btn) {
    return;
  }
  openPositionModalForEdit(btn.dataset.positionId);
}

function onKpiCardClick(event) {
  const tile = event.target.closest("button[data-kpi-type]");
  if (!tile) {
    return;
  }
  ui.kpiFilter = {
    type: tile.dataset.kpiType,
    value: tile.dataset.kpiValue || "",
    label: tile.dataset.kpiLabel || "",
  };
  renderKpiDrilldown();
}

function onKpiDrilldownClick(event) {
  const detailsBtn = event.target.closest("button[data-action='open-kpi-order-details']");
  if (!detailsBtn) {
    return;
  }
  openOrderDetails(detailsBtn.dataset.orderId, { switchToOrders: false });
}

function onGanttOrderClick(event) {
  const clearManualBtn = event.target.closest("button[data-gantt-clear-manual]");
  if (clearManualBtn) {
    const orderId = clearManualBtn.dataset.ganttClearManual;
    if (!orderId) {
      return;
    }
    safeAction(() => updateOrderManualStart(orderId, null));
    return;
  }

  const toggleBtn = event.target.closest("button[data-gantt-toggle-order]");
  if (toggleBtn) {
    const orderId = toggleBtn.dataset.ganttToggleOrder;
    if (!orderId) {
      return;
    }
    ui.ganttExpandedOrders[orderId] = !ui.ganttExpandedOrders[orderId];
    renderGantt();
    return;
  }

  const orderBtn = event.target.closest("button[data-gantt-order-id]");
  if (!orderBtn) {
    return;
  }
  const orderId = orderBtn.dataset.ganttOrderId;
  if (!orderId) {
    return;
  }
  openOrderDetails(orderId, { switchToOrders: true });
}

function onGanttDragStart(event) {
  const dragCell = event.target.closest("[data-gantt-drag-order-cell]");
  if (!dragCell) {
    return;
  }
  const orderId = String(dragCell.dataset.ganttDragOrderCell || "");
  if (!orderId) {
    return;
  }
  ui.ganttDragOrderId = orderId;
  event.dataTransfer.effectAllowed = "move";
  event.dataTransfer.setData("text/plain", orderId);
  el.ganttBoard.classList.add("cg-drag-mode");
}

function onGanttDragOver(event) {
  if (!ui.ganttDragOrderId) {
    return;
  }
  const dropCell = event.target.closest("td[data-gantt-date]");
  if (!dropCell) {
    return;
  }
  event.preventDefault();
  const current = el.ganttBoard.querySelector(".cg-day-cell.cg-drop-target");
  if (current && current !== dropCell) {
    current.classList.remove("cg-drop-target");
  }
  dropCell.classList.add("cg-drop-target");
}

function onGanttDrop(event) {
  if (!ui.ganttDragOrderId) {
    return;
  }
  const dropCell = event.target.closest("td[data-gantt-date]");
  if (!dropCell) {
    return;
  }
  event.preventDefault();
  const date = String(dropCell.dataset.ganttDate || "");
  const orderId = ui.ganttDragOrderId;
  clearGanttDragState();
  safeAction(() => updateOrderManualStart(orderId, date));
}

function onGanttDragEnd() {
  clearGanttDragState();
}

function clearGanttDragState() {
  ui.ganttDragOrderId = "";
  el.ganttBoard.classList.remove("cg-drag-mode");
  el.ganttBoard.querySelectorAll(".cg-day-cell.cg-drop-target").forEach((cell) => {
    cell.classList.remove("cg-drop-target");
  });
}

function onGanttCalendarToggleChange(event) {
  const checkbox = event.target.closest('input[data-gantt-workday-toggle="1"]');
  if (!checkbox) {
    return;
  }
  const date = String(checkbox.dataset.date || "");
  if (!date) {
    return;
  }
  safeAction(() => saveCalendarDayWorking(date, checkbox.checked));
}

function applyGanttDaysCount() {
  const rawBack = toInt(el.ganttDaysBackInput?.value);
  const rawForward = toInt(el.ganttDaysForwardInput?.value);
  const boundedBack = clamp(rawBack, 0, 730);
  const boundedForward = clamp(rawForward, 1, 730);
  ui.ganttDaysBackToShow = boundedBack;
  ui.ganttDaysForwardToShow = boundedForward;
  if (el.ganttDaysBackInput) {
    el.ganttDaysBackInput.value = String(boundedBack);
  }
  if (el.ganttDaysForwardInput) {
    el.ganttDaysForwardInput.value = String(boundedForward);
  }
  renderGantt();
}

function clearGanttFilters() {
  if (el.ganttSearchInput) {
    el.ganttSearchInput.value = "";
  }
  if (el.ganttStatusFilter) {
    el.ganttStatusFilter.value = "Wszystkie";
  }
  renderGantt();
}

function clearFeedbackFilters() {
  if (el.feedbackSearchInput) {
    el.feedbackSearchInput.value = "";
  }
  if (el.feedbackOrderStatusFilter) {
    el.feedbackOrderStatusFilter.value = "Wszystkie";
  }
  if (el.feedbackOrderPositionStatusFilter) {
    el.feedbackOrderPositionStatusFilter.value = "Wszystkie";
  }
  if (el.feedbackWorkflowStatusSelect) {
    el.feedbackWorkflowStatusSelect.value = "";
  }
  if (el.feedbackProgressPercentInput) {
    el.feedbackProgressPercentInput.value = "";
  }
  renderOrderSelectors();
  renderFeedbackPositions();
}

function clearOrdersFilters() {
  if (el.ordersSearchInput) {
    el.ordersSearchInput.value = "";
  }
  if (el.ordersPositionSearchInput) {
    el.ordersPositionSearchInput.value = "";
  }
  if (el.ordersStatusFilter) {
    el.ordersStatusFilter.value = "Wszystkie";
  }
  if (el.ordersWeekFilterInput) {
    el.ordersWeekFilterInput.value = "";
  }
  renderOrdersTable();
}

function archiveSelectableOrderIds(orders) {
  const source = Array.isArray(orders) ? orders : operationalOrders();
  return source
    .filter((order) => !order.archived && (order.orderStatus || "") === "Zakonczone")
    .map((order) => String(order.id || ""))
    .filter(Boolean);
}

function pruneArchiveSelection() {
  const allowed = new Set(archiveSelectableOrderIds());
  ui.archiveSelectionOrderIds = (ui.archiveSelectionOrderIds || []).filter((orderId) => allowed.has(orderId));
}

function selectCompletedOrdersForArchive() {
  const visibleOrders = sortOrdersForTable(operationalOrders().filter(orderMatchesMainFilters));
  const ids = archiveSelectableOrderIds(visibleOrders);
  ui.archiveSelectionOrderIds = ids;
  renderOrdersTable();
}

function clearArchiveFilters() {
  if (el.archiveSearchInput) {
    el.archiveSearchInput.value = "";
  }
  if (el.archiveWeekFilterInput) {
    el.archiveWeekFilterInput.value = "";
  }
  if (el.archiveDateFilterInput) {
    el.archiveDateFilterInput.value = "";
  }
  renderArchiveTable();
}

async function archiveCompletedOrders() {
  pruneArchiveSelection();
  const selectedIds = Array.from(new Set(ui.archiveSelectionOrderIds || []));
  if (selectedIds.length === 0) {
    alert("Najpierw zaznacz zakonczone zamowienia do archiwizacji.");
    return;
  }
  const payload = await api("/api/orders/archive-completed", { method: "POST", body: { orderIds: selectedIds } });
  ui.archiveSelectionOrderIds = [];
  await reloadAndRender();
  const count = toInt(payload.archivedCount);
  alert(count > 0 ? `Przeniesiono do archiwum: ${count}` : "Brak zaznaczonych zakonczonych zamowien do archiwizacji.");
}

function setView(name, force = false) {
  const target = String(name || "dashboard");
  if (!force && !canAccessSection(target)) {
    return;
  }
  ui.currentView = target;
  el.navButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.section === target));
  el.views.forEach((view) => view.classList.toggle("active", view.id === `view-${target}`));
}

async function reloadAndRender() {
  if (!isLoggedIn()) {
    renderAll();
    return;
  }
  await reloadState();
  renderAll();
}

async function reloadState() {
  const data = await api("/api/bootstrap");
  state.orders = Array.isArray(data.orders) ? data.orders : [];
  state.feedbackEvents = Array.isArray(data.feedbackEvents) ? data.feedbackEvents : [];
  state.users = Array.isArray(data.users) ? data.users : [];
  state.settings = normalizeSettings(data.settings || {});
  state.stations = normalizeStations(data.stations);
  state.stationSettings = data.stationSettings || {};
  state.technologies = data.technologies || {};
  state.materialRules = normalizeMaterialRules(data.materialRules || {});
  state.skillWorkers = normalizeSkillWorkers(data.skillWorkers || []);
  state.skillAvailability = normalizeSkillAvailability(data.skillAvailability || []);
  state.databases = Array.isArray(data.databases) ? data.databases : [];
  state.activeDatabaseKey = String(data.activeDatabase || "default");
  state.databaseAccessMap = normalizeDatabaseAccessMap(data.databaseAccessMap || {});
  recalculateOrders();
}

function isLoggedIn() {
  return Boolean(ui.currentUser?.id);
}

function isAdminUser() {
  return String(ui.currentUser?.role || "").toLowerCase() === "admin";
}

function canCreateDatabaseVariants() {
  if (isAdminUser()) {
    return true;
  }
  return Boolean(ui.currentUser?.canCreateDatabases);
}

function normalizeUserLoginKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function databaseCatalogSource() {
  if (Array.isArray(state.databases) && state.databases.length > 0) {
    return state.databases;
  }
  return [{ key: "default", name: "Domyslna baza", fileName: "planner.db", active: true, variant: false }];
}

function normalizeDatabaseAccessMap(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return {};
  }
  const validKeys = new Set(
    databaseCatalogSource()
      .map((item) => String(item?.key || "").trim())
      .filter((item) => Boolean(item)),
  );
  const output = {};
  Object.entries(raw).forEach(([rawLogin, rawKeys]) => {
    const loginKey = normalizeUserLoginKey(rawLogin);
    if (!loginKey) {
      return;
    }
    const source = Array.isArray(rawKeys) ? rawKeys : [];
    const keys = [];
    source.forEach((entry) => {
      const key = String(entry || "").trim();
      if (!key) {
        return;
      }
      if (validKeys.size > 0 && !validKeys.has(key)) {
        return;
      }
      if (keys.includes(key)) {
        return;
      }
      keys.push(key);
    });
    output[loginKey] = keys;
  });
  return output;
}

function allowedDatabaseKeysForUser(user) {
  const source = databaseCatalogSource();
  const allKeys = source.map((item) => String(item?.key || "")).filter((item) => Boolean(item));
  const role = String(user?.role || "user").toLowerCase();
  if (role === "admin") {
    return allKeys;
  }
  const loginKey = normalizeUserLoginKey(user?.login);
  if (!loginKey) {
    return [];
  }
  const mapped = state.databaseAccessMap?.[loginKey];
  if (!Array.isArray(mapped)) {
    return allKeys;
  }
  return mapped.filter((key) => allKeys.includes(String(key || "")));
}

function normalizeVisibleSections(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }
  const out = [];
  raw.forEach((section) => {
    const value = String(section || "").trim();
    if (!USER_VISIBLE_SECTIONS.includes(value)) {
      return;
    }
    if (!out.includes(value)) {
      out.push(value);
    }
  });
  return out;
}

function normalizeSkillWorkers(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }
  const departments = availableSkillDepartments();
  const stationIds = new Set((state.stations || []).map((station) => String(station?.id || "").trim()));
  const workers = [];
  raw.forEach((item) => {
    if (!item || typeof item !== "object") {
      return;
    }
    const workerId = String(item.id || "").trim();
    if (!workerId) {
      return;
    }
    const department = departments.includes(String(item.department || "").trim())
      ? String(item.department || "").trim()
      : departments[0] || DEPARTMENTS[0];
    const primaryStationIdRaw = String(item.primaryStationId || item.primary_station_id || "").trim();
    const primaryStationId = stationIds.has(primaryStationIdRaw) ? primaryStationIdRaw : "";
    const assignedShift = clamp(toInt(item.assignedShift ?? item.assigned_shift ?? 1), 1, 3);
    const skillMap = {};
    if (item.skills && typeof item.skills === "object" && !Array.isArray(item.skills)) {
      Object.entries(item.skills).forEach(([rawStationId, rawLevel]) => {
        const stationId = String(rawStationId || "").trim();
        if (!stationId || (stationIds.size > 0 && !stationIds.has(stationId))) {
          return;
        }
        const level = clamp(toInt(rawLevel), 0, 3);
        if (level > 0) {
          skillMap[stationId] = level;
        }
      });
    }
    workers.push({
      id: workerId,
      name: String(item.name || "").trim(),
      department,
      primaryStationId,
      assignedShift,
      active: item.active !== false,
      skills: skillMap,
    });
  });
  workers.sort((a, b) => {
    const depDiff = String(a.department || "").localeCompare(String(b.department || ""), "pl", { sensitivity: "base" });
    if (depDiff !== 0) {
      return depDiff;
    }
    const nameDiff = String(a.name || "").localeCompare(String(b.name || ""), "pl", { sensitivity: "base" });
    if (nameDiff !== 0) {
      return nameDiff;
    }
    return String(a.id || "").localeCompare(String(b.id || ""));
  });
  return workers;
}

function normalizeSkillAvailability(raw) {
  const out = {};
  if (!Array.isArray(raw)) {
    return out;
  }
  raw.forEach((item) => {
    const workerId = String(item?.workerId || item?.worker_id || "").trim();
    const date = normalizeOptionalDate(item?.date || item?.day);
    const shift = clamp(toInt(item?.shift), 1, 3);
    const minutes = clamp(toInt(item?.minutes), 0, 1440);
    if (!workerId || !date) {
      return;
    }
    if (!out[workerId]) {
      out[workerId] = {};
    }
    if (!out[workerId][date]) {
      out[workerId][date] = {};
    }
    out[workerId][date][String(shift)] = minutes;
  });
  return out;
}

function normalizeCurrentUser(raw) {
  const role = String(raw?.role || "user").toLowerCase() === "admin" ? "admin" : "user";
  return {
    id: String(raw?.id || "").trim(),
    name: String(raw?.name || "").trim(),
    department: String(raw?.department || "").trim(),
    login: String(raw?.login || "").trim(),
    role,
    visibleSections: normalizeVisibleSections(raw?.visibleSections ?? raw?.visible_sections),
    canCreateDatabases: Boolean(raw?.canCreateDatabases ?? raw?.can_create_databases),
  };
}

function visibleSectionsForCurrentUser() {
  if (!isLoggedIn()) {
    return [];
  }
  if (isAdminUser()) {
    return USER_VISIBLE_SECTIONS.slice();
  }
  return normalizeVisibleSections(ui.currentUser?.visibleSections || []);
}

function canAccessSection(section) {
  const target = String(section || "");
  if (target === "dashboard") {
    return true;
  }
  return visibleSectionsForCurrentUser().includes(target);
}

function firstVisibleSectionForCurrentUser() {
  const sections = visibleSectionsForCurrentUser();
  return sections[0] || "dashboard";
}

function applyAccessControl() {
  const visible = new Set(["dashboard", ...visibleSectionsForCurrentUser()]);
  el.navButtons.forEach((btn) => {
    const section = String(btn.dataset.section || "");
    const allow = visible.has(section);
    btn.classList.toggle("auth-hidden", !allow);
    btn.disabled = !allow;
  });
  document.body.classList.toggle("app-locked", !isLoggedIn());
  const desiredView = visible.has(ui.currentView) ? ui.currentView : firstVisibleSectionForCurrentUser();
  setView(desiredView || "dashboard", true);
}

function normalizeSettings(raw) {
  const days = Array.isArray(raw.workingDays)
    ? raw.workingDays.map((item) => toInt(item)).filter((item) => item >= 0 && item <= 6)
    : Array.isArray(raw.working_days)
    ? raw.working_days.map((item) => toInt(item)).filter((item) => item >= 0 && item <= 6)
    : [1, 2, 3, 4, 5];
  const weekdayShifts = normalizeWeekdayShifts(raw.weekdayShifts ?? raw.weekday_shifts, days);
  const daysFromShifts = Object.entries(weekdayShifts)
    .filter(([, shifts]) => toInt(shifts) > 0)
    .map(([day]) => toInt(day))
    .sort((a, b) => a - b);
  const calendarOverrides = normalizeCalendarOverrides(raw.calendarOverrides ?? raw.calendar_overrides);
  const stationOvertime = normalizeStationOvertime(raw.stationOvertime ?? raw.station_overtime);
  return {
    minutesPerShift: toInt(raw.minutesPerShift ?? raw.minutes_per_shift ?? 480),
    workingDays: daysFromShifts.length > 0 ? daysFromShifts : days.length > 0 ? days : [1, 2, 3, 4, 5],
    weekdayShifts,
    stationOvertime,
    calendarOverrides,
  };
}

function normalizeWeekdayShifts(raw, fallbackDays = [1, 2, 3, 4, 5]) {
  const base = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
  const defaultDays = Array.isArray(fallbackDays) && fallbackDays.length > 0 ? fallbackDays : [1, 2, 3, 4, 5];
  defaultDays.forEach((day) => {
    const dayIndex = clamp(toInt(day), 0, 6);
    base[dayIndex] = 2;
  });
  if (raw && typeof raw === "object" && !Array.isArray(raw)) {
    Object.entries(raw).forEach(([day, shifts]) => {
      const dayIndex = clamp(toInt(day), 0, 6);
      base[dayIndex] = clamp(toInt(shifts), 0, 3);
    });
  }
  const hasWorkingDay = Object.values(base).some((value) => toInt(value) > 0);
  if (!hasWorkingDay) {
    return { 0: 0, 1: 2, 2: 2, 3: 2, 4: 2, 5: 2, 6: 0 };
  }
  return base;
}

function normalizeCalendarOverrides(raw) {
  if (!raw || typeof raw !== "object") {
    return {};
  }
  const result = {};
  Object.entries(raw).forEach(([date, value]) => {
    const key = String(date || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) {
      return;
    }
    result[key] = Boolean(value);
  });
  return result;
}

function normalizeStationOvertime(raw) {
  if (!raw || typeof raw !== "object") {
    return {};
  }
  const result = {};
  Object.entries(raw).forEach(([date, value]) => {
    const key = String(date || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) {
      return;
    }
    if (!value || typeof value !== "object") {
      return;
    }
    const dayMap = {};
    Object.entries(value).forEach(([stationId, minutes]) => {
      const normalizedStationId = String(stationId || "").trim();
      if (!normalizedStationId) {
        return;
      }
      const normalizedMinutes = clamp(toInt(minutes), 0, 1440);
      if (normalizedMinutes > 0) {
        dayMap[normalizedStationId] = normalizedMinutes;
      }
    });
    if (Object.keys(dayMap).length > 0) {
      result[key] = dayMap;
    }
  });
  return result;
}

function normalizeMaterialRules(raw) {
  const defaults = {
    Maszynownia: ["wood", "corpus"],
    Lakiernia: ["wood", "corpus"],
    Kompletacja: ["glass", "hardware", "accessories"],
    Kosmetyka: ["accessories"],
  };
  const result = structuredClone(defaults);
  Object.keys(raw || {}).forEach((dept) => {
    if (Array.isArray(raw[dept])) {
      result[dept] = raw[dept].filter((item) => MATERIALS.some((mat) => mat.key === item));
    }
  });
  return result;
}

function normalizeStations(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }
  return raw
    .map((item, index) => ({
      id: String(item?.id || "").trim(),
      name: String(item?.name || "").trim(),
      department: String(item?.department || "").trim() || "Maszynownia",
      active: item?.active !== false,
      sortOrder: Math.max(1, toInt(item?.sortOrder ?? item?.sort_order ?? index + 1)),
    }))
    .filter((item) => item.id && item.name)
    .sort((a, b) => {
      const diff = toInt(a.sortOrder) - toInt(b.sortOrder);
      if (diff !== 0) {
        return diff;
      }
      return a.id.localeCompare(b.id);
    });
}

function operationalOrders() {
  return state.orders.filter((order) => !order.archived);
}

function archivedOrders() {
  return state.orders.filter((order) => Boolean(order.archived));
}

function isoWeekNumber(dateValue) {
  const raw = normalizeOptionalDate(dateValue);
  if (!raw) {
    return null;
  }
  const date = toDate(raw);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  const day = (date.getDay() + 6) % 7;
  const thursday = new Date(date);
  thursday.setDate(date.getDate() - day + 3);
  const firstThursday = new Date(thursday.getFullYear(), 0, 4);
  const firstDay = (firstThursday.getDay() + 6) % 7;
  firstThursday.setDate(firstThursday.getDate() - firstDay + 3);
  const diff = thursday.getTime() - firstThursday.getTime();
  return 1 + Math.round(diff / 604800000);
}

function formatWeekLabel(dateValue) {
  const week = isoWeekNumber(dateValue);
  return week ? `T${week}` : "-";
}

function isPlannedDateEdited(order) {
  return Boolean(normalizeOptionalDate(order?.manualPlannedDate));
}

function formatPlannedDateCell(order) {
  const planned = effectivePlannedDate(order);
  if (!planned) {
    return "-";
  }
  const label = escapeHtml(formatDate(planned));
  if (!isPlannedDateEdited(order)) {
    return label;
  }
  return `${label} <span class="edited-plan-badge" title="Data planowana edytowana recznie">E</span>`;
}

function orderMatchesMainFilters(order) {
  const search = normalizeSearchText(el.ordersSearchInput?.value).trim();
  const positionSearch = normalizeSearchText(el.ordersPositionSearchInput?.value).trim();
  const status = String(el.ordersStatusFilter?.value || "Wszystkie");
  const weekFilter = toInt(el.ordersWeekFilterInput?.value || "");
  if (search) {
    const haystack = normalizeSearchText([order.orderNumber, order.client, order.owner].join(" "));
    if (!haystack.includes(search)) {
      return false;
    }
  }
  if (status && status !== "Wszystkie" && (order.orderStatus || "Dokumentacja") !== status) {
    return false;
  }
  if (positionSearch && !orderMatchesPositionSearch(order, positionSearch)) {
    return false;
  }
  if (weekFilter > 0) {
    const week = isoWeekNumber(effectivePlannedDate(order));
    if (week !== weekFilter) {
      return false;
    }
  }
  return true;
}

function orderMatchesPositionSearch(order, search) {
  const headerText = [order.orderNumber, order.color, order.extras]
    .map((value) => normalizeSearchText(value))
    .join(" ");
  const positionsText = (order.positions || [])
    .map((position) => {
      const materialsText = Object.entries(position.materials || {})
        .map(([key, value]) => `${key} ${value?.date || ""} ${value?.toOrder ? "do zamowienia" : ""}`)
        .join(" ");
      return [
        position.positionNumber,
        position.technology,
        position.system,
        position.color,
        position.line,
        position.width,
        position.height,
        `${position.width}x${position.height}`,
        positionShapeLabel(position),
        positionElementsLabel(position),
        position.currentDepartmentStatus,
        position.notes,
        materialsText,
      ]
        .map((value) => normalizeSearchText(value))
        .join(" ");
    })
    .join(" ");

  return `${headerText} ${positionsText}`.includes(search);
}

function normalizeSearchText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ł/g, "l");
}

function sortOrdersForTable(orders) {
  const items = (orders || []).slice();
  const sortKey = ui.ordersSortKey || "createdAt";
  const direction = ui.ordersSortDir === "asc" ? 1 : -1;
  items.sort((a, b) => {
    const av = orderSortValue(a, sortKey);
    const bv = orderSortValue(b, sortKey);
    const cmp = compareOrderSortValues(av, bv);
    if (cmp !== 0) {
      return cmp * direction;
    }
    return String(a.orderNumber || "").localeCompare(String(b.orderNumber || ""), "pl", { numeric: true });
  });
  return items;
}

function orderSortValue(order, key) {
  switch (key) {
    case "orderNumber":
      return String(order.orderNumber || "");
    case "client":
      return String(order.client || "");
    case "extras":
      return String(order.extras || "");
    case "entryDate":
      return dateSortNumber(order.entryDate);
    case "startDate":
      return dateSortNumber(effectiveStartDate(order));
    case "plannedProductionDate":
      return dateSortNumber(effectivePlannedDate(order));
    case "plannedWeek":
      return isoWeekNumber(effectivePlannedDate(order)) || 0;
    case "positionsCount":
      return Array.isArray(order.positions) ? order.positions.length : 0;
    case "orderStatus":
      return String(order.orderStatus || "Dokumentacja");
    case "completedAt":
      return dateSortNumber(order.completedAt);
    case "blockedMaterials":
      return pendingMaterialsCount(order);
    case "createdAt":
      return dateSortNumber(order.createdAt);
    default:
      return String(order.orderNumber || "");
  }
}

function dateSortNumber(value) {
  const parsed = Date.parse(String(value || ""));
  return Number.isFinite(parsed) ? parsed : -1;
}

function compareOrderSortValues(a, b) {
  const bothNumbers = typeof a === "number" && typeof b === "number";
  if (bothNumbers) {
    return a - b;
  }
  return String(a || "").localeCompare(String(b || ""), "pl", { numeric: true, sensitivity: "base" });
}

function syncOrdersSortHeaderState() {
  if (!el.ordersTable) {
    return;
  }
  const headers = Array.from(el.ordersTable.querySelectorAll("th[data-order-sort]"));
  headers.forEach((header) => {
    header.classList.remove("sort-asc", "sort-desc");
    if (String(header.dataset.orderSort || "") !== ui.ordersSortKey) {
      return;
    }
    header.classList.add(ui.ordersSortDir === "asc" ? "sort-asc" : "sort-desc");
  });
}

function orderMatchesArchiveFilters(order) {
  const search = String(el.archiveSearchInput?.value || "")
    .trim()
    .toLowerCase();
  const weekFilter = toInt(el.archiveWeekFilterInput?.value || "");
  const dateFilter = normalizeOptionalDate(el.archiveDateFilterInput?.value || "");
  if (search) {
    const haystack = [order.orderNumber, order.client, order.owner].join(" ").toLowerCase();
    if (!haystack.includes(search)) {
      return false;
    }
  }
  if (weekFilter > 0) {
    const week = isoWeekNumber(effectivePlannedDate(order));
    if (week !== weekFilter) {
      return false;
    }
  }
  if (dateFilter) {
    const completedDate = normalizeOptionalDate(order?.completedAt);
    if (completedDate !== dateFilter) {
      return false;
    }
  }
  return true;
}

function orderMatchesGanttFilters(order) {
  const search = String(el.ganttSearchInput?.value || "")
    .trim()
    .toLowerCase();
  const status = String(el.ganttStatusFilter?.value || "Wszystkie");
  if (search) {
    const haystack = [order.orderNumber, order.client, order.owner].join(" ").toLowerCase();
    if (!haystack.includes(search)) {
      return false;
    }
  }
  if (status && status !== "Wszystkie" && (order.orderStatus || "Dokumentacja") !== status) {
    return false;
  }
  return true;
}

function renderAll() {
  renderLoginPanel();
  applyAccessControl();
  if (!isLoggedIn()) {
    el.activeOrdersCount.textContent = "0";
    return;
  }
  syncGanttDaysControl();
  renderTechnologySelects();
  renderOrderSelectors();
  ensureSelectedOrder();
  renderOrdersTable();
  renderArchiveTable();
  renderSelectedOrderDetails();
  renderKpiCards();
  renderKpiDrilldown();
  renderGantt();
  renderReportStationSelect();
  ensureReportDefaults();
  renderReport();
  renderExecution();
  renderSkillWorkerForm();
  renderSkillWorkersList();
  renderSkillAvailabilityControls();
  renderSkillAvailabilityCalendar();
  ensureSkillsAllocationDefaults();
  renderSkillsAllocation();
  renderFeedbackPositions();
  renderUsers();
  renderDatabaseAccessMatrix();
  renderDatabaseManager();
  syncSettingsForm();
  renderStationOvertimeEditor();
  renderStationSettingsEditor();
  renderMaterialRulesEditor();
  renderTechnologyAllocationEditor();
}

function syncLoginDatabaseSelect() {
  if (!el.loginDatabaseSelect) {
    return;
  }
  const source =
    Array.isArray(state.databases) && state.databases.length > 0
      ? state.databases
      : [{ key: "default", name: "Domyslna baza", fileName: "planner.db", active: true, variant: false }];
  const currentValue = String(el.loginDatabaseSelect.value || "").trim();
  el.loginDatabaseSelect.innerHTML = source
    .map((item) => {
      const key = String(item.key || "");
      const label = `${item.name || key} - ${item.fileName || ""}`;
      return `<option value="${escapeHtml(key)}">${escapeHtml(label)}</option>`;
    })
    .join("");
  const validCurrent = source.some((item) => String(item.key || "") === currentValue);
  const validActive = source.some((item) => String(item.key || "") === state.activeDatabaseKey);
  el.loginDatabaseSelect.value = validCurrent ? currentValue : validActive ? state.activeDatabaseKey : String(source[0]?.key || "default");
}

function renderLoginPanel() {
  const logged = isLoggedIn();
  syncLoginDatabaseSelect();
  if (el.loginStatusText) {
    el.loginStatusText.textContent = logged
      ? "Zalogowano. Masz dostep tylko do przypisanych sekcji."
      : "Zaloguj sie, aby odblokowac pozostale sekcje.";
  }
  if (el.loginBtn) {
    el.loginBtn.classList.toggle("panel-hidden", logged);
  }
  if (el.logoutBtn) {
    el.logoutBtn.classList.toggle("panel-hidden", !logged);
  }
  if (el.loginInput) {
    el.loginInput.disabled = logged;
  }
  if (el.loginPasswordInput) {
    el.loginPasswordInput.disabled = logged;
    if (logged) {
      el.loginPasswordInput.value = "";
    }
  }
  if (el.loginDatabaseSelect) {
    el.loginDatabaseSelect.disabled = logged;
  }
  if (el.currentUserInfo) {
    if (logged) {
      const roleLabel = isAdminUser() ? "Admin" : "Uzytkownik";
      const sections = isAdminUser()
        ? "Wszystkie sekcje"
        : visibleSectionsForCurrentUser()
            .map((section) => SECTION_LABELS[section] || section)
            .join(", ") || "Brak sekcji";
      const dbLabel =
        state.databases.find((item) => String(item.key || "") === state.activeDatabaseKey)?.name || state.activeDatabaseKey;
      const creationLabel = canCreateDatabaseVariants() ? "Tworzenie baz: TAK" : "Tworzenie baz: NIE";
      el.currentUserInfo.innerHTML = `<strong>${escapeHtml(
        ui.currentUser.name || ui.currentUser.login || "Uzytkownik",
      )}</strong> | Rola: ${escapeHtml(roleLabel)} | Baza: ${escapeHtml(dbLabel)} | ${escapeHtml(
        creationLabel,
      )} | Dostep: ${escapeHtml(sections)}`;
      el.currentUserInfo.classList.remove("panel-hidden");
    } else {
      el.currentUserInfo.textContent = "";
      el.currentUserInfo.classList.add("panel-hidden");
    }
  }
  if (el.postLoginLinks) {
    if (!logged) {
      el.postLoginLinks.innerHTML = "";
      el.postLoginLinks.classList.add("panel-hidden");
      return;
    }
    const cards = visibleSectionsForCurrentUser().map((section) => {
      return `
        <article class="card tile">
          <h3>${escapeHtml(SECTION_LABELS[section] || section)}</h3>
          <p>Przejdz do sekcji.</p>
          <button type="button" data-go="${escapeHtml(section)}">Przejdz</button>
        </article>
      `;
    });
    el.postLoginLinks.innerHTML = cards.join("");
    el.postLoginLinks.classList.toggle("panel-hidden", cards.length === 0);
  }
}

function syncGanttDaysControl() {
  if (el.ganttDaysBackInput) {
    el.ganttDaysBackInput.value = String(clamp(toInt(ui.ganttDaysBackToShow), 0, 730));
  }
  if (el.ganttDaysForwardInput) {
    el.ganttDaysForwardInput.value = String(clamp(toInt(ui.ganttDaysForwardToShow), 1, 730));
  }
}

function renderTechnologySelects() {
  const names = Object.keys(state.technologies);
  const markup = names.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  el.technologySelect.innerHTML = markup || '<option value="">Brak technologii</option>';
  el.technologyEditorSelect.innerHTML = markup || '<option value="">Brak technologii</option>';

  const selectedPosition = names.includes(el.technologySelect.value) ? el.technologySelect.value : names[0] || "";
  const selectedEditor = names.includes(ui.selectedTechnology) ? ui.selectedTechnology : names[0] || "";
  el.technologySelect.value = selectedPosition;
  ui.selectedTechnology = selectedEditor;
  el.technologyEditorSelect.value = selectedEditor;
}

function renderOrderSelectors() {
  const source = feedbackFilteredOrders();
  const options = ['<option value="">Wybierz zamowienie</option>']
    .concat(
      source.map(
        (order) => `<option value="${order.id}">${escapeHtml(order.orderNumber)} - ${escapeHtml(order.client)}</option>`,
      ),
    )
    .join("");
  const validIds = source.map((item) => item.id);
  const feedbackCurrent = validIds.includes(el.feedbackOrderSelect.value) ? el.feedbackOrderSelect.value : validIds[0] || "";

  el.feedbackOrderSelect.innerHTML = options;
  el.feedbackOrderSelect.value = feedbackCurrent;
}

function feedbackFilteredOrders() {
  const orderStatusFilter = String(el.feedbackOrderStatusFilter?.value || "Wszystkie");
  const positionStatusFilter = String(el.feedbackOrderPositionStatusFilter?.value || "Wszystkie");
  return operationalOrders().filter((order) => {
    if (orderStatusFilter !== "Wszystkie" && (order.orderStatus || "Dokumentacja") !== orderStatusFilter) {
      return false;
    }
    if (positionStatusFilter !== "Wszystkie" && orderKpiPositionStatus(order) !== positionStatusFilter) {
      return false;
    }
    return true;
  });
}

function ensureSelectedOrder() {
  const allIds = state.orders.map((item) => item.id);
  if (allIds.includes(ui.selectedOrderId)) {
    return;
  }
  const activeIds = operationalOrders().map((item) => item.id);
  ui.selectedOrderId = activeIds[0] || allIds[0] || "";
}

function orderHasAnyAttachment(order) {
  if (!order || !Array.isArray(order.positions)) {
    return false;
  }
  return order.positions.some((position) => String(position?.attachmentName || "").trim().length > 0);
}

function renderOrdersTable() {
  const active = operationalOrders().filter((order) => (order.orderStatus || "") !== "Zakonczone").length;
  const filteredOrders = operationalOrders().filter(orderMatchesMainFilters);
  const visibleOrders = sortOrdersForTable(filteredOrders);
  pruneArchiveSelection();
  const completedVisibleIds = archiveSelectableOrderIds(visibleOrders);
  const selectedArchiveCount = (ui.archiveSelectionOrderIds || []).length;
  if (el.selectCompletedOrdersBtn) {
    el.selectCompletedOrdersBtn.disabled = completedVisibleIds.length === 0;
  }
  if (el.archiveCompletedOrdersBtn) {
    el.archiveCompletedOrdersBtn.disabled = selectedArchiveCount === 0;
  }
  syncOrdersSortHeaderState();
  if (visibleOrders.length === 0) {
    el.ordersTableBody.innerHTML = `<tr><td colspan="12">Brak zamowien dla aktualnego filtra.</td></tr>`;
    el.activeOrdersCount.textContent = String(active);
    return;
  }

  const selectedArchiveOrders = new Set(ui.archiveSelectionOrderIds || []);
  el.ordersTableBody.innerHTML = visibleOrders
    .map((order) => {
      const blockedMaterials = pendingMaterialsCount(order);
      const hasAttachment = orderHasAnyAttachment(order);
      const selectedClass = order.id === ui.selectedOrderId ? "selected-row" : "";
      const canArchive = (order.orderStatus || "") === "Zakonczone";
      const archiveChecked = canArchive && selectedArchiveOrders.has(order.id) ? "checked" : "";
      const archiveControl = canArchive
        ? `<label class="archive-order-toggle"><input type="checkbox" data-action="toggle-archive-order" data-order-id="${escapeHtml(
            order.id,
          )}" ${archiveChecked} /> Do archiwum</label>`
        : `<span class="archive-order-disabled">-</span>`;
      return `
        <tr data-order-id="${order.id}" class="${selectedClass}">
          <td>
            <div class="order-number-cell">
              <span>${escapeHtml(order.orderNumber)}</span>
              ${hasAttachment ? '<span class="order-attachment-marker" title="Co najmniej jedna pozycja ma zalacznik">ZAL</span>' : ""}
            </div>
          </td>
          <td>${escapeHtml(order.client)}</td>
          <td>${formatDate(order.entryDate)}</td>
          <td>${formatDate(effectiveStartDate(order))}</td>
          <td>${formatPlannedDateCell(order)}</td>
          <td>${formatWeekLabel(effectivePlannedDate(order))}</td>
          <td>${order.positions.length}</td>
          <td>${statusPill(order.orderStatus || "Dokumentacja")}</td>
          <td>${formatDate(order.completedAt)}</td>
          <td>${blockedMaterials}</td>
          <td>
            <button type="button" data-action="open-order-details" data-order-id="${order.id}">Szczegoly</button>
            ${archiveControl}
          </td>
          <td>${escapeHtml(order.extras || "-")}</td>
        </tr>
      `;
    })
    .join("");

  el.activeOrdersCount.textContent = String(active);
}

function renderArchiveTable() {
  const items = archivedOrders().filter(orderMatchesArchiveFilters);
  if (items.length === 0) {
    el.archiveTableBody.innerHTML = `<tr><td colspan="9">Brak zamowien w archiwum dla aktualnego filtra.</td></tr>`;
    return;
  }

  el.archiveTableBody.innerHTML = items
    .map(
      (order) => `
      <tr data-order-id="${order.id}">
        <td>${escapeHtml(order.orderNumber)}</td>
        <td>${escapeHtml(order.client)}</td>
        <td>${formatDate(order.entryDate)}</td>
        <td>${formatPlannedDateCell(order)}</td>
        <td>${formatWeekLabel(effectivePlannedDate(order))}</td>
        <td>${formatDate(order.completedAt)}</td>
        <td>${statusPill(order.orderStatus || "Dokumentacja")}</td>
        <td>${order.positions.length}</td>
        <td><button type="button" data-action="open-order-details" data-order-id="${order.id}">Szczegoly</button></td>
      </tr>`,
    )
    .join("");
}

function renderSelectedOrderDetails() {
  const order = state.orders.find((item) => item.id === ui.selectedOrderId);
  if (!order) {
    el.orderDetailsTitle.textContent = "Szczegoly zamowienia";
    fillSelect(el.selectedOrderStatusSelect, ORDER_STATUSES, ORDER_STATUSES[0]);
    setSelectedOrderEditorValues({
      orderNumber: "",
      entryDate: "",
      plannedProductionDate: "",
      manualPlannedDate: "",
      owner: "",
      client: "",
      color: "",
      framesCount: 0,
      sashesCount: 0,
      extras: "",
      orderStatus: ORDER_STATUSES[0],
    });
    setSelectedOrderEditorDisabled(true);
    setOrderDetailsActionState(false);
    el.selectedOrderPositionsBody.innerHTML = `<tr><td colspan="12">Wybierz zamowienie z listy.</td></tr>`;
    return;
  }

  el.orderDetailsTitle.textContent = `Szczegoly zamowienia ${order.orderNumber} (${order.client})`;
  setSelectedOrderEditorDisabled(!ui.orderDetailsEditMode);
  setOrderDetailsActionState(true);
  setSelectedOrderEditorValues({
    orderNumber: order.orderNumber,
    entryDate: order.entryDate,
    plannedProductionDate: order.plannedProductionDate,
    manualPlannedDate: order.manualPlannedDate,
    owner: order.owner,
    client: order.client,
    color: order.color,
    framesCount: order.framesCount,
    sashesCount: order.sashesCount,
    extras: order.extras,
    orderStatus: ORDER_STATUSES.includes(order.orderStatus) ? order.orderStatus : ORDER_STATUSES[0],
  });

  if (order.positions.length === 0) {
    el.selectedOrderPositionsBody.innerHTML = `<tr><td colspan="12">Brak pozycji w tym zamowieniu.</td></tr>`;
    return;
  }

  el.selectedOrderPositionsBody.innerHTML = order.positions
    .map((position) => {
      const materialsHtml = MATERIALS.map((material) => {
        const entry = position.materials?.[material.key] || { date: null, toOrder: false };
        return `${material.label}: ${materialStateLabel(entry)}`;
      }).join("<br>");

      const attachmentMarker = position.attachmentName
        ? '<span class="attachment-flag attachment-flag--has" title="Pozycja ma zalacznik">ZAL</span>'
        : '<span class="attachment-flag attachment-flag--none" title="Brak zalacznika">BRAK</span>';

      const attachment = position.attachmentName
        ? `<div class="position-attachment-cell">
            ${attachmentMarker}
            <div class="position-attachment-actions">
              <button type="button" data-action="open-attachment" data-position-id="${position.id}" data-file-name="${escapeHtml(
            position.attachmentName,
          )}">Otworz</button>
              <button type="button" data-action="upload-attachment" data-position-id="${position.id}">Zmien</button>
              <button type="button" class="ghost-btn" data-action="delete-attachment" data-position-id="${position.id}">Usun</button>
            </div>
          </div>`
        : `<div class="position-attachment-cell">
            ${attachmentMarker}
            <div class="position-attachment-actions">
              <button type="button" data-action="upload-attachment" data-position-id="${position.id}">Dodaj</button>
            </div>
          </div>`;

      return `
        <tr data-position-id="${position.id}">
          <td>${escapeHtml(position.positionNumber)}</td>
          <td>${position.width}x${position.height}</td>
          <td>${Math.max(0, toInt(position.framesCount))} / ${Math.max(0, toInt(position.sashesCount))}</td>
          <td>${escapeHtml(positionShapeLabel(position))}</td>
          <td>${escapeHtml(positionElementsLabel(position))}</td>
          <td>${escapeHtml(position.technology)}</td>
          <td>${toFloat(position.times.machining).toFixed(0)} / ${toFloat(position.times.painting).toFixed(0)} / ${toFloat(
        position.times.assembly,
      ).toFixed(0)}</td>
          <td>${statusPill(normalizeKpiPositionStatus(position.currentDepartmentStatus || POSITION_DEPARTMENT_STATUSES[0]))}</td>
          <td>${materialsHtml}</td>
          <td>${escapeHtml(position.notes || "-")}</td>
          <td>${attachment}</td>
          <td><button type="button" data-action="edit-position" data-position-id="${position.id}">Edytuj</button></td>
        </tr>
      `;
    })
    .join("");
}

function setSelectedOrderEditorValues(order) {
  el.selectedOrderNumberInput.value = order.orderNumber || "";
  el.selectedOrderEntryDateInput.value = order.entryDate || "";
  const edited = isPlannedDateEdited(order);
  el.selectedOrderPlannedDateInput.value = normalizeOptionalDate(order.plannedProductionDate) || "";
  el.selectedOrderPlannedDateInput.title = edited
    ? "Data planowana edytowana recznie"
    : "Data planowana wyliczona automatycznie";
  el.selectedOrderPlannedDateInput.classList.toggle("input-manual-date", edited);
  el.selectedOrderOwnerInput.value = order.owner || "";
  el.selectedOrderClientInput.value = order.client || "";
  el.selectedOrderColorInput.value = order.color || "";
  el.selectedOrderFramesInput.value = String(Math.max(0, toInt(order.framesCount)));
  el.selectedOrderSashesInput.value = String(Math.max(0, toInt(order.sashesCount)));
  el.selectedOrderExtrasInput.value = order.extras || "";
  el.selectedOrderStatusSelect.value = order.orderStatus || ORDER_STATUSES[0];
}

function setSelectedOrderEditorDisabled(disabled) {
  const fields = [
    el.selectedOrderNumberInput,
    el.selectedOrderEntryDateInput,
    el.selectedOrderPlannedDateInput,
    el.selectedOrderOwnerInput,
    el.selectedOrderClientInput,
    el.selectedOrderColorInput,
    el.selectedOrderExtrasInput,
  ];
  fields.forEach((field) => {
    if (field) {
      field.disabled = Boolean(disabled);
    }
  });
  if (el.selectedOrderFramesInput) {
    el.selectedOrderFramesInput.disabled = true;
  }
  if (el.selectedOrderSashesInput) {
    el.selectedOrderSashesInput.disabled = true;
  }
  if (el.selectedOrderStatusSelect) {
    el.selectedOrderStatusSelect.disabled = true;
  }
}

function setOrderDetailsActionState(hasOrder) {
  if (el.editSelectedOrderBtn) {
    el.editSelectedOrderBtn.disabled = !hasOrder || ui.orderDetailsEditMode;
  }
  if (el.saveSelectedOrderBtn) {
    el.saveSelectedOrderBtn.disabled = !hasOrder || !ui.orderDetailsEditMode;
  }
  if (el.openPositionModalBtn) {
    el.openPositionModalBtn.disabled = !hasOrder;
  }
}

function renderKpiCards() {
  const source = operationalOrders();
  const statusBuckets = [...KPI_POSITION_STATUSES, KPI_POSITION_MIXED_STATUS, KPI_POSITION_EMPTY_STATUS];
  const cards = [];
  statusBuckets.forEach((status) => {
    const count = source.filter((order) => orderKpiPositionStatus(order) === status).length;
    cards.push(makeKpiCard(`Status pozycji: ${status}`, count, "positionStatus", status));
  });
  el.kpiCards.innerHTML = cards.join("");
}

function makeKpiCard(title, value, type, key) {
  return `
    <button class="card kpi-card kpi-clickable" data-kpi-type="${escapeHtml(type)}" data-kpi-value="${escapeHtml(
    key,
  )}" data-kpi-label="${escapeHtml(title)}">
      <h3>${escapeHtml(title)}</h3>
      <p>${value}</p>
    </button>
  `;
}

function renderKpiDrilldown() {
  const filter = ui.kpiFilter || {
    type: "positionStatus",
    value: KPI_POSITION_STATUSES[0],
    label: `Status pozycji: ${KPI_POSITION_STATUSES[0]}`,
  };
  const filtered = filterOrdersByKpi(filter);
  el.kpiDrilldownTitle.textContent = `Lista zamowien po KPI: ${filter.label}`;
  if (filtered.length === 0) {
    el.kpiDrilldownBody.innerHTML = `<tr><td colspan="5">Brak zamowien dla wybranego kryterium.</td></tr>`;
    return;
  }
  el.kpiDrilldownBody.innerHTML = filtered
    .map((order) => {
      const positionStatus = orderKpiPositionStatus(order);
      return `
        <tr>
          <td>
            <button type="button" class="kpi-order-link-btn" data-action="open-kpi-order-details" data-order-id="${escapeHtml(
              order.id,
            )}">${escapeHtml(order.orderNumber)}</button>
          </td>
          <td>${statusPill(positionStatus)}</td>
          <td>${Math.max(0, toInt(order.positions?.length || 0))}</td>
          <td>${formatPlannedDateCell(order)}</td>
          <td><button type="button" data-action="open-kpi-order-details" data-order-id="${escapeHtml(order.id)}">Szczegoly</button></td>
        </tr>
      `;
    })
    .join("");
}

function filterOrdersByKpi(filter) {
  const source = operationalOrders();
  if (filter.type === "positionStatus") {
    return source.filter((order) => orderKpiPositionStatus(order) === filter.value);
  }
  return source;
}

function orderKpiPositionStatus(order) {
  const positions = Array.isArray(order?.positions) ? order.positions : [];
  if (positions.length === 0) {
    return KPI_POSITION_EMPTY_STATUS;
  }
  const uniqueStatuses = Array.from(
    new Set(
      positions.map((position) =>
        normalizeKpiPositionStatus(position?.currentDepartmentStatus || POSITION_DEPARTMENT_STATUSES[0]),
      ),
    ),
  );
  return uniqueStatuses.length === 1 ? uniqueStatuses[0] : KPI_POSITION_MIXED_STATUS;
}

function normalizeKpiPositionStatus(status) {
  const normalized = String(status || "").trim();
  if (normalized === "Planowanie") {
    return "Dokumentacja";
  }
  if (POSITION_DEPARTMENT_STATUSES.includes(normalized)) {
    return normalized;
  }
  return POSITION_DEPARTMENT_STATUSES[0];
}

function renderReportStationSelect() {
  if (!el.reportStationSelect) {
    return;
  }
  const activeStations = state.stations.filter((station) => station.active !== false);
  if (activeStations.length === 0) {
    el.reportStationSelect.innerHTML = '<option value="">Brak aktywnych stanowisk</option>';
    return;
  }
  const previous = el.reportStationSelect.value;
  el.reportStationSelect.innerHTML = activeStations
    .map((station) => `<option value="${escapeHtml(station.id)}">${escapeHtml(station.name)}</option>`)
    .join("");
  const valid = activeStations.some((station) => station.id === previous);
  el.reportStationSelect.value = valid ? previous : activeStations[0].id;
}

function ensureReportDefaults() {
  if (el.reportDate && !el.reportDate.value) {
    el.reportDate.value = isoDate(new Date());
  }
  if (el.executionAnchorDate && !el.executionAnchorDate.value) {
    el.executionAnchorDate.value = isoDate(new Date());
  }
  if (el.executionPeriodMode && !el.executionPeriodMode.value) {
    el.executionPeriodMode.value = "day";
  }
}

function sortedUniqueDates(values) {
  return Array.from(new Set((values || []).filter(Boolean))).sort();
}

function renderGanttChart(container, tasks, options = {}) {
  if (!container) {
    return;
  }
  if (!Array.isArray(tasks) || tasks.length === 0) {
    container.innerHTML = `<p>${options.emptyMessage || "Brak danych do Gantta."}</p>`;
    return;
  }

  const bounds = ganttBounds(tasks);
  const timelineDates = dateRange(bounds.start, bounds.end);
  if (timelineDates.length === 0) {
    container.innerHTML = `<p>${options.emptyMessage || "Brak danych do Gantta."}</p>`;
    return;
  }

  const dayWidth = 28;
  const timelineWidth = timelineDates.length * dayWidth;
  const indexByDate = {};
  timelineDates.forEach((date, index) => {
    indexByDate[date] = index;
  });
  const workingDaysCount = timelineDates.filter((date) => isWorkingDay(toDate(date))).length;
  const nonWorkingDaysCount = timelineDates.length - workingDaysCount;
  const weekGroups = groupTimelineByWeek(timelineDates);
  const dayMeta = timelineDates.map((date) => {
    const dt = toDate(date);
    const dayIndex = dt.getDay();
    return {
      date,
      dt,
      dayIndex,
      isSaturday: dayIndex === 6,
      isSunday: dayIndex === 0,
      working: isWorkingDay(dt),
    };
  });
  const offDayCells = timelineDates
    .map((date, idx) => {
      const info = dayMeta[idx];
      const cls = [
        "gantt-day-cell",
        info.working ? "" : "off",
        info.isSaturday ? "saturday" : "",
        info.isSunday ? "sunday" : "",
      ]
        .filter(Boolean)
        .join(" ");
      return `<span class="${cls}"></span>`;
    })
    .join("");

  const leftRows = tasks
    .map((task, index) => {
      const nameControl = options.clickable
        ? `<button type="button" class="gantt-task-link" data-gantt-order-id="${escapeHtml(task.id)}">${escapeHtml(task.name)}</button>`
        : `<span class="gantt-task-name-text">${escapeHtml(task.name)}</span>`;
      return `
        <div class="gantt-task-row ${index % 2 === 0 ? "even" : "odd"}">
          <div class="gantt-task-col name">
            ${nameControl}
            <small>${escapeHtml(task.meta || "")}</small>
          </div>
          <div class="gantt-task-col">${task.durationDays}</div>
          <div class="gantt-task-col">${formatDate(task.start)}</div>
          <div class="gantt-task-col">${formatDate(task.end)}</div>
        </div>
      `;
    })
    .join("");

  const timeRows = tasks
    .map((task, index) => {
      const bars = (task.segments || [])
        .map((segment) => {
          const startIndex = indexByDate[segment.start];
          const endIndex = indexByDate[segment.end];
          if (typeof startIndex !== "number" || typeof endIndex !== "number") {
            return "";
          }
          const left = startIndex * dayWidth + 2;
          const width = Math.max(8, (endIndex - startIndex + 1) * dayWidth - 4);
          const label = `${task.name} (${formatDate(segment.start)} - ${formatDate(segment.end)})`;
          return `<span class="gantt-segment" style="left:${left}px; width:${width}px; background:${segment.color};" title="${escapeHtml(
            label,
          )}"></span>`;
        })
        .join("");

      return `
        <div class="gantt-time-row ${index % 2 === 0 ? "even" : "odd"}">
          <div class="gantt-day-grid">${offDayCells}</div>
          <div class="gantt-bars-layer">
            ${bars}
            <span class="gantt-bar-note">${escapeHtml(task.sideNote || "")}</span>
          </div>
        </div>
      `;
    })
    .join("");

  const weeksHeader = weekGroups
    .map((group) => {
      const width = group.count * dayWidth;
      return `<div class="gantt-week-cell" style="width:${width}px">${escapeHtml(group.label)}</div>`;
    })
    .join("");

  const daysHeader = timelineDates
    .map((date, idx) => {
      const info = dayMeta[idx];
      const dayName = ["Ndz", "Pon", "Wt", "Sr", "Czw", "Pt", "Sob"][info.dayIndex];
      const cls = [
        "gantt-day-head",
        info.working ? "" : "off",
        info.isSaturday ? "saturday" : "",
        info.isSunday ? "sunday" : "",
      ]
        .filter(Boolean)
        .join(" ");
      const workMark = info.working ? "R" : "N";
      return `<div class="${cls}" title="${info.working ? "Dzien roboczy" : "Dzien niepracujacy"}"><span>${dayName}</span><strong>${info.dt.getDate()}</strong><em>${workMark}</em></div>`;
    })
    .join("");

  const togglesHeader = timelineDates
    .map((date, idx) => {
      const info = dayMeta[idx];
      const cls = [
        "gantt-work-toggle-cell",
        info.isSaturday ? "saturday" : "",
        info.isSunday ? "sunday" : "",
      ]
        .filter(Boolean)
        .join(" ");
      const checked = info.working ? "checked" : "";
      const disabled = options.editableCalendar ? "" : "disabled";
      const title = `${date}: ${info.working ? "roboczy" : "niepracujacy"}`;
      return `<label class="${cls}" title="${title}"><input type="checkbox" data-gantt-workday-toggle="1" data-date="${date}" ${checked} ${disabled} /></label>`;
    })
    .join("");

  container.innerHTML = `
    <div class="gantt-pro-wrapper">
      ${options.title ? `<p class="gantt-pro-title">${escapeHtml(options.title)}</p>` : ""}
      <div class="gantt-workday-info">
        <span class="legend-item work"><i></i>Dzien roboczy (${workingDaysCount})</span>
        <span class="legend-item off"><i></i>Dzien niepracujacy (${nonWorkingDaysCount})</span>
        <span class="legend-item sat"><i></i>Sobota</span>
        <span class="legend-item sun"><i></i>Niedziela</span>
        <span class="legend-note">Harmonogram omija dni niepracujace i przeskakuje na kolejny roboczy.</span>
      </div>
      <div class="gantt-pro">
        <div class="gantt-left">
          <div class="gantt-task-header">
            <div class="gantt-task-col name">Task Name</div>
            <div class="gantt-task-col">Duration</div>
            <div class="gantt-task-col">Start</div>
            <div class="gantt-task-col">End</div>
          </div>
          <div class="gantt-task-body">${leftRows}</div>
        </div>
        <div class="gantt-right-scroll">
          <div class="gantt-right" style="width:${timelineWidth}px">
            ${options.editableCalendar ? `<div class="gantt-work-toggle-row">${togglesHeader}</div>` : ""}
            <div class="gantt-weeks-row">${weeksHeader}</div>
            <div class="gantt-days-row">${daysHeader}</div>
            <div class="gantt-time-body">${timeRows}</div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function ganttBounds(tasks) {
  const starts = tasks.map((task) => task.start).filter(Boolean);
  const ends = tasks.map((task) => task.end).filter(Boolean);
  if (starts.length === 0 || ends.length === 0) {
    const today = isoDate(new Date());
    return { start: today, end: today };
  }
  return {
    start: starts.sort()[0],
    end: ends.sort()[ends.length - 1],
  };
}

function dateRange(startDate, endDate) {
  const start = toDate(startDate);
  const end = toDate(endDate);
  if (start > end) {
    return [];
  }
  const result = [];
  const cursor = new Date(start);
  let guard = 0;
  while (cursor <= end && guard < 1200) {
    result.push(isoDate(cursor));
    cursor.setDate(cursor.getDate() + 1);
    guard += 1;
  }
  return result;
}

function groupTimelineByWeek(dates) {
  const groups = [];
  let current = null;
  dates.forEach((date) => {
    const weekStart = mondayOf(date);
    const key = isoDate(weekStart);
    if (!current || current.key !== key) {
      current = {
        key,
        label: `Tydz ${formatDate(key)}`,
        count: 0,
      };
      groups.push(current);
    }
    current.count += 1;
  });
  return groups;
}

function mondayOf(dateValue) {
  const date = toDate(dateValue);
  const day = (date.getDay() + 6) % 7;
  date.setDate(date.getDate() - day);
  return date;
}

function buildDateSegments(dates) {
  const sorted = sortedUniqueDates(dates);
  if (sorted.length === 0) {
    return [];
  }
  const segments = [];
  let start = sorted[0];
  let previous = sorted[0];

  for (let index = 1; index < sorted.length; index += 1) {
    const current = sorted[index];
    if (daysBetween(previous, current) === 1) {
      previous = current;
      continue;
    }
    segments.push({ start, end: previous });
    start = current;
    previous = current;
  }

  segments.push({ start, end: previous });
  return segments;
}

function orderStatusColor(status) {
  if (status === "Produkcja") {
    return "#0a7f5a";
  }
  if (status === "Kosmetyka") {
    return "#cc7a00";
  }
  if (status === "Zakonczone") {
    return "#72849d";
  }
  return "#264f7b";
}

function departmentColor(department) {
  if (department === "Maszynownia") {
    return "#0f4c81";
  }
  if (department === "Lakiernia") {
    return "#b31f2e";
  }
  if (department === "Kompletacja") {
    return "#1d8758";
  }
  if (department === "Kosmetyka") {
    return "#c98900";
  }
  return "#3e5e7c";
}

function renderGanttCalendarMatrix(container, orders) {
  if (!container) {
    return;
  }
  const sourceOrders = Array.isArray(orders) ? orders : [];
  const dates = buildGanttTimelineDates(sourceOrders);
  if (dates.length === 0) {
    container.innerHTML = "<p>Brak dni do wyswietlenia.</p>";
    return;
  }

  const workInfo = dates.map((date) => {
    const dt = toDate(date);
    return {
      date,
      dt,
      dayIndex: dt.getDay(),
      saturday: dt.getDay() === 6,
      sunday: dt.getDay() === 0,
      working: isWorkingDay(dt),
    };
  });

  const leftHeader = `
    <th class="cg-left-head sticky-col c1" rowspan="4">Numer zamowienia</th>
    <th class="cg-left-head sticky-col c2" rowspan="4">Obszar</th>
  `;
  const leftColumnsWidth = 480;
  const dayColumnWidth = 78;
  const tableWidth = leftColumnsWidth + dates.length * dayColumnWidth;

  const toggleHeader = workInfo
    .map((item) => {
      const classes = ["cg-toggle-cell"];
      if (item.saturday) {
        classes.push("sat");
      }
      if (item.sunday) {
        classes.push("sun");
      }
      const checked = item.working ? "checked" : "";
      return `<th class="${classes.join(" ")}"><input type="checkbox" data-gantt-workday-toggle="1" data-date="${
        item.date
      }" ${checked} /></th>`;
    })
    .join("");

  const weekdayHeader = workInfo
    .map((item) => {
      const names = ["Ndz", "Pon", "Wt", "Sr", "Czw", "Pt", "Sob"];
      const classes = ["cg-dayname"];
      if (item.saturday) {
        classes.push("sat");
      }
      if (item.sunday) {
        classes.push("sun");
      }
      if (!item.working) {
        classes.push("off");
      }
      return `<th class="${classes.join(" ")}">${names[item.dayIndex]}</th>`;
    })
    .join("");

  const dateHeader = workInfo
    .map((item) => {
      const classes = ["cg-date"];
      if (item.saturday) {
        classes.push("sat");
      }
      if (item.sunday) {
        classes.push("sun");
      }
      if (!item.working) {
        classes.push("off");
      }
      return `<th class="${classes.join(" ")}">${formatDate(item.date)}</th>`;
    })
    .join("");

  const bodyRows = [];
  bodyRows.push(`
    <tr class="cg-spacer-row">
      <td colspan="${dates.length + 2}"></td>
    </tr>
  `);
  if (sourceOrders.length === 0) {
    bodyRows.push(`
      <tr class="cg-empty-row">
        <td class="sticky-col c1">-</td>
        <td class="sticky-col c2">Brak zamowien</td>
        ${renderCalendarCells(dates, {}, "empty", workInfo)}
      </tr>
    `);
  } else {
    sourceOrders.forEach((order) => {
      const orderDaily = order.calculation?.orderDaily || {};
      const expanded = Boolean(ui.ganttExpandedOrders[order.id]);
      const stationRows = buildOrderStationRows(order);
      const symbol = expanded ? "-" : "+";
      const manualBadge = order.manualStartDate
        ? `<span class="cg-manual-badge" title="Reczny start: ${escapeHtml(formatDate(order.manualStartDate))}">R: ${escapeHtml(
            formatDate(order.manualStartDate),
          )}</span>`
        : "";
      const manualPlannedBadge = isPlannedDateEdited(order)
        ? `<span class="cg-manual-badge" title="Recznie edytowana planowana data produkcji">P: ${escapeHtml(
            formatDate(order.plannedProductionDate),
          )}</span>`
        : "";
      const clearManualBtn = order.manualStartDate
        ? `<button type="button" class="cg-clear-manual-btn" data-gantt-clear-manual="${escapeHtml(
            order.id,
          )}" title="Usun reczny start">Reset</button>`
        : "";

      bodyRows.push(`
        <tr class="cg-order-row" data-gantt-order-row="1" data-order-id="${escapeHtml(order.id)}">
          <td class="sticky-col c1">
            <button type="button" class="cg-toggle-btn" data-gantt-toggle-order="${escapeHtml(order.id)}">${symbol}</button>
            <button type="button" class="cg-order-link" data-gantt-order-id="${escapeHtml(order.id)}">${escapeHtml(
              order.orderNumber,
            )}</button>
            ${manualBadge}
            ${manualPlannedBadge}
            ${clearManualBtn}
          </td>
          <td class="sticky-col c2">Zamowienie</td>
          ${renderCalendarCells(dates, orderDaily, "order", workInfo, { dragOrderId: order.id })}
        </tr>
      `);

      if (expanded) {
        if (stationRows.length === 0) {
          bodyRows.push(`
            <tr class="cg-station-row">
              <td class="sticky-col c1"></td>
              <td class="sticky-col c2">Brak stanowisk</td>
              ${renderCalendarCells(dates, {}, "empty", workInfo)}
            </tr>
          `);
        } else {
          stationRows.forEach((row) => {
            bodyRows.push(`
              <tr class="cg-station-row">
                <td class="sticky-col c1"></td>
                <td class="sticky-col c2">${escapeHtml(row.areaLabel)}</td>
                ${renderCalendarCells(dates, row.daily, `station:${row.department}`, workInfo)}
              </tr>
            `);
          });
        }
      }
    });
  }

  container.innerHTML = `
    <div class="calendar-gantt-wrap">
      <div class="calendar-gantt-top-scroll" data-gantt-top-scroll="1">
        <div class="calendar-gantt-top-scroll-inner" style="width:${tableWidth}px"></div>
      </div>
      <div class="calendar-gantt-body" data-gantt-body-scroll="1">
        <table class="calendar-gantt-table">
          <thead>
            <tr>
              ${leftHeader}
              <th class="cg-days-title" colspan="${dates.length}">Dni Wolne/Pracujace</th>
            </tr>
            <tr>${toggleHeader}</tr>
            <tr>${weekdayHeader}</tr>
            <tr>${dateHeader}</tr>
          </thead>
          <tbody>
            ${bodyRows.join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;
  syncCalendarGanttScroll(container);
}

function syncCalendarGanttScroll(container) {
  const topScroll = container.querySelector('[data-gantt-top-scroll="1"]');
  const bodyScroll = container.querySelector('[data-gantt-body-scroll="1"]');
  if (!topScroll || !bodyScroll) {
    return;
  }

  bodyScroll.scrollLeft = Math.max(0, toInt(ui.ganttCalendarScrollLeft));
  bodyScroll.scrollTop = Math.max(0, toInt(ui.ganttCalendarScrollTop));
  topScroll.scrollLeft = bodyScroll.scrollLeft;

  let syncing = false;
  topScroll.addEventListener("scroll", () => {
    if (syncing) {
      return;
    }
    syncing = true;
    bodyScroll.scrollLeft = topScroll.scrollLeft;
    ui.ganttCalendarScrollLeft = bodyScroll.scrollLeft;
    syncing = false;
  });

  bodyScroll.addEventListener("scroll", () => {
    if (syncing) {
      return;
    }
    syncing = true;
    topScroll.scrollLeft = bodyScroll.scrollLeft;
    ui.ganttCalendarScrollLeft = bodyScroll.scrollLeft;
    ui.ganttCalendarScrollTop = bodyScroll.scrollTop;
    syncing = false;
  });
}

function buildGanttTimelineDates(orders) {
  const monday = mondayOf(new Date());
  const daysBack = clamp(toInt(ui.ganttDaysBackToShow), 0, 730);
  const daysForward = clamp(toInt(ui.ganttDaysForwardToShow), 1, 730);
  const startDate = new Date(monday);
  startDate.setDate(startDate.getDate() - daysBack);
  const defaultStart = isoDate(startDate);
  const endDate = new Date(monday);
  endDate.setDate(endDate.getDate() + daysForward - 1);
  const defaultEnd = isoDate(endDate);
  return dateRange(defaultStart, defaultEnd);
}

function buildOrderStationRows(order) {
  const stationDaily = order.calculation?.stationDaily || {};
  const stationIndex = stationOrderIndexMap();
  const departmentPriority = {
    Maszynownia: 1,
    Lakiernia: 2,
    Kompletacja: 3,
    Kosmetyka: 4,
  };
  const output = Object.entries(stationDaily)
    .map(([stationId, daily]) => {
      const hasLoad = Object.values(daily || {}).some((value) => toFloat(value) > 0);
      if (!hasLoad) {
        return null;
      }
      const station = state.stations.find((item) => item.id === stationId);
      const stationName = station ? station.name : stationId;
      const department = station ? station.department : "-";
      return {
        stationId,
        stationName,
        department,
        areaLabel: `${department} / ${stationName}`,
        daily,
      };
    })
    .filter(Boolean)
    .sort((a, b) => {
      const pa = departmentPriority[a.department] || 99;
      const pb = departmentPriority[b.department] || 99;
      const oa = stationIndex[a.stationId] ?? 999999;
      const ob = stationIndex[b.stationId] ?? 999999;
      return pa - pb || oa - ob || a.stationName.localeCompare(b.stationName);
    });
  return output;
}

function renderCalendarCells(dates, dailyMap, profile, dayInfo, options = {}) {
  const values = dates.map((date) => toFloat(dailyMap?.[date] || 0));
  const max = Math.max(...values, 0);
  const dragOrderId = String(options.dragOrderId || "");
  return dates
    .map((date, idx) => {
      const value = toFloat(dailyMap?.[date] || 0);
      const info = dayInfo[idx];
      const classes = ["cg-day-cell"];
      if (info.saturday) {
        classes.push("sat");
      }
      if (info.sunday) {
        classes.push("sun");
      }
      if (!info.working) {
        classes.push("off");
      }
      const color = value > 0 ? calendarCellColor(profile, value, max) : "";
      const style = color ? ` style="background:${color}"` : "";
      const text = value > 0 ? value.toFixed(2).replace(".", ",") : "";
      const dragAttrs = dragOrderId
        ? ` data-gantt-drag-order-cell="${escapeHtml(dragOrderId)}" draggable="true" title="Przeciagnij zamowienie na inny dzien"`
        : "";
      return `<td class="${classes.join(" ")}" data-gantt-date="${date}"${dragAttrs}${style}>${text}</td>`;
    })
    .join("");
}

function calendarCellColor(profile, value, maxValue) {
  if (maxValue <= 0 || value <= 0) {
    return "";
  }
  const intensity = clamp(value / maxValue, 0.2, 1);
  if (profile === "order") {
    return `rgba(66, 133, 244, ${0.14 + 0.55 * intensity})`;
  }
  if (profile.startsWith("station:")) {
    const dept = profile.split(":")[1] || "";
    if (dept === "Maszynownia") {
      return `rgba(255, 82, 82, ${0.2 + 0.65 * intensity})`;
    }
    if (dept === "Lakiernia") {
      return `rgba(255, 193, 7, ${0.2 + 0.65 * intensity})`;
    }
    if (dept === "Kompletacja") {
      return `rgba(181, 255, 55, ${0.2 + 0.65 * intensity})`;
    }
    if (dept === "Kosmetyka") {
      return `rgba(76, 175, 80, ${0.2 + 0.65 * intensity})`;
    }
  }
  return `rgba(66, 133, 244, ${0.1 + 0.45 * intensity})`;
}

function buildRealizationDaily(order) {
  const result = {
    machining: {},
    painting: {},
    assembly: {},
  };
  const rank = departmentRankMap();

  order.positions.forEach((position) => {
    PROCESS_FLOW.forEach((step, index) => {
      const planned = toFloat(position.times?.[step.key] || 0);
      if (planned <= 0) {
        return;
      }
      const fraction = processRealizationFraction(position, step.key, index + 1, rank);
      if (fraction <= 0) {
        return;
      }
      const minutes = planned * fraction;
      const day = normalizeEventDate(position.finishedAt || position.startedAt, order.entryDate);
      result[step.key][day] = (result[step.key][day] || 0) + minutes;
    });
  });
  return result;
}

function processRealizationFraction(position, stepKey, stepRank, rankMap) {
  if ((position.status || "") === "done") {
    return 1;
  }
  const current = rankMap[position.currentDepartmentStatus || "Dokumentacja"] || 0;
  if (current > stepRank) {
    return 1;
  }
  if (current === stepRank) {
    if ((position.status || "") === "in_progress") {
      return 0.5;
    }
    if ((position.status || "") === "done") {
      return 1;
    }
  }
  if ((position.currentDepartmentStatus || "") === "Zakonczone") {
    return 1;
  }
  return 0;
}

function departmentRankMap() {
  return {
    Dokumentacja: 0,
    Maszynownia: 1,
    Lakiernia: 2,
    Kompletacja: 3,
    Kosmetyka: 4,
    Zakonczone: 5,
  };
}

function normalizeEventDate(value, fallbackDate) {
  const raw = String(value || "").trim();
  if (!raw) {
    return normalizeOptionalDate(fallbackDate) || isoDate(new Date());
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }
  const dt = new Date(raw);
  if (Number.isNaN(dt.getTime())) {
    return normalizeOptionalDate(fallbackDate) || isoDate(new Date());
  }
  return isoDate(dt);
}

function renderGantt() {
  const items = operationalOrders()
    .filter(orderMatchesGanttFilters)
    .slice()
    .sort(
      (a, b) =>
        effectiveStartDate(a).localeCompare(effectiveStartDate(b)) || a.orderNumber.localeCompare(b.orderNumber),
    );
  renderGanttCalendarMatrix(el.ganttBoard, items);
}

function renderReport() {
  const selectedDate = normalizeOptionalDate(el.reportDate?.value);
  const stationId = String(el.reportStationSelect?.value || "").trim();
  const station = state.stations.find((item) => item.id === stationId);
  const stationName = station ? station.name : stationId;

  if (!selectedDate || !stationId) {
    el.reportCockpit.innerHTML = `<p>Wybierz dzien i stanowisko.</p>`;
    el.reportSummary.textContent = "Brak danych.";
    return;
  }

  const rows = [];
  const ganttOrdered = operationalOrders()
    .slice()
    .sort(
      (a, b) =>
        effectiveStartDate(a).localeCompare(effectiveStartDate(b)) || a.orderNumber.localeCompare(b.orderNumber),
    );

  ganttOrdered.forEach((order) => {
    const stationDaily = order.calculation?.stationDaily || {};
    const minutes = toFloat(stationDaily?.[stationId]?.[selectedDate] || 0);
    if (minutes <= 0) {
      return;
    }
    rows.push({
      orderNumber: order.orderNumber,
      positionsCount: Array.isArray(order.positions) ? order.positions.length : 0,
      owner: order.owner,
      color: order.color,
      framesCount: Math.max(0, toInt(order.framesCount)),
      sashesCount: Math.max(0, toInt(order.sashesCount)),
      minutes,
    });
  });

  const capacities = dailyCapacitiesByStation(selectedDate);
  const stationCapacity = Math.max(0, toFloat(capacities[stationId] || 0));
  const totalMinutes = rows.reduce((sum, row) => sum + row.minutes, 0);
  const overload = Math.max(0, totalMinutes - stationCapacity);
  el.reportSummary.textContent = `${formatDate(selectedDate)} | ${stationName} | zamowien: ${rows.length} | suma wg Gantt: ${totalMinutes.toFixed(
    0,
  )} min | pojemnosc: ${stationCapacity.toFixed(0)} min | przeciazenie: ${overload.toFixed(0)} min`;

  el.reportCockpit.innerHTML = buildProductionCockpit({
    date: selectedDate,
    station,
    stationCapacity,
    totalMinutes,
    overload,
    rows,
  });
}

function renderExecution() {
  if (!el.executionSummary || !el.executionKpiCards || !el.executionDepartmentBody || !el.executionOrdersBody) {
    return;
  }
  const mode = String(el.executionPeriodMode?.value || "day");
  const anchor = normalizeOptionalDate(el.executionAnchorDate?.value);
  if (!anchor) {
    el.executionSummary.textContent = "Wybierz date bazowa.";
    el.executionKpiCards.innerHTML = "";
    el.executionDepartmentBody.innerHTML = `<tr><td colspan="6">Brak danych.</td></tr>`;
    el.executionOrdersBody.innerHTML = `<tr><td colspan="5">Brak danych.</td></tr>`;
    return;
  }

  const range = executionRangeForSelection(mode, anchor);
  const executionDepartments = ["Maszynownia", "Lakiernia", "Kompletacja", "Kosmetyka"];
  const totalsByDepartment = executionDepartments.reduce((acc, department) => {
    acc[department] = { planned: 0, actual: 0, capacity: 0 };
    return acc;
  }, {});

  const allDates = dateRange(range.start, range.end);
  allDates.forEach((date) => {
    if (!isWorkingDay(toDate(date))) {
      return;
    }
    const dailyCapacityByDepartment = dailyCapacitiesByDepartment(date);
    totalsByDepartment.Maszynownia.capacity += toFloat(dailyCapacityByDepartment.machining);
    totalsByDepartment.Lakiernia.capacity += toFloat(dailyCapacityByDepartment.painting);
    totalsByDepartment.Kompletacja.capacity += toFloat(dailyCapacityByDepartment.assembly);
  });

  const orderRows = [];
  state.orders.forEach((order) => {
    const plannedOrder = sumDailyMapInRange(order?.calculation?.orderDaily || {}, range.start, range.end);
    const realization = buildRealizationDaily(order);
    const actualOrder =
      sumDailyMapInRange(realization.machining, range.start, range.end) +
      sumDailyMapInRange(realization.painting, range.start, range.end) +
      sumDailyMapInRange(realization.assembly, range.start, range.end);

    const processDaily = order?.calculation?.processDaily || {};
    totalsByDepartment.Maszynownia.planned += sumDailyMapInRange(processDaily.machining, range.start, range.end);
    totalsByDepartment.Lakiernia.planned += sumDailyMapInRange(processDaily.painting, range.start, range.end);
    totalsByDepartment.Kompletacja.planned += sumDailyMapInRange(processDaily.assembly, range.start, range.end);

    totalsByDepartment.Maszynownia.actual += sumDailyMapInRange(realization.machining, range.start, range.end);
    totalsByDepartment.Lakiernia.actual += sumDailyMapInRange(realization.painting, range.start, range.end);
    totalsByDepartment.Kompletacja.actual += sumDailyMapInRange(realization.assembly, range.start, range.end);

    if (plannedOrder <= 0 && actualOrder <= 0) {
      return;
    }
    orderRows.push({
      orderNumber: order.orderNumber,
      status: order.orderStatus || "Dokumentacja",
      planned: plannedOrder,
      actual: actualOrder,
      completion: completionPercent(actualOrder, plannedOrder),
    });
  });

  orderRows.sort((a, b) => b.planned - a.planned || b.actual - a.actual || a.orderNumber.localeCompare(b.orderNumber, "pl"));

  const totalPlanned = executionDepartments.reduce((sum, department) => sum + totalsByDepartment[department].planned, 0);
  const totalActual = executionDepartments.reduce((sum, department) => sum + totalsByDepartment[department].actual, 0);
  const totalCapacity = executionDepartments.reduce((sum, department) => sum + totalsByDepartment[department].capacity, 0);
  const completion = completionPercent(totalActual, totalPlanned);
  const efficiency = completionPercent(totalActual, totalCapacity);

  el.executionSummary.textContent = `${range.label} | Zamowienia: ${orderRows.length} | Plan: ${totalPlanned.toFixed(
    0,
  )} min | Wykonanie: ${totalActual.toFixed(0)} min | Realizacja planu: ${completion.toFixed(1)}% | Wydajnosc: ${efficiency.toFixed(1)}%`;

  el.executionKpiCards.innerHTML = [
    executionKpiCard("Plan (min)", totalPlanned.toFixed(0)),
    executionKpiCard("Wykonanie (min)", totalActual.toFixed(0)),
    executionKpiCard("Realizacja planu", `${completion.toFixed(1)}%`),
    executionKpiCard("Pojemnosc", `${totalCapacity.toFixed(0)} min`),
    executionKpiCard("Wydajnosc", `${efficiency.toFixed(1)}%`),
  ].join("");

  el.executionDepartmentBody.innerHTML = executionDepartments
    .map((department) => {
      const row = totalsByDepartment[department];
      const departmentCompletion = completionPercent(row.actual, row.planned);
      const departmentEfficiency = completionPercent(row.actual, row.capacity);
      return `
        <tr>
          <td>${escapeHtml(department)}</td>
          <td>${row.planned.toFixed(0)}</td>
          <td>${row.actual.toFixed(0)}</td>
          <td>${departmentCompletion.toFixed(1)}%</td>
          <td>${row.capacity.toFixed(0)}</td>
          <td>${departmentEfficiency.toFixed(1)}%</td>
        </tr>
      `;
    })
    .join("");

  if (orderRows.length === 0) {
    el.executionOrdersBody.innerHTML = `<tr><td colspan="5">Brak zamowien z planem lub wykonaniem dla wybranego zakresu.</td></tr>`;
    return;
  }

  el.executionOrdersBody.innerHTML = orderRows
    .map((row) => {
      return `
        <tr>
          <td>${escapeHtml(row.orderNumber)}</td>
          <td>${statusPill(row.status)}</td>
          <td>${row.planned.toFixed(0)}</td>
          <td>${row.actual.toFixed(0)}</td>
          <td>${row.completion.toFixed(1)}%</td>
        </tr>
      `;
    })
    .join("");
}

function executionKpiCard(label, value) {
  return `
    <article class="card execution-kpi-card">
      <h3>${escapeHtml(label)}</h3>
      <p>${escapeHtml(value)}</p>
    </article>
  `;
}

function executionRangeForSelection(mode, anchorDate) {
  if (mode === "week") {
    const weekStart = mondayOf(anchorDate);
    const weekEnd = toDate(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    const start = isoDate(weekStart);
    const end = isoDate(weekEnd);
    return {
      start,
      end,
      label: `Tydzien ${isoWeekNumber(anchorDate) || "-"}: ${formatDate(start)} - ${formatDate(end)}`,
    };
  }
  const day = normalizeOptionalDate(anchorDate) || isoDate(new Date());
  return {
    start: day,
    end: day,
    label: `Dzien: ${formatDate(day)}`,
  };
}

function sumDailyMapInRange(dailyMap, startDate, endDate) {
  return Object.entries(dailyMap || {}).reduce((sum, [date, minutes]) => {
    if (date < startDate || date > endDate) {
      return sum;
    }
    return sum + toFloat(minutes);
  }, 0);
}

function completionPercent(actual, base) {
  const denominator = Math.max(0, toFloat(base));
  if (denominator <= 0) {
    return toFloat(actual) > 0 ? 100 : 0;
  }
  return (toFloat(actual) / denominator) * 100;
}

function buildProductionCockpit({ date, station, stationCapacity, totalMinutes, overload, rows }) {
  const settings = state.stationSettings?.[station?.id] || { shiftCount: 2, peopleCount: 1 };
  const shiftCount = clamp(toInt(settings.shiftCount), 1, 3);
  const peopleCount = Math.max(1, toInt(settings.peopleCount));
  const minutesPerShift = Math.max(1, toInt(state.settings?.minutesPerShift));
  const utilization = stationCapacity > 0 ? (totalMinutes / stationCapacity) * 100 : 0;
  const dayNames = ["Niedziela", "Poniedzialek", "Wtorek", "Sroda", "Czwartek", "Piatek", "Sobota"];
  const dayName = dayNames[toDate(date).getDay()] || "-";

  const rowsHtml =
    rows.length > 0
      ? rows
          .map(
            (row, index) => `
            <tr>
              <td>${index + 1}</td>
              <td>${escapeHtml(row.orderNumber)}</td>
              <td>${Math.max(0, toInt(row.positionsCount))}</td>
              <td>${escapeHtml(row.color || "-")}</td>
              <td>${Math.max(0, toInt(row.framesCount))}</td>
              <td>${Math.max(0, toInt(row.sashesCount))}</td>
              <td>${escapeHtml(row.owner || "-")}</td>
              <td>${row.minutes.toFixed(0)}</td>
              <td class="print-notes-cell"></td>
            </tr>`,
          )
          .join("")
      : `<tr><td colspan="9">Brak zaplanowanych zamowien na ten dzien i stanowisko.</td></tr>`;

  return `
    <div class="cockpit-sheet">
      <div class="cockpit-header">
        <div>
          <p class="cockpit-eyebrow">Kokpit Produkcyjny</p>
          <h3>Plan dzienny stanowiska</h3>
        </div>
        <div class="cockpit-meta-right">
          <strong>Data: ${escapeHtml(formatDate(date))}</strong>
          <span>${escapeHtml(dayName)}</span>
        </div>
      </div>

      <div class="cockpit-meta-grid">
        <div class="cockpit-meta-item"><span>Stanowisko</span><strong>${escapeHtml(station?.name || "-")}</strong></div>
        <div class="cockpit-meta-item"><span>Dzial</span><strong>${escapeHtml(station?.department || "-")}</strong></div>
        <div class="cockpit-meta-item"><span>Zmiany</span><strong>${shiftCount}</strong></div>
        <div class="cockpit-meta-item"><span>Osoby</span><strong>${peopleCount}</strong></div>
        <div class="cockpit-meta-item"><span>Min/zmiane</span><strong>${minutesPerShift}</strong></div>
        <div class="cockpit-meta-item"><span>Pojemnosc dnia</span><strong>${stationCapacity.toFixed(0)} min</strong></div>
      </div>

      <div class="cockpit-kpi-strip">
        <div class="cockpit-kpi"><span>Zamowienia</span><strong>${rows.length}</strong></div>
        <div class="cockpit-kpi"><span>Suma minut</span><strong>${totalMinutes.toFixed(0)}</strong></div>
        <div class="cockpit-kpi"><span>Wykorzystanie</span><strong>${utilization.toFixed(1)}%</strong></div>
        <div class="cockpit-kpi ${overload > 0 ? "warn" : "ok"}"><span>Przeciazenie</span><strong>${overload.toFixed(
          0,
        )} min</strong></div>
      </div>

      <div class="table-wrap">
        <table class="cockpit-table">
          <thead>
            <tr>
              <th>Lp</th>
              <th>Zamowienie</th>
              <th>Pozycje</th>
              <th>Kolor</th>
              <th>Ramy</th>
              <th>Skrzydla</th>
              <th>Opracowuje</th>
              <th>Minuty</th>
              <th>Uwagi brygady</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>

      <div class="cockpit-signatures">
        <div><span>Brygadzista:</span><b></b></div>
        <div><span>Operator / Brygada:</span><b></b></div>
        <div><span>Data i godzina wydruku:</span><b>${escapeHtml(formatDate(isoDate(new Date())))} ${escapeHtml(
          new Date().toLocaleTimeString("pl-PL", { hour: "2-digit", minute: "2-digit" }),
        )}</b></div>
      </div>
    </div>
  `;
}

function printReportCockpit() {
  document.body.classList.add("print-report-mode");
  const cleanup = () => {
    document.body.classList.remove("print-report-mode");
    window.removeEventListener("afterprint", cleanup);
  };
  window.addEventListener("afterprint", cleanup);
  window.print();
  window.setTimeout(cleanup, 1200);
}

function renderWorkload() {
  if (state.stations.length === 0) {
    el.workloadTableBody.innerHTML = `<tr><td colspan="3">Brak stanowisk.</td></tr>`;
    return;
  }
  const workload = calculateStationWorkload();
  el.workloadTableBody.innerHTML = state.stations
    .map((station) => {
      return `
        <tr>
          <td>${escapeHtml(station.name)}</td>
          <td>${escapeHtml(station.department)}</td>
          <td>${(workload[station.id] || 0).toFixed(1)}</td>
        </tr>
      `;
    })
    .join("");
}

function renderSkillWorkerForm() {
  if (!el.skillWorkerForm || !el.skillWorkerSkillsWrap) {
    return;
  }
  const departments = availableSkillDepartments();
  const editingWorker = state.skillWorkers.find((item) => item.id === String(ui.editingSkillWorkerId || "").trim()) || null;
  if (!editingWorker && ui.editingSkillWorkerId) {
    ui.editingSkillWorkerId = "";
  }
  const workerDepartment = String(editingWorker?.department || "").trim();
  const currentDepartment = String(el.skillWorkerDepartmentSelect?.value || "").trim();
  let selectedDepartment = departments.includes(currentDepartment) ? currentDepartment : "";
  if (!selectedDepartment) {
    selectedDepartment = departments.includes(workerDepartment) ? workerDepartment : departments[0] || DEPARTMENTS[0];
  }
  const selectedAssignedShift = clamp(
    toInt(el.skillWorkerAssignedShiftSelect?.value || editingWorker?.assignedShift || 1),
    1,
    3,
  );
  if (el.skillWorkerDepartmentSelect) {
    el.skillWorkerDepartmentSelect.innerHTML = departments.map(
      (department) => `<option value="${escapeHtml(department)}">${escapeHtml(department)}</option>`,
    ).join("");
    el.skillWorkerDepartmentSelect.value = selectedDepartment;
  }
  if (el.skillWorkerAssignedShiftSelect) {
    el.skillWorkerAssignedShiftSelect.value = String(selectedAssignedShift);
  }
  const stations = (Array.isArray(state.stations) ? state.stations : []).filter(
    (station) => station?.active !== false && String(station?.department || "").trim() === selectedDepartment,
  );
  const primarySelect = el.skillWorkerPrimaryStationSelect;
  if (primarySelect) {
    const currentPrimary = String(primarySelect.value || "").trim();
    const editingPrimary = String(editingWorker?.primaryStationId || "").trim();
    const stationOptions = stations.map(
      (station) =>
        `<option value="${escapeHtml(station.id)}">${escapeHtml(station.name)} (${escapeHtml(station.id)})</option>`,
    );
    primarySelect.innerHTML = `<option value="">Brak przypisania</option>${stationOptions.join("")}`;
    const validIds = new Set(stations.map((station) => String(station.id || "").trim()));
    let selectedPrimary = "";
    if (editingPrimary && validIds.has(editingPrimary)) {
      selectedPrimary = editingPrimary;
    } else if (currentPrimary && validIds.has(currentPrimary)) {
      selectedPrimary = currentPrimary;
    }
    primarySelect.value = selectedPrimary;
    primarySelect.disabled = stations.length === 0;
  }
  if (editingWorker) {
    if (el.skillWorkerNameInput) {
      el.skillWorkerNameInput.value = editingWorker.name || "";
    }
    if (el.skillWorkerActiveInput) {
      el.skillWorkerActiveInput.checked = editingWorker.active !== false;
    }
    if (el.skillWorkerSubmitBtn) {
      el.skillWorkerSubmitBtn.textContent = "Zapisz pracownika";
    }
    el.cancelSkillWorkerEditBtn?.classList.remove("panel-hidden");
  } else {
    if (el.skillWorkerSubmitBtn) {
      el.skillWorkerSubmitBtn.textContent = "Dodaj pracownika";
    }
    el.cancelSkillWorkerEditBtn?.classList.add("panel-hidden");
  }

  if (stations.length === 0) {
    el.skillWorkerSkillsWrap.innerHTML = `<p>Brak stanowisk przypisanych do dzialu bazowego: <strong>${escapeHtml(
      selectedDepartment,
    )}</strong>.</p>`;
    if (el.skillWorkerSubmitBtn) {
      el.skillWorkerSubmitBtn.disabled = true;
    }
    return;
  }

  const rows = stations
    .map((station) => {
      const level = clamp(toInt(editingWorker?.skills?.[station.id] || 0), 0, 3);
      const activeBadge = station.active === false ? " (nieaktywne)" : "";
      return `
        <label class="skills-level-row">
          <span class="skills-level-station">${escapeHtml(station.name)} (${escapeHtml(station.department)})${escapeHtml(activeBadge)}</span>
          <input
            type="number"
            min="0"
            max="3"
            step="1"
            value="${level}"
            data-skill-level="1"
            data-station-id="${escapeHtml(station.id)}"
            title="0 = brak umiejetnosci, 1 = podstawowy, 2 = samodzielny, 3 = ekspert"
          />
        </label>
      `;
    })
    .join("");
  el.skillWorkerSkillsWrap.innerHTML = `
    <p class="skills-level-hint">Poziom umiejetnosci: 0 (brak), 1 (podstawowy), 2 (samodzielny), 3 (ekspert)</p>
    <div class="skills-level-list">${rows}</div>
  `;
  if (el.skillWorkerSubmitBtn) {
    el.skillWorkerSubmitBtn.disabled = false;
  }
}

function renderSkillWorkersList() {
  if (!el.skillWorkersList) {
    return;
  }
  if (!Array.isArray(state.skillWorkers) || state.skillWorkers.length === 0) {
    el.skillWorkersList.innerHTML = "<p>Brak pracownikow w matrycy umiejetnosci.</p>";
    updateSkillWorkerSelectionStatus();
    return;
  }
  const selected = getSkillWorkerSelectionSet();
  setSkillWorkerSelectionSet(selected);
  const stationNameById = {};
  (state.stations || []).forEach((station) => {
    stationNameById[station.id] = station.name;
  });
  const rows = state.skillWorkers
    .slice()
    .sort((a, b) => {
      const depDiff = String(a.department || "").localeCompare(String(b.department || ""), "pl", { sensitivity: "base" });
      if (depDiff !== 0) {
        return depDiff;
      }
      return String(a.name || "").localeCompare(String(b.name || ""), "pl", { sensitivity: "base" });
    })
    .map((worker) => {
      const entries = Object.entries(worker.skills || {})
        .map(([stationId, level]) => ({
          stationId,
          level: clamp(toInt(level), 0, 3),
          stationName: stationNameById[stationId] || stationId,
        }))
        .filter((item) => item.level > 0)
        .sort((a, b) => b.level - a.level || a.stationName.localeCompare(b.stationName, "pl", { sensitivity: "base" }));
      const summary = entries.length
        ? entries.map((item) => `${item.stationName} (L${item.level})`).join(", ")
        : "Brak stanowisk";
      const primaryStationName = worker.primaryStationId
        ? stationNameById[worker.primaryStationId] || worker.primaryStationId
        : "Brak przypisania";
      const assignedShift = clamp(toInt(worker.assignedShift || 1), 1, 3);
      const activeText = worker.active === false ? "Nieaktywny" : "Aktywny";
      const checkedAttr = selected.has(String(worker.id || "")) ? "checked" : "";
      return `
        <li>
          <label class="inline-checkbox">
            <input type="checkbox" data-action="select-skill-worker" data-worker-id="${escapeHtml(worker.id)}" ${checkedAttr} />
            Zaznacz
          </label>
          <strong>${escapeHtml(worker.name || "-")}</strong>
          <span> | Dzial: ${escapeHtml(worker.department || "-")}</span>
          <span> | Zmiana: ${assignedShift}</span>
          <span> | Stanowisko glowne: ${escapeHtml(primaryStationName)}</span>
          <span> | Status: ${escapeHtml(activeText)}</span>
          <span> | Umiejetnosci: ${escapeHtml(summary)}</span>
          <span class="user-actions">
            <button type="button" data-action="edit-skill-worker" data-worker-id="${escapeHtml(worker.id)}">Edytuj</button>
            <button type="button" data-action="delete-skill-worker" data-worker-id="${escapeHtml(worker.id)}">Usun</button>
          </span>
        </li>
      `;
    })
    .join("");
  el.skillWorkersList.innerHTML = `<ul>${rows}</ul>`;
  updateSkillWorkerSelectionStatus();
}

function ensureSkillsAllocationDefaults() {
  if (el.skillsAllocationMode && !el.skillsAllocationMode.value) {
    el.skillsAllocationMode.value = "day";
  }
  if (el.skillsAllocationAnchorDate && !el.skillsAllocationAnchorDate.value) {
    el.skillsAllocationAnchorDate.value = normalizeOptionalDate(el.reportDate?.value) || isoDate(new Date());
  }
}

function stationDemandForDateByStation(date) {
  const output = {};
  operationalOrders().forEach((order) => {
    const stationDaily = order.calculation?.stationDaily || {};
    Object.entries(stationDaily).forEach(([stationId, daily]) => {
      const minutes = toFloat(daily?.[date] || 0);
      if (minutes <= 0) {
        return;
      }
      output[stationId] = (output[stationId] || 0) + minutes;
    });
  });
  return output;
}

function ensureSkillAvailabilityDefaults() {
  if (el.skillAvailabilityStartDate && !el.skillAvailabilityStartDate.value) {
    el.skillAvailabilityStartDate.value = isoDate(new Date());
  }
  if (el.skillAvailabilityDays) {
    const days = clamp(toInt(el.skillAvailabilityDays.value || 14), 1, 60);
    el.skillAvailabilityDays.value = String(days);
  }
}

function skillAvailabilityMinutes(workerId, date, shift) {
  const worker = (state.skillWorkers || []).find((item) => String(item?.id || "").trim() === String(workerId || "").trim());
  const assignedShift = clamp(toInt(worker?.assignedShift || 1), 1, 3);
  if (shift !== assignedShift) {
    return 0;
  }
  const dayShifts = clamp(toInt(shiftsForDate(date)), 0, 3);
  if (dayShifts <= 0 && !hasAnyStationOvertimeOnDate(date)) {
    return 0;
  }
  const key = String(shift);
  const explicit = state.skillAvailability?.[workerId]?.[date];
  if (explicit && Object.prototype.hasOwnProperty.call(explicit, key)) {
    return clamp(toInt(explicit[key]), 0, 1440);
  }
  return clamp(toInt(state.settings?.minutesPerShift), 0, 1440);
}

function renderSkillAvailabilityControls() {
  if (!el.skillAvailabilityWorkerSelect) {
    return;
  }
  ensureSkillAvailabilityDefaults();
  const workers = (state.skillWorkers || [])
    .filter((worker) => worker.active !== false)
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pl", { sensitivity: "base" }));
  const currentValue = String(el.skillAvailabilityWorkerSelect.value || "");
  el.skillAvailabilityWorkerSelect.innerHTML =
    workers.length === 0
      ? '<option value="">Brak pracownikow</option>'
      : workers.map((worker) => `<option value="${escapeHtml(worker.id)}">${escapeHtml(worker.name)}</option>`).join("");
  if (workers.length === 0) {
    el.skillAvailabilityWorkerSelect.value = "";
    if (el.saveSkillAvailabilityBtn) {
      el.saveSkillAvailabilityBtn.disabled = true;
    }
    if (el.renderSkillAvailabilityBtn) {
      el.renderSkillAvailabilityBtn.disabled = true;
    }
    return;
  }
  const validCurrent = workers.some((worker) => worker.id === currentValue);
  el.skillAvailabilityWorkerSelect.value = validCurrent ? currentValue : workers[0].id;
  if (el.saveSkillAvailabilityBtn) {
    el.saveSkillAvailabilityBtn.disabled = false;
  }
  if (el.renderSkillAvailabilityBtn) {
    el.renderSkillAvailabilityBtn.disabled = false;
  }
}

function availabilityCalendarDates() {
  ensureSkillAvailabilityDefaults();
  const start = normalizeOptionalDate(el.skillAvailabilityStartDate?.value) || isoDate(new Date());
  const days = clamp(toInt(el.skillAvailabilityDays?.value || 14), 1, 60);
  const end = isoDate(addDays(start, days - 1));
  return dateRange(start, end);
}

function renderSkillAvailabilityCalendar() {
  if (!el.skillAvailabilityGridWrap) {
    return;
  }
  const workerId = String(el.skillAvailabilityWorkerSelect?.value || "").trim();
  if (!workerId) {
    el.skillAvailabilityGridWrap.innerHTML = "<p>Brak pracownikow do planowania dostepnosci.</p>";
    if (el.skillAvailabilityStatus) {
      el.skillAvailabilityStatus.textContent = "Brak danych.";
    }
    return;
  }
  const dates = availabilityCalendarDates();
  if (dates.length === 0) {
    el.skillAvailabilityGridWrap.innerHTML = "<p>Nieprawidlowy zakres dat.</p>";
    return;
  }
  const worker = (state.skillWorkers || []).find((item) => String(item?.id || "").trim() === workerId) || null;
  const assignedShift = clamp(toInt(worker?.assignedShift || 1), 1, 3);
  const head = dates
    .map((date) => {
      const dt = toDate(date);
      const weekdayShort = ["Ndz", "Pon", "Wt", "Sr", "Czw", "Pt", "Sob"][dt.getDay()];
      const cls = dt.getDay() === 0 ? "sun" : dt.getDay() === 6 ? "sat" : "";
      return `<th class="${cls}">${escapeHtml(weekdayShort)}<br /><strong>${escapeHtml(formatDate(date))}</strong></th>`;
    })
    .join("");
  const rows = [1, 2, 3]
    .map((shift) => {
      const editableShift = shift === assignedShift;
      const cells = dates
        .map((date) => {
          const value = skillAvailabilityMinutes(workerId, date, shift);
          const readonlyAttr = editableShift ? "" : "disabled";
          return `
            <td>
              <input
                type="number"
                min="0"
                max="1440"
                step="1"
                value="${value}"
                ${readonlyAttr}
                data-skill-availability-cell="1"
                data-worker-id="${escapeHtml(workerId)}"
                data-date="${escapeHtml(date)}"
                data-shift="${shift}"
              />
            </td>
          `;
        })
        .join("");
      const shiftLabel = editableShift ? `Zmiana ${shift}` : `Zmiana ${shift} (nieprzypisana)`;
      return `<tr><th>${shiftLabel}</th>${cells}</tr>`;
    })
    .join("");
  el.skillAvailabilityGridWrap.innerHTML = `
    <div class="table-wrap">
      <table class="skill-availability-table">
        <thead>
          <tr>
            <th>Zmiana</th>
            ${head}
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
  if (el.skillAvailabilityStatus) {
    el.skillAvailabilityStatus.textContent = `Wprowadz minuty dla przypisanej zmiany ${assignedShift} i kliknij Zapisz kalendarz.`;
  }
}

async function saveSkillAvailabilityCalendar() {
  const workerId = String(el.skillAvailabilityWorkerSelect?.value || "").trim();
  if (!workerId) {
    throw new Error("Wybierz pracownika.");
  }
  const dates = availabilityCalendarDates();
  if (dates.length === 0) {
    throw new Error("Nieprawidlowy zakres dat.");
  }
  const entries = Array.from(el.skillAvailabilityGridWrap?.querySelectorAll("input[data-skill-availability-cell='1']") || []).map(
    (input) => {
      const date = String(input.dataset.date || "");
      const shift = clamp(toInt(input.dataset.shift), 1, 3);
      const minutes = clamp(toInt(input.value), 0, 1440);
      input.value = String(minutes);
      return { date, shift, minutes };
    },
  );
  await api("/api/skills/availability/bulk", {
    method: "PUT",
    body: {
      workerId,
      startDate: dates[0],
      endDate: dates[dates.length - 1],
      entries,
    },
  });
  if (el.skillAvailabilityStatus) {
    el.skillAvailabilityStatus.textContent = "Kalendarz dostepnosci zapisany.";
  }
  await reloadAndRender();
}

function stationShiftCapacitiesForDate(station, date) {
  const cfg = state.stationSettings?.[station.id] || { shiftCount: 2, peopleCount: 1 };
  const peopleCount = Math.max(1, toInt(cfg.peopleCount));
  const stationShiftLimit = clamp(toInt(cfg.shiftCount), 1, 3);
  const overtime = clamp(toInt(stationOvertimeMinutesForDate(station.id, date)), 0, 1440);
  const dayShiftLimit = clamp(toInt(shiftsForDate(date)), 0, 3);
  const effectiveShiftCount = dayShiftLimit <= 0 ? 0 : stationShiftLimit;
  const base = Math.max(0, toInt(state.settings?.minutesPerShift)) * peopleCount;
  const output = [];
  if (effectiveShiftCount <= 0) {
    if (overtime > 0) {
      output.push({ shift: 1, capacity: overtime });
    }
    return output;
  }
  for (let shift = 1; shift <= effectiveShiftCount; shift += 1) {
    const capacity = base + (shift === effectiveShiftCount ? overtime : 0);
    output.push({ shift, capacity });
  }
  return output;
}

function buildShiftDemandRowsForDate(date, stations) {
  const demandByStation = stationDemandForDateByStation(date);
  const stationIndex = stationOrderIndexMap();
  const rows = [];
  stations.forEach((station) => {
    let remainingDemand = Math.max(0, toFloat(demandByStation[station.id] || 0));
    const shiftCaps = stationShiftCapacitiesForDate(station, date);
    if (remainingDemand <= 0 && shiftCaps.length === 0) {
      return;
    }
    if (shiftCaps.length === 0 && remainingDemand > 0) {
      rows.push({
        date,
        shift: 1,
        stationId: station.id,
        stationName: station.name,
        department: station.department,
        demand: remainingDemand,
        capacity: 0,
      });
      return;
    }
    shiftCaps.forEach((item) => {
      const shiftDemand = Math.max(0, Math.min(remainingDemand, item.capacity));
      if (shiftDemand > 0 || item.capacity > 0) {
        rows.push({
          date,
          shift: item.shift,
          stationId: station.id,
          stationName: station.name,
          department: station.department,
          demand: shiftDemand,
          capacity: item.capacity,
        });
      }
      remainingDemand = Math.max(0, remainingDemand - shiftDemand);
    });
    if (remainingDemand > 0) {
      const lastShift = shiftCaps.length > 0 ? shiftCaps[shiftCaps.length - 1].shift : 1;
      rows.push({
        date,
        shift: lastShift,
        stationId: station.id,
        stationName: station.name,
        department: station.department,
        demand: remainingDemand,
        capacity: 0,
      });
    }
  });
  rows.sort((a, b) => {
    if (a.date !== b.date) {
      return a.date.localeCompare(b.date);
    }
    if (a.shift !== b.shift) {
      return a.shift - b.shift;
    }
    const demandDiff = b.demand - a.demand;
    if (Math.abs(demandDiff) > 0.0001) {
      return demandDiff;
    }
    return (stationIndex[a.stationId] ?? 999999) - (stationIndex[b.stationId] ?? 999999);
  });
  return rows;
}

function workersSortedForStation(station, candidates) {
  return candidates.sort((a, b) => {
    const sameDepartmentA = String(a.department || "") === String(station.department || "") ? 1 : 0;
    const sameDepartmentB = String(b.department || "") === String(station.department || "") ? 1 : 0;
    if (sameDepartmentA !== sameDepartmentB) {
      return sameDepartmentB - sameDepartmentA;
    }
    const levelDiff = clamp(toInt(b.skills?.[station.id] || 0), 0, 3) - clamp(toInt(a.skills?.[station.id] || 0), 0, 3);
    if (levelDiff !== 0) {
      return levelDiff;
    }
    return String(a.name || "").localeCompare(String(b.name || ""), "pl", { sensitivity: "base" });
  });
}

function buildSkillsAllocationForDate(date, stations, workers) {
  const shiftRows = buildShiftDemandRowsForDate(date, stations);
  const poolsByShift = { 1: {}, 2: {}, 3: {} };
  workers.forEach((worker) => {
    const workerId = String(worker.id || "").trim();
    if (!workerId || worker.active === false) {
      return;
    }
    const assignedShift = clamp(toInt(worker.assignedShift || 1), 1, 3);
    const available = skillAvailabilityMinutes(workerId, date, assignedShift);
    if (available <= 0) {
      return;
    }
    poolsByShift[assignedShift][workerId] = available;
  });

  const workerById = Object.fromEntries(
    workers
      .map((worker) => [String(worker.id || "").trim(), worker])
      .filter(([workerId]) => Boolean(workerId)),
  );
  const stationById = Object.fromEntries(stations.map((station) => [station.id, station]));
  const outputRows = shiftRows.map((row) => ({
    ...row,
    assignedMinutes: 0,
    requiredWorkers: 0,
    assignedWorkers: [],
    missingWorkers: 0,
    suggestions: [],
  }));
  const workerUsageByShift = { 1: {}, 2: {}, 3: {} };

  outputRows.forEach((row) => {
    const shiftPool = poolsByShift[row.shift] || {};
    let remaining = Math.max(0, toFloat(row.demand));
    const station = stationById[row.stationId];
    if (!station || remaining <= 0) {
      return;
    }
    const candidateWorkers = Object.keys(shiftPool)
      .map((workerId) => workerById[workerId])
      .filter((worker) => worker && clamp(toInt(worker.skills?.[row.stationId] || 0), 0, 3) > 0);
    const sorted = workersSortedForStation(station, candidateWorkers);
    sorted.forEach((worker) => {
      if (remaining <= 0) {
        return;
      }
      const workerId = String(worker.id || "").trim();
      const available = Math.max(0, toFloat(shiftPool[workerId] || 0));
      if (available <= 0) {
        return;
      }
      const used = Math.min(available, remaining);
      row.assignedMinutes += used;
      remaining -= used;
      const primaryStationId = String(worker.primaryStationId || "").trim();
      const primaryStationName = primaryStationId ? stationById[primaryStationId]?.name || primaryStationId : "";
      row.assignedWorkers.push({
        workerId,
        name: worker.name,
        level: clamp(toInt(worker.skills?.[row.stationId] || 0), 0, 3),
        minutes: used,
        primaryStationId,
        primaryStationName,
        isPrimary: primaryStationId && primaryStationId === row.stationId,
      });
      const remainingWorkerMinutes = Math.max(0, available - used);
      if (remainingWorkerMinutes > 0.001) {
        shiftPool[workerId] = remainingWorkerMinutes;
      } else {
        delete shiftPool[workerId];
      }
      if (!workerUsageByShift[row.shift][workerId]) {
        workerUsageByShift[row.shift][workerId] = { worker, assignments: [] };
      }
      workerUsageByShift[row.shift][workerId].assignments.push({
        rowRef: row,
        stationId: row.stationId,
        minutes: used,
      });
    });
  });

  [1, 2, 3].forEach((shift) => {
    const usageMap = workerUsageByShift[shift] || {};
    Object.values(usageMap).forEach((usage) => {
      const worker = usage.worker || {};
      const primaryStationId = String(worker.primaryStationId || "").trim();
      if (!primaryStationId) {
        return;
      }
      const primaryStationName = stationById[primaryStationId]?.name || primaryStationId;
      usage.assignments.forEach((assignment) => {
        if (assignment.stationId === primaryStationId) {
          return;
        }
        const suggestionText =
          usage.assignments.length > 1
            ? `${worker.name}: uzupelnienie poza stanowiskiem glownym (${primaryStationName})`
            : `${worker.name}: brak pelnego oblozenia na stanowisku glownym (${primaryStationName})`;
        if (!assignment.rowRef.suggestions.includes(suggestionText)) {
          assignment.rowRef.suggestions.push(suggestionText);
        }
      });
    });
  });

  outputRows.forEach((row) => {
    const perPersonMinutes = Math.max(1, toInt(state.settings?.minutesPerShift));
    row.requiredWorkers = row.demand > 0 ? Math.max(1, Math.ceil(row.demand / perPersonMinutes)) : 0;
    row.missingWorkers = Math.max(0, row.requiredWorkers - row.assignedWorkers.length);
  });

  return {
    rows: outputRows.filter((row) => row.demand > 0 || row.assignedWorkers.length > 0),
  };
}

function assignmentLabelForRow(row) {
  if (!Array.isArray(row.assignedWorkers) || row.assignedWorkers.length === 0) {
    return "-";
  }
  return row.assignedWorkers
    .map((item) => {
      const crossText = item.isPrimary ? "" : item.primaryStationName ? `, uzup. z: ${item.primaryStationName}` : "";
      return `${item.name} (L${item.level}, ${item.minutes.toFixed(0)} min${crossText})`;
    })
    .join(", ");
}

function renderSkillsAllocation() {
  if (!el.skillsAllocationSummary || !el.skillsAllocationTableBody) {
    return;
  }
  const mode = String(el.skillsAllocationMode?.value || "day");
  const anchor = normalizeOptionalDate(el.skillsAllocationAnchorDate?.value);
  if (!anchor) {
    el.skillsAllocationSummary.textContent = "Wybierz date bazowa.";
    el.skillsAllocationTableBody.innerHTML = `<tr><td colspan="12">Brak danych.</td></tr>`;
    return;
  }
  const activeWorkers = (state.skillWorkers || []).filter((worker) => worker.active !== false);
  const activeStations = (state.stations || []).filter((station) => station.active !== false);
  if (activeStations.length === 0) {
    el.skillsAllocationSummary.textContent = "Brak aktywnych stanowisk.";
    el.skillsAllocationTableBody.innerHTML = `<tr><td colspan="12">Dodaj i aktywuj stanowiska w ustawieniach.</td></tr>`;
    return;
  }
  if (activeWorkers.length === 0) {
    const range = executionRangeForSelection(mode, anchor);
    el.skillsAllocationSummary.textContent = `${range.label} | Brak aktywnych pracownikow w matrycy.`;
    el.skillsAllocationTableBody.innerHTML = `<tr><td colspan="12">Dodaj pracownikow i umiejetnosci, aby wygenerowac przydzial.</td></tr>`;
    return;
  }

  const range = executionRangeForSelection(mode, anchor);
  const dates = dateRange(range.start, range.end);
  const rows = [];
  let totalDemand = 0;
  let totalAssignedMinutes = 0;
  let totalRequired = 0;
  let totalAssignedWorkers = 0;

  dates.forEach((date) => {
    const dayPlan = buildSkillsAllocationForDate(date, activeStations, activeWorkers);
    rows.push(...dayPlan.rows);
  });

  rows.forEach((row) => {
    totalDemand += toFloat(row.demand);
    totalAssignedMinutes += toFloat(row.assignedMinutes);
    totalRequired += toInt(row.requiredWorkers);
    totalAssignedWorkers += Array.isArray(row.assignedWorkers) ? row.assignedWorkers.length : 0;
  });

  const totalMissing = Math.max(0, totalRequired - totalAssignedWorkers);
  const totalCoverage = totalRequired > 0 ? (totalAssignedWorkers / totalRequired) * 100 : 100;
  const minuteCoverage = totalDemand > 0 ? (totalAssignedMinutes / totalDemand) * 100 : 100;
  el.skillsAllocationSummary.textContent = `${range.label} | Zapotrzebowanie: ${totalDemand.toFixed(
    0,
  )} min | Przydzielone minuty: ${totalAssignedMinutes.toFixed(0)} | Pokrycie minut: ${minuteCoverage.toFixed(
    1,
  )}% | Potrzeba: ${totalRequired} os. | Przydzielono: ${totalAssignedWorkers} os. | Braki: ${totalMissing} os. | Pokrycie obsady: ${totalCoverage.toFixed(
    1,
  )}%`;

  if (rows.length === 0) {
    el.skillsAllocationTableBody.innerHTML = `<tr><td colspan="12">Brak zapotrzebowania na stanowiska w wybranym zakresie.</td></tr>`;
    return;
  }
  el.skillsAllocationTableBody.innerHTML = rows
    .map((row) => {
      const coverage = row.requiredWorkers > 0 ? (row.assignedWorkers.length / row.requiredWorkers) * 100 : 100;
      return `
        <tr>
          <td>${escapeHtml(formatDate(row.date))}</td>
          <td>${row.shift}</td>
          <td>${escapeHtml(row.stationName)}</td>
          <td>${escapeHtml(row.department)}</td>
          <td>${row.demand.toFixed(0)}</td>
          <td>${row.assignedMinutes.toFixed(0)}</td>
          <td>${row.requiredWorkers}</td>
          <td>${row.assignedWorkers.length}</td>
          <td>${escapeHtml(assignmentLabelForRow(row))}</td>
          <td>${escapeHtml((row.suggestions || []).join(" | ") || "-")}</td>
          <td>${coverage.toFixed(1)}%</td>
          <td>${row.missingWorkers}</td>
        </tr>
      `;
    })
    .join("");
}

function renderFeedbackPositions() {
  const order = state.orders.find((item) => item.id === el.feedbackOrderSelect.value);
  if (!order) {
    el.feedbackPositions.innerHTML = "<p>Wybierz zamowienie.</p>";
    if (el.feedbackSelectionInfo) {
      el.feedbackSelectionInfo.textContent = "Brak danych.";
    }
    applyFeedbackActionState();
    return;
  }
  if (order.positions.length === 0) {
    el.feedbackPositions.innerHTML = "<p>Zamowienie nie ma jeszcze pozycji.</p>";
    if (el.feedbackSelectionInfo) {
      el.feedbackSelectionInfo.textContent = "Brak pozycji.";
    }
    setFeedbackSelectionSet(order.id, new Set());
    applyFeedbackActionState();
    return;
  }
  const allIds = new Set(order.positions.map((position) => String(position.id)));
  const selected = getFeedbackSelectionSet(order.id);
  Array.from(selected).forEach((positionId) => {
    if (!allIds.has(positionId)) {
      selected.delete(positionId);
    }
  });
  setFeedbackSelectionSet(order.id, selected);

  const search = normalizeSearchText(el.feedbackSearchInput?.value).trim();
  const visible = order.positions.filter((position) => {
    if (!search) {
      return true;
    }
    const haystack = normalizeSearchText(
      [
        position.positionNumber,
        position.technology,
        position.line,
        position.currentDepartmentStatus,
        statusLabel(position.status),
        position.width,
        position.height,
        positionProgressPercent(position),
      ].join(" "),
    );
    return haystack.includes(search);
  });

  if (visible.length === 0) {
    el.feedbackPositions.innerHTML = "<p>Brak pozycji dla aktualnego filtra.</p>";
    if (el.feedbackSelectionInfo) {
      el.feedbackSelectionInfo.textContent = `Widoczne: 0 | Zaznaczone: ${selected.size} / ${order.positions.length}`;
    }
    applyFeedbackActionState();
    return;
  }

  el.feedbackPositions.innerHTML = "";
  visible.forEach((position) => {
    const node = el.feedbackItemTemplate.content.firstElementChild.cloneNode(true);
    node.dataset.positionId = String(position.id);
    const checkbox = node.querySelector("input[data-feedback-position-id]");
    const tag = node.querySelector(".tag");
    const meta = node.querySelector(".meta");
    const departmentSelect = node.querySelector('[data-feedback-field="department"]');
    const workflowSelect = node.querySelector('[data-feedback-field="workflow"]');
    const progressInput = node.querySelector('[data-feedback-field="progress"]');
    const saveBtn = node.querySelector('button[data-action="save-feedback-position"]');
    checkbox.value = String(position.id);
    checkbox.checked = selected.has(String(position.id));
    tag.textContent = `Pozycja ${position.positionNumber}`;
    const progress = positionProgressPercent(position);
    meta.textContent = `${position.width}x${position.height} | ${position.technology} | Dzial: ${
      position.currentDepartmentStatus || "-"
    } | Status: ${statusLabel(position.status)} | Postep: ${progress}%`;
    fillSelect(
      departmentSelect,
      POSITION_DEPARTMENT_STATUSES,
      String(position.currentDepartmentStatus || POSITION_DEPARTMENT_STATUSES[0]),
    );
    fillSelect(
      workflowSelect,
      POSITION_WORKFLOW_STATUS_OPTIONS,
      normalizeWorkflowStatusValue(position.status, "pending"),
    );
    if (progressInput) {
      progressInput.value = String(progress);
    }
    if (saveBtn) {
      saveBtn.dataset.positionId = String(position.id);
    }
    refreshFeedbackQuickButtons(node);
    el.feedbackPositions.append(node);
  });

  if (el.feedbackSelectionInfo) {
    el.feedbackSelectionInfo.textContent = `Widoczne: ${visible.length} | Zaznaczone: ${selected.size} / ${order.positions.length}`;
  }
  applyFeedbackActionState();
}

function applyFeedbackActionState() {
  const selectedCount = getSelectedFeedbackPositionIds().length;
  const visibleCount = Array.from(el.feedbackPositions.querySelectorAll("input[data-feedback-position-id]")).length;
  if (el.selectAllPositionsBtn) {
    el.selectAllPositionsBtn.disabled = visibleCount === 0;
  }
  if (el.clearSelectedPositionsBtn) {
    el.clearSelectedPositionsBtn.disabled = selectedCount === 0;
  }
  if (el.startPositionsBtn) {
    el.startPositionsBtn.disabled = selectedCount === 0;
  }
  if (el.finishPositionsBtn) {
    el.finishPositionsBtn.disabled = selectedCount === 0;
  }
  if (el.setFeedbackDepartmentBtn) {
    el.setFeedbackDepartmentBtn.disabled = selectedCount === 0;
  }
  if (el.applyFeedbackChangesBtn) {
    el.applyFeedbackChangesBtn.disabled = selectedCount === 0;
  }
}

function renderUsers() {
  const admin = isAdminUser();
  if (el.usersAdminHint) {
    el.usersAdminHint.classList.toggle("panel-hidden", admin);
  }
  if (!admin && ui.editingUserId) {
    resetUserFormState();
  }
  Array.from(el.userForm?.querySelectorAll("input, select, button") || []).forEach((node) => {
    if (!node) {
      return;
    }
    node.disabled = !admin;
  });
  syncUserFormMode();

  if (!Array.isArray(state.users) || state.users.length === 0) {
    el.usersList.innerHTML = "<p>Brak uzytkownikow.</p>";
    return;
  }
  const rows = state.users
    .map((user) => {
      const isUserAdmin = String(user.role || "user").toLowerCase() === "admin";
      const role = isUserAdmin ? "Admin" : "Uzytkownik";
      const createDbPermission = isUserAdmin || user.canCreateDatabases ? "TAK" : "NIE";
      const loginKey = normalizeUserLoginKey(user.login);
      const hasExplicitMap = Object.prototype.hasOwnProperty.call(state.databaseAccessMap || {}, loginKey);
      const allowedDbCount = allowedDatabaseKeysForUser(user).length;
      const allDbCount = databaseCatalogSource().length;
      const dbAccessLabel = isUserAdmin
        ? "Wszystkie bazy (admin)"
        : hasExplicitMap
          ? `${allowedDbCount}/${allDbCount} przypisane`
          : "Wszystkie (domyslnie)";
      const sections =
        isUserAdmin
          ? "Wszystkie"
          : normalizeVisibleSections(user.visibleSections)
              .map((section) => SECTION_LABELS[section] || section)
              .join(", ") || "-";
      const actions = admin
        ? `<span class="user-actions">
            <button type="button" data-action="edit-user" data-user-id="${escapeHtml(user.id)}">Edytuj</button>
            ${
              isUserAdmin
                ? '<button type="button" disabled title="Konto administratora nie moze byc usuniete.">Usun</button>'
                : `<button type="button" data-action="delete-user" data-user-id="${escapeHtml(user.id)}">Usun</button>`
            }
          </span>`
        : "";
      return `
        <li>
          <strong>${escapeHtml(user.name || "-")}</strong>
          <span> | Login: ${escapeHtml(user.login || "-")}</span>
          <span> | Dzial: ${escapeHtml(user.department || "-")}</span>
          <span> | Rola: ${escapeHtml(role)}</span>
          <span> | Dostep do baz: ${escapeHtml(dbAccessLabel)}</span>
          <span> | Tworzenie baz: ${escapeHtml(createDbPermission)}</span>
          <span> | Widoki: ${escapeHtml(sections)}</span>
          ${actions}
        </li>
      `;
    })
    .join("");
  el.usersList.innerHTML = `<ul>${rows}</ul>`;
}

function onDatabaseAccessMatrixChange(event) {
  const checkbox = event.target.closest("input[data-db-access-checkbox='1']");
  if (!checkbox) {
    return;
  }
  if (el.databaseAccessStatusText) {
    el.databaseAccessStatusText.textContent = "Niezapisane zmiany przypisan baz.";
  }
}

function collectDatabaseAccessMapFromMatrix() {
  const output = {};
  if (!el.databaseAccessMatrixWrap) {
    return output;
  }
  const rows = Array.from(el.databaseAccessMatrixWrap.querySelectorAll("tr[data-db-access-login]"));
  rows.forEach((row) => {
    const loginKey = normalizeUserLoginKey(row.dataset.dbAccessLogin || "");
    const role = String(row.dataset.role || "user").toLowerCase();
    if (!loginKey || role === "admin") {
      return;
    }
    const checkedKeys = Array.from(row.querySelectorAll("input[data-db-access-checkbox='1']:checked"))
      .map((node) => String(node.dataset.dbAccessKey || "").trim())
      .filter((item) => Boolean(item));
    output[loginKey] = checkedKeys;
  });
  return output;
}

async function saveDatabaseAccessMatrix() {
  if (!isAdminUser()) {
    throw new Error("Tylko administrator moze zapisywac przypisania baz.");
  }
  const payload = await api("/api/database-access", {
    method: "PUT",
    body: {
      databaseAccessMap: collectDatabaseAccessMapFromMatrix(),
    },
  });
  state.databaseAccessMap = normalizeDatabaseAccessMap(payload.databaseAccessMap || {});
  state.databases = Array.isArray(payload.databases) ? payload.databases : state.databases;
  state.activeDatabaseKey = String(payload.activeDatabase || state.activeDatabaseKey || "default");
  syncLoginDatabaseSelect();
  renderUsers();
  renderDatabaseAccessMatrix();
  renderDatabaseManager();
  if (el.databaseAccessStatusText) {
    el.databaseAccessStatusText.textContent = "Przypisania baz zostaly zapisane.";
  }
}

function renderDatabaseAccessMatrix() {
  if (!el.databaseAccessMatrixWrap) {
    return;
  }
  if (!isLoggedIn()) {
    el.databaseAccessMatrixWrap.innerHTML = "";
    if (el.saveDatabaseAccessBtn) {
      el.saveDatabaseAccessBtn.disabled = true;
    }
    if (el.databaseAccessStatusText) {
      el.databaseAccessStatusText.textContent = "";
    }
    return;
  }
  const admin = isAdminUser();
  if (!admin) {
    el.databaseAccessMatrixWrap.innerHTML = "<p>Tylko administrator moze edytowac przypisania uzytkownikow do baz.</p>";
    if (el.saveDatabaseAccessBtn) {
      el.saveDatabaseAccessBtn.disabled = true;
    }
    if (el.databaseAccessStatusText) {
      el.databaseAccessStatusText.textContent = "";
    }
    return;
  }

  const users = Array.isArray(state.users)
    ? state.users.filter((user) => String(user?.login || "").trim().length > 0)
    : [];
  const databases = databaseCatalogSource();
  if (users.length === 0) {
    el.databaseAccessMatrixWrap.innerHTML = "<p>Brak uzytkownikow z loginem.</p>";
    if (el.saveDatabaseAccessBtn) {
      el.saveDatabaseAccessBtn.disabled = true;
    }
    return;
  }
  if (databases.length === 0) {
    el.databaseAccessMatrixWrap.innerHTML = "<p>Brak dostepnych baz.</p>";
    if (el.saveDatabaseAccessBtn) {
      el.saveDatabaseAccessBtn.disabled = true;
    }
    return;
  }

  const head = databases
    .map((dbItem) => `<th>${escapeHtml(dbItem.name || dbItem.key || "-")}</th>`)
    .join("");
  const rows = users
    .map((user) => {
      const isUserAdmin = String(user.role || "user").toLowerCase() === "admin";
      const loginKey = normalizeUserLoginKey(user.login);
      const allowedSet = new Set(allowedDatabaseKeysForUser(user));
      const checkboxes = databases
        .map((dbItem) => {
          const dbKey = String(dbItem.key || "");
          const checked = allowedSet.has(dbKey);
          const disabled = isUserAdmin ? "disabled" : "";
          return `<td class="db-access-cell">
              <input
                type="checkbox"
                data-db-access-checkbox="1"
                data-db-access-login="${escapeHtml(loginKey)}"
                data-db-access-key="${escapeHtml(dbKey)}"
                ${checked ? "checked" : ""}
                ${disabled}
              />
            </td>`;
        })
        .join("");
      const roleLabel = isUserAdmin ? "Admin" : "Uzytkownik";
      return `<tr data-db-access-login="${escapeHtml(loginKey)}" data-role="${escapeHtml(String(user.role || "user"))}">
          <th class="db-access-user-col">
            <div>${escapeHtml(user.name || "-")}</div>
            <small>${escapeHtml(user.login || "-")} | ${escapeHtml(roleLabel)}</small>
          </th>
          ${checkboxes}
        </tr>`;
    })
    .join("");

  el.databaseAccessMatrixWrap.innerHTML = `
    <div class="table-wrap database-access-wrap">
      <table class="database-access-table">
        <thead>
          <tr>
            <th class="db-access-user-col">Uzytkownik</th>
            ${head}
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
  if (el.saveDatabaseAccessBtn) {
    el.saveDatabaseAccessBtn.disabled = false;
  }
  if (el.databaseAccessStatusText && !el.databaseAccessStatusText.textContent.trim()) {
    el.databaseAccessStatusText.textContent = "Admin ma zawsze dostep do wszystkich baz.";
  }
}

function setUserPermissionsSelection(selected) {
  const selectedSet = new Set(normalizeVisibleSections(Array.isArray(selected) ? selected : []));
  Array.from(el.userForm?.querySelectorAll('input[name="visibleSection"]') || []).forEach((input) => {
    input.checked = selectedSet.has(String(input.value || ""));
  });
}

function syncUserPermissionsEnabledState() {
  const role = String(el.userRoleSelect?.value || "user").toLowerCase();
  const admin = role === "admin";
  const inputs = Array.from(el.userForm?.querySelectorAll('input[name="visibleSection"]') || []);
  const canCreateDbInput = el.userForm?.querySelector('input[name="canCreateDatabases"]');
  inputs.forEach((input) => {
    input.disabled = admin;
    if (admin) {
      input.checked = true;
    }
  });
  if (canCreateDbInput) {
    canCreateDbInput.disabled = admin;
    if (admin) {
      canCreateDbInput.checked = true;
    }
  }
}

function renderDatabaseManager() {
  if (!el.databaseSelect || !el.switchDatabaseBtn || !el.createDatabaseBtn || !el.newDatabaseNameInput) {
    return;
  }
  const canSwitch = isLoggedIn();
  const canCreate = canCreateDatabaseVariants();
  if (el.databaseAdminHint) {
    el.databaseAdminHint.classList.toggle("panel-hidden", canCreate);
  }

  const databaseItems = Array.isArray(state.databases) && state.databases.length > 0 ? state.databases : [];
  const fallback = [{ key: "default", name: "Domyslna baza", fileName: "planner.db", active: true, variant: false }];
  const source = databaseItems.length > 0 ? databaseItems : fallback;
  el.databaseSelect.innerHTML = source
    .map((item) => {
      const key = String(item.key || "");
      const activeBadge = key === state.activeDatabaseKey ? " (aktywna)" : "";
      const label = `${item.name || key}${activeBadge} - ${item.fileName || ""}`;
      return `<option value="${escapeHtml(key)}">${escapeHtml(label)}</option>`;
    })
    .join("");
  const valid = source.some((item) => String(item.key || "") === state.activeDatabaseKey);
  el.databaseSelect.value = valid ? state.activeDatabaseKey : String(source[0]?.key || "default");

  el.databaseSelect.disabled = !canSwitch;
  el.switchDatabaseBtn.disabled = !canSwitch;
  el.newDatabaseNameInput.disabled = !canCreate;
  el.createDatabaseBtn.disabled = !canCreate;
  if (el.databaseStatusText && !canCreate) {
    el.databaseStatusText.textContent = "Mozesz przelaczac baze. Tworzenie baz jest dostepne dla admina lub wyznaczonego uzytkownika.";
  } else if (el.databaseStatusText && canCreate) {
    if (el.databaseStatusText.textContent.includes("Tworzenie baz jest dostepne")) {
      el.databaseStatusText.textContent = "";
    }
  }
}

function syncSettingsForm() {
  el.settingsForm.minutesPerShift.value = state.settings.minutesPerShift;
  Array.from(el.settingsForm.querySelectorAll('input[name="weekdayShift"]')).forEach((input) => {
    const day = clamp(toInt(input.dataset.day), 0, 6);
    const value = state.settings.weekdayShifts?.[day] ?? state.settings.weekdayShifts?.[String(day)] ?? 0;
    input.value = clamp(toInt(value), 0, 3);
  });
}

function renderStationOvertimeEditor() {
  if (!el.overtimeStationSelect || !el.stationOvertimeTableBody) {
    return;
  }
  const activeStations = (state.stations || []).filter((item) => item?.active !== false);
  const stationItems = activeStations.length > 0 ? activeStations : state.stations || [];
  const hasStations = stationItems.length > 0;
  const stationOptions = stationItems
    .map((station) => {
      const label = `${station.name} (${station.id})`;
      return `<option value="${escapeHtml(station.id)}">${escapeHtml(label)}</option>`;
    })
    .join("");
  const currentStationValue = String(el.overtimeStationSelect.value || "");
  el.overtimeStationSelect.innerHTML = stationOptions || '<option value="">Brak stanowisk</option>';
  el.overtimeStationSelect.disabled = !hasStations;
  if (el.addStationOvertimeBtn) {
    el.addStationOvertimeBtn.disabled = !hasStations;
  }
  if (currentStationValue && Array.from(el.overtimeStationSelect.options).some((opt) => opt.value === currentStationValue)) {
    el.overtimeStationSelect.value = currentStationValue;
  }
  if (el.overtimeStationSelect.options.length > 0 && !el.overtimeStationSelect.value) {
    el.overtimeStationSelect.value = el.overtimeStationSelect.options[0].value;
  }
  if (el.overtimeDateInput && !el.overtimeDateInput.value) {
    el.overtimeDateInput.value = isoDate(new Date());
  }
  if (el.overtimeMinutesInput) {
    const current = clamp(toInt(el.overtimeMinutesInput.value), 1, 1440);
    el.overtimeMinutesInput.value = String(current || 60);
  }

  const stationNameById = {};
  (state.stations || []).forEach((station) => {
    stationNameById[station.id] = station.name;
  });
  const flat = [];
  const overtime = state.settings?.stationOvertime || {};
  Object.entries(overtime).forEach(([date, dayMap]) => {
    Object.entries(dayMap || {}).forEach(([stationId, minutes]) => {
      flat.push({
        date,
        stationId,
        stationName: stationNameById[stationId] || stationId,
        minutes: Math.max(0, toInt(minutes)),
      });
    });
  });
  flat.sort(
    (a, b) => a.date.localeCompare(b.date) || a.stationName.localeCompare(b.stationName, "pl", { sensitivity: "base" }),
  );
  if (flat.length === 0) {
    el.stationOvertimeTableBody.innerHTML = `<tr><td colspan="4">Brak nadgodzin dla stanowisk.</td></tr>`;
    return;
  }
  el.stationOvertimeTableBody.innerHTML = flat
    .map(
      (item) => `
        <tr>
          <td>${escapeHtml(formatDate(item.date))}</td>
          <td>${escapeHtml(item.stationName)} (${escapeHtml(item.stationId)})</td>
          <td>${escapeHtml(String(item.minutes))}</td>
          <td>
            <button type="button" data-overtime-remove="1" data-date="${escapeHtml(item.date)}" data-station-id="${escapeHtml(
              item.stationId,
            )}">Usun</button>
          </td>
        </tr>
      `,
    )
    .join("");
}

function renderStationSettingsEditor() {
  if (state.stations.length === 0) {
    el.stationSettingsTableBody.innerHTML = `<tr><td colspan="7">Brak stanowisk. Uzyj przycisku "Dodaj stanowisko".</td></tr>`;
    return;
  }
  el.stationSettingsTableBody.innerHTML = state.stations
    .map((station) => {
      const cfg = state.stationSettings[station.id] || { shiftCount: 2, peopleCount: 1 };
      return makeStationEditorRowHtml({
        id: station.id,
        name: station.name,
        department: station.department,
        active: station.active !== false,
        shiftCount: cfg.shiftCount,
        peopleCount: cfg.peopleCount,
      });
    })
    .join("");
}

function makeStationEditorRowHtml(station) {
  return `
    <tr data-station-row="1">
      <td><input data-kind="stationId" value="${escapeHtml(station.id || "")}" ${station.id ? "readonly" : ""} placeholder="AUTO" /></td>
      <td><input data-kind="name" value="${escapeHtml(station.name || "")}" placeholder="Nazwa stanowiska" /></td>
      <td>
        <select data-kind="department">
          ${DEPARTMENTS.map(
            (department) =>
              `<option value="${escapeHtml(department)}" ${
                department === (station.department || "Maszynownia") ? "selected" : ""
              }>${escapeHtml(department)}</option>`,
          ).join("")}
        </select>
      </td>
      <td><input data-kind="active" type="checkbox" ${station.active === false ? "" : "checked"} /></td>
      <td><input data-kind="shift" type="number" min="1" max="3" step="1" value="${clamp(toInt(station.shiftCount), 1, 3)}" /></td>
      <td><input data-kind="people" type="number" min="1" step="1" value="${Math.max(1, toInt(station.peopleCount))}" /></td>
      <td class="station-row-actions">
        <button type="button" data-action="move-up-station-row">W gore</button>
        <button type="button" data-action="move-down-station-row">W dol</button>
        <button type="button" data-action="delete-station-row">Usun</button>
      </td>
    </tr>
  `;
}

function addStationEditorRow() {
  const emptyRow = makeStationEditorRowHtml({
    id: "",
    name: "",
    department: "Maszynownia",
    active: true,
    shiftCount: 2,
    peopleCount: 1,
  });
  const emptyPlaceholder = el.stationSettingsTableBody.querySelector("tr td[colspan='7']");
  if (emptyPlaceholder) {
    el.stationSettingsTableBody.innerHTML = "";
  }
  el.stationSettingsTableBody.insertAdjacentHTML("beforeend", emptyRow);
  const newRows = Array.from(el.stationSettingsTableBody.querySelectorAll("tr[data-station-row='1']"));
  const last = newRows[newRows.length - 1];
  last?.querySelector('input[data-kind="name"]')?.focus();
}

function onStationSettingsTableClick(event) {
  const moveUpBtn = event.target.closest("button[data-action='move-up-station-row']");
  if (moveUpBtn) {
    const row = moveUpBtn.closest("tr[data-station-row='1']");
    if (!row) {
      return;
    }
    const prev = row.previousElementSibling;
    if (prev && prev.matches("tr[data-station-row='1']")) {
      row.parentElement?.insertBefore(row, prev);
    }
    return;
  }

  const moveDownBtn = event.target.closest("button[data-action='move-down-station-row']");
  if (moveDownBtn) {
    const row = moveDownBtn.closest("tr[data-station-row='1']");
    if (!row) {
      return;
    }
    const next = row.nextElementSibling;
    if (next && next.matches("tr[data-station-row='1']")) {
      row.parentElement?.insertBefore(next, row);
    }
    return;
  }

  const deleteBtn = event.target.closest("button[data-action='delete-station-row']");
  if (!deleteBtn) {
    return;
  }
  const row = deleteBtn.closest("tr[data-station-row='1']");
  if (!row) {
    return;
  }
  row.remove();
  const rowsLeft = el.stationSettingsTableBody.querySelectorAll("tr[data-station-row='1']").length;
  if (rowsLeft === 0) {
    el.stationSettingsTableBody.innerHTML = `<tr><td colspan="7">Brak stanowisk. Uzyj przycisku "Dodaj stanowisko".</td></tr>`;
  }
}

function renderMaterialRulesEditor() {
  const rows = MATERIALS.map((material) => {
    const cols = DEPARTMENTS.map((department) => {
      const checked = (state.materialRules[department] || []).includes(material.key) ? "checked" : "";
      return `<td><input type="checkbox" data-material="${material.key}" data-department="${department}" ${checked} /></td>`;
    }).join("");
    return `<tr><td>${material.label}</td>${cols}</tr>`;
  }).join("");
  el.materialRulesTableBody.innerHTML = rows;
}

function renderTechnologyAllocationEditor() {
  const technology = ui.selectedTechnology || el.technologyEditorSelect.value;
  const process = ui.selectedProcess || el.processEditorSelect.value;
  if (!technology || !process) {
    el.technologyAllocationTableBody.innerHTML = `<tr><td colspan="3">Brak danych.</td></tr>`;
    return;
  }
  const department = processToDepartment(process);
  const allocations = (state.technologies[technology] && state.technologies[technology][process]) || {};
  const rows = state.stations.filter((station) => station.department === department && station.active !== false);
  if (rows.length === 0) {
    el.technologyAllocationTableBody.innerHTML = `<tr><td colspan="3">Brak aktywnych stanowisk dla procesu.</td></tr>`;
    return;
  }
  el.technologyAllocationTableBody.innerHTML = rows
    .map((station) => {
      const value = toFloat(allocations[station.id]);
      return `
        <tr>
          <td>${escapeHtml(station.name)}</td>
          <td>${escapeHtml(station.department)}</td>
          <td><input type="number" min="0" step="0.01" data-station-id="${station.id}" value="${value.toFixed(2)}" /></td>
        </tr>
      `;
    })
    .join("");
}

function recalculateOrders() {
  const stationUsage = {};
  const departmentUsage = {
    machining: {},
    painting: {},
    assembly: {},
  };
  const ordered = state.orders
    .slice()
    .sort((a, b) => {
      const aa = planningSortDate(a);
      const bb = planningSortDate(b);
      return aa.localeCompare(bb) || (a.entryDate || "").localeCompare(b.entryDate || "") || a.orderNumber.localeCompare(b.orderNumber);
    });
  ordered.forEach((order) => recalculateOrder(order, { stationUsage, departmentUsage }));
}

function recalculateOrder(order, sharedPlan = null) {
  const nominalCapacitiesByDepartment = dailyCapacitiesByDepartment(null, true);
  const nominalCapacitiesByStation = dailyCapacitiesByStation(null, true);
  const stationCapacityCache = {};
  const processCapacityCache = {};
  const capacityForStationOnDate = (stationId, dateValue) => {
    const key = `${stationId}|${dateValue}`;
    if (Object.prototype.hasOwnProperty.call(stationCapacityCache, key)) {
      return stationCapacityCache[key];
    }
    const capacities = dailyCapacitiesByStation(dateValue, false);
    const value = Math.max(0, toFloat(capacities[stationId] || 0));
    stationCapacityCache[key] = value;
    return value;
  };
  const capacityForProcessOnDate = (processKey, dateValue) => {
    const key = `${processKey}|${dateValue}`;
    if (Object.prototype.hasOwnProperty.call(processCapacityCache, key)) {
      return processCapacityCache[key];
    }
    const capacities = dailyCapacitiesByDepartment(dateValue, false);
    const value = Math.max(0, toFloat(capacities[processKey] || 0));
    processCapacityCache[key] = value;
    return value;
  };
  const stationUsage = sharedPlan?.stationUsage || null;
  const departmentUsage = sharedPlan?.departmentUsage || null;
  const localStationUsage = {};
  const localDepartmentUsage = {
    machining: {},
    painting: {},
    assembly: {},
  };
  const totals = { machining: 0, painting: 0, assembly: 0 };
  let currentEndDate = toDate(planningAnchorDate(order));
  let firstProcessStart = null;
  const orderDaily = {};
  const stationDaily = {};
  const processDaily = {
    machining: {},
    painting: {},
    assembly: {},
  };

  PROCESS_FLOW.forEach((step) => {
    let processMinutes = 0;
    const readyPositions = [];
    const startConstraints = [isoDate(currentEndDate), order.entryDate];

    order.positions.forEach((position) => {
      const ready = positionReadinessForDepartment(position, step.department, order.entryDate);
      if (!ready.ready) {
        return;
      }
      processMinutes += toFloat(position.times[step.key]);
      readyPositions.push(position);
      startConstraints.push(ready.availableFrom);
    });

    totals[step.key] = processMinutes;
    if (processMinutes <= 0) {
      return;
    }

    const startDate = maxDate(startConstraints);
    if (!firstProcessStart) {
      firstProcessStart = startDate;
    }
    const split = splitProcessMinutesByFramesAndSashes(readyPositions, step.key, processMinutes);
    const stationWeights = weightedStationMapForPositions(readyPositions, step.key);
    const processStationIds = activeStationIdsForProcess(step.key);
    const streams = [];
    if (split.frameMinutes > 0) {
      streams.push({ kind: "frames", minutes: split.frameMinutes });
    }
    if (split.sashMinutes > 0) {
      streams.push({ kind: "sashes", minutes: split.sashMinutes });
    }
    if (streams.length === 0 && processMinutes > 0) {
      streams.push({ kind: "frames", minutes: processMinutes });
    }

    let processEndDate = toDate(startDate);
    const assignedStreams = assignStreamsToStations(streams, processStationIds, stationWeights, nominalCapacitiesByStation);

    if (assignedStreams.length === 0) {
      const minutesCapacity = (date) =>
        Math.max(1, capacityForProcessOnDate(step.key, date) || nominalCapacitiesByDepartment[step.key] || 1);
      const baseUsageByDay = departmentUsage ? departmentUsage[step.key] || (departmentUsage[step.key] = {}) : null;
      const localUsageByDay = departmentUsage ? localDepartmentUsage[step.key] || (localDepartmentUsage[step.key] = {}) : null;
      const allocation = baseUsageByDay
        ? allocateMinutesByDayWithBaseUsage(startDate, processMinutes, minutesCapacity, baseUsageByDay, localUsageByDay)
        : allocateMinutesByDay(startDate, processMinutes, minutesCapacity);
      allocation.forEach((dayEntry) => {
        addMinutesToDailyMap(orderDaily, dayEntry.date, dayEntry.minutes);
        addMinutesToDailyMap(processDaily[step.key], dayEntry.date, dayEntry.minutes);
      });
      if (allocation.length > 0) {
        processEndDate = toDate(allocation[allocation.length - 1].date);
      }
    } else {
      assignedStreams.forEach((stream) => {
        const minutesCapacity = (date) =>
          Math.max(
            1,
            capacityForStationOnDate(stream.stationId, date) ||
              capacityForProcessOnDate(step.key, date) ||
              nominalCapacitiesByStation[stream.stationId] ||
              nominalCapacitiesByDepartment[step.key] ||
              1,
          );
        const baseUsageByDay = stationUsage ? stationUsage[stream.stationId] || (stationUsage[stream.stationId] = {}) : null;
        const localUsageByDay = stationUsage ? localStationUsage[stream.stationId] || (localStationUsage[stream.stationId] = {}) : null;
        const allocation = baseUsageByDay
          ? allocateMinutesByDayWithBaseUsage(startDate, stream.minutes, minutesCapacity, baseUsageByDay, localUsageByDay)
          : allocateMinutesByDay(startDate, stream.minutes, minutesCapacity);
        allocation.forEach((dayEntry) => {
          addMinutesToDailyMap(orderDaily, dayEntry.date, dayEntry.minutes);
          addMinutesToDailyMap(processDaily[step.key], dayEntry.date, dayEntry.minutes);
          if (!stationDaily[stream.stationId]) {
            stationDaily[stream.stationId] = {};
          }
          addMinutesToDailyMap(stationDaily[stream.stationId], dayEntry.date, dayEntry.minutes);
        });
        if (allocation.length > 0) {
          const streamEndDate = toDate(allocation[allocation.length - 1].date);
          if (streamEndDate > processEndDate) {
            processEndDate = streamEndDate;
          }
        }
      });
    }
    currentEndDate = processEndDate;
  });

  const calculatedPlannedDate = order.positions.length === 0 ? null : isoDate(currentEndDate);
  const manualPlannedDate = resolveManualPlannedDate(order);
  let startDate = firstProcessStart || planningAnchorDate(order);

  if (calculatedPlannedDate && manualPlannedDate) {
    const offset = workingDayOffset(calculatedPlannedDate, manualPlannedDate);
    if (offset !== 0) {
      startDate = shiftDateByWorkingDays(startDate, offset);
      shiftDailyMapByWorkingOffset(orderDaily, offset);
      shiftScopedDailyByWorkingOffset(processDaily, offset);
      shiftScopedDailyByWorkingOffset(stationDaily, offset);
      shiftScopedDailyByWorkingOffset(localDepartmentUsage, offset);
      shiftScopedDailyByWorkingOffset(localStationUsage, offset);
    }
  }

  if (stationUsage) {
    mergeScopedUsage(stationUsage, localStationUsage);
  }
  if (departmentUsage) {
    mergeScopedUsage(departmentUsage, localDepartmentUsage);
  }

  order.calculation = {
    totals,
    startDate,
    calculatedPlannedDate,
    orderDaily,
    stationDaily,
    processDaily,
  };
  order.plannedProductionDate = manualPlannedDate || calculatedPlannedDate;
  const positionTotals = sumOrderFramesAndSashes(order);
  order.framesCount = positionTotals.framesCount;
  order.sashesCount = positionTotals.sashesCount;
}

function addMinutesToDailyMap(map, date, minutes) {
  map[date] = (map[date] || 0) + toFloat(minutes);
}

function shiftDailyMapByWorkingOffset(dailyMap, offset) {
  if (!dailyMap || offset === 0) {
    return;
  }
  const original = Object.entries(dailyMap);
  Object.keys(dailyMap).forEach((key) => {
    delete dailyMap[key];
  });
  original.forEach(([date, value]) => {
    const shiftedDate = shiftDateByWorkingDays(date, offset);
    addMinutesToDailyMap(dailyMap, shiftedDate, value);
  });
}

function shiftScopedDailyByWorkingOffset(scopedDailyMap, offset) {
  if (!scopedDailyMap || offset === 0) {
    return;
  }
  Object.values(scopedDailyMap).forEach((dailyMap) => {
    shiftDailyMapByWorkingOffset(dailyMap, offset);
  });
}

function mergeScopedUsage(target, source) {
  if (!target || !source) {
    return;
  }
  Object.entries(source).forEach(([scopeKey, dailyMap]) => {
    const targetDaily = target[scopeKey] || (target[scopeKey] = {});
    Object.entries(dailyMap || {}).forEach(([date, minutes]) => {
      addMinutesToDailyMap(targetDaily, date, minutes);
    });
  });
}

function sumOrderFramesAndSashes(order) {
  const positions = Array.isArray(order?.positions) ? order.positions : [];
  return positions.reduce(
    (acc, position) => {
      acc.framesCount += Math.max(0, toInt(position.framesCount));
      acc.sashesCount += Math.max(0, toInt(position.sashesCount));
      return acc;
    },
    { framesCount: 0, sashesCount: 0 },
  );
}

function normalizeIsoDateKey(dateValue) {
  if (dateValue === null || dateValue === undefined || dateValue === "") {
    return null;
  }
  const text = String(dateValue).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }
  const parsed = toDate(dateValue);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  return isoDate(parsed);
}

function stationOvertimeMinutesForDate(stationId, dateValue) {
  const dateKey = normalizeIsoDateKey(dateValue);
  if (!dateKey || !stationId) {
    return 0;
  }
  const dayMap = state.settings?.stationOvertime?.[dateKey] || {};
  return clamp(toInt(dayMap?.[stationId] || 0), 0, 1440);
}

function stationCapacityForDate(station, dateValue, ignoreWeekdayLimit = false) {
  if (!station || station.active === false) {
    return 0;
  }
  const dateKey = normalizeIsoDateKey(dateValue);
  const minutesPerShift = Math.max(1, state.settings.minutesPerShift);
  const cfg = state.stationSettings[station.id] || { shiftCount: 2, peopleCount: 1 };
  const stationShiftCount = clamp(toInt(cfg.shiftCount), 1, 3);
  const peopleCount = Math.max(1, toInt(cfg.peopleCount));
  const dayShifts = clamp(toInt(shiftsForDate(dateKey || new Date())), 0, 3);
  const effectiveShiftCount = ignoreWeekdayLimit ? stationShiftCount : dayShifts <= 0 ? 0 : stationShiftCount;
  const overtimeMinutes = dateKey ? stationOvertimeMinutesForDate(station.id, dateKey) : 0;
  if (effectiveShiftCount <= 0) {
    return overtimeMinutes;
  }
  return minutesPerShift * effectiveShiftCount * peopleCount + overtimeMinutes;
}

function dailyCapacitiesByDepartment(dateValue = null, ignoreWeekdayLimit = false) {
  const result = { machining: 0, painting: 0, assembly: 0 };
  state.stations.forEach((station) => {
    const minutes = stationCapacityForDate(station, dateValue, ignoreWeekdayLimit);
    if (minutes <= 0) {
      return;
    }
    if (station.department === "Maszynownia") {
      result.machining += minutes;
    } else if (station.department === "Lakiernia") {
      result.painting += minutes;
    } else if (station.department === "Kompletacja") {
      result.assembly += minutes;
    }
  });
  return result;
}

function dailyCapacitiesByStation(dateValue = null, ignoreWeekdayLimit = false) {
  const result = {};
  state.stations.forEach((station) => {
    result[station.id] = stationCapacityForDate(station, dateValue, ignoreWeekdayLimit);
  });
  return result;
}

function activeStationIdsForProcess(process) {
  const department = processToDepartment(process);
  return state.stations
    .filter((station) => station.department === department && station.active !== false)
    .map((station) => station.id);
}

function splitProcessMinutesByFramesAndSashes(positions, process, fallbackMinutes = 0) {
  let frameMinutes = 0;
  let sashMinutes = 0;
  positions.forEach((position) => {
    const total = Math.max(0, toFloat(position.times?.[process] || 0));
    if (total <= 0) {
      return;
    }
    const frames = Math.max(0, toInt(position.framesCount));
    const sashes = Math.max(0, toInt(position.sashesCount));
    if (frames > 0 && sashes > 0) {
      const sum = frames + sashes;
      frameMinutes += total * (frames / sum);
      sashMinutes += total * (sashes / sum);
    } else if (frames > 0) {
      frameMinutes += total;
    } else if (sashes > 0) {
      sashMinutes += total;
    } else {
      frameMinutes += total;
    }
  });
  const targetTotal = Math.max(0, toFloat(fallbackMinutes));
  const currentTotal = frameMinutes + sashMinutes;
  if (currentTotal <= 0 && targetTotal > 0) {
    return { frameMinutes: targetTotal, sashMinutes: 0 };
  }
  if (targetTotal > 0 && currentTotal > 0 && Math.abs(targetTotal - currentTotal) > 0.0001) {
    frameMinutes += targetTotal - currentTotal;
  }
  return {
    frameMinutes: Math.max(0, frameMinutes),
    sashMinutes: Math.max(0, sashMinutes),
  };
}

function allocateMinutesByDay(startDateValue, totalMinutes, minutesPerDay) {
  const output = [];
  let remaining = Math.max(0, toFloat(totalMinutes));
  let day = normalizeWorkday(startDateValue);
  let safety = 0;
  while (remaining > 0 && safety < 2500) {
    safety += 1;
    const date = isoDate(day);
    const dayCapacity = resolveDailyCapacity(minutesPerDay, date);
    if (dayCapacity <= 0) {
      day = nextWorkingDay(day);
      continue;
    }
    const current = Math.min(dayCapacity, remaining);
    output.push({ date, minutes: current });
    remaining -= current;
    if (remaining > 0) {
      day = nextWorkingDay(day);
    }
  }
  return output;
}

function allocateMinutesByDayWithBaseUsage(
  startDateValue,
  totalMinutes,
  minutesPerDay,
  baseUsageByDay,
  localUsageByDay,
) {
  const output = [];
  let remaining = Math.max(0, toFloat(totalMinutes));
  let day = normalizeWorkday(startDateValue);
  let safety = 0;
  while (remaining > 0 && safety < 2500) {
    safety += 1;
    const date = isoDate(day);
    const dayCapacity = resolveDailyCapacity(minutesPerDay, date);
    if (dayCapacity <= 0) {
      day = nextWorkingDay(day);
      continue;
    }
    const usedBase = Math.max(0, toFloat(baseUsageByDay?.[date] || 0));
    const usedLocal = Math.max(0, toFloat(localUsageByDay?.[date] || 0));
    const available = Math.max(0, dayCapacity - usedBase - usedLocal);
    if (available <= 0) {
      day = nextWorkingDay(day);
      continue;
    }
    const current = Math.min(available, remaining);
    output.push({ date, minutes: current });
    localUsageByDay[date] = usedLocal + current;
    remaining -= current;
    if (remaining > 0) {
      day = nextWorkingDay(day);
    }
  }
  return output;
}

function resolveDailyCapacity(source, date) {
  const raw = typeof source === "function" ? source(date) : source;
  return Math.max(0, toFloat(raw));
}

function weightedStationMapForPositions(positions, process) {
  const aggregate = {};
  let totalMinutes = 0;
  positions.forEach((position) => {
    const minutes = toFloat(position.times[process]);
    if (minutes <= 0) {
      return;
    }
    totalMinutes += minutes;
    const map = normalizedProcessMap(position.technology, process);
    Object.entries(map).forEach(([stationId, ratio]) => {
      aggregate[stationId] = (aggregate[stationId] || 0) + ratio * minutes;
    });
  });
  if (totalMinutes <= 0) {
    return {};
  }
  Object.keys(aggregate).forEach((key) => {
    aggregate[key] /= totalMinutes;
  });
  return aggregate;
}

function assignStreamsToStations(streams, stationIds, stationWeights, stationCapacities) {
  if (!Array.isArray(stationIds) || stationIds.length === 0) {
    return [];
  }
  const queue = streams
    .map((stream, index) => ({ ...stream, index }))
    .filter((stream) => toFloat(stream.minutes) > 0)
    .sort((a, b) => toFloat(b.minutes) - toFloat(a.minutes));
  const used = new Set();
  const out = new Array(streams.length).fill(null);
  queue.forEach((stream) => {
    let stationId = selectPreferredStation(stationIds, stationWeights, stationCapacities, used);
    if (!stationId) {
      stationId = selectPreferredStation(stationIds, stationWeights, stationCapacities, new Set());
    }
    if (!stationId) {
      return;
    }
    out[stream.index] = {
      kind: stream.kind,
      minutes: toFloat(stream.minutes),
      stationId,
    };
    used.add(stationId);
  });
  return out.filter(Boolean);
}

function selectPreferredStation(stationIds, stationWeights, stationCapacities, excludedStations = new Set()) {
  const stationIndex = stationOrderIndexMap();
  const filtered = stationIds.filter((stationId) => !excludedStations.has(stationId));
  const candidates = filtered.length > 0 ? filtered : stationIds.slice();
  if (candidates.length === 0) {
    return null;
  }
  candidates.sort((a, b) => {
    const weightDiff = toFloat(stationWeights?.[b]) - toFloat(stationWeights?.[a]);
    if (Math.abs(weightDiff) > 0.0001) {
      return weightDiff;
    }
    const capacityDiff = toFloat(stationCapacities?.[b]) - toFloat(stationCapacities?.[a]);
    if (Math.abs(capacityDiff) > 0.0001) {
      return capacityDiff;
    }
    const orderDiff = (stationIndex[a] ?? 999999) - (stationIndex[b] ?? 999999);
    if (orderDiff !== 0) {
      return orderDiff;
    }
    return String(a).localeCompare(String(b));
  });
  return candidates[0];
}

function stationOrderIndexMap() {
  const output = {};
  state.stations.forEach((station, index) => {
    output[station.id] = index;
  });
  return output;
}

function calculateStationWorkload() {
  const result = {};
  state.stations.forEach((station) => {
    result[station.id] = 0;
  });
  operationalOrders().forEach((order) => {
    const stationDaily = order.calculation?.stationDaily || {};
    Object.entries(stationDaily).forEach(([stationId, days]) => {
      if (result[stationId] === undefined) {
        result[stationId] = 0;
      }
      Object.values(days || {}).forEach((minutes) => {
        result[stationId] += toFloat(minutes);
      });
    });
  });
  return result;
}

function positionReadinessForDepartment(position, department, fallbackDate) {
  const requiredMaterials = state.materialRules[department] || [];
  const materials = position.materials || {};
  let available = toDate(fallbackDate);
  let blocked = false;
  requiredMaterials.forEach((materialKey) => {
    const mat = materials[materialKey] || { date: null, toOrder: false };
    if (mat.toOrder && !mat.date) {
      blocked = true;
    }
    if (mat.date) {
      available = toDate(maxDate([isoDate(available), mat.date]));
    }
  });
  return {
    ready: !blocked,
    availableFrom: isoDate(available),
  };
}

function normalizedProcessMap(technology, process) {
  const department = processToDepartment(process);
  const stationIds = state.stations
    .filter((item) => item.department === department && item.active !== false)
    .map((item) => item.id);
  const source = (state.technologies[technology] && state.technologies[technology][process]) || {};
  const filtered = {};
  stationIds.forEach((stationId) => {
    filtered[stationId] = Math.max(0, toFloat(source[stationId]));
  });
  const sum = Object.values(filtered).reduce((acc, value) => acc + value, 0);
  if (sum <= 0) {
    const equal = stationIds.length > 0 ? 1 / stationIds.length : 0;
    const fallback = {};
    stationIds.forEach((stationId) => {
      fallback[stationId] = equal;
    });
    return fallback;
  }
  Object.keys(filtered).forEach((stationId) => {
    filtered[stationId] = filtered[stationId] / sum;
  });
  return filtered;
}

function effectiveStartDate(order) {
  return (order.calculation && order.calculation.startDate) || order.entryDate;
}

function effectivePlannedDate(order) {
  return normalizeOptionalDate(order?.plannedProductionDate);
}

function resolveManualPlannedDate(order) {
  const manual = normalizeOptionalDate(order?.manualPlannedDate);
  if (!manual) {
    return null;
  }
  const entry = normalizeOptionalDate(order?.entryDate) || isoDate(new Date());
  const snapped = isoDate(normalizeWorkday(manual));
  return maxDate([entry, snapped]);
}

function planningSortDate(order) {
  return resolveManualPlannedDate(order) || planningAnchorDate(order);
}

function planningAnchorDate(order) {
  const entry = normalizeOptionalDate(order?.entryDate) || isoDate(new Date());
  const manual = normalizeOptionalDate(order?.manualStartDate);
  if (!manual) {
    return entry;
  }
  return maxDate([entry, manual]);
}

function pendingMaterialsCount(order) {
  return order.positions.reduce((sum, position) => sum + pendingMaterialsCountForPosition(position), 0);
}

function pendingMaterialsCountForPosition(position) {
  return MATERIALS.reduce((sum, material) => {
    const entry = position.materials?.[material.key] || { toOrder: false, date: null };
    return sum + (entry.toOrder && !entry.date ? 1 : 0);
  }, 0);
}

function nearestDueOrder() {
  const items = operationalOrders().filter((order) => effectivePlannedDate(order));
  if (items.length === 0) {
    return null;
  }
  return items.sort((a, b) => toDate(effectivePlannedDate(a)) - toDate(effectivePlannedDate(b)))[0];
}

function statusPill(status) {
  if (status === KPI_POSITION_MIXED_STATUS) {
    return `<span class="status-pill status-mixed">${escapeHtml(status)}</span>`;
  }
  if (status === KPI_POSITION_EMPTY_STATUS) {
    return `<span class="status-pill status-empty">${escapeHtml(status)}</span>`;
  }
  if (status === "Zakonczone") {
    return `<span class="status-pill status-done">${escapeHtml(status)}</span>`;
  }
  if (status === "Produkcja" || status === "Kosmetyka") {
    return `<span class="status-pill status-progress">${escapeHtml(status)}</span>`;
  }
  return `<span class="status-pill status-pending">${escapeHtml(status)}</span>`;
}

function statusLabel(status) {
  if (status === "in_progress") {
    return "W toku";
  }
  if (status === "done") {
    return "Zakonczona";
  }
  return "Planowana";
}

function normalizeWorkflowStatusValue(value, fallback = "pending") {
  const raw = String(value || "").trim();
  const allowed = new Set(["pending", "in_progress", "done"]);
  if (allowed.has(raw)) {
    return raw;
  }
  return fallback;
}

function parseFeedbackProgressPercent(value) {
  const raw = String(value ?? "").trim();
  if (!raw) {
    return null;
  }
  const parsed = parseFloat(raw);
  if (!Number.isFinite(parsed)) {
    return null;
  }
  return Math.round(clamp(parsed, 0, 100) * 10) / 10;
}

function positionProgressPercent(position) {
  const status = normalizeWorkflowStatusValue(position?.status, "pending");
  if (status === "done") {
    return 100;
  }
  const parsed = parseFloat(String(position?.progressPercent ?? ""));
  if (!Number.isFinite(parsed)) {
    return status === "in_progress" ? 1 : 0;
  }
  return Math.round(clamp(parsed, 0, 100) * 10) / 10;
}

function processToDepartment(process) {
  if (process === "machining") {
    return "Maszynownia";
  }
  if (process === "painting") {
    return "Lakiernia";
  }
  return "Kompletacja";
}

function materialStateLabel(entry) {
  if (!entry) {
    return "-";
  }
  if (entry.toOrder && !entry.date) {
    return "Do zamowienia (brak daty)";
  }
  if (entry.toOrder && entry.date) {
    return `Do zamowienia, dostepne: ${formatDate(entry.date)}`;
  }
  if (entry.date) {
    return `Dostepne: ${formatDate(entry.date)}`;
  }
  return "Nie dotyczy / bez blokady";
}

function positionShapeLabel(position) {
  const labels = [];
  if (position?.shapeRect) {
    labels.push("Prostokat");
  }
  if (position?.shapeSkos) {
    labels.push("Skos");
  }
  if (position?.shapeLuk) {
    labels.push("Luk");
  }
  return labels.length > 0 ? labels.join(", ") : "-";
}

function positionElementsLabel(position) {
  return [
    `Slemie: ${Math.max(0, toInt(position?.slemieCount))}`,
    `Slupek staly: ${Math.max(0, toInt(position?.slupekStalyCount))}`,
    `Przymyk: ${Math.max(0, toInt(position?.przymykCount))}`,
    `Niski prog: ${Math.max(0, toInt(position?.niskiProgCount))}`,
  ].join(" | ");
}

function hydrateDates() {
  const today = isoDate(new Date());
  el.todayDate.textContent = formatDate(today);
  if (!el.orderForm.entryDate.value) {
    el.orderForm.entryDate.value = today;
  }
  if (el.archiveDateFilterInput && !el.archiveDateFilterInput.value) {
    el.archiveDateFilterInput.value = today;
  }
}

async function api(path, options = {}) {
  const mergedHeaders = { "Content-Type": "application/json", ...(options.headers || {}) };
  const requestOptions = {
    method: options.method || "GET",
    headers: mergedHeaders,
    credentials: "same-origin",
  };
  if (options.body) {
    requestOptions.body = JSON.stringify(options.body);
  }
  const response = await fetch(path, requestOptions);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (response.status === 401 && isLoggedIn()) {
      resetLocalSessionState();
      renderAll();
    }
    throw new Error(payload.error || `Blad API (${response.status})`);
  }
  return payload;
}

async function apiForm(path, formData, options = {}) {
  const requestOptions = {
    method: options.method || "POST",
    body: formData,
    credentials: "same-origin",
    headers: options.headers || {},
  };
  const response = await fetch(path, requestOptions);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (response.status === 401 && isLoggedIn()) {
      resetLocalSessionState();
      renderAll();
    }
    throw new Error(payload.error || `Blad API (${response.status})`);
  }
  return payload;
}

function addWorkingDays(startDate, daysToAdd) {
  const date = normalizeWorkday(startDate);
  let remaining = Math.max(0, daysToAdd);
  while (remaining > 0) {
    const next = nextWorkingDay(date);
    date.setTime(next.getTime());
    remaining -= 1;
  }
  return date;
}

function shiftDateByWorkingDays(dateValue, offset) {
  const start = normalizeWorkday(dateValue);
  let remaining = Math.abs(toInt(offset));
  let current = start;
  while (remaining > 0) {
    current = offset >= 0 ? nextWorkingDay(current) : previousWorkingDay(current);
    remaining -= 1;
  }
  return isoDate(current);
}

function workingDayOffset(fromDateValue, toDateValue) {
  const from = normalizeWorkday(fromDateValue);
  const to = normalizeWorkday(toDateValue);
  const fromIso = isoDate(from);
  const toIso = isoDate(to);
  if (fromIso === toIso) {
    return 0;
  }
  let steps = 0;
  let current = from;
  let guard = 0;
  if (to > from) {
    while (isoDate(current) !== toIso && guard < 5000) {
      current = nextWorkingDay(current);
      steps += 1;
      guard += 1;
    }
    return steps;
  }
  while (isoDate(current) !== toIso && guard < 5000) {
    current = previousWorkingDay(current);
    steps -= 1;
    guard += 1;
  }
  return steps;
}

function normalizeWorkday(dateValue) {
  const date = toDate(dateValue);
  let safety = 0;
  while (!isWorkingDay(date) && safety < 14) {
    date.setDate(date.getDate() + 1);
    safety += 1;
  }
  return date;
}

function nextWorkingDay(dateValue) {
  const date = toDate(dateValue);
  let safety = 0;
  do {
    date.setDate(date.getDate() + 1);
    safety += 1;
  } while (!isWorkingDay(date) && safety < 14);
  return date;
}

function previousWorkingDay(dateValue) {
  const date = toDate(dateValue);
  let safety = 0;
  do {
    date.setDate(date.getDate() - 1);
    safety += 1;
  } while (!isWorkingDay(date) && safety < 14);
  return date;
}

function isWorkingDay(date) {
  return shiftsForDate(date) > 0 || hasAnyStationOvertimeOnDate(date);
}

function hasAnyStationOvertimeOnDate(dateValue) {
  const dateKey = normalizeIsoDateKey(dateValue);
  if (!dateKey) {
    return false;
  }
  const dayMap = state.settings?.stationOvertime?.[dateKey] || {};
  return Object.values(dayMap).some((minutes) => toInt(minutes) > 0);
}

function weekdayShiftCount(dayIndex) {
  const key = clamp(toInt(dayIndex), 0, 6);
  const source = state.settings?.weekdayShifts || {};
  const raw = source[key] ?? source[String(key)] ?? 0;
  return clamp(toInt(raw), 0, 3);
}

function shiftsForDate(dateValue) {
  const date = toDate(dateValue);
  const key = isoDate(date);
  const overrides = state.settings.calendarOverrides || {};
  if (Object.prototype.hasOwnProperty.call(overrides, key)) {
    if (!Boolean(overrides[key])) {
      return 0;
    }
    return Math.max(1, weekdayShiftCount(date.getDay()));
  }
  return weekdayShiftCount(date.getDay());
}

function maxDate(values) {
  const valid = values.filter(Boolean).map((value) => toDate(value));
  if (valid.length === 0) {
    return isoDate(new Date());
  }
  const max = new Date(Math.max(...valid.map((item) => item.getTime())));
  return isoDate(max);
}

function isoDate(dateValue) {
  const date = dateValue instanceof Date ? new Date(dateValue) : new Date(dateValue);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function toDate(dateValue) {
  if (dateValue instanceof Date) {
    const copy = new Date(dateValue);
    copy.setHours(0, 0, 0, 0);
    return copy;
  }
  const value = String(dateValue || "");
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) {
    const fallback = new Date();
    fallback.setHours(0, 0, 0, 0);
    return fallback;
  }
  return date;
}

function addDays(dateValue, days) {
  const date = toDate(dateValue);
  date.setDate(date.getDate() + toInt(days));
  return date;
}

function formatDate(dateValue) {
  if (!dateValue) {
    return "-";
  }
  return toDate(dateValue).toLocaleDateString("pl-PL");
}

function daysBetween(start, end) {
  return Math.round((normalizeMidnight(end).getTime() - normalizeMidnight(start).getTime()) / 86400000);
}

function normalizeMidnight(dateValue) {
  const date = new Date(dateValue);
  date.setHours(0, 0, 0, 0);
  return date;
}

function normalizeOptionalDate(value) {
  const text = String(value || "").trim();
  return text || null;
}

function clamp(value, minimum, maximum) {
  return Math.max(minimum, Math.min(maximum, value));
}

function toBoolean(value) {
  return value === "on" || value === true || value === "true" || value === 1;
}

function toInt(value) {
  const num = parseInt(String(value ?? "0"), 10);
  return Number.isFinite(num) ? num : 0;
}

function toFloat(value) {
  const num = parseFloat(String(value ?? "0"));
  return Number.isFinite(num) ? num : 0;
}

function cssEscape(value) {
  return String(value).replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      const comma = result.indexOf(",");
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.onerror = () => reject(new Error("Nie udalo sie odczytac zalacznika."));
    reader.readAsDataURL(file);
  });
}

