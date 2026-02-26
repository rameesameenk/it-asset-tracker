const STORAGE_KEY = "it_assets_v5";
const USERS_KEY = "it_users_v1";
const DEPARTMENTS_KEY = "it_departments_v1";
const EMPLOYEES_KEY = "it_employees_v1";
const EMPLOYEE_BULK_LOG_KEY = "it_employee_bulk_log_v1";
const CATALOG_BULK_LOG_KEY = "it_catalog_bulk_log_v1";
const COMPANIES_KEY = "it_companies_v1";
const CUSTOM_TYPES_KEY = "it_custom_asset_types_v1";
const CUSTOM_STATUSES_KEY = "it_custom_asset_statuses_v1";
const CUSTOM_TYPE_CODES_KEY = "it_custom_asset_type_codes_v1";
const SOFTWARE_MASTER_KEY = "it_software_master_v1";
const SOFTWARE_INVENTORY_KEY = "it_software_inventory_v1";
const SOFTWARE_PURCHASE_KEY = "it_software_purchase_v1";
const DESIGNATIONS_KEY = "it_designations_v1";
const WORK_SITES_KEY = "it_work_sites_v1";
const SOFTWARE_CHECK_LOCK_KEY = "it_software_check_lock_v1";
const ACTIVITY_LOG_KEY = "it_activity_log_v1";
const RETURN_LOG_KEY = "it_return_logs_v1";
const SESSION_KEY = "it_session_v1";
const LOCATION_KEY = "it_location_v1";
const COMPANY_KEY = "it_company_v1";
const MODULE_KEY = "it_module_v1";
const AUDITS_KEY = "it_audits_v1";
const CLEANING_KEY = "it_cleaning_v1";
const LOCATIONS_KEY = "it_locations_v1";
const LIST_PAGE_SIZE_KEY = "it_list_page_sizes_v1";
const BACKUP_KEY = "it_backups_v1";
const BACKUP_DATE_KEY = "it_backup_last_date";
const FULL_BACKUP_SNAPSHOT_KEY = "it_full_backup_snapshot_v1";
const FULL_BACKUP_SNAPSHOT_AT_KEY = "it_full_backup_snapshot_at_v1";
const CLOUD_SYNC_CONFIG_KEY = "it_cloud_sync_config_v1";
const CLOUD_LAST_REMOTE_AT_KEY = "it_cloud_last_remote_at_v1";
const TRACKED_BACKUP_KEYS = new Set([
  STORAGE_KEY,
  USERS_KEY,
  DEPARTMENTS_KEY,
  EMPLOYEES_KEY,
  EMPLOYEE_BULK_LOG_KEY,
  CATALOG_BULK_LOG_KEY,
  COMPANIES_KEY,
  CUSTOM_TYPES_KEY,
  CUSTOM_STATUSES_KEY,
  CUSTOM_TYPE_CODES_KEY,
  SOFTWARE_MASTER_KEY,
  SOFTWARE_INVENTORY_KEY,
  SOFTWARE_PURCHASE_KEY,
  DESIGNATIONS_KEY,
  WORK_SITES_KEY,
  SOFTWARE_CHECK_LOCK_KEY,
  ACTIVITY_LOG_KEY,
  RETURN_LOG_KEY,
  AUDITS_KEY,
  CLEANING_KEY,
  LOCATIONS_KEY
]);

const DEFAULT_COMPANY_META = {
  Sribal: { prefix: "SB", logo: "assets/sribal-logo.png" },
  Wudless: { prefix: "WD", logo: "assets/wudless-logo.png" }
};
let COMPANY_META = { ...DEFAULT_COMPANY_META };

const TYPE_CODES = {
  Laptop: "LAP",
  Desktop: "DES",
  "Mobile Phone": "MOB",
  Printer: "PRI",
  Monitor: "MON",
  Network: "NET",
  Peripheral: "PER",
  Server: "SRV",
  Other: "OTH"
};
const DEFAULT_ASSET_TYPES = ["Laptop", "Desktop", "Mobile Phone", "Printer", "Monitor", "Network", "Peripheral", "Server", "Other"];
const DEFAULT_STATUSES = ["Available", "Assigned", "Returned", "Repair", "Service", "Not Usable", "Retired"];
const DEFAULT_LOCATIONS = ["Site", "HeadOffice", "Factory"];

const DEVICE_TYPES = new Set(["Laptop", "Desktop", "Mobile Phone"]);

const CODE39_MAP = {
  "0": "nnnwwnwnn", "1": "wnnwnnnnw", "2": "nnwwnnnnw", "3": "wnwwnnnnn", "4": "nnnwwnnnw",
  "5": "wnnwwnnnn", "6": "nnwwwnnnn", "7": "nnnwnnwnw", "8": "wnnwnnwnn", "9": "nnwwnnwnn",
  A: "wnnnnwnnw", B: "nnwnnwnnw", C: "wnwnnwnnn", D: "nnnnwwnnw", E: "wnnnwwnnn", F: "nnwnwwnnn",
  G: "nnnnnwwnw", H: "wnnnnwwnn", I: "nnwnnwwnn", J: "nnnnwwwnn", K: "wnnnnnnww", L: "nnwnnnnww",
  M: "wnwnnnnwn", N: "nnnnwnnww", O: "wnnnwnnwn", P: "nnwnwnnwn", Q: "nnnnnnwww", R: "wnnnnnwwn",
  S: "nnwnnnwwn", T: "nnnnwnwwn", U: "wwnnnnnnw", V: "nwwnnnnnw", W: "wwwnnnnnn", X: "nwnnwnnnw",
  Y: "wwnnwnnnn", Z: "nwwnwnnnn", "-": "nwnnnnwnw", ".": "wwnnnnwnn", " ": "nwwnnnwnn", "$": "nwnwnwnnn",
  "/": "nwnwnnnwn", "+": "nwnnnwnwn", "%": "nnnwnwnwn", "*": "nwnnwnwnn"
};

const screens = {
  auth: document.getElementById("auth-screen"),
  company: document.getElementById("company-screen"),
  module: document.getElementById("module-screen"),
  settings: document.getElementById("settings-screen"),
  app: document.getElementById("app-screen")
};

const pages = {
  catalog: document.getElementById("catalog-page"),
  employees: document.getElementById("employees-page"),
  tracker: document.getElementById("tracker-page"),
  "software-check": document.getElementById("software-check-page"),
  "software-inventory": document.getElementById("software-inventory-page"),
  return: document.getElementById("return-page"),
  audit: document.getElementById("audit-page"),
  "audit-report": document.getElementById("audit-report-page"),
  "audit-status": document.getElementById("audit-status-page"),
  cleaning: document.getElementById("cleaning-page"),
  "activity-history": document.getElementById("activity-history-page"),
  "admin-users": document.getElementById("admin-users-page")
};

const tabs = {
  catalog: document.getElementById("tab-catalog"),
  employees: document.getElementById("tab-employees"),
  tracker: document.getElementById("tab-tracker"),
  "software-check": document.getElementById("tab-software-check"),
  "software-inventory": document.getElementById("tab-software-inventory"),
  return: document.getElementById("tab-return"),
  audit: document.getElementById("tab-audit"),
  "audit-status": document.getElementById("tab-audit-status"),
  cleaning: document.getElementById("tab-cleaning"),
  "activity-history": document.getElementById("tab-activity-history")
};

const loginForm = document.getElementById("login-form");
const authStatus = document.getElementById("auth-status");
const companyStatus = document.getElementById("company-status");
const moduleStatus = document.getElementById("module-status");
const companyGrid = document.getElementById("company-grid");
const moduleCards = Array.from(document.querySelectorAll(".module-card"));
const adminModuleCard = document.getElementById("admin-module-card");
const activityModuleCard = document.getElementById("activity-module-card");
const companyAdminUsersBtn = document.getElementById("company-admin-users-btn");
const moduleAdminUsersBtn = document.getElementById("module-admin-users-btn");
const appAdminUsersBtn = document.getElementById("app-admin-users-btn");
const moduleLocationSelect = document.getElementById("module-location");
const moduleCompanyMeta = document.getElementById("module-company-meta");
const companyAdminSettingsBtn = document.getElementById("company-admin-settings-btn");
const moduleChangeCompanyBtn = document.getElementById("module-change-company");
const moduleResetAllBtn = document.getElementById("module-reset-all");
const moduleAdminSettingsBtn = document.getElementById("module-admin-settings-btn");
const moduleAdminSettingsPanel = document.getElementById("module-admin-settings-panel");
const appAdminSettingsBtn = document.getElementById("app-admin-settings-btn");
const settingsCompanyMeta = document.getElementById("settings-company-meta");
const settingsBackBtn = document.getElementById("settings-back");
const settingsLogoutBtn = document.getElementById("settings-logout");
const downloadFullBackupBtn = document.getElementById("download-full-backup");
const selectSaveFolderBtn = document.getElementById("select-save-folder");
const saveFolderNowBtn = document.getElementById("save-folder-now");
const saveFolderLabel = document.getElementById("save-folder-label");
const uploadBackupFileInput = document.getElementById("upload-backup-file");
const restoreBackupFileBtn = document.getElementById("restore-backup-file");
const cloudUrlInput = document.getElementById("cloud-url");
const cloudAnonKeyInput = document.getElementById("cloud-anon-key");
const cloudAppIdInput = document.getElementById("cloud-app-id");
const cloudAutoSyncInput = document.getElementById("cloud-auto-sync");
const cloudSaveConfigBtn = document.getElementById("cloud-save-config");
const cloudPushNowBtn = document.getElementById("cloud-push-now");
const cloudPullNowBtn = document.getElementById("cloud-pull-now");
const adminSettingsStatus = document.getElementById("admin-settings-status");
const newCompanyNameInput = document.getElementById("new-company-name");
const newCompanyPrefixInput = document.getElementById("new-company-prefix");
const newCompanyLogoInput = document.getElementById("new-company-logo");
const addCompanyBtn = document.getElementById("add-company-btn");
const newAssetTypeInput = document.getElementById("new-asset-type");
const addAssetTypeBtn = document.getElementById("add-asset-type-btn");
const newAssetStatusInput = document.getElementById("new-asset-status");
const addAssetStatusBtn = document.getElementById("add-asset-status-btn");
const deleteDepartmentSelect = document.getElementById("delete-department-select");
const deleteDepartmentBtn = document.getElementById("delete-department-btn");
const newLocationInput = document.getElementById("new-location");
const addLocationBtn = document.getElementById("add-location-btn");
const newDesignationInput = document.getElementById("new-designation");
const addDesignationBtn = document.getElementById("add-designation-btn");
const newWorkSiteInput = document.getElementById("new-work-site-name");
const addWorkSiteBtn = document.getElementById("add-work-site-btn");
const deleteLocationSelect = document.getElementById("delete-location-select");
const deleteLocationBtn = document.getElementById("delete-location-btn");
const deleteDesignationSelect = document.getElementById("delete-designation-select");
const deleteDesignationBtn = document.getElementById("delete-designation-btn");
const deleteWorkSiteSelect = document.getElementById("delete-work-site-select");
const deleteWorkSiteBtn = document.getElementById("delete-work-site-btn");
const renameLocationSelect = document.getElementById("rename-location-select");
const renameLocationInput = document.getElementById("rename-location-input");
const renameLocationBtn = document.getElementById("rename-location-btn");
const renameDepartmentSelect = document.getElementById("rename-department-select");
const renameDepartmentInput = document.getElementById("rename-department-input");
const renameDepartmentBtn = document.getElementById("rename-department-btn");
const renameAssetTypeSelect = document.getElementById("rename-asset-type-select");
const renameAssetTypeInput = document.getElementById("rename-asset-type-input");
const renameAssetTypeBtn = document.getElementById("rename-asset-type-btn");
const renameStatusSelect = document.getElementById("rename-status-select");
const renameStatusInput = document.getElementById("rename-status-input");
const renameStatusBtn = document.getElementById("rename-status-btn");
const renameSoftwareSelect = document.getElementById("rename-software-select");
const renameSoftwareInput = document.getElementById("rename-software-input");
const renameSoftwareBtn = document.getElementById("rename-software-btn");
const renameDesignationSelect = document.getElementById("rename-designation-select");
const renameDesignationInput = document.getElementById("rename-designation-input");
const renameDesignationBtn = document.getElementById("rename-designation-btn");
const deleteAssetTypeSelect = document.getElementById("delete-asset-type-select");
const deleteAssetTypeBtn = document.getElementById("delete-asset-type-btn");
const deleteStatusSelect = document.getElementById("delete-status-select");
const deleteStatusBtn = document.getElementById("delete-status-btn");
const authBackupBtn = document.getElementById("auth-backup");
const companyBackupBtn = document.getElementById("company-backup");
const moduleBackupBtn = document.getElementById("module-backup");
const settingsBackupBtn = document.getElementById("settings-backup");
const appBackupBtn = document.getElementById("app-backup");
const authLogoutBtn = document.getElementById("auth-logout");
const companyLogoutBtn = document.getElementById("company-logout");
const moduleLogoutBtn = document.getElementById("module-logout");

const sessionMeta = document.getElementById("session-meta");
const sessionAvatar = document.getElementById("session-avatar");
const companyLogo = document.getElementById("company-logo");
const goModuleSelectorBtn = document.getElementById("go-module-selector");
const changeCompanyBtn = document.getElementById("change-company");
const logoutBtn = document.getElementById("logout");
const adminUserForm = document.getElementById("admin-user-form");
const newUserUsername = document.getElementById("new-user-username");
const newUserPassword = document.getElementById("new-user-password");
const newUserPhoto = document.getElementById("new-user-photo");
const newUserRole = document.getElementById("new-user-role");
const adminUserPhotoEditInput = document.getElementById("admin-user-photo-edit-input");
const adminUserStatus = document.getElementById("admin-user-status");
const adminUserTbody = document.getElementById("admin-user-tbody");
const adminUserEditModal = document.getElementById("admin-user-edit-modal");
const adminUserEditClose = document.getElementById("admin-user-edit-close");
const adminUserEditCancel = document.getElementById("admin-user-edit-cancel");
const adminUserEditForm = document.getElementById("admin-user-edit-form");
const adminUserEditUsername = document.getElementById("admin-user-edit-username");
const adminUserEditPassword = document.getElementById("admin-user-edit-password");
const adminUserEditRole = document.getElementById("admin-user-edit-role");
const adminUserEditPhoto = document.getElementById("admin-user-edit-photo");
const adminUserEditPhotoPreview = document.getElementById("admin-user-edit-photo-preview");
const adminUserEditStatus = document.getElementById("admin-user-edit-status");

const catalogForm = document.getElementById("catalog-form");
const catalogFormTitle = document.getElementById("catalog-form-title");
const catalogStatusMsg = document.getElementById("catalog-status-msg");
const catalogSaveBtn = document.getElementById("catalog-save");
const catalogCancelBtn = document.getElementById("catalog-cancel");
const catalogDeleteAllBtn = document.getElementById("catalog-delete-all");
const catalogImportCsvInput = document.getElementById("catalog-import-csv");
const catalogSampleCsvBtn = document.getElementById("catalog-sample-csv");
const catalogExportCsvBtn = document.getElementById("catalog-export-csv");
const catalogBulkHistorySelect = document.getElementById("catalog-bulk-history");
const catalogUndoLastBulkBtn = document.getElementById("catalog-undo-last-bulk");
const catalogUndoAllBulkBtn = document.getElementById("catalog-undo-all-bulk");
const catalogDeleteBulkSelectedBtn = document.getElementById("catalog-delete-bulk-selected");
const catalogViewStatus = document.getElementById("catalog-view-status");
const catalogSearchInput = document.getElementById("catalog-search");
const catalogSelectAllBtn = document.getElementById("catalog-select-all");
const catalogDeleteSelectedBtn = document.getElementById("catalog-delete-selected");
const catalogListTbody = document.getElementById("catalog-list-tbody");
const toggleAdminInputBtn = document.getElementById("toggle-admin-input");
const employeeForm = document.getElementById("employee-form");
const employeeName = document.getElementById("employee-name");
const employeeIdInput = document.getElementById("employee-id");
const employeeMobileInput = document.getElementById("employee-mobile");
const employeeEmailInput = document.getElementById("employee-email");
const employeeDesignationInput = document.getElementById("employee-designation");
const employeeGender = document.getElementById("employee-gender");
const employeeDepartment = document.getElementById("employee-department");
const employeeCompany = document.getElementById("employee-company");
const employeeLocation = document.getElementById("employee-location");
const employeeSiteNameWrap = document.getElementById("employee-site-name-wrap");
const employeeSiteNameInput = document.getElementById("employee-site-name");
const employeePhotoInput = document.getElementById("employee-photo");
const employeePhotoPreview = document.getElementById("employee-photo-preview");
const employeePhotoEditInput = document.getElementById("employee-photo-edit-input");
const employeeSaveBtn = document.getElementById("employee-save");
const employeeCancelBtn = document.getElementById("employee-cancel");
const employeeStatusMsg = document.getElementById("employee-status-msg");
const employeeSearchInput = document.getElementById("employee-search");
const employeeSortSelect = document.getElementById("employee-sort");
const employeeSelectAllBtn = document.getElementById("employee-select-all");
const employeeDeleteSelectedBtn = document.getElementById("employee-delete-selected");
const employeeBulkHistorySelect = document.getElementById("employee-bulk-history");
const employeeUndoLastBulkBtn = document.getElementById("employee-undo-last-bulk");
const employeeUndoAllBulkBtn = document.getElementById("employee-undo-all-bulk");
const employeeDeleteBulkSelectedBtn = document.getElementById("employee-delete-bulk-selected");
const employeeTbody = document.getElementById("employee-tbody");
const employeeSiteNameCol = document.getElementById("employee-site-name-col");
const employeeImportCsvInput = document.getElementById("employee-import-csv");
const employeeSampleCsvBtn = document.getElementById("employee-sample-csv");
const employeeCsvHeadersHint = document.getElementById("employee-csv-headers-hint");
const employeePhotoModal = document.getElementById("employee-photo-modal");
const employeePhotoModalClose = document.getElementById("employee-photo-modal-close");
const employeePhotoModalImage = document.getElementById("employee-photo-modal-image");
const employeeEditModal = document.getElementById("employee-edit-modal");
const employeeEditModalClose = document.getElementById("employee-edit-modal-close");
const employeeEditModalCancel = document.getElementById("employee-edit-modal-cancel");
const employeeEditForm = document.getElementById("employee-edit-form");
const employeeEditName = document.getElementById("employee-edit-name");
const employeeEditId = document.getElementById("employee-edit-id");
const employeeEditMobile = document.getElementById("employee-edit-mobile");
const employeeEditEmail = document.getElementById("employee-edit-email");
const employeeEditDesignation = document.getElementById("employee-edit-designation");
const employeeEditGender = document.getElementById("employee-edit-gender");
const employeeEditDepartment = document.getElementById("employee-edit-department");
const employeeEditCompany = document.getElementById("employee-edit-company");
const employeeEditLocation = document.getElementById("employee-edit-location");
const employeeEditSiteNameWrap = document.getElementById("employee-edit-site-name-wrap");
const employeeEditSiteName = document.getElementById("employee-edit-site-name");
const employeeEditPhoto = document.getElementById("employee-edit-photo");
const employeeEditPhotoPreview = document.getElementById("employee-edit-photo-preview");
const employeeEditStatusMsg = document.getElementById("employee-edit-status-msg");
const employeeNameDatalist = document.getElementById("employee-name-list");
const employeeDepartmentDatalist = document.getElementById("employee-department-list");
const employeeWorkSiteDatalist = document.getElementById("employee-work-site-list");
const employeeDesignationDatalist = document.getElementById("employee-designation-list");

const catalogFields = {
  tag: document.getElementById("catalog-tag"),
  name: document.getElementById("catalog-name"),
  type: document.getElementById("catalog-type"),
  serial: document.getElementById("catalog-serial"),
  macAddress: document.getElementById("catalog-mac"),
  imei: document.getElementById("catalog-imei"),
  os: document.getElementById("catalog-os"),
  storage: document.getElementById("catalog-storage"),
  ram: document.getElementById("catalog-ram"),
  graphics: document.getElementById("catalog-graphics"),
  printerMode: document.getElementById("catalog-printer-mode"),
  printerIp: document.getElementById("catalog-printer-ip"),
  printerPassword: document.getElementById("catalog-printer-password"),
  adminPassword: document.getElementById("catalog-admin-password"),
  company: document.getElementById("catalog-company"),
  location: document.getElementById("catalog-location"),
  purchaseDate: document.getElementById("catalog-date"),
  warrantyDate: document.getElementById("catalog-warranty-date"),
  status: document.getElementById("catalog-status"),
  notes: document.getElementById("catalog-notes")
};
const catalogAssignedOwner = document.getElementById("catalog-assigned-owner");
const catalogAssignedGender = document.getElementById("catalog-assigned-gender");
const catalogAssignedDepartment = document.getElementById("catalog-assigned-department");
const catalogAssignedSystem = document.getElementById("catalog-assigned-system");
const catalogAssignedLocation = document.getElementById("catalog-assigned-location");
const catalogAssignedPhotoInput = document.getElementById("catalog-assigned-photo");
const catalogAssignedPhotoPreview = document.getElementById("catalog-assigned-photo-preview");
const catalogAssignedOwnerWrap = document.getElementById("catalog-assigned-owner-wrap");
const catalogAssignedGenderWrap = document.getElementById("catalog-assigned-gender-wrap");
const catalogAssignedDepartmentWrap = document.getElementById("catalog-assigned-department-wrap");
const catalogAssignedSystemWrap = document.getElementById("catalog-assigned-system-wrap");
const catalogAssignedLocationWrap = document.getElementById("catalog-assigned-location-wrap");
const catalogAssignedPhotoWrap = document.getElementById("catalog-assigned-photo-wrap");
const catalogAssignedPhotoPreviewWrap = document.getElementById("catalog-assigned-photo-preview-wrap");
const catalogEditModal = document.getElementById("catalog-edit-modal");
const catalogEditModalTitle = document.getElementById("catalog-edit-modal-title");
const catalogEditModalClose = document.getElementById("catalog-edit-modal-close");
const catalogEditModalCancel = document.getElementById("catalog-edit-modal-cancel");
const catalogEditForm = document.getElementById("catalog-edit-form");
const catalogEditStatusMsg = document.getElementById("catalog-edit-status-msg");
const catalogEditFields = {
  tag: document.getElementById("catalog-edit-tag"),
  name: document.getElementById("catalog-edit-name"),
  type: document.getElementById("catalog-edit-type"),
  serial: document.getElementById("catalog-edit-serial"),
  macAddress: document.getElementById("catalog-edit-mac"),
  imei: document.getElementById("catalog-edit-imei"),
  os: document.getElementById("catalog-edit-os"),
  storage: document.getElementById("catalog-edit-storage"),
  ram: document.getElementById("catalog-edit-ram"),
  graphics: document.getElementById("catalog-edit-graphics"),
  printerMode: document.getElementById("catalog-edit-printer-mode"),
  printerIp: document.getElementById("catalog-edit-printer-ip"),
  printerPassword: document.getElementById("catalog-edit-printer-password"),
  adminPassword: document.getElementById("catalog-edit-admin-password"),
  company: document.getElementById("catalog-edit-company"),
  location: document.getElementById("catalog-edit-location"),
  purchaseDate: document.getElementById("catalog-edit-date"),
  warrantyDate: document.getElementById("catalog-edit-warranty-date"),
  status: document.getElementById("catalog-edit-status"),
  notes: document.getElementById("catalog-edit-notes")
};
const catalogEditAssignedOwner = document.getElementById("catalog-edit-assigned-owner");
const catalogEditAssignedGender = document.getElementById("catalog-edit-assigned-gender");
const catalogEditAssignedDepartment = document.getElementById("catalog-edit-assigned-department");
const catalogEditAssignedSystem = document.getElementById("catalog-edit-assigned-system");
const catalogEditAssignedLocation = document.getElementById("catalog-edit-assigned-location");
const catalogEditAssignedPhotoInput = document.getElementById("catalog-edit-assigned-photo");
const catalogEditAssignedPhotoPreview = document.getElementById("catalog-edit-assigned-photo-preview");
const catalogEditMacWrap = document.getElementById("catalog-edit-mac-wrap");
const catalogEditImeiWrap = document.getElementById("catalog-edit-imei-wrap");
const catalogEditOsWrap = document.getElementById("catalog-edit-os-wrap");
const catalogEditStorageWrap = document.getElementById("catalog-edit-storage-wrap");
const catalogEditRamWrap = document.getElementById("catalog-edit-ram-wrap");
const catalogEditGraphicsWrap = document.getElementById("catalog-edit-graphics-wrap");
const catalogEditPrinterModeWrap = document.getElementById("catalog-edit-printer-mode-wrap");
const catalogEditPrinterIpWrap = document.getElementById("catalog-edit-printer-ip-wrap");
const catalogEditPrinterPasswordWrap = document.getElementById("catalog-edit-printer-password-wrap");
const catalogEditAdminPasswordWrap = document.getElementById("catalog-edit-admin-password-wrap");
const catalogEditAssignedOwnerWrap = document.getElementById("catalog-edit-assigned-owner-wrap");
const catalogEditAssignedGenderWrap = document.getElementById("catalog-edit-assigned-gender-wrap");
const catalogEditAssignedDepartmentWrap = document.getElementById("catalog-edit-assigned-department-wrap");
const catalogEditAssignedSystemWrap = document.getElementById("catalog-edit-assigned-system-wrap");
const catalogEditAssignedLocationWrap = document.getElementById("catalog-edit-assigned-location-wrap");
const catalogEditAssignedPhotoWrap = document.getElementById("catalog-edit-assigned-photo-wrap");
const catalogEditAssignedPhotoPreviewWrap = document.getElementById("catalog-edit-assigned-photo-preview-wrap");

const macWrap = document.getElementById("mac-wrap");
const imeiWrap = document.getElementById("imei-wrap");
const osWrap = document.getElementById("os-wrap");
const storageWrap = document.getElementById("storage-wrap");
const ramWrap = document.getElementById("ram-wrap");
const graphicsWrap = document.getElementById("graphics-wrap");
const adminPasswordWrap = document.getElementById("admin-password-wrap");
const printerModeWrap = document.getElementById("printer-mode-wrap");
const printerIpWrap = document.getElementById("printer-ip-wrap");
const printerPasswordWrap = document.getElementById("printer-password-wrap");

const assignStatusMsg = document.getElementById("tracker-status-msg");
const assignAssetSelect = document.getElementById("assign-asset-select");
const assignAssetSuggestions = document.getElementById("assign-asset-suggestions");
const assignAssetName = document.getElementById("assign-asset-name");
const assignAssetType = document.getElementById("assign-asset-type");
const assignSystemName = document.getElementById("assign-system-name");
const assignLocationDetail = document.getElementById("assign-location-detail");
const assignLocationDetailWrap = document.getElementById("assign-location-detail-wrap");
const assignDepartment = document.getElementById("assign-department");
const assignOwner = document.getElementById("assign-owner");
const assignGender = document.getElementById("assign-gender");
const assignStatus = document.getElementById("assign-status");
const statusAction = document.getElementById("status-action");
const assignPhoto = document.getElementById("assign-photo");
const assignPhotoPreview = document.getElementById("assign-photo-preview");
const addDepartmentBtn = document.getElementById("add-department");
const addDepartmentWrap = document.getElementById("add-department-wrap");
const newDepartmentInput = document.getElementById("new-department");
const assignSaveBtn = document.getElementById("assign-save");
const deassignBtn = document.getElementById("deassign-asset");
const returnBtn = document.getElementById("return-asset");
const applyStatusBtn = document.getElementById("apply-status");
const returnForm = document.getElementById("return-form");
const returnStatusMsg = document.getElementById("return-status-msg");
const returnAssetSelect = document.getElementById("return-asset-select");
const returnAssetSuggestions = document.getElementById("return-asset-suggestions");
const returnAssetName = document.getElementById("return-asset-name");
const returnAssetType = document.getElementById("return-asset-type");
const returnCurrentOwner = document.getElementById("return-current-owner");
const returnCurrentStatus = document.getElementById("return-current-status");
const returnDateInput = document.getElementById("return-date");
const returnReasonSelect = document.getElementById("return-reason");
const returnNoteInput = document.getElementById("return-note");
const returnSaveBtn = document.getElementById("return-save");
const returnClearBtn = document.getElementById("return-clear");
const returnExportCsvBtn = document.getElementById("return-export-csv");
const returnHistoryTbody = document.getElementById("return-history-tbody");

const searchInput = document.getElementById("search");
const scanInput = document.getElementById("scan-input");
const filterType = document.getElementById("filter-type");
const filterStatus = document.getElementById("filter-status");
const scanFindBtn = document.getElementById("scan-find");
const manualBackupBtn = document.getElementById("manual-backup");
const exportMode = document.getElementById("export-mode");
const exportFrom = document.getElementById("export-from");
const exportTo = document.getElementById("export-to");
const exportCsvBtn = document.getElementById("export-csv");

const tableBody = document.getElementById("asset-tbody");
const thMac = document.getElementById("th-mac");
const thImei = document.getElementById("th-imei");
const thOs = document.getElementById("th-os");
const thStorage = document.getElementById("th-storage");
const thRam = document.getElementById("th-ram");
const thGraphics = document.getElementById("th-graphics");

const barcodePreview = document.getElementById("barcode-preview");
const printBtn = document.getElementById("print-label");
const fieldSuggestionBoxes = {
  name: document.getElementById("catalog-name-suggestions-box"),
  serial: document.getElementById("catalog-serial-suggestions-box"),
  os: document.getElementById("catalog-os-suggestions-box"),
  storage: document.getElementById("catalog-storage-suggestions-box"),
  ram: document.getElementById("catalog-ram-suggestions-box"),
  graphics: document.getElementById("catalog-graphics-suggestions-box"),
  printerIp: document.getElementById("catalog-printer-ip-suggestions-box"),
  owner: document.getElementById("assign-owner-suggestions-box")
};
const fieldSuggestionInputs = {
  name: catalogFields.name,
  serial: catalogFields.serial,
  os: catalogFields.os,
  storage: catalogFields.storage,
  ram: catalogFields.ram,
  graphics: catalogFields.graphics,
  printerIp: catalogFields.printerIp,
  owner: assignOwner
};

const auditLocation = document.getElementById("audit-location");
const auditCount = document.getElementById("audit-count");
const runAuditBtn = document.getElementById("run-audit");
const openAuditReportBtn = document.getElementById("open-audit-report");
const viewAuditReportsBtn = document.getElementById("view-audit-reports");
const viewAuditDueBtn = document.getElementById("view-audit-due");
const saveAuditBtn = document.getElementById("save-audit");
const auditStatus = document.getElementById("audit-status");
const auditTbody = document.getElementById("audit-tbody");
const auditReportForm = document.getElementById("audit-report-form");
const auditReportAsset = document.getElementById("audit-report-asset");
const auditReportBackBtn = document.getElementById("audit-report-back");
const auditReportStatus = document.getElementById("audit-report-status");
const auditReportHistoryTbody = document.getElementById("audit-report-history-tbody");
const auditReportHistorySystem = document.getElementById("audit-report-history-system");
const auditReportSearch = document.getElementById("audit-report-search");
const auditReportPageSize = document.getElementById("audit-report-page-size");
const refreshAuditHistoryBtn = document.getElementById("refresh-audit-history");
const auditReportDate = document.getElementById("audit-report-date");
const auditReportPcName = document.getElementById("audit-report-pc-name");
const auditReportDepartment = document.getElementById("audit-report-department");
const auditReportUserName = document.getElementById("audit-report-user-name");
const auditReportOsVersion = document.getElementById("audit-report-os-version");
const auditReportWinActivated = document.getElementById("audit-report-win-activated");
const auditReportAdminVerified = document.getElementById("audit-report-admin-verified");
const auditReportUnauthorizedSw = document.getElementById("audit-report-unauthorized-sw");
const auditReportAntivirus = document.getElementById("audit-report-antivirus");
const auditReportFirewall = document.getElementById("audit-report-firewall");
const auditReportUpdatesPending = document.getElementById("audit-report-updates-pending");
const auditReportUsbRestricted = document.getElementById("audit-report-usb-restricted");
const auditReportPrinterConfig = document.getElementById("audit-report-printer-config");
const auditReportBackupVerified = document.getElementById("audit-report-backup-verified");
const auditReportIssues = document.getElementById("audit-report-issues");
const auditReportRiskLevel = document.getElementById("audit-report-risk-level");
const auditReportActionTaken = document.getElementById("audit-report-action-taken");
const auditReportAuditorName = document.getElementById("audit-report-auditor-name");
const auditStatusLocation = document.getElementById("audit-status-location");
const auditStatusFilter = document.getElementById("audit-status-filter");
const refreshAuditStatusBtn = document.getElementById("refresh-audit-status");
const auditStatusBackBtn = document.getElementById("audit-status-back");
const auditStatusMsg = document.getElementById("audit-status-msg");
const auditStatusTbody = document.getElementById("audit-status-tbody");
const auditReportViewPanel = document.getElementById("audit-report-view-panel");
const auditReportViewTemplate = document.getElementById("audit-report-view-template");
const auditReportViewCloseBtn = document.getElementById("audit-report-view-close");
const auditEditModal = document.getElementById("audit-edit-modal");
const auditEditModalClose = document.getElementById("audit-edit-modal-close");
const auditEditModalCancel = document.getElementById("audit-edit-modal-cancel");
const auditEditForm = document.getElementById("audit-edit-form");
const auditEditStatusMsg = document.getElementById("audit-edit-status-msg");
const auditEditAssetTag = document.getElementById("audit-edit-asset-tag");
const auditEditDate = document.getElementById("audit-edit-date");
const auditEditPcName = document.getElementById("audit-edit-pc-name");
const auditEditDepartment = document.getElementById("audit-edit-department");
const auditEditUserName = document.getElementById("audit-edit-user-name");
const auditEditOsVersion = document.getElementById("audit-edit-os-version");
const auditEditWinActivated = document.getElementById("audit-edit-win-activated");
const auditEditAdminVerified = document.getElementById("audit-edit-admin-verified");
const auditEditUnauthorizedSw = document.getElementById("audit-edit-unauthorized-sw");
const auditEditAntivirus = document.getElementById("audit-edit-antivirus");
const auditEditFirewall = document.getElementById("audit-edit-firewall");
const auditEditUpdatesPending = document.getElementById("audit-edit-updates-pending");
const auditEditUsbRestricted = document.getElementById("audit-edit-usb-restricted");
const auditEditPrinterConfig = document.getElementById("audit-edit-printer-config");
const auditEditBackupVerified = document.getElementById("audit-edit-backup-verified");
const auditEditRiskLevel = document.getElementById("audit-edit-risk-level");
const auditEditAuditorName = document.getElementById("audit-edit-auditor-name");
const auditEditIssues = document.getElementById("audit-edit-issues");
const auditEditActionTaken = document.getElementById("audit-edit-action-taken");
const activitySearchInput = document.getElementById("activity-search");
const activityFilterUser = document.getElementById("activity-filter-user");
const activityFilterAction = document.getElementById("activity-filter-action");
const activityDateFrom = document.getElementById("activity-date-from");
const activityDateTo = document.getElementById("activity-date-to");
const activitySort = document.getElementById("activity-sort");
const activityRefreshBtn = document.getElementById("activity-refresh");
const activityExportCsvBtn = document.getElementById("activity-export-csv");
const activityExportAllBtn = document.getElementById("activity-export-all");
const clearActivityHistoryBtn = document.getElementById("clear-activity-history");
const activityStatusMsg = document.getElementById("activity-status-msg");
const activityTbody = document.getElementById("activity-tbody");

const cleaningLocation = document.getElementById("cleaning-location");
const refreshCleaningBtn = document.getElementById("refresh-cleaning");
const cleaningStatus = document.getElementById("cleaning-status");
const cleaningTbody = document.getElementById("cleaning-tbody");
const newSoftwareNameInput = document.getElementById("new-software-name");
const addSoftwareBtn = document.getElementById("add-software-btn");
const deleteSoftwareSelect = document.getElementById("delete-software-select");
const deleteSoftwareBtn = document.getElementById("delete-software-btn");
const softwareCheckLockToggle = document.getElementById("software-check-lock-toggle");
const softwareForm = document.getElementById("software-form");
const softwareAssetSearch = document.getElementById("software-asset-search");
const softwareAssetSelect = document.getElementById("software-asset-select");
const softwareAssignedOwner = document.getElementById("software-assigned-owner");
const softwareSelectedMac = document.getElementById("software-selected-mac");
const softwareNameSelect = document.getElementById("software-name-select");
const softwarePiratedSelect = document.getElementById("software-pirated-select");
const softwareStatusMsg = document.getElementById("software-status-msg");
const softwareSearchInput = document.getElementById("software-search");
const softwareViewScope = document.getElementById("software-view-scope");
const softwareSortSoftware = document.getElementById("software-sort-software");
const softwareSortPirated = document.getElementById("software-sort-pirated");
const softwareTbody = document.getElementById("software-tbody");
const softwareDeviceTbody = document.getElementById("software-device-tbody");
const softwareDeviceSearch = document.getElementById("software-device-search");
const softwareDeviceViewModal = document.getElementById("software-device-view-modal");
const softwareDeviceViewTitle = document.getElementById("software-device-view-title");
const softwareDeviceViewSummary = document.getElementById("software-device-view-summary");
const softwareDeviceViewTbody = document.getElementById("software-device-view-tbody");
const softwareDeviceViewClose = document.getElementById("software-device-view-close");
const softwareInventoryForm = document.getElementById("software-inventory-form");
const inventorySoftwareName = document.getElementById("inventory-software-name");
const inventoryPurchaseDate = document.getElementById("inventory-purchase-date");
const inventoryExpiryDate = document.getElementById("inventory-expiry-date");
const inventoryCost = document.getElementById("inventory-cost");
const inventoryActivationCode = document.getElementById("inventory-activation-code");
const inventoryMaxUsers = document.getElementById("inventory-max-users");
const inventoryAssignedSearch = document.getElementById("inventory-assigned-search");
const inventoryAssignedSystems = document.getElementById("inventory-assigned-systems");
const inventorySaveBtn = document.getElementById("inventory-save");
const inventoryCancelBtn = document.getElementById("inventory-cancel");
const inventoryStatusMsg = document.getElementById("inventory-status-msg");
const inventorySearchInput = document.getElementById("inventory-search");
const inventoryTbody = document.getElementById("inventory-tbody");
const inventoryEditModal = document.getElementById("inventory-edit-modal");
const inventoryModalCloseBtn = document.getElementById("inventory-modal-close");
const inventoryModalCancelBtn = document.getElementById("inventory-modal-cancel");
const inventoryModalForm = document.getElementById("inventory-modal-form");
const inventoryModalSoftwareName = document.getElementById("inventory-modal-software-name");
const inventoryModalPurchaseDate = document.getElementById("inventory-modal-purchase-date");
const inventoryModalExpiryDate = document.getElementById("inventory-modal-expiry-date");
const inventoryModalCost = document.getElementById("inventory-modal-cost");
const inventoryModalActivationCode = document.getElementById("inventory-modal-activation-code");
const inventoryModalMaxUsers = document.getElementById("inventory-modal-max-users");
const inventoryModalAssignedSearch = document.getElementById("inventory-modal-assigned-search");
const inventoryModalAssignedSystems = document.getElementById("inventory-modal-assigned-systems");
const inventoryModalStatusMsg = document.getElementById("inventory-modal-status-msg");

let users = [];
let departments = [];
let locations = [];
let designations = [];
let workSites = [];
let employees = [];
let customAssetTypes = [];
let customAssetStatuses = [];
let customTypeCodes = {};
let softwareMaster = [];
let softwareInventory = [];
let softwarePurchases = [];
let softwareCheckRequiresAdminPassword = false;
let activityLogs = [];
let returnLogs = [];
let assets = [];
let audits = [];
let cleaningRecords = {};
let sessionUser = null;
let currentCompany = null;
let currentLocation = null;
let activePage = "tracker";
let editingCatalogId = null;
let selectedAssetId = null;
let assignPhotoDataUrl = "";
let auditSelection = [];
let selectedAuditAssetId = null;
let selectedAuditHistorySystemId = "";
let assignOptionLookup = new Map();
let assignOptions = [];
let returnOptionLookup = new Map();
let returnOptions = [];
let settingsReturnScreen = "module";
let employeePhotoDataUrl = "";
let editingEmployeePhotoId = null;
let editingEmployeeId = null;
let editingEmployeeModalId = null;
let employeeEditPhotoDataUrl = "";
let selectedEmployeeIds = new Set();
let employeeBulkHistory = [];
let catalogBulkHistory = [];
let catalogAssignedPhotoDataUrl = "";
let selectedCatalogAssetIds = new Set();
let editingCatalogModalId = null;
let catalogEditAssignedPhotoDataUrl = "";
let selectedSoftwareAssetId = null;
let editingSoftwarePurchaseId = null;
let editingAuditReportId = null;
let editingAuditModalId = null;
let modalInventorySelectedIds = new Set();
let saveFolderHandle = null;
let saveFolderWriteTimer = null;
let saveFolderWriteInFlight = false;
let saveFolderQueuedReason = "";
let listPageSizes = {};
let cloudSyncConfig = { url: "", anonKey: "", appId: "it-asset-tracker", autoSync: false };
let cloudLastRemoteAt = "";
let cloudPushTimer = null;
let cloudPushInFlight = false;
let cloudPullInFlight = false;
let cloudSyncSuspendLocalHooks = false;
let cloudPullInterval = null;
let editingAdminUserId = null;
let editingAdminUserPhotoDataUrl = "";
const AUDIT_YN_OPTIONS = ["", "Y", "N"];
const AUDIT_RISK_OPTIONS = ["Low", "Medium", "High", "Critical"];
const LIST_PAGE_SIZE_TARGETS = new Set([
  "admin-user-tbody",
  "activity-tbody",
  "catalog-list-tbody",
  "software-device-tbody",
  "software-tbody",
  "inventory-tbody",
  "employee-tbody",
  "asset-tbody",
  "return-history-tbody",
  "audit-tbody",
  "audit-report-history-tbody",
  "audit-status-tbody",
  "cleaning-tbody"
]);
const fieldSuggestionValues = {
  name: [],
  serial: [],
  os: [],
  storage: [],
  ram: [],
  graphics: [],
  printerIp: [],
  owner: []
};

function uid() {
  return window.crypto?.randomUUID ? window.crypto.randomUUID() : `id_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function todayKey() {
  return new Date().toISOString().slice(0, 10);
}

function normalizeListPageSize(value, fallback = "20") {
  const raw = String(value || fallback || "20").toLowerCase();
  if (raw === "all") return "all";
  return ["10", "20", "50", "100"].includes(raw) ? raw : String(fallback || "20");
}

function loadListPageSizes() {
  const saved = loadJson(LIST_PAGE_SIZE_KEY, {});
  if (!saved || typeof saved !== "object" || Array.isArray(saved)) return {};
  return saved;
}

function getListPageSizeValue(listId, fallback = "20") {
  return normalizeListPageSize(listPageSizes[listId], fallback);
}

function setListPageSizeValue(listId, value) {
  listPageSizes[listId] = normalizeListPageSize(value);
  saveJson(LIST_PAGE_SIZE_KEY, listPageSizes);
}

function getPagedRows(listId, rows, fallback = "20") {
  const source = Array.isArray(rows) ? rows : [];
  const sizeValue = getListPageSizeValue(listId, fallback);
  if (sizeValue === "all") {
    return { rows: source, shown: source.length, total: source.length, sizeValue };
  }
  const size = Math.max(1, Number.parseInt(sizeValue, 10) || Number.parseInt(fallback, 10) || 20);
  return { rows: source.slice(0, size), shown: Math.min(source.length, size), total: source.length, sizeValue };
}

function rerenderListById(listId) {
  if (listId === "admin-user-tbody") return renderUserTable();
  if (listId === "activity-tbody") return renderActivityHistoryPage();
  if (listId === "catalog-list-tbody") return renderCatalogList();
  if (listId === "software-device-tbody") return renderSoftwareDeviceList();
  if (listId === "software-tbody") return renderSoftwareInventoryTable();
  if (listId === "inventory-tbody") return renderSoftwareInventoryList();
  if (listId === "employee-tbody") return renderEmployeeTable();
  if (listId === "asset-tbody") return renderTrackerTable();
  if (listId === "return-history-tbody") return renderReturnHistory();
  if (listId === "audit-tbody") return renderAuditTable(auditSelection.map((id) => assets.find((a) => a.id === id)).filter(Boolean));
  if (listId === "audit-report-history-tbody") return renderAuditReportHistory();
  if (listId === "audit-status-tbody") return renderAuditStatusList();
  if (listId === "cleaning-tbody") return renderCleaningTable();
}

function initListPageSizeControls() {
  document.querySelectorAll("tbody[id]").forEach((tbody) => {
    const listId = tbody.id;
    if (!LIST_PAGE_SIZE_TARGETS.has(listId)) return;

    if (listId === "audit-report-history-tbody") {
      if (auditReportPageSize) {
        auditReportPageSize.value = getListPageSizeValue(listId, auditReportPageSize.value || "20");
      }
      return;
    }

    const tableWrap = tbody.closest(".table-wrap");
    if (!tableWrap || !tableWrap.parentElement) return;
    if (tableWrap.parentElement.querySelector(`[data-list-page-size-for="${listId}"]`)) return;

    const toolbar = document.createElement("div");
    toolbar.className = "toolbar list-page-size-toolbar";
    toolbar.dataset.listPageSizeFor = listId;
    toolbar.innerHTML = `
      <label class="inline-input">
        <span>Show</span>
        <select data-list-page-size-select="${listId}">
          <option value="10">10</option>
          <option value="20">20</option>
          <option value="50">50</option>
          <option value="100">100</option>
          <option value="all">All</option>
        </select>
      </label>
    `;
    tableWrap.parentElement.insertBefore(toolbar, tableWrap);
    const select = toolbar.querySelector("select");
    if (!select) return;
    select.value = getListPageSizeValue(listId, "20");
    select.addEventListener("change", () => {
      setListPageSizeValue(listId, select.value);
      rerenderListById(listId);
    });
  });
}

function formatDdMmYyyy(iso) {
  const value = String(iso || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return "";
  const [y, m, d] = value.split("-");
  return `${d}/${m}/${y}`;
}

function parseDdMmYyyy(text) {
  const value = String(text || "").trim();
  const m = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return "";
  const day = Number(m[1]);
  const month = Number(m[2]);
  const year = Number(m[3]);
  if (month < 1 || month > 12) return "";
  const maxDay = new Date(year, month, 0).getDate();
  if (day < 1 || day > maxDay) return "";
  return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function buildMonthOptions(selected = 1) {
  const months = [
    "01", "02", "03", "04", "05", "06",
    "07", "08", "09", "10", "11", "12"
  ];
  return months.map((m, idx) => `<option value="${m}" ${Number(selected) === idx + 1 ? "selected" : ""}>${m}</option>`).join("");
}

function buildDayOptions(year, month, selected = 1) {
  const y = Number(year) || new Date().getFullYear();
  const m = Number(month) || 1;
  const maxDay = new Date(y, m, 0).getDate();
  let html = "";
  for (let d = 1; d <= maxDay; d += 1) {
    html += `<option value="${String(d).padStart(2, "0")}" ${d === Number(selected) ? "selected" : ""}>${String(d).padStart(2, "0")}</option>`;
  }
  return html;
}

function buildYearOptions(selectedYear) {
  const nowYear = new Date().getFullYear();
  const start = nowYear - 20;
  const end = nowYear + 20;
  const chosen = Number(selectedYear) || nowYear;
  let html = "";
  for (let y = start; y <= end; y += 1) {
    html += `<option value="${y}" ${y === chosen ? "selected" : ""}>${y}</option>`;
  }
  return html;
}

function initDateInputEnhancements() {
  const dateInputs = Array.from(document.querySelectorAll("input[type='date']"));
  dateInputs.forEach((nativeInput) => {
    if (!nativeInput.id || nativeInput.dataset.enhancedDate === "1") return;
    nativeInput.dataset.enhancedDate = "1";
    nativeInput.setAttribute("lang", "en-GB");

    const wrapper = document.createElement("div");
    wrapper.className = "date-enhanced-wrap";
    wrapper.style.display = "grid";
    wrapper.style.gridTemplateColumns = "1fr 1fr 1fr";
    wrapper.style.gap = "8px";
    wrapper.style.alignItems = "center";
    wrapper.style.marginTop = "6px";

    const day = document.createElement("select");
    day.dataset.dateDayFor = nativeInput.id;

    const month = document.createElement("select");
    month.dataset.dateMonthFor = nativeInput.id;
    month.innerHTML = buildMonthOptions(1);

    const year = document.createElement("select");
    year.dataset.dateYearFor = nativeInput.id;
    year.innerHTML = buildYearOptions(new Date().getFullYear());

    const syncFromIso = (iso) => {
      const value = String(iso || "");
      if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
        const [y, m, d] = value.split("-");
        month.value = m;
        if (!Array.from(year.options).some((o) => o.value === y)) {
          year.innerHTML = buildYearOptions(Number(y));
        }
        year.value = y;
        day.innerHTML = buildDayOptions(Number(y), Number(m), Number(d));
        day.value = d;
      } else {
        day.innerHTML = buildDayOptions(Number(year.value), Number(month.value), 1);
        day.value = "01";
      }
    };

    const updateFromSelects = () => {
      let selectedDay = Number(day.value || 1);
      const currentIso = String(nativeInput.value || "");
      if (/^\d{4}-\d{2}-\d{2}$/.test(currentIso) && !day.value) selectedDay = Number(currentIso.slice(8, 10));
      const maxDay = new Date(Number(year.value), Number(month.value), 0).getDate();
      if (selectedDay > maxDay) selectedDay = maxDay;
      day.innerHTML = buildDayOptions(Number(year.value), Number(month.value), selectedDay);
      day.value = String(selectedDay).padStart(2, "0");
      const iso = `${year.value}-${month.value}-${day.value}`;
      nativeInput.value = iso;
      nativeInput.dispatchEvent(new Event("input", { bubbles: true }));
      nativeInput.dispatchEvent(new Event("change", { bubbles: true }));
    };

    day.addEventListener("change", updateFromSelects);
    month.addEventListener("change", updateFromSelects);
    year.addEventListener("change", updateFromSelects);
    nativeInput.addEventListener("change", () => syncFromIso(nativeInput.value));

    syncFromIso(nativeInput.value);
    nativeInput.style.display = "none";
    nativeInput.parentNode.insertBefore(wrapper, nativeInput.nextSibling);
    wrapper.appendChild(day);
    wrapper.appendChild(month);
    wrapper.appendChild(year);
  });
}

function enforceUppercaseOnInput(input) {
  if (!input) return;
  input.addEventListener("input", () => {
    const start = input.selectionStart;
    const end = input.selectionEnd;
    const upper = input.value.toUpperCase();
    if (upper === input.value) return;
    input.value = upper;
    if (start !== null && end !== null) input.setSelectionRange(start, end);
  });
}

function initUppercaseEnforcements() {
  [
    "catalog-serial",
    "catalog-mac",
    "catalog-edit-serial",
    "catalog-edit-mac"
  ].forEach((id) => enforceUppercaseOnInput(document.getElementById(id)));
}

function saveJson(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
  if (TRACKED_BACKUP_KEYS.has(key)) {
    persistLiveBackupSnapshot(`save:${key}`);
    scheduleSaveFolderWrite(`save:${key}`);
    scheduleCloudPush(`save:${key}`);
  }
}

function setSaveFolderLabel(text) {
  if (!saveFolderLabel) return;
  saveFolderLabel.textContent = text;
}

function canUseSaveFolderApi() {
  return typeof window.showDirectoryPicker === "function";
}

async function writeSaveFolderFile(fileName, data) {
  if (!saveFolderHandle) return;
  const fileHandle = await saveFolderHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(JSON.stringify(data, null, 2));
  await writable.close();
}

async function flushSaveFolderWrite(reason = "auto") {
  if (!saveFolderHandle || !canUseSaveFolderApi()) return;
  if (saveFolderWriteInFlight) {
    saveFolderQueuedReason = reason;
    return;
  }
  saveFolderWriteInFlight = true;
  try {
    const payload = buildFullBackupPayload(reason);
    const now = new Date().toISOString();
    const day = now.slice(0, 10);
    const envelope = {
      generatedAt: now,
      reason,
      company: currentCompany || "",
      location: currentLocation || "",
      payload
    };
    await writeSaveFolderFile("it-data-latest.json", envelope);
    await writeSaveFolderFile(`it-data-${day}.json`, envelope);
    setSaveFolderLabel(`Save folder: ${saveFolderHandle.name || "Selected"} (Last save: ${now.replace("T", " ").slice(0, 19)})`);
  } catch (error) {
    setStatus(adminSettingsStatus, `Save folder write failed: ${error.message || error}`, "err");
  } finally {
    saveFolderWriteInFlight = false;
    if (saveFolderQueuedReason) {
      const nextReason = saveFolderQueuedReason;
      saveFolderQueuedReason = "";
      void flushSaveFolderWrite(nextReason);
    }
  }
}

function scheduleSaveFolderWrite(reason = "auto") {
  if (!saveFolderHandle || !canUseSaveFolderApi()) return;
  if (saveFolderWriteTimer) clearTimeout(saveFolderWriteTimer);
  saveFolderWriteTimer = setTimeout(() => {
    void flushSaveFolderWrite(reason);
  }, 700);
}

async function selectSaveFolderByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can set save folder.", "err");
    return;
  }
  if (!canUseSaveFolderApi()) {
    setStatus(adminSettingsStatus, "Save folder is not supported in this browser.", "err");
    return;
  }
  try {
    const handle = await window.showDirectoryPicker({ mode: "readwrite" });
    saveFolderHandle = handle;
    setSaveFolderLabel(`Save folder: ${saveFolderHandle.name || "Selected"}`);
    await flushSaveFolderWrite("select-save-folder");
    setStatus(adminSettingsStatus, "Save folder selected. Data will auto-save to this folder.", "ok");
  } catch (error) {
    if (error?.name === "AbortError") return;
    setStatus(adminSettingsStatus, `Unable to select save folder: ${error.message || error}`, "err");
  }
}

async function saveToFolderNowByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can save data to folder.", "err");
    return;
  }
  if (!saveFolderHandle || !canUseSaveFolderApi()) {
    setStatus(adminSettingsStatus, "Choose save folder first.", "err");
    return;
  }
  await flushSaveFolderWrite("manual-save-folder");
  setStatus(adminSettingsStatus, "All data saved to selected folder.", "ok");
}

function loadJson(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function normalizeCloudSyncConfig(raw = {}) {
  const url = String(raw.url || "").trim().replace(/\/+$/, "");
  const anonKey = String(raw.anonKey || "").trim();
  const appId = String(raw.appId || "it-asset-tracker").trim() || "it-asset-tracker";
  const autoSync = !!raw.autoSync;
  return { url, anonKey, appId, autoSync };
}

function loadCloudSyncConfig() {
  cloudSyncConfig = normalizeCloudSyncConfig(loadJson(CLOUD_SYNC_CONFIG_KEY, {}));
  cloudLastRemoteAt = String(localStorage.getItem(CLOUD_LAST_REMOTE_AT_KEY) || "");
}

function saveCloudSyncConfig() {
  cloudSyncConfig = normalizeCloudSyncConfig({
    url: cloudUrlInput ? cloudUrlInput.value : cloudSyncConfig.url,
    anonKey: cloudAnonKeyInput ? cloudAnonKeyInput.value : cloudSyncConfig.anonKey,
    appId: cloudAppIdInput ? cloudAppIdInput.value : cloudSyncConfig.appId,
    autoSync: !!cloudAutoSyncInput?.checked
  });
  saveJson(CLOUD_SYNC_CONFIG_KEY, cloudSyncConfig);
  refreshCloudSyncForm();
}

function refreshCloudSyncForm() {
  if (cloudUrlInput) cloudUrlInput.value = cloudSyncConfig.url || "";
  if (cloudAnonKeyInput) cloudAnonKeyInput.value = cloudSyncConfig.anonKey || "";
  if (cloudAppIdInput) cloudAppIdInput.value = cloudSyncConfig.appId || "it-asset-tracker";
  if (cloudAutoSyncInput) cloudAutoSyncInput.checked = !!cloudSyncConfig.autoSync;
}

function isCloudSyncConfigured() {
  return !!(cloudSyncConfig.url && cloudSyncConfig.anonKey && cloudSyncConfig.appId);
}

function getCloudStateApiUrl() {
  const base = `${cloudSyncConfig.url}/rest/v1/app_state`;
  const params = new URLSearchParams();
  params.set("app_id", `eq.${cloudSyncConfig.appId}`);
  params.set("select", "app_id,payload,updated_at,updated_by");
  params.set("limit", "1");
  return `${base}?${params.toString()}`;
}

function getCloudUpsertApiUrl() {
  const base = `${cloudSyncConfig.url}/rest/v1/app_state`;
  const params = new URLSearchParams();
  params.set("on_conflict", "app_id");
  return `${base}?${params.toString()}`;
}

function cloudHeaders(extra = {}) {
  return {
    apikey: cloudSyncConfig.anonKey,
    Authorization: `Bearer ${cloudSyncConfig.anonKey}`,
    "Content-Type": "application/json",
    ...extra
  };
}

async function pushCloudState(reason = "manual") {
  if (!isCloudSyncConfigured()) {
    setStatus(adminSettingsStatus, "Cloud sync config missing. Save URL, key, and App ID first.", "err");
    return false;
  }
  if (cloudPushInFlight) return false;
  cloudPushInFlight = true;
  try {
    const now = new Date().toISOString();
    const payload = buildFullBackupPayload(`cloud:${reason}`);
    const body = [{
      app_id: cloudSyncConfig.appId,
      payload,
      updated_at: now,
      updated_by: sessionUser || "system"
    }];
    const response = await fetch(getCloudUpsertApiUrl(), {
      method: "POST",
      headers: cloudHeaders({
        Prefer: "resolution=merge-duplicates,return=representation"
      }),
      body: JSON.stringify(body)
    });
    if (!response.ok) {
      const txt = await response.text();
      throw new Error(`Cloud push failed (${response.status}): ${txt}`);
    }
    cloudLastRemoteAt = now;
    localStorage.setItem(CLOUD_LAST_REMOTE_AT_KEY, cloudLastRemoteAt);
    return true;
  } catch (error) {
    setStatus(adminSettingsStatus, error.message || "Cloud push failed.", "err");
    return false;
  } finally {
    cloudPushInFlight = false;
  }
}

async function pullCloudState(reason = "manual") {
  if (!isCloudSyncConfigured()) {
    setStatus(adminSettingsStatus, "Cloud sync config missing. Save URL, key, and App ID first.", "err");
    return false;
  }
  if (cloudPullInFlight) return false;
  cloudPullInFlight = true;
  try {
    const response = await fetch(getCloudStateApiUrl(), {
      method: "GET",
      headers: cloudHeaders()
    });
    if (!response.ok) {
      const txt = await response.text();
      throw new Error(`Cloud pull failed (${response.status}): ${txt}`);
    }
    const rows = await response.json();
    const row = Array.isArray(rows) ? rows[0] : null;
    if (!row || !row.payload || typeof row.payload !== "object") {
      if (reason === "manual") setStatus(adminSettingsStatus, "No cloud state found for this App ID.", "err");
      return false;
    }
    const remoteAt = String(row.updated_at || "");
    if (reason !== "manual" && remoteAt && cloudLastRemoteAt && new Date(remoteAt).getTime() <= new Date(cloudLastRemoteAt).getTime()) {
      return false;
    }
    cloudSyncSuspendLocalHooks = true;
    applyBackupPayload(row.payload, { statusEl: null, successMessage: "", logReason: "" });
    cloudSyncSuspendLocalHooks = false;
    cloudLastRemoteAt = remoteAt || new Date().toISOString();
    localStorage.setItem(CLOUD_LAST_REMOTE_AT_KEY, cloudLastRemoteAt);
    if (reason === "manual") setStatus(adminSettingsStatus, "Cloud data pulled successfully.", "ok");
    return true;
  } catch (error) {
    cloudSyncSuspendLocalHooks = false;
    setStatus(adminSettingsStatus, error.message || "Cloud pull failed.", "err");
    return false;
  } finally {
    cloudPullInFlight = false;
  }
}

function scheduleCloudPush(reason = "auto") {
  if (!cloudSyncConfig.autoSync || !isCloudSyncConfigured() || cloudSyncSuspendLocalHooks) return;
  if (cloudPushTimer) clearTimeout(cloudPushTimer);
  cloudPushTimer = setTimeout(() => {
    void pushCloudState(reason);
  }, 1200);
}

function startCloudAutoPull() {
  if (cloudPullInterval) clearInterval(cloudPullInterval);
  if (!cloudSyncConfig.autoSync || !isCloudSyncConfigured()) return;
  cloudPullInterval = setInterval(() => {
    void pullCloudState("auto");
  }, 60000);
}

function logActivity(action, details = "", meta = {}) {
  const entry = {
    id: uid(),
    at: new Date().toISOString(),
    user: meta.user || sessionUser || "system",
    action: String(action || "UNKNOWN"),
    details: String(details || ""),
    company: meta.company || currentCompany || "",
    location: meta.location || currentLocation || ""
  };
  activityLogs.push(entry);
  if (activityLogs.length > 5000) activityLogs = activityLogs.slice(activityLogs.length - 5000);
  saveJson(ACTIVITY_LOG_KEY, activityLogs);
}

function normalizeCompanyPrefix(name, inputPrefix = "") {
  const cleaned = String(inputPrefix || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (cleaned) return cleaned.slice(0, 3);
  const fromName = String(name || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
  return (fromName.slice(0, 2) || "CO").padEnd(2, "X");
}

function loadCompanyMeta() {
  const raw = loadJson(COMPANIES_KEY, null);
  if (!raw || typeof raw !== "object") {
    COMPANY_META = { ...DEFAULT_COMPANY_META };
    saveJson(COMPANIES_KEY, COMPANY_META);
    return;
  }

  const next = {};
  Object.entries(raw).forEach(([name, meta]) => {
    const companyName = String(name || "").trim();
    if (!companyName) return;
    const prefix = normalizeCompanyPrefix(companyName, meta?.prefix || "");
    const logo = typeof meta?.logo === "string" ? meta.logo : "";
    next[companyName] = { prefix, logo };
  });

  Object.entries(DEFAULT_COMPANY_META).forEach(([name, meta]) => {
    if (!next[name]) next[name] = { ...meta };
  });

  COMPANY_META = next;
  saveJson(COMPANIES_KEY, COMPANY_META);
}

function renderCompanyButtons() {
  if (!companyGrid) return;
  const names = Object.keys(COMPANY_META).sort((a, b) => a.localeCompare(b));
  companyGrid.innerHTML = names.map((name) => {
    const logo = COMPANY_META[name]?.logo || "";
    return `<button type="button" class="company-option" data-company="${escapeHtml(name)}">
      ${logo ? `<img src="${escapeHtml(logo)}" alt="${escapeHtml(name)}" onerror="this.style.display='none'" />` : ""}
      <span>${escapeHtml(name)}</span>
    </button>`;
  }).join("");
}

function normalizeTypeCode(typeName) {
  const raw = String(typeName || "").toUpperCase().replace(/[^A-Z0-9 ]/g, " ").trim();
  if (!raw) return "OTH";
  const parts = raw.split(/\s+/).filter(Boolean);
  const fromParts = parts.map((p) => p[0]).join("").slice(0, 3);
  const compact = raw.replace(/\s+/g, "");
  return (fromParts || compact.slice(0, 3) || "OTH").padEnd(3, "X").slice(0, 3);
}

function uniqueCaseInsensitive(values) {
  const map = new Map();
  values.forEach((v) => {
    const text = String(v || "").trim();
    if (!text) return;
    const key = text.toLowerCase();
    if (!map.has(key)) map.set(key, text);
  });
  return [...map.values()];
}

function getAllAssetTypes() {
  return uniqueCaseInsensitive([...DEFAULT_ASSET_TYPES, ...customAssetTypes]);
}

function getAllAssetStatuses() {
  return uniqueCaseInsensitive([...DEFAULT_STATUSES, ...customAssetStatuses]);
}

function loadCustomTypeAndStatusConfig() {
  customAssetTypes = loadJson(CUSTOM_TYPES_KEY, []);
  if (!Array.isArray(customAssetTypes)) customAssetTypes = [];
  customAssetTypes = uniqueCaseInsensitive(customAssetTypes);

  customAssetStatuses = loadJson(CUSTOM_STATUSES_KEY, []);
  if (!Array.isArray(customAssetStatuses)) customAssetStatuses = [];
  customAssetStatuses = uniqueCaseInsensitive(customAssetStatuses);

  const rawCodes = loadJson(CUSTOM_TYPE_CODES_KEY, {});
  customTypeCodes = rawCodes && typeof rawCodes === "object" ? rawCodes : {};

  saveJson(CUSTOM_TYPES_KEY, customAssetTypes);
  saveJson(CUSTOM_STATUSES_KEY, customAssetStatuses);
  saveJson(CUSTOM_TYPE_CODES_KEY, customTypeCodes);
}

function setSelectOptions(selectEl, values, includeAll = false) {
  if (!selectEl) return;
  const prev = selectEl.value;
  const opts = includeAll
    ? [`<option value="">All</option>${values.map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("")}`]
    : values.map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`);
  selectEl.innerHTML = opts.join("");
  if (values.includes(prev)) selectEl.value = prev;
  else if (includeAll) selectEl.value = "";
  else if (values.length) selectEl.value = values[0];
}

function refreshTypeAndStatusOptions() {
  const types = getAllAssetTypes();
  const statuses = getAllAssetStatuses();
  const actionStatuses = statuses.filter((s) => s !== "Returned");

  setSelectOptions(catalogFields.type, types, false);
  setSelectOptions(filterType, types, true);
  setSelectOptions(catalogFields.status, statuses, false);
  setSelectOptions(filterStatus, statuses, true);
  setSelectOptions(statusAction, actionStatuses, false);
  setSelectOptions(catalogViewStatus, statuses, true);

  if (!catalogFields.type.value && types.length) catalogFields.type.value = types[0];
  if (!catalogFields.status.value && statuses.length) catalogFields.status.value = statuses[0];
  if (!statusAction.value && actionStatuses.length) statusAction.value = actionStatuses[0];
}

function getAllLocations() {
  const merged = [...DEFAULT_LOCATIONS];
  for (const loc of locations) {
    const name = String(loc || "").trim();
    if (!name) continue;
    if (!merged.some((x) => x.toLowerCase() === name.toLowerCase())) merged.push(name);
  }
  return merged;
}

function getDefaultLocation() {
  return getAllLocations()[0] || "Site";
}

function getAllDesignations() {
  const merged = uniqueCaseInsensitive([
    ...designations,
    ...employees.map((e) => String(e.designation || "").trim()).filter(Boolean)
  ]);
  return merged.sort((a, b) => a.localeCompare(b));
}

function getAllWorkSites() {
  const merged = uniqueCaseInsensitive([
    ...workSites,
    ...employees.map((e) => String(e.siteName || "").trim()).filter(Boolean)
  ]);
  return merged.sort((a, b) => a.localeCompare(b));
}

function refreshDesignationOptions() {
  const values = getAllDesignations();
  if (employeeDesignationDatalist) {
    employeeDesignationDatalist.innerHTML = values.map((v) => `<option value="${escapeHtml(v)}"></option>`).join("");
  }
  if (deleteDesignationSelect) {
    const options = values.map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    deleteDesignationSelect.innerHTML = `<option value="">Select designation</option>${options}`;
  }
  if (renameDesignationSelect) {
    const options = values.map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    renameDesignationSelect.innerHTML = `<option value="">Select designation</option>${options}`;
  }
}

function refreshWorkSiteOptions() {
  const values = getAllWorkSites();
  if (employeeWorkSiteDatalist) {
    employeeWorkSiteDatalist.innerHTML = values.map((v) => `<option value="${escapeHtml(v)}"></option>`).join("");
  }
  if (deleteWorkSiteSelect) {
    const options = values.map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    deleteWorkSiteSelect.innerHTML = `<option value="">Select working site</option>${options}`;
  }
}

function ensureDesignationExists(value) {
  const name = String(value || "").trim();
  if (!name) return false;
  if (designations.some((d) => d.toLowerCase() === name.toLowerCase())) return false;
  designations.push(name);
  designations = uniqueCaseInsensitive(designations);
  saveJson(DESIGNATIONS_KEY, designations);
  refreshDesignationOptions();
  refreshAdminOptionLists();
  return true;
}

function ensureDepartmentExists(value) {
  const name = String(value || "").trim();
  if (!name) return false;
  if (departments.some((d) => d.toLowerCase() === name.toLowerCase())) return false;
  departments.push(name);
  refreshDepartmentSelect();
  return true;
}

function ensureWorkSiteExists(value) {
  const name = String(value || "").trim();
  if (!name) return false;
  if (workSites.some((s) => s.toLowerCase() === name.toLowerCase())) return false;
  workSites.push(name);
  workSites = uniqueCaseInsensitive(workSites);
  saveJson(WORK_SITES_KEY, workSites);
  refreshWorkSiteOptions();
  refreshAdminOptionLists();
  return true;
}

function refreshLocationOptions() {
  const allLocations = getAllLocations();
  setSelectOptions(moduleLocationSelect, allLocations, false);
  setSelectOptions(auditLocation, allLocations, false);
  setSelectOptions(auditStatusLocation, allLocations, false);
  setSelectOptions(cleaningLocation, allLocations, false);

  if (currentLocation && allLocations.includes(currentLocation)) {
    if (moduleLocationSelect) moduleLocationSelect.value = currentLocation;
    if (auditLocation) auditLocation.value = currentLocation;
    if (auditStatusLocation) auditStatusLocation.value = currentLocation;
    if (cleaningLocation) cleaningLocation.value = currentLocation;
  }
}

function refreshAdminOptionLists() {
  if (deleteDepartmentSelect) {
    const deptOptions = departments.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("");
    deleteDepartmentSelect.innerHTML = `<option value="">Select department</option>${deptOptions}`;
  }
  if (deleteAssetTypeSelect) {
    const typeOptions = getAllAssetTypes()
      .map((t) => `<option value="${escapeHtml(t)}">${escapeHtml(t)}${DEFAULT_ASSET_TYPES.includes(t) ? " (Default)" : ""}</option>`)
      .join("");
    deleteAssetTypeSelect.innerHTML = `<option value="">Select asset type</option>${typeOptions}`;
  }
  if (deleteStatusSelect) {
    const statusOptions = getAllAssetStatuses()
      .map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}${DEFAULT_STATUSES.includes(s) ? " (Default)" : ""}</option>`)
      .join("");
    deleteStatusSelect.innerHTML = `<option value="">Select status</option>${statusOptions}`;
  }
  if (deleteLocationSelect) {
    const locationOptions = getAllLocations()
      .map((l) => `<option value="${escapeHtml(l)}">${escapeHtml(l)}${DEFAULT_LOCATIONS.includes(l) ? " (Default)" : ""}</option>`)
      .join("");
    deleteLocationSelect.innerHTML = `<option value="">Select location</option>${locationOptions}`;
  }
  if (renameLocationSelect) {
    const locationOptions = getAllLocations()
      .map((l) => `<option value="${escapeHtml(l)}">${escapeHtml(l)}${DEFAULT_LOCATIONS.includes(l) ? " (Default)" : ""}</option>`)
      .join("");
    renameLocationSelect.innerHTML = `<option value="">Select location</option>${locationOptions}`;
  }
  if (renameDepartmentSelect) {
    const options = departments.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("");
    renameDepartmentSelect.innerHTML = `<option value="">Select department</option>${options}`;
  }
  if (renameAssetTypeSelect) {
    const options = getAllAssetTypes()
      .map((t) => `<option value="${escapeHtml(t)}">${escapeHtml(t)}${DEFAULT_ASSET_TYPES.includes(t) ? " (Default)" : ""}</option>`)
      .join("");
    renameAssetTypeSelect.innerHTML = `<option value="">Select asset type</option>${options}`;
  }
  if (renameStatusSelect) {
    const options = getAllAssetStatuses()
      .map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}${DEFAULT_STATUSES.includes(s) ? " (Default)" : ""}</option>`)
      .join("");
    renameStatusSelect.innerHTML = `<option value="">Select status</option>${options}`;
  }
  if (renameSoftwareSelect) {
    const options = uniqueCaseInsensitive(softwareMaster)
      .sort((a, b) => a.localeCompare(b))
      .map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`)
      .join("");
    renameSoftwareSelect.innerHTML = `<option value="">Select software</option>${options}`;
  }
  refreshDesignationOptions();
  refreshWorkSiteOptions();
}

function setStatus(el, message, type = "ok") {
  el.textContent = message;
  el.className = `status ${type}`;
}

function normalizeTag(value) {
  return String(value || "").trim().toUpperCase().replace(/\s+/g, "-");
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function buildTopValues(items, max = 12) {
  const counts = new Map();
  for (const raw of items) {
    const v = String(raw || "").trim();
    if (!v) continue;
    counts.set(v, (counts.get(v) || 0) + 1);
  }
  return Array.from(counts.entries())
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, max)
    .map(([value]) => value);
}

function hideAllFieldSuggestions() {
  Object.values(fieldSuggestionBoxes).forEach((box) => box?.classList.add("hidden"));
}

function renderFieldSuggestions(key, query = "") {
  const box = fieldSuggestionBoxes[key];
  const input = fieldSuggestionInputs[key];
  if (!box || !input) return;

  const q = query.trim().toLowerCase();
  const values = fieldSuggestionValues[key] || [];
  const matches = values.filter((v) => !q || v.toLowerCase().includes(q)).slice(0, 12);

  if (!matches.length) {
    box.innerHTML = "";
    box.classList.add("hidden");
    return;
  }

  box.innerHTML = matches
    .map((v) => `<div class="suggestion-item" data-field="${key}" data-value="${escapeHtml(v)}">${escapeHtml(v)}</div>`)
    .join("");
  box.classList.remove("hidden");
}

function setupFieldSuggestions() {
  Object.entries(fieldSuggestionInputs).forEach(([key, input]) => {
    if (!input) return;
    input.addEventListener("focus", () => renderFieldSuggestions(key, input.value));
    input.addEventListener("input", () => renderFieldSuggestions(key, input.value));
  });

  Object.entries(fieldSuggestionBoxes).forEach(([key, box]) => {
    if (!box) return;
    box.addEventListener("click", (event) => {
      const item = event.target.closest(".suggestion-item[data-value]");
      if (!item) return;
      const value = item.dataset.value || "";
      const input = fieldSuggestionInputs[key];
      if (!input) return;
      input.value = value;
      if (key === "owner") applyEmployeeDataToAssignFields(value);
      box.classList.add("hidden");
      input.focus();
    });
  });

  document.addEventListener("click", (event) => {
    const isFieldInput = Object.values(fieldSuggestionInputs).some((input) => input === event.target);
    const inFieldBox = Object.values(fieldSuggestionBoxes).some((box) => box?.contains(event.target));
    if (isFieldInput || inFieldBox) return;
    hideAllFieldSuggestions();
  });
}

function refreshAutoSuggestions() {
  const scoped = getScopedAssets();
  const scopedEmployees = getScopedEmployees();
  fieldSuggestionValues.name = buildTopValues(scoped.map((a) => a.name));
  fieldSuggestionValues.serial = buildTopValues(scoped.map((a) => a.serial));
  fieldSuggestionValues.os = buildTopValues(scoped.map((a) => a.os));
  fieldSuggestionValues.storage = buildTopValues(scoped.map((a) => a.storage));
  fieldSuggestionValues.ram = buildTopValues(scoped.map((a) => a.ram));
  fieldSuggestionValues.graphics = buildTopValues(scoped.map((a) => a.graphics));
  fieldSuggestionValues.printerIp = buildTopValues(scoped.map((a) => a.printerIp));
  fieldSuggestionValues.owner = buildTopValues(scopedEmployees.map((e) => e.name));
  refreshEmployeeNameDatalist();
}

function refreshEmployeeNameDatalist() {
  if (!employeeNameDatalist) return;
  const names = buildTopValues(getScopedEmployees().map((e) => e.name), 500);
  employeeNameDatalist.innerHTML = names.map((name) => `<option value="${escapeHtml(name)}"></option>`).join("");
}

function applyEmployeeDataToAssignFields(ownerName) {
  const picked = getScopedEmployees().find((e) => e.name.toLowerCase() === String(ownerName || "").trim().toLowerCase());
  if (!picked) return;
  if (!assignGender.value && picked.gender) assignGender.value = picked.gender;
  if (!assignDepartment.value && picked.department) assignDepartment.value = picked.department;
  if (!assignPhotoDataUrl && picked.photo) {
    assignPhotoDataUrl = picked.photo;
    updateAssignPhotoPreview(assignPhotoDataUrl);
  }
}

function applyEmployeeDataToCatalogAssignedFields(ownerName) {
  const picked = getScopedEmployees().find((e) => e.name.toLowerCase() === String(ownerName || "").trim().toLowerCase());
  if (!picked) return;
  if (catalogAssignedGender && !catalogAssignedGender.value && picked.gender) catalogAssignedGender.value = picked.gender;
  if (catalogAssignedDepartment && !catalogAssignedDepartment.value && picked.department) catalogAssignedDepartment.value = picked.department;
  if (!catalogAssignedPhotoDataUrl && picked.photo) {
    catalogAssignedPhotoDataUrl = picked.photo;
    if (catalogAssignedPhotoPreview) {
      catalogAssignedPhotoPreview.hidden = false;
      catalogAssignedPhotoPreview.src = catalogAssignedPhotoDataUrl;
    }
  }
}

function isDeviceType(type) {
  return DEVICE_TYPES.has(type);
}

function isPrinterType(type) {
  return type === "Printer";
}

function parseLegacyAsset(raw) {
  const assignment = raw.assignment || (raw.owner ? {
    owner: raw.owner,
    systemName: raw.systemName || "",
    locationDetail: raw.locationDetail || "",
    gender: raw.gender || "",
    department: raw.department || "",
    assigneePhoto: raw.assigneePhoto || "",
    assignedAt: raw.updatedAt || raw.createdAt || new Date().toISOString()
  } : null);

  return {
    id: raw.id || uid(),
    assetTag: normalizeTag(raw.assetTag),
    name: String(raw.name || "").trim(),
    type: String(raw.type || "Other"),
    serial: String(raw.serial || "").trim(),
    macAddress: String(raw.macAddress || "").trim(),
    imei: String(raw.imei || "").trim(),
    os: String(raw.os || "").trim(),
    storage: String(raw.storage || "").trim(),
    ram: String(raw.ram || "").trim(),
    graphics: String(raw.graphics || "").trim(),
    printerMode: String(raw.printerMode || "Direct"),
    printerIp: String(raw.printerIp || "").trim(),
    printerPassword: String(raw.printerPassword || ""),
    adminPassword: String(raw.adminPassword || ""),
    company: String(raw.company || "").trim(),
    location: String(raw.location || "").trim(),
    purchaseDate: String(raw.purchaseDate || ""),
    warrantyDate: String(raw.warrantyDate || ""),
    status: String(raw.status || (assignment ? "Assigned" : "Available")),
    notes: String(raw.notes || ""),
    assignment,
    returnedAt: String(raw.returnedAt || ""),
    returnReason: String(raw.returnReason || ""),
    returnNote: String(raw.returnNote || ""),
    createdAt: raw.createdAt || new Date().toISOString(),
    updatedAt: raw.updatedAt || new Date().toISOString()
  };
}

function seedDefaultAdmin() {
  users = users.map((u) => ({
    id: u.id || uid(),
    username: String(u.username || "").trim(),
    password: String(u.password || ""),
    role: u.role === "admin" ? "admin" : (String(u.username || "").trim().toLowerCase() === "ramees" ? "admin" : "user"),
    profilePhoto: String(u.profilePhoto || ""),
    createdAt: u.createdAt || new Date().toISOString()
  })).filter((u) => u.username && u.password);

  if (!users.some((u) => u.username.toLowerCase() === "ramees")) {
    users.push({
      id: uid(),
      username: "ramees",
      password: "IT@Admin",
      role: "admin",
      profilePhoto: "",
      createdAt: new Date().toISOString()
    });
  }
  saveJson(USERS_KEY, users);
}

function loadState() {
  loadCloudSyncConfig();
  loadCompanyMeta();
  loadCustomTypeAndStatusConfig();
  users = loadJson(USERS_KEY, []);
  if (!Array.isArray(users)) users = [];
  seedDefaultAdmin();
  departments = loadJson(DEPARTMENTS_KEY, ["IT", "HR", "Finance", "Operations"]);
  locations = loadJson(LOCATIONS_KEY, DEFAULT_LOCATIONS);
  if (!Array.isArray(locations)) locations = [...DEFAULT_LOCATIONS];
  locations = getAllLocations();
  saveJson(LOCATIONS_KEY, locations);
  designations = loadJson(DESIGNATIONS_KEY, []);
  if (!Array.isArray(designations)) designations = [];
  designations = uniqueCaseInsensitive(designations);
  saveJson(DESIGNATIONS_KEY, designations);
  workSites = loadJson(WORK_SITES_KEY, []);
  if (!Array.isArray(workSites)) workSites = [];
  workSites = uniqueCaseInsensitive(workSites);
  saveJson(WORK_SITES_KEY, workSites);
  employees = loadJson(EMPLOYEES_KEY, []);
  if (!Array.isArray(employees)) employees = [];
  employees = employees.map((e) => ({
    ...e,
    mobile: String(e.mobile || ""),
    email: String(e.email || ""),
    designation: String(e.designation || ""),
    siteName: String(e.siteName || "")
  }));
  employeeBulkHistory = loadJson(EMPLOYEE_BULK_LOG_KEY, []);
  if (!Array.isArray(employeeBulkHistory)) employeeBulkHistory = [];
  catalogBulkHistory = loadJson(CATALOG_BULK_LOG_KEY, []);
  if (!Array.isArray(catalogBulkHistory)) catalogBulkHistory = [];
  softwareMaster = loadJson(SOFTWARE_MASTER_KEY, []);
  if (!Array.isArray(softwareMaster)) softwareMaster = [];
  softwareMaster = uniqueCaseInsensitive(softwareMaster);
  softwareInventory = loadJson(SOFTWARE_INVENTORY_KEY, []);
  if (!Array.isArray(softwareInventory)) softwareInventory = [];
  softwareInventory = softwareInventory.filter((row) => row && row.assetId && row.softwareName);
  softwarePurchases = loadJson(SOFTWARE_PURCHASE_KEY, []);
  if (!Array.isArray(softwarePurchases)) softwarePurchases = [];
  softwarePurchases = softwarePurchases
    .filter((row) => row && row.softwareName)
    .map((row) => ({
      ...row,
      maxUsers: row.maxUsers === "" || row.maxUsers === null || row.maxUsers === undefined
        ? ""
        : String(Math.max(0, Number.parseInt(row.maxUsers, 10) || 0))
    }));
  softwareCheckRequiresAdminPassword = !!loadJson(SOFTWARE_CHECK_LOCK_KEY, false);
  activityLogs = loadJson(ACTIVITY_LOG_KEY, []);
  if (!Array.isArray(activityLogs)) activityLogs = [];
  returnLogs = loadJson(RETURN_LOG_KEY, []);
  if (!Array.isArray(returnLogs)) returnLogs = [];
  assets = loadJson(STORAGE_KEY, []).map(parseLegacyAsset).filter((a) => a.assetTag && a.name);
  audits = loadJson(AUDITS_KEY, []);
  cleaningRecords = loadJson(CLEANING_KEY, {});
  listPageSizes = loadListPageSizes();
  sessionUser = sessionStorage.getItem(SESSION_KEY) || null;
  localStorage.removeItem(SESSION_KEY);
  currentCompany = localStorage.getItem(COMPANY_KEY) || null;
  currentLocation = localStorage.getItem(LOCATION_KEY) || null;
  if (currentLocation && !getAllLocations().includes(currentLocation)) {
    currentLocation = getDefaultLocation();
    localStorage.setItem(LOCATION_KEY, currentLocation);
  }
  activePage = localStorage.getItem(MODULE_KEY) || "tracker";
}

function hardResetAllData() {
  const ok = window.confirm("Reset everything? This clears assets, audits, cleaning logs, users (except default admin), and restarts asset tags from 001.");
  if (!ok) return;
  logActivity("RESET_ALL_DATA", "Admin reset all data");

  [
    STORAGE_KEY,
    DEPARTMENTS_KEY,
    EMPLOYEES_KEY,
    EMPLOYEE_BULK_LOG_KEY,
    CATALOG_BULK_LOG_KEY,
    COMPANIES_KEY,
    CUSTOM_TYPES_KEY,
    CUSTOM_STATUSES_KEY,
    CUSTOM_TYPE_CODES_KEY,
    SOFTWARE_MASTER_KEY,
    SOFTWARE_INVENTORY_KEY,
    SOFTWARE_PURCHASE_KEY,
    DESIGNATIONS_KEY,
    WORK_SITES_KEY,
    SOFTWARE_CHECK_LOCK_KEY,
    ACTIVITY_LOG_KEY,
    RETURN_LOG_KEY,
    AUDITS_KEY,
    CLEANING_KEY,
    LOCATIONS_KEY,
    BACKUP_KEY,
    BACKUP_DATE_KEY,
    FULL_BACKUP_SNAPSHOT_KEY,
    FULL_BACKUP_SNAPSHOT_AT_KEY,
    LIST_PAGE_SIZE_KEY,
    CLOUD_LAST_REMOTE_AT_KEY,
    SESSION_KEY,
    COMPANY_KEY,
    LOCATION_KEY,
    MODULE_KEY,
    USERS_KEY
  ].forEach((k) => localStorage.removeItem(k));
  sessionStorage.removeItem(SESSION_KEY);

  users = [];
  seedDefaultAdmin();
  COMPANY_META = { ...DEFAULT_COMPANY_META };
  saveJson(COMPANIES_KEY, COMPANY_META);
  customAssetTypes = [];
  customAssetStatuses = [];
  customTypeCodes = {};
  softwareMaster = [];
  softwareInventory = [];
  softwarePurchases = [];
  softwareCheckRequiresAdminPassword = false;
  activityLogs = [];
  returnLogs = [];
  saveJson(CUSTOM_TYPES_KEY, customAssetTypes);
  saveJson(CUSTOM_STATUSES_KEY, customAssetStatuses);
  saveJson(CUSTOM_TYPE_CODES_KEY, customTypeCodes);
  saveJson(SOFTWARE_MASTER_KEY, softwareMaster);
  saveJson(SOFTWARE_INVENTORY_KEY, softwareInventory);
  saveJson(SOFTWARE_PURCHASE_KEY, softwarePurchases);
  saveJson(SOFTWARE_CHECK_LOCK_KEY, softwareCheckRequiresAdminPassword);
  saveJson(ACTIVITY_LOG_KEY, activityLogs);
  saveJson(RETURN_LOG_KEY, returnLogs);
  departments = ["IT", "HR", "Finance", "Operations"];
  locations = [...DEFAULT_LOCATIONS];
  designations = [];
  workSites = [];
  saveJson(LOCATIONS_KEY, locations);
  saveJson(DESIGNATIONS_KEY, designations);
  saveJson(WORK_SITES_KEY, workSites);
  employees = [];
  employeeBulkHistory = [];
  catalogBulkHistory = [];
  assets = [];
  audits = [];
  cleaningRecords = {};
  sessionUser = null;
  currentCompany = null;
  currentLocation = null;
  activePage = "tracker";
  selectedAssetId = null;
  editingCatalogId = null;
  assignOptionLookup = new Map();
  assignOptions = [];
  auditSelection = [];
  selectedEmployeeIds = new Set();
  listPageSizes = {};
  cloudLastRemoteAt = "";
  localStorage.removeItem(CLOUD_LAST_REMOTE_AT_KEY);
  editingEmployeeId = null;

  setStatus(moduleStatus, "All data reset. Login with ramees / IT@Admin.", "ok");
  route();
}

function buildFullBackupPayload(reason = "manual") {
  const payloadUsers = loadJson(USERS_KEY, users);
  const payloadDepartments = loadJson(DEPARTMENTS_KEY, departments);
  const payloadLocations = loadJson(LOCATIONS_KEY, locations);
  const payloadDesignations = loadJson(DESIGNATIONS_KEY, designations);
  const payloadWorkSites = loadJson(WORK_SITES_KEY, workSites);
  const payloadEmployees = loadJson(EMPLOYEES_KEY, employees);
  const payloadEmployeeBulkHistory = loadJson(EMPLOYEE_BULK_LOG_KEY, employeeBulkHistory);
  const payloadCatalogBulkHistory = loadJson(CATALOG_BULK_LOG_KEY, catalogBulkHistory);
  const payloadCompanies = loadJson(COMPANIES_KEY, COMPANY_META);
  const payloadCustomTypes = loadJson(CUSTOM_TYPES_KEY, customAssetTypes);
  const payloadCustomStatuses = loadJson(CUSTOM_STATUSES_KEY, customAssetStatuses);
  const payloadCustomTypeCodes = loadJson(CUSTOM_TYPE_CODES_KEY, customTypeCodes);
  const payloadSoftwareMaster = loadJson(SOFTWARE_MASTER_KEY, softwareMaster);
  const payloadSoftwareInventory = loadJson(SOFTWARE_INVENTORY_KEY, softwareInventory);
  const payloadSoftwarePurchases = loadJson(SOFTWARE_PURCHASE_KEY, softwarePurchases);
  const payloadSoftwareCheckLock = !!loadJson(SOFTWARE_CHECK_LOCK_KEY, softwareCheckRequiresAdminPassword);
  const payloadActivityLogs = loadJson(ACTIVITY_LOG_KEY, activityLogs);
  const payloadReturnLogs = loadJson(RETURN_LOG_KEY, returnLogs);
  const payloadAssets = loadJson(STORAGE_KEY, assets).map(parseLegacyAsset).filter((a) => a.assetTag && a.name);
  const payloadAudits = loadJson(AUDITS_KEY, audits);
  const payloadCleaning = loadJson(CLEANING_KEY, cleaningRecords);
  const payloadDailyBackups = loadJson(BACKUP_KEY, []);

  return {
    backupVersion: 1,
    reason,
    generatedAt: new Date().toISOString(),
    users: Array.isArray(payloadUsers) ? payloadUsers : users,
    departments: Array.isArray(payloadDepartments) ? payloadDepartments : departments,
    locations: Array.isArray(payloadLocations) ? payloadLocations : locations,
    designations: Array.isArray(payloadDesignations) ? payloadDesignations : designations,
    workSites: Array.isArray(payloadWorkSites) ? payloadWorkSites : workSites,
    employees: Array.isArray(payloadEmployees) ? payloadEmployees : employees,
    employeeBulkHistory: Array.isArray(payloadEmployeeBulkHistory) ? payloadEmployeeBulkHistory : employeeBulkHistory,
    catalogBulkHistory: Array.isArray(payloadCatalogBulkHistory) ? payloadCatalogBulkHistory : catalogBulkHistory,
    companies: payloadCompanies && typeof payloadCompanies === "object" ? payloadCompanies : COMPANY_META,
    customAssetTypes: Array.isArray(payloadCustomTypes) ? payloadCustomTypes : customAssetTypes,
    customAssetStatuses: Array.isArray(payloadCustomStatuses) ? payloadCustomStatuses : customAssetStatuses,
    customTypeCodes: payloadCustomTypeCodes && typeof payloadCustomTypeCodes === "object" ? payloadCustomTypeCodes : customTypeCodes,
    softwareMaster: Array.isArray(payloadSoftwareMaster) ? payloadSoftwareMaster : softwareMaster,
    softwareInventory: Array.isArray(payloadSoftwareInventory) ? payloadSoftwareInventory : softwareInventory,
    softwarePurchases: Array.isArray(payloadSoftwarePurchases) ? payloadSoftwarePurchases : softwarePurchases,
    softwareCheckLockEnabled: payloadSoftwareCheckLock,
    activityLogs: Array.isArray(payloadActivityLogs) ? payloadActivityLogs : activityLogs,
    returnLogs: Array.isArray(payloadReturnLogs) ? payloadReturnLogs : returnLogs,
    assets: Array.isArray(payloadAssets) ? payloadAssets : assets,
    audits: Array.isArray(payloadAudits) ? payloadAudits : audits,
    cleaningRecords: payloadCleaning && typeof payloadCleaning === "object" ? payloadCleaning : cleaningRecords,
    dailySnapshots: Array.isArray(payloadDailyBackups) ? payloadDailyBackups : [],
    liveSnapshotAt: localStorage.getItem(FULL_BACKUP_SNAPSHOT_AT_KEY) || null
  };
}

function persistLiveBackupSnapshot(reason = "change") {
  try {
    const payload = buildFullBackupPayload(reason);
    localStorage.setItem(FULL_BACKUP_SNAPSHOT_KEY, JSON.stringify(payload));
    localStorage.setItem(FULL_BACKUP_SNAPSHOT_AT_KEY, new Date().toISOString());
  } catch {
    // non-blocking: app should continue even if snapshot write fails
  }
}

function downloadFullBackup(reason = "manual") {
  const payload = buildFullBackupPayload(reason);
  const stamp = new Date().toISOString().replaceAll(":", "-");
  downloadFile(`it-full-backup-${stamp}.json`, JSON.stringify(payload, null, 2), "application/json");
}

function triggerChangeBackup(reason = "change") {
  void reason;
}

async function restoreFromBackupFile() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can restore backups.", "err");
    return;
  }
  const file = uploadBackupFileInput.files?.[0];
  if (!file) {
    setStatus(adminSettingsStatus, "Choose a backup JSON file first.", "err");
    return;
  }

  try {
    const parsed = JSON.parse(await file.text());
    const raw = parsed?.payload && typeof parsed.payload === "object" ? parsed.payload : parsed;
    applyBackupPayload(raw, { statusEl: adminSettingsStatus, successMessage: "Backup restored successfully.", logReason: "RESTORE_BACKUP" });
  } catch (error) {
    setStatus(adminSettingsStatus, `Restore failed: ${error.message}`, "err");
  } finally {
    uploadBackupFileInput.value = "";
  }
}

function applyBackupPayload(raw, { statusEl = null, successMessage = "", logReason = "" } = {}) {
  const nextUsers = Array.isArray(raw.users) ? raw.users : [];
  const nextDepartments = Array.isArray(raw.departments) ? raw.departments : ["IT", "HR", "Finance", "Operations"];
  const nextLocations = Array.isArray(raw.locations) ? raw.locations : [...DEFAULT_LOCATIONS];
  const nextDesignations = Array.isArray(raw.designations) ? raw.designations : [];
  const nextWorkSites = Array.isArray(raw.workSites) ? raw.workSites : [];
  const nextEmployees = Array.isArray(raw.employees) ? raw.employees : [];
  const nextEmployeeBulkHistory = Array.isArray(raw.employeeBulkHistory) ? raw.employeeBulkHistory : [];
  const nextCatalogBulkHistory = Array.isArray(raw.catalogBulkHistory) ? raw.catalogBulkHistory : [];
  const nextCompanies = raw.companies && typeof raw.companies === "object" ? raw.companies : { ...DEFAULT_COMPANY_META };
  const nextCustomTypes = Array.isArray(raw.customAssetTypes) ? raw.customAssetTypes : [];
  const nextCustomStatuses = Array.isArray(raw.customAssetStatuses) ? raw.customAssetStatuses : [];
  const nextCustomTypeCodes = raw.customTypeCodes && typeof raw.customTypeCodes === "object" ? raw.customTypeCodes : {};
  const nextSoftwareMaster = Array.isArray(raw.softwareMaster) ? raw.softwareMaster : [];
  const nextSoftwareInventory = Array.isArray(raw.softwareInventory) ? raw.softwareInventory : [];
  const nextSoftwarePurchases = Array.isArray(raw.softwarePurchases) ? raw.softwarePurchases : [];
  const nextSoftwareCheckLock = !!raw.softwareCheckLockEnabled;
  const nextActivityLogs = Array.isArray(raw.activityLogs) ? raw.activityLogs : [];
  const nextReturnLogs = Array.isArray(raw.returnLogs) ? raw.returnLogs : [];
  const nextAssets = Array.isArray(raw.assets) ? raw.assets.map(parseLegacyAsset).filter((a) => a.assetTag && a.name) : [];
  const nextAudits = Array.isArray(raw.audits) ? raw.audits : [];
  const nextCleaning = raw.cleaningRecords && typeof raw.cleaningRecords === "object" ? raw.cleaningRecords : {};

  users = nextUsers;
  seedDefaultAdmin();
  COMPANY_META = {};
  Object.entries(nextCompanies).forEach(([name, meta]) => {
    const companyName = String(name || "").trim();
    if (!companyName) return;
    COMPANY_META[companyName] = {
      prefix: normalizeCompanyPrefix(companyName, meta?.prefix || ""),
      logo: typeof meta?.logo === "string" ? meta.logo : ""
    };
  });
  Object.entries(DEFAULT_COMPANY_META).forEach(([name, meta]) => {
    if (!COMPANY_META[name]) COMPANY_META[name] = { ...meta };
  });
  customAssetTypes = uniqueCaseInsensitive(nextCustomTypes);
  customAssetStatuses = uniqueCaseInsensitive(nextCustomStatuses);
  customTypeCodes = nextCustomTypeCodes;
  softwareMaster = uniqueCaseInsensitive(nextSoftwareMaster);
  softwareInventory = nextSoftwareInventory.filter((row) => row && row.assetId && row.softwareName);
  softwarePurchases = nextSoftwarePurchases
    .filter((row) => row && row.softwareName)
    .map((row) => ({
      ...row,
      maxUsers: row.maxUsers === "" || row.maxUsers === null || row.maxUsers === undefined
        ? ""
        : String(Math.max(0, Number.parseInt(row.maxUsers, 10) || 0))
    }));
  softwareCheckRequiresAdminPassword = nextSoftwareCheckLock;
  activityLogs = nextActivityLogs;
  returnLogs = nextReturnLogs;
  departments = nextDepartments;
  locations = [...DEFAULT_LOCATIONS];
  for (const loc of nextLocations) {
    const name = String(loc || "").trim();
    if (!name) continue;
    if (!locations.some((x) => x.toLowerCase() === name.toLowerCase())) locations.push(name);
  }
  employees = nextEmployees.map((e) => ({
    ...e,
    mobile: String(e.mobile || ""),
    email: String(e.email || ""),
    designation: String(e.designation || ""),
    siteName: String(e.siteName || "")
  }));
  employeeBulkHistory = nextEmployeeBulkHistory;
  catalogBulkHistory = nextCatalogBulkHistory;
  designations = uniqueCaseInsensitive(nextDesignations);
  workSites = uniqueCaseInsensitive(nextWorkSites);
  assets = nextAssets;
  audits = nextAudits;
  cleaningRecords = nextCleaning;

  saveJson(DEPARTMENTS_KEY, departments);
  saveJson(LOCATIONS_KEY, locations);
  saveJson(DESIGNATIONS_KEY, designations);
  saveJson(WORK_SITES_KEY, workSites);
  saveJson(EMPLOYEES_KEY, employees);
  saveJson(EMPLOYEE_BULK_LOG_KEY, employeeBulkHistory);
  saveJson(CATALOG_BULK_LOG_KEY, catalogBulkHistory);
  saveJson(COMPANIES_KEY, COMPANY_META);
  saveJson(CUSTOM_TYPES_KEY, customAssetTypes);
  saveJson(CUSTOM_STATUSES_KEY, customAssetStatuses);
  saveJson(CUSTOM_TYPE_CODES_KEY, customTypeCodes);
  saveJson(SOFTWARE_MASTER_KEY, softwareMaster);
  saveJson(SOFTWARE_INVENTORY_KEY, softwareInventory);
  saveJson(SOFTWARE_PURCHASE_KEY, softwarePurchases);
  saveJson(SOFTWARE_CHECK_LOCK_KEY, softwareCheckRequiresAdminPassword);
  saveJson(ACTIVITY_LOG_KEY, activityLogs);
  saveJson(RETURN_LOG_KEY, returnLogs);
  saveJson(STORAGE_KEY, assets);
  saveJson(AUDITS_KEY, audits);
  saveJson(CLEANING_KEY, cleaningRecords);

  refreshDepartmentSelect();
  refreshDesignationOptions();
  refreshWorkSiteOptions();
  refreshTypeAndStatusOptions();
  refreshLocationOptions();
  refreshAdminOptionLists();
  renderCompanyButtons();
  if (softwareCheckLockToggle) softwareCheckLockToggle.checked = softwareCheckRequiresAdminPassword;
  refreshAutoSuggestions();
  resetCatalogForm();
  renderEmployeesPage();
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
  renderSoftwareCheckPage();
  renderSoftwareInventoryPage();
  renderActivityHistoryPage();
  renderReturnHistory();
  renderCleaningTable();
  if (logReason) logActivity(logReason, "Backup restored");
  if (statusEl && successMessage) setStatus(statusEl, successMessage, "ok");
}

function hasLocalBusinessData() {
  return (
    assets.length > 0 ||
    employees.length > 0 ||
    audits.length > 0 ||
    returnLogs.length > 0 ||
    activityLogs.length > 0 ||
    softwareInventory.length > 0 ||
    softwarePurchases.length > 0
  );
}

async function tryAutoRestoreFromSharedSaveFile() {
  if (hasLocalBusinessData()) return false;
  try {
    const response = await fetch("./save/it-data-latest.json", { cache: "no-store" });
    if (!response.ok) {
      if (window.location.protocol === "file:") {
        setStatus(authStatus, "Open using localhost (not file://) to auto-load shared save data.", "err");
      }
      return false;
    }
    const envelope = await response.json();
    const raw = envelope?.payload && typeof envelope.payload === "object" ? envelope.payload : envelope;
    if (!raw || typeof raw !== "object") return false;
    applyBackupPayload(raw, { logReason: "" });
    setStatus(authStatus, "Loaded shared data from save/it-data-latest.json.", "ok");
    return true;
  } catch {
    if (window.location.protocol === "file:") {
      setStatus(authStatus, "Open using localhost (not file://) to auto-load shared save data.", "err");
    }
    return false;
  }
}

function performLogout() {
  const userBefore = sessionUser;
  sessionUser = null;
  currentCompany = null;
  currentLocation = null;
  if (cloudPullInterval) {
    clearInterval(cloudPullInterval);
    cloudPullInterval = null;
  }
  localStorage.removeItem(SESSION_KEY);
  sessionStorage.removeItem(SESSION_KEY);
  localStorage.removeItem(COMPANY_KEY);
  localStorage.removeItem(LOCATION_KEY);
  if (userBefore) logActivity("LOGOUT", "User logged out", { user: userBefore, company: "", location: "" });
  route();
}

function persistAssets() {
  saveJson(STORAGE_KEY, assets);
  createDailyBackupSnapshot();
  refreshAutoSuggestions();
  triggerChangeBackup("assets");
  if (activePage === "software-check") renderSoftwareCheckPage();
  if (activePage === "software-inventory") renderSoftwareInventoryPage();
}

function refreshHeader() {
  sessionMeta.textContent = `User: ${sessionUser || "-"} | Company: ${currentCompany || "-"} | Location: ${currentLocation || "-"}`;
  const current = getCurrentUser();
  if (sessionAvatar) {
    if (current?.profilePhoto) {
      sessionAvatar.hidden = false;
      sessionAvatar.src = current.profilePhoto;
    } else {
      sessionAvatar.hidden = true;
      sessionAvatar.removeAttribute("src");
    }
  }
  const logo = COMPANY_META[currentCompany]?.logo;
  if (logo) {
    companyLogo.hidden = false;
    companyLogo.src = logo;
  } else {
    companyLogo.hidden = true;
    companyLogo.removeAttribute("src");
  }
  adminModuleCard?.classList.toggle("hidden", !isAdminUser());
  activityModuleCard?.classList.toggle("hidden", !isAdminUser());
  companyAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
  moduleAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
  appAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
  tabs["activity-history"]?.classList.toggle("hidden", !isAdminUser());
  moduleAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
  companyAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
  appAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
  if (addDepartmentWrap) addDepartmentWrap.classList.toggle("hidden", !isAdminUser());
  if (!isAdminUser() && (activePage === "admin-users" || activePage === "activity-history")) {
    activePage = "tracker";
    localStorage.setItem(MODULE_KEY, "tracker");
  }
  if (isAdminUser()) renderUserTable();
}

function refreshModuleMeta() {
  moduleCompanyMeta.textContent = `Company: ${currentCompany || "-"}`;
}

function refreshSettingsMeta() {
  settingsCompanyMeta.textContent = `Company: ${currentCompany || "-"} | Location: ${currentLocation || "-"}`;
}

function openAdminSettingsFromAnywhere(source = "module") {
  if (!isAdminUser()) return;
  settingsReturnScreen = source;
  showScreen("settings");
  refreshSettingsMeta();
  refreshCloudSyncForm();
  refreshAdminOptionLists();
  renderSoftwareMasterOptions();
  if (softwareCheckLockToggle) softwareCheckLockToggle.checked = softwareCheckRequiresAdminPassword;
}

function openAdminUsersFromShortcut() {
  if (!isAdminUser()) return;
  if (!currentCompany) {
    const firstCompany = Object.keys(COMPANY_META)[0] || "";
    if (!firstCompany) {
      setStatus(companyStatus, "No company configured. Add a company first.", "err");
      return;
    }
    currentCompany = firstCompany;
    localStorage.setItem(COMPANY_KEY, currentCompany);
    setStatus(companyStatus, `Auto-selected ${currentCompany} to open Admin User Management.`, "ok");
  }
  if (!currentLocation) {
    currentLocation = moduleLocationSelect?.value || getDefaultLocation();
    localStorage.setItem(LOCATION_KEY, currentLocation);
  }
  activePage = "admin-users";
  localStorage.setItem(MODULE_KEY, "admin-users");
  route();
}

async function addCompanyByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add companies.", "err");
    return;
  }
  const name = newCompanyNameInput.value.trim();
  const prefix = normalizeCompanyPrefix(name, newCompanyPrefixInput.value);
  if (!name) {
    setStatus(adminSettingsStatus, "Company name is required.", "err");
    return;
  }
  if (COMPANY_META[name]) {
    setStatus(adminSettingsStatus, "Company already exists.", "err");
    return;
  }

  let logo = "";
  const file = newCompanyLogoInput.files?.[0];
  if (file) {
    try {
      logo = await readFileAsDataUrl(file);
    } catch (error) {
      setStatus(adminSettingsStatus, error.message, "err");
      return;
    }
  }

  COMPANY_META[name] = { prefix, logo };
  saveJson(COMPANIES_KEY, COMPANY_META);
  logActivity("ADD_COMPANY", `${name} (${prefix})`);
  renderCompanyButtons();
  newCompanyNameInput.value = "";
  newCompanyPrefixInput.value = "";
  newCompanyLogoInput.value = "";
  setStatus(adminSettingsStatus, `Added company ${name} (${prefix}).`, "ok");
}

function addAssetTypeByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add asset types.", "err");
    return;
  }
  const typeName = String(newAssetTypeInput.value || "").trim();
  if (!typeName) {
    setStatus(adminSettingsStatus, "Asset type is required.", "err");
    return;
  }
  const allTypes = getAllAssetTypes();
  if (allTypes.some((t) => t.toLowerCase() === typeName.toLowerCase())) {
    setStatus(adminSettingsStatus, "Asset type already exists.", "err");
    return;
  }
  customAssetTypes.push(typeName);
  customAssetTypes = uniqueCaseInsensitive(customAssetTypes);
  customTypeCodes[typeName] = normalizeTypeCode(typeName);
  saveJson(CUSTOM_TYPES_KEY, customAssetTypes);
  saveJson(CUSTOM_TYPE_CODES_KEY, customTypeCodes);
  logActivity("ADD_ASSET_TYPE", typeName);
  refreshTypeAndStatusOptions();
  refreshAdminOptionLists();
  newAssetTypeInput.value = "";
  setStatus(adminSettingsStatus, `Added asset type ${typeName}.`, "ok");
}

function addAssetStatusByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add statuses.", "err");
    return;
  }
  const statusName = String(newAssetStatusInput.value || "").trim();
  if (!statusName) {
    setStatus(adminSettingsStatus, "Status is required.", "err");
    return;
  }
  const allStatuses = getAllAssetStatuses();
  if (allStatuses.some((s) => s.toLowerCase() === statusName.toLowerCase())) {
    setStatus(adminSettingsStatus, "Status already exists.", "err");
    return;
  }
  customAssetStatuses.push(statusName);
  customAssetStatuses = uniqueCaseInsensitive(customAssetStatuses);
  saveJson(CUSTOM_STATUSES_KEY, customAssetStatuses);
  logActivity("ADD_STATUS", statusName);
  refreshTypeAndStatusOptions();
  refreshAdminOptionLists();
  newAssetStatusInput.value = "";
  setStatus(adminSettingsStatus, `Added status ${statusName}.`, "ok");
}

function addLocationByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add locations.", "err");
    return;
  }
  const name = String(newLocationInput?.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Location is required.", "err");
    return;
  }
  if (getAllLocations().some((l) => l.toLowerCase() === name.toLowerCase())) {
    setStatus(adminSettingsStatus, "Location already exists.", "err");
    return;
  }
  locations.push(name);
  locations = getAllLocations();
  saveJson(LOCATIONS_KEY, locations);
  refreshLocationOptions();
  refreshAdminOptionLists();
  if (!currentLocation) {
    currentLocation = name;
    localStorage.setItem(LOCATION_KEY, currentLocation);
  }
  if (newLocationInput) newLocationInput.value = "";
  logActivity("ADD_LOCATION", name);
  setStatus(adminSettingsStatus, `Added location ${name}.`, "ok");
}

function deleteLocationByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can delete locations.", "err");
    return;
  }
  const name = String(deleteLocationSelect?.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Select a location to delete.", "err");
    return;
  }
  if (DEFAULT_LOCATIONS.includes(name)) {
    setStatus(adminSettingsStatus, "Default locations cannot be deleted.", "err");
    return;
  }
  const usedByAssets = assets.some((a) => String(a.location || "").toLowerCase() === name.toLowerCase());
  const usedByEmployees = employees.some((e) => String(e.location || "").toLowerCase() === name.toLowerCase());
  const usedByReturns = returnLogs.some((r) => String(r.location || "").toLowerCase() === name.toLowerCase());
  const usedByAudits = audits.some((a) => String(a.location || "").toLowerCase() === name.toLowerCase());
  const usedBySoftware = softwareInventory.some((s) => {
    const asset = assets.find((a) => a.id === s.assetId);
    return String(asset?.location || "").toLowerCase() === name.toLowerCase();
  }) || softwarePurchases.some((s) => String(s.location || "").toLowerCase() === name.toLowerCase());
  const usedByCleaning = Object.keys(cleaningRecords || {}).some((assetId) => {
    const asset = assets.find((a) => a.id === assetId);
    return String(asset?.location || "").toLowerCase() === name.toLowerCase();
  });
  if (usedByAssets || usedByEmployees || usedByReturns || usedByAudits || usedBySoftware || usedByCleaning) {
    setStatus(adminSettingsStatus, "Location is in use. Move related records before deleting.", "err");
    return;
  }
  locations = getAllLocations().filter((l) => l.toLowerCase() !== name.toLowerCase());
  saveJson(LOCATIONS_KEY, locations);
  if (currentLocation && currentLocation.toLowerCase() === name.toLowerCase()) {
    currentLocation = getDefaultLocation();
    localStorage.setItem(LOCATION_KEY, currentLocation);
  }
  refreshLocationOptions();
  refreshAdminOptionLists();
  logActivity("DELETE_LOCATION", name);
  setStatus(adminSettingsStatus, `Deleted location ${name}.`, "ok");
}

function renameLocationByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can rename locations.", "err");
    return;
  }
  const oldName = String(renameLocationSelect?.value || "").trim();
  const newName = String(renameLocationInput?.value || "").trim();
  if (!oldName) {
    setStatus(adminSettingsStatus, "Select a location to rename.", "err");
    return;
  }
  if (!newName) {
    setStatus(adminSettingsStatus, "New location name is required.", "err");
    return;
  }
  if (DEFAULT_LOCATIONS.includes(oldName)) {
    setStatus(adminSettingsStatus, "Default locations cannot be renamed.", "err");
    return;
  }
  if (oldName.toLowerCase() === newName.toLowerCase()) {
    setStatus(adminSettingsStatus, "New location must be different.", "err");
    return;
  }
  if (getAllLocations().some((l) => l.toLowerCase() === newName.toLowerCase())) {
    setStatus(adminSettingsStatus, "New location already exists.", "err");
    return;
  }

  assets = assets.map((a) => (
    String(a.location || "").toLowerCase() === oldName.toLowerCase() ? { ...a, location: newName } : a
  ));
  employees = employees.map((e) => (
    String(e.location || "").toLowerCase() === oldName.toLowerCase() ? { ...e, location: newName } : e
  ));
  returnLogs = returnLogs.map((r) => (
    String(r.location || "").toLowerCase() === oldName.toLowerCase() ? { ...r, location: newName } : r
  ));
  audits = audits.map((a) => (
    String(a.location || "").toLowerCase() === oldName.toLowerCase() ? { ...a, location: newName } : a
  ));
  softwarePurchases = softwarePurchases.map((s) => (
    String(s.location || "").toLowerCase() === oldName.toLowerCase() ? { ...s, location: newName } : s
  ));
  activityLogs = activityLogs.map((row) => (
    String(row.location || "").toLowerCase() === oldName.toLowerCase() ? { ...row, location: newName } : row
  ));
  employeeBulkHistory = employeeBulkHistory.map((b) => (
    String(b.location || "").toLowerCase() === oldName.toLowerCase() ? { ...b, location: newName } : b
  ));

  locations = getAllLocations().map((l) => (l.toLowerCase() === oldName.toLowerCase() ? newName : l));
  locations = getAllLocations();

  if (currentLocation && currentLocation.toLowerCase() === oldName.toLowerCase()) {
    currentLocation = newName;
    localStorage.setItem(LOCATION_KEY, currentLocation);
  }

  saveJson(STORAGE_KEY, assets);
  saveJson(EMPLOYEES_KEY, employees);
  saveJson(RETURN_LOG_KEY, returnLogs);
  saveJson(AUDITS_KEY, audits);
  saveJson(SOFTWARE_PURCHASE_KEY, softwarePurchases);
  saveJson(ACTIVITY_LOG_KEY, activityLogs);
  saveJson(EMPLOYEE_BULK_LOG_KEY, employeeBulkHistory);
  saveJson(LOCATIONS_KEY, locations);

  refreshLocationOptions();
  refreshAdminOptionLists();
  refreshModuleMeta();
  refreshSettingsMeta();
  refreshHeader();
  refreshAutoSuggestions();
  if (activePage === "employees") renderEmployeeTable();
  if (activePage === "catalog") renderCatalogList();
  if (activePage === "tracker") renderTrackerTable();
  if (activePage === "return") renderReturnHistory();
  if (activePage === "audit") renderAuditTable(auditSelection || []);
  if (activePage === "audit-report") renderAuditReportHistory();
  if (activePage === "audit-status") renderAuditStatusList();
  if (activePage === "cleaning") renderCleaningTable();
  if (activePage === "software-inventory") renderSoftwareInventoryList();

  if (renameLocationInput) renameLocationInput.value = "";
  logActivity("RENAME_LOCATION", `${oldName} -> ${newName}`);
  setStatus(adminSettingsStatus, `Renamed location ${oldName} to ${newName}. Records migrated.`, "ok");
}

function renameDepartmentByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can rename departments.", "err");
    return;
  }
  const oldName = String(renameDepartmentSelect?.value || "").trim();
  const newName = String(renameDepartmentInput?.value || "").trim();
  if (!oldName || !newName) {
    setStatus(adminSettingsStatus, "Select department and enter new name.", "err");
    return;
  }
  if (oldName.toLowerCase() === newName.toLowerCase()) {
    setStatus(adminSettingsStatus, "New department must be different.", "err");
    return;
  }
  if (departments.some((d) => d.toLowerCase() === newName.toLowerCase())) {
    setStatus(adminSettingsStatus, "Department already exists.", "err");
    return;
  }

  departments = departments.map((d) => (d.toLowerCase() === oldName.toLowerCase() ? newName : d));
  employees = employees.map((e) => (
    String(e.department || "").toLowerCase() === oldName.toLowerCase() ? { ...e, department: newName } : e
  ));
  assets = assets.map((a) => {
    if (!a.assignment) return a;
    if (String(a.assignment.department || "").toLowerCase() !== oldName.toLowerCase()) return a;
    return { ...a, assignment: { ...a.assignment, department: newName, updatedAt: new Date().toISOString() } };
  });
  audits = audits.map((a) => (
    String(a.department || "").toLowerCase() === oldName.toLowerCase() ? { ...a, department: newName } : a
  ));

  saveJson(DEPARTMENTS_KEY, departments);
  saveJson(EMPLOYEES_KEY, employees);
  saveJson(STORAGE_KEY, assets);
  saveJson(AUDITS_KEY, audits);
  refreshDepartmentSelect();
  refreshAdminOptionLists();
  renderEmployeeTable();
  renderTrackerTable();
  logActivity("RENAME_DEPARTMENT", `${oldName} -> ${newName}`);
  if (renameDepartmentInput) renameDepartmentInput.value = "";
  setStatus(adminSettingsStatus, `Renamed department ${oldName} to ${newName}.`, "ok");
}

function renameAssetTypeByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can rename asset types.", "err");
    return;
  }
  const oldName = String(renameAssetTypeSelect?.value || "").trim();
  const newName = String(renameAssetTypeInput?.value || "").trim();
  if (!oldName || !newName) {
    setStatus(adminSettingsStatus, "Select asset type and enter new name.", "err");
    return;
  }
  if (DEFAULT_ASSET_TYPES.includes(oldName)) {
    setStatus(adminSettingsStatus, "Default asset types cannot be renamed.", "err");
    return;
  }
  if (oldName.toLowerCase() === newName.toLowerCase()) {
    setStatus(adminSettingsStatus, "New asset type must be different.", "err");
    return;
  }
  if (getAllAssetTypes().some((t) => t.toLowerCase() === newName.toLowerCase())) {
    setStatus(adminSettingsStatus, "Asset type already exists.", "err");
    return;
  }

  customAssetTypes = customAssetTypes.map((t) => (t.toLowerCase() === oldName.toLowerCase() ? newName : t));
  if (customTypeCodes[oldName]) {
    customTypeCodes[newName] = customTypeCodes[oldName];
    delete customTypeCodes[oldName];
  }
  assets = assets.map((a) => (
    String(a.type || "").toLowerCase() === oldName.toLowerCase() ? { ...a, type: newName, updatedAt: new Date().toISOString() } : a
  ));
  softwareInventory = softwareInventory.map((row) => (
    String(row.type || "").toLowerCase() === oldName.toLowerCase() ? { ...row, type: newName } : row
  ));

  saveJson(CUSTOM_TYPES_KEY, customAssetTypes);
  saveJson(CUSTOM_TYPE_CODES_KEY, customTypeCodes);
  saveJson(STORAGE_KEY, assets);
  saveJson(SOFTWARE_INVENTORY_KEY, softwareInventory);
  refreshTypeAndStatusOptions();
  refreshAdminOptionLists();
  renderCatalogList();
  renderTrackerTable();
  logActivity("RENAME_ASSET_TYPE", `${oldName} -> ${newName}`);
  if (renameAssetTypeInput) renameAssetTypeInput.value = "";
  setStatus(adminSettingsStatus, `Renamed asset type ${oldName} to ${newName}.`, "ok");
}

function renameStatusByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can rename statuses.", "err");
    return;
  }
  const oldName = String(renameStatusSelect?.value || "").trim();
  const newName = String(renameStatusInput?.value || "").trim();
  if (!oldName || !newName) {
    setStatus(adminSettingsStatus, "Select status and enter new name.", "err");
    return;
  }
  if (DEFAULT_STATUSES.includes(oldName)) {
    setStatus(adminSettingsStatus, "Default statuses cannot be renamed.", "err");
    return;
  }
  if (oldName.toLowerCase() === newName.toLowerCase()) {
    setStatus(adminSettingsStatus, "New status must be different.", "err");
    return;
  }
  if (getAllAssetStatuses().some((s) => s.toLowerCase() === newName.toLowerCase())) {
    setStatus(adminSettingsStatus, "Status already exists.", "err");
    return;
  }

  customAssetStatuses = customAssetStatuses.map((s) => (s.toLowerCase() === oldName.toLowerCase() ? newName : s));
  assets = assets.map((a) => (
    String(a.status || "").toLowerCase() === oldName.toLowerCase() ? { ...a, status: newName, updatedAt: new Date().toISOString() } : a
  ));
  saveJson(CUSTOM_STATUSES_KEY, customAssetStatuses);
  saveJson(STORAGE_KEY, assets);
  refreshTypeAndStatusOptions();
  refreshAdminOptionLists();
  renderCatalogList();
  renderTrackerTable();
  logActivity("RENAME_STATUS", `${oldName} -> ${newName}`);
  if (renameStatusInput) renameStatusInput.value = "";
  setStatus(adminSettingsStatus, `Renamed status ${oldName} to ${newName}.`, "ok");
}

function renameSoftwareByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can rename software.", "err");
    return;
  }
  const oldName = String(renameSoftwareSelect?.value || "").trim();
  const newName = String(renameSoftwareInput?.value || "").trim();
  if (!oldName || !newName) {
    setStatus(adminSettingsStatus, "Select software and enter new name.", "err");
    return;
  }
  if (oldName.toLowerCase() === newName.toLowerCase()) {
    setStatus(adminSettingsStatus, "New software name must be different.", "err");
    return;
  }
  if (softwareMaster.some((s) => s.toLowerCase() === newName.toLowerCase())) {
    setStatus(adminSettingsStatus, "Software already exists.", "err");
    return;
  }

  softwareMaster = softwareMaster.map((s) => (s.toLowerCase() === oldName.toLowerCase() ? newName : s));
  softwareInventory = softwareInventory.map((row) => (
    String(row.softwareName || "").toLowerCase() === oldName.toLowerCase() ? { ...row, softwareName: newName } : row
  ));
  softwarePurchases = softwarePurchases.map((row) => (
    String(row.softwareName || "").toLowerCase() === oldName.toLowerCase() ? { ...row, softwareName: newName } : row
  ));
  persistSoftwareData();
  renderSoftwareMasterOptions();
  refreshAdminOptionLists();
  renderSoftwareInventoryTable();
  renderSoftwareInventoryPage();
  logActivity("RENAME_SOFTWARE", `${oldName} -> ${newName}`);
  if (renameSoftwareInput) renameSoftwareInput.value = "";
  setStatus(adminSettingsStatus, `Renamed software ${oldName} to ${newName}.`, "ok");
}

function addDesignationByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add designations.", "err");
    return;
  }
  const name = String(newDesignationInput?.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Designation is required.", "err");
    return;
  }
  if (getAllDesignations().some((d) => d.toLowerCase() === name.toLowerCase())) {
    setStatus(adminSettingsStatus, "Designation already exists.", "err");
    return;
  }
  designations.push(name);
  designations = uniqueCaseInsensitive(designations);
  saveJson(DESIGNATIONS_KEY, designations);
  refreshDesignationOptions();
  refreshAdminOptionLists();
  logActivity("ADD_DESIGNATION", name);
  if (newDesignationInput) newDesignationInput.value = "";
  setStatus(adminSettingsStatus, `Added designation ${name}.`, "ok");
}

function addWorkSiteByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add working site names.", "err");
    return;
  }
  const name = String(newWorkSiteInput?.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Working site name is required.", "err");
    return;
  }
  if (getAllWorkSites().some((s) => s.toLowerCase() === name.toLowerCase())) {
    setStatus(adminSettingsStatus, "Working site already exists.", "err");
    return;
  }
  workSites.push(name);
  workSites = uniqueCaseInsensitive(workSites);
  saveJson(WORK_SITES_KEY, workSites);
  refreshWorkSiteOptions();
  refreshAdminOptionLists();
  logActivity("ADD_WORK_SITE", name);
  if (newWorkSiteInput) newWorkSiteInput.value = "";
  setStatus(adminSettingsStatus, `Added working site ${name}.`, "ok");
}

function deleteDesignationByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can delete designations.", "err");
    return;
  }
  const name = String(deleteDesignationSelect?.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Select a designation to delete.", "err");
    return;
  }
  const usedByEmployees = employees.some((e) => String(e.designation || "").toLowerCase() === name.toLowerCase());
  if (usedByEmployees) {
    setStatus(adminSettingsStatus, "Designation is in use by employees.", "err");
    return;
  }
  designations = designations.filter((d) => d.toLowerCase() !== name.toLowerCase());
  saveJson(DESIGNATIONS_KEY, designations);
  refreshDesignationOptions();
  refreshAdminOptionLists();
  logActivity("DELETE_DESIGNATION", name);
  setStatus(adminSettingsStatus, `Deleted designation ${name}.`, "ok");
}

function deleteWorkSiteByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can delete working site names.", "err");
    return;
  }
  const name = String(deleteWorkSiteSelect?.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Select a working site to delete.", "err");
    return;
  }
  const used = employees.some((e) => String(e.siteName || "").toLowerCase() === name.toLowerCase());
  if (used) {
    setStatus(adminSettingsStatus, "Working site is in use by employees.", "err");
    return;
  }
  workSites = workSites.filter((s) => s.toLowerCase() !== name.toLowerCase());
  saveJson(WORK_SITES_KEY, workSites);
  refreshWorkSiteOptions();
  refreshAdminOptionLists();
  logActivity("DELETE_WORK_SITE", name);
  setStatus(adminSettingsStatus, `Deleted working site ${name}.`, "ok");
}

function renameDesignationByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can rename designations.", "err");
    return;
  }
  const oldName = String(renameDesignationSelect?.value || "").trim();
  const newName = String(renameDesignationInput?.value || "").trim();
  if (!oldName || !newName) {
    setStatus(adminSettingsStatus, "Select designation and enter new name.", "err");
    return;
  }
  if (oldName.toLowerCase() === newName.toLowerCase()) {
    setStatus(adminSettingsStatus, "New designation must be different.", "err");
    return;
  }
  if (getAllDesignations().some((d) => d.toLowerCase() === newName.toLowerCase())) {
    setStatus(adminSettingsStatus, "Designation already exists.", "err");
    return;
  }

  designations = uniqueCaseInsensitive(designations.map((d) => (d.toLowerCase() === oldName.toLowerCase() ? newName : d)));
  employees = employees.map((e) => (
    String(e.designation || "").toLowerCase() === oldName.toLowerCase() ? { ...e, designation: newName } : e
  ));
  saveJson(DESIGNATIONS_KEY, designations);
  saveJson(EMPLOYEES_KEY, employees);
  refreshDesignationOptions();
  refreshAdminOptionLists();
  renderEmployeeTable();
  logActivity("RENAME_DESIGNATION", `${oldName} -> ${newName}`);
  if (renameDesignationInput) renameDesignationInput.value = "";
  setStatus(adminSettingsStatus, `Renamed designation ${oldName} to ${newName}.`, "ok");
}

function deleteDepartmentByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can delete departments.", "err");
    return;
  }
  const name = String(deleteDepartmentSelect?.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Select a department to delete.", "err");
    return;
  }
  const usedByAssets = assets.some((a) => (a.assignment?.department || "").toLowerCase() === name.toLowerCase());
  const usedByEmployees = employees.some((e) => (e.department || "").toLowerCase() === name.toLowerCase());
  if (usedByAssets || usedByEmployees) {
    setStatus(adminSettingsStatus, "Department is in use. Reassign records before deleting.", "err");
    return;
  }
  departments = departments.filter((d) => d.toLowerCase() !== name.toLowerCase());
  saveJson(DEPARTMENTS_KEY, departments);
  logActivity("DELETE_DEPARTMENT", name);
  refreshDepartmentSelect();
  refreshAdminOptionLists();
  setStatus(adminSettingsStatus, `Deleted department ${name}.`, "ok");
}

function deleteAssetTypeByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can delete asset types.", "err");
    return;
  }
  const typeName = String(deleteAssetTypeSelect?.value || "").trim();
  if (!typeName) {
    setStatus(adminSettingsStatus, "Select an asset type to delete.", "err");
    return;
  }
  if (DEFAULT_ASSET_TYPES.includes(typeName)) {
    setStatus(adminSettingsStatus, "Default asset types cannot be deleted.", "err");
    return;
  }
  const usedByAssets = assets.some((a) => String(a.type || "").toLowerCase() === typeName.toLowerCase());
  if (usedByAssets) {
    setStatus(adminSettingsStatus, "Asset type is in use by existing assets.", "err");
    return;
  }
  customAssetTypes = customAssetTypes.filter((t) => t.toLowerCase() !== typeName.toLowerCase());
  delete customTypeCodes[typeName];
  saveJson(CUSTOM_TYPES_KEY, customAssetTypes);
  saveJson(CUSTOM_TYPE_CODES_KEY, customTypeCodes);
  logActivity("DELETE_ASSET_TYPE", typeName);
  refreshTypeAndStatusOptions();
  refreshAdminOptionLists();
  setStatus(adminSettingsStatus, `Deleted asset type ${typeName}.`, "ok");
}

function deleteStatusByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can delete statuses.", "err");
    return;
  }
  const statusName = String(deleteStatusSelect?.value || "").trim();
  if (!statusName) {
    setStatus(adminSettingsStatus, "Select a status to delete.", "err");
    return;
  }
  if (DEFAULT_STATUSES.includes(statusName)) {
    setStatus(adminSettingsStatus, "Default statuses cannot be deleted.", "err");
    return;
  }
  const usedByAssets = assets.some((a) => String(a.status || "").toLowerCase() === statusName.toLowerCase());
  if (usedByAssets) {
    setStatus(adminSettingsStatus, "Status is in use by existing assets.", "err");
    return;
  }
  customAssetStatuses = customAssetStatuses.filter((s) => s.toLowerCase() !== statusName.toLowerCase());
  saveJson(CUSTOM_STATUSES_KEY, customAssetStatuses);
  logActivity("DELETE_STATUS", statusName);
  refreshTypeAndStatusOptions();
  refreshAdminOptionLists();
  setStatus(adminSettingsStatus, `Deleted status ${statusName}.`, "ok");
}

function showScreen(name) {
  Object.entries(screens).forEach(([key, el]) => {
    el.classList.toggle("hidden", key !== name);
  });
}

function setActivePage(page, options = {}) {
  refreshLocationOptions();
  const previousPage = activePage;
  const targetPage = (page === "admin-users" || page === "activity-history") && !isAdminUser() ? "tracker" : page;
  if (targetPage === "software-check" && softwareCheckRequiresAdminPassword && !options.skipLockPrompt) {
    const input = window.prompt("Admin password required to open Software:");
    if (input === null) return false;
    if (!verifyAnyAdminPassword(input)) {
      setStatus(moduleStatus, "Invalid admin password. Software is locked.", "err");
      setStatus(assignStatusMsg, "Invalid admin password. Software is locked.", "err");
      return false;
    }
  }
  activePage = targetPage;
  localStorage.setItem(MODULE_KEY, targetPage);

  Object.entries(pages).forEach(([k, el]) => el?.classList.toggle("hidden", k !== targetPage));
  Object.entries(tabs).forEach(([k, el]) => el?.classList.toggle("is-active", k === targetPage));

  if (targetPage === "catalog") {
    resetCatalogForm();
    renderCatalogList();
  }
  if (targetPage === "employees") {
    renderEmployeesPage();
  }
  if (targetPage === "tracker") {
    selectedAssetId = null;
    refreshAssignAssetSelect();
    renderTrackerTable();
    renderBarcodePreview(null);
  }
  if (targetPage === "software-check") {
    renderSoftwareCheckPage();
  }
  if (targetPage === "software-inventory") {
    resetSoftwareInventoryForm();
    renderSoftwareInventoryPage();
  } else {
    closeInventoryEditModal();
  }
  if (targetPage === "return") {
    refreshReturnAssetSelect();
    clearReturnForm();
    renderReturnHistory();
  }
  if (targetPage === "audit") {
    if (auditLocation) auditLocation.value = currentLocation || auditLocation.value;
    renderAuditTable([]);
  }
  if (targetPage === "audit-report") {
    renderAuditReportAssetOptions();
    if (!selectedAuditAssetId && auditSelection.length) selectedAuditAssetId = auditSelection[0];
    fillAuditReportFormByAsset(selectedAuditAssetId);
    renderAuditReportHistory();
  }
  if (targetPage === "audit-status") {
    if (auditStatusLocation) auditStatusLocation.value = currentLocation || auditStatusLocation.value;
    renderAuditStatusList();
  }
  if (targetPage === "cleaning") renderCleaningTable();
  if (targetPage === "admin-users") renderUserTable();
  if (targetPage === "activity-history") renderActivityHistoryPage();
  if (!options.silent && sessionUser && previousPage !== targetPage) {
    logActivity("OPEN_PAGE", targetPage);
  }
  return true;
}

function route() {
  if (!sessionUser && cloudPullInterval) {
    clearInterval(cloudPullInterval);
    cloudPullInterval = null;
  }
  if (sessionUser) startCloudAutoPull();
  refreshLocationOptions();
  if (currentLocation && !getAllLocations().includes(currentLocation)) {
    currentLocation = getDefaultLocation();
    localStorage.setItem(LOCATION_KEY, currentLocation);
  }
  if (!sessionUser) {
    showScreen("auth");
    return;
  }
  if (!currentCompany) {
    showScreen("company");
    renderCompanyButtons();
    companyAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
    companyAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
    return;
  }
  if (!currentLocation) {
    showScreen("module");
    adminModuleCard?.classList.toggle("hidden", !isAdminUser());
    activityModuleCard?.classList.toggle("hidden", !isAdminUser());
    companyAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
    moduleAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
    appAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
    moduleAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
    companyAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
    appAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
    refreshModuleMeta();
    moduleLocationSelect.value = getDefaultLocation();
    return;
  }

  showScreen("app");
  moduleLocationSelect.value = currentLocation || getDefaultLocation();
  refreshHeader();
  refreshDepartmentSelect();
  refreshDesignationOptions();
  refreshWorkSiteOptions();
  refreshTypeAndStatusOptions();
  refreshAutoSuggestions();
  setActivePage(activePage, { silent: true });
}

function getCurrentUser() {
  return users.find((u) => u.username === sessionUser) || null;
}

function isAdminUser() {
  return getCurrentUser()?.role === "admin";
}

function renderUserTable() {
  if (!isAdminUser()) {
    adminUserTbody.innerHTML = "";
    return;
  }

  const allRows = [...users].sort((a, b) => a.username.localeCompare(b.username));
  const { rows } = getPagedRows("admin-user-tbody", allRows);
  adminUserTbody.innerHTML = rows.map((u) => {
    const editBtn = `<button type="button" class="ghost" data-action="edit-user" data-id="${u.id}">Edit</button>`;
    const promoteBtn = u.role === "admin"
      ? ""
      : `<button type="button" class="secondary" data-action="promote-admin" data-id="${u.id}">Make Admin</button>`;
    const deleteBtn = `<button type="button" class="danger" data-action="delete-user" data-id="${u.id}">Delete</button>`;
    const photoCell = u.profilePhoto
      ? `<img class="admin-user-thumb" src="${u.profilePhoto}" alt="${escapeHtml(u.username)}" />`
      : "";
    return `<tr>
      <td>${photoCell}</td>
      <td>${escapeHtml(u.username)}</td>
      <td>${escapeHtml(u.role)}</td>
      <td><div class="row-actions">${editBtn} ${promoteBtn} ${deleteBtn}</div></td>
    </tr>`;
  }).join("");
}

async function addUserByAdmin(event) {
  event.preventDefault();
  if (!isAdminUser()) {
    setStatus(adminUserStatus, "Only admins can manage users.", "err");
    return;
  }

  const username = newUserUsername.value.trim();
  const password = newUserPassword.value;
  const role = newUserRole.value === "admin" ? "admin" : "user";

  if (username.length < 3) return setStatus(adminUserStatus, "Username must be at least 3 characters.", "err");
  if (password.length < 4) return setStatus(adminUserStatus, "Password must be at least 4 characters.", "err");
  if (users.some((u) => u.username.toLowerCase() === username.toLowerCase())) {
    return setStatus(adminUserStatus, "Username already exists.", "err");
  }
  let profilePhoto = "";
  if (newUserPhoto?.files?.[0]) {
    try {
      profilePhoto = await readCompressedImage(newUserPhoto.files[0]);
    } catch (error) {
      setStatus(adminUserStatus, error.message || "Failed to process user photo.", "err");
      return;
    }
  }

  users.push({
    id: uid(),
    username,
    password,
    role,
    profilePhoto,
    createdAt: new Date().toISOString()
  });
  saveJson(USERS_KEY, users);
  logActivity("ADD_USER", `${username} (${role})`);
  triggerChangeBackup("users");
  adminUserForm.reset();
  renderUserTable();
  setStatus(adminUserStatus, `Created ${role}: ${username}`, "ok");
}

function promoteUserToAdmin(id) {
  if (!isAdminUser()) return;
  users = users.map((u) => (u.id === id ? { ...u, role: "admin" } : u));
  saveJson(USERS_KEY, users);
  logActivity("PROMOTE_USER", id);
  triggerChangeBackup("users");
  renderUserTable();
  setStatus(adminUserStatus, "User promoted to admin.", "ok");
}

function deleteUserByAdmin(id) {
  if (!isAdminUser()) return;
  const target = users.find((u) => u.id === id);
  if (!target) return;

  if (target.username === sessionUser) {
    setStatus(adminUserStatus, "You cannot delete your own account.", "err");
    return;
  }

  const adminCount = users.filter((u) => u.role === "admin").length;
  if (target.role === "admin" && adminCount <= 1) {
    setStatus(adminUserStatus, "Cannot delete the last admin.", "err");
    return;
  }

  if (!window.confirm(`Delete user ${target.username}?`)) return;

  users = users.filter((u) => u.id !== id);
  saveJson(USERS_KEY, users);
  logActivity("DELETE_USER", target.username);
  triggerChangeBackup("users");
  renderUserTable();
  setStatus(adminUserStatus, `Deleted user ${target.username}.`, "ok");
}

function editUserByAdmin(id) {
  if (!isAdminUser()) return;
  const target = users.find((u) => u.id === id);
  if (!target) return;
  editingAdminUserId = target.id;
  editingAdminUserPhotoDataUrl = String(target.profilePhoto || "");
  adminUserEditUsername.value = target.username;
  adminUserEditPassword.value = "";
  adminUserEditRole.value = target.role;
  if (adminUserEditPhoto) adminUserEditPhoto.value = "";
  if (editingAdminUserPhotoDataUrl) {
    adminUserEditPhotoPreview.hidden = false;
    adminUserEditPhotoPreview.src = editingAdminUserPhotoDataUrl;
  } else {
    adminUserEditPhotoPreview.hidden = true;
    adminUserEditPhotoPreview.removeAttribute("src");
  }
  setStatus(adminUserEditStatus, "", "ok");
  adminUserEditModal.classList.remove("hidden");
  adminUserEditModal.setAttribute("aria-hidden", "false");
}

function closeAdminUserEditModal() {
  if (!adminUserEditModal) return;
  adminUserEditModal.classList.add("hidden");
  adminUserEditModal.setAttribute("aria-hidden", "true");
  editingAdminUserId = null;
  editingAdminUserPhotoDataUrl = "";
}

async function saveAdminUserEdit(event) {
  event.preventDefault();
  if (!isAdminUser()) return;
  const target = users.find((u) => u.id === editingAdminUserId);
  if (!target) return;
  const nextUsername = String(adminUserEditUsername.value || "").trim();
  if (nextUsername.length < 3) {
    setStatus(adminUserEditStatus, "Username must be at least 3 characters.", "err");
    return;
  }
  if (users.some((u) => u.id !== target.id && u.username.toLowerCase() === nextUsername.toLowerCase())) {
    setStatus(adminUserEditStatus, "Username already exists.", "err");
    return;
  }
  const nextPassword = String(adminUserEditPassword.value || "");
  if (nextPassword && nextPassword.length < 4) {
    setStatus(adminUserEditStatus, "Password must be at least 4 characters.", "err");
    return;
  }
  const nextRole = adminUserEditRole.value === "admin" ? "admin" : "user";

  const adminCount = users.filter((u) => u.role === "admin").length;
  if (target.role === "admin" && nextRole !== "admin" && adminCount <= 1) {
    setStatus(adminUserEditStatus, "Cannot demote the last admin.", "err");
    return;
  }
  if (adminUserEditPhoto?.files?.[0]) {
    try {
      editingAdminUserPhotoDataUrl = await readCompressedImage(adminUserEditPhoto.files[0]);
    } catch (error) {
      setStatus(adminUserEditStatus, error.message || "Failed to process user photo.", "err");
      return;
    }
  }

  users = users.map((u) => (u.id === target.id
    ? {
      ...u,
      username: nextUsername,
      password: nextPassword || u.password,
      role: nextRole,
      profilePhoto: editingAdminUserPhotoDataUrl || ""
    }
    : u));

  if (target.username === sessionUser && nextUsername !== target.username) {
    sessionUser = nextUsername;
    sessionStorage.setItem(SESSION_KEY, sessionUser);
    localStorage.removeItem(SESSION_KEY);
    refreshHeader();
  }

  saveJson(USERS_KEY, users);
  triggerChangeBackup("users");
  logActivity("EDIT_USER", `${target.username} -> ${nextUsername} (${nextRole})`);
  renderUserTable();
  closeAdminUserEditModal();
  refreshHeader();
  setStatus(adminUserStatus, `Updated user ${nextUsername}.`, "ok");
}

function verifyLoginPassword(password) {
  const user = getCurrentUser();
  return !!(user && user.password === password);
}

function verifyAnyAdminPassword(password) {
  const input = String(password || "");
  if (!input) return false;
  return users.some((u) => u.role === "admin" && u.password === input);
}

function getTypeCode(type) {
  return TYPE_CODES[type] || customTypeCodes[type] || normalizeTypeCode(type);
}

function getScopedAssets(company = currentCompany, location = currentLocation) {
  return assets.filter((a) => a.company === company && a.location === location);
}

function getScopedEmployees(company = currentCompany, location = currentLocation) {
  return employees.filter((e) => e.company === company && e.location === location);
}

function getScopedDeviceAssets(company = currentCompany, location = currentLocation) {
  return getScopedAssets(company, location)
    .filter((a) => isDeviceType(a.type))
    .sort((a, b) => a.assetTag.localeCompare(b.assetTag));
}

function persistSoftwareData() {
  saveJson(SOFTWARE_MASTER_KEY, softwareMaster);
  saveJson(SOFTWARE_INVENTORY_KEY, softwareInventory);
  saveJson(SOFTWARE_PURCHASE_KEY, softwarePurchases);
  triggerChangeBackup("software");
}

function getScopedSoftwareInventory(company = currentCompany, location = currentLocation) {
  return softwareInventory
    .filter((row) => row.company === company && row.location === location)
    .filter((row) => {
      const asset = assets.find((a) => a.id === row.assetId);
      return !!asset && isDeviceType(asset.type);
    });
}

function getNextTag(type) {
  const prefix = COMPANY_META[currentCompany]?.prefix || "SB";
  const code = getTypeCode(type);
  const re = new RegExp(`^${prefix}-${code}-(\\d+)$`);
  let max = 0;
  for (const asset of assets) {
    const m = String(asset.assetTag).match(re);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return `${prefix}-${code}-${String(max + 1).padStart(3, "0")}`;
}

function updateCatalogTypeVisibility() {
  const type = catalogFields.type.value;
  const device = isDeviceType(type);
  const printer = isPrinterType(type);
  const assignedStatus = catalogFields.status.value === "Assigned";

  [macWrap, imeiWrap, osWrap, storageWrap, ramWrap, graphicsWrap, adminPasswordWrap].forEach((el) => {
    el.classList.toggle("hidden-field", !device);
  });

  printerModeWrap.classList.toggle("hidden-field", !printer);
  printerPasswordWrap.classList.toggle("hidden-field", !printer);
  printerIpWrap.classList.toggle("hidden-field", !(printer && catalogFields.printerMode.value === "Network Enabled"));

  if (!device) {
    catalogFields.macAddress.value = "";
    catalogFields.imei.value = "";
    catalogFields.os.value = "";
    catalogFields.storage.value = "";
    catalogFields.ram.value = "";
    catalogFields.graphics.value = "";
    catalogFields.adminPassword.value = "";
  }

  if (!printer) {
    catalogFields.printerMode.value = "Direct";
    catalogFields.printerIp.value = "";
    catalogFields.printerPassword.value = "";
  }

  [
    catalogAssignedOwnerWrap,
    catalogAssignedGenderWrap,
    catalogAssignedDepartmentWrap,
    catalogAssignedSystemWrap,
    catalogAssignedLocationWrap,
    catalogAssignedPhotoWrap,
    catalogAssignedPhotoPreviewWrap
  ].forEach((el) => el?.classList.toggle("hidden-field", !assignedStatus));

  if (!assignedStatus) {
    if (catalogAssignedOwner) catalogAssignedOwner.value = "";
    if (catalogAssignedGender) catalogAssignedGender.value = "";
    if (catalogAssignedDepartment) catalogAssignedDepartment.value = "";
    if (catalogAssignedSystem) catalogAssignedSystem.value = "";
    if (catalogAssignedLocation) catalogAssignedLocation.value = "";
    if (catalogAssignedPhotoInput) catalogAssignedPhotoInput.value = "";
    catalogAssignedPhotoDataUrl = "";
    if (catalogAssignedPhotoPreview) {
      catalogAssignedPhotoPreview.hidden = true;
      catalogAssignedPhotoPreview.removeAttribute("src");
    }
  }
}

function resetCatalogForm() {
  catalogForm.reset();
  editingCatalogId = null;
  catalogFormTitle.textContent = "Add Inventory Asset";
  catalogSaveBtn.textContent = "Save Asset";
  catalogCancelBtn.hidden = true;
  catalogFields.company.value = currentCompany || "";
  catalogFields.location.value = currentLocation || "";
  catalogFields.type.value = "Laptop";
  catalogFields.status.value = "Available";
  catalogFields.tag.value = getNextTag(catalogFields.type.value);
  catalogFields.adminPassword.type = "password";
  catalogFields.printerPassword.type = "password";
  toggleAdminInputBtn.textContent = "Show";
  catalogAssignedPhotoDataUrl = "";
  if (catalogAssignedPhotoPreview) {
    catalogAssignedPhotoPreview.hidden = true;
    catalogAssignedPhotoPreview.removeAttribute("src");
  }
  updateCatalogTypeVisibility();
  renderCatalogList();
}

function validateCatalogAsset(asset) {
  if (!asset.assetTag || !asset.name) return "Asset tag and name are required.";
  const duplicate = assets.find((a) => a.assetTag === asset.assetTag && a.id !== asset.id);
  if (duplicate) return "Asset tag already exists.";
  if (isPrinterType(asset.type) && asset.printerMode === "Network Enabled" && !asset.printerIp) {
    return "Printer IP is required for network enabled printers.";
  }
  return "";
}

function saveCatalogAsset(event) {
  event.preventDefault();

  const type = catalogFields.type.value;
  const device = isDeviceType(type);
  const printer = isPrinterType(type);

  const asset = {
    id: editingCatalogId || uid(),
    assetTag: editingCatalogId ? catalogFields.tag.value : getNextTag(type),
    name: catalogFields.name.value.trim(),
    type,
    serial: catalogFields.serial.value.trim(),
    macAddress: device ? catalogFields.macAddress.value.trim() : "",
    imei: device ? catalogFields.imei.value.trim() : "",
    os: device ? catalogFields.os.value.trim() : "",
    storage: device ? catalogFields.storage.value.trim() : "",
    ram: device ? catalogFields.ram.value.trim() : "",
    graphics: device ? catalogFields.graphics.value.trim() : "",
    printerMode: printer ? catalogFields.printerMode.value : "",
    printerIp: printer && catalogFields.printerMode.value === "Network Enabled" ? catalogFields.printerIp.value.trim() : "",
    printerPassword: printer ? catalogFields.printerPassword.value : "",
    adminPassword: device ? catalogFields.adminPassword.value : "",
    company: currentCompany,
    location: currentLocation,
    purchaseDate: catalogFields.purchaseDate.value,
    warrantyDate: catalogFields.warrantyDate.value,
    status: catalogFields.status.value,
    notes: catalogFields.notes.value.trim(),
    assignment: null,
    returnedAt: "",
    createdAt: editingCatalogId ? assets.find((a) => a.id === editingCatalogId)?.createdAt || new Date().toISOString() : new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  if (asset.status === "Assigned") {
    const owner = String(catalogAssignedOwner?.value || "").trim();
    const gender = String(catalogAssignedGender?.value || "");
    const department = String(catalogAssignedDepartment?.value || "");
    if (!owner || !gender || !department) {
      setStatus(catalogStatusMsg, "For Assigned status, owner/gender/department are required.", "err");
      return;
    }
    if (!catalogAssignedPhotoDataUrl) {
      setStatus(catalogStatusMsg, "Assigned person photo is required for Assigned status.", "err");
      return;
    }
    asset.assignment = {
      owner,
      systemName: String(catalogAssignedSystem?.value || "").trim(),
      locationDetail: String(catalogAssignedLocation?.value || "").trim(),
      gender,
      department,
      assigneePhoto: catalogAssignedPhotoDataUrl || "",
      assignedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    upsertEmployeeFromAssignment(owner, gender, department, catalogAssignedPhotoDataUrl || "");
  } else {
    asset.assignment = null;
  }

  const error = validateCatalogAsset(asset);
  if (error) {
    setStatus(catalogStatusMsg, error, "err");
    return;
  }

  if (editingCatalogId) {
    assets = assets.map((a) => (a.id === editingCatalogId ? asset : a));
    logActivity("UPDATE_ASSET", asset.assetTag);
    setStatus(catalogStatusMsg, `Updated ${asset.assetTag}.`, "ok");
  } else {
    assets.push(asset);
    logActivity("ADD_ASSET", asset.assetTag);
    setStatus(catalogStatusMsg, `Saved ${asset.assetTag}.`, "ok");
  }

  persistAssets();
  resetCatalogForm();
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
}

function updateCatalogEditTypeVisibility() {
  const type = catalogEditFields.type.value;
  const device = isDeviceType(type);
  const printer = isPrinterType(type);
  const assignedStatus = catalogEditFields.status.value === "Assigned";

  [catalogEditMacWrap, catalogEditImeiWrap, catalogEditOsWrap, catalogEditStorageWrap, catalogEditRamWrap, catalogEditGraphicsWrap, catalogEditAdminPasswordWrap]
    .forEach((el) => el?.classList.toggle("hidden-field", !device));
  catalogEditPrinterModeWrap?.classList.toggle("hidden-field", !printer);
  catalogEditPrinterPasswordWrap?.classList.toggle("hidden-field", !printer);
  catalogEditPrinterIpWrap?.classList.toggle("hidden-field", !(printer && catalogEditFields.printerMode.value === "Network Enabled"));

  if (!device) {
    catalogEditFields.macAddress.value = "";
    catalogEditFields.imei.value = "";
    catalogEditFields.os.value = "";
    catalogEditFields.storage.value = "";
    catalogEditFields.ram.value = "";
    catalogEditFields.graphics.value = "";
    catalogEditFields.adminPassword.value = "";
  }
  if (!printer) {
    catalogEditFields.printerMode.value = "Direct";
    catalogEditFields.printerIp.value = "";
    catalogEditFields.printerPassword.value = "";
  }

  [
    catalogEditAssignedOwnerWrap,
    catalogEditAssignedGenderWrap,
    catalogEditAssignedDepartmentWrap,
    catalogEditAssignedSystemWrap,
    catalogEditAssignedLocationWrap,
    catalogEditAssignedPhotoWrap,
    catalogEditAssignedPhotoPreviewWrap
  ].forEach((el) => el?.classList.toggle("hidden-field", !assignedStatus));

  if (!assignedStatus) {
    if (catalogEditAssignedOwner) catalogEditAssignedOwner.value = "";
    if (catalogEditAssignedGender) catalogEditAssignedGender.value = "";
    if (catalogEditAssignedDepartment) catalogEditAssignedDepartment.value = "";
    if (catalogEditAssignedSystem) catalogEditAssignedSystem.value = "";
    if (catalogEditAssignedLocation) catalogEditAssignedLocation.value = "";
    if (catalogEditAssignedPhotoInput) catalogEditAssignedPhotoInput.value = "";
    catalogEditAssignedPhotoDataUrl = "";
    if (catalogEditAssignedPhotoPreview) {
      catalogEditAssignedPhotoPreview.hidden = true;
      catalogEditAssignedPhotoPreview.removeAttribute("src");
    }
  }
}

function closeCatalogEditModal() {
  if (!catalogEditModal) return;
  catalogEditModal.classList.add("hidden");
  catalogEditModal.setAttribute("aria-hidden", "true");
  editingCatalogModalId = null;
  if (catalogEditForm) catalogEditForm.reset();
  catalogEditAssignedPhotoDataUrl = "";
  if (catalogEditAssignedPhotoPreview) {
    catalogEditAssignedPhotoPreview.hidden = true;
    catalogEditAssignedPhotoPreview.removeAttribute("src");
  }
  setStatus(catalogEditStatusMsg, "");
}

function openCatalogEditModal(id) {
  const asset = assets.find((a) => a.id === id);
  if (!asset || !catalogEditModal) return;

  editingCatalogModalId = id;
  setSelectOptions(catalogEditFields.type, getAllAssetTypes(), false);
  setSelectOptions(catalogEditFields.status, getAllAssetStatuses(), false);
  catalogEditModalTitle.textContent = `Edit Asset: ${asset.assetTag}`;
  catalogEditFields.tag.value = asset.assetTag;
  catalogEditFields.name.value = asset.name;
  catalogEditFields.type.value = asset.type;
  catalogEditFields.serial.value = asset.serial || "";
  catalogEditFields.macAddress.value = asset.macAddress || "";
  catalogEditFields.imei.value = asset.imei || "";
  catalogEditFields.os.value = asset.os || "";
  catalogEditFields.storage.value = asset.storage || "";
  catalogEditFields.ram.value = asset.ram || "";
  catalogEditFields.graphics.value = asset.graphics || "";
  catalogEditFields.printerMode.value = asset.printerMode || "Direct";
  catalogEditFields.printerIp.value = asset.printerIp || "";
  catalogEditFields.printerPassword.value = asset.printerPassword || "";
  catalogEditFields.adminPassword.value = asset.adminPassword || "";
  catalogEditFields.company.value = asset.company || "";
  catalogEditFields.location.value = asset.location || "";
  catalogEditFields.purchaseDate.value = asset.purchaseDate || "";
  catalogEditFields.warrantyDate.value = asset.warrantyDate || "";
  catalogEditFields.status.value = asset.status || "Available";
  catalogEditFields.notes.value = asset.notes || "";
  if (catalogEditAssignedOwner) catalogEditAssignedOwner.value = asset.assignment?.owner || "";
  if (catalogEditAssignedGender) catalogEditAssignedGender.value = asset.assignment?.gender || "";
  if (catalogEditAssignedDepartment) catalogEditAssignedDepartment.value = asset.assignment?.department || "";
  if (catalogEditAssignedSystem) catalogEditAssignedSystem.value = asset.assignment?.systemName || "";
  if (catalogEditAssignedLocation) catalogEditAssignedLocation.value = asset.assignment?.locationDetail || "";
  if (catalogEditAssignedPhotoInput) catalogEditAssignedPhotoInput.value = "";
  catalogEditAssignedPhotoDataUrl = asset.assignment?.assigneePhoto || "";
  if (catalogEditAssignedPhotoPreview) {
    if (catalogEditAssignedPhotoDataUrl) {
      catalogEditAssignedPhotoPreview.hidden = false;
      catalogEditAssignedPhotoPreview.src = catalogEditAssignedPhotoDataUrl;
    } else {
      catalogEditAssignedPhotoPreview.hidden = true;
      catalogEditAssignedPhotoPreview.removeAttribute("src");
    }
  }
  updateCatalogEditTypeVisibility();
  setStatus(catalogEditStatusMsg, "");
  catalogEditModal.classList.remove("hidden");
  catalogEditModal.setAttribute("aria-hidden", "false");
}

function saveCatalogEditModal(event) {
  event.preventDefault();
  if (!editingCatalogModalId) return;
  const existing = assets.find((a) => a.id === editingCatalogModalId);
  if (!existing) return;

  const type = catalogEditFields.type.value;
  const device = isDeviceType(type);
  const printer = isPrinterType(type);
  const asset = {
    id: existing.id,
    assetTag: catalogEditFields.tag.value,
    name: catalogEditFields.name.value.trim(),
    type,
    serial: catalogEditFields.serial.value.trim(),
    macAddress: device ? catalogEditFields.macAddress.value.trim() : "",
    imei: device ? catalogEditFields.imei.value.trim() : "",
    os: device ? catalogEditFields.os.value.trim() : "",
    storage: device ? catalogEditFields.storage.value.trim() : "",
    ram: device ? catalogEditFields.ram.value.trim() : "",
    graphics: device ? catalogEditFields.graphics.value.trim() : "",
    printerMode: printer ? catalogEditFields.printerMode.value : "",
    printerIp: printer && catalogEditFields.printerMode.value === "Network Enabled" ? catalogEditFields.printerIp.value.trim() : "",
    printerPassword: printer ? catalogEditFields.printerPassword.value : "",
    adminPassword: device ? catalogEditFields.adminPassword.value : "",
    company: existing.company,
    location: existing.location,
    purchaseDate: catalogEditFields.purchaseDate.value,
    warrantyDate: catalogEditFields.warrantyDate.value,
    status: catalogEditFields.status.value,
    notes: catalogEditFields.notes.value.trim(),
    assignment: null,
    returnedAt: "",
    createdAt: existing.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  if (asset.status === "Assigned") {
    const owner = String(catalogEditAssignedOwner?.value || "").trim();
    const gender = String(catalogEditAssignedGender?.value || "");
    const department = String(catalogEditAssignedDepartment?.value || "");
    if (!owner || !gender || !department) {
      setStatus(catalogEditStatusMsg, "For Assigned status, owner/gender/department are required.", "err");
      return;
    }
    if (!catalogEditAssignedPhotoDataUrl) {
      setStatus(catalogEditStatusMsg, "Assigned person photo is required for Assigned status.", "err");
      return;
    }
    asset.assignment = {
      owner,
      systemName: String(catalogEditAssignedSystem?.value || "").trim(),
      locationDetail: String(catalogEditAssignedLocation?.value || "").trim(),
      gender,
      department,
      assigneePhoto: catalogEditAssignedPhotoDataUrl || "",
      assignedAt: existing.assignment?.assignedAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    upsertEmployeeFromAssignment(owner, gender, department, catalogEditAssignedPhotoDataUrl || "");
  }

  const error = validateCatalogAsset(asset);
  if (error) {
    setStatus(catalogEditStatusMsg, error, "err");
    return;
  }

  assets = assets.map((a) => (a.id === editingCatalogModalId ? asset : a));
  logActivity("UPDATE_ASSET", asset.assetTag);
  persistAssets();
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
  closeCatalogEditModal();
  setStatus(catalogStatusMsg, `Updated ${asset.assetTag}.`, "ok");
}

function startCatalogEdit(id) {
  setActivePage("catalog");
  openCatalogEditModal(id);
}

function deleteOneAsset(id) {
  const asset = assets.find((a) => a.id === id);
  if (!asset || !window.confirm(`Delete ${asset.assetTag}?`)) return;

  assets = assets.filter((a) => a.id !== id);
  selectedCatalogAssetIds.delete(id);
  logActivity("DELETE_ASSET", asset.assetTag);
  purgeSoftwareForDeletedAssets(new Set(assets.map((a) => a.id)));
  persistAssets();
  renderTrackerTable();
  refreshAssignAssetSelect();
  renderCatalogList();
  setStatus(catalogStatusMsg, `Deleted ${asset.assetTag}.`, "ok");
}

function deleteScopedAssets() {
  const scoped = getScopedAssets();
  if (!scoped.length) {
    setStatus(catalogStatusMsg, "No scoped assets to delete.", "err");
    return;
  }
  if (!window.confirm(`Delete all assets for ${currentCompany} / ${currentLocation}?`)) return;

  const ids = new Set(scoped.map((a) => a.id));
  assets = assets.filter((a) => !ids.has(a.id));
  selectedCatalogAssetIds = new Set([...selectedCatalogAssetIds].filter((id) => !ids.has(id)));
  logActivity("DELETE_SCOPED_ASSETS", `${currentCompany}/${currentLocation} count=${scoped.length}`);
  purgeSoftwareForDeletedAssets(new Set(assets.map((a) => a.id)));
  persistAssets();
  resetCatalogForm();
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
  setStatus(catalogStatusMsg, "Deleted all scoped assets.", "ok");
}

function getVisibleCatalogRows() {
  const scoped = getScopedAssets().sort((a, b) => a.assetTag.localeCompare(b.assetTag));
  const statusFilter = catalogViewStatus.value;
  const search = String(catalogSearchInput?.value || "").trim().toLowerCase();
  return scoped
    .filter((a) => !statusFilter || a.status === statusFilter)
    .filter((a) => {
      if (!search) return true;
      const hay = [
        a.assetTag,
        a.name,
        a.type,
        a.serial,
        a.status,
        a.assignment?.owner,
        a.macAddress,
        a.imei
      ].map((v) => String(v || "").toLowerCase()).join(" ");
      return hay.includes(search);
    });
}

function toggleSelectAllCatalogVisible() {
  const rows = getPagedRows("catalog-list-tbody", getVisibleCatalogRows()).rows;
  if (!rows.length) return;
  const ids = rows.map((a) => a.id);
  const allSelected = ids.every((id) => selectedCatalogAssetIds.has(id));
  if (allSelected) ids.forEach((id) => selectedCatalogAssetIds.delete(id));
  else ids.forEach((id) => selectedCatalogAssetIds.add(id));
  renderCatalogList();
}

function deleteSelectedCatalogAssets() {
  if (!selectedCatalogAssetIds.size) {
    setStatus(catalogStatusMsg, "Select assets to delete.", "err");
    return;
  }
  if (!window.confirm(`Delete ${selectedCatalogAssetIds.size} selected asset(s)?`)) return;

  const before = assets.length;
  assets = assets.filter((a) => !selectedCatalogAssetIds.has(a.id));
  selectedCatalogAssetIds = new Set();
  const removed = before - assets.length;
  if (!removed) {
    setStatus(catalogStatusMsg, "No selected assets found to delete.", "err");
    return;
  }
  logActivity("DELETE_SELECTED_ASSETS", `count=${removed} ${currentCompany}/${currentLocation}`);
  purgeSoftwareForDeletedAssets(new Set(assets.map((a) => a.id)));
  persistAssets();
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
  setStatus(catalogStatusMsg, `Deleted ${removed} selected asset(s).`, "ok");
}

function renderCatalogList() {
  refreshCatalogBulkHistoryOptions();
  const scopedIds = new Set(getScopedAssets().map((a) => a.id));
  selectedCatalogAssetIds = new Set([...selectedCatalogAssetIds].filter((id) => scopedIds.has(id)));
  const allRows = getVisibleCatalogRows();
  const { rows } = getPagedRows("catalog-list-tbody", allRows);

  if (!allRows.length) {
    catalogListTbody.innerHTML = `<tr><td colspan="10">No assets for selected filters in current company/location.</td></tr>`;
    return;
  }

  catalogListTbody.innerHTML = rows.map((a) => `
    <tr>
      <td><input type="checkbox" data-action="catalog-select" data-id="${a.id}" ${selectedCatalogAssetIds.has(a.id) ? "checked" : ""} /></td>
      <td>${escapeHtml(a.assetTag)}</td>
      <td>${escapeHtml(a.name)}</td>
      <td>${escapeHtml(a.type)}</td>
      <td>${escapeHtml(a.serial || "-")}</td>
      <td>${escapeHtml(a.status)}</td>
      <td>${escapeHtml(a.assignment?.owner || "-")}</td>
      <td>${escapeHtml(a.company)}</td>
      <td>${escapeHtml(a.location)}</td>
      <td>
        <button type="button" class="ghost" data-action="catalog-edit" data-id="${a.id}">Edit</button>
        ${isDeviceType(a.type) ? `<button type="button" class="secondary" data-action="catalog-software" data-id="${a.id}">Software</button>` : ""}
        <button type="button" class="danger" data-action="catalog-delete" data-id="${a.id}">Delete</button>
      </td>
    </tr>
  `).join("");
}

function refreshDepartmentSelect() {
  departments = Array.from(new Set(departments.map((d) => d.trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b));
  saveJson(DEPARTMENTS_KEY, departments);
  assignDepartment.innerHTML = `<option value="">Select</option>${departments.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("")}`;
  if (catalogAssignedDepartment) {
    catalogAssignedDepartment.innerHTML = `<option value="">Select</option>${departments.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("")}`;
  }
  if (catalogEditAssignedDepartment) {
    catalogEditAssignedDepartment.innerHTML = `<option value="">Select</option>${departments.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("")}`;
  }
  if (employeeDepartmentDatalist) {
    employeeDepartmentDatalist.innerHTML = departments.map((d) => `<option value="${escapeHtml(d)}"></option>`).join("");
  }
  if (employeeDepartment) {
    if (employeeDepartment.tagName === "SELECT") {
      employeeDepartment.innerHTML = `<option value="">Select</option>${departments.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("")}`;
    }
  }
  if (employeeEditDepartment) {
    if (employeeEditDepartment.tagName === "SELECT") {
      employeeEditDepartment.innerHTML = `<option value="">Select</option>${departments.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("")}`;
    }
  }
  refreshAdminOptionLists();
}

function addDepartment() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add departments.", "err");
    return;
  }
  const name = newDepartmentInput.value.trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Department name is required.", "err");
    return;
  }
  if (departments.some((d) => d.toLowerCase() === name.toLowerCase())) {
    setStatus(adminSettingsStatus, "Department already exists.", "err");
    return;
  }
  departments.push(name);
  refreshDepartmentSelect();
  triggerChangeBackup("departments");
  assignDepartment.value = name;
  if (employeeDepartment && !employeeDepartment.value) employeeDepartment.value = name;
  newDepartmentInput.value = "";
  logActivity("ADD_DEPARTMENT", name);
  setStatus(adminSettingsStatus, `Added department ${name}.`, "ok");
}

function renderEmployeesPage() {
  if (!employeeForm) return;
  const isSiteLocation = currentLocation === "Site";
  employeeCompany.value = currentCompany || "";
  employeeLocation.value = currentLocation || "";
  if (employeeSiteNameWrap) employeeSiteNameWrap.classList.toggle("hidden", !isSiteLocation);
  if (employeeSiteNameCol) employeeSiteNameCol.classList.toggle("hidden", !isSiteLocation);
  if (employeeCsvHeadersHint) {
    employeeCsvHeadersHint.textContent = isSiteLocation
      ? "Headers: name,employeeId,mobile,email,designation,gender,department,siteName"
      : "Headers: name,employeeId,mobile,email,designation,gender,department";
  }
  if (!isSiteLocation && employeeSiteNameInput) employeeSiteNameInput.value = "";
  employeePhotoDataUrl = "";
  editingEmployeeId = null;
  if (employeeCancelBtn) employeeCancelBtn.hidden = true;
  if (employeeSaveBtn) employeeSaveBtn.textContent = "Add Employee";
  if (employeePhotoPreview) {
    employeePhotoPreview.hidden = true;
    employeePhotoPreview.removeAttribute("src");
  }
  selectedEmployeeIds = new Set();
  refreshEmployeeBulkHistoryOptions();
  renderEmployeeTable();
}

function renderEmployeeTable() {
  if (!employeeTbody) return;
  const isSiteLocation = currentLocation === "Site";
  if (employeeSiteNameCol) employeeSiteNameCol.classList.toggle("hidden", !isSiteLocation);
  const allRows = getFilteredScopedEmployees();
  const { rows } = getPagedRows("employee-tbody", allRows);
  if (!allRows.length) {
    employeeTbody.innerHTML = `<tr><td colspan="${isSiteLocation ? 13 : 12}">No employees in current company/location.</td></tr>`;
    return;
  }

  employeeTbody.innerHTML = rows.map((e) => `
    <tr>
      <td><input type="checkbox" data-action="select-employee" data-id="${e.id}" ${selectedEmployeeIds.has(e.id) ? "checked" : ""} /></td>
      <td>${e.photo ? `<img class="assignee-thumb employee-photo-thumb" src="${e.photo}" alt="${escapeHtml(e.name)}" data-action="view-employee-photo" data-id="${e.id}" />` : "-"}</td>
      <td>${escapeHtml(e.name)}</td>
      <td>${escapeHtml(e.employeeId || "-")}</td>
      <td>${escapeHtml(e.mobile || "-")}</td>
      <td>${escapeHtml(e.email || "-")}</td>
      <td>${escapeHtml(e.designation || "-")}</td>
      <td>${escapeHtml(e.gender || "-")}</td>
      <td>${escapeHtml(e.department || "-")}</td>
      ${isSiteLocation ? `<td>${escapeHtml(e.siteName || "-")}</td>` : ""}
      <td>${escapeHtml(e.company || "-")}</td>
      <td>${escapeHtml(e.location || "-")}</td>
      <td class="employee-row-actions">
        <button type="button" class="ghost" data-action="edit-employee" data-id="${e.id}">Edit</button>
        <button type="button" class="secondary" data-action="edit-employee-photo" data-id="${e.id}">${e.photo ? "Edit Photo" : "Add Photo"}</button>
        <button type="button" class="danger" data-action="delete-employee" data-id="${e.id}">Delete</button>
      </td>
    </tr>
  `).join("");
}

function getFilteredScopedEmployees() {
  const q = String(employeeSearchInput?.value || "").trim().toLowerCase();
  const rows = getScopedEmployees()
    .filter((e) => {
      if (!q) return true;
      const hay = `${e.name || ""} ${e.employeeId || ""} ${e.mobile || ""} ${e.email || ""} ${e.designation || ""} ${e.department || ""} ${e.siteName || ""}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a, b) => a.name.localeCompare(b.name));
  const sortMode = String(employeeSortSelect?.value || "name-asc");
  rows.sort((a, b) => {
    if (sortMode === "name-desc") return String(b.name || "").localeCompare(String(a.name || ""));
    if (sortMode === "site-asc") {
      const siteCmp = String(a.siteName || "").localeCompare(String(b.siteName || ""));
      if (siteCmp !== 0) return siteCmp;
      return String(a.name || "").localeCompare(String(b.name || ""));
    }
    if (sortMode === "site-desc") {
      const siteCmp = String(b.siteName || "").localeCompare(String(a.siteName || ""));
      if (siteCmp !== 0) return siteCmp;
      return String(a.name || "").localeCompare(String(b.name || ""));
    }
    return String(a.name || "").localeCompare(String(b.name || ""));
  });
  return rows;
}

function refreshEmployeeBulkHistoryOptions() {
  if (!employeeBulkHistorySelect) return;
  const scopedIds = new Set(getScopedEmployees().map((e) => e.id));
  const items = employeeBulkHistory
    .filter((b) => (b.company === currentCompany && b.location === currentLocation))
    .filter((b) => Array.isArray(b.employeeIds) && b.employeeIds.some((id) => scopedIds.has(id)))
    .sort((a, b) => String(b.at || "").localeCompare(String(a.at || "")));
  employeeBulkHistorySelect.innerHTML = `<option value="">Select upload batch</option>${items
    .map((b) => `<option value="${b.id}">${escapeHtml((b.at || "").replace("T", " ").slice(0, 19))} - ${b.count || b.employeeIds.length} employees</option>`)
    .join("")}`;
}

async function addEmployee(event) {
  event.preventDefault();
  const wasEdit = !!editingEmployeeId;
  const name = employeeName.value.trim();
  const empId = employeeIdInput.value.trim();
  const mobile = String(employeeMobileInput?.value || "").trim();
  const email = String(employeeEmailInput?.value || "").trim();
  const designation = String(employeeDesignationInput?.value || "").trim();
  const gender = employeeGender.value;
  const department = employeeDepartment.value;
  const siteName = currentLocation === "Site" ? String(employeeSiteNameInput?.value || "").trim() : "";

  if (!name) {
    setStatus(employeeStatusMsg, "Employee name is required.", "err");
    return;
  }

  const duplicate = getScopedEmployees().find((e) => e.id !== editingEmployeeId && e.name.toLowerCase() === name.toLowerCase() && (!empId || String(e.employeeId || "").toLowerCase() === empId.toLowerCase()));
  if (duplicate) {
    setStatus(employeeStatusMsg, "Employee already exists in this location.", "err");
    return;
  }

  if (employeePhotoInput?.files?.[0]) {
    try {
      employeePhotoDataUrl = await readCompressedImage(employeePhotoInput.files[0]);
    } catch (error) {
      setStatus(employeeStatusMsg, error.message, "err");
      return;
    }
  }

  const payload = {
    id: editingEmployeeId || uid(),
    name,
    employeeId: empId,
    mobile,
    email,
    designation,
    gender,
    department,
    siteName,
    photo: employeePhotoDataUrl || "",
    company: currentCompany,
    location: currentLocation,
    createdAt: editingEmployeeId ? (employees.find((e) => e.id === editingEmployeeId)?.createdAt || new Date().toISOString()) : new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  if (editingEmployeeId) {
    employees = employees.map((e) => (e.id === editingEmployeeId ? payload : e));
  } else {
    employees.push(payload);
  }
  if (siteName) ensureWorkSiteExists(siteName);
  if (department) ensureDepartmentExists(department);
  if (designation) ensureDesignationExists(designation);
  saveJson(EMPLOYEES_KEY, employees);
  refreshAutoSuggestions();
  renderEmployeeTable();
  employeeForm.reset();
  employeePhotoDataUrl = "";
  if (employeePhotoPreview) {
    employeePhotoPreview.hidden = true;
    employeePhotoPreview.removeAttribute("src");
  }
  employeeCompany.value = currentCompany || "";
  employeeLocation.value = currentLocation || "";
  editingEmployeeId = null;
  if (employeeCancelBtn) employeeCancelBtn.hidden = true;
  if (employeeSaveBtn) employeeSaveBtn.textContent = "Add Employee";
  setStatus(employeeStatusMsg, `${wasEdit ? "Updated" : "Added"} employee ${name}.`, "ok");
}

function startEmployeeEdit(id) {
  const employee = employees.find((e) => e.id === id && e.company === currentCompany && e.location === currentLocation);
  if (!employee) return;
  editingEmployeeModalId = id;
  employeeEditPhotoDataUrl = employee.photo || "";
  if (employeeEditName) employeeEditName.value = employee.name || "";
  if (employeeEditId) employeeEditId.value = employee.employeeId || "";
  if (employeeEditMobile) employeeEditMobile.value = employee.mobile || "";
  if (employeeEditEmail) employeeEditEmail.value = employee.email || "";
  if (employeeEditDesignation) employeeEditDesignation.value = employee.designation || "";
  if (employeeEditGender) employeeEditGender.value = employee.gender || "";
  if (employeeEditDepartment) employeeEditDepartment.value = employee.department || "";
  if (employeeEditCompany) employeeEditCompany.value = employee.company || currentCompany || "";
  if (employeeEditLocation) employeeEditLocation.value = employee.location || currentLocation || "";
  if (employeeEditSiteName) employeeEditSiteName.value = employee.siteName || "";
  if (employeeEditSiteNameWrap) employeeEditSiteNameWrap.classList.toggle("hidden", currentLocation !== "Site");
  if (employeeEditPhoto) employeeEditPhoto.value = "";
  if (employeeEditPhotoPreview) {
    if (employeeEditPhotoDataUrl) {
      employeeEditPhotoPreview.hidden = false;
      employeeEditPhotoPreview.src = employeeEditPhotoDataUrl;
    } else {
      employeeEditPhotoPreview.hidden = true;
      employeeEditPhotoPreview.removeAttribute("src");
    }
  }
  setStatus(employeeEditStatusMsg, "", "ok");
  employeeEditModal?.classList.remove("hidden");
}

function closeEmployeeEditModal() {
  editingEmployeeModalId = null;
  employeeEditPhotoDataUrl = "";
  if (employeeEditForm) employeeEditForm.reset();
  if (employeeEditPhotoPreview) {
    employeeEditPhotoPreview.hidden = true;
    employeeEditPhotoPreview.removeAttribute("src");
  }
  setStatus(employeeEditStatusMsg, "", "ok");
  employeeEditModal?.classList.add("hidden");
}

async function saveEmployeeEditModal(event) {
  event.preventDefault();
  if (!editingEmployeeModalId) return;
  const employee = employees.find((e) => e.id === editingEmployeeModalId);
  if (!employee) {
    setStatus(employeeEditStatusMsg, "Employee not found.", "err");
    return;
  }

  const name = String(employeeEditName?.value || "").trim();
  const employeeId = String(employeeEditId?.value || "").trim();
  const mobile = String(employeeEditMobile?.value || "").trim();
  const email = String(employeeEditEmail?.value || "").trim();
  const designation = String(employeeEditDesignation?.value || "").trim();
  const gender = String(employeeEditGender?.value || "");
  const department = String(employeeEditDepartment?.value || "");
  const siteName = currentLocation === "Site" ? String(employeeEditSiteName?.value || "").trim() : "";

  if (!name) {
    setStatus(employeeEditStatusMsg, "Employee name is required.", "err");
    return;
  }

  const duplicate = getScopedEmployees().find((e) => e.id !== employee.id && e.name.toLowerCase() === name.toLowerCase() && (!employeeId || String(e.employeeId || "").toLowerCase() === employeeId.toLowerCase()));
  if (duplicate) {
    setStatus(employeeEditStatusMsg, "Employee already exists in this location.", "err");
    return;
  }

  if (employeeEditPhoto?.files?.[0]) {
    try {
      employeeEditPhotoDataUrl = await readCompressedImage(employeeEditPhoto.files[0]);
    } catch (error) {
      setStatus(employeeEditStatusMsg, error.message, "err");
      return;
    }
  }

  if (siteName) ensureWorkSiteExists(siteName);
  if (department) ensureDepartmentExists(department);
  if (designation) ensureDesignationExists(designation);

  employees = employees.map((e) => {
    if (e.id !== employee.id) return e;
    return {
      ...e,
      name,
      employeeId,
      mobile,
      email,
      designation,
      gender,
      department,
      siteName,
      photo: employeeEditPhotoDataUrl || "",
      updatedAt: new Date().toISOString()
    };
  });
  saveJson(EMPLOYEES_KEY, employees);
  refreshAutoSuggestions();
  renderEmployeeTable();
  refreshEmployeeBulkHistoryOptions();
  closeEmployeeEditModal();
  setStatus(employeeStatusMsg, `Updated employee ${name}.`, "ok");
}

function cancelEmployeeEdit() {
  editingEmployeeId = null;
  employeeForm.reset();
  employeePhotoDataUrl = "";
  if (employeePhotoPreview) {
    employeePhotoPreview.hidden = true;
    employeePhotoPreview.removeAttribute("src");
  }
  employeeCompany.value = currentCompany || "";
  employeeLocation.value = currentLocation || "";
  if (employeeSaveBtn) employeeSaveBtn.textContent = "Add Employee";
  if (employeeCancelBtn) employeeCancelBtn.hidden = true;
}

function openEmployeePhotoModal(id) {
  const employee = employees.find((e) => e.id === id);
  if (!employee?.photo || !employeePhotoModal || !employeePhotoModalImage) return;
  employeePhotoModalImage.src = employee.photo;
  employeePhotoModal.classList.remove("hidden");
}

function closeEmployeePhotoModal() {
  if (!employeePhotoModal) return;
  employeePhotoModal.classList.add("hidden");
  if (employeePhotoModalImage) employeePhotoModalImage.removeAttribute("src");
}

function upsertEmployeeFromAssignment(ownerName, gender, department, photo) {
  const name = String(ownerName || "").trim();
  if (!name) return;
  if (department) ensureDepartmentExists(department);

  const existing = getScopedEmployees().find((e) => e.name.toLowerCase() === name.toLowerCase());
  if (existing) {
    let changed = false;
    if (!existing.gender && gender) {
      existing.gender = gender;
      changed = true;
    }
    if (!existing.department && department) {
      existing.department = department;
      changed = true;
    }
    if (!existing.photo && photo) {
      existing.photo = photo;
      changed = true;
    }
    if (changed) {
      saveJson(EMPLOYEES_KEY, employees);
      refreshAutoSuggestions();
      if (activePage === "employees") renderEmployeeTable();
    }
    return;
  }

  employees.push({
    id: uid(),
    name,
    employeeId: "",
    gender: gender || "",
    department: department || "",
    photo: photo || "",
    company: currentCompany,
    location: currentLocation,
    createdAt: new Date().toISOString()
  });
  saveJson(EMPLOYEES_KEY, employees);
  refreshAutoSuggestions();
  if (activePage === "employees") renderEmployeeTable();
}

function downloadEmployeeSampleCsv() {
  const isSiteLocation = currentLocation === "Site";
  const sample = isSiteLocation
    ? [
      "name,employeeId,mobile,email,designation,gender,department,siteName",
      "Ramees,E001,9876543210,ramees@example.com,IT Admin,Male,IT,Site-A",
      "Aisha,E002,9000000002,aisha@example.com,Accounts,Female,Finance,Site-B",
      "Rahul,E003,9000000003,rahul@example.com,Supervisor,Male,Operations,Site-C"
    ].join("\n")
    : [
      "name,employeeId,mobile,email,designation,gender,department",
      "Ramees,E001,9876543210,ramees@example.com,IT Admin,Male,IT",
      "Aisha,E002,9000000002,aisha@example.com,Accounts,Female,Finance",
      "Rahul,E003,9000000003,rahul@example.com,Supervisor,Male,Operations"
    ].join("\n");
  downloadFile("employees-bulk-sample.csv", sample, "text/csv");
  setStatus(employeeStatusMsg, "Employee sample CSV downloaded.", "ok");
}

async function importEmployeeCsv(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const rows = parseCsv(await file.text());
    let added = 0;
    let skipped = 0;
    const batchIds = [];

    for (const row of rows) {
      const name = String(row.name || "").trim();
      const employeeId = String(row.employeeId || "").trim();
      const gender = String(row.gender || "").trim();
      const department = String(row.department || "").trim();
      const mobile = String(row.mobile || "").trim();
      const email = String(row.email || "").trim();
      const designation = String(row.designation || "").trim();
      const siteName = currentLocation === "Site" ? String(row.siteName || "").trim() : "";

      if (!name) {
        skipped += 1;
        continue;
      }

      const duplicate = getScopedEmployees().find(
        (e) => e.name.toLowerCase() === name.toLowerCase()
          && (!employeeId || String(e.employeeId || "").toLowerCase() === employeeId.toLowerCase())
      );
      if (duplicate) {
        skipped += 1;
        continue;
      }

      const nextId = uid();
      employees.push({
        id: nextId,
        name,
        employeeId,
        mobile,
        email,
        designation,
        gender,
        department,
        siteName,
        photo: "",
        company: currentCompany,
        location: currentLocation,
        createdAt: new Date().toISOString()
      });
      if (siteName) ensureWorkSiteExists(siteName);
      if (department) ensureDepartmentExists(department);
      if (designation) ensureDesignationExists(designation);
      batchIds.push(nextId);
      added += 1;
    }

    saveJson(EMPLOYEES_KEY, employees);
    if (batchIds.length) {
      employeeBulkHistory.push({
        id: uid(),
        at: new Date().toISOString(),
        company: currentCompany,
        location: currentLocation,
        count: batchIds.length,
        employeeIds: batchIds
      });
      saveJson(EMPLOYEE_BULK_LOG_KEY, employeeBulkHistory);
    }
    refreshAutoSuggestions();
    renderEmployeeTable();
    refreshEmployeeBulkHistoryOptions();
    setStatus(employeeStatusMsg, `Employee CSV imported. Added: ${added}, Skipped: ${skipped}.`, "ok");
  } catch (error) {
    setStatus(employeeStatusMsg, `CSV import failed: ${error.message}`, "err");
  } finally {
    event.target.value = "";
  }
}

function deleteEmployee(id) {
  const target = employees.find((e) => e.id === id);
  if (!target) return;
  if (!window.confirm(`Delete employee ${target.name}?`)) return;
  employees = employees.filter((e) => e.id !== id);
  saveJson(EMPLOYEES_KEY, employees);
  refreshAutoSuggestions();
  renderEmployeeTable();
  refreshEmployeeBulkHistoryOptions();
  setStatus(employeeStatusMsg, `Deleted employee ${target.name}.`, "ok");
}

function toggleSelectAllEmployees() {
  const rows = getPagedRows("employee-tbody", getFilteredScopedEmployees()).rows;
  const allSelected = rows.length > 0 && rows.every((e) => selectedEmployeeIds.has(e.id));
  if (allSelected) {
    selectedEmployeeIds = new Set();
  } else {
    selectedEmployeeIds = new Set(rows.map((e) => e.id));
  }
  renderEmployeeTable();
}

function deleteSelectedEmployees() {
  if (!selectedEmployeeIds.size) {
    setStatus(employeeStatusMsg, "Select employees to delete.", "err");
    return;
  }
  if (!window.confirm(`Delete ${selectedEmployeeIds.size} selected employee(s)?`)) return;
  employees = employees.filter((e) => !selectedEmployeeIds.has(e.id));
  selectedEmployeeIds = new Set();
  saveJson(EMPLOYEES_KEY, employees);
  renderEmployeeTable();
  refreshEmployeeBulkHistoryOptions();
  setStatus(employeeStatusMsg, "Selected employees deleted.", "ok");
}

function undoLastEmployeeBulkUpload() {
  const scoped = employeeBulkHistory
    .filter((b) => b.company === currentCompany && b.location === currentLocation)
    .sort((a, b) => String(b.at || "").localeCompare(String(a.at || "")));
  const last = scoped[0];
  if (!last) {
    setStatus(employeeStatusMsg, "No bulk upload history found.", "err");
    return;
  }
  const idSet = new Set(last.employeeIds || []);
  const before = employees.length;
  employees = employees.filter((e) => !idSet.has(e.id));
  employeeBulkHistory = employeeBulkHistory.filter((b) => b.id !== last.id);
  saveJson(EMPLOYEES_KEY, employees);
  saveJson(EMPLOYEE_BULK_LOG_KEY, employeeBulkHistory);
  const removed = before - employees.length;
  renderEmployeeTable();
  refreshEmployeeBulkHistoryOptions();
  setStatus(employeeStatusMsg, `Undo complete. Removed ${removed} employees from last bulk upload.`, "ok");
}

function undoAllEmployeeBulkUploads() {
  const scopedBatches = employeeBulkHistory.filter((b) => b.company === currentCompany && b.location === currentLocation);
  if (!scopedBatches.length) {
    setStatus(employeeStatusMsg, "No bulk upload history found.", "err");
    return;
  }
  if (!window.confirm(`Undo all bulk uploads for ${currentCompany} / ${currentLocation}?`)) return;
  const idSet = new Set(scopedBatches.flatMap((b) => b.employeeIds || []));
  const before = employees.length;
  employees = employees.filter((e) => !idSet.has(e.id));
  employeeBulkHistory = employeeBulkHistory.filter((b) => !(b.company === currentCompany && b.location === currentLocation));
  saveJson(EMPLOYEES_KEY, employees);
  saveJson(EMPLOYEE_BULK_LOG_KEY, employeeBulkHistory);
  selectedEmployeeIds = new Set();
  const removed = before - employees.length;
  renderEmployeeTable();
  refreshEmployeeBulkHistoryOptions();
  setStatus(employeeStatusMsg, `Undo complete. Removed ${removed} employees from all bulk uploads.`, "ok");
}

function deleteSelectedEmployeeBulkUpload() {
  const batchId = String(employeeBulkHistorySelect?.value || "");
  if (!batchId) {
    setStatus(employeeStatusMsg, "Select a bulk upload batch.", "err");
    return;
  }
  const batch = employeeBulkHistory.find((b) => b.id === batchId);
  if (!batch) {
    setStatus(employeeStatusMsg, "Batch not found.", "err");
    return;
  }
  if (!window.confirm(`Delete employees from selected bulk upload (${batch.count || 0})?`)) return;
  const idSet = new Set(batch.employeeIds || []);
  const before = employees.length;
  employees = employees.filter((e) => !idSet.has(e.id));
  employeeBulkHistory = employeeBulkHistory.filter((b) => b.id !== batchId);
  saveJson(EMPLOYEES_KEY, employees);
  saveJson(EMPLOYEE_BULK_LOG_KEY, employeeBulkHistory);
  const removed = before - employees.length;
  renderEmployeeTable();
  refreshEmployeeBulkHistoryOptions();
  setStatus(employeeStatusMsg, `Removed ${removed} employees from selected bulk upload.`, "ok");
}

function startEmployeePhotoEdit(id) {
  const employee = employees.find((e) => e.id === id);
  if (!employee) return;
  editingEmployeePhotoId = id;
  employeePhotoEditInput?.click();
}

async function saveEditedEmployeePhoto() {
  const file = employeePhotoEditInput?.files?.[0];
  if (!file || !editingEmployeePhotoId) return;
  try {
    const photo = await readCompressedImage(file);
    let updatedName = "";
    employees = employees.map((e) => {
      if (e.id !== editingEmployeePhotoId) return e;
      updatedName = e.name;
      return { ...e, photo };
    });
    saveJson(EMPLOYEES_KEY, employees);
    refreshAutoSuggestions();
    renderEmployeeTable();
    setStatus(employeeStatusMsg, `Photo updated for ${updatedName}.`, "ok");
  } catch (error) {
    setStatus(employeeStatusMsg, error.message, "err");
  } finally {
    editingEmployeePhotoId = null;
    if (employeePhotoEditInput) employeePhotoEditInput.value = "";
  }
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Unable to read image file."));
    reader.readAsDataURL(file);
  });
}

function compressImageDataUrl(dataUrl, maxWidth = 320, maxHeight = 320, quality = 0.78) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(maxWidth / img.width, maxHeight / img.height, 1);
      const w = Math.max(1, Math.round(img.width * scale));
      const h = Math.max(1, Math.round(img.height * scale));
      const canvas = document.createElement("canvas");
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        reject(new Error("Unable to process image."));
        return;
      }
      ctx.drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL("image/jpeg", quality));
    };
    img.onerror = () => reject(new Error("Unable to process image."));
    img.src = dataUrl;
  });
}

async function readCompressedImage(file) {
  const raw = await readFileAsDataUrl(file);
  return compressImageDataUrl(raw, 320, 320, 0.78);
}

function updateAssignPhotoPreview(dataUrl) {
  if (!dataUrl) {
    assignPhotoPreview.hidden = true;
    assignPhotoPreview.removeAttribute("src");
    return;
  }
  assignPhotoPreview.hidden = false;
  assignPhotoPreview.src = dataUrl;
}

function refreshAssignAssetSelect() {
  const scoped = getScopedAssets().sort((a, b) => a.assetTag.localeCompare(b.assetTag));
  assignOptionLookup = new Map();
  assignOptions = [];

  if (!scoped.length) {
    assignAssetSuggestions.innerHTML = "";
    assignAssetSuggestions.classList.add("hidden");
    assignAssetSelect.value = "";
    fillAssignForm(null);
    return;
  }

  scoped.forEach((a) => {
    const label = `${a.assetTag} - ${a.name}`;
    assignOptionLookup.set(label, a.id);
    assignOptions.push({
      id: a.id,
      label,
      tag: a.assetTag,
      owner: a.assignment?.owner || ""
    });
  });

  const selected = selectedAssetId ? scoped.find((a) => a.id === selectedAssetId) : null;
  if (selected) {
    assignAssetSelect.value = `${selected.assetTag} - ${selected.name}`;
    fillAssignForm(selected);
    return;
  }

  assignAssetSelect.value = "";
  fillAssignForm(null);
}

function refreshReturnAssetSelect() {
  const scoped = getScopedAssets().sort((a, b) => a.assetTag.localeCompare(b.assetTag));
  returnOptionLookup = new Map();
  returnOptions = [];

  scoped.forEach((a) => {
    const label = `${a.assetTag} - ${a.name}`;
    returnOptionLookup.set(label, a.id);
    returnOptions.push({ id: a.id, label, tag: a.assetTag });
  });

  if (!scoped.length) {
    returnAssetSuggestions.innerHTML = "";
    returnAssetSuggestions.classList.add("hidden");
    returnAssetSelect.value = "";
    fillReturnForm(null);
  }
}

function fillReturnForm(asset) {
  if (!asset) {
    returnAssetName.value = "";
    returnAssetType.value = "";
    returnCurrentOwner.value = "";
    returnCurrentStatus.value = "";
    returnDateInput.value = new Date().toISOString().slice(0, 10);
    returnReasonSelect.value = "";
    returnNoteInput.value = "";
    return;
  }

  returnAssetName.value = asset.name;
  returnAssetType.value = asset.type;
  returnCurrentOwner.value = asset.assignment?.owner || "-";
  returnCurrentStatus.value = asset.status;
  returnDateInput.value = asset.returnedAt ? asset.returnedAt.slice(0, 10) : new Date().toISOString().slice(0, 10);
  returnReasonSelect.value = asset.returnReason || "";
  returnNoteInput.value = asset.returnNote || "";
}

function clearReturnForm() {
  returnAssetSelect.value = "";
  returnAssetSuggestions.innerHTML = "";
  returnAssetSuggestions.classList.add("hidden");
  fillReturnForm(null);
}

function getScopedReturnLogs(company = currentCompany, location = currentLocation) {
  return returnLogs
    .filter((r) => r.company === company && r.location === location)
    .sort((a, b) => String(b.returnedAt || "").localeCompare(String(a.returnedAt || "")));
}

function renderReturnHistory() {
  if (!returnHistoryTbody) return;
  const allRows = getScopedReturnLogs();
  const { rows } = getPagedRows("return-history-tbody", allRows);
  if (!allRows.length) {
    returnHistoryTbody.innerHTML = `<tr><td colspan="10">No return history in current company/location.</td></tr>`;
    return;
  }
  returnHistoryTbody.innerHTML = rows.map((r) => `<tr>
    <td>${fmtDate(r.returnedAt)}</td>
    <td>${escapeHtml(r.assetTag || "-")}</td>
    <td>${escapeHtml(r.assetName || "-")}</td>
    <td>${escapeHtml(r.assetType || "-")}</td>
    <td>${escapeHtml(r.previousOwner || "-")}</td>
    <td>${escapeHtml(r.reason || "-")}</td>
    <td>${escapeHtml(r.note || "-")}</td>
    <td>${escapeHtml(r.company || "-")}</td>
    <td>${escapeHtml(r.location || "-")}</td>
    <td>${escapeHtml(r.returnedBy || "-")}</td>
  </tr>`).join("");
}

function exportReturnHistoryCsv() {
  const rows = getScopedReturnLogs().map((r) => ({
    returnDate: fmtDate(r.returnedAt),
    assetTag: r.assetTag || "",
    assetName: r.assetName || "",
    assetType: r.assetType || "",
    previousOwner: r.previousOwner || "",
    reason: r.reason || "",
    note: r.note || "",
    company: r.company || "",
    location: r.location || "",
    returnedBy: r.returnedBy || ""
  }));
  const csv = toCsvRows(
    ["returnDate", "assetTag", "assetName", "assetType", "previousOwner", "reason", "note", "company", "location", "returnedBy"],
    rows
  );
  downloadFile(`return-history-${currentCompany || "all"}-${currentLocation || "all"}-${todayKey()}.csv`, csv, "text/csv");
  setStatus(returnStatusMsg, `Exported ${rows.length} return rows.`, "ok");
}

function readReturnAsset() {
  const raw = returnAssetSelect.value.trim();
  if (!raw) return null;

  const lookupId = returnOptionLookup.get(raw);
  if (lookupId) return assets.find((a) => a.id === lookupId) || null;

  const typedTag = normalizeTag(raw.includes(" - ") ? raw.split(" - ")[0] : raw);
  return getScopedAssets().find((a) => a.assetTag === typedTag) || null;
}

function renderReturnSuggestions(query = "") {
  const q = query.trim().toLowerCase();
  const matches = returnOptions
    .filter((o) => !q || o.label.toLowerCase().includes(q) || o.tag.toLowerCase().includes(q))
    .slice(0, 12);

  if (!matches.length) {
    returnAssetSuggestions.innerHTML = "";
    returnAssetSuggestions.classList.add("hidden");
    return;
  }

  returnAssetSuggestions.innerHTML = matches
    .map((m) => `<div class="suggestion-item" data-id="${m.id}" data-label="${escapeHtml(m.label)}">${escapeHtml(m.label)}</div>`)
    .join("");
  returnAssetSuggestions.classList.remove("hidden");
}

function fillAssignForm(asset) {
  if (!asset) {
    assignAssetName.value = "";
    assignAssetType.value = "";
    assignSystemName.value = "";
    assignLocationDetail.value = "";
    assignOwner.value = "";
    assignGender.value = "";
    assignDepartment.value = "";
    assignStatus.value = "";
    statusAction.value = "Assigned";
    assignPhoto.value = "";
    assignPhotoDataUrl = "";
    updateAssignPhotoPreview("");
    return;
  }

  assignAssetName.value = asset.name;
  assignAssetType.value = asset.type;
  assignSystemName.value = asset.assignment?.systemName || "";
  assignLocationDetail.value = asset.assignment?.locationDetail || "";
  assignOwner.value = asset.assignment?.owner || "";
  assignGender.value = asset.assignment?.gender || "";
  assignDepartment.value = asset.assignment?.department || "";
  assignStatus.value = asset.status;
  statusAction.value = asset.status;
  assignPhoto.value = "";
  assignPhotoDataUrl = asset.assignment?.assigneePhoto || "";
  updateAssignPhotoPreview(assignPhotoDataUrl);
}

function clearAssignActionPanel() {
  selectedAssetId = null;
  assignAssetSelect.value = "";
  assignAssetSuggestions.innerHTML = "";
  assignAssetSuggestions.classList.add("hidden");
  fillAssignForm(null);
  renderBarcodePreview(null);
}

function readAssignAsset() {
  const raw = assignAssetSelect.value.trim();
  if (!raw) return null;

  const lookupId = assignOptionLookup.get(raw);
  if (lookupId) return assets.find((a) => a.id === lookupId) || null;

  const typedTag = normalizeTag(raw.includes(" - ") ? raw.split(" - ")[0] : raw);
  return getScopedAssets().find((a) => a.assetTag === typedTag) || null;
}

function findScopedAssetByTag(rawTag, { deviceOnly = false } = {}) {
  const tag = normalizeTag(rawTag);
  if (!tag) return null;
  const list = getScopedAssets();
  const hit = list.find((a) => a.assetTag === tag) || null;
  if (!hit) return null;
  if (deviceOnly && !isDeviceType(hit.type)) return null;
  return hit;
}

function renderAssignSuggestions(query = "") {
  const q = query.trim().toLowerCase();
  const matches = assignOptions
    .filter((o) => (
      !q ||
      o.label.toLowerCase().includes(q) ||
      o.tag.toLowerCase().includes(q) ||
      o.owner.toLowerCase().includes(q)
    ))
    .slice(0, 12);

  if (!matches.length) {
    assignAssetSuggestions.innerHTML = "";
    assignAssetSuggestions.classList.add("hidden");
    return;
  }

  assignAssetSuggestions.innerHTML = matches
    .map((m) => `<div class="suggestion-item" data-id="${m.id}" data-label="${escapeHtml(m.label)}">${escapeHtml(m.label)}</div>`)
    .join("");
  assignAssetSuggestions.classList.remove("hidden");
}

async function assignOrUpdateAsset() {
  const asset = readAssignAsset();
  if (!asset) {
    setStatus(assignStatusMsg, "Select a valid asset.", "err");
    return;
  }

  if (assignPhoto.files?.[0]) {
    try {
      assignPhotoDataUrl = await readFileAsDataUrl(assignPhoto.files[0]);
      updateAssignPhotoPreview(assignPhotoDataUrl);
    } catch (error) {
      setStatus(assignStatusMsg, error.message, "err");
      return;
    }
  }

  const owner = assignOwner.value.trim();
  const systemName = assignSystemName.value.trim();
  const locationDetail = assignLocationDetail.value.trim();
  const gender = assignGender.value;
  const department = assignDepartment.value;

  if (!owner || !gender || !department) {
    setStatus(assignStatusMsg, "Owner, gender, and department are required.", "err");
    return;
  }

  upsertEmployeeFromAssignment(owner, gender, department, assignPhotoDataUrl);

  const updated = {
    ...asset,
    status: "Assigned",
    assignment: {
      owner,
      systemName,
      locationDetail,
      gender,
      department,
      assigneePhoto: assignPhotoDataUrl || asset.assignment?.assigneePhoto || "",
      assignedAt: asset.assignment?.assignedAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    },
    returnedAt: "",
    returnReason: "",
    returnNote: "",
    updatedAt: new Date().toISOString()
  };

  assets = assets.map((a) => (a.id === asset.id ? updated : a));
  logActivity("ASSIGN_ASSET", `${updated.assetTag} -> ${owner}`);
  persistAssets();
  renderTrackerTable();
  clearAssignActionPanel();
  setStatus(assignStatusMsg, `Assigned/updated ${updated.assetTag}.`, "ok");
}

function deassignAsset() {
  const asset = readAssignAsset();
  if (!asset) return;

  const updated = {
    ...asset,
    status: "Available",
    assignment: null,
    returnedAt: "",
    returnReason: "",
    returnNote: "",
    updatedAt: new Date().toISOString()
  };

  assets = assets.map((a) => (a.id === asset.id ? updated : a));
  logActivity("DEASSIGN_ASSET", updated.assetTag);
  persistAssets();
  renderTrackerTable();
  clearAssignActionPanel();
  setStatus(assignStatusMsg, `${updated.assetTag} deassigned.`, "ok");
}

function openReturnPageFromTracker() {
  const asset = readAssignAsset();
  setActivePage("return");
  if (!asset) return;
  returnAssetSelect.value = `${asset.assetTag} - ${asset.name}`;
  fillReturnForm(asset);
}

function returnAssetFromReturnPage() {
  const asset = readReturnAsset();
  if (!asset) {
    setStatus(returnStatusMsg, "Select a valid asset.", "err");
    return;
  }

  const dateValue = returnDateInput.value;
  const reason = returnReasonSelect.value;
  if (!dateValue) {
    setStatus(returnStatusMsg, "Return date is required.", "err");
    return;
  }
  if (!reason) {
    setStatus(returnStatusMsg, "Return reason is required.", "err");
    return;
  }

  const returnedAt = new Date(`${dateValue}T00:00:00`).toISOString();
  const note = returnNoteInput.value.trim();
  const previousOwner = asset.assignment?.owner || "";

  returnLogs.push({
    id: uid(),
    assetId: asset.id,
    assetTag: asset.assetTag,
    assetName: asset.name,
    assetType: asset.type,
    previousOwner,
    reason,
    note,
    company: asset.company,
    location: asset.location,
    returnedAt,
    returnedBy: sessionUser || "",
    createdAt: new Date().toISOString()
  });
  saveJson(RETURN_LOG_KEY, returnLogs);

  const updated = {
    ...asset,
    status: "Available",
    assignment: null,
    returnedAt,
    returnReason: reason,
    returnNote: note,
    updatedAt: new Date().toISOString()
  };

  assets = assets.map((a) => (a.id === asset.id ? updated : a));
  logActivity("RETURN_ASSET", `${updated.assetTag} reason=${reason}`);
  persistAssets();
  renderTrackerTable();
  refreshAssignAssetSelect();
  refreshReturnAssetSelect();
  renderReturnHistory();
  clearReturnForm();
  clearAssignActionPanel();
  setStatus(returnStatusMsg, `${updated.assetTag} returned and moved to Available.`, "ok");
  setStatus(assignStatusMsg, `${updated.assetTag} returned and moved to Available.`, "ok");
}

function applyStatusChange() {
  const asset = readAssignAsset();
  if (!asset) {
    setStatus(assignStatusMsg, "Select a valid asset before applying status.", "err");
    return;
  }
  const nextStatus = statusAction.value;

  if (nextStatus === "Returned") {
    setStatus(assignStatusMsg, "Use the Return Asset page to mark return with date and reason.", "err");
    return;
  }

  const updated = {
    ...asset,
    status: nextStatus,
    returnedAt: nextStatus === "Returned" ? asset.returnedAt : "",
    returnReason: nextStatus === "Returned" ? asset.returnReason : "",
    returnNote: nextStatus === "Returned" ? asset.returnNote : "",
    updatedAt: new Date().toISOString()
  };

  assets = assets.map((a) => (a.id === asset.id ? updated : a));
  logActivity("UPDATE_ASSET_STATUS", `${updated.assetTag} -> ${nextStatus}`);
  persistAssets();
  renderTrackerTable();
  clearAssignActionPanel();
  setStatus(assignStatusMsg, `${updated.assetTag} status changed to ${nextStatus}.`, "ok");
}

function matchesSearch(asset, query) {
  if (!query) return true;
  const hay = [
    asset.assetTag, asset.name, asset.type, asset.serial,
    asset.macAddress, asset.imei, asset.os, asset.storage, asset.ram, asset.graphics,
    asset.company, asset.location, asset.status,
    asset.assignment?.owner || "", asset.assignment?.systemName || "", asset.assignment?.department || ""
  ].join(" ").toLowerCase();
  return hay.includes(query);
}

function getVisibleTrackerAssets() {
  const query = searchInput.value.trim().toLowerCase();
  const fType = filterType.value;
  const fStatus = filterStatus.value;

  return getScopedAssets()
    .filter((a) => matchesSearch(a, query))
    .filter((a) => (fType ? a.type === fType : true))
    .filter((a) => (fStatus ? a.status === fStatus : true))
    .sort((a, b) => a.type.localeCompare(b.type) || a.status.localeCompare(b.status) || a.assetTag.localeCompare(b.assetTag));
}

function shouldShowDeviceColumns(rows) {
  if (isDeviceType(filterType.value)) return true;
  return rows.some((r) => isDeviceType(r.type));
}

function renderTrackerTable() {
  const allRows = getVisibleTrackerAssets();
  const { rows } = getPagedRows("asset-tbody", allRows);
  const showDevice = shouldShowDeviceColumns(allRows);

  [thMac, thImei, thOs, thStorage, thRam, thGraphics].forEach((th) => th.classList.toggle("hidden", !showDevice));

  if (!allRows.length) {
    tableBody.innerHTML = `<tr><td colspan="19">No assets found for current company/location.</td></tr>`;
    return;
  }

  tableBody.innerHTML = rows.map((asset) => {
    const owner = asset.assignment?.owner || "-";
    const systemName = asset.assignment?.systemName || "-";
    const locationDetail = asset.assignment?.locationDetail || "-";
    const gender = asset.assignment?.gender || "-";
    const dept = asset.assignment?.department || "-";
    const deviceCols = showDevice
      ? `<td>${escapeHtml(asset.macAddress || "-")}</td><td>${escapeHtml(asset.imei || "-")}</td><td>${escapeHtml(asset.os || "-")}</td><td>${escapeHtml(asset.storage || "-")}</td><td>${escapeHtml(asset.ram || "-")}</td><td>${escapeHtml(asset.graphics || "-")}</td>`
      : "";
    const canShowAdmin = isDeviceType(asset.type) && asset.adminPassword;
    const adminCell = canShowAdmin
      ? `<button type="button" class="eye-btn secondary" data-action="show-admin" data-id="${asset.id}">Eye</button>`
      : "-";

    return `<tr>
      <td>${escapeHtml(asset.assetTag)}</td>
      <td>${escapeHtml(asset.name)}</td>
      <td>${escapeHtml(asset.type)}</td>
      <td>${escapeHtml(asset.serial || "-")}</td>
      ${deviceCols}
      <td>${escapeHtml(owner)}</td>
      <td>${escapeHtml(systemName)}</td>
      <td>${escapeHtml(locationDetail)}</td>
      <td>${escapeHtml(gender)}</td>
      <td>${escapeHtml(dept)}</td>
      <td>${escapeHtml(asset.company)}</td>
      <td>${escapeHtml(asset.location)}</td>
      <td>${escapeHtml(asset.status)}</td>
      <td>${adminCell}</td>
      <td>
        <div class="row-actions">
          <button type="button" class="secondary" data-action="select" data-id="${asset.id}">Barcode</button>
          <button type="button" class="ghost" data-action="edit" data-id="${asset.id}">Edit</button>
          <button type="button" class="danger" data-action="delete" data-id="${asset.id}">Delete</button>
        </div>
      </td>
    </tr>`;
  }).join("");
}

function requirePasswordForAdminReveal(id) {
  const asset = assets.find((a) => a.id === id);
  if (!asset || !isDeviceType(asset.type) || !asset.adminPassword) {
    setStatus(assignStatusMsg, "Admin access not available for this asset type.", "err");
    return;
  }

  const input = window.prompt("Enter your login password to view admin password:");
  if (input === null) return;
  if (!verifyLoginPassword(input)) {
    setStatus(assignStatusMsg, "Password check failed.", "err");
    return;
  }

  window.alert(`Admin Password (${asset.assetTag}): ${asset.adminPassword}`);
  setStatus(assignStatusMsg, "Admin password shown after verification.", "ok");
}

function sanitizeCode39Value(value) {
  return value
    .toUpperCase()
    .split("")
    .map((ch) => (CODE39_MAP[ch] && ch !== "*" ? ch : "-"))
    .join("");
}

function createCode39Svg(rawValue, options = {}) {
  const value = sanitizeCode39Value(rawValue);
  const payload = `*${value}*`;
  const narrow = options.narrow ?? 2;
  const wide = options.wide ?? 5;
  const barHeight = options.barHeight ?? 90;
  const gap = options.gap ?? narrow;
  const margin = options.margin ?? 12;

  let width = margin * 2;
  let x = margin;
  const rects = [];

  for (let c = 0; c < payload.length; c += 1) {
    const pattern = CODE39_MAP[payload[c]];
    if (!pattern) continue;
    for (let i = 0; i < pattern.length; i += 1) {
      const isBar = i % 2 === 0;
      const w = pattern[i] === "w" ? wide : narrow;
      if (isBar) rects.push(`<rect x="${x}" y="${margin}" width="${w}" height="${barHeight}" fill="#000" />`);
      x += w;
      width += w;
    }
    if (c < payload.length - 1) {
      x += gap;
      width += gap;
    }
  }

  const textY = margin + barHeight + 18;
  const height = margin + barHeight + 26;
  return `
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="Barcode ${escapeHtml(value)}">
  <rect x="0" y="0" width="${width}" height="${height}" fill="#fff" />
  ${rects.join("\n  ")}
  <text x="50%" y="${textY}" text-anchor="middle" font-family="monospace" font-size="14" fill="#111">${escapeHtml(value)}</text>
</svg>`.trim();
}

function renderBarcodePreview(asset) {
  if (!asset) {
    barcodePreview.classList.add("empty");
    barcodePreview.textContent = "No asset selected.";
    printBtn.disabled = true;
    return;
  }

  const owner = asset.assignment?.owner || "Unassigned";
  const photo = asset.assignment?.assigneePhoto || "";
  barcodePreview.classList.remove("empty");
  barcodePreview.innerHTML = `
    <div class="barcode-content">
      ${createCode39Svg(asset.assetTag)}
      <div class="asset-meta"><strong>${escapeHtml(asset.assetTag)}</strong> - ${escapeHtml(asset.name)}</div>
      <div class="asset-meta">${escapeHtml(owner)} | ${escapeHtml(asset.company)} | ${escapeHtml(asset.location)} | ${escapeHtml(asset.status)}</div>
      ${photo ? `<img class="assignee-thumb" src="${photo}" alt="${escapeHtml(owner)}" />` : ""}
    </div>
  `;
  printBtn.disabled = false;
}

function printSelectedBarcode() {
  const asset = assets.find((a) => a.id === selectedAssetId);
  if (!asset) return;

  const owner = asset.assignment?.owner || "Unassigned";
  const photo = asset.assignment?.assigneePhoto || "";
  const logo = COMPANY_META[asset.company]?.logo || "";

  const html = `<!doctype html><html><head><meta charset="utf-8" />
<title>Print Label - ${escapeHtml(asset.assetTag)}</title>
<style>
body{font-family:Arial,sans-serif;padding:20px}
.label{border:1px solid #ccc;padding:12px;max-width:980px}
.logo{height:38px;max-width:160px;object-fit:contain;display:block;margin-bottom:10px}
.row{display:flex;align-items:center;justify-content:space-between;gap:14px}
.left,.middle,.right{display:flex;flex-direction:column;justify-content:center}
.left{width:110px;align-items:center}
.middle{flex:1;min-width:260px}
.right{width:430px;align-items:flex-end}
.photo{width:90px;height:90px;border:1px solid #bbb;object-fit:cover;border-radius:8px}
.meta{font-size:14px;line-height:1.45;margin:2px 0}
.barcode-wrap svg{max-width:100%;height:auto}
@media print{body{padding:0}.label{border:none;max-width:none}.right{width:48%}}
</style></head><body><div class="label">
${logo ? `<img class="logo" src="${logo}" alt="${escapeHtml(asset.company)}" />` : ""}
<div class="row">
  <div class="left">
    ${photo ? `<img class="photo" src="${photo}" alt="${escapeHtml(owner)}" />` : `<div class="meta">No Photo</div>`}
  </div>
  <div class="middle">
    <div class="meta"><strong>Asset:</strong> ${escapeHtml(asset.name)}</div>
    <div class="meta"><strong>Tag:</strong> ${escapeHtml(asset.assetTag)}</div>
    <div class="meta"><strong>Assigned Person:</strong> ${escapeHtml(owner)}</div>
    <div class="meta"><strong>Assigned Location:</strong> ${escapeHtml(asset.location)}</div>
    <div class="meta"><strong>Status:</strong> ${escapeHtml(asset.status)}</div>
  </div>
  <div class="right barcode-wrap">
    ${createCode39Svg(asset.assetTag, { narrow: 2, wide: 5, barHeight: 86, margin: 10 })}
  </div>
</div>
</div><script>window.onload=()=>window.print();</script></body></html>`;

  const w = window.open("", "_blank", "width=820,height=620");
  if (!w) {
    setStatus(assignStatusMsg, "Pop-up blocked. Allow pop-ups for printing.", "err");
    return;
  }
  w.document.write(html);
  w.document.close();
}

function findByBarcode() {
  const code = normalizeTag(scanInput.value);
  if (!code) {
    setStatus(assignStatusMsg, "Scan or enter barcode value.", "err");
    return;
  }

  const asset = assets.find((a) => a.assetTag === code);
  if (!asset) {
    setStatus(assignStatusMsg, `No asset found for ${code}.`, "err");
    return;
  }

  selectedAssetId = asset.id;
  renderBarcodePreview(asset);

  if (asset.company === currentCompany && asset.location === currentLocation) {
    searchInput.value = code;
    renderTrackerTable();
    assignAssetSelect.value = `${asset.assetTag} - ${asset.name}`;
    fillAssignForm(asset);
    setStatus(assignStatusMsg, `Matched ${code}.`, "ok");
  } else {
    setStatus(assignStatusMsg, `Matched ${code} in ${asset.company} / ${asset.location}. Switch location/module to edit.`, "ok");
  }
}

function toCsv(rows) {
  const headers = [
    "assetTag", "name", "type", "serial", "macAddress", "imei", "os", "storage", "ram", "graphics",
    "printerMode", "printerIp", "printerPassword", "adminPassword", "company", "location", "status",
    "owner", "systemName", "locationDetail", "gender", "department", "purchaseDate", "warrantyDate", "notes", "createdAt", "updatedAt"
  ];

  const lines = [headers.join(",")];
  for (const row of rows) {
    const mapped = {
      ...row,
      owner: row.assignment?.owner || "",
      systemName: row.assignment?.systemName || "",
      locationDetail: row.assignment?.locationDetail || "",
      gender: row.assignment?.gender || "",
      department: row.assignment?.department || ""
    };

    lines.push(headers.map((h) => {
      const v = String(mapped[h] || "");
      return v.includes(",") || v.includes('"') || v.includes("\n") ? `"${v.replaceAll('"', '""')}"` : v;
    }).join(","));
  }
  return lines.join("\n");
}

function parseCsv(text) {
  const rows = [];
  let i = 0;
  let cell = "";
  let row = [];
  let inQuote = false;

  while (i < text.length) {
    const ch = text[i];
    const next = text[i + 1];
    if (inQuote) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 2;
        continue;
      }
      if (ch === '"') {
        inQuote = false;
        i += 1;
        continue;
      }
      cell += ch;
      i += 1;
      continue;
    }
    if (ch === '"') {
      inQuote = true;
      i += 1;
      continue;
    }
    if (ch === ",") {
      row.push(cell);
      cell = "";
      i += 1;
      continue;
    }
    if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      i += 1;
      continue;
    }
    if (ch !== "\r") cell += ch;
    i += 1;
  }

  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }

  if (!rows.length) return [];
  const headers = rows[0].map((h) => h.trim());
  return rows.slice(1).filter((r) => r.some((v) => String(v).trim())).map((r) => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = r[idx] ?? "";
    });
    return obj;
  });
}

function toCsvRows(headers, rows) {
  const lines = [headers.join(",")];
  rows.forEach((row) => {
    lines.push(headers.map((h) => {
      const v = String(row[h] ?? "");
      return v.includes(",") || v.includes('"') || v.includes("\n") ? `"${v.replaceAll('"', '""')}"` : v;
    }).join(","));
  });
  return lines.join("\n");
}

function refreshActivityFilterOptions() {
  if (!activityFilterUser || !activityFilterAction) return;
  const prevUser = activityFilterUser.value;
  const prevAction = activityFilterAction.value;
  const usersList = Array.from(new Set(activityLogs.map((r) => String(r.user || "").trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b));
  const actionsList = Array.from(new Set(activityLogs.map((r) => String(r.action || "").trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b));
  activityFilterUser.innerHTML = `<option value="">All</option>${usersList.map((u) => `<option value="${escapeHtml(u)}">${escapeHtml(u)}</option>`).join("")}`;
  activityFilterAction.innerHTML = `<option value="">All</option>${actionsList.map((a) => `<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join("")}`;
  if (usersList.includes(prevUser)) activityFilterUser.value = prevUser;
  if (actionsList.includes(prevAction)) activityFilterAction.value = prevAction;
}

function getVisibleActivityLogs() {
  const q = String(activitySearchInput?.value || "").trim().toLowerCase();
  const fUser = String(activityFilterUser?.value || "");
  const fAction = String(activityFilterAction?.value || "");
  const from = activityDateFrom?.value ? new Date(`${activityDateFrom.value}T00:00:00`).getTime() : null;
  const to = activityDateTo?.value ? new Date(`${activityDateTo.value}T23:59:59`).getTime() : null;
  const sortMode = String(activitySort?.value || "newest");

  const rows = activityLogs
    .filter((r) => (fUser ? String(r.user || "") === fUser : true))
    .filter((r) => (fAction ? String(r.action || "") === fAction : true))
    .filter((r) => {
      const t = new Date(r.at || 0).getTime();
      if (from && t < from) return false;
      if (to && t > to) return false;
      return true;
    })
    .filter((r) => {
      if (!q) return true;
      const hay = `${r.user || ""} ${r.action || ""} ${r.details || ""} ${r.company || ""} ${r.location || ""}`.toLowerCase();
      return hay.includes(q);
    });

  rows.sort((a, b) => {
    if (sortMode === "oldest") return new Date(a.at || 0).getTime() - new Date(b.at || 0).getTime();
    if (sortMode === "user-az") return String(a.user || "").localeCompare(String(b.user || "")) || new Date(b.at || 0).getTime() - new Date(a.at || 0).getTime();
    if (sortMode === "action-az") return String(a.action || "").localeCompare(String(b.action || "")) || new Date(b.at || 0).getTime() - new Date(a.at || 0).getTime();
    return new Date(b.at || 0).getTime() - new Date(a.at || 0).getTime();
  });
  return rows;
}

function renderActivityHistoryPage() {
  if (!activityTbody) return;
  refreshActivityFilterOptions();
  const allRows = getVisibleActivityLogs();
  const { rows } = getPagedRows("activity-tbody", allRows);
  if (!allRows.length) {
    activityTbody.innerHTML = `<tr><td colspan="6">No activity found for current filters.</td></tr>`;
    setStatus(activityStatusMsg, "No activity found.", "err");
    return;
  }
  activityTbody.innerHTML = rows.map((r) => `<tr>
    <td>${escapeHtml(String(r.at || "").replace("T", " ").slice(0, 19))}</td>
    <td>${escapeHtml(r.user || "-")}</td>
    <td>${escapeHtml(r.action || "-")}</td>
    <td>${escapeHtml(r.details || "-")}</td>
    <td>${escapeHtml(r.company || "-")}</td>
    <td>${escapeHtml(r.location || "-")}</td>
  </tr>`).join("");
  setStatus(activityStatusMsg, `Showing ${rows.length} of ${allRows.length} activities.`, "ok");
}

function exportActivityCsv() {
  if (!isAdminUser()) return;
  const rows = getVisibleActivityLogs().map((r) => ({
    dateTime: String(r.at || "").replace("T", " ").slice(0, 19),
    user: r.user || "",
    action: r.action || "",
    details: r.details || "",
    company: r.company || "",
    location: r.location || ""
  }));
  const csv = toCsvRows(["dateTime", "user", "action", "details", "company", "location"], rows);
  downloadFile(`activity-history-${todayKey()}.csv`, csv, "text/csv");
  setStatus(activityStatusMsg, `Exported ${rows.length} activity rows.`, "ok");
}

function exportAllData() {
  if (!isAdminUser()) return;
  triggerManualBackup("export-all-data");
  setStatus(activityStatusMsg, "All data export downloaded.", "ok");
}

function clearAllActivityHistoryByAdmin() {
  if (!isAdminUser()) {
    setStatus(activityStatusMsg, "Only admins can clear activity history.", "err");
    return;
  }
  const input = window.prompt("Enter admin password to clear all activity history:");
  if (input === null) return;
  if (!verifyAnyAdminPassword(input)) {
    setStatus(activityStatusMsg, "Invalid admin password.", "err");
    return;
  }
  if (!window.confirm("Clear all activity history records? This cannot be undone.")) return;

  activityLogs = [];
  saveJson(ACTIVITY_LOG_KEY, activityLogs);
  renderActivityHistoryPage();
  setStatus(activityStatusMsg, "All activity history cleared.", "ok");
}

function downloadFile(name, text, mime = "text/plain") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

function downloadSampleCsv() {
  const sample = [
    "name,type,serial,purchaseDate,warrantyDate,status,assignedOwner,assignedEmployeeId,macAddress,imei,os,storage,ram,graphics,printerMode,printerIp,printerPassword,adminPassword,notes",
    "Dell Latitude 7440,Laptop,SN-1234,2026-01-15,2029-01-14,Assigned,Ramees,EMP-001,00:11:22:33:44:55,,Windows 11,512GB SSD,16GB,Intel Iris Xe,,,,admin123,Office laptop",
    "HP LaserJet 400,Printer,PR-9382,2025-11-10,2027-11-09,Available,,,,,,,,,Network Enabled,192.168.1.50,printer@123,,Network printer"
  ].join("\n");
  downloadFile("asset-bulk-sample.csv", sample, "text/csv");
  setStatus(catalogStatusMsg, "Sample CSV downloaded.", "ok");
}

function refreshCatalogBulkHistoryOptions() {
  if (!catalogBulkHistorySelect) return;
  const scopedIds = new Set(getScopedAssets().map((a) => a.id));
  const items = catalogBulkHistory
    .filter((b) => b.company === currentCompany && b.location === currentLocation)
    .filter((b) => Array.isArray(b.assetIds) && b.assetIds.some((id) => scopedIds.has(id)))
    .sort((a, b) => String(b.at || "").localeCompare(String(a.at || "")));
  catalogBulkHistorySelect.innerHTML = `<option value="">Select upload batch</option>${items
    .map((b) => `<option value="${b.id}">${escapeHtml((b.at || "").replace("T", " ").slice(0, 19))} - ${b.count || b.assetIds.length} assets</option>`)
    .join("")}`;
}

function undoLastCatalogBulkUpload() {
  const scoped = catalogBulkHistory
    .filter((b) => b.company === currentCompany && b.location === currentLocation)
    .sort((a, b) => String(b.at || "").localeCompare(String(a.at || "")));
  const last = scoped[0];
  if (!last) {
    setStatus(catalogStatusMsg, "No catalog bulk upload history found.", "err");
    return;
  }
  const idSet = new Set(last.assetIds || []);
  const before = assets.length;
  assets = assets.filter((a) => !idSet.has(a.id));
  catalogBulkHistory = catalogBulkHistory.filter((b) => b.id !== last.id);
  saveJson(STORAGE_KEY, assets);
  saveJson(CATALOG_BULK_LOG_KEY, catalogBulkHistory);
  purgeSoftwareForDeletedAssets(new Set(assets.map((a) => a.id)));
  const removed = before - assets.length;
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
  refreshCatalogBulkHistoryOptions();
  setStatus(catalogStatusMsg, `Undo complete. Removed ${removed} assets from last catalog bulk upload.`, "ok");
}

function undoAllCatalogBulkUploads() {
  const scopedBatches = catalogBulkHistory.filter((b) => b.company === currentCompany && b.location === currentLocation);
  if (!scopedBatches.length) {
    setStatus(catalogStatusMsg, "No catalog bulk upload history found.", "err");
    return;
  }
  if (!window.confirm(`Undo all catalog bulk uploads for ${currentCompany} / ${currentLocation}?`)) return;
  const idSet = new Set(scopedBatches.flatMap((b) => b.assetIds || []));
  const before = assets.length;
  assets = assets.filter((a) => !idSet.has(a.id));
  catalogBulkHistory = catalogBulkHistory.filter((b) => !(b.company === currentCompany && b.location === currentLocation));
  saveJson(STORAGE_KEY, assets);
  saveJson(CATALOG_BULK_LOG_KEY, catalogBulkHistory);
  purgeSoftwareForDeletedAssets(new Set(assets.map((a) => a.id)));
  const removed = before - assets.length;
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
  refreshCatalogBulkHistoryOptions();
  setStatus(catalogStatusMsg, `Undo complete. Removed ${removed} assets from all catalog bulk uploads.`, "ok");
}

function deleteSelectedCatalogBulkUpload() {
  const batchId = String(catalogBulkHistorySelect?.value || "");
  if (!batchId) {
    setStatus(catalogStatusMsg, "Select a catalog bulk upload batch.", "err");
    return;
  }
  const batch = catalogBulkHistory.find((b) => b.id === batchId);
  if (!batch) {
    setStatus(catalogStatusMsg, "Batch not found.", "err");
    return;
  }
  if (!window.confirm(`Delete assets from selected bulk upload (${batch.count || 0})?`)) return;
  const idSet = new Set(batch.assetIds || []);
  const before = assets.length;
  assets = assets.filter((a) => !idSet.has(a.id));
  catalogBulkHistory = catalogBulkHistory.filter((b) => b.id !== batchId);
  saveJson(STORAGE_KEY, assets);
  saveJson(CATALOG_BULK_LOG_KEY, catalogBulkHistory);
  purgeSoftwareForDeletedAssets(new Set(assets.map((a) => a.id)));
  const removed = before - assets.length;
  refreshAssignAssetSelect();
  renderTrackerTable();
  renderCatalogList();
  refreshCatalogBulkHistoryOptions();
  setStatus(catalogStatusMsg, `Removed ${removed} assets from selected catalog bulk upload.`, "ok");
}

function exportCatalogCsv() {
  const rows = getScopedAssets().sort((a, b) => a.assetTag.localeCompare(b.assetTag));
  downloadFile(`catalog-${currentCompany}-${currentLocation}-${todayKey()}.csv`, toCsv(rows), "text/csv");
  setStatus(catalogStatusMsg, `Exported ${rows.length} rows.`, "ok");
}

async function importCatalogCsv(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const rows = parseCsv(await file.text());
    let added = 0;
    let skipped = 0;
    let assignedMatched = 0;
    let assignedFallbackAvailable = 0;
    const batchIds = [];

    for (const row of rows) {
      const type = String(row.type || "Other").trim();
      const device = isDeviceType(type);
      const printer = isPrinterType(type);
      const ownerInput = String(row.assignedOwner || row.owner || "").trim();
      const employeeIdInput = String(row.assignedEmployeeId || row.employeeId || "").trim();
      const wantedAssigned = String(row.status || "").trim().toLowerCase() === "assigned" || !!ownerInput || !!employeeIdInput;
      const scopedEmployees = getScopedEmployees();
      const matchedById = employeeIdInput
        ? scopedEmployees.find((e) => String(e.employeeId || "").trim().toLowerCase() === employeeIdInput.toLowerCase())
        : null;
      const matchedByName = ownerInput
        ? scopedEmployees.find((e) => String(e.name || "").trim().toLowerCase() === ownerInput.toLowerCase())
        : null;
      const matchedEmployee = matchedById || matchedByName || null;
      let resolvedStatus = String(row.status || "Available").trim() || "Available";
      let assignment = null;
      if (wantedAssigned) {
        if (matchedEmployee) {
          resolvedStatus = "Assigned";
          assignment = {
            owner: matchedEmployee.name || ownerInput,
            systemName: "",
            locationDetail: currentLocation === "Site" ? String(matchedEmployee.siteName || "").trim() : "",
            gender: String(matchedEmployee.gender || ""),
            department: String(matchedEmployee.department || ""),
            assigneePhoto: String(matchedEmployee.photo || ""),
            assignedAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          };
          assignedMatched += 1;
        } else {
          resolvedStatus = "Available";
          assignment = null;
          assignedFallbackAvailable += 1;
        }
      }
      const asset = {
        id: uid(),
        assetTag: getNextTag(type),
        name: String(row.name || "").trim(),
        type,
        serial: String(row.serial || "").trim(),
        macAddress: device ? String(row.macAddress || "").trim() : "",
        imei: device ? String(row.imei || "").trim() : "",
        os: device ? String(row.os || "").trim() : "",
        storage: device ? String(row.storage || "").trim() : "",
        ram: device ? String(row.ram || "").trim() : "",
        graphics: device ? String(row.graphics || "").trim() : "",
        printerMode: printer ? String(row.printerMode || "Direct") : "",
        printerIp: printer ? String(row.printerIp || "").trim() : "",
        printerPassword: printer ? String(row.printerPassword || "") : "",
        adminPassword: device ? String(row.adminPassword || "") : "",
        company: currentCompany,
        location: currentLocation,
        purchaseDate: String(row.purchaseDate || ""),
        warrantyDate: String(row.warrantyDate || ""),
        status: resolvedStatus,
        notes: String(row.notes || ""),
        assignment,
        returnedAt: "",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };

      const error = validateCatalogAsset(asset);
      if (error) {
        skipped += 1;
        continue;
      }
      assets.push(asset);
      batchIds.push(asset.id);
      added += 1;
    }

    if (batchIds.length) {
      catalogBulkHistory.push({
        id: uid(),
        at: new Date().toISOString(),
        company: currentCompany,
        location: currentLocation,
        count: batchIds.length,
        assetIds: batchIds
      });
      saveJson(CATALOG_BULK_LOG_KEY, catalogBulkHistory);
    }

    persistAssets();
    resetCatalogForm();
    refreshAssignAssetSelect();
    renderTrackerTable();
    renderCatalogList();
    refreshCatalogBulkHistoryOptions();
    let msg = `CSV import done. Added ${added}, skipped ${skipped}.`;
    if (assignedMatched || assignedFallbackAvailable) {
      msg += ` Assigned matched: ${assignedMatched}. Fallback to Available: ${assignedFallbackAvailable}.`;
    }
    setStatus(catalogStatusMsg, msg, skipped ? "err" : "ok");
  } catch (error) {
    setStatus(catalogStatusMsg, `CSV import failed: ${error.message}`, "err");
  } finally {
    catalogImportCsvInput.value = "";
  }
}

function getExportRowsByMode() {
  const mode = exportMode.value;
  if (mode === "all") return assets;
  if (mode === "dateRange") {
    const from = exportFrom.value ? new Date(`${exportFrom.value}T00:00:00`) : null;
    const to = exportTo.value ? new Date(`${exportTo.value}T23:59:59`) : null;
    return assets.filter((a) => {
      const d = new Date(a.createdAt || a.updatedAt || Date.now());
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });
  }
  return getVisibleTrackerAssets();
}

function exportTrackerCsv() {
  const rows = getExportRowsByMode();
  downloadFile(`tracker-${todayKey()}.csv`, toCsv(rows), "text/csv");
  setStatus(assignStatusMsg, `Exported ${rows.length} tracker rows.`, "ok");
}

function createDailyBackupSnapshot() {
  const today = todayKey();
  const last = localStorage.getItem(BACKUP_DATE_KEY);
  if (last === today) return;

  const backups = loadJson(BACKUP_KEY, []);
  backups.push({ date: today, generatedAt: new Date().toISOString(), count: assets.length, assets });
  while (backups.length > 30) backups.shift();
  saveJson(BACKUP_KEY, backups);
  localStorage.setItem(BACKUP_DATE_KEY, today);
}

function getTodayBackup() {
  createDailyBackupSnapshot();
  const backups = loadJson(BACKUP_KEY, []);
  return backups.find((b) => b.date === todayKey()) || backups[backups.length - 1] || null;
}

function downloadTodayBackup(reason = "manual") {
  const backup = getTodayBackup();
  if (!backup) {
    setStatus(assignStatusMsg, "No backup data available.", "err");
    return;
  }

  downloadFile(`${todayKey()}-it-assets-backup.json`, JSON.stringify(backup, null, 2), "application/json");
  setStatus(assignStatusMsg, "Backup downloaded. Save it in backups folder.", "ok");
}

function triggerManualBackup(source = "manual") {
  persistLiveBackupSnapshot(`manual:${source}`);
  const payload = buildFullBackupPayload(source);
  const stamp = new Date().toISOString().replaceAll(":", "-");
  downloadFile(`it-full-backup-${stamp}.json`, JSON.stringify(payload, null, 2), "application/json");
  const summary = `Backup: assets ${payload.assets.length}, users ${payload.users.length}, employees ${payload.employees.length}, software ${payload.softwareInventory.length}, licenses ${payload.softwarePurchases.length}.`;
  if (!screens.auth.classList.contains("hidden")) setStatus(authStatus, summary, "ok");
  else if (!screens.company.classList.contains("hidden")) setStatus(companyStatus, summary, "ok");
  else if (!screens.module.classList.contains("hidden")) setStatus(moduleStatus, summary, "ok");
  else if (!screens.settings.classList.contains("hidden")) setStatus(adminSettingsStatus, summary, "ok");
  else setStatus(assignStatusMsg, summary, "ok");
  logActivity("EXPORT_ALL_DATA", `source=${source}`);
}

function shuffle(array) {
  const a = [...array];
  for (let i = a.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function fmtDate(value) {
  if (!value) return "-";
  return String(value).slice(0, 10);
}

function getLastAuditForAsset(assetId) {
  const rows = audits
    .filter((a) => a.assetId === assetId)
    .sort((x, y) => new Date(y.changedAt || y.auditDate || y.doneAt || y.date || 0).getTime() - new Date(x.changedAt || x.auditDate || x.doneAt || x.date || 0).getTime());
  return rows[0] || null;
}

function getLastCompletedAuditForAsset(assetId) {
  const rows = audits
    .filter((a) => a.assetId === assetId && a.auditDone === true)
    .sort((x, y) => new Date(y.auditDate || y.doneAt || y.date || 0).getTime() - new Date(x.auditDate || x.doneAt || x.date || 0).getTime());
  return rows[0] || null;
}

function getNextAuditDueDate(lastAudit) {
  if (!lastAudit) return null;
  if (lastAudit.nextDueAt) return lastAudit.nextDueAt;
  const base = new Date(lastAudit.auditDate || lastAudit.doneAt || lastAudit.date || Date.now());
  base.setDate(base.getDate() + 30);
  return base.toISOString();
}

function isAuditDueForAsset(asset) {
  const last = getLastAuditForAsset(asset.id);
  if (last?.statusOverride === "due") return true;
  if (!last) return true;
  const dueAt = getNextAuditDueDate(last);
  return !dueAt || new Date(dueAt).getTime() <= Date.now();
}

function renderAuditTable(rows) {
  const allRows = Array.isArray(rows) ? rows : [];
  const paged = getPagedRows("audit-tbody", allRows);
  if (!allRows.length) {
    auditTbody.innerHTML = `<tr><td colspan="8">No audit sample generated.</td></tr>`;
    return;
  }
  auditTbody.innerHTML = paged.rows.map((a) => {
    const last = getLastAuditForAsset(a.id);
    const dueAt = getNextAuditDueDate(last);
    return `
      <tr>
        <td><input type="radio" name="audit-pick" data-id="${a.id}" ${selectedAuditAssetId === a.id ? "checked" : ""} /></td>
        <td>${escapeHtml(a.assetTag)}</td>
        <td>${escapeHtml(a.assignment?.systemName || a.name)}</td>
        <td>${escapeHtml(a.assignment?.department || "-")}</td>
        <td>${escapeHtml(a.assignment?.owner || "-")}</td>
        <td>${escapeHtml(a.os || "-")}</td>
        <td>${fmtDate(last?.auditDate || last?.doneAt || last?.date)}</td>
        <td>${fmtDate(dueAt)}</td>
      </tr>
    `;
  }).join("");
}

function runRandomAudit() {
  const location = auditLocation.value;
  const count = Math.max(1, Number(auditCount.value || 1));
  const pool = assets
    .filter((a) => a.company === currentCompany && a.location === location && isDeviceType(a.type))
    .filter((a) => isAuditDueForAsset(a));

  if (!pool.length) {
    auditSelection = [];
    selectedAuditAssetId = null;
    renderAuditTable([]);
    setStatus(auditStatus, "No due systems available in selected location.", "err");
    return;
  }

  const sample = shuffle(pool).slice(0, Math.min(count, pool.length));
  auditSelection = sample.map((a) => a.id);
  selectedAuditAssetId = auditSelection[0] || null;
  renderAuditTable(sample);
  setStatus(auditStatus, `Generated ${sample.length} random due systems for audit.`, "ok");
}

function renderAuditReportAssetOptions() {
  const list = auditSelection
    .map((id) => assets.find((a) => a.id === id))
    .filter(Boolean);

  if (!list.length) {
    auditReportAsset.innerHTML = `<option value="">No selected system</option>`;
    return;
  }

  auditReportAsset.innerHTML = list.map((a) => `<option value="${a.id}">${escapeHtml(a.assetTag)} - ${escapeHtml(a.assignment?.systemName || a.name)}</option>`).join("");
  if (selectedAuditAssetId && list.some((a) => a.id === selectedAuditAssetId)) {
    auditReportAsset.value = selectedAuditAssetId;
  } else {
    selectedAuditAssetId = list[0].id;
    auditReportAsset.value = selectedAuditAssetId;
  }
}

function fillAuditReportFormByAsset(assetId) {
  const asset = assets.find((a) => a.id === assetId);
  if (!asset) return;
  editingAuditReportId = null;
  saveAuditBtn.textContent = "Mark Audit Done";
  selectedAuditAssetId = asset.id;
  auditReportDate.value = new Date().toISOString().slice(0, 10);
  auditReportPcName.value = asset.assignment?.systemName || asset.name || "";
  auditReportDepartment.value = asset.assignment?.department || "";
  auditReportUserName.value = asset.assignment?.owner || "";
  auditReportOsVersion.value = asset.os || "";
  auditReportAuditorName.value = sessionUser || "";
  auditReportWinActivated.value = "";
  auditReportAdminVerified.value = "";
  auditReportUnauthorizedSw.value = "";
  auditReportAntivirus.value = "";
  auditReportFirewall.value = "";
  auditReportUpdatesPending.value = "";
  auditReportUsbRestricted.value = "";
  auditReportPrinterConfig.value = "";
  auditReportBackupVerified.value = "";
  auditReportIssues.value = "";
  auditReportRiskLevel.value = "Low";
  auditReportActionTaken.value = "";
}

function openAuditReportPage() {
  if (!auditSelection.length) {
    setStatus(auditStatus, "Generate random systems first.", "err");
    return;
  }
  if (!selectedAuditAssetId) selectedAuditAssetId = auditSelection[0];
  setActivePage("audit-report");
}

function renderAuditReportHistory() {
  const scopedSystems = assets
    .filter((a) => a.company === currentCompany && isDeviceType(a.type))
    .sort((a, b) => a.assetTag.localeCompare(b.assetTag));
  if (auditReportHistorySystem) {
    const prev = selectedAuditHistorySystemId || auditReportHistorySystem.value || "";
    auditReportHistorySystem.innerHTML = `<option value="">All Systems</option>${scopedSystems
      .map((a) => `<option value="${a.id}">${escapeHtml(a.assetTag)} - ${escapeHtml(a.assignment?.systemName || a.name)}</option>`)
      .join("")}`;
    if (prev && scopedSystems.some((a) => a.id === prev)) {
      auditReportHistorySystem.value = prev;
      selectedAuditHistorySystemId = prev;
    } else {
      auditReportHistorySystem.value = "";
      selectedAuditHistorySystemId = "";
    }
  }

  const rowsAll = audits
    .filter((a) => a.auditDone === true)
    .filter((a) => a.company === currentCompany)
    .filter((a) => !selectedAuditHistorySystemId || a.assetId === selectedAuditHistorySystemId)
    .filter((a) => {
      const q = String(auditReportSearch?.value || "").trim().toLowerCase();
      if (!q) return true;
      const hay = [
        a.assetTag,
        a.pcName,
        a.department,
        a.userName,
        a.riskLevel,
        a.issuesFound,
        a.auditorName,
        a.location
      ].map((v) => String(v || "").toLowerCase()).join(" ");
      return hay.includes(q);
    })
    .sort((x, y) => new Date(y.auditDate || y.doneAt || y.date || 0).getTime() - new Date(x.auditDate || x.doneAt || x.date || 0).getTime());
  if (auditReportPageSize) {
    auditReportPageSize.value = getListPageSizeValue("audit-report-history-tbody", auditReportPageSize.value || "20");
  }
  const paged = getPagedRows("audit-report-history-tbody", rowsAll, "20");
  const rows = paged.rows;

  if (!rowsAll.length) {
    auditReportHistoryTbody.innerHTML = `<tr><td colspan="10">No saved audit reports.</td></tr>`;
    return;
  }

  auditReportHistoryTbody.innerHTML = rows.map((r) => `
    <tr>
      <td>${fmtDate(r.auditDate || r.doneAt || r.date)}</td>
      <td>${escapeHtml(r.assetTag || "-")}</td>
      <td>${escapeHtml(r.pcName || "-")}</td>
      <td>${escapeHtml(r.department || "-")}</td>
      <td>${escapeHtml(r.userName || "-")}</td>
      <td>${escapeHtml(r.riskLevel || "-")}</td>
      <td>${escapeHtml(r.issuesFound || "-")}</td>
      <td>${escapeHtml(r.auditorName || "-")}</td>
      <td>${escapeHtml(r.location || "-")}</td>
      <td>${isAdminUser()
    ? `<button type="button" class="secondary" data-action="edit-audit-history" data-id="${r.id}">Edit</button> <button type="button" class="danger" data-action="delete-audit-history" data-id="${r.id}">Delete</button>`
    : "-"}</td>
    </tr>
  `).join("");
  if (rowsAll.length > rows.length) {
    setStatus(auditReportStatus, `Showing ${rows.length} of ${rowsAll.length} reports. Use Show to view more.`, "ok");
  }
}

function editAuditHistoryByAdmin(id) {
  if (!isAdminUser()) {
    setStatus(auditReportStatus, "Only admins can edit audit reports.", "err");
    return;
  }
  const report = audits.find((a) => a.id === id && a.auditDone === true);
  if (!report) return;
  openAuditEditModal(report);
}

function openAuditEditModal(report) {
  if (!auditEditModal || !report) return;
  editingAuditModalId = report.id;
  if (auditEditAssetTag) auditEditAssetTag.value = report.assetTag || "-";
  if (auditEditDate) auditEditDate.value = fmtDate(report.auditDate || report.doneAt || report.date) || new Date().toISOString().slice(0, 10);
  if (auditEditPcName) auditEditPcName.value = report.pcName || "";
  if (auditEditDepartment) auditEditDepartment.value = report.department || "";
  if (auditEditUserName) auditEditUserName.value = report.userName || "";
  if (auditEditOsVersion) auditEditOsVersion.value = report.osVersion || "";
  if (auditEditWinActivated) auditEditWinActivated.value = report.windowsActivated || "";
  if (auditEditAdminVerified) auditEditAdminVerified.value = report.adminAccountsVerified || "";
  if (auditEditUnauthorizedSw) auditEditUnauthorizedSw.value = report.unauthorizedSoftware || "";
  if (auditEditAntivirus) auditEditAntivirus.value = report.antivirusActive || "";
  if (auditEditFirewall) auditEditFirewall.value = report.firewallOn || "";
  if (auditEditUpdatesPending) auditEditUpdatesPending.value = report.updatesPending || "";
  if (auditEditUsbRestricted) auditEditUsbRestricted.value = report.usbRestricted || "";
  if (auditEditPrinterConfig) auditEditPrinterConfig.value = report.printerConfigOk || "";
  if (auditEditBackupVerified) auditEditBackupVerified.value = report.backupVerified || "";
  if (auditEditRiskLevel) auditEditRiskLevel.value = report.riskLevel || "Low";
  if (auditEditAuditorName) auditEditAuditorName.value = report.auditorName || sessionUser || "";
  if (auditEditIssues) auditEditIssues.value = report.issuesFound || "";
  if (auditEditActionTaken) auditEditActionTaken.value = report.actionTaken || "";
  setStatus(auditEditStatusMsg, "", "ok");
  auditEditModal.classList.remove("hidden");
  auditEditModal.setAttribute("aria-hidden", "false");
}

function closeAuditEditModal() {
  if (!auditEditModal) return;
  auditEditModal.classList.add("hidden");
  auditEditModal.setAttribute("aria-hidden", "true");
  editingAuditModalId = null;
  if (auditEditForm) auditEditForm.reset();
  setStatus(auditEditStatusMsg, "", "ok");
}

function saveAuditEditModal(event) {
  event?.preventDefault();
  if (!isAdminUser()) {
    setStatus(auditEditStatusMsg, "Only admins can edit audit reports.", "err");
    return;
  }
  if (!editingAuditModalId) return;
  const existing = audits.find((a) => a.id === editingAuditModalId && a.auditDone === true);
  if (!existing) {
    setStatus(auditEditStatusMsg, "Audit report not found.", "err");
    return;
  }
  const asset = assets.find((a) => a.id === existing.assetId);
  if (!asset) {
    setStatus(auditEditStatusMsg, "Linked asset not found.", "err");
    return;
  }
  const nowIso = new Date().toISOString();
  const auditDateValue = auditEditDate?.value || new Date().toISOString().slice(0, 10);
  const dateIso = new Date(`${auditDateValue}T00:00:00`).toISOString();
  const dueAt = new Date(`${auditDateValue}T00:00:00`);
  dueAt.setDate(dueAt.getDate() + 30);

  const payload = {
    ...existing,
    changedAt: nowIso,
    auditDate: dateIso,
    doneAt: nowIso,
    nextDueAt: dueAt.toISOString(),
    pcName: auditEditPcName?.value.trim() || "",
    department: auditEditDepartment?.value.trim() || "",
    userName: auditEditUserName?.value.trim() || "",
    osVersion: auditEditOsVersion?.value.trim() || "",
    windowsActivated: auditEditWinActivated?.value || "",
    adminAccountsVerified: auditEditAdminVerified?.value || "",
    unauthorizedSoftware: auditEditUnauthorizedSw?.value || "",
    antivirusActive: auditEditAntivirus?.value || "",
    firewallOn: auditEditFirewall?.value || "",
    updatesPending: auditEditUpdatesPending?.value || "",
    usbRestricted: auditEditUsbRestricted?.value || "",
    printerConfigOk: auditEditPrinterConfig?.value || "",
    backupVerified: auditEditBackupVerified?.value || "",
    issuesFound: auditEditIssues?.value.trim() || "",
    riskLevel: auditEditRiskLevel?.value || "Low",
    actionTaken: auditEditActionTaken?.value.trim() || "",
    auditorName: auditEditAuditorName?.value.trim() || sessionUser || ""
  };

  audits = audits.map((a) => (a.id === editingAuditModalId ? payload : a));
  saveJson(AUDITS_KEY, audits);
  logActivity("UPDATE_AUDIT_REPORT", payload.assetTag || "");
  renderAuditReportHistory();
  renderAuditStatusList();
  closeAuditEditModal();
  setStatus(auditReportStatus, `Audit updated for ${payload.assetTag}.`, "ok");
}

function deleteAuditHistoryByAdmin(id) {
  if (!isAdminUser()) {
    setStatus(auditReportStatus, "Only admins can delete audit history.", "err");
    return;
  }
  const record = audits.find((a) => a.id === id && a.auditDone === true);
  if (!record) return;
  if (!window.confirm(`Delete audit report for ${record.assetTag} dated ${fmtDate(record.auditDate || record.doneAt || record.date)}?`)) return;
  audits = audits.filter((a) => a.id !== id);
  saveJson(AUDITS_KEY, audits);
  logActivity("DELETE_AUDIT_REPORT", record.assetTag || "");
  renderAuditReportHistory();
  renderAuditStatusList();
  setStatus(auditReportStatus, "Audit report deleted.", "ok");
}

function saveAuditRecord(event) {
  event?.preventDefault();
  if (editingAuditReportId && !isAdminUser()) {
    setStatus(auditReportStatus, "Only admins can edit audit reports.", "err");
    return;
  }
  const wasEdit = !!editingAuditReportId;
  const assetId = auditReportAsset.value || selectedAuditAssetId;
  const asset = assets.find((a) => a.id === assetId);
  if (!asset) {
    setStatus(auditReportStatus, "Select a valid system for audit report.", "err");
    return;
  }

  const nowIso = new Date().toISOString();
  const auditDate = auditReportDate.value || nowIso.slice(0, 10);
  const dateIso = new Date(`${auditDate}T00:00:00`).toISOString();
  const dueAt = new Date(`${auditDate}T00:00:00`);
  dueAt.setDate(dueAt.getDate() + 30);

  const payload = {
    id: uid(),
    auditDone: true,
    done: true,
    statusOverride: "audited",
    assetId: asset.id,
    assetTag: asset.assetTag,
    company: currentCompany,
    location: asset.location,
    changedAt: nowIso,
    auditDate: dateIso,
    doneAt: nowIso,
    nextDueAt: dueAt.toISOString(),
    pcName: auditReportPcName.value.trim(),
    department: auditReportDepartment.value.trim(),
    userName: auditReportUserName.value.trim(),
    osVersion: auditReportOsVersion.value.trim(),
    windowsActivated: auditReportWinActivated.value,
    adminAccountsVerified: auditReportAdminVerified.value,
    unauthorizedSoftware: auditReportUnauthorizedSw.value,
    antivirusActive: auditReportAntivirus.value,
    firewallOn: auditReportFirewall.value,
    updatesPending: auditReportUpdatesPending.value,
    usbRestricted: auditReportUsbRestricted.value,
    printerConfigOk: auditReportPrinterConfig.value,
    backupVerified: auditReportBackupVerified.value,
    issuesFound: auditReportIssues.value.trim(),
    riskLevel: auditReportRiskLevel.value,
    actionTaken: auditReportActionTaken.value.trim(),
    auditorName: auditReportAuditorName.value.trim() || sessionUser || ""
  };
  if (editingAuditReportId) {
    const existing = audits.find((a) => a.id === editingAuditReportId);
    if (existing) {
      payload.id = existing.id;
      payload.createdAt = existing.createdAt || existing.changedAt || nowIso;
    } else {
      payload.id = editingAuditReportId;
    }
    audits = audits.map((a) => (a.id === editingAuditReportId ? payload : a));
  } else {
    audits.push(payload);
  }
  saveJson(AUDITS_KEY, audits);
  logActivity(wasEdit ? "UPDATE_AUDIT_REPORT" : "ADD_AUDIT_REPORT", asset.assetTag);

  if (!wasEdit) {
    auditSelection = auditSelection.filter((id) => id !== asset.id);
    selectedAuditAssetId = auditSelection[0] || null;
    renderAuditReportAssetOptions();
    if (selectedAuditAssetId) fillAuditReportFormByAsset(selectedAuditAssetId);
  }
  renderAuditTable(
    auditSelection
      .map((id) => assets.find((a) => a.id === id))
      .filter(Boolean)
  );
  renderAuditReportHistory();
  setStatus(auditReportStatus, wasEdit ? `Audit updated for ${asset.assetTag}.` : `Audit saved for ${asset.assetTag}. Next due after 30 days.`, "ok");
  renderAuditStatusList();
  editingAuditReportId = null;
  saveAuditBtn.textContent = "Mark Audit Done";
}

function renderAuditStatusList() {
  const location = auditStatusLocation.value || currentLocation;
  const filter = auditStatusFilter?.value || "";
  const scoped = assets
    .filter((a) => a.company === currentCompany && a.location === location && isDeviceType(a.type))
    .sort((a, b) => a.assetTag.localeCompare(b.assetTag));

  if (!scoped.length) {
    auditStatusTbody.innerHTML = `<tr><td colspan="10">No device systems in selected location.</td></tr>`;
    setStatus(auditStatusMsg, "No systems available.", "err");
    return;
  }

  const filteredRows = scoped.filter((a) => {
    const completed = getLastCompletedAuditForAsset(a.id);
    if (filter === "audited") return !!completed;
    if (filter === "notAudited") return !completed;
    return true;
  });
  const dueCount = filteredRows.reduce((sum, a) => {
    const last = getLastAuditForAsset(a.id);
    const dueAt = getNextAuditDueDate(last);
    const due = last?.statusOverride === "due" || !last || (dueAt && new Date(dueAt).getTime() <= Date.now());
    return sum + (due ? 1 : 0);
  }, 0);
  const { rows } = getPagedRows("audit-status-tbody", filteredRows);

  auditStatusTbody.innerHTML = rows.map((a) => {
    const last = getLastAuditForAsset(a.id);
    const completed = getLastCompletedAuditForAsset(a.id);
    const dueAt = getNextAuditDueDate(last);
    const due = last?.statusOverride === "due" || !last || (dueAt && new Date(dueAt).getTime() <= Date.now());
    const viewReportBtn = completed
      ? `<button type="button" class="secondary" data-action="view-audit-report" data-id="${a.id}">View Report</button>`
      : "-";
    const actionCell = isAdminUser()
      ? (completed && !due
        ? `<div class="audit-admin-actions"><button type="button" class="audited-btn" disabled>Audited</button><button type="button" class="danger" data-action="audit-mark-due" data-id="${a.id}">Set Due Now</button></div>`
        : `<button type="button" class="secondary" data-action="audit-mark-audited" data-id="${a.id}">Mark Audited Today</button>`)
      : (completed && !due ? `<span class="audited-text">Audited</span>` : "-");
    return `
      <tr>
        <td>${escapeHtml(a.assetTag)}</td>
        <td>${escapeHtml(a.assignment?.systemName || a.name)}</td>
        <td>${escapeHtml(a.assignment?.department || "-")}</td>
        <td>${escapeHtml(a.assignment?.owner || "-")}</td>
        <td>${fmtDate(last?.auditDate || last?.doneAt || last?.date)}</td>
        <td>${fmtDate(dueAt)}</td>
        <td>${completed ? "Yes" : "No"}</td>
        <td>${due ? "Due" : "Audited"}</td>
        <td>${viewReportBtn}</td>
        <td>${actionCell}</td>
      </tr>
    `;
  }).join("");
  setStatus(auditStatusMsg, `Total: ${filteredRows.length}. Due: ${dueCount}. Audited: ${filteredRows.length - dueCount}.`, "ok");
}

function viewAuditReportTemplate(assetId) {
  const asset = assets.find((a) => a.id === assetId);
  if (!asset) return;
  const report = getLastCompletedAuditForAsset(assetId);
  if (!report) {
    setStatus(auditStatusMsg, "No completed audit report for this system.", "err");
    return;
  }

  const fields = [
    ["Audit Date", fmtDate(report.auditDate || report.doneAt || report.date)],
    ["PC Name", report.pcName || asset.assignment?.systemName || asset.name || "-"],
    ["Department", report.department || asset.assignment?.department || "-"],
    ["User Name", report.userName || asset.assignment?.owner || "-"],
    ["OS Version", report.osVersion || asset.os || "-"],
    ["Windows Activated (Y/N)", report.windowsActivated || "-"],
    ["Admin Accounts Verified (Y/N)", report.adminAccountsVerified || "-"],
    ["Unauthorized Software (Y/N)", report.unauthorizedSoftware || "-"],
    ["Antivirus Active (Y/N)", report.antivirusActive || "-"],
    ["Firewall ON (Y/N)", report.firewallOn || "-"],
    ["Updates Pending (Y/N)", report.updatesPending || "-"],
    ["USB Restricted (Y/N)", report.usbRestricted || "-"],
    ["Printer Config OK (Y/N)", report.printerConfigOk || "-"],
    ["Backup Verified (Y/N)", report.backupVerified || "-"],
    ["Issues Found", report.issuesFound || "-"],
    ["Risk Level", report.riskLevel || "-"],
    ["Action Taken", report.actionTaken || "-"],
    ["Auditor Name", report.auditorName || "-"],
    ["Asset Tag", report.assetTag || asset.assetTag || "-"],
    ["Location", report.location || asset.location || "-"]
  ];

  auditReportViewTemplate.innerHTML = fields
    .map(([label, value]) => `<label>${escapeHtml(label)}<input type="text" readonly value="${escapeHtml(value)}" /></label>`)
    .join("");
  auditReportViewPanel.classList.remove("hidden");
}

function updateAuditStatusByAdmin(assetId, targetState) {
  if (!isAdminUser()) {
    setStatus(auditStatusMsg, "Only admins can change audit status.", "err");
    return;
  }
  const asset = assets.find((a) => a.id === assetId);
  if (!asset) return;
  const now = new Date();
  const record = {
    id: uid(),
    assetId: asset.id,
    assetTag: asset.assetTag,
    company: asset.company,
    location: asset.location,
    changedAt: now.toISOString(),
    changedBy: sessionUser,
    statusOverride: targetState
  };
  if (targetState === "audited") {
    const nextDue = new Date(now);
    nextDue.setDate(nextDue.getDate() + 30);
    record.auditDone = true;
    record.auditDate = now.toISOString();
    record.doneAt = now.toISOString();
    record.nextDueAt = nextDue.toISOString();
    record.auditorName = sessionUser || "";
  } else if (targetState === "due") {
    record.nextDueAt = now.toISOString();
  }
  audits.push(record);
  saveJson(AUDITS_KEY, audits);
  renderAuditStatusList();
  setStatus(auditStatusMsg, `Audit status updated for ${asset.assetTag}.`, "ok");
}

function getLastCleaned(assetId) {
  return cleaningRecords[assetId]?.lastCleaned || "";
}

function monthlyStatus(lastCleaned) {
  if (!lastCleaned) return "Due";
  const last = new Date(lastCleaned).getTime();
  const now = Date.now();
  return now - last > 30 * 24 * 60 * 60 * 1000 ? "Due" : "OK";
}

function renderCleaningTable() {
  const location = cleaningLocation.value;
  const allRows = assets
    .filter((a) => a.company === currentCompany && a.location === location && isDeviceType(a.type))
    .sort((a, b) => a.assetTag.localeCompare(b.assetTag));
  const { rows } = getPagedRows("cleaning-tbody", allRows);

  if (!allRows.length) {
    cleaningTbody.innerHTML = `<tr><td colspan="6">No laptop/desktop/mobile assets in selected location.</td></tr>`;
    return;
  }

  cleaningTbody.innerHTML = rows.map((a) => {
    const last = getLastCleaned(a.id);
    const state = monthlyStatus(last);
    const cleaned = state === "OK";
    return `
      <tr>
        <td>${escapeHtml(a.assetTag)}</td>
        <td>${escapeHtml(a.name)}</td>
        <td>${escapeHtml(a.type)}</td>
        <td>${last ? escapeHtml(last.slice(0, 10)) : "-"}</td>
        <td>${state}</td>
        <td>
          <button type="button" class="${cleaned ? "cleaned-btn" : "secondary"}" data-action="mark-cleaned" data-id="${a.id}" ${cleaned ? "disabled" : ""}>
            ${cleaned ? "Cleaned" : "Mark Cleaned"}
          </button>
        </td>
      </tr>
    `;
  }).join("");
}

function markAssetCleaned(id) {
  cleaningRecords[id] = { lastCleaned: new Date().toISOString(), updatedBy: sessionUser };
  saveJson(CLEANING_KEY, cleaningRecords);
  triggerChangeBackup("cleaning");
  const asset = assets.find((a) => a.id === id);
  logActivity("MARK_CLEANED", asset?.assetTag || id);
  renderCleaningTable();
  setStatus(cleaningStatus, "Cleaning status updated.", "ok");
}

function renderSoftwareMasterOptions() {
  const values = uniqueCaseInsensitive(softwareMaster);
  softwareMaster = values;
  const options = values.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  softwareNameSelect.innerHTML = options || `<option value="">No software in master list</option>`;
  deleteSoftwareSelect.innerHTML = `<option value="">Select software</option>${options}`;
  if (softwareSortSoftware) {
    const prev = softwareSortSoftware.value;
    softwareSortSoftware.innerHTML = `<option value="">All Software</option>${options}`;
    if (prev && values.includes(prev)) softwareSortSoftware.value = prev;
    else softwareSortSoftware.value = "";
  }
}

function updateSelectedSoftwareAssetDetails() {
  const assetId = softwareAssetSelect.value || selectedSoftwareAssetId || "";
  const asset = assets.find((a) => a.id === assetId) || null;
  softwareAssignedOwner.value = asset?.assignment?.owner || "-";
  softwareSelectedMac.value = asset?.macAddress || "-";
}

function showSoftwareSeatUsage() {
  const softwareName = String(softwareNameSelect?.value || "").trim();
  if (!softwareName) return;
  const seats = getPurchasedSeatTotal(softwareName);
  const used = getLicensedInstallCount(softwareName);
  if (seats > 0) {
    setStatus(softwareStatusMsg, `${softwareName}: seats used ${used}/${seats}.`, "ok");
  } else {
    setStatus(softwareStatusMsg, `${softwareName}: no purchased seats configured in Software Inventory.`, "err");
  }
}

function renderSoftwareAssetOptions() {
  const devices = getScopedDeviceAssets();
  const q = String(softwareAssetSearch?.value || "").trim().toLowerCase();
  const filtered = q
    ? devices.filter((a) => `${a.assetTag} ${a.name} ${a.macAddress || ""} ${a.assignment?.systemName || ""} ${a.assignment?.owner || ""}`.toLowerCase().includes(q))
    : devices;
  if (!devices.length) {
    softwareAssetSelect.innerHTML = `<option value="">No laptop/desktop/mobile assets</option>`;
    selectedSoftwareAssetId = null;
    updateSelectedSoftwareAssetDetails();
    return;
  }
  if (!filtered.length) {
    softwareAssetSelect.innerHTML = `<option value="">No matching assets</option>`;
    selectedSoftwareAssetId = null;
    updateSelectedSoftwareAssetDetails();
    return;
  }
  softwareAssetSelect.innerHTML = filtered
    .map((a) => `<option value="${a.id}">${escapeHtml(a.assetTag)} - ${escapeHtml(a.assignment?.systemName || a.name)}</option>`)
    .join("");
  if (selectedSoftwareAssetId && filtered.some((a) => a.id === selectedSoftwareAssetId)) {
    softwareAssetSelect.value = selectedSoftwareAssetId;
  } else {
    selectedSoftwareAssetId = filtered[0].id;
    softwareAssetSelect.value = selectedSoftwareAssetId;
  }
  updateSelectedSoftwareAssetDetails();
}

function renderSoftwareDeviceList() {
  const devices = getScopedDeviceAssets();
  if (!devices.length) {
    softwareDeviceTbody.innerHTML = `<tr><td colspan="7">No laptop/desktop/mobile assets in current company/location.</td></tr>`;
    return;
  }
  const q = String(softwareDeviceSearch?.value || "").trim().toLowerCase();
  const filtered = q
    ? devices.filter((a) => `${a.assetTag} ${a.name} ${a.type} ${a.assignment?.owner || ""} ${a.assignment?.systemName || ""} ${a.macAddress || ""}`.toLowerCase().includes(q))
    : devices;
  const { rows } = getPagedRows("software-device-tbody", filtered);
  if (!filtered.length) {
    softwareDeviceTbody.innerHTML = `<tr><td colspan="7">No matching devices.</td></tr>`;
    return;
  }
  softwareDeviceTbody.innerHTML = rows.map((a) => {
    const count = getScopedSoftwareInventory().filter((row) => row.assetId === a.id).length;
    return `<tr>
      <td>${escapeHtml(a.assetTag)}</td>
      <td>${escapeHtml(a.name)}</td>
      <td>${escapeHtml(a.type)}</td>
      <td>${escapeHtml(a.assignment?.owner || "-")}</td>
      <td>${escapeHtml(a.assignment?.systemName || "-")}</td>
      <td>${count}</td>
      <td><button type="button" class="secondary" data-action="software-view-device" data-id="${a.id}">View</button></td>
    </tr>`;
  }).join("");
}

function closeSoftwareDeviceViewModal() {
  if (!softwareDeviceViewModal) return;
  softwareDeviceViewModal.classList.add("hidden");
  softwareDeviceViewModal.setAttribute("aria-hidden", "true");
}

function openSoftwareDeviceViewModal(assetId) {
  const asset = assets.find((a) => a.id === assetId && a.company === currentCompany && a.location === currentLocation);
  if (!asset || !softwareDeviceViewModal || !softwareDeviceViewTbody) return;
  const rows = getScopedSoftwareInventory()
    .filter((row) => row.assetId === asset.id)
    .sort((a, b) => String(a.softwareName || "").localeCompare(String(b.softwareName || "")));
  const piratedCount = rows.filter((r) => String(r.pirated || "No").toLowerCase() === "yes").length;
  const okCount = rows.length - piratedCount;
  if (softwareDeviceViewTitle) {
    softwareDeviceViewTitle.textContent = `Installed Software: ${asset.assetTag} - ${asset.assignment?.systemName || asset.name}`;
  }
  if (softwareDeviceViewSummary) {
    softwareDeviceViewSummary.textContent = `Total: ${rows.length} | Not Pirated: ${okCount} | Pirated: ${piratedCount}`;
  }
  if (!rows.length) {
    softwareDeviceViewTbody.innerHTML = `<tr><td colspan="4">No software installed records for this system.</td></tr>`;
  } else {
    softwareDeviceViewTbody.innerHTML = rows.map((row) => `
      <tr>
        <td>${escapeHtml(row.softwareName || "-")}</td>
        <td>${escapeHtml(row.pirated || "No")}</td>
        <td>${escapeHtml(row.updatedBy || "-")}</td>
        <td>${fmtDate(row.updatedAt || row.createdAt)}</td>
      </tr>
    `).join("");
  }
  softwareDeviceViewModal.classList.remove("hidden");
  softwareDeviceViewModal.setAttribute("aria-hidden", "false");
}

function renderSoftwareInventoryTable() {
  const q = String(softwareSearchInput.value || "").trim().toLowerCase();
  const selectedId = softwareAssetSelect.value || selectedSoftwareAssetId || "";
  const scope = softwareViewScope?.value || "all";
  const softwareFilterName = softwareSortSoftware?.value || "";
  const piratedFilterMode = softwareSortPirated?.value || "all";
  const orderByMaster = new Map(softwareMaster.map((name, idx) => [String(name).toLowerCase(), idx]));
  const allRows = getScopedSoftwareInventory()
    .filter((row) => (scope === "selected" ? (!!selectedId && row.assetId === selectedId) : true))
    .filter((row) => (softwareFilterName ? String(row.softwareName || "").toLowerCase() === softwareFilterName.toLowerCase() : true))
    .filter((row) => {
      if (piratedFilterMode === "yes") return String(row.pirated || "No").toLowerCase() === "yes";
      if (piratedFilterMode === "no") return String(row.pirated || "No").toLowerCase() !== "yes";
      return true;
    })
    .filter((row) => {
      if (!q) return true;
      const linkedOwner = assets.find((a) => a.id === row.assetId)?.assignment?.owner || "";
      const hay = `${row.assetTag || ""} ${row.assetName || ""} ${row.assetSystemName || ""} ${row.assetOwner || linkedOwner} ${row.assetMac || ""} ${row.softwareName || ""} ${row.type || ""}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a, b) => {
      const aKey = String(a.softwareName || "").toLowerCase();
      const bKey = String(b.softwareName || "").toLowerCase();
      const aIdx = orderByMaster.has(aKey) ? orderByMaster.get(aKey) : Number.MAX_SAFE_INTEGER;
      const bIdx = orderByMaster.has(bKey) ? orderByMaster.get(bKey) : Number.MAX_SAFE_INTEGER;
      if (aIdx !== bIdx) return aIdx - bIdx;
      const nameCmp = String(a.softwareName || "").localeCompare(String(b.softwareName || ""));
      if (nameCmp !== 0) return nameCmp;
      return String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || ""));
    });

  const { rows } = getPagedRows("software-tbody", allRows);
  if (!allRows.length) {
    softwareTbody.innerHTML = `<tr><td colspan="8">No software records for this view.</td></tr>`;
    return;
  }
  softwareTbody.innerHTML = rows.map((row) => `<tr>
    <td>${escapeHtml(row.assetTag || "-")}</td>
    <td>${escapeHtml(row.assetSystemName || row.assetName || "-")}</td>
    <td>${escapeHtml(row.type || "-")}</td>
    <td>${escapeHtml(row.softwareName || "-")}</td>
    <td>${escapeHtml(row.pirated || "No")}</td>
    <td>${escapeHtml(row.updatedBy || "-")}</td>
    <td>${fmtDate(row.updatedAt || row.createdAt)}</td>
    <td><button type="button" class="danger" data-action="software-delete" data-id="${row.id}">Delete</button></td>
  </tr>`).join("");
}

function renderSoftwareCheckPage(preselectAssetId = null) {
  if (!softwareAssetSelect) return;
  if (preselectAssetId) selectedSoftwareAssetId = preselectAssetId;
  renderSoftwareMasterOptions();
  renderSoftwareAssetOptions();
  renderSoftwareDeviceList();
  renderSoftwareInventoryTable();
  renderInventorySoftwareOptions();
  showSoftwareSeatUsage();
}

function addSoftwareMasterByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can add software names.", "err");
    return;
  }
  const name = String(newSoftwareNameInput.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Software name is required.", "err");
    return;
  }
  if (softwareMaster.some((s) => s.toLowerCase() === name.toLowerCase())) {
    setStatus(adminSettingsStatus, "Software already exists.", "err");
    return;
  }
  softwareMaster.push(name);
  persistSoftwareData();
  refreshAdminOptionLists();
  logActivity("ADD_SOFTWARE_MASTER", name);
  newSoftwareNameInput.value = "";
  renderSoftwareCheckPage();
  setStatus(adminSettingsStatus, `Added software: ${name}.`, "ok");
}

function deleteSoftwareMasterByAdmin() {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can delete software names.", "err");
    return;
  }
  const name = String(deleteSoftwareSelect.value || "").trim();
  if (!name) {
    setStatus(adminSettingsStatus, "Select software to delete.", "err");
    return;
  }
  softwareMaster = softwareMaster.filter((s) => s.toLowerCase() !== name.toLowerCase());
  softwareInventory = softwareInventory.filter((row) => String(row.softwareName || "").toLowerCase() !== name.toLowerCase());
  softwarePurchases = softwarePurchases.filter((row) => String(row.softwareName || "").toLowerCase() !== name.toLowerCase());
  persistSoftwareData();
  refreshAdminOptionLists();
  logActivity("DELETE_SOFTWARE_MASTER", name);
  renderSoftwareCheckPage();
  setStatus(adminSettingsStatus, `Deleted software: ${name}.`, "ok");
}

function saveSoftwareInventory(event) {
  event.preventDefault();
  const assetId = softwareAssetSelect.value || selectedSoftwareAssetId;
  const asset = assets.find((a) => a.id === assetId && a.company === currentCompany && a.location === currentLocation);
  if (!asset || !isDeviceType(asset.type)) {
    setStatus(softwareStatusMsg, "Select a valid laptop/desktop/mobile system.", "err");
    return;
  }
  const softwareName = String(softwareNameSelect.value || "").trim();
  if (!softwareName) {
    setStatus(softwareStatusMsg, "Add software in master list first.", "err");
    return;
  }
  const pirated = softwarePiratedSelect.value === "Yes" ? "Yes" : "No";
  const existing = softwareInventory.find((row) => row.assetId === asset.id && row.softwareName.toLowerCase() === softwareName.toLowerCase());
  const piratedBypass = pirated === "Yes";
  if (!piratedBypass) {
    const seats = getPurchasedSeatTotal(softwareName);
    const used = getLicensedInstallCount(softwareName);
    const existingCounts = existing && String(existing.pirated || "No").toLowerCase() !== "yes" ? 1 : 0;
    const nextUsed = used - existingCounts + 1;
    if (seats <= 0) {
      setStatus(softwareStatusMsg, `No purchased seats found for ${softwareName}. Add seats in Software Inventory.`, "err");
      return;
    }
    if (nextUsed > seats) {
      setStatus(softwareStatusMsg, `Seat limit reached for ${softwareName}. Used ${used}/${seats}.`, "err");
      return;
    }
  }
  const now = new Date().toISOString();
  if (existing) {
    existing.pirated = pirated;
    existing.assetTag = asset.assetTag;
    existing.assetName = asset.name;
    existing.assetSystemName = asset.assignment?.systemName || "";
    existing.assetOwner = asset.assignment?.owner || "";
    existing.assetMac = asset.macAddress || "";
    existing.type = asset.type;
    existing.updatedAt = now;
    existing.updatedBy = sessionUser || "";
  } else {
    softwareInventory.push({
      id: uid(),
      assetId: asset.id,
      assetTag: asset.assetTag,
      assetName: asset.name,
      assetSystemName: asset.assignment?.systemName || "",
      assetOwner: asset.assignment?.owner || "",
      assetMac: asset.macAddress || "",
      type: asset.type,
      company: asset.company,
      location: asset.location,
      softwareName,
      pirated,
      createdAt: now,
      updatedAt: now,
      updatedBy: sessionUser || ""
    });
  }
  persistSoftwareData();
  logActivity("SAVE_SOFTWARE_CHECK", `${asset.assetTag} / ${softwareName} / pirated=${pirated}`);
  renderSoftwareDeviceList();
  renderSoftwareInventoryTable();
  renderSoftwareInventoryList();
  if (piratedBypass) {
    setStatus(softwareStatusMsg, `Saved software for ${asset.assetTag}. Pirated entry bypassed purchased-seat checks.`, "ok");
  } else {
    const seats = getPurchasedSeatTotal(softwareName);
    const used = getLicensedInstallCount(softwareName);
    const seatMsg = seats > 0 ? ` Seats used ${used}/${seats}.` : "";
    setStatus(softwareStatusMsg, `Saved software for ${asset.assetTag}.${seatMsg}`, "ok");
  }
}

function openSoftwareCheckForAsset(assetId) {
  const asset = assets.find((a) => a.id === assetId);
  if (!asset || !isDeviceType(asset.type)) {
    setStatus(catalogStatusMsg, "Software check is only for laptop/desktop/mobile.", "err");
    return;
  }
  const opened = setActivePage("software-check");
  if (!opened) return;
  if (softwareAssetSearch) softwareAssetSearch.value = "";
  selectedSoftwareAssetId = assetId;
  renderSoftwareCheckPage(assetId);
}

function deleteSoftwareInventoryRow(id) {
  const target = softwareInventory.find((row) => row.id === id);
  if (!target) return;
  softwareInventory = softwareInventory.filter((row) => row.id !== id);
  persistSoftwareData();
  logActivity("DELETE_SOFTWARE_CHECK", `${target.assetTag} / ${target.softwareName}`);
  renderSoftwareDeviceList();
  renderSoftwareInventoryTable();
  setStatus(softwareStatusMsg, `Deleted ${target.softwareName} from ${target.assetTag}.`, "ok");
}

function purgeSoftwareForDeletedAssets(validAssetIds = new Set(assets.map((a) => a.id))) {
  const before = softwareInventory.length;
  softwareInventory = softwareInventory.filter((row) => validAssetIds.has(row.assetId));
  softwarePurchases = softwarePurchases.map((row) => ({
    ...row,
    assignedSystemIds: Array.isArray(row.assignedSystemIds) ? row.assignedSystemIds.filter((id) => validAssetIds.has(id)) : []
  }));
  if (softwareInventory.length !== before) persistSoftwareData();
}

function renderInventorySoftwareOptions() {
  const values = uniqueCaseInsensitive(softwareMaster).sort((a, b) => a.localeCompare(b));
  if (!values.length) {
    inventorySoftwareName.innerHTML = `<option value="">No software in master list</option>`;
    if (inventoryModalSoftwareName) inventoryModalSoftwareName.innerHTML = `<option value="">No software in master list</option>`;
    return;
  }
  const options = values.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  inventorySoftwareName.innerHTML = options;
  if (inventoryModalSoftwareName) inventoryModalSoftwareName.innerHTML = options;
}

function renderInventorySystemOptions() {
  if (!inventoryAssignedSystems) return;
  const devices = getScopedDeviceAssets();
  const fromUi = new Set(Array.from(inventoryAssignedSystems?.selectedOptions || []).map((opt) => opt.value));
  const fromEdit = Array.isArray(softwarePurchases.find((row) => row.id === editingSoftwarePurchaseId)?.assignedSystemIds)
    ? softwarePurchases.find((row) => row.id === editingSoftwarePurchaseId).assignedSystemIds
    : [];
  const selected = fromUi.size ? fromUi : new Set(fromEdit);
  const q = String(inventoryAssignedSearch?.value || "").trim().toLowerCase();
  const filtered = q
    ? devices.filter((a) => `${a.assetTag} ${a.name} ${a.assignment?.owner || ""} ${a.assignment?.systemName || ""}`.toLowerCase().includes(q))
    : devices;
  if (!devices.length) {
    inventoryAssignedSystems.innerHTML = "";
    return;
  }
  inventoryAssignedSystems.innerHTML = filtered
    .map((a) => `<option value="${a.id}" ${selected.has(a.id) ? "selected" : ""}>${escapeHtml(a.assetTag)} - ${escapeHtml(a.assignment?.systemName || a.name)}</option>`)
    .join("");
}

function renderInventoryModalSystemOptions() {
  if (!inventoryModalAssignedSystems) return;
  const devices = getScopedDeviceAssets();
  const q = String(inventoryModalAssignedSearch?.value || "").trim().toLowerCase();
  const filtered = q
    ? devices.filter((a) => `${a.assetTag} ${a.name} ${a.assignment?.owner || ""} ${a.assignment?.systemName || ""}`.toLowerCase().includes(q))
    : devices;
  if (!devices.length) {
    inventoryModalAssignedSystems.innerHTML = "";
    return;
  }
  inventoryModalAssignedSystems.innerHTML = filtered
    .map((a) => `<option value="${a.id}" ${modalInventorySelectedIds.has(a.id) ? "selected" : ""}>${escapeHtml(a.assetTag)} - ${escapeHtml(a.assignment?.systemName || a.name)}</option>`)
    .join("");
}

function openInventoryEditModal(row) {
  if (!inventoryEditModal || !row) return;
  editingSoftwarePurchaseId = row.id;
  renderInventorySoftwareOptions();
  inventoryModalSoftwareName.value = row.softwareName || "";
  inventoryModalPurchaseDate.value = row.purchaseDate || "";
  inventoryModalExpiryDate.value = row.expiryDate || "";
  inventoryModalCost.value = row.cost || "";
  inventoryModalActivationCode.value = row.activationCode || "";
  if (inventoryModalMaxUsers) inventoryModalMaxUsers.value = row.maxUsers || "";
  setStatus(inventoryModalStatusMsg, "", "ok");
  inventoryEditModal.classList.remove("hidden");
  inventoryEditModal.setAttribute("aria-hidden", "false");
}

function closeInventoryEditModal() {
  if (!inventoryEditModal) return;
  inventoryEditModal.classList.add("hidden");
  inventoryEditModal.setAttribute("aria-hidden", "true");
  modalInventorySelectedIds = new Set();
  editingSoftwarePurchaseId = null;
}

function saveInventoryModalEdit(event) {
  event.preventDefault();
  if (!editingSoftwarePurchaseId) return;
  const existing = softwarePurchases.find((r) => r.id === editingSoftwarePurchaseId && r.company === currentCompany && r.location === currentLocation);
  if (!existing) {
    setStatus(inventoryModalStatusMsg, "Record not found.", "err");
    return;
  }
  const softwareName = String(inventoryModalSoftwareName.value || "").trim();
  if (!softwareName) {
    setStatus(inventoryModalStatusMsg, "Select software name.", "err");
    return;
  }
  const maxUsersValue = Math.max(0, Number.parseInt(inventoryModalMaxUsers?.value || "0", 10) || 0);
  const used = getLicensedInstallCount(softwareName, currentCompany, currentLocation);
  const seatsOther = getPurchasedSeatTotalExcluding(editingSoftwarePurchaseId, softwareName, currentCompany, currentLocation);
  if (used > seatsOther + maxUsersValue) {
    setStatus(inventoryModalStatusMsg, `Cannot save. ${softwareName} is already used on ${used} systems, but seats would be ${seatsOther + maxUsersValue}.`, "err");
    return;
  }
  const payload = {
    ...existing,
    softwareName,
    purchaseDate: inventoryModalPurchaseDate.value || "",
    expiryDate: inventoryModalExpiryDate.value || "",
    cost: inventoryModalCost.value || "",
    activationCode: inventoryModalActivationCode.value.trim(),
    maxUsers: String(maxUsersValue),
    assignedSystemIds: Array.isArray(existing.assignedSystemIds) ? existing.assignedSystemIds : [],
    updatedBy: sessionUser || "",
    updatedAt: new Date().toISOString()
  };
  softwarePurchases = softwarePurchases.map((row) => (row.id === editingSoftwarePurchaseId ? payload : row));
  persistSoftwareData();
  logActivity("UPDATE_SOFTWARE_INVENTORY", softwareName);
  renderSoftwareInventoryList();
  closeInventoryEditModal();
  setStatus(inventoryStatusMsg, `Updated ${softwareName}.`, "ok");
}

function formatMoney(value) {
  if (value === "" || value === null || value === undefined) return "-";
  const n = Number(value);
  if (!Number.isFinite(n)) return "-";
  return n.toFixed(2);
}

function getPurchasedSeatTotal(softwareName, company = currentCompany, location = currentLocation) {
  const key = String(softwareName || "").trim().toLowerCase();
  if (!key) return 0;
  return softwarePurchases
    .filter((row) => row.company === company && row.location === location)
    .filter((row) => String(row.softwareName || "").trim().toLowerCase() === key)
    .reduce((sum, row) => sum + Math.max(0, Number.parseInt(row.maxUsers, 10) || 0), 0);
}

function getPurchasedSeatTotalExcluding(purchaseId, softwareName, company = currentCompany, location = currentLocation) {
  const key = String(softwareName || "").trim().toLowerCase();
  if (!key) return 0;
  return softwarePurchases
    .filter((row) => row.id !== purchaseId)
    .filter((row) => row.company === company && row.location === location)
    .filter((row) => String(row.softwareName || "").trim().toLowerCase() === key)
    .reduce((sum, row) => sum + Math.max(0, Number.parseInt(row.maxUsers, 10) || 0), 0);
}

function getLicensedInstallCount(softwareName, company = currentCompany, location = currentLocation) {
  const key = String(softwareName || "").trim().toLowerCase();
  if (!key) return 0;
  return softwareInventory.filter((row) => (
    row.company === company &&
    row.location === location &&
    String(row.softwareName || "").trim().toLowerCase() === key &&
    String(row.pirated || "No").toLowerCase() !== "yes"
  )).length;
}

function getSystemSummary(ids = []) {
  const labels = ids
    .map((id) => assets.find((a) => a.id === id))
    .filter((a) => a && a.company === currentCompany && a.location === currentLocation && isDeviceType(a.type))
    .map((a) => `${a.assetTag} (${a.assignment?.systemName || a.name})`);
  return labels.length ? labels.join(", ") : "-";
}

function resetSoftwareInventoryForm() {
  if (!softwareInventoryForm) return;
  editingSoftwarePurchaseId = null;
  softwareInventoryForm.reset();
  if (inventoryMaxUsers) inventoryMaxUsers.value = "";
  inventorySaveBtn.textContent = "Save Inventory";
  inventoryCancelBtn.hidden = true;
  renderInventorySoftwareOptions();
}

function renderSoftwareInventoryList() {
  const q = String(inventorySearchInput.value || "").trim().toLowerCase();
  const allRows = softwarePurchases
    .filter((row) => row.company === currentCompany && row.location === currentLocation)
    .filter((row) => {
      if (!q) return true;
      const hay = `${row.softwareName || ""} ${row.activationCode || ""}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")));

  const { rows } = getPagedRows("inventory-tbody", allRows);
  if (!allRows.length) {
    inventoryTbody.innerHTML = `<tr><td colspan="11">No software purchase records.</td></tr>`;
    return;
  }
  inventoryTbody.innerHTML = rows.map((row) => `<tr>
    <td>${escapeHtml(row.softwareName || "-")}</td>
    <td>${fmtDate(row.purchaseDate)}</td>
    <td>${fmtDate(row.expiryDate)}</td>
    <td>${escapeHtml(formatMoney(row.cost))}</td>
    <td>${escapeHtml(row.activationCode || "-")}</td>
    <td>${escapeHtml(row.maxUsers || "0")}</td>
    <td>${getLicensedInstallCount(row.softwareName, row.company, row.location)}</td>
    <td>${Math.max(0, (Number.parseInt(row.maxUsers, 10) || 0) - getLicensedInstallCount(row.softwareName, row.company, row.location))}</td>
    <td>${escapeHtml(row.updatedBy || "-")}</td>
    <td>${fmtDate(row.updatedAt || row.createdAt)}</td>
    <td>
      <button type="button" class="secondary" data-action="inventory-edit" data-id="${row.id}">Edit</button>
      <button type="button" class="danger" data-action="inventory-delete" data-id="${row.id}">Delete</button>
    </td>
  </tr>`).join("");
}

function renderSoftwareInventoryPage() {
  if (!inventoryTbody) return;
  renderInventorySoftwareOptions();
  renderSoftwareInventoryList();
}

function getSelectedInventorySystemIds() {
  if (!inventoryAssignedSystems) return [];
  return Array.from(inventoryAssignedSystems.selectedOptions || []).map((o) => o.value).filter(Boolean);
}

function saveSoftwarePurchase(event) {
  event.preventDefault();
  const softwareName = String(inventorySoftwareName.value || "").trim();
  if (!softwareName) {
    setStatus(inventoryStatusMsg, "Select software name.", "err");
    return;
  }
  const maxUsersValue = Math.max(0, Number.parseInt(inventoryMaxUsers?.value || "0", 10) || 0);
  const used = getLicensedInstallCount(softwareName, currentCompany, currentLocation);
  const seatsOther = getPurchasedSeatTotalExcluding(editingSoftwarePurchaseId || "", softwareName, currentCompany, currentLocation);
  if (used > seatsOther + maxUsersValue) {
    setStatus(inventoryStatusMsg, `Cannot save. ${softwareName} is already used on ${used} systems, but seats would be ${seatsOther + maxUsersValue}.`, "err");
    return;
  }
  const payload = {
    id: editingSoftwarePurchaseId || uid(),
    softwareName,
    purchaseDate: inventoryPurchaseDate.value || "",
    expiryDate: inventoryExpiryDate.value || "",
    cost: inventoryCost.value || "",
    activationCode: inventoryActivationCode.value.trim(),
    maxUsers: String(maxUsersValue),
    assignedSystemIds: editingSoftwarePurchaseId
      ? (softwarePurchases.find((r) => r.id === editingSoftwarePurchaseId)?.assignedSystemIds || [])
      : [],
    company: currentCompany,
    location: currentLocation,
    updatedBy: sessionUser || "",
    updatedAt: new Date().toISOString(),
    createdAt: editingSoftwarePurchaseId
      ? (softwarePurchases.find((r) => r.id === editingSoftwarePurchaseId)?.createdAt || new Date().toISOString())
      : new Date().toISOString()
  };

  if (editingSoftwarePurchaseId) {
    softwarePurchases = softwarePurchases.map((row) => (row.id === editingSoftwarePurchaseId ? payload : row));
    logActivity("UPDATE_SOFTWARE_INVENTORY", softwareName);
    setStatus(inventoryStatusMsg, `Updated ${softwareName}.`, "ok");
  } else {
    softwarePurchases.push(payload);
    logActivity("ADD_SOFTWARE_INVENTORY", softwareName);
    setStatus(inventoryStatusMsg, `Saved ${softwareName}.`, "ok");
  }
  persistSoftwareData();
  resetSoftwareInventoryForm();
  renderSoftwareInventoryList();
}

function editSoftwarePurchase(id) {
  const row = softwarePurchases.find((r) => r.id === id && r.company === currentCompany && r.location === currentLocation);
  if (!row) return;
  openInventoryEditModal(row);
}

function deleteSoftwarePurchase(id) {
  const row = softwarePurchases.find((r) => r.id === id);
  if (!row) return;
  if (!window.confirm(`Delete software inventory: ${row.softwareName}?`)) return;
  softwarePurchases = softwarePurchases.filter((r) => r.id !== id);
  persistSoftwareData();
  logActivity("DELETE_SOFTWARE_INVENTORY", row.softwareName);
  if (editingSoftwarePurchaseId === id) resetSoftwareInventoryForm();
  renderSoftwareInventoryList();
  setStatus(inventoryStatusMsg, "Software inventory deleted.", "ok");
}

loginForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const username = document.getElementById("login-username").value.trim();
  const password = document.getElementById("login-password").value;
  const user = users.find((u) => u.username === username && u.password === password);
  if (!user) return setStatus(authStatus, "Invalid username or password.", "err");

  sessionUser = username;
  sessionStorage.setItem(SESSION_KEY, username);
  localStorage.removeItem(SESSION_KEY);
  logActivity("LOGIN", "User login successful", { user: username, company: currentCompany || "", location: currentLocation || "" });
  setStatus(authStatus, "Login successful.", "ok");
  route();
});

companyGrid?.addEventListener("click", (event) => {
  const btn = event.target.closest(".company-option[data-company]");
  if (!btn) return;
  currentCompany = btn.dataset.company || "";
  localStorage.setItem(COMPANY_KEY, currentCompany);
  logActivity("SELECT_COMPANY", currentCompany);
  setStatus(companyStatus, `Selected ${currentCompany}.`, "ok");
  showScreen("module");
  refreshModuleMeta();
  adminModuleCard?.classList.toggle("hidden", !isAdminUser());
  activityModuleCard?.classList.toggle("hidden", !isAdminUser());
  moduleAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
  moduleAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
  refreshLocationOptions();
  moduleLocationSelect.value = currentLocation || getDefaultLocation();
});

moduleCards.forEach((card) => {
  card.addEventListener("click", () => {
    const module = card.dataset.module;
    if ((module === "admin-users" || module === "activity-history") && !isAdminUser()) {
      setStatus(moduleStatus, "Only admins can open this module.", "err");
      return;
    }
    const location = moduleLocationSelect.value;
    if (!location) {
      setStatus(moduleStatus, "Select location.", "err");
      return;
    }
    currentLocation = location;
    activePage = module;
    localStorage.setItem(LOCATION_KEY, location);
    localStorage.setItem(MODULE_KEY, module);
    logActivity("OPEN_MODULE", `${module} @ ${location}`);
    setStatus(moduleStatus, `Opening ${module} for ${location}.`, "ok");
    route();
  });
});

Object.entries(tabs).forEach(([key, btn]) => {
  btn?.addEventListener("click", () => setActivePage(key));
});

goModuleSelectorBtn.addEventListener("click", () => {
  showScreen("module");
  refreshModuleMeta();
  adminModuleCard?.classList.toggle("hidden", !isAdminUser());
  activityModuleCard?.classList.toggle("hidden", !isAdminUser());
  moduleAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
  moduleAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
  refreshLocationOptions();
  moduleLocationSelect.value = currentLocation || getDefaultLocation();
});

moduleAdminSettingsBtn.addEventListener("click", () => {
  openAdminSettingsFromAnywhere("module");
});
companyAdminSettingsBtn.addEventListener("click", () => openAdminSettingsFromAnywhere("company"));
appAdminSettingsBtn.addEventListener("click", () => openAdminSettingsFromAnywhere("app"));
companyAdminUsersBtn?.addEventListener("click", openAdminUsersFromShortcut);
moduleAdminUsersBtn?.addEventListener("click", openAdminUsersFromShortcut);
appAdminUsersBtn?.addEventListener("click", () => setActivePage("admin-users"));

settingsBackBtn.addEventListener("click", () => {
  if (settingsReturnScreen === "company") {
    showScreen("company");
    companyAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
    return;
  }
  if (settingsReturnScreen === "app" && currentCompany && currentLocation) {
    showScreen("app");
    refreshHeader();
    setActivePage(activePage);
    return;
  }
  showScreen("module");
  refreshModuleMeta();
  adminModuleCard?.classList.toggle("hidden", !isAdminUser());
  activityModuleCard?.classList.toggle("hidden", !isAdminUser());
  moduleAdminUsersBtn?.classList.toggle("hidden", !isAdminUser());
  moduleAdminSettingsBtn.classList.toggle("hidden", !isAdminUser());
  refreshLocationOptions();
  moduleLocationSelect.value = currentLocation || getDefaultLocation();
});
settingsLogoutBtn.addEventListener("click", performLogout);

downloadFullBackupBtn.addEventListener("click", () => {
  triggerManualBackup("manual-settings");
});
selectSaveFolderBtn?.addEventListener("click", () => {
  void selectSaveFolderByAdmin();
});
saveFolderNowBtn?.addEventListener("click", () => {
  void saveToFolderNowByAdmin();
});
addCompanyBtn?.addEventListener("click", addCompanyByAdmin);
addAssetTypeBtn?.addEventListener("click", addAssetTypeByAdmin);
addAssetStatusBtn?.addEventListener("click", addAssetStatusByAdmin);
addLocationBtn?.addEventListener("click", addLocationByAdmin);
addDesignationBtn?.addEventListener("click", addDesignationByAdmin);
addWorkSiteBtn?.addEventListener("click", addWorkSiteByAdmin);
deleteDepartmentBtn?.addEventListener("click", deleteDepartmentByAdmin);
deleteLocationBtn?.addEventListener("click", deleteLocationByAdmin);
renameLocationBtn?.addEventListener("click", renameLocationByAdmin);
renameDepartmentBtn?.addEventListener("click", renameDepartmentByAdmin);
renameAssetTypeBtn?.addEventListener("click", renameAssetTypeByAdmin);
renameStatusBtn?.addEventListener("click", renameStatusByAdmin);
renameSoftwareBtn?.addEventListener("click", renameSoftwareByAdmin);
renameDesignationBtn?.addEventListener("click", renameDesignationByAdmin);
deleteAssetTypeBtn?.addEventListener("click", deleteAssetTypeByAdmin);
deleteStatusBtn?.addEventListener("click", deleteStatusByAdmin);
deleteDesignationBtn?.addEventListener("click", deleteDesignationByAdmin);
deleteWorkSiteBtn?.addEventListener("click", deleteWorkSiteByAdmin);
employeeSortSelect?.addEventListener("change", renderEmployeeTable);
softwareCheckLockToggle?.addEventListener("change", () => {
  if (!isAdminUser()) {
    softwareCheckLockToggle.checked = softwareCheckRequiresAdminPassword;
    setStatus(adminSettingsStatus, "Only admins can change this setting.", "err");
    return;
  }
  softwareCheckRequiresAdminPassword = !!softwareCheckLockToggle.checked;
  saveJson(SOFTWARE_CHECK_LOCK_KEY, softwareCheckRequiresAdminPassword);
  logActivity("UPDATE_SOFTWARE_CHECK_LOCK", softwareCheckRequiresAdminPassword ? "enabled" : "disabled");
  setStatus(
    adminSettingsStatus,
    softwareCheckRequiresAdminPassword
      ? "Software is now locked with admin password."
      : "Software lock disabled.",
    "ok"
  );
});

restoreBackupFileBtn.addEventListener("click", restoreFromBackupFile);
cloudSaveConfigBtn?.addEventListener("click", () => {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can update cloud sync settings.", "err");
    return;
  }
  saveCloudSyncConfig();
  startCloudAutoPull();
  setStatus(adminSettingsStatus, "Cloud sync config saved.", "ok");
});
cloudPushNowBtn?.addEventListener("click", async () => {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can push cloud sync.", "err");
    return;
  }
  saveCloudSyncConfig();
  const ok = await pushCloudState("manual");
  if (ok) setStatus(adminSettingsStatus, "Cloud push completed.", "ok");
});
cloudPullNowBtn?.addEventListener("click", async () => {
  if (!isAdminUser()) {
    setStatus(adminSettingsStatus, "Only admins can pull cloud sync.", "err");
    return;
  }
  saveCloudSyncConfig();
  const ok = await pullCloudState("manual");
  if (ok) setStatus(adminSettingsStatus, "Cloud pull completed.", "ok");
});

moduleChangeCompanyBtn.addEventListener("click", () => {
  currentCompany = null;
  currentLocation = null;
  localStorage.removeItem(COMPANY_KEY);
  localStorage.removeItem(LOCATION_KEY);
  route();
});

moduleResetAllBtn.addEventListener("click", hardResetAllData);

changeCompanyBtn.addEventListener("click", () => {
  currentCompany = null;
  currentLocation = null;
  localStorage.removeItem(COMPANY_KEY);
  localStorage.removeItem(LOCATION_KEY);
  route();
});
companyLogo.addEventListener("click", () => {
  currentCompany = null;
  currentLocation = null;
  localStorage.removeItem(COMPANY_KEY);
  localStorage.removeItem(LOCATION_KEY);
  route();
});

logoutBtn.addEventListener("click", () => {
  performLogout();
});
authLogoutBtn?.addEventListener("click", performLogout);
companyLogoutBtn.addEventListener("click", performLogout);
moduleLogoutBtn.addEventListener("click", performLogout);

catalogFields.type.addEventListener("change", () => {
  updateCatalogTypeVisibility();
  if (!editingCatalogId) catalogFields.tag.value = getNextTag(catalogFields.type.value);
});
catalogFields.printerMode.addEventListener("change", updateCatalogTypeVisibility);
catalogFields.status.addEventListener("change", updateCatalogTypeVisibility);
catalogAssignedPhotoInput?.addEventListener("change", async () => {
  const file = catalogAssignedPhotoInput.files?.[0];
  if (!file) {
    catalogAssignedPhotoDataUrl = "";
    if (catalogAssignedPhotoPreview) {
      catalogAssignedPhotoPreview.hidden = true;
      catalogAssignedPhotoPreview.removeAttribute("src");
    }
    return;
  }
  try {
    catalogAssignedPhotoDataUrl = await readCompressedImage(file);
    if (catalogAssignedPhotoPreview) {
      catalogAssignedPhotoPreview.hidden = false;
      catalogAssignedPhotoPreview.src = catalogAssignedPhotoDataUrl;
    }
  } catch (error) {
    setStatus(catalogStatusMsg, error.message, "err");
  }
});
catalogEditAssignedPhotoInput?.addEventListener("change", async () => {
  const file = catalogEditAssignedPhotoInput.files?.[0];
  if (!file) return;
  try {
    catalogEditAssignedPhotoDataUrl = await readCompressedImage(file);
    if (catalogEditAssignedPhotoPreview) {
      catalogEditAssignedPhotoPreview.hidden = false;
      catalogEditAssignedPhotoPreview.src = catalogEditAssignedPhotoDataUrl;
    }
  } catch (error) {
    setStatus(catalogEditStatusMsg, error.message, "err");
  }
});

toggleAdminInputBtn.addEventListener("click", () => {
  const show = catalogFields.adminPassword.type === "password";
  catalogFields.adminPassword.type = show ? "text" : "password";
  toggleAdminInputBtn.textContent = show ? "Hide" : "Show";
});

catalogForm.addEventListener("submit", saveCatalogAsset);
catalogCancelBtn.addEventListener("click", resetCatalogForm);
catalogDeleteAllBtn.addEventListener("click", deleteScopedAssets);
catalogImportCsvInput.addEventListener("change", importCatalogCsv);
catalogSampleCsvBtn.addEventListener("click", downloadSampleCsv);
catalogExportCsvBtn.addEventListener("click", exportCatalogCsv);
catalogUndoLastBulkBtn?.addEventListener("click", undoLastCatalogBulkUpload);
catalogUndoAllBulkBtn?.addEventListener("click", undoAllCatalogBulkUploads);
catalogDeleteBulkSelectedBtn?.addEventListener("click", deleteSelectedCatalogBulkUpload);
catalogViewStatus.addEventListener("change", renderCatalogList);
catalogSearchInput?.addEventListener("input", renderCatalogList);
catalogSelectAllBtn?.addEventListener("click", toggleSelectAllCatalogVisible);
catalogDeleteSelectedBtn?.addEventListener("click", deleteSelectedCatalogAssets);
catalogEditFields.type?.addEventListener("change", updateCatalogEditTypeVisibility);
catalogEditFields.printerMode?.addEventListener("change", updateCatalogEditTypeVisibility);
catalogEditFields.status?.addEventListener("change", updateCatalogEditTypeVisibility);
catalogEditForm?.addEventListener("submit", saveCatalogEditModal);
catalogEditModalClose?.addEventListener("click", closeCatalogEditModal);
catalogEditModalCancel?.addEventListener("click", closeCatalogEditModal);
catalogEditModal?.addEventListener("click", (event) => {
  if (event.target === catalogEditModal) closeCatalogEditModal();
});
catalogListTbody?.addEventListener("click", (event) => {
  const check = event.target.closest("input[type='checkbox'][data-action='catalog-select'][data-id]");
  if (check) {
    if (check.checked) selectedCatalogAssetIds.add(check.dataset.id);
    else selectedCatalogAssetIds.delete(check.dataset.id);
    return;
  }
  const btn = event.target.closest("button[data-action][data-id]");
  if (!btn) return;
  if (btn.dataset.action === "catalog-software") {
    openSoftwareCheckForAsset(btn.dataset.id);
    return;
  }
  if (btn.dataset.action === "catalog-edit") {
    startCatalogEdit(btn.dataset.id);
    return;
  }
  if (btn.dataset.action === "catalog-delete") {
    deleteOneAsset(btn.dataset.id);
  }
});
employeeForm?.addEventListener("submit", addEmployee);
employeeCancelBtn?.addEventListener("click", cancelEmployeeEdit);
employeeImportCsvInput?.addEventListener("change", importEmployeeCsv);
employeeSampleCsvBtn?.addEventListener("click", downloadEmployeeSampleCsv);
employeeSearchInput?.addEventListener("input", renderEmployeeTable);
employeeSelectAllBtn?.addEventListener("click", toggleSelectAllEmployees);
employeeDeleteSelectedBtn?.addEventListener("click", deleteSelectedEmployees);
employeeUndoLastBulkBtn?.addEventListener("click", undoLastEmployeeBulkUpload);
employeeUndoAllBulkBtn?.addEventListener("click", undoAllEmployeeBulkUploads);
employeeDeleteBulkSelectedBtn?.addEventListener("click", deleteSelectedEmployeeBulkUpload);
employeeTbody?.addEventListener("click", (event) => {
  const photoThumb = event.target.closest("[data-action='view-employee-photo'][data-id]");
  if (photoThumb) {
    openEmployeePhotoModal(photoThumb.dataset.id);
    return;
  }
  const btn = event.target.closest("button[data-action]");
  if (btn) {
    if (btn.dataset.action === "edit-employee") {
      startEmployeeEdit(btn.dataset.id);
      return;
    }
    if (btn.dataset.action === "edit-employee-photo") {
      startEmployeePhotoEdit(btn.dataset.id);
      return;
    }
    if (btn.dataset.action === "delete-employee") {
      deleteEmployee(btn.dataset.id);
    }
  }
});
employeeTbody?.addEventListener("change", (event) => {
  const box = event.target.closest("input[type='checkbox'][data-action='select-employee'][data-id]");
  if (!box) return;
  if (box.checked) selectedEmployeeIds.add(box.dataset.id);
  else selectedEmployeeIds.delete(box.dataset.id);
});
employeePhotoEditInput?.addEventListener("change", saveEditedEmployeePhoto);
employeePhotoModalClose?.addEventListener("click", closeEmployeePhotoModal);
employeePhotoModal?.addEventListener("click", (event) => {
  if (event.target === employeePhotoModal) closeEmployeePhotoModal();
});
employeeEditForm?.addEventListener("submit", saveEmployeeEditModal);
employeeEditModalClose?.addEventListener("click", closeEmployeeEditModal);
employeeEditModalCancel?.addEventListener("click", closeEmployeeEditModal);
employeeEditModal?.addEventListener("click", (event) => {
  if (event.target === employeeEditModal) closeEmployeeEditModal();
});
employeeEditPhoto?.addEventListener("change", async () => {
  const file = employeeEditPhoto.files?.[0];
  if (!file) return;
  try {
    employeeEditPhotoDataUrl = await readCompressedImage(file);
    if (employeeEditPhotoPreview) {
      employeeEditPhotoPreview.hidden = false;
      employeeEditPhotoPreview.src = employeeEditPhotoDataUrl;
    }
  } catch (error) {
    setStatus(employeeEditStatusMsg, error.message, "err");
  }
});

addDepartmentBtn.addEventListener("click", addDepartment);
assignOwner.addEventListener("input", () => applyEmployeeDataToAssignFields(assignOwner.value));
assignOwner.addEventListener("change", () => applyEmployeeDataToAssignFields(assignOwner.value));
catalogAssignedOwner?.addEventListener("input", () => applyEmployeeDataToCatalogAssignedFields(catalogAssignedOwner.value));
catalogAssignedOwner?.addEventListener("change", () => applyEmployeeDataToCatalogAssignedFields(catalogAssignedOwner.value));
assignAssetSelect.addEventListener("change", () => {
  const asset = readAssignAsset();
  fillAssignForm(asset);
  selectedAssetId = asset?.id || null;
  renderBarcodePreview(asset || null);
  assignAssetSuggestions.classList.add("hidden");
});
assignAssetSelect.addEventListener("input", () => {
  renderAssignSuggestions(assignAssetSelect.value);
  const asset = readAssignAsset();
  if (!asset) return;
  selectedAssetId = asset.id;
  fillAssignForm(asset);
  renderBarcodePreview(asset);
});
assignAssetSelect.addEventListener("focus", () => {
  renderAssignSuggestions(assignAssetSelect.value);
});
assignAssetSelect.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  const asset = findScopedAssetByTag(assignAssetSelect.value);
  if (!asset) return;
  assignAssetSelect.value = `${asset.assetTag} - ${asset.name}`;
  selectedAssetId = asset.id;
  fillAssignForm(asset);
  renderBarcodePreview(asset);
  assignAssetSuggestions.classList.add("hidden");
});
assignAssetSuggestions.addEventListener("click", (event) => {
  const item = event.target.closest(".suggestion-item");
  if (!item) return;
  const id = item.dataset.id;
  const label = item.dataset.label || "";
  const asset = assets.find((a) => a.id === id) || null;
  if (!asset) return;
  assignAssetSelect.value = label;
  selectedAssetId = asset.id;
  fillAssignForm(asset);
  renderBarcodePreview(asset);
  assignAssetSuggestions.classList.add("hidden");
});
returnAssetSelect.addEventListener("change", () => {
  const asset = readReturnAsset();
  fillReturnForm(asset);
  returnAssetSuggestions.classList.add("hidden");
});
returnAssetSelect.addEventListener("input", () => {
  renderReturnSuggestions(returnAssetSelect.value);
  const asset = readReturnAsset();
  if (!asset) return;
  fillReturnForm(asset);
});
returnAssetSelect.addEventListener("focus", () => {
  renderReturnSuggestions(returnAssetSelect.value);
});
returnAssetSelect.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  const asset = findScopedAssetByTag(returnAssetSelect.value);
  if (!asset) return;
  returnAssetSelect.value = `${asset.assetTag} - ${asset.name}`;
  fillReturnForm(asset);
  returnAssetSuggestions.classList.add("hidden");
});
returnAssetSuggestions.addEventListener("click", (event) => {
  const item = event.target.closest(".suggestion-item");
  if (!item) return;
  const id = item.dataset.id;
  const label = item.dataset.label || "";
  const asset = assets.find((a) => a.id === id) || null;
  if (!asset) return;
  returnAssetSelect.value = label;
  fillReturnForm(asset);
  returnAssetSuggestions.classList.add("hidden");
});
document.addEventListener("click", (event) => {
  if (
    event.target === assignAssetSelect ||
    assignAssetSuggestions.contains(event.target) ||
    event.target === returnAssetSelect ||
    returnAssetSuggestions.contains(event.target)
  ) return;
  assignAssetSuggestions.classList.add("hidden");
  returnAssetSuggestions.classList.add("hidden");
});

assignPhoto.addEventListener("change", async () => {
  const file = assignPhoto.files?.[0];
  if (!file) return;
  try {
    assignPhotoDataUrl = await readFileAsDataUrl(file);
    updateAssignPhotoPreview(assignPhotoDataUrl);
  } catch (error) {
    setStatus(assignStatusMsg, error.message, "err");
  }
});

employeePhotoInput?.addEventListener("change", async () => {
  const file = employeePhotoInput.files?.[0];
  if (!file) {
    employeePhotoDataUrl = "";
    if (employeePhotoPreview) {
      employeePhotoPreview.hidden = true;
      employeePhotoPreview.removeAttribute("src");
    }
    return;
  }
  try {
    employeePhotoDataUrl = await readCompressedImage(file);
    if (employeePhotoPreview) {
      employeePhotoPreview.hidden = false;
      employeePhotoPreview.src = employeePhotoDataUrl;
    }
  } catch (error) {
    setStatus(employeeStatusMsg, error.message, "err");
  }
});

assignSaveBtn.addEventListener("click", assignOrUpdateAsset);
deassignBtn.addEventListener("click", deassignAsset);
returnBtn.addEventListener("click", openReturnPageFromTracker);
applyStatusBtn.addEventListener("click", applyStatusChange);
returnSaveBtn.addEventListener("click", returnAssetFromReturnPage);
returnClearBtn.addEventListener("click", clearReturnForm);
returnExportCsvBtn?.addEventListener("click", exportReturnHistoryCsv);

searchInput.addEventListener("input", renderTrackerTable);
filterType.addEventListener("change", renderTrackerTable);
filterStatus.addEventListener("change", renderTrackerTable);

scanFindBtn.addEventListener("click", findByBarcode);
scanInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    findByBarcode();
  }
});

manualBackupBtn.addEventListener("click", () => triggerManualBackup("manual-tracker"));
authBackupBtn?.addEventListener("click", () => triggerManualBackup("manual-auth"));
companyBackupBtn.addEventListener("click", () => triggerManualBackup("manual-company"));
moduleBackupBtn.addEventListener("click", () => triggerManualBackup("manual-module"));
settingsBackupBtn.addEventListener("click", () => triggerManualBackup("manual-settings-header"));
appBackupBtn.addEventListener("click", () => triggerManualBackup("manual-app"));

softwareAssetSearch?.addEventListener("input", () => {
  renderSoftwareAssetOptions();
  renderSoftwareInventoryTable();
});
softwareAssetSearch?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  const asset = findScopedAssetByTag(softwareAssetSearch.value, { deviceOnly: true });
  if (!asset) return;
  selectedSoftwareAssetId = asset.id;
  softwareAssetSearch.value = asset.assetTag;
  renderSoftwareAssetOptions();
  renderSoftwareInventoryTable();
});
softwareAssetSelect?.addEventListener("change", () => {
  selectedSoftwareAssetId = softwareAssetSelect.value || null;
  updateSelectedSoftwareAssetDetails();
  renderSoftwareInventoryTable();
});
softwareNameSelect?.addEventListener("change", showSoftwareSeatUsage);
softwareForm?.addEventListener("submit", saveSoftwareInventory);
softwareSearchInput?.addEventListener("input", renderSoftwareInventoryTable);
softwareViewScope?.addEventListener("change", renderSoftwareInventoryTable);
softwareSortSoftware?.addEventListener("change", renderSoftwareInventoryTable);
softwareSortPirated?.addEventListener("change", renderSoftwareInventoryTable);
softwareDeviceSearch?.addEventListener("input", renderSoftwareDeviceList);
softwareDeviceTbody?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action='software-view-device'][data-id]");
  if (!btn) return;
  openSoftwareDeviceViewModal(btn.dataset.id);
});
softwareDeviceViewClose?.addEventListener("click", closeSoftwareDeviceViewModal);
softwareDeviceViewModal?.addEventListener("click", (event) => {
  if (event.target === softwareDeviceViewModal) closeSoftwareDeviceViewModal();
});
softwareTbody?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action='software-delete'][data-id]");
  if (!btn) return;
  deleteSoftwareInventoryRow(btn.dataset.id);
});
addSoftwareBtn?.addEventListener("click", addSoftwareMasterByAdmin);
deleteSoftwareBtn?.addEventListener("click", deleteSoftwareMasterByAdmin);
softwareInventoryForm?.addEventListener("submit", saveSoftwarePurchase);
inventoryCancelBtn?.addEventListener("click", resetSoftwareInventoryForm);
inventorySearchInput?.addEventListener("input", renderSoftwareInventoryList);
inventoryTbody?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action][data-id]");
  if (!btn) return;
  if (btn.dataset.action === "inventory-edit") {
    editSoftwarePurchase(btn.dataset.id);
    return;
  }
  if (btn.dataset.action === "inventory-delete") {
    deleteSoftwarePurchase(btn.dataset.id);
  }
});
inventoryModalForm?.addEventListener("submit", saveInventoryModalEdit);
inventoryModalCloseBtn?.addEventListener("click", closeInventoryEditModal);
inventoryModalCancelBtn?.addEventListener("click", closeInventoryEditModal);
inventoryEditModal?.addEventListener("click", (event) => {
  if (event.target === inventoryEditModal) closeInventoryEditModal();
});
auditEditForm?.addEventListener("submit", saveAuditEditModal);
auditEditModalClose?.addEventListener("click", closeAuditEditModal);
auditEditModalCancel?.addEventListener("click", closeAuditEditModal);
auditEditModal?.addEventListener("click", (event) => {
  if (event.target === auditEditModal) closeAuditEditModal();
});
activityRefreshBtn?.addEventListener("click", renderActivityHistoryPage);
activitySearchInput?.addEventListener("input", renderActivityHistoryPage);
activityFilterUser?.addEventListener("change", renderActivityHistoryPage);
activityFilterAction?.addEventListener("change", renderActivityHistoryPage);
activityDateFrom?.addEventListener("change", renderActivityHistoryPage);
activityDateTo?.addEventListener("change", renderActivityHistoryPage);
activitySort?.addEventListener("change", renderActivityHistoryPage);
activityExportCsvBtn?.addEventListener("click", exportActivityCsv);
activityExportAllBtn?.addEventListener("click", exportAllData);
clearActivityHistoryBtn?.addEventListener("click", clearAllActivityHistoryByAdmin);

exportCsvBtn.addEventListener("click", exportTrackerCsv);

runAuditBtn.addEventListener("click", runRandomAudit);
openAuditReportBtn.addEventListener("click", openAuditReportPage);
viewAuditReportsBtn.addEventListener("click", () => setActivePage("audit-report"));
viewAuditDueBtn.addEventListener("click", () => setActivePage("audit-status"));
auditTbody.addEventListener("change", (event) => {
  const pick = event.target.closest("input[name='audit-pick'][data-id]");
  if (!pick) return;
  selectedAuditAssetId = pick.dataset.id;
});
auditReportAsset.addEventListener("change", () => {
  selectedAuditAssetId = auditReportAsset.value || null;
  fillAuditReportFormByAsset(selectedAuditAssetId);
});
auditReportHistorySystem.addEventListener("change", () => {
  selectedAuditHistorySystemId = auditReportHistorySystem.value || "";
  renderAuditReportHistory();
});
auditReportSearch?.addEventListener("input", renderAuditReportHistory);
auditReportPageSize?.addEventListener("change", () => {
  setListPageSizeValue("audit-report-history-tbody", auditReportPageSize.value);
  renderAuditReportHistory();
});
refreshAuditHistoryBtn.addEventListener("click", renderAuditReportHistory);
auditReportBackBtn.addEventListener("click", () => {
  editingAuditReportId = null;
  saveAuditBtn.textContent = "Mark Audit Done";
  setActivePage("audit");
});
auditReportForm.addEventListener("submit", saveAuditRecord);
auditReportHistoryTbody.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action][data-id]");
  if (!btn) return;
  if (btn.dataset.action === "delete-audit-history") {
    deleteAuditHistoryByAdmin(btn.dataset.id);
    return;
  }
  if (btn.dataset.action === "edit-audit-history") {
    editAuditHistoryByAdmin(btn.dataset.id);
  }
});
auditLocation.addEventListener("change", runRandomAudit);
refreshAuditStatusBtn.addEventListener("click", renderAuditStatusList);
auditStatusLocation.addEventListener("change", renderAuditStatusList);
auditStatusFilter?.addEventListener("change", renderAuditStatusList);
auditStatusBackBtn.addEventListener("click", () => setActivePage("audit"));
auditReportViewCloseBtn?.addEventListener("click", () => {
  auditReportViewPanel.classList.add("hidden");
  auditReportViewTemplate.innerHTML = "";
});
auditStatusTbody.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action][data-id]");
  if (!btn) return;
  if (btn.dataset.action === "view-audit-report") {
    viewAuditReportTemplate(btn.dataset.id);
    return;
  }
  if (btn.dataset.action === "audit-mark-audited") {
    updateAuditStatusByAdmin(btn.dataset.id, "audited");
    return;
  }
  if (btn.dataset.action === "audit-mark-due") {
    updateAuditStatusByAdmin(btn.dataset.id, "due");
  }
});

refreshCleaningBtn.addEventListener("click", renderCleaningTable);
cleaningLocation.addEventListener("change", renderCleaningTable);
cleaningTbody.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action='mark-cleaned']");
  if (!btn) return;
  markAssetCleaned(btn.dataset.id);
});

tableBody.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-id]");
  if (!btn) return;

  const id = btn.dataset.id;
  const action = btn.dataset.action;

  if (action === "select") {
    const asset = assets.find((a) => a.id === id);
    if (!asset) return;
    selectedAssetId = id;
    renderBarcodePreview(asset);
    if (asset.company === currentCompany && asset.location === currentLocation) {
      assignAssetSelect.value = `${asset.assetTag} - ${asset.name}`;
      fillAssignForm(asset);
    }
    return;
  }

  if (action === "edit") {
    startCatalogEdit(id);
    return;
  }

  if (action === "delete") {
    deleteOneAsset(id);
    return;
  }

  if (action === "show-admin") {
    requirePasswordForAdminReveal(id);
  }
});

printBtn.addEventListener("click", printSelectedBarcode);
adminUserForm.addEventListener("submit", addUserByAdmin);
adminUserTbody.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action]");
  if (!btn) return;
  if (btn.dataset.action === "edit-user") {
    editUserByAdmin(btn.dataset.id);
    return;
  }
  if (btn.dataset.action === "promote-admin") {
    promoteUserToAdmin(btn.dataset.id);
    return;
  }
  if (btn.dataset.action === "delete-user") {
    deleteUserByAdmin(btn.dataset.id);
  }
});
adminUserEditForm?.addEventListener("submit", saveAdminUserEdit);
adminUserEditClose?.addEventListener("click", closeAdminUserEditModal);
adminUserEditCancel?.addEventListener("click", closeAdminUserEditModal);
adminUserEditModal?.addEventListener("click", (event) => {
  if (event.target === adminUserEditModal) closeAdminUserEditModal();
});
adminUserEditPhoto?.addEventListener("change", async () => {
  const file = adminUserEditPhoto.files?.[0];
  if (!file) return;
  try {
    editingAdminUserPhotoDataUrl = await readCompressedImage(file);
    if (adminUserEditPhotoPreview) {
      adminUserEditPhotoPreview.hidden = false;
      adminUserEditPhotoPreview.src = editingAdminUserPhotoDataUrl;
    }
  } catch (error) {
    setStatus(adminUserEditStatus, error.message || "Failed to process user photo.", "err");
  }
});

async function bootstrapApp() {
  loadState();
  await tryAutoRestoreFromSharedSaveFile();
  initDateInputEnhancements();
  initUppercaseEnforcements();
  if (canUseSaveFolderApi()) setSaveFolderLabel("No save folder selected.");
  else setSaveFolderLabel("Save folder not supported in this browser.");
  createDailyBackupSnapshot();
  persistLiveBackupSnapshot("startup");
  setupFieldSuggestions();
  initListPageSizeControls();
  refreshCloudSyncForm();
  if (cloudSyncConfig.autoSync && isCloudSyncConfigured()) {
    await pullCloudState("startup");
  }
  startCloudAutoPull();
  route();
}

void bootstrapApp();
