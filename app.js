const PROFILE_STORAGE_KEY = "batch_query_profiles_v1";
const TASK_STORAGE_KEY = "batch_query_tasks_v1";
const LARGE_ROW_THRESHOLD = 200000;
const EXCEL_MAX_ROWS_PER_SHEET = 1048576;
const BRIDGE_API_BASE = "http://127.0.0.1:8765";
const XLSX_CDN = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";
const DUCKDB_CDN = "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

let XLSX = null;
let duckdb = null;

const state = {
  profiles: [],
  activeProfileId: "",
  fileContexts: [],
  activeFileId: "",
  hasHeader: true,
  tasks: [],
  runningTasks: new Map(),
  progressTicker: null
};

const els = {
  profileTabs: document.getElementById("profileTabs"),
  addProfileBtn: document.getElementById("addProfileBtn"),
  exportProfilesBtn: document.getElementById("exportProfilesBtn"),
  importProfilesInput: document.getElementById("importProfilesInput"),
  profileName: document.getElementById("profileName"),
  profileDbType: document.getElementById("profileDbType"),
  profileHost: document.getElementById("profileHost"),
  profilePort: document.getElementById("profilePort"),
  profileUser: document.getElementById("profileUser"),
  profilePassword: document.getElementById("profilePassword"),
  profileDatabase: document.getElementById("profileDatabase"),
  saveProfileBtn: document.getElementById("saveProfileBtn"),
  testProfileBtn: document.getElementById("testProfileBtn"),
  testProfileStatus: document.getElementById("testProfileStatus"),
  deleteProfileBtn: document.getElementById("deleteProfileBtn"),
  dataFileInput: document.getElementById("dataFileInput"),
  dropZone: document.getElementById("dropZone"),
  fileMeta: document.getElementById("fileMeta"),
  importedFilesList: document.getElementById("importedFilesList"),
  sheetSelect: document.getElementById("sheetSelect"),
  previewTable: document.getElementById("previewTable"),
  selectedColumnMeta: document.getElementById("selectedColumnMeta"),
  addTaskBtn: document.getElementById("addTaskBtn"),
  runAllBtn: document.getElementById("runAllBtn"),
  exportAllXlsxBtn: document.getElementById("exportAllXlsxBtn"),
  globalProgressBar: document.getElementById("globalProgressBar"),
  globalProgressText: document.getElementById("globalProgressText"),
  globalProgressMeta: document.getElementById("globalProgressMeta"),
  taskList: document.getElementById("taskList")
};

boot();

function boot() {
  try {
    loadProfiles();
    loadTasks();
    bindEvents();
    renderProfiles();
    renderImportedFiles();
    renderSheetPreview();
    renderTasks();
    window.__APP_READY = true;
  } catch (error) {
    console.error("Boot failed:", error);
    window.__APP_BOOT_ERROR = String(error?.message || error || "Unknown boot error");
    throw error;
  }
}

function bindEvents() {
  els.addProfileBtn.addEventListener("click", () => {
    const profile = createProfile();
    state.profiles.push(profile);
    state.activeProfileId = profile.id;
    saveProfiles();
    renderProfiles();
    renderTasks();
  });

  els.exportProfilesBtn.addEventListener("click", exportProfilesToFile);
  els.importProfilesInput.addEventListener("change", onImportProfiles);
  els.saveProfileBtn.addEventListener("click", saveActiveProfileFromForm);
  els.testProfileBtn.addEventListener("click", testActiveProfileConnection);
  els.deleteProfileBtn.addEventListener("click", deleteActiveProfile);
  els.profileDbType.addEventListener("change", () => {
    const profile = getActiveProfile();
    if (!profile) return;
    profile.dbType = normalizeDbType(els.profileDbType.value);
    if (!els.profilePort.value.trim()) {
      els.profilePort.value = getDefaultPort(profile.dbType);
    }
    els.profilePort.placeholder = getDefaultPort(profile.dbType);
    saveProfiles();
    renderTasks();
  });

  els.profileTabs.addEventListener("click", (event) => {
    const tab = event.target.closest(".tab");
    if (!tab) return;
    state.activeProfileId = tab.dataset.id || "";
    renderProfiles();
  });

  els.dataFileInput.addEventListener("change", async (event) => {
    const [file] = event.target.files || [];
    if (file) await importDataFile(file);
  });

  ["dragenter", "dragover"].forEach((type) => {
    els.dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      els.dropZone.classList.add("dragging");
    });
  });

  ["dragleave", "drop"].forEach((type) => {
    els.dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      els.dropZone.classList.remove("dragging");
    });
  });

  els.dropZone.addEventListener("drop", async (event) => {
    const [file] = event.dataTransfer?.files || [];
    if (file) await importDataFile(file);
  });

  document.querySelectorAll("input[name='hasHeader']").forEach((input) => {
    input.addEventListener("change", () => {
      const fileCtx = getActiveFileContext();
      if (!fileCtx) return;
      fileCtx.hasHeader = input.value === "true";
      state.hasHeader = fileCtx.hasHeader;
      renderSheetPreview();
      renderImportedFiles();
    });
  });

  els.sheetSelect.addEventListener("change", () => {
    const fileCtx = getActiveFileContext();
    if (fileCtx) {
      fileCtx.activeSheetIndex = Number(els.sheetSelect.value || 0);
    }
    renderSheetPreview();
  });

  els.previewTable.addEventListener("click", (event) => {
    const th = event.target.closest("th[data-col-index]");
    if (!th) return;
    const fileCtx = getActiveFileContext();
    if (!fileCtx) return;
    fileCtx.selectedColumnIndex = Number(th.dataset.colIndex);
    renderSheetPreview();
    renderImportedFiles();
  });

  els.importedFilesList.addEventListener("click", (event) => {
    const delBtn = event.target.closest("[data-file-del-id]");
    if (delBtn) {
      event.preventDefault();
      event.stopPropagation();
      removeImportedFile(delBtn.dataset.fileDelId || "");
      return;
    }
    const row = event.target.closest("[data-file-id]");
    if (!row) return;
    setActiveFile(row.dataset.fileId || "");
  });

  els.addTaskBtn.addEventListener("click", () => {
    state.tasks.push(createTask());
    saveTasks();
    renderTasks();
  });

  els.runAllBtn.addEventListener("click", runAllTasksInParallel);
  els.exportAllXlsxBtn.addEventListener("click", exportAllTasksAsXlsx);
  els.taskList.addEventListener("click", onTaskListClick);
  els.taskList.addEventListener("input", onTaskListInput);
}

function loadProfiles() {
  try {
    const raw = localStorage.getItem(PROFILE_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    state.profiles = Array.isArray(parsed) ? parsed.map((item) => normalizeProfile(item)) : [];
  } catch (error) {
    state.profiles = [];
  }

  if (state.profiles.length === 0) {
    state.profiles.push(createProfile());
  }
  state.activeProfileId = state.profiles[0].id;
}

function saveProfiles() {
  localStorage.setItem(PROFILE_STORAGE_KEY, JSON.stringify(state.profiles));
}

function createProfile(overrides = {}) {
  const profile = {
    id: generateId(),
    name: "默认配置",
    dbType: "mysql",
    host: "",
    port: "3306",
    username: "",
    password: "",
    database: "",
    ...overrides
  };
  profile.dbType = normalizeDbType(profile.dbType);
  profile.port = String(profile.port || getDefaultPort(profile.dbType));
  return profile;
}

function getActiveProfile() {
  return state.profiles.find((p) => p.id === state.activeProfileId) || null;
}

function renderProfiles() {
  els.profileTabs.innerHTML = state.profiles
    .map((profile) => {
      const active = profile.id === state.activeProfileId ? "active" : "";
      return `<button type="button" class="tab ${active}" data-id="${profile.id}">${escapeHtml(profile.name || "未命名")}</button>`;
    })
    .join("");

  const profile = getActiveProfile();
  if (!profile) return;

  els.profileName.value = profile.name || "";
  els.profileDbType.value = normalizeDbType(profile.dbType);
  els.profileHost.value = profile.host || "";
  els.profilePort.value = profile.port || "";
  els.profilePort.placeholder = getDefaultPort(els.profileDbType.value);
  els.profileUser.value = profile.username || "";
  els.profilePassword.value = profile.password || "";
  els.profileDatabase.value = profile.database || "";
  setProfileTestStatus("未测试", "");
}

function saveActiveProfileFromForm() {
  const profile = getActiveProfile();
  if (!profile) return;
  profile.name = els.profileName.value.trim() || "未命名配置";
  profile.dbType = normalizeDbType(els.profileDbType.value);
  profile.host = els.profileHost.value.trim();
  profile.port = els.profilePort.value.trim();
  profile.username = els.profileUser.value.trim();
  profile.password = els.profilePassword.value;
  profile.database = els.profileDatabase.value.trim();
  saveProfiles();
  renderProfiles();
  renderTasks();
}

async function testActiveProfileConnection() {
  const profile = getActiveProfile();
  if (!profile) return;
  saveActiveProfileFromForm();

  setProfileTestStatus("测试中...", "");
  try {
    const resp = await postBridgeJSON("/api/test-connection", { profile });
    if (!resp?.ok) {
      throw new Error(resp?.message || "连接失败");
    }
    setProfileTestStatus("连接成功", "ok");
  } catch (error) {
    console.error(error);
    const msg = String(error?.message || "连接失败");
    const lowerMsg = msg.toLowerCase();
    const isBridgeUnavailable =
      msg.includes("无法连接本地桥接服务") ||
      msg.includes("python src/bridge_server.py") ||
      lowerMsg.includes("failed to fetch") ||
      lowerMsg.includes("networkerror");
    const isAccessDenied = lowerMsg.includes("access denied") || msg.includes("1045");

    let finalMsg = msg;
    if (isBridgeUnavailable) {
      finalMsg = msg.includes("python src/bridge_server.py")
        ? msg
        : `${msg}，请先启动 python src/bridge_server.py`;
    } else if (isAccessDenied) {
      finalMsg = `${msg}。请检查用户名/密码，以及数据库侧是否放通你的来源 IP。`;
    }
    setProfileTestStatus(`连接失败: ${finalMsg}`, "err");
  }
}

function setProfileTestStatus(text, statusClass) {
  if (!els.testProfileStatus) return;
  els.testProfileStatus.textContent = text;
  els.testProfileStatus.classList.remove("ok", "err");
  if (statusClass) {
    els.testProfileStatus.classList.add(statusClass);
  }
}

function deleteActiveProfile() {
  if (state.profiles.length <= 1) {
    alert("至少保留一个配置。");
    return;
  }
  state.profiles = state.profiles.filter((p) => p.id !== state.activeProfileId);
  state.activeProfileId = state.profiles[0].id;
  state.tasks.forEach((task) => {
    if (!state.profiles.some((p) => p.id === task.profileId)) {
      task.profileId = state.profiles[0].id;
    }
  });
  saveProfiles();
  saveTasks();
  renderProfiles();
  renderTasks();
}

function exportProfilesToFile() {
  downloadTextFile(
    `db_profiles_${formatNow()}.json`,
    JSON.stringify(state.profiles, null, 2),
    "application/json;charset=utf-8"
  );
}

async function onImportProfiles(event) {
  const files = Array.from(event.target.files || []);
  if (files.length === 0) return;

  const imported = [];
  for (const file of files) {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        parsed.forEach((item) => imported.push(normalizeProfile(item)));
      } else {
        imported.push(normalizeProfile(parsed));
      }
    } catch (error) {
      console.error(error);
    }
  }

  if (imported.length === 0) {
    alert("没有读取到有效的配置 JSON。");
    return;
  }

  state.profiles.push(...imported);
  state.activeProfileId = imported[0].id;
  saveProfiles();
  renderProfiles();
  renderTasks();
  event.target.value = "";
}

function normalizeProfile(raw) {
  const dbType = normalizeDbType(
    raw?.dbType || raw?.db_type || raw?.type || raw?.dialect || raw?.DIALECT
  );
  return createProfile({
    name: raw?.name || raw?.database || "导入配置",
    dbType,
    host: raw?.host || raw?.HOST || "",
    port: String(raw?.port || raw?.PORT || getDefaultPort(dbType)),
    username: raw?.username || raw?.USERNAME || "",
    password: raw?.password || raw?.PASSWORD || "",
    database: raw?.database || raw?.DATABASE || ""
  });
}

function normalizeDbType(value) {
  const text = String(value || "").trim().toLowerCase();
  if (text === "clickhouse" || text === "ch" || text.includes("clickhouse")) return "clickhouse";
  return "mysql";
}

function getDefaultPort(dbType) {
  return normalizeDbType(dbType) === "clickhouse" ? "8123" : "3306";
}

function loadTasks() {
  try {
    const raw = localStorage.getItem(TASK_STORAGE_KEY);
    state.tasks = raw ? JSON.parse(raw) : [];
  } catch (error) {
    state.tasks = [];
  }
}

function saveTasks() {
  localStorage.setItem(TASK_STORAGE_KEY, JSON.stringify(state.tasks));
}

function createTask(overrides = {}) {
  return {
    id: generateId(),
    name: `任务 ${state.tasks.length + 1}`,
    profileId: state.activeProfileId,
    fileId: state.activeFileId || "",
    sql: "SELECT * FROM source_data WHERE 1=1 AND id IN ({})",
    batchSize: 3000,
    status: "idle",
    progress: 0,
    message: "待执行",
    totalBatches: 0,
    doneBatches: 0,
    totalRows: 0,
    doneRows: 0,
    startedAt: null,
    endedAt: null,
    exportType: "csv",
    csvEncoding: "auto",
    resultRows: [],
    resultColumns: [],
    ...overrides
  };
}

function renderTasks() {
  if (state.tasks.length === 0) {
    els.taskList.innerHTML = '<div class="task-card">暂无任务，点击“新增任务”。</div>';
    return;
  }

  const options = state.profiles
    .map((profile) => `<option value="${profile.id}">${escapeHtml(profile.name)}</option>`)
    .join("");
  const fileOptions = state.fileContexts.length > 0
    ? state.fileContexts.map((file) => `<option value="${file.id}">${escapeHtml(file.fileName)}</option>`).join("")
    : '<option value="">未导入文件</option>';

  els.taskList.innerHTML = state.tasks
    .map((task) => {
      const isExporting = task.status === "exporting";
      const exportBtnLabel = isExporting ? "导出中..." : "导出";
      const exportBtnDisabled = isExporting ? "disabled" : "";
      const statusClass =
        task.status === "done"
          ? "ok"
          : task.status === "error"
            ? "err"
            : task.status === "aborted"
              ? "warn"
              : task.status === "exporting"
                ? "warn"
              : "";
      return `
        <article class="task-card" data-task-id="${task.id}">
          <div class="task-head">
            <strong>${escapeHtml(task.name)}</strong>
            <div class="head-actions">
              <button type="button" data-act="run" data-task-id="${task.id}">运行</button>
              <button type="button" data-act="abort" data-task-id="${task.id}">中止</button>
              <button type="button" data-act="export" data-task-id="${task.id}" ${exportBtnDisabled}>${exportBtnLabel}</button>
              <button type="button" class="danger" data-act="delete" data-task-id="${task.id}">删除</button>
            </div>
          </div>
          <div class="task-grid">
            <label>任务名<input data-field="name" value="${escapeAttr(task.name)}"></label>
            <label>配置
              <select data-field="profileId">${options}</select>
            </label>
            <label>文件
              <select data-field="fileId">${fileOptions}</select>
            </label>
            <label>批大小<input data-field="batchSize" type="number" min="1" value="${task.batchSize}"></label>
            <label>导出格式
              <select data-field="exportType">
                <option value="csv">CSV</option>
                <option value="xlsx">XLSX(多Sheet)</option>
              </select>
            </label>
            <label>CSV编码
              <select data-field="csvEncoding">
                <option value="auto">自动(推荐)</option>
                <option value="utf8">UTF-8</option>
                <option value="utf8bom">UTF-8 BOM(Win兼容)</option>
              </select>
            </label>
            <label class="sql">SQL(支持 {} 占位符或 {{values}} 占位符)
              <textarea data-field="sql">${escapeHtml(task.sql)}</textarea>
            </label>
          </div>
          <progress max="100" value="${Number(task.progress) || 0}"></progress>
          <div class="status ${statusClass}">${escapeHtml(task.message || "")}</div>
          <div class="status">${escapeHtml(buildTaskProgressMeta(task))}</div>
        </article>
      `;
    })
    .join("");

  state.tasks.forEach((task) => {
    const card = els.taskList.querySelector(`[data-task-id="${task.id}"]`);
    if (!card) return;
    const profileSelect = card.querySelector('[data-field="profileId"]');
    const fileSelect = card.querySelector('[data-field="fileId"]');
    const exportSelect = card.querySelector('[data-field="exportType"]');
    const csvEncodingSelect = card.querySelector('[data-field="csvEncoding"]');
    if (profileSelect) profileSelect.value = task.profileId;
    if (fileSelect) fileSelect.value = task.fileId || "";
    if (exportSelect) exportSelect.value = task.exportType;
    if (csvEncodingSelect) csvEncodingSelect.value = task.csvEncoding || "auto";
  });

  renderGlobalProgress();
  syncProgressTicker();
}

function onTaskListInput(event) {
  const input = event.target;
  const card = input.closest("[data-task-id]");
  if (!card) return;
  const task = state.tasks.find((item) => item.id === card.dataset.taskId);
  if (!task) return;
  const field = input.dataset.field;
  if (!field) return;

  if (field === "batchSize") {
    task.batchSize = Math.max(1, Number(input.value) || 3000);
  } else {
    task[field] = input.value;
  }
  saveTasks();
}

function buildTaskProgressMeta(task) {
  const totalBatches = Number(task.totalBatches) || 0;
  const doneBatches = Number(task.doneBatches) || 0;
  const totalRows = Number(task.totalRows) || 0;
  const doneRows = Number(task.doneRows) || 0;
  const start = Number(task.startedAt) || 0;
  const end = Number(task.endedAt) || 0;
  if (!start) return "耗时: 00:00:00 | 批次: 0/0 | 行数: 0/0 | ETA: --";

  const now = task.status === "running" ? Date.now() : end || Date.now();
  const elapsedMs = Math.max(0, now - start);
  const etaText = task.status === "running" ? formatEta(elapsedMs, doneRows, totalRows) : "--";
  return `耗时: ${formatDuration(elapsedMs)} | 批次: ${doneBatches}/${totalBatches} | 行数: ${doneRows}/${totalRows} | ETA: ${etaText}`;
}

function renderGlobalProgress() {
  if (!els.globalProgressBar || !els.globalProgressText || !els.globalProgressMeta) return;

  const targetTasks = state.tasks.filter(
    (task) => Number(task.totalRows) > 0 || Number(task.totalBatches) > 0 || Number(task.progress) > 0
  );
  if (targetTasks.length === 0) {
    els.globalProgressBar.value = 0;
    els.globalProgressText.textContent = "0%";
    els.globalProgressMeta.textContent = "等待任务启动";
    return;
  }

  const totalRows = targetTasks.reduce((sum, task) => sum + (Number(task.totalRows) || 0), 0);
  const doneRows = targetTasks.reduce((sum, task) => sum + (Number(task.doneRows) || 0), 0);
  const totalBatches = targetTasks.reduce((sum, task) => sum + (Number(task.totalBatches) || 0), 0);
  const doneBatches = targetTasks.reduce((sum, task) => sum + (Number(task.doneBatches) || 0), 0);
  const runningCount = targetTasks.filter((task) => task.status === "running").length;

  const progressPercent = totalRows > 0 ? Math.round((doneRows / totalRows) * 100) : 0;
  els.globalProgressBar.value = Math.min(100, Math.max(0, progressPercent));
  els.globalProgressText.textContent = `${progressPercent}%`;

  const startedTasks = targetTasks.filter((task) => Number(task.startedAt) > 0);
  const earliestStart = startedTasks.length > 0
    ? Math.min(...startedTasks.map((task) => Number(task.startedAt)))
    : 0;
  const elapsedMs = earliestStart ? Date.now() - earliestStart : 0;
  const etaText = runningCount > 0 ? formatEta(elapsedMs, doneRows, totalRows) : "--";

  els.globalProgressMeta.textContent =
    `运行中任务: ${runningCount} | 耗时: ${formatDuration(elapsedMs)} | ETA: ${etaText} | 批次: ${doneBatches}/${totalBatches} | 行数: ${doneRows}/${totalRows}`;
}

function syncProgressTicker() {
  const hasRunning = state.tasks.some((task) => task.status === "running");
  if (hasRunning && !state.progressTicker) {
    state.progressTicker = setInterval(() => {
      renderGlobalProgress();
      const runningCards = state.tasks.filter((task) => task.status === "running");
      if (runningCards.length > 0) {
        const statusNodes = document.querySelectorAll(".task-card");
        statusNodes.forEach((node) => {
          const taskId = node.getAttribute("data-task-id");
          const task = state.tasks.find((item) => item.id === taskId);
          if (!task) return;
          const infoNodes = node.querySelectorAll(".status");
          if (infoNodes.length >= 2) {
            infoNodes[1].textContent = buildTaskProgressMeta(task);
          }
        });
      }
    }, 1000);
  }
  if (!hasRunning && state.progressTicker) {
    clearInterval(state.progressTicker);
    state.progressTicker = null;
  }
}

async function onTaskListClick(event) {
  const btn = event.target.closest("button[data-act]");
  if (!btn) return;
  event.preventDefault();
  event.stopPropagation();

  const taskId = btn.dataset.taskId || btn.closest("[data-task-id]")?.dataset.taskId || "";
  if (!taskId) return;
  const task = state.tasks.find((item) => item.id === taskId);
  if (!task) return;

  const act = btn.dataset.act;
  if (act === "delete") {
    if (state.runningTasks.has(task.id)) {
      await abortTask(task.id);
    }
    state.tasks = state.tasks.filter((item) => item.id !== task.id);
    saveTasks();
    renderTasks();
    return;
  }

  if (act === "run") {
    await runTask(task.id);
    return;
  }

  if (act === "abort") {
    await abortTask(task.id);
    return;
  }

  if (act === "export") {
    if (task.status === "exporting") {
      showToast("该任务正在导出，请稍候。", "warn");
      return;
    }
    await exportSingleTask(task);
  }
}

async function importDataFile(file) {
  const name = file.name.toLowerCase();
  const isCsv = name.endsWith(".csv");
  const isExcel = name.endsWith(".xlsx") || name.endsWith(".xls");
  if (!(isCsv || isExcel)) {
    alert("仅支持 .csv / .xlsx / .xls");
    return;
  }

  try {
    let sheets = [];
    if (isCsv) {
      const baseName = file.name.replace(/\.[^./\\]+$/, "");
      const displayName = baseName || file.name || "Sheet1";
      const viewName = normalizeViewName(displayName);
      const fileToken = generateId().replace(/-/g, "");
      const virtualFileName = `${viewName}_${fileToken}.csv`;
      const csvBytes = new Uint8Array(await file.arrayBuffer());
      const analyzed = await analyzeCsvWithDuckDB(virtualFileName, csvBytes.slice(), 11);
      sheets = [
        {
          name: "Sheet1",
          viewName,
          rowCount: analyzed.rowCount,
          previewRows: analyzed.previewRows,
          csvBytes
        }
      ];
    } else {
      await ensureXLSX();
      const workbook = await readWorkbook(file);
      sheets = workbook.SheetNames.map((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const ref = sheet["!ref"] || "A1:A1";
        const range = XLSX.utils.decode_range(ref);
        const rowCount = range.e.r - range.s.r + 1;
        const csvText = XLSX.utils.sheet_to_csv(sheet, { FS: ",", RS: "\n" });
        const previewRows = getPreviewRowsFromCsv(csvText, 11);
        return {
          name: sheetName,
          viewName: normalizeViewName(sheetName),
          rowCount,
          csvText,
          previewRows
        };
      });
    }

    const totalRows = sheets.reduce((sum, sheet) => sum + sheet.rowCount, 0);
    const fileContext = {
      id: generateId(),
      fileName: file.name,
      sheets,
      totalRows,
      isLarge: totalRows > LARGE_ROW_THRESHOLD,
      selectedColumnIndex: null,
      hasHeader: true,
      activeSheetIndex: 0
    };
    state.fileContexts.push(fileContext);
    setActiveFile(fileContext.id);
    state.tasks.forEach((task) => {
      if (!task.fileId) task.fileId = fileContext.id;
    });
    renderImportedFiles();
    renderTasks();
    renderSheetPreview();
  } catch (error) {
    console.error(error);
    alert(`文件读取失败: ${safeErrorMessage(error, "未知错误")}`);
  }
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array" });
}

function getPreviewRowsFromCsv(csvText, rowLimit) {
  return parseCsvPreviewAndCount(csvText, rowLimit).previewRows;
}

async function analyzeCsvWithDuckDB(fileName, csvBytes, rowLimit = 11) {
  const db = await createDuckDBInstance();
  const conn = await db.connect();
  try {
    await registerCsvInput(db, fileName, { csvBytes });
    const escapedFile = sqlQuote(fileName);
    const safeLimit = Math.max(1, Number(rowLimit) || 11);
    const previewTable = await conn.query(
      `SELECT * FROM read_csv_auto('${escapedFile}', header=false, all_varchar=true) LIMIT ${safeLimit}`
    );
    const countTable = await conn.query(
      `SELECT COUNT(*) AS cnt FROM read_csv_auto('${escapedFile}', header=false, all_varchar=true)`
    );
    const previewRows = arrowTableToRows(previewTable);
    const countRows = arrowTableToObjects(countTable);
    const rowCount = Number(countRows[0]?.cnt ?? 0) || 0;
    return { previewRows, rowCount };
  } finally {
    try {
      await conn.close();
    } catch (_) {}
    try {
      await db.terminate();
    } catch (_) {}
  }
}

function parseCsvPreviewAndCount(csvText, rowLimit = 11) {
  const previewRows = [];
  let rowCount = 0;
  let row = [];
  let field = "";
  let inQuotes = false;

  function finishField() {
    row.push(field);
    field = "";
  }

  function finishRow() {
    finishField();
    rowCount += 1;
    if (previewRows.length < rowLimit) {
      previewRows.push(row);
    }
    row = [];
  }

  for (let i = 0; i < csvText.length; i += 1) {
    const ch = csvText[i];
    if (ch === '"') {
      if (inQuotes && csvText[i + 1] === '"') {
        field += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && ch === ",") {
      finishField();
      continue;
    }

    if (!inQuotes && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && csvText[i + 1] === "\n") {
        i += 1;
      }
      finishRow();
      continue;
    }

    field += ch;
  }

  if (field.length > 0 || row.length > 0) {
    finishRow();
  }

  return { previewRows, rowCount };
}

function safeErrorMessage(error, fallback = "未知错误") {
  try {
    const msg = error?.message;
    if (typeof msg === "string" && msg.trim()) return msg.trim();
  } catch (_) {}
  try {
    const raw = String(error ?? "").trim();
    if (raw && raw !== "[object Object]") return raw;
  } catch (_) {}
  return fallback;
}

function getCurrentSheet() {
  const fileCtx = getActiveFileContext();
  if (!fileCtx || fileCtx.sheets.length === 0) return null;
  const idx = Number(fileCtx.activeSheetIndex || 0);
  return fileCtx.sheets[idx] || fileCtx.sheets[0];
}

function renderSheetPreview() {
  const fileCtx = getActiveFileContext();
  if (!fileCtx) {
    els.fileMeta.textContent = "未导入文件";
    els.sheetSelect.innerHTML = "";
    els.previewTable.innerHTML = "";
    els.selectedColumnMeta.textContent = "未选择列";
    return;
  }
  state.hasHeader = fileCtx.hasHeader;
  document.querySelectorAll("input[name='hasHeader']").forEach((input) => {
    input.checked = input.value === String(fileCtx.hasHeader);
  });
  els.fileMeta.textContent = `${fileCtx.fileName} | Sheet: ${fileCtx.sheets.length} | 总行数: ${fileCtx.totalRows}${fileCtx.isLarge ? " | 大文件: 使用 DuckDB 处理" : ""}`;
  els.sheetSelect.innerHTML = fileCtx.sheets
    .map((sheet, idx) => `<option value="${idx}">${escapeHtml(sheet.name)} (${sheet.rowCount} 行)</option>`)
    .join("");
  els.sheetSelect.value = String(fileCtx.activeSheetIndex || 0);

  const sheet = getCurrentSheet();
  if (!sheet) {
    els.previewTable.innerHTML = "";
    els.selectedColumnMeta.textContent = "未选择列";
    return;
  }

  const rows = sheet.previewRows || [];
  const headerRow = rows[0] || [];
  const bodyRows = fileCtx.hasHeader ? rows.slice(1, 11) : rows.slice(0, 10);
  const colCount = Math.max(headerRow.length, ...bodyRows.map((row) => row.length), 1);

  const headers = [];
  for (let i = 0; i < colCount; i += 1) {
    if (fileCtx.hasHeader) {
      headers.push(String(headerRow[i] ?? `col_${i + 1}`));
    } else {
      headers.push(`col_${i + 1}`);
    }
  }

  const headerHtml = headers
    .map((name, idx) => {
      const selected = idx === fileCtx.selectedColumnIndex ? "selected" : "";
      return `<th data-col-index="${idx}" class="${selected}" title="点击选择此列">${escapeHtml(name)}</th>`;
    })
    .join("");

  const bodyHtml = bodyRows
    .map((row) => {
      const cells = Array.from({ length: colCount }, (_, idx) => `<td>${escapeHtml(String(row[idx] ?? ""))}</td>`).join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  els.previewTable.innerHTML = `
    <thead><tr>${headerHtml}</tr></thead>
    <tbody>${bodyHtml || `<tr><td colspan="${colCount}">暂无数据</td></tr>`}</tbody>
  `;

  if (fileCtx.selectedColumnIndex === null) {
    els.selectedColumnMeta.textContent = "未选择列";
  } else {
    els.selectedColumnMeta.textContent = `已选择第 ${fileCtx.selectedColumnIndex + 1} 列`;
  }
}

function getActiveFileContext() {
  return state.fileContexts.find((item) => item.id === state.activeFileId) || null;
}

function setActiveFile(fileId) {
  state.activeFileId = fileId;
  const ctx = getActiveFileContext();
  if (ctx) {
    state.hasHeader = ctx.hasHeader;
  }
  renderImportedFiles();
  renderSheetPreview();
}

function renderImportedFiles() {
  if (!els.importedFilesList) return;
  if (state.fileContexts.length === 0) {
    els.importedFilesList.innerHTML = '<div class="imported-file">暂无导入文件</div>';
    return;
  }
  els.importedFilesList.innerHTML = state.fileContexts
    .map((file) => {
      const active = file.id === state.activeFileId ? "active" : "";
      const colText = file.selectedColumnIndex === null ? "未选列" : `第 ${file.selectedColumnIndex + 1} 列`;
      return `<div class="imported-file ${active}" data-file-id="${file.id}">
        <div class="imported-file-head">
          <div>${escapeHtml(file.fileName)}</div>
          <button type="button" class="file-del-btn" data-file-del-id="${file.id}">删除</button>
        </div>
        <div>列: ${colText}</div>
      </div>`;
    })
    .join("");
}

function removeImportedFile(fileId) {
  if (!fileId) return;
  const target = state.fileContexts.find((f) => f.id === fileId);
  if (!target) return;
  const usedTaskCount = state.tasks.filter((task) => task.fileId === fileId).length;
  const ok = window.confirm(
    usedTaskCount > 0
      ? `确认删除文件「${target.fileName}」？将同时清空 ${usedTaskCount} 个任务的文件绑定。`
      : `确认删除文件「${target.fileName}」？`
  );
  if (!ok) return;

  state.fileContexts = state.fileContexts.filter((f) => f.id !== fileId);
  state.tasks.forEach((task) => {
    if (task.fileId === fileId) {
      task.fileId = "";
    }
  });

  if (state.activeFileId === fileId) {
    state.activeFileId = state.fileContexts[0]?.id || "";
  }

  renderImportedFiles();
  renderSheetPreview();
  renderTasks();
  saveTasks();
}

async function runAllTasksInParallel() {
  if (state.tasks.length === 0) {
    alert("请先新增任务。");
    return;
  }
  await Promise.allSettled(state.tasks.map((task) => runTask(task.id)));
}

async function runTask(taskId) {
  const task = state.tasks.find((item) => item.id === taskId);
  if (!task) return;
  if (state.runningTasks.has(taskId)) return;
  const fileCtx = state.fileContexts.find((file) => file.id === task.fileId);
  if (!fileCtx) {
    task.status = "error";
    task.message = "失败: 任务未绑定有效导入文件";
    renderTasks();
    saveTasks();
    return;
  }
  if (fileCtx.selectedColumnIndex === null) {
    task.status = "error";
    task.message = "失败: 该文件尚未选择批处理列";
    renderTasks();
    saveTasks();
    return;
  }

  task.status = "running";
  task.progress = 0;
  task.resultRows = [];
  task.resultColumns = [];
  task.totalBatches = 0;
  task.doneBatches = 0;
  task.totalRows = 0;
  task.doneRows = 0;
  task.startedAt = Date.now();
  task.endedAt = null;
  task.message = "准备批处理...";
  const execution = {
    cancelled: false,
    db: null,
    conn: null,
    fetchControllers: []
  };
  state.runningTasks.set(taskId, execution);
  renderTasks();

  try {
    const values = await extractDistinctValues(fileCtx, fileCtx.selectedColumnIndex, fileCtx.hasHeader, execution);
    if (execution.cancelled) {
      throw new TaskAbortError();
    }
    if (values.length === 0) {
      task.status = "done";
      task.progress = 100;
      task.endedAt = Date.now();
      task.message = "查询列没有可用值。";
      renderTasks();
      saveTasks();
      return;
    }

    const chunks = splitIntoChunks(values, Math.max(1, Number(task.batchSize) || 3000));
    task.totalBatches = chunks.length;
    task.totalRows = values.length;
    task.doneBatches = 0;
    task.doneRows = 0;
    task.message = `执行中: 0/${chunks.length} 批，累计 0 行`;
    renderTasks();
    let mergedRows = [];
    let columns = [];
    const profile = state.profiles.find((item) => item.id === task.profileId) || getActiveProfile();
    if (!profile) {
      throw new Error("未找到任务对应的数据库配置");
    }

    for (let i = 0; i < chunks.length; i += 1) {
      if (execution.cancelled) {
        throw new TaskAbortError();
      }
      const chunk = chunks[i];
      const controller = new AbortController();
      execution.fetchControllers.push(controller);
      let resp;
      try {
        resp = await postBridgeJSON(
          "/api/query-batch",
          {
            profile,
            sql_template: task.sql,
            values: chunk
          },
          controller.signal
        );
      } finally {
        execution.fetchControllers = execution.fetchControllers.filter((c) => c !== controller);
      }
      if (!resp?.ok) {
        throw new Error(resp?.message || "批次查询失败");
      }
      const rows = Array.isArray(resp.rows) ? resp.rows : [];

      if (rows.length > 0 && columns.length === 0) {
        columns = Object.keys(rows[0]);
      }
      mergedRows = mergedRows.concat(rows);
      task.doneBatches = i + 1;
      task.doneRows += chunk.length;
      task.progress = Math.round(((i + 1) / chunks.length) * 100);
      task.message = `执行中: ${i + 1}/${chunks.length} 批，结果 ${mergedRows.length} 行`;
      renderTasks();
    }

    task.status = "done";
    task.progress = 100;
    task.endedAt = Date.now();
    task.resultRows = mergedRows;
    task.resultColumns = columns;
    task.message = `完成: ${mergedRows.length} 行`;
  } catch (error) {
    if (isAbortError(error) || execution.cancelled) {
      task.status = "aborted";
      task.endedAt = Date.now();
      task.message = "任务已中止";
    } else {
      console.error(error);
      task.status = "error";
      task.endedAt = Date.now();
      task.message = `失败: ${error.message}`;
    }
  } finally {
    await closeExecution(execution);
    state.runningTasks.delete(taskId);
  }

  renderTasks();
  saveTasks();
}

async function abortTask(taskId) {
  const execution = state.runningTasks.get(taskId);
  const task = state.tasks.find((item) => item.id === taskId);
  if (!task) return;
  if (!execution) {
    task.message = "当前任务未在运行";
    renderTasks();
    return;
  }

  execution.cancelled = true;
  task.status = "aborted";
  task.endedAt = Date.now();
  task.message = "正在中止...";
  task.progress = Math.min(100, Number(task.progress) || 0);
  renderTasks();
  saveTasks();
  await closeExecution(execution);
}

async function extractDistinctValues(fileCtx, colIndex, hasHeader, execution) {
  const db = await createDuckDBInstance();
  const conn = await db.connect();
  if (execution) {
    execution.db = db;
    execution.conn = conn;
  }
  const values = new Set();

  for (const sheet of fileCtx.sheets) {
    if (execution?.cancelled) {
      throw new TaskAbortError();
    }
    const fileName = `${sheet.viewName}.csv`;
    await registerCsvInput(db, fileName, sheet);
    const headerFlag = hasHeader ? "true" : "false";
    const escapedFile = sqlQuote(fileName);
    const schemaTable = await conn.query(
      `SELECT * FROM read_csv_auto('${escapedFile}', header=${headerFlag}, all_varchar=true) LIMIT 0`
    );
    const schemaNames = schemaTable.schema.fields.map((field) => field.name);
    const colName = schemaNames[colIndex];
    if (!colName) continue;

    const escapedCol = quoteIdentifier(colName);
    const table = await conn.query(
      `SELECT DISTINCT ${escapedCol} AS val
       FROM read_csv_auto('${escapedFile}', header=${headerFlag}, all_varchar=true)
       WHERE ${escapedCol} IS NOT NULL AND TRIM(CAST(${escapedCol} AS VARCHAR)) <> ''`
    );
    const rows = arrowTableToObjects(table);
    rows.forEach((row) => {
      const value = String(row.val ?? "").trim();
      if (value) values.add(value);
    });
  }

  await conn.close();
  await db.terminate();
  if (execution) {
    execution.conn = null;
    execution.db = null;
  }
  return Array.from(values);
}

function splitIntoChunks(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

async function exportSingleTask(task) {
  if (!task.resultRows || task.resultRows.length === 0) {
    alert("该任务暂无结果可导出。");
    return;
  }
  if (task.status === "exporting") return;
  if (task.status === "running") {
    alert("任务运行中，暂不支持导出。请等待运行完成后再导出。");
    return;
  }

  const exportRows = normalizeExportRows(task.resultRows);
  const totalRows = exportRows.length;
  if (totalRows === 0) {
    alert("该任务结果格式异常，无法导出。");
    return;
  }
  const previousStatus = task.status;
  const previousMessage = task.message;
  const previousProgress = Number(task.progress) || 0;
  task.status = "exporting";
  task.progress = 0;
  task.message = `导出中: 0/${totalRows} 行`;
  renderTasks();

  const onProgress = (percent, message) => {
    task.progress = Math.max(0, Math.min(100, Number(percent) || 0));
    task.message = message || `导出中: ${Math.round(task.progress)}%`;
    renderTasks();
  };

  const baseName = `${safeFileName(task.name || "task")}_${formatNow()}`;
  try {
    if (task.exportType === "xlsx") {
      try {
        await exportRowsAsXlsx(exportRows, baseName, onProgress);
      } catch (xlsxError) {
        const xlsxMsg = safeErrorMessage(xlsxError, "");
        if (/Invalid array length/i.test(xlsxMsg)) {
          showToast("XLSX 导出失败，自动降级为 CSV 导出。", "warn", 3200);
          await exportRowsAsCsv(exportRows, baseName, task.csvEncoding || "auto", onProgress);
        } else {
          throw xlsxError;
        }
      }
    } else {
      await exportRowsAsCsv(exportRows, baseName, task.csvEncoding || "auto", onProgress);
    }
    task.status = "done";
    task.progress = 100;
    task.message = `导出完成: ${totalRows} 行`;
    showToast(task.message, "ok");
  } catch (error) {
    console.error(error);
    const errMsg = safeErrorMessage(error, "未知错误");
    if (errMsg.includes("已取消导出")) {
      task.status = previousStatus || "done";
      task.progress = previousProgress;
      task.message = previousMessage || "已取消导出";
      showToast("已取消导出", "warn");
    } else {
      task.status = "error";
      task.message = `导出失败: ${errMsg}`;
      alert(task.message);
      showToast(task.message, "err");
    }
  } finally {
    if (task.status === "exporting") {
      task.status = previousStatus || "done";
      task.message = previousMessage || task.message;
      task.progress = previousProgress;
    }
    renderTasks();
    saveTasks();
  }
}

async function exportRowsAsCsv(rows, baseName, encodingMode = "auto", onProgress = () => {}) {
  const total = rows.length;
  const headers = total > 0 ? Object.keys(rows[0]) : [];
  const chunkSize = 3000;
  const csvChunks = [];
  let buffer = "";

  onProgress(3, "导出中: 准备 CSV 列头");
  buffer += `${headers.map(csvEscapeCell).join(",")}\r\n`;
  await yieldToUI();

  for (let i = 0; i < total; i += 1) {
    const row = rows[i] || {};
    const line = headers.map((key) => csvEscapeCell(row[key])).join(",");
    buffer += `${line}\r\n`;

    if ((i + 1) % chunkSize === 0 || i === total - 1) {
      csvChunks.push(buffer);
      buffer = "";
      const ratio = total <= 0 ? 1 : (i + 1) / total;
      const percent = Math.min(90, 5 + Math.round(ratio * 85));
      onProgress(percent, `导出中: CSV ${i + 1}/${total} 行`);
      await yieldToUI();
    }
  }

  const rawCsv = csvChunks.join("");
  const resolvedEncoding = resolveCsvEncoding(encodingMode);
  onProgress(93, `导出中: 写入 CSV (${resolvedEncoding.toUpperCase()})`);
  const csv = resolvedEncoding === "utf8bom" ? `\uFEFF${rawCsv}` : rawCsv;
  const saved = await saveFileWithPickerOrDownload(
    `${safeFileName(baseName)}.csv`,
    "text/csv;charset=utf-8",
    csv
  );
  if (!saved) {
    throw new Error("已取消导出");
  }
  onProgress(100, `导出完成: CSV ${total} 行`);
}

async function exportRowsAsXlsx(rows, baseName, onProgress = () => {}) {
  await ensureXLSX();
  const wb = XLSX.utils.book_new();
  const total = rows.length;
  const sheetCount = Math.ceil(total / EXCEL_MAX_ROWS_PER_SHEET);
  onProgress(3, `导出中: 准备 XLSX，共 ${sheetCount} 个 Sheet`);
  for (let i = 0; i < sheetCount; i += 1) {
    const start = i * EXCEL_MAX_ROWS_PER_SHEET;
    const end = Math.min(start + EXCEL_MAX_ROWS_PER_SHEET, total);
    const chunk = rows.slice(start, end);
    const sheetProgress = 5 + Math.round(((i + 1) / Math.max(1, sheetCount)) * 70);
    onProgress(sheetProgress, `导出中: 生成 Sheet ${i + 1}/${sheetCount} (${end - start} 行)`);
    const ws = XLSX.utils.json_to_sheet(chunk);
    XLSX.utils.book_append_sheet(wb, ws, `${truncateSheetName(baseName)}_${i + 1}`);
    await yieldToUI();
  }
  onProgress(80, "导出中: 生成 XLSX 文件");
  const xlsxArray = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  onProgress(93, "导出中: 写入磁盘");
  const saved = await saveFileWithPickerOrDownload(
    `${safeFileName(baseName)}.xlsx`,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    xlsxArray,
    false
  );
  if (!saved) {
    throw new Error("已取消导出");
  }
  onProgress(100, `导出完成: XLSX ${total} 行`);
}

async function exportAllTasksAsXlsx() {
  await ensureXLSX();
  const doneTasks = state.tasks.filter((task) => task.resultRows && task.resultRows.length > 0);
  if (doneTasks.length === 0) {
    alert("暂无可导出的任务结果。");
    return;
  }
  const wb = XLSX.utils.book_new();
  doneTasks.forEach((task) => {
    const ws = XLSX.utils.json_to_sheet(task.resultRows);
    XLSX.utils.book_append_sheet(wb, ws, truncateSheetName(task.name || "task"));
  });
  const xlsxArray = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  await saveFileWithPickerOrDownload(
    `all_tasks_${formatNow()}.xlsx`,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    xlsxArray,
    false
  );
}

async function createDuckDBInstance() {
  await ensureDuckDB();
  const bundles = duckdb.getJsDelivrBundles();
  const bundle = await duckdb.selectBundle(bundles);
  const worker = createDuckDBWorker(bundle.mainWorker);
  const logger = new duckdb.ConsoleLogger();
  const instance = new duckdb.AsyncDuckDB(logger, worker);
  await instance.instantiate(bundle.mainModule, bundle.pthreadWorker);
  return instance;
}

function createDuckDBWorker(workerURL) {
  try {
    return new Worker(workerURL);
  } catch (error) {
    const blob = new Blob([`importScripts("${workerURL}");`], { type: "text/javascript" });
    const blobURL = URL.createObjectURL(blob);
    return new Worker(blobURL);
  }
}

function arrowTableToObjects(table) {
  const rows = table.toArray();
  return rows.map((row) => {
    if (row && typeof row.toJSON === "function") return row.toJSON();
    if (row && typeof row === "object") return { ...row };
    return { value: row };
  });
}

function arrowTableToRows(table) {
  const fields = Array.isArray(table?.schema?.fields) ? table.schema.fields : [];
  const objects = arrowTableToObjects(table);
  return objects.map((obj) => fields.map((field) => String(obj?.[field.name] ?? "")));
}

async function registerCsvInput(db, fileName, sheet) {
  if (sheet?.csvBytes instanceof Uint8Array) {
    const safeBytes = sheet.csvBytes.slice();
    if (typeof db.registerFileBuffer === "function") {
      await db.registerFileBuffer(fileName, safeBytes);
      return;
    }
    const decoder = new TextDecoder("utf-8");
    await db.registerFileText(fileName, decoder.decode(safeBytes));
    return;
  }
  await db.registerFileText(fileName, String(sheet?.csvText ?? ""));
}

function normalizeViewName(name) {
  const normalized = name.toLowerCase().replace(/[^a-z0-9_]+/g, "_");
  return normalized.replace(/^_+|_+$/g, "") || `sheet_${Math.floor(Math.random() * 10000)}`;
}

async function closeExecution(execution) {
  if (!execution) return;
  const conn = execution.conn;
  const db = execution.db;
  const controllers = Array.isArray(execution.fetchControllers) ? execution.fetchControllers : [];
  execution.conn = null;
  execution.db = null;
  execution.fetchControllers = [];
  controllers.forEach((controller) => {
    try {
      controller.abort();
    } catch (error) {
      // ignore abort errors
    }
  });
  try {
    if (conn) await conn.close();
  } catch (error) {
    // ignore close errors
  }
  try {
    if (db) await db.terminate();
  } catch (error) {
    // ignore terminate errors
  }
}

async function postBridgeJSON(path, payload, signal) {
  let response;
  try {
    response = await fetch(`${BRIDGE_API_BASE}${path}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload),
      signal
    });
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new TaskAbortError();
    }
    throw new Error(`无法连接本地桥接服务 ${BRIDGE_API_BASE}，请先运行 python src/bridge_server.py`);
  }

  let data = {};
  try {
    data = await response.json();
  } catch (error) {
    data = { ok: false, message: `接口返回非 JSON (${response.status})` };
  }
  if (!response.ok) {
    return {
      ok: false,
      message: data?.message || `HTTP ${response.status}`
    };
  }
  return data;
}

class TaskAbortError extends Error {
  constructor() {
    super("Task aborted");
    this.name = "TaskAbortError";
  }
}

function isAbortError(error) {
  return error instanceof TaskAbortError || String(error?.name || "") === "TaskAbortError";
}

function quoteIdentifier(identifier) {
  return `"${String(identifier).replaceAll('"', '""')}"`;
}

function sqlQuote(value) {
  return String(value).replaceAll("'", "''");
}

function csvEscapeCell(value) {
  const text = String(value ?? "");
  const escaped = text.replaceAll('"', '""');
  if (/[",\r\n]/.test(escaped)) {
    return `"${escaped}"`;
  }
  return escaped;
}

async function yieldToUI() {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function formatDuration(ms) {
  const totalSeconds = Math.max(0, Math.floor((ms || 0) / 1000));
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
}

function formatEta(elapsedMs, doneRows, totalRows) {
  if (!totalRows || doneRows <= 0 || doneRows >= totalRows) return "--";
  const speed = doneRows / Math.max(1, elapsedMs / 1000);
  if (!speed || !Number.isFinite(speed)) return "--";
  const remainingRows = totalRows - doneRows;
  const remainingMs = (remainingRows / speed) * 1000;
  return formatDuration(remainingMs);
}

function formatNow() {
  const now = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${now.getFullYear()}${p(now.getMonth() + 1)}${p(now.getDate())}_${p(now.getHours())}${p(now.getMinutes())}${p(now.getSeconds())}`;
}

function downloadTextFile(name, text, mime) {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

async function saveFileWithPickerOrDownload(fileName, mime, content, usePicker = true) {
  const blob = content instanceof Blob ? content : new Blob([content], { type: mime });
  if (usePicker && typeof window.showSaveFilePicker === "function") {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: fileName,
        types: [
          {
            description: fileName.toLowerCase().endsWith(".xlsx") ? "Excel 文件" : "CSV 文件",
            accept: { [mime]: [fileName.slice(fileName.lastIndexOf("."))] }
          }
        ]
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true;
    } catch (error) {
      if (error?.name === "AbortError") return false;
    }
  }
  downloadBlobFile(fileName, blob);
  return true;
}

function downloadBlobFile(name, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

function safeFileName(name) {
  return String(name || "result").replace(/[\\/:*?"<>|]+/g, "_").slice(0, 80);
}

function normalizeExportRows(rawRows) {
  if (!Array.isArray(rawRows)) {
    throw new Error("结果数据不是数组");
  }
  return rawRows.map((row) => {
    if (row && typeof row === "object" && !Array.isArray(row)) {
      return row;
    }
    return { value: String(row ?? "") };
  });
}

function showToast(message, type = "info", duration = 2600) {
  const text = String(message || "").trim();
  if (!text) return;
  let container = document.getElementById("toastContainer");
  if (!container) {
    container = document.createElement("div");
    container.id = "toastContainer";
    container.className = "toast-container";
    document.body.appendChild(container);
  }
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = text;
  container.appendChild(toast);

  requestAnimationFrame(() => {
    toast.classList.add("show");
  });
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => toast.remove(), 220);
  }, Math.max(800, Number(duration) || 2600));
}

function truncateSheetName(name) {
  return safeFileName(name).slice(0, 25) || "Sheet";
}

function generateId() {
  if (typeof crypto !== "undefined" && crypto && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function resolveCsvEncoding(mode) {
  if (mode === "utf8" || mode === "utf8bom") return mode;
  const platform = String(navigator.platform || "").toLowerCase();
  const userAgent = String(navigator.userAgent || "").toLowerCase();
  const isMacLike = platform.includes("mac") || /iphone|ipad|ipod/.test(userAgent);
  return isMacLike ? "utf8" : "utf8bom";
}

async function ensureXLSX() {
  if (XLSX) return XLSX;
  try {
    XLSX = await import(XLSX_CDN);
    return XLSX;
  } catch (error) {
    throw new Error("XLSX 依赖加载失败，请检查网络后重试。");
  }
}

async function ensureDuckDB() {
  if (duckdb) return duckdb;
  try {
    duckdb = await import(DUCKDB_CDN);
    return duckdb;
  } catch (error) {
    throw new Error("DuckDB 依赖加载失败，请检查网络后重试。");
  }
}

function escapeHtml(input) {
  return String(input ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttr(input) {
  return escapeHtml(input).replaceAll("\n", " ");
}
