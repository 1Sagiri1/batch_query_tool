import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";
import * as duckdb from "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

const PROFILE_STORAGE_KEY = "batch_query_profiles_v1";
const TASK_STORAGE_KEY = "batch_query_tasks_v1";
const LARGE_ROW_THRESHOLD = 200000;
const EXCEL_MAX_ROWS_PER_SHEET = 1048576;

const state = {
  profiles: [],
  activeProfileId: "",
  fileContext: null,
  selectedColumnIndex: null,
  hasHeader: true,
  tasks: []
};

const els = {
  profileTabs: document.getElementById("profileTabs"),
  addProfileBtn: document.getElementById("addProfileBtn"),
  exportProfilesBtn: document.getElementById("exportProfilesBtn"),
  importProfilesInput: document.getElementById("importProfilesInput"),
  profileName: document.getElementById("profileName"),
  profileMode: document.getElementById("profileMode"),
  profileDialect: document.getElementById("profileDialect"),
  profileDriver: document.getElementById("profileDriver"),
  profileHost: document.getElementById("profileHost"),
  profilePort: document.getElementById("profilePort"),
  profileUser: document.getElementById("profileUser"),
  profilePassword: document.getElementById("profilePassword"),
  profileDatabase: document.getElementById("profileDatabase"),
  saveProfileBtn: document.getElementById("saveProfileBtn"),
  deleteProfileBtn: document.getElementById("deleteProfileBtn"),
  dataFileInput: document.getElementById("dataFileInput"),
  dropZone: document.getElementById("dropZone"),
  fileMeta: document.getElementById("fileMeta"),
  sheetSelect: document.getElementById("sheetSelect"),
  previewTable: document.getElementById("previewTable"),
  selectedColumnMeta: document.getElementById("selectedColumnMeta"),
  addTaskBtn: document.getElementById("addTaskBtn"),
  runAllBtn: document.getElementById("runAllBtn"),
  exportAllXlsxBtn: document.getElementById("exportAllXlsxBtn"),
  taskList: document.getElementById("taskList")
};

boot();

function boot() {
  loadProfiles();
  loadTasks();
  bindEvents();
  renderProfiles();
  renderTasks();
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
  els.deleteProfileBtn.addEventListener("click", deleteActiveProfile);

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
      state.hasHeader = input.value === "true";
      renderSheetPreview();
    });
  });

  els.sheetSelect.addEventListener("change", () => {
    renderSheetPreview();
  });

  els.previewTable.addEventListener("click", (event) => {
    const th = event.target.closest("th[data-col-index]");
    if (!th) return;
    state.selectedColumnIndex = Number(th.dataset.colIndex);
    renderSheetPreview();
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
    state.profiles = raw ? JSON.parse(raw) : [];
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
  return {
    id: crypto.randomUUID(),
    name: "默认配置",
    mode: "duckdb-local",
    dialect: "mysql",
    driver: "pymysql",
    host: "",
    port: "3306",
    username: "",
    password: "",
    database: "",
    ...overrides
  };
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
  els.profileMode.value = profile.mode || "duckdb-local";
  els.profileDialect.value = profile.dialect || "";
  els.profileDriver.value = profile.driver || "";
  els.profileHost.value = profile.host || "";
  els.profilePort.value = profile.port || "";
  els.profileUser.value = profile.username || "";
  els.profilePassword.value = profile.password || "";
  els.profileDatabase.value = profile.database || "";
}

function saveActiveProfileFromForm() {
  const profile = getActiveProfile();
  if (!profile) return;
  profile.name = els.profileName.value.trim() || "未命名配置";
  profile.mode = els.profileMode.value;
  profile.dialect = els.profileDialect.value.trim();
  profile.driver = els.profileDriver.value.trim();
  profile.host = els.profileHost.value.trim();
  profile.port = els.profilePort.value.trim();
  profile.username = els.profileUser.value.trim();
  profile.password = els.profilePassword.value;
  profile.database = els.profileDatabase.value.trim();
  saveProfiles();
  renderProfiles();
  renderTasks();
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
  return createProfile({
    name: raw?.name || raw?.database || "导入配置",
    mode: raw?.mode || "duckdb-local",
    dialect: raw?.dialect || raw?.DIALECT || "mysql",
    driver: raw?.driver || raw?.DRIVER || "pymysql",
    host: raw?.host || raw?.HOST || "",
    port: String(raw?.port || raw?.PORT || "3306"),
    username: raw?.username || raw?.USERNAME || "",
    password: raw?.password || raw?.PASSWORD || "",
    database: raw?.database || raw?.DATABASE || ""
  });
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
    id: crypto.randomUUID(),
    name: `任务 ${state.tasks.length + 1}`,
    profileId: state.activeProfileId,
    sql: "SELECT * FROM source_data WHERE 1=1 AND id IN ({})",
    batchSize: 3000,
    status: "idle",
    progress: 0,
    message: "待执行",
    exportType: "csv",
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

  els.taskList.innerHTML = state.tasks
    .map((task) => {
      const statusClass = task.status === "done" ? "ok" : task.status === "error" ? "err" : "";
      return `
        <article class="task-card" data-task-id="${task.id}">
          <div class="task-head">
            <strong>${escapeHtml(task.name)}</strong>
            <div class="head-actions">
              <button type="button" data-act="run">运行</button>
              <button type="button" data-act="export">导出</button>
              <button type="button" class="danger" data-act="delete">删除</button>
            </div>
          </div>
          <div class="task-grid">
            <label>任务名<input data-field="name" value="${escapeAttr(task.name)}"></label>
            <label>配置
              <select data-field="profileId">${options}</select>
            </label>
            <label>批大小<input data-field="batchSize" type="number" min="1" value="${task.batchSize}"></label>
            <label>导出格式
              <select data-field="exportType">
                <option value="csv">CSV</option>
                <option value="xlsx">XLSX(多Sheet)</option>
              </select>
            </label>
            <label class="sql">SQL(支持 {} 占位符或 {{values}} 占位符)
              <textarea data-field="sql">${escapeHtml(task.sql)}</textarea>
            </label>
          </div>
          <progress max="100" value="${Number(task.progress) || 0}"></progress>
          <div class="status ${statusClass}">${escapeHtml(task.message || "")}</div>
        </article>
      `;
    })
    .join("");

  state.tasks.forEach((task) => {
    const card = els.taskList.querySelector(`[data-task-id="${task.id}"]`);
    if (!card) return;
    const profileSelect = card.querySelector('[data-field="profileId"]');
    const exportSelect = card.querySelector('[data-field="exportType"]');
    if (profileSelect) profileSelect.value = task.profileId;
    if (exportSelect) exportSelect.value = task.exportType;
  });
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

async function onTaskListClick(event) {
  const btn = event.target.closest("button[data-act]");
  if (!btn) return;
  const card = btn.closest("[data-task-id]");
  if (!card) return;
  const task = state.tasks.find((item) => item.id === card.dataset.taskId);
  if (!task) return;

  const act = btn.dataset.act;
  if (act === "delete") {
    state.tasks = state.tasks.filter((item) => item.id !== task.id);
    saveTasks();
    renderTasks();
    return;
  }

  if (act === "run") {
    await runTask(task.id);
    return;
  }

  if (act === "export") {
    exportSingleTask(task);
  }
}

async function importDataFile(file) {
  const name = file.name.toLowerCase();
  if (!(name.endsWith(".csv") || name.endsWith(".xlsx") || name.endsWith(".xls"))) {
    alert("仅支持 .csv / .xlsx / .xls");
    return;
  }

  try {
    const workbook = await readWorkbook(file);
    const sheets = workbook.SheetNames.map((sheetName) => {
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

    const totalRows = sheets.reduce((sum, sheet) => sum + sheet.rowCount, 0);
    state.fileContext = {
      fileName: file.name,
      sheets,
      totalRows,
      isLarge: totalRows > LARGE_ROW_THRESHOLD
    };
    state.selectedColumnIndex = null;

    els.fileMeta.textContent = `${file.name} | Sheet: ${sheets.length} | 总行数: ${totalRows}${state.fileContext.isLarge ? " | 大文件: 使用 DuckDB 处理" : ""}`;
    els.sheetSelect.innerHTML = sheets
      .map((sheet, idx) => `<option value="${idx}">${escapeHtml(sheet.name)} (${sheet.rowCount} 行)</option>`)
      .join("");
    renderSheetPreview();
  } catch (error) {
    console.error(error);
    alert(`文件读取失败: ${error.message}`);
  }
}

async function readWorkbook(file) {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".csv")) {
    const text = await file.text();
    return XLSX.read(text, { type: "string" });
  }
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array" });
}

function getPreviewRowsFromCsv(csvText, rowLimit) {
  const wb = XLSX.read(csvText, { type: "string" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
  return rows.slice(0, rowLimit);
}

function getCurrentSheet() {
  if (!state.fileContext || state.fileContext.sheets.length === 0) return null;
  const idx = Number(els.sheetSelect.value || 0);
  return state.fileContext.sheets[idx] || state.fileContext.sheets[0];
}

function renderSheetPreview() {
  const sheet = getCurrentSheet();
  if (!sheet) {
    els.previewTable.innerHTML = "";
    els.selectedColumnMeta.textContent = "未选择列";
    return;
  }

  const rows = sheet.previewRows || [];
  const headerRow = rows[0] || [];
  const bodyRows = state.hasHeader ? rows.slice(1, 11) : rows.slice(0, 10);
  const colCount = Math.max(headerRow.length, ...bodyRows.map((row) => row.length), 1);

  const headers = [];
  for (let i = 0; i < colCount; i += 1) {
    if (state.hasHeader) {
      headers.push(String(headerRow[i] ?? `col_${i + 1}`));
    } else {
      headers.push(`col_${i + 1}`);
    }
  }

  const headerHtml = headers
    .map((name, idx) => {
      const selected = idx === state.selectedColumnIndex ? "selected" : "";
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

  if (state.selectedColumnIndex === null) {
    els.selectedColumnMeta.textContent = "未选择列";
  } else {
    els.selectedColumnMeta.textContent = `已选择第 ${state.selectedColumnIndex + 1} 列`;
  }
}

async function runAllTasksInParallel() {
  if (state.tasks.length === 0) {
    alert("请先新增任务。");
    return;
  }
  if (!state.fileContext) {
    alert("请先导入文件并选择批处理列。");
    return;
  }
  if (state.selectedColumnIndex === null) {
    alert("请在预览表头点击选择用于批处理的列。");
    return;
  }
  await Promise.allSettled(state.tasks.map((task) => runTask(task.id)));
}

async function runTask(taskId) {
  const task = state.tasks.find((item) => item.id === taskId);
  if (!task) return;
  if (!state.fileContext) {
    alert("请先导入文件。");
    return;
  }
  if (state.selectedColumnIndex === null) {
    alert("请先选择批处理列。");
    return;
  }

  task.status = "running";
  task.progress = 0;
  task.resultRows = [];
  task.resultColumns = [];
  task.message = "初始化 DuckDB...";
  renderTasks();

  try {
    const values = await extractDistinctValues(state.selectedColumnIndex, state.hasHeader);
    if (values.length === 0) {
      task.status = "done";
      task.progress = 100;
      task.message = "查询列没有可用值。";
      renderTasks();
      saveTasks();
      return;
    }

    const db = await createDuckDBInstance();
    const conn = await db.connect();
    await registerSheetsAsViews(db, conn, state.fileContext.sheets, state.hasHeader);

    const chunks = splitIntoChunks(values, Math.max(1, Number(task.batchSize) || 3000));
    let mergedRows = [];
    let columns = [];

    for (let i = 0; i < chunks.length; i += 1) {
      const chunk = chunks[i];
      const query = formatSQL(task.sql, chunk);
      const table = await conn.query(query);
      const rows = arrowTableToObjects(table);

      if (rows.length > 0 && columns.length === 0) {
        columns = Object.keys(rows[0]);
      }
      mergedRows = mergedRows.concat(rows);
      task.progress = Math.round(((i + 1) / chunks.length) * 100);
      task.message = `执行中: ${i + 1}/${chunks.length} 批，累计 ${mergedRows.length} 行`;
      renderTasks();
    }

    await conn.close();
    await db.terminate();

    task.status = "done";
    task.progress = 100;
    task.resultRows = mergedRows;
    task.resultColumns = columns;
    task.message = `完成: ${mergedRows.length} 行`;
  } catch (error) {
    console.error(error);
    task.status = "error";
    task.message = `失败: ${error.message}`;
  }

  renderTasks();
  saveTasks();
}

async function extractDistinctValues(colIndex, hasHeader) {
  const db = await createDuckDBInstance();
  const conn = await db.connect();
  const values = new Set();

  for (const sheet of state.fileContext.sheets) {
    const fileName = `${sheet.viewName}.csv`;
    await db.registerFileText(fileName, sheet.csvText);
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
  return Array.from(values);
}

async function registerSheetsAsViews(db, conn, sheets, hasHeader) {
  const headerFlag = hasHeader ? "true" : "false";
  const viewNames = [];
  for (const sheet of sheets) {
    const fileName = `${sheet.viewName}.csv`;
    const escapedFile = sqlQuote(fileName);
    await db.registerFileText(fileName, sheet.csvText);
    await conn.query(
      `CREATE OR REPLACE VIEW ${quoteIdentifier(sheet.viewName)} AS
       SELECT * FROM read_csv_auto('${escapedFile}', header=${headerFlag}, all_varchar=true)`
    );
    viewNames.push(sheet.viewName);
  }

  if (viewNames.length > 0) {
    const unionSQL = viewNames
      .map((name) => `SELECT * FROM ${quoteIdentifier(name)}`)
      .join(" UNION ALL BY NAME ");
    await conn.query(`CREATE OR REPLACE VIEW source_data AS ${unionSQL}`);
  }
}

function formatSQL(sqlTemplate, batchValues) {
  const values = batchValues.map((item) => `'${sqlQuote(item)}'`).join(", ");
  if (sqlTemplate.includes("{}")) return sqlTemplate.replaceAll("{}", values);
  if (sqlTemplate.includes("{{values}}")) return sqlTemplate.replaceAll("{{values}}", values);
  return sqlTemplate;
}

function splitIntoChunks(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

function exportSingleTask(task) {
  if (!task.resultRows || task.resultRows.length === 0) {
    alert("该任务暂无结果可导出。");
    return;
  }
  if (task.exportType === "xlsx") {
    exportRowsAsXlsx(task.resultRows, task.name || "task");
  } else {
    exportRowsAsCsv(task.resultRows, task.name || "task");
  }
}

function exportRowsAsCsv(rows, baseName) {
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  downloadTextFile(`${safeFileName(baseName)}_${formatNow()}.csv`, csv, "text/csv;charset=utf-8");
}

function exportRowsAsXlsx(rows, baseName) {
  const wb = XLSX.utils.book_new();
  const total = rows.length;
  const sheetCount = Math.ceil(total / EXCEL_MAX_ROWS_PER_SHEET);
  for (let i = 0; i < sheetCount; i += 1) {
    const start = i * EXCEL_MAX_ROWS_PER_SHEET;
    const end = Math.min(start + EXCEL_MAX_ROWS_PER_SHEET, total);
    const chunk = rows.slice(start, end);
    const ws = XLSX.utils.json_to_sheet(chunk);
    XLSX.utils.book_append_sheet(wb, ws, `${truncateSheetName(baseName)}_${i + 1}`);
  }
  XLSX.writeFile(wb, `${safeFileName(baseName)}_${formatNow()}.xlsx`);
}

function exportAllTasksAsXlsx() {
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
  XLSX.writeFile(wb, `all_tasks_${formatNow()}.xlsx`);
}

async function createDuckDBInstance() {
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

function normalizeViewName(name) {
  const normalized = name.toLowerCase().replace(/[^a-z0-9_]+/g, "_");
  return normalized.replace(/^_+|_+$/g, "") || `sheet_${Math.floor(Math.random() * 10000)}`;
}

function quoteIdentifier(identifier) {
  return `"${String(identifier).replaceAll('"', '""')}"`;
}

function sqlQuote(value) {
  return String(value).replaceAll("'", "''");
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

function safeFileName(name) {
  return String(name || "result").replace(/[\\/:*?"<>|]+/g, "_").slice(0, 80);
}

function truncateSheetName(name) {
  return safeFileName(name).slice(0, 25) || "Sheet";
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
