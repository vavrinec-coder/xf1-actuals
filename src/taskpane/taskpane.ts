/* eslint-disable no-undef */
/* global Office */

import * as XLSX from "xlsx";

type StatusKind = "info" | "success" | "error";
type EntityDepartmentMode = "blank" | "constant" | "range";
type DateMode = "constant" | "range";
type DimensionShape = "blank" | "constant" | "singleCell" | "singleColumn" | "multiColumn";

interface SourceState {
  id: number;
  file: File | null;
  workbook: XLSX.WorkBook | null;
  sheetNames: string[];
  sourceSheet: string;
  sourcePath: string;
  accountRange: string;
  valueRange: string;
  entityMode: EntityDepartmentMode;
  entityConstant: string;
  entityRange: string;
  departmentMode: EntityDepartmentMode;
  departmentConstant: string;
  departmentRange: string;
  dateMode: DateMode;
  dateConstant: string;
  dateRange: string;
}

interface MatrixData {
  rows: number;
  cols: number;
  values: unknown[][];
}

interface PreparedDimension {
  shape: DimensionShape;
  constantValue: unknown;
  matrix: MatrixData | null;
}

interface PreparedSource {
  source: SourceState;
  accountMatrix: MatrixData;
  valueMatrix: MatrixData;
  entity: PreparedDimension;
  department: PreparedDimension;
  date: PreparedDimension;
}

interface ConsolidatedRow {
  Account: string;
  Entity: string;
  Department: string;
  Date: unknown;
  Value: number;
  SourceFile: string;
  SourceSheet: string;
}

interface ConfigSource {
  sourcePath: string;
  sourceSheet: string;
  accountRange: string;
  valueRange: string;
  entityMode: EntityDepartmentMode;
  entityConstant: string;
  entityRange: string;
  departmentMode: EntityDepartmentMode;
  departmentConstant: string;
  departmentRange: string;
  dateMode: DateMode;
  dateConstant: string;
  dateRange: string;
}

interface ConfigPayload {
  version: string;
  createdAtUtc: string;
  outputFileName: string;
  sources: ConfigSource[];
}

const MAX_SOURCES = 12;
const OUTPUT_HEADERS = [
  "Account",
  "Entity",
  "Department",
  "Date",
  "Value",
  "SourceFile",
  "SourceSheet",
];
const OUTPUT_FILE_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CONFIG_FILE_EXTENSION = ".xf1config.json";

let nextSourceId = 1;
let panelExpanded = false;
let sources: SourceState[] = [];
let outputFileName = "consolidated-actuals.xlsx";
let outputFileHandle: any = null;
let isRunning = false;

let sideloadMsgElement: HTMLElement;
let appBodyElement: HTMLElement;
let panelElement: HTMLElement;
let sourceContainerElement: HTMLElement;
let sourceCountLabelElement: HTMLElement;
let outputFileNameElement: HTMLInputElement;
let outputLocationLabelElement: HTMLElement;
let statusElement: HTMLElement;
let addSourceButtonElement: HTMLButtonElement;
let runButtonElement: HTMLButtonElement;
let loadConfigInputElement: HTMLInputElement;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    return;
  }

  sideloadMsgElement = requireElement<HTMLElement>("sideload-msg");
  appBodyElement = requireElement<HTMLElement>("app-body");
  panelElement = requireElement<HTMLElement>("consolidate-panel");
  sourceContainerElement = requireElement<HTMLElement>("sources-container");
  sourceCountLabelElement = requireElement<HTMLElement>("source-count-label");
  outputFileNameElement = requireElement<HTMLInputElement>("output-file-name");
  outputLocationLabelElement = requireElement<HTMLElement>("output-location-label");
  statusElement = requireElement<HTMLElement>("status-message");
  addSourceButtonElement = requireElement<HTMLButtonElement>("add-source-btn");
  runButtonElement = requireElement<HTMLButtonElement>("run-consolidation-btn");
  loadConfigInputElement = requireElement<HTMLInputElement>("load-config-input");

  sideloadMsgElement.style.display = "none";
  appBodyElement.style.display = "flex";

  initializeUi();
});

function initializeUi(): void {
  const toggleButton = requireElement<HTMLButtonElement>("toggle-consolidate-btn");
  const browseOutputButton = requireElement<HTMLButtonElement>("browse-output-btn");
  const saveConfigButton = requireElement<HTMLButtonElement>("save-config-btn");
  const loadConfigButton = requireElement<HTMLButtonElement>("load-config-btn");

  toggleButton.addEventListener("click", toggleConsolidatePanel);
  addSourceButtonElement.addEventListener("click", addSourceBlock);
  browseOutputButton.addEventListener("click", browseOutputLocation);
  saveConfigButton.addEventListener("click", saveConfig);
  loadConfigButton.addEventListener("click", () => loadConfigInputElement.click());
  loadConfigInputElement.addEventListener("change", loadConfig);
  runButtonElement.addEventListener("click", runConsolidation);
  outputFileNameElement.addEventListener("input", () => {
    outputFileName = outputFileNameElement.value.trim();
  });

  sources = [createDefaultSource()];
  renderSourceCards();
  refreshUiState();
  setStatus("info", 'Click "Consolidate from multiple files" to start configuring File 1.');
}

function toggleConsolidatePanel(): void {
  panelExpanded = !panelExpanded;
  panelElement.classList.toggle("hidden", !panelExpanded);
  setStatus(
    "info",
    panelExpanded ? "Configuration section opened." : "Configuration section collapsed."
  );
}

function createDefaultSource(): SourceState {
  return {
    id: nextSourceId++,
    file: null,
    workbook: null,
    sheetNames: [],
    sourceSheet: "",
    sourcePath: "",
    accountRange: "",
    valueRange: "",
    entityMode: "blank",
    entityConstant: "",
    entityRange: "",
    departmentMode: "blank",
    departmentConstant: "",
    departmentRange: "",
    dateMode: "constant",
    dateConstant: "",
    dateRange: "",
  };
}

function addSourceBlock(): void {
  if (sources.length >= MAX_SOURCES) {
    setStatus("error", "You can configure a maximum of 12 source files in one run.");
    return;
  }

  sources.push(createDefaultSource());
  renderSourceCards();
  refreshUiState();
  setStatus("info", "Added a new source block.");
}

function removeSourceBlock(sourceId: number): void {
  if (sources.length === 1) {
    setStatus("error", "At least one source block is required.");
    return;
  }

  sources = sources.filter((item) => item.id !== sourceId);
  renderSourceCards();
  refreshUiState();
  setStatus("info", "Removed the selected source block.");
}

function renderSourceCards(): void {
  if (sources.length === 0) {
    sourceContainerElement.innerHTML = '<p class="meta-text">No source blocks configured.</p>';
    return;
  }

  const markup = sources
    .map((source, index) => {
      const sheetOptions = source.sheetNames
        .map((sheet) => {
          return `<option value="${escapeHtml(sheet)}" ${source.sourceSheet === sheet ? "selected" : ""}>${escapeHtml(sheet)}</option>`;
        })
        .join("");

      const fileLabel = source.file
        ? `Loaded file: ${escapeHtml(source.file.name)}`
        : "No file selected yet.";
      const removeButton =
        sources.length > 1
          ? `<button type="button" class="btn btn-danger source-remove-btn" data-source-id="${source.id}">Remove</button>`
          : "";

      return `
        <div class="source-card">
          <div class="source-header">
            <h4 class="source-title">File ${index + 1}</h4>
            ${removeButton}
          </div>
          <div class="grid">
            <div class="grid-full">
              <label class="field-label" for="source-file-${source.id}">Source file (.xlsx)</label>
              <input id="source-file-${source.id}" class="file-input source-file-input" type="file" accept=".xlsx" data-source-id="${source.id}" />
              <p class="hint">${fileLabel}</p>
            </div>

            <div class="grid-full">
              <label class="field-label" for="source-path-${source.id}">Source full path (for config reuse)</label>
              <input id="source-path-${source.id}" class="text-input source-bind-input" type="text" data-source-id="${source.id}" data-field="sourcePath" value="${escapeHtml(source.sourcePath)}" placeholder="Example: D:\\Data\\Book1.xlsx" />
            </div>

            <div class="grid-full">
              <label class="field-label" for="source-sheet-${source.id}">Source sheet</label>
              <select id="source-sheet-${source.id}" class="select-input source-bind-input" data-source-id="${source.id}" data-field="sourceSheet">
                <option value="">Select sheet</option>
                ${sheetOptions}
              </select>
            </div>

            <div>
              <label class="field-label" for="account-range-${source.id}">Account cell range</label>
              <input id="account-range-${source.id}" class="text-input source-bind-input" type="text" data-source-id="${source.id}" data-field="accountRange" value="${escapeHtml(source.accountRange)}" placeholder="B10:B110" />
            </div>

            <div>
              <label class="field-label" for="value-range-${source.id}">Value cell range</label>
              <input id="value-range-${source.id}" class="text-input source-bind-input" type="text" data-source-id="${source.id}" data-field="valueRange" value="${escapeHtml(source.valueRange)}" placeholder="C10:K110" />
            </div>

            ${renderDimensionEditor("Entity", "entity", source)}
            ${renderDimensionEditor("Department", "department", source)}
            ${renderDateEditor(source)}
          </div>
        </div>`;
    })
    .join("");

  sourceContainerElement.innerHTML = markup;
  wireSourceCardEvents();
}

function renderDimensionEditor(
  label: string,
  key: "entity" | "department",
  source: SourceState
): string {
  const modeField = key === "entity" ? "entityMode" : "departmentMode";
  const constantField = key === "entity" ? "entityConstant" : "departmentConstant";
  const rangeField = key === "entity" ? "entityRange" : "departmentRange";
  const mode = source[modeField];
  const constantDisabled = mode !== "constant" ? "disabled" : "";
  const rangeDisabled = mode !== "range" ? "disabled" : "";

  return `
    <div>
      <label class="field-label" for="${key}-mode-${source.id}">${label} mode</label>
      <select id="${key}-mode-${source.id}" class="select-input source-bind-input" data-source-id="${source.id}" data-field="${modeField}">
        <option value="blank" ${mode === "blank" ? "selected" : ""}>Blank</option>
        <option value="constant" ${mode === "constant" ? "selected" : ""}>Constant</option>
        <option value="range" ${mode === "range" ? "selected" : ""}>Range</option>
      </select>
    </div>
    <div>
      <label class="field-label" for="${key}-constant-${source.id}">${label} constant</label>
      <input id="${key}-constant-${source.id}" class="text-input source-bind-input" type="text" data-source-id="${source.id}" data-field="${constantField}" value="${escapeHtml(source[constantField])}" ${constantDisabled} />
    </div>
    <div class="grid-full">
      <label class="field-label" for="${key}-range-${source.id}">${label} range</label>
      <input id="${key}-range-${source.id}" class="text-input source-bind-input" type="text" data-source-id="${source.id}" data-field="${rangeField}" value="${escapeHtml(source[rangeField])}" ${rangeDisabled} placeholder="Single column or 1 x N row when Value is multi-column" />
    </div>`;
}

function renderDateEditor(source: SourceState): string {
  const constantDisabled = source.dateMode !== "constant" ? "disabled" : "";
  const rangeDisabled = source.dateMode !== "range" ? "disabled" : "";

  return `
    <div>
      <label class="field-label" for="date-mode-${source.id}">Date mode</label>
      <select id="date-mode-${source.id}" class="select-input source-bind-input" data-source-id="${source.id}" data-field="dateMode">
        <option value="constant" ${source.dateMode === "constant" ? "selected" : ""}>Constant</option>
        <option value="range" ${source.dateMode === "range" ? "selected" : ""}>Range</option>
      </select>
    </div>
    <div>
      <label class="field-label" for="date-constant-${source.id}">Date constant</label>
      <input id="date-constant-${source.id}" class="text-input source-bind-input" type="date" data-source-id="${source.id}" data-field="dateConstant" value="${escapeHtml(source.dateConstant)}" ${constantDisabled} />
    </div>
    <div class="grid-full">
      <label class="field-label" for="date-range-${source.id}">Date range</label>
      <input id="date-range-${source.id}" class="text-input source-bind-input" type="text" data-source-id="${source.id}" data-field="dateRange" value="${escapeHtml(source.dateRange)}" ${rangeDisabled} placeholder="Single column or 1 x N row when Value is multi-column" />
    </div>`;
}

function wireSourceCardEvents(): void {
  const removeButtons =
    sourceContainerElement.querySelectorAll<HTMLButtonElement>(".source-remove-btn");
  removeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const sourceId = Number(button.dataset.sourceId);
      removeSourceBlock(sourceId);
    });
  });

  const fileInputs =
    sourceContainerElement.querySelectorAll<HTMLInputElement>(".source-file-input");
  fileInputs.forEach((input) => {
    input.addEventListener("change", async () => {
      const sourceId = Number(input.dataset.sourceId);
      const selectedFile = input.files && input.files.length > 0 ? input.files[0] : null;
      await handleSourceFileSelection(sourceId, selectedFile);
    });
  });

  const bindInputs = sourceContainerElement.querySelectorAll<HTMLInputElement | HTMLSelectElement>(
    ".source-bind-input"
  );
  bindInputs.forEach((input) => {
    const handler = () => {
      const sourceId = Number(input.dataset.sourceId);
      const field = input.dataset.field || "";
      updateSourceField(sourceId, field, input.value);
    };

    input.addEventListener("input", handler);
    input.addEventListener("change", handler);
  });
}

function updateSourceField(sourceId: number, field: string, value: string): void {
  const source = sources.find((item) => item.id === sourceId);
  if (!source) {
    return;
  }

  const supportedFields = [
    "sourcePath",
    "sourceSheet",
    "accountRange",
    "valueRange",
    "entityMode",
    "entityConstant",
    "entityRange",
    "departmentMode",
    "departmentConstant",
    "departmentRange",
    "dateMode",
    "dateConstant",
    "dateRange",
  ];

  if (supportedFields.indexOf(field) < 0) {
    return;
  }

  (source as any)[field] = value;
  renderSourceCards();
  refreshUiState();
}

async function handleSourceFileSelection(
  sourceId: number,
  selectedFile: File | null
): Promise<void> {
  const source = sources.find((item) => item.id === sourceId);
  if (!source) {
    return;
  }

  if (!selectedFile) {
    source.file = null;
    source.workbook = null;
    source.sheetNames = [];
    source.sourceSheet = "";
    renderSourceCards();
    setStatus("info", "Source file cleared.");
    return;
  }

  if (!selectedFile.name.toLowerCase().endsWith(".xlsx")) {
    setStatus("error", `File ${selectedFile.name} is not .xlsx. Please choose a .xlsx file.`);
    return;
  }

  try {
    const buffer = await selectedFile.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false });

    source.file = selectedFile;
    source.workbook = workbook;
    source.sheetNames = workbook.SheetNames.slice();
    source.sourceSheet = source.sheetNames.length > 0 ? source.sheetNames[0] : "";

    if (!source.sourcePath) {
      source.sourcePath = selectedFile.name;
    }

    renderSourceCards();
    refreshUiState();
    setStatus("success", `Loaded ${selectedFile.name}. Select the sheet and mapping ranges.`);
  } catch (error) {
    setStatus("error", `Unable to read ${selectedFile.name}. ${toErrorMessage(error)}`);
  }
}

async function browseOutputLocation(): Promise<void> {
  const picker = (window as any).showSaveFilePicker;
  const normalizedName = normalizeOutputFileName(outputFileName);

  if (typeof picker !== "function") {
    outputFileHandle = null;
    outputLocationLabelElement.textContent =
      "Save picker unavailable here. Output will download to default folder.";
    setStatus(
      "info",
      "Save picker unavailable. The consolidated file will be downloaded in browser default folder."
    );
    return;
  }

  try {
    const handle = await picker({
      suggestedName: normalizedName,
      types: [
        {
          description: "Excel Workbook",
          accept: {
            [OUTPUT_FILE_TYPE]: [".xlsx"],
          },
        },
      ],
    });

    outputFileHandle = handle;
    outputLocationLabelElement.textContent = `Selected file: ${handle.name || normalizedName}`;
    setStatus(
      "success",
      "Output destination selected. Existing file at this location will be overwritten."
    );
  } catch (error: any) {
    if (error && error.name === "AbortError") {
      setStatus("info", "Save location selection cancelled.");
      return;
    }

    setStatus("error", `Unable to choose save location. ${toErrorMessage(error)}`);
  }
}

function saveConfig(): void {
  try {
    const payload: ConfigPayload = {
      version: "1.0.0",
      createdAtUtc: new Date().toISOString(),
      outputFileName: normalizeOutputFileName(outputFileName),
      sources: sources.map((source) => {
        return {
          sourcePath: source.sourcePath,
          sourceSheet: source.sourceSheet,
          accountRange: source.accountRange,
          valueRange: source.valueRange,
          entityMode: source.entityMode,
          entityConstant: source.entityConstant,
          entityRange: source.entityRange,
          departmentMode: source.departmentMode,
          departmentConstant: source.departmentConstant,
          departmentRange: source.departmentRange,
          dateMode: source.dateMode,
          dateConstant: source.dateConstant,
          dateRange: source.dateRange,
        };
      }),
    };

    const configName = buildConfigFileName();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    downloadBlob(blob, configName);
    setStatus("success", `Config exported: ${configName}`);
  } catch (error) {
    setStatus("error", `Unable to save config. ${toErrorMessage(error)}`);
  }
}

async function loadConfig(): Promise<void> {
  const file =
    loadConfigInputElement.files && loadConfigInputElement.files.length > 0
      ? loadConfigInputElement.files[0]
      : null;
  loadConfigInputElement.value = "";

  if (!file) {
    return;
  }

  try {
    const rawText = await file.text();
    const parsed = JSON.parse(rawText) as ConfigPayload;

    if (!parsed || !Array.isArray(parsed.sources)) {
      throw new Error("Invalid config format.");
    }

    if (parsed.sources.length < 1 || parsed.sources.length > MAX_SOURCES) {
      throw new Error(`Config must include between 1 and ${MAX_SOURCES} sources.`);
    }

    outputFileName = normalizeOutputFileName(parsed.outputFileName || "consolidated-actuals.xlsx");
    outputFileNameElement.value = outputFileName;
    outputFileHandle = null;
    outputLocationLabelElement.textContent = "No save location selected yet.";

    sources = parsed.sources.map((sourceConfig) => {
      return {
        id: nextSourceId++,
        file: null,
        workbook: null,
        sheetNames: [],
        sourceSheet: sourceConfig.sourceSheet || "",
        sourcePath: sourceConfig.sourcePath || "",
        accountRange: sourceConfig.accountRange || "",
        valueRange: sourceConfig.valueRange || "",
        entityMode: sourceConfig.entityMode || "blank",
        entityConstant: sourceConfig.entityConstant || "",
        entityRange: sourceConfig.entityRange || "",
        departmentMode: sourceConfig.departmentMode || "blank",
        departmentConstant: sourceConfig.departmentConstant || "",
        departmentRange: sourceConfig.departmentRange || "",
        dateMode: sourceConfig.dateMode || "constant",
        dateConstant: sourceConfig.dateConstant || "",
        dateRange: sourceConfig.dateRange || "",
      };
    });

    panelExpanded = true;
    panelElement.classList.remove("hidden");
    renderSourceCards();
    refreshUiState();
    setStatus(
      "info",
      "Config loaded. Please reselect each source file before running consolidation."
    );
  } catch (error) {
    setStatus("error", `Unable to load config. ${toErrorMessage(error)}`);
  }
}

async function runConsolidation(): Promise<void> {
  if (isRunning) {
    return;
  }

  isRunning = true;
  refreshUiState();
  setStatus("info", "Validating source mappings...");

  try {
    const preparedSources = prepareSourcesForConsolidation();
    setStatus("info", "Consolidating data...");
    const rows = buildConsolidatedRows(preparedSources);

    const worksheet = XLSX.utils.json_to_sheet(rows, { header: OUTPUT_HEADERS });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Consolidated");
    const outputBytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
    const outputBlob = new Blob([outputBytes], { type: OUTPUT_FILE_TYPE });
    const normalizedName = normalizeOutputFileName(outputFileName);

    await saveOutputFile(outputBlob, normalizedName);
    setStatus(
      "success",
      `Consolidation complete. ${rows.length} rows written to ${normalizedName}.`
    );
  } catch (error) {
    setStatus("error", toErrorMessage(error));
  } finally {
    isRunning = false;
    refreshUiState();
  }
}

function prepareSourcesForConsolidation(): PreparedSource[] {
  if (sources.length < 1) {
    throw new Error("At least one source block is required.");
  }

  const preparedSources: PreparedSource[] = [];

  sources.forEach((source, index) => {
    const sourceNumber = index + 1;

    if (!source.file || !source.workbook) {
      throw new Error(`File ${sourceNumber}: Source file is required. Please select a .xlsx file.`);
    }

    if (!source.file.name.toLowerCase().endsWith(".xlsx")) {
      throw new Error(`File ${sourceNumber}: Only .xlsx source files are supported.`);
    }

    if (!source.sourceSheet) {
      throw new Error(`File ${sourceNumber}: Source sheet is required.`);
    }

    const worksheet = source.workbook.Sheets[source.sourceSheet];
    if (!worksheet) {
      throw new Error(
        `File ${sourceNumber}: Sheet "${source.sourceSheet}" was not found in selected file.`
      );
    }

    const accountMatrix = readRangeMatrix(
      worksheet,
      source.accountRange,
      `File ${sourceNumber} Account range`
    );
    const valueMatrix = readRangeMatrix(
      worksheet,
      source.valueRange,
      `File ${sourceNumber} Value range`
    );

    if (accountMatrix.cols !== 1) {
      throw new Error(`File ${sourceNumber}: Account range must be a single column.`);
    }

    if (accountMatrix.rows !== valueMatrix.rows) {
      throw new Error(
        `File ${sourceNumber}: Account and Value ranges must have the same row count.`
      );
    }

    const entity = prepareEntityOrDepartmentDimension(
      source.entityMode,
      source.entityConstant,
      source.entityRange,
      worksheet,
      valueMatrix,
      `File ${sourceNumber} Entity`
    );

    const department = prepareEntityOrDepartmentDimension(
      source.departmentMode,
      source.departmentConstant,
      source.departmentRange,
      worksheet,
      valueMatrix,
      `File ${sourceNumber} Department`
    );

    const date = prepareDateDimension(
      source.dateMode,
      source.dateConstant,
      source.dateRange,
      worksheet,
      valueMatrix,
      `File ${sourceNumber} Date`
    );

    const multiColumnCount = [entity, department, date].filter(
      (item) => item.shape === "multiColumn"
    ).length;

    if (valueMatrix.cols > 1 && multiColumnCount !== 1) {
      throw new Error(
        `File ${sourceNumber}: When Value has multiple columns, exactly one of Entity/Department/Date must be multi-column (1 row x N columns).`
      );
    }

    if (valueMatrix.cols === 1 && multiColumnCount > 0) {
      throw new Error(
        `File ${sourceNumber}: Multi-column Entity/Department/Date is not allowed when Value has a single column.`
      );
    }

    preparedSources.push({
      source,
      accountMatrix,
      valueMatrix,
      entity,
      department,
      date,
    });
  });

  return preparedSources;
}

function prepareEntityOrDepartmentDimension(
  mode: EntityDepartmentMode,
  constantValue: string,
  rangeText: string,
  worksheet: XLSX.WorkSheet,
  valueMatrix: MatrixData,
  label: string
): PreparedDimension {
  if (mode === "blank") {
    return {
      shape: "blank",
      constantValue: "",
      matrix: null,
    };
  }

  if (mode === "constant") {
    return {
      shape: "constant",
      constantValue: trimText(constantValue),
      matrix: null,
    };
  }

  const matrix = readRangeMatrix(worksheet, rangeText, `${label} range`);
  return classifyRangeMatrix(matrix, valueMatrix, label);
}

function prepareDateDimension(
  mode: DateMode,
  constantValue: string,
  rangeText: string,
  worksheet: XLSX.WorkSheet,
  valueMatrix: MatrixData,
  label: string
): PreparedDimension {
  if (mode === "constant") {
    return {
      shape: "constant",
      constantValue: constantValue || "",
      matrix: null,
    };
  }

  const matrix = readRangeMatrix(worksheet, rangeText, `${label} range`);
  return classifyRangeMatrix(matrix, valueMatrix, label);
}

function classifyRangeMatrix(
  matrix: MatrixData,
  valueMatrix: MatrixData,
  label: string
): PreparedDimension {
  if (matrix.rows === 1 && matrix.cols === 1) {
    return {
      shape: "singleCell",
      constantValue: null,
      matrix,
    };
  }

  if (matrix.cols === 1 && matrix.rows === valueMatrix.rows) {
    return {
      shape: "singleColumn",
      constantValue: null,
      matrix,
    };
  }

  if (valueMatrix.cols > 1 && matrix.rows === 1 && matrix.cols === valueMatrix.cols) {
    return {
      shape: "multiColumn",
      constantValue: null,
      matrix,
    };
  }

  throw new Error(
    `${label} is incompatible with Value range shape. Use constant, blank, single-column with same rows, or 1 row x N columns matching Value columns.`
  );
}

function buildConsolidatedRows(preparedSources: PreparedSource[]): ConsolidatedRow[] {
  const rows: ConsolidatedRow[] = [];

  preparedSources.forEach((prepared) => {
    const sourceFile = prepared.source.sourcePath
      ? prepared.source.sourcePath
      : prepared.source.file
        ? prepared.source.file.name
        : "";
    const sourceSheet = prepared.source.sourceSheet;

    for (let rowIndex = 0; rowIndex < prepared.valueMatrix.rows; rowIndex += 1) {
      const accountRaw = prepared.accountMatrix.values[rowIndex][0];
      const account = trimText(accountRaw);

      if (!account) {
        continue;
      }

      for (let colIndex = 0; colIndex < prepared.valueMatrix.cols; colIndex += 1) {
        const valueRaw = prepared.valueMatrix.values[rowIndex][colIndex];
        const valueNumber = parseNumericValue(valueRaw);

        if (valueNumber === null) {
          continue;
        }

        const entityValue = trimText(resolveDimensionValue(prepared.entity, rowIndex, colIndex));
        const departmentValue = trimText(
          resolveDimensionValue(prepared.department, rowIndex, colIndex)
        );
        const dateValue = resolveDimensionValue(prepared.date, rowIndex, colIndex);

        rows.push({
          Account: account,
          Entity: entityValue,
          Department: departmentValue,
          Date: dateValue,
          Value: valueNumber,
          SourceFile: sourceFile,
          SourceSheet: sourceSheet,
        });
      }
    }
  });

  return rows;
}

function resolveDimensionValue(
  dimension: PreparedDimension,
  rowIndex: number,
  colIndex: number
): unknown {
  if (dimension.shape === "blank") {
    return "";
  }

  if (dimension.shape === "constant") {
    return dimension.constantValue;
  }

  if (!dimension.matrix) {
    return "";
  }

  if (dimension.shape === "singleCell") {
    return dimension.matrix.values[0][0];
  }

  if (dimension.shape === "singleColumn") {
    return dimension.matrix.values[rowIndex][0];
  }

  return dimension.matrix.values[0][colIndex];
}

function parseNumericValue(value: unknown): number | null {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== "string") {
    return null;
  }

  let textValue = value.trim();
  if (!textValue || textValue === "-") {
    return null;
  }

  let negative = false;
  if (textValue.startsWith("(") && textValue.endsWith(")")) {
    negative = true;
    textValue = textValue.slice(1, -1);
  }

  const cleaned = textValue.replace(/\$/g, "").replace(/,/g, "").replace(/\s+/g, "");
  if (!cleaned) {
    return null;
  }

  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) {
    return null;
  }

  return negative ? parsed * -1 : parsed;
}

function readRangeMatrix(worksheet: XLSX.WorkSheet, rangeInput: string, label: string): MatrixData {
  const normalizedRange = normalizeRangeInput(rangeInput);

  if (!normalizedRange) {
    throw new Error(`${label} is required.`);
  }

  try {
    const range = XLSX.utils.decode_range(normalizedRange);
    const rows = range.e.r - range.s.r + 1;
    const cols = range.e.c - range.s.c + 1;

    if (rows < 1 || cols < 1) {
      throw new Error("Range must include at least one cell.");
    }

    const values: unknown[][] = [];
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      const rowValues: unknown[] = [];
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[address];
        rowValues.push(cell ? cell.v : null);
      }
      values.push(rowValues);
    }

    return {
      rows,
      cols,
      values,
    };
  } catch {
    throw new Error(`${label} is invalid. Please provide a valid A1 range like B10:B110.`);
  }
}

function normalizeRangeInput(input: string): string {
  let value = (input || "").trim().replace(/\$/g, "");
  if (!value) {
    return "";
  }

  const bangIndex = value.lastIndexOf("!");
  if (bangIndex >= 0) {
    value = value.slice(bangIndex + 1);
  }

  value = value.toUpperCase();

  const parts = value.split(":");
  if (parts.length === 1) {
    if (!isCellAddress(parts[0])) {
      return "";
    }
    return `${parts[0]}:${parts[0]}`;
  }

  if (parts.length === 2 && isCellAddress(parts[0]) && isCellAddress(parts[1])) {
    return `${parts[0]}:${parts[1]}`;
  }

  return "";
}

function isCellAddress(value: string): boolean {
  return /^[A-Z]+[1-9][0-9]*$/.test(value);
}

function normalizeOutputFileName(fileName: string): string {
  const trimmed = (fileName || "").trim();
  if (!trimmed) {
    return "consolidated-actuals.xlsx";
  }

  return trimmed.toLowerCase().endsWith(".xlsx") ? trimmed : `${trimmed}.xlsx`;
}

function buildConfigFileName(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = pad2(now.getMonth() + 1);
  const day = pad2(now.getDate());
  const hours = pad2(now.getHours());
  const minutes = pad2(now.getMinutes());
  const seconds = pad2(now.getSeconds());
  return `xf1-config-${year}${month}${day}-${hours}${minutes}${seconds}${CONFIG_FILE_EXTENSION}`;
}

function pad2(value: number): string {
  return value < 10 ? `0${value}` : `${value}`;
}

async function saveOutputFile(blob: Blob, fileName: string): Promise<void> {
  if (outputFileHandle) {
    const writable = await outputFileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return;
  }

  downloadBlob(blob, fileName);
}

function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

function refreshUiState(): void {
  sourceCountLabelElement.textContent = `${sources.length}/${MAX_SOURCES} source files configured`;
  addSourceButtonElement.disabled = sources.length >= MAX_SOURCES || isRunning;
  runButtonElement.disabled = isRunning;
}

function setStatus(kind: StatusKind, message: string): void {
  statusElement.classList.remove("status-info", "status-success", "status-error");
  if (kind === "info") {
    statusElement.classList.add("status-info");
  } else if (kind === "success") {
    statusElement.classList.add("status-success");
  } else {
    statusElement.classList.add("status-error");
  }
  statusElement.textContent = message;
}

function trimText(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Required element not found: ${id}`);
  }
  return element as T;
}
