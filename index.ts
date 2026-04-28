
/// <reference types="powerapps-component-framework" />
import { IInputs, IOutputs } from "./generated/ManifestTypes";
interface UnitConversion {
  FromUnit: string;
  MultiplyBy: string;
  ToUnit: string;
}
interface RowItem {
  Code: string;
  DataID?: string | null;
  PreFilled?: number | string | null;
  FormulaText?: string | null;
  selectedUnit?: string | null;
  categoryUnits?: UnitConversion[] | null;
  Previousvalue?: number | string | null;
  Value?: number | string | null;
  Type?: string | number | null;
  [key: string]: unknown;
}

interface DefaultValueItem {
  Code?: string | null;
  IndicatorID?: string | null;
  DataID?: string | null;
  Value?: number | string | null;
}

export class LiveCalculationComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  // Host and output callback from PCF runtime.
  private container!: HTMLDivElement;
  private notifyOutputChanged!: () => void;

  // Last input snapshots used to avoid unnecessary full re-renders.
  private lastJson: string | null = null;
  private lastDefaultValue: string | null = null;
  private lastGapPercentThreshold: number = 50;
  private lastShowValidation: boolean = false;

  // Runtime maps for row inputs, units, computed ratios, and stable row keys.
  private _allItems: RowItem[] = [];
  private _rowInputMap: Map<string, HTMLInputElement> = new Map();
  private _rowUnitMap: Map<string, HTMLSelectElement> = new Map();
  private _rowRatioMap: Map<string, number | null> = new Map();
  private _rowKeyByItem: WeakMap<RowItem, string> = new WeakMap();

  // Deferred recompute callbacks for formula rows and validation UI.
  private _dependentComputes: Array<() => void> = [];
  private _validationUiRefreshers: Array<() => void> = [];

  // Output properties exposed back to Canvas.
  private _outputData?: string;
  private _hasValidationError: boolean = false;
  private _validationErrorRows: string = "[]";
  private _gapPercentThreshold: number = 50;
  private _showValidation: boolean = false;
  private _showCommentPopup: boolean = false;
  private _commentPopupData: string = "";
  private _totalItemsCount: number = 0;
  private _validationErrorCount: number = 0;
  private _validationStatusSummary: string = "";

  private normalizeCodeKey(code: string | null | undefined): string {
    return (code ?? "").trim().toLowerCase();
  }

  // Build a stable row key to prevent collisions when codes repeat across categories.
  private getRowKey(item: RowItem): string {
    const existing = this._rowKeyByItem.get(item);
    if (existing) {
      return existing;
    }

    const dataId = typeof item.DataID === "string" ? item.DataID.trim().toLowerCase() : "";
    const itemRec = item as Record<string, unknown>;
    const rawCategory = this.getFieldValue(
      itemRec,
      ["Category", "category", "CategoryCode", "categoryCode", "CategoryName", "categoryName"],
      true
    );
    const category = typeof rawCategory === "string" ? rawCategory.trim().toLowerCase() : "";
    const code = this.normalizeCodeKey(item.Code);

    if (dataId && category) {
      return `dataid:${dataId}|category:${category}|code:${code}`;
    }
    if (dataId) {
      return `dataid:${dataId}|code:${code}`;
    }
    return category ? `category:${category}|code:${code}` : `code:${code}`;
  }

  private findLiveInputByCode(codeKey: string): HTMLInputElement | undefined {
    for (const row of this._allItems) {
      if (this.normalizeCodeKey(row.Code) !== codeKey) {
        continue;
      }
      const key = this.getRowKey(row);
      const input = this._rowInputMap.get(key);
      if (input) {
        return input;
      }
    }
    return undefined;
  }

  private findLiveUnitSelectByCode(codeKey: string): HTMLSelectElement | undefined {
    for (const row of this._allItems) {
      if (this.normalizeCodeKey(row.Code) !== codeKey) {
        continue;
      }
      const key = this.getRowKey(row);
      const select = this._rowUnitMap.get(key);
      if (select) {
        return select;
      }
    }
    return undefined;
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;
    this.render(context);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    const raw = context.parameters.indicesJson.raw ?? "";
    const defaultValue = context.parameters.DefaultValue.raw ?? null;
    const gapPercentThreshold = this.resolveGapPercentThreshold(context.parameters.GapPercentThreshold.raw);
    const showValidation = context.parameters.ShowValidation.raw === true;
    const dataChanged =
      raw !== this.lastJson ||
      defaultValue !== this.lastDefaultValue ||
      gapPercentThreshold !== this.lastGapPercentThreshold;
    const showValidationChanged = showValidation !== this.lastShowValidation;

    if (dataChanged) {
      this.render(context);
      return;
    }

    if (showValidationChanged) {
      this.lastShowValidation = showValidation;
      this._showValidation = showValidation;
      this.recomputeValidationState();
      this._validationUiRefreshers.forEach(refresh => refresh());
      this.notifyOutputChanged();
    }
  }

  public getOutputs(): IOutputs {
    return {
      OutputData: this._outputData,
      HasValidationError: this._hasValidationError,
      ValidationErrorRows: this._validationErrorRows,
      ShowCommentPopup: this._showCommentPopup,
      CommentPopupData: this._commentPopupData,
      ValidationStatusSummary: this._validationStatusSummary
    };
  }

  // Serialize row output and recompute aggregated validation counters.
  private buildOutputData(): string {
    const rows: Array<{
      Code: string;
      DataID: string | null;
      Value: number | null;
      selectedUnit: string | null;
      Ratio: number | null;
    }> = [];
    const validationErrorRows: Array<{
      Code: string;
      DataID: string | null;
      Reason: string;
    }> = [];


    for (const item of this._allItems) {
      const key = this.getRowKey(item);
      const input = this._rowInputMap.get(key);
      const select = this._rowUnitMap.get(key);
      const ratio = this._rowRatioMap.has(key) ? this._rowRatioMap.get(key) : null;
      const valueText = input?.value ?? "";
      const value = input ? this.parseNumber(valueText) : null;

      rows.push({
        Code: item.Code,
        DataID: item.DataID ?? null,
        Value: value,
        selectedUnit: select ? select.value : (item.selectedUnit ?? null),
        Ratio: ratio ?? null
      });

      const rowReason = this.getValueValidationReason(
        this.getInputMode(item) === 1,
        valueText,
        this.getPreviousYearValue(item)
      );
      if (rowReason) {
        validationErrorRows.push({
          Code: item.Code,
          DataID: item.DataID ?? null,
          Reason: rowReason
        });
      }
    }

    const valueRequiredCount = validationErrorRows.filter(r => r.Reason === "Value is required").length;
    const justificationRequiredCount = validationErrorRows.filter(r => r.Reason === "Justification required").length;
    this._totalItemsCount = valueRequiredCount;
    this._hasValidationError = validationErrorRows.length > 0;
    this._validationErrorCount = justificationRequiredCount;
    this._validationStatusSummary = `Value is required: ${this._totalItemsCount} | Justification required: ${this._validationErrorCount}`;
    this._validationErrorRows = JSON.stringify(validationErrorRows);
    return JSON.stringify(rows);
  }

  private recomputeValidationState(): void {
    const validationErrorRows: Array<{
      Code: string;
      DataID: string | null;
      Reason: string;
    }> = [];

    for (const item of this._allItems) {
      const key = this.getRowKey(item);
      const input = this._rowInputMap.get(key);
      const valueText = input?.value ?? "";
      const editableInput = this.getInputMode(item) === 1;
      const prevYearValue = this.getPreviousYearValue(item);

      const rowReason = this.getValueValidationReason(editableInput, valueText, prevYearValue);

      if (rowReason) {
        validationErrorRows.push({
          Code: item.Code,
          DataID: item.DataID ?? null,
          Reason: rowReason
        });
      }
    }

    const valueRequiredCount = validationErrorRows.filter(r => r.Reason === "Value is required").length;
    const justificationRequiredCount = validationErrorRows.filter(r => r.Reason === "Justification required").length;
    this._totalItemsCount = valueRequiredCount;
    this._hasValidationError = validationErrorRows.length > 0;
    this._validationErrorCount = justificationRequiredCount;
    this._validationStatusSummary = `Value is required: ${this._totalItemsCount} | Justification required: ${this._validationErrorCount}`;
    this._validationErrorRows = JSON.stringify(validationErrorRows);
  }

  public destroy(): void {
    while (this.container && this.container.firstChild) {
      this.container.removeChild(this.container.firstChild);
    }
  }

  // ===================== RENDER =====================
  private render(context: ComponentFramework.Context<IInputs>): void {
    const raw = context.parameters.indicesJson.raw ?? "";
    this.lastJson = raw;
    const defaultValue = context.parameters.DefaultValue.raw ?? null;
    this.lastDefaultValue = defaultValue;
    const gapPercentThreshold = this.resolveGapPercentThreshold(context.parameters.GapPercentThreshold.raw);
    this.lastGapPercentThreshold = gapPercentThreshold;
    this._gapPercentThreshold = gapPercentThreshold;
    const showValidation = context.parameters.ShowValidation.raw === true;
    this.lastShowValidation = showValidation;
    this._showValidation = showValidation;

    this.container.innerHTML = "";
    this.injectStyles(this.container);

    const wrapper = document.createElement("div");
    wrapper.className = "pcf-table";

    wrapper.appendChild(this.renderHeader());

    const items = this.safeParse(raw);
    this._allItems = items; // Store for cross-row formula resolution.
    this._rowKeyByItem = new WeakMap();
    this._allItems.forEach((item, index) => {
      const dataId = typeof item.DataID === "string" ? item.DataID.trim().toLowerCase() : "";
      const itemRec = item as Record<string, unknown>;
      const rawCategory = this.getFieldValue(
        itemRec,
        ["Category", "category", "CategoryCode", "categoryCode", "CategoryName", "categoryName"],
        true
      );
      const category = typeof rawCategory === "string" ? rawCategory.trim().toLowerCase() : "";
      const code = this.normalizeCodeKey(item.Code);
      const stableKey = `idx:${index}|dataid:${dataId || "_"}|category:${category || "_"}|code:${code || "_"}`;
      this._rowKeyByItem.set(item, stableKey);
    });
    this._rowInputMap.clear(); // Reset live input map on each render.
    this._rowUnitMap.clear(); // Reset live unit select map on each render.
    this._rowRatioMap.clear(); // Reset ratio map on each render.
    this._dependentComputes = []; // Reset dependent compute callbacks.
    this._validationUiRefreshers = []; // Reset validation UI callbacks.
    const parsedDefaultValue = defaultValue !== null ? this.parseNumber(defaultValue) : null;
    const perRowDefaults = this.parseDefaultValueMap(defaultValue);
    if (items.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pcf-empty";
      empty.innerText = "No data";
      wrapper.appendChild(empty);
      this.container.appendChild(wrapper);
      return;
    }

    const sortedItems = [...items].sort((a, b) => this.getInputMode(a) - this.getInputMode(b));

    for (const item of sortedItems) {
      wrapper.appendChild(this.renderRow(item, parsedDefaultValue, perRowDefaults));
    }

    this.container.appendChild(wrapper);

    // Defer final notification so it fires after init() completes.
    // notifyOutputChanged() is ignored when called synchronously during init().
    setTimeout(() => {
      this._outputData = this.buildOutputData();
      this.notifyOutputChanged();
    }, 0);
  }

  private renderHeader(): HTMLDivElement {
    const currentYear = this.getCurrentYear();
    const previousYear = currentYear - 1;

    const header = document.createElement("div");
    header.className = "pcf-row header";
    header.innerHTML = `
      <div>Indicator</div>
      <div>Unit</div>
      <div>${previousYear}</div>
      <div>Pre-filled</div>
      <div>${currentYear}</div>
      <div>% N/N‑1</div>
      <div>Actions</div>
  `;
    return header;
  }

  // Render one row with units, values, validation, formula and comment action.
  private renderRow(
    item: RowItem,
    parsedDefaultValue: number | null,
    perRowDefaults: Map<string, number>
  ): HTMLDivElement {
    const row = document.createElement("div");
    row.className = "pcf-row";

    // INDICATEURS
    const colIndicator = document.createElement("div");
    colIndicator.className = "indicator";
    colIndicator.innerHTML = `
      <div class="title">${this.escape(item.Code)}</div>
     
    `;

    // Unit selector for the row. For editable rows, unit changes are propagated per category.
    const colUnit = document.createElement("div");
    const unitSelect = document.createElement("select");
    unitSelect.className = "unitSelect";

    const inputMode = this.getInputMode(item);
    const editableInput = inputMode === 1;
    unitSelect.disabled = !editableInput;

    const itemRec = item as Record<string, unknown>;
    const rawUnits =
      this.getFieldValue(itemRec, ["categoryUnits", "Units", "UnitConversions", "UnitsJson"], true) ??
      item.categoryUnits;
    const toUnits = this.extractUnitOptions(rawUnits);
    const selectedUnitRaw =
      this.getFieldValue(itemRec, ["selectedUnit", "SelectedUnit", "Unit"], true) ??
      item.selectedUnit;
    const selectedUnit = this.normalizeUnit(typeof selectedUnitRaw === "string" ? selectedUnitRaw : null);
    const baseUnit = this.getBaseUnit(item);
    const rowCodeKey = this.getRowKey(item);
    const categoryKey = this.getCategoryKey(item, toUnits);

    if (toUnits.length === 0) {
      this.addOption(unitSelect, "No unit", "No unit");
    } else {
      const matchedSelectedUnit = this.findMatchingUnit(toUnits, selectedUnit);
      if (selectedUnit && !matchedSelectedUnit) {
        toUnits.unshift(selectedUnit);
      }
      toUnits.forEach(u => this.addOption(unitSelect, u, u));
      const finalSelected = this.findMatchingUnit(toUnits, selectedUnit) ?? toUnits[0];
      unitSelect.value = finalSelected;
      const selectedIndex = toUnits.findIndex(u => this.areUnitsEqual(u, finalSelected));
      unitSelect.selectedIndex = selectedIndex >= 0 ? selectedIndex : 0;
    }

    colUnit.appendChild(unitSelect);
    if (rowCodeKey) {
      this._rowUnitMap.set(rowCodeKey, unitSelect);
    }

    if (editableInput) {
      unitSelect.addEventListener("change", () => {
        const changedUnit = this.normalizeUnit(unitSelect.value);
        if (!changedUnit) {
          return;
        }
        this.applyUnitToCategory(categoryKey, changedUnit);
      });
    }

    const currentYear = this.getCurrentYear();
    const previousYear = currentYear - 1;

    // Previous-year static value
    const colPrevYear = document.createElement("div");
    const prevYearValue = this.getPreviousYearValue(item);
    const basePrevYear = Number(prevYearValue ?? 0);
    colPrevYear.innerText =
      prevYearValue === null ? "-" : this.formatNumber(basePrevYear);

    // Pre-filled static value from JSON
    const colPreFilled = document.createElement("div");
    const preFilledValue = this.getPreFilledValue(item);
    colPreFilled.innerText =
      preFilledValue === null ? "-" : this.formatNumber(preFilledValue);

    // Current-year input (single input)
    const colCurrentYear = document.createElement("div");
    colCurrentYear.className = "valueCell";
    const inputCurrentYear = document.createElement("input");
    inputCurrentYear.type = "text";
    inputCurrentYear.className = "valueInput";
    const valueError = document.createElement("div");
    valueError.className = "valueError";
    valueError.innerText = "";
    const currentYearValue = this.getCurrentYearValue(item);
    const rowDefaultValue = this.getPerRowDefaultValue(item, perRowDefaults);
    const initialEditableValue =
      currentYearValue ??
      (editableInput ? (rowDefaultValue ?? parsedDefaultValue) : null);

    inputCurrentYear.placeholder = editableInput ? "Enter data" : "No data";
    inputCurrentYear.readOnly = !editableInput;
    if (!editableInput) {
      inputCurrentYear.classList.add("valueInputReadonly");
    }

    if (initialEditableValue !== null) {
      const formattedInitialValue = this.formatNumber(initialEditableValue);
      inputCurrentYear.value = formattedInitialValue;
      inputCurrentYear.defaultValue = formattedInitialValue;
    }
    colCurrentYear.appendChild(inputCurrentYear);
    colCurrentYear.appendChild(valueError);
    const itemCodeKey = this.getRowKey(item);
    const updateValueValidationUi = (): void => {
      const reason = this.getValueValidationReason(editableInput, inputCurrentYear.value, prevYearValue);
      const shouldShowError = !!reason && this._showValidation;
      valueError.innerText = shouldShowError ? reason! : "";
      valueError.style.display = shouldShowError ? "block" : "none";
      inputCurrentYear.classList.toggle("valueInputError", shouldShowError);
    };
    this._validationUiRefreshers.push(updateValueValidationUi);

    if (itemCodeKey) {
      this._rowInputMap.set(itemCodeKey, inputCurrentYear);
    }

    // % N/N‑1 column
    const colPct = document.createElement("div");
    colPct.className = "pctCell";
    colPct.innerText = "-";

    // ACTIONS
    const colActions = document.createElement("div");
    colActions.className = "actions";
    colActions.innerHTML = `
      <span class="action" title="Comment" aria-label="Comment">
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
          <path d="M4 5h16v11H9l-5 4V5z" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>
        </svg>
      </span>
    `;
    const commentAction = colActions.querySelector(".action");
    if (commentAction) {
      commentAction.addEventListener("click", (evt) => {
        evt.preventDefault();
        this._showCommentPopup = !this._showCommentPopup;
        this._commentPopupData = JSON.stringify({
          Code: item.Code,
          DataID: item.DataID ?? null,
          Timestamp: Date.now()
        });
        this.notifyOutputChanged();
      });
    }

    // -------- compute when user types current-year value --------
    const compute = (): void => {
      if (!editableInput) {
        // For readonly rows, evaluate the formula (e.g. cross-row code references)
        // and display the result directly in the input field.
        const convertedFormula = this.convertFormulaTextToFormula(item.FormulaText);
        if (convertedFormula) {
          const x = basePrevYear;
          const yVal = currentYearValue !== null ? currentYearValue : 0;
          const result = this.evalFormula(convertedFormula, x, yVal);
          if (result !== null && !Number.isNaN(result)) {
            const r = this.round(result, 6);
            inputCurrentYear.value = this.formatNumber(r);
            if (prevYearValue !== null && basePrevYear !== 0) {
              const pct = ((r - basePrevYear) / basePrevYear) * 100;
              const pctRounded = this.round(pct, 2);
              colPct.innerText = `${pctRounded}%`;
              this._rowRatioMap.set(rowCodeKey, pctRounded);
            } else {
              colPct.innerText = "-";
              this._rowRatioMap.set(rowCodeKey, null);
            }
            this._outputData = this.buildOutputData();
            this.notifyOutputChanged();
            return;
          }
        }
        colPct.innerText = "-";
        this._rowRatioMap.set(rowCodeKey, null);
        this._outputData = this.buildOutputData();
        this.notifyOutputChanged();
        return;
      }

      const yRaw = inputCurrentYear.value;

      // If current year is empty -> show '-' (no fake 0)
      if (!yRaw || yRaw.trim() === "") {
        colPct.innerText = "-";
        this._rowRatioMap.set(rowCodeKey, null);
        updateValueValidationUi();
        this._outputData = this.buildOutputData();
        this.notifyOutputChanged();
        return;
      }

      const x = basePrevYear;                 // Previous year baseline is X
      const y = this.parseNumber(yRaw);       // Current year typed value is Y
      if (y === null) {
        colPct.innerText = "-";
        this._rowRatioMap.set(rowCodeKey, null);
        updateValueValidationUi();
        this._outputData = this.buildOutputData();
        this.notifyOutputChanged();
        return;
      }

      const selectedDisplayUnit = this.normalizeUnit(unitSelect.value);
      const yForFormula = this.convertInputFromBaseToSelectedUnit(item, y, selectedDisplayUnit, baseUnit);
      if (yForFormula === null) {
        colPct.innerText = "-";
        this._rowRatioMap.set(rowCodeKey, null);
        updateValueValidationUi();
        this._outputData = this.buildOutputData();
        this.notifyOutputChanged();
        return;
      }

      if (prevYearValue === null || basePrevYear === 0) {
        colPct.innerText = "-";
        this._rowRatioMap.set(rowCodeKey, null);
      } else {
        // For editable rows, ratio is always based on current year input vs previous year value.
        const r = this.round(yForFormula, 6);
        const pct = ((r - basePrevYear) / basePrevYear) * 100;
        const pctRounded = this.round(pct, 2);
        colPct.innerText = `${pctRounded}%`;
        this._rowRatioMap.set(rowCodeKey, pctRounded);
      }

      this._outputData = this.buildOutputData();
      updateValueValidationUi();
      this.notifyOutputChanged();
    };

    inputCurrentYear.addEventListener("input", compute);
    if (!editableInput && item.FormulaText) {
      // Register this row so editable rows can trigger its recomputation
      this._dependentComputes.push(compute);
    }

    if (editableInput) {
      // When an editable row changes, re-run all readonly computed rows
      inputCurrentYear.addEventListener("input", () => {
        this._dependentComputes.forEach(fn => fn());
      });

      // Recompute when display unit changes because input is converted before formula evaluation.
      unitSelect.addEventListener("change", () => {
        compute();
        this._dependentComputes.forEach(fn => fn());
      });
    }

    compute(); // initialize
    updateValueValidationUi();

    row.append(colIndicator, colUnit, colPrevYear, colPreFilled, colCurrentYear, colPct, colActions);
    return row;
  }

  private getCurrentYear(): number {
    return new Date().getFullYear();
  }

  // Type 1 => editable with "Enter data", Type 2 => readonly with "No data".
  // Any missing/unknown type defaults to display-only for safety.
  private getInputMode(item: RowItem): 1 | 2 {
    const itemRec = item as Record<string, unknown>;
    const rawType =
      this.getFieldValue(itemRec, ["Type", "InputMode", "Mode"], true) ??
      item.Type;

    if (typeof rawType === "number") {
      return rawType === 1 ? 1 : 2;
    }

    if (typeof rawType === "string") {
      const normalized = rawType.trim().toLowerCase();
      if (normalized === "1" || normalized === "type1" || normalized === "editable") {
        return 1;
      }
      if (
        normalized === "2" ||
        normalized === "type2" ||
        normalized === "readonly" ||
        normalized === "read-only"
      ) {
        return 2;
      }
    }

    if (typeof rawType === "boolean") {
      return rawType ? 1 : 2;
    }

    return 2;
  }

  // Read N-1 value from common legacy and current payload keys.
  private getPreviousYearValue(item: RowItem): number | null {
    const previousYear = this.getCurrentYear() - 1;
    const record = item as Record<string, unknown>;
    const candidates: unknown[] = [
      record["Previousvalue"],
      record["ValueYnMinus1"],
      record["ValueYn-1"],
      record[`Value${previousYear}`]
    ];

    return this.getFirstNumericValue(candidates);
  }

  private getPreFilledValue(item: RowItem): number | null {
    const record = item as Record<string, unknown>;
    const candidates: unknown[] = [
      this.getFieldValue(record, ["PreFilled", "Prefilled", "PreFilledValue", "DefaultValue"], true),
      item.PreFilled
    ];

    return this.getFirstNumericValue(candidates);
  }

  // Read the current year value from payload.
  private getCurrentYearValue(item: RowItem): number | null {
    const currentYear = this.getCurrentYear();
    const record = item as Record<string, unknown>;
    const candidates: unknown[] = [
      record["Value"],
      record[`Value${currentYear}`]
    ];

    return this.getFirstNumericValue(candidates);
  }

  // Return the first numeric candidate from mixed unknown/string values.
  private getFirstNumericValue(candidates: unknown[]): number | null {
    for (const candidate of candidates) {
      if (typeof candidate === "number" && Number.isFinite(candidate)) {
        return candidate;
      }

      if (typeof candidate === "string") {
        if (candidate.trim().toLowerCase() === "null" || candidate.trim() === "") {
          continue;
        }

        const parsed = this.parseNumber(candidate);
        if (parsed !== null) {
          return parsed;
        }
      }
    }

    return null;
  }

  private normalizeFieldKey(key: string): string {
    return key.replace(/[^a-z0-9]/gi, "").toLowerCase();
  }

  // Flexible key resolver to support different payload naming conventions.
  private getFieldValue(
    record: Record<string, unknown>,
    candidateKeys: string[],
    allowContains: boolean = false
  ): unknown {
    for (const key of candidateKeys) {
      if (key in record) {
        return record[key];
      }
    }

    const normalizedCandidates = candidateKeys.map(k => this.normalizeFieldKey(k));
    for (const [rawKey, value] of Object.entries(record)) {
      const normalizedKey = this.normalizeFieldKey(rawKey);
      for (const candidate of normalizedCandidates) {
        if (normalizedKey === candidate || normalizedKey.endsWith(candidate)) {
          return value;
        }
        if (allowContains && normalizedKey.includes(candidate)) {
          return value;
        }
      }
    }

    return undefined;
  }

  // Extract distinct unit options from strings, objects or JSON arrays.
  private extractUnitOptions(rawUnits: unknown): string[] {
    const values: string[] = [];

    const addUnit = (value: unknown): void => {
      if (typeof value !== "string") {
        return;
      }
      const normalized = value.trim();
      if (!normalized || normalized.toLowerCase() === "null") {
        return;
      }
      values.push(normalized);
    };

    const readUnitFromObject = (obj: Record<string, unknown>): void => {
      const unitValue = this.getFieldValue(obj, ["ToUnit", "Unit", "FromUnit", "Name"], true);
      addUnit(unitValue);
    };

    if (typeof rawUnits === "string") {
      const trimmed = rawUnits.trim();
      if (trimmed.startsWith("[") || trimmed.startsWith("{")) {
        try {
          return this.extractUnitOptions(JSON.parse(trimmed));
        } catch {
          addUnit(trimmed);
        }
      } else {
        addUnit(trimmed);
      }
    } else if (Array.isArray(rawUnits)) {
      for (const entry of rawUnits) {
        if (typeof entry === "string") {
          addUnit(entry);
          continue;
        }
        if (entry && typeof entry === "object") {
          readUnitFromObject(entry as Record<string, unknown>);
        }
      }
    } else if (rawUnits && typeof rawUnits === "object") {
      readUnitFromObject(rawUnits as Record<string, unknown>);
    }

    return Array.from(new Set(values));
  }

  private normalizeUnit(unit: string | null | undefined): string | null {
    if (typeof unit !== "string") {
      return null;
    }

    const normalized = unit.trim();
    if (!normalized || normalized.toLowerCase() === "null") {
      return null;
    }

    return normalized;
  }

  private areUnitsEqual(left: string | null | undefined, right: string | null | undefined): boolean {
    const l = this.normalizeUnit(left)?.toLowerCase();
    const r = this.normalizeUnit(right)?.toLowerCase();
    return !!l && !!r && l === r;
  }

  private findMatchingUnit(options: string[], target: string | null | undefined): string | null {
    if (!target) {
      return null;
    }

    for (const option of options) {
      if (this.areUnitsEqual(option, target)) {
        return option;
      }
    }

    return null;
  }

  // Use category as synchronization boundary for unit selection.
  private getCategoryKey(item: RowItem, resolvedUnits: string[]): string {
    const record = item as Record<string, unknown>;
    const rawCategory = this.getFieldValue(
      record,
      ["Category", "category", "CategoryCode", "categoryCode", "CategoryName", "categoryName"],
      true
    );

    if (typeof rawCategory === "string") {
      const normalizedCategory = rawCategory.trim();
      if (normalizedCategory.length > 0) {
        return `category:${normalizedCategory.toLowerCase()}`;
      }
    }

    if (resolvedUnits.length > 0) {
      const signature = resolvedUnits.map(u => u.toLowerCase()).sort().join("|");
      return `units:${signature}`;
    }

    return "default";
  }

  // When one unit changes, apply the same unit to rows in the same category.
  private applyUnitToCategory(categoryKey: string, selectedUnit: string): void {
    for (const row of this._allItems) {
      const rowRec = row as Record<string, unknown>;
      const rowUnitsRaw =
        this.getFieldValue(rowRec, ["categoryUnits", "Units", "UnitConversions", "UnitsJson"], true) ??
        row.categoryUnits;
      const rowUnits = this.extractUnitOptions(rowUnitsRaw);
      const rowCategoryKey = this.getCategoryKey(row, rowUnits);

      if (rowCategoryKey !== categoryKey) {
        continue;
      }

      row.selectedUnit = selectedUnit;
      rowRec["selectedUnit"] = selectedUnit;

      const rowCodeKey = this.getRowKey(row);
      const select = rowCodeKey ? this._rowUnitMap.get(rowCodeKey) : undefined;
      if (!select) {
        continue;
      }

      const existingOptions = Array.from(select.options).map(o => o.value);
      const matchedExisting = this.findMatchingUnit(existingOptions, selectedUnit);

      if (matchedExisting) {
        select.value = matchedExisting;
      } else {
        this.addOption(select, selectedUnit, selectedUnit);
        select.value = selectedUnit;
      }

      const selectedIndex = Array.from(select.options).findIndex(option =>
        this.areUnitsEqual(option.value, select.value)
      );
      select.selectedIndex = selectedIndex >= 0 ? selectedIndex : 0;
    }
  }

  // Resolve and cache the row base unit used for conversion.
  private getBaseUnit(item: RowItem): string | null {
    const itemRec = item as Record<string, unknown>;
    const persistedBase = this.normalizeUnit(
      typeof itemRec["__baseUnit"] === "string" ? (itemRec["__baseUnit"] as string) : null
    );
    if (persistedBase) {
      return persistedBase;
    }

    const rawBase = this.getFieldValue(itemRec, ["baseUnit", "BaseUnit", "selectedUnit", "SelectedUnit"], true);
    const base = this.normalizeUnit(typeof rawBase === "string" ? rawBase : null);
    if (base) {
      itemRec["__baseUnit"] = base;
    }
    return base;
  }

  // Convert a typed value from base unit into the selected display unit.
  private convertInputFromBaseToSelectedUnit(
    item: RowItem,
    inputValue: number,
    selectedDisplayUnit: string | null,
    baseUnit: string | null
  ): number | null {
    if (!Number.isFinite(inputValue)) {
      return null;
    }

    const fromUnit = this.normalizeUnit(baseUnit);
    const toUnit = this.normalizeUnit(selectedDisplayUnit);
    if (!fromUnit || !toUnit || this.areUnitsEqual(fromUnit, toUnit)) {
      return inputValue;
    }

    const itemRec = item as Record<string, unknown>;
    const rawUnits =
      this.getFieldValue(itemRec, ["categoryUnits", "Units", "UnitConversions", "UnitsJson"], true) ??
      item.categoryUnits;
    const factor = this.getConversionFactor(rawUnits, fromUnit, toUnit);
    if (factor === null) {
      return null;
    }

    return inputValue * factor;
  }

  // Build a graph from conversion rules and find factor via BFS.
  private getConversionFactor(rawUnits: unknown, fromUnit: string, toUnit: string): number | null {
    const from = this.normalizeUnit(fromUnit);
    const to = this.normalizeUnit(toUnit);
    if (!from || !to) {
      return null;
    }
    if (this.areUnitsEqual(from, to)) {
      return 1;
    }

    const edges = new Map<string, Array<{ to: string; factor: number }>>();
    const addEdge = (fromKey: string, toKey: string, factor: number): void => {
      if (!edges.has(fromKey)) {
        edges.set(fromKey, []);
      }
      edges.get(fromKey)!.push({ to: toKey, factor });
    };

    const list = Array.isArray(rawUnits) ? rawUnits : [];
    for (const entry of list) {
      if (!entry || typeof entry !== "object") {
        continue;
      }
      const rec = entry as Record<string, unknown>;
      const fromVal = this.normalizeUnit(
        typeof this.getFieldValue(rec, ["FromUnit", "fromUnit"], true) === "string"
          ? (this.getFieldValue(rec, ["FromUnit", "fromUnit"], true) as string)
          : null
      );
      const toVal = this.normalizeUnit(
        typeof this.getFieldValue(rec, ["ToUnit", "toUnit"], true) === "string"
          ? (this.getFieldValue(rec, ["ToUnit", "toUnit"], true) as string)
          : null
      );
      const multiplyRaw = this.getFieldValue(rec, ["MultiplyBy", "multiplyBy"], true);
      const factor = this.parseMultiplyFactor(typeof multiplyRaw === "string" ? multiplyRaw : null);

      if (!fromVal || !toVal || factor === null || factor === 0) {
        continue;
      }

      const fromKey = fromVal.toLowerCase();
      const toKey = toVal.toLowerCase();
      addEdge(fromKey, toKey, factor);
      addEdge(toKey, fromKey, 1 / factor);
    }

    const start = from.toLowerCase();
    const target = to.toLowerCase();
    const queue: Array<{ unit: string; factor: number }> = [{ unit: start, factor: 1 }];
    const visited = new Set<string>([start]);

    while (queue.length > 0) {
      const current = queue.shift();
      if (!current) {
        break;
      }
      if (current.unit === target) {
        return current.factor;
      }
      const neighbors = edges.get(current.unit) ?? [];
      for (const neighbor of neighbors) {
        if (visited.has(neighbor.to)) {
          continue;
        }
        visited.add(neighbor.to);
        queue.push({ unit: neighbor.to, factor: current.factor * neighbor.factor });
      }
    }

    return null;
  }

  private parseMultiplyFactor(raw: string | null | undefined): number | null {
    if (!raw) {
      return null;
    }

    let cleaned = raw
      .replace(/[×xX*]/g, "")
      .replace(/\s+/g, "")
      .trim();

    if (!cleaned) {
      return null;
    }

    const hasComma = cleaned.includes(",");
    const hasDot = cleaned.includes(".");
    if (hasComma && hasDot) {
      cleaned = cleaned.replace(/,/g, "");
    } else if (hasComma) {
      const thousandPattern = /^-?\d{1,3}(,\d{3})+$/;
      cleaned = thousandPattern.test(cleaned) ? cleaned.replace(/,/g, "") : cleaned.replace(/,/g, ".");
    }

    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  // ===================== FORMULA ENGINE =====================
  private cleanFormula(text: string | null | undefined): string {
    return (text ?? "").replace(/\s+/g, "").toLowerCase();
  }

  // Resolve row code references and normalize aliases (x/y) for evaluation.
  private convertFormulaTextToFormula(formulaText: string | null | undefined): string {
    const raw = (formulaText ?? "").trim();
    if (!raw || raw.toLowerCase() === "null") {
      return "";
    }

    // Normalize common symbols/aliases coming from Dataverse expressions.
    let converted = raw
      .replace(/[×]/g, "*")
      .replace(/[÷]/g, "/");

    // Resolve 'CODE' references → their current-year numeric value, e.g. 'WB-C-2' → 150
    // Resolve 'CODE' references → live input value of that row
    converted = converted.replace(/'([^']+)'/g, (_match, code: string) => {
      const codeKey = this.normalizeCodeKey(code);
      const refRow = this._allItems.find(r => this.normalizeCodeKey(r.Code) === codeKey);

      // Prefer live input value (typed or previously computed) over raw JSON value
      const inputEl = this.findLiveInputByCode(codeKey);
      if (inputEl && inputEl.value.trim() !== "") {
        const liveVal = this.parseNumber(inputEl.value);
        if (liveVal !== null) {
          const selectEl = this.findLiveUnitSelectByCode(codeKey);
          const selectedDisplayUnit = this.normalizeUnit(selectEl?.value ?? null);
          const baseUnit = refRow ? this.getBaseUnit(refRow) : null;
          const convertedLiveVal = refRow
            ? this.convertInputFromBaseToSelectedUnit(refRow, liveVal, selectedDisplayUnit, baseUnit)
            : liveVal;

          if (convertedLiveVal !== null) {
            return String(convertedLiveVal);
          }

          return String(liveVal);
        }
      }
      // Fallback to JSON value
      if (refRow) {
        const val = this.getCurrentYearValue(refRow);
        if (val !== null) {
          return String(val);
        }

        // If referenced code exists but has no current-year value yet,
        // keep formula unresolved instead of silently substituting 0.
        return "NaN";
      }

      // Unknown code reference should not produce a numeric result.
      return "NaN";
    });

    // Normalize x/y aliases (for rows that use previousvalue/value pattern)
    converted = converted
      .replace(/previousvalue|yn-1|n-1/gi, "x")
      .replace(/value|yn|\bn\b/gi, "y");

    converted = this.cleanFormula(converted);

    return converted;
  }

  // Evaluate a sanitized arithmetic expression containing x/y placeholders.
  private evalFormula(formula: string, x: number, y: number): number | null {
    // Accept only simple arithmetic expressions containing x/y (or a/b), numbers, and parentheses.
    const normalized = (formula ?? "").replace(/\s+/g, "").toLowerCase();
    if (!normalized) {
      return null;
    }

    if (!/^[0-9xyab+\-*/().]+$/.test(normalized)) {
      return null;
    }

    // Map legacy aliases a/b to x/y.
    const mapped = normalized.replace(/a/g, "x").replace(/b/g, "y");
    const expression = mapped
      .replace(/x/g, `(${x})`)
      .replace(/y/g, `(${y})`);

    try {
      const fn = new Function(`"use strict"; return (${expression});`);
      const result = Number(fn());

      if (!Number.isFinite(result)) {
        return null;
      }

      return result;
    } catch {
      return null;
    }
  }

  // ===================== JSON =====================
  // Parse indicesJson string payload safely.
  private safeParse(raw: string): RowItem[] {
    if (!raw || raw.trim().length < 2) return [];
    try {
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? (parsed as RowItem[]) : [];
    } catch {
      return [];
    }
  }

  // Parse DefaultValue payload into lookup map keyed by code and id.
  private parseDefaultValueMap(rawDefault: string | null): Map<string, number> {
    const defaults = new Map<string, number>();
    if (!rawDefault) {
      return defaults;
    }

    const trimmed = rawDefault.trim();
    let parsed: unknown = null;

    // Canvas can send JSON directly, or a quoted/escaped JSON string.
    if (trimmed.startsWith("[") || trimmed.startsWith("{")) {
      try {
        parsed = JSON.parse(trimmed);
      } catch {
        return defaults;
      }
    } else if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
      const unwrapped = trimmed.substring(1, trimmed.length - 1);
      const candidate = unwrapped.replace(/\\"/g, '"');
      try {
        parsed = JSON.parse(candidate);
      } catch {
        return defaults;
      }
    } else {
      return defaults;
    }

    if (typeof parsed === "string") {
      const nested = parsed.trim();
      if (nested.startsWith("[") || nested.startsWith("{")) {
        try {
          parsed = JSON.parse(nested);
        } catch {
          return defaults;
        }
      }
    }

    const entries: unknown[] = Array.isArray(parsed) ? parsed : [parsed];

    for (const entry of entries) {
      if (!entry || typeof entry !== "object") {
        continue;
      }

      const record = entry as Record<string, unknown>;
      const defaultItem = record as DefaultValueItem;

      const rawValue =
        this.getFieldValue(record, ["Value", "defaultValue", "DefaultValue"], true) ??
        defaultItem.Value;

      let numericValue: number | null = null;
      if (typeof rawValue === "number" && Number.isFinite(rawValue)) {
        numericValue = rawValue;
      } else if (typeof rawValue === "string") {
        numericValue = this.parseNumber(rawValue);
      }

      if (numericValue === null) {
        continue;
      }

      const rawCode =
        this.getFieldValue(record, ["Code", "code"], true) ??
        defaultItem.Code;
      const code = typeof rawCode === "string" ? this.normalizeCodeKey(rawCode) : "";
      if (code) {
        defaults.set(`code:${code}`, numericValue);
      }

      const rawId =
        this.getFieldValue(record, ["IndicatorID", "IndicatorId", "DataID", "DataId"], true) ??
        defaultItem.IndicatorID ??
        defaultItem.DataID;
      const id = typeof rawId === "string" ? rawId.trim().toLowerCase() : "";
      if (id) {
        defaults.set(`id:${id}`, numericValue);
      }
    }

    return defaults;
  }

  // Resolve row-level default value from Code first, then DataID/IndicatorID.
  private getPerRowDefaultValue(item: RowItem, defaults: Map<string, number>): number | null {
    const codeKey = this.normalizeCodeKey(item.Code);
    if (codeKey && defaults.has(`code:${codeKey}`)) {
      return defaults.get(`code:${codeKey}`) ?? null;
    }

    const dataId = typeof item.DataID === "string" ? item.DataID.trim().toLowerCase() : "";
    if (dataId && defaults.has(`id:${dataId}`)) {
      return defaults.get(`id:${dataId}`) ?? null;
    }

    return null;
  }

  // ===================== HELPERS =====================
  private addOption(select: HTMLSelectElement, value: string, text: string): void {
    const opt = document.createElement("option");
    opt.value = value;
    opt.text = text;
    select.appendChild(opt);
  }

  private parseNumber(v: string): number | null {
    if (!v) return null;
    const normalized = v.replace(",", ".").trim();
    const n = Number(normalized);
    return Number.isFinite(n) ? n : null;
  }

  private isValueEmpty(v: string | null | undefined): boolean {
    return !v || v.trim() === "";
  }

  private getValueValidationReason(
    editableInput: boolean,
    currentValueText: string | null | undefined,
    previousValue: number | null
  ): string | null {
    if (!editableInput) {
      return null;
    }

    if (this.isValueEmpty(currentValueText)) {
      return "Value is required";
    }

    const currentValue = this.parseNumber(currentValueText ?? "");
    if (currentValue === null || previousValue === null || previousValue === 0) {
      return null;
    }

    const gapPercent = (Math.abs(currentValue - previousValue) / Math.abs(previousValue)) * 100;
    return gapPercent > this._gapPercentThreshold ? "Justification required" : null;
  }

  private resolveGapPercentThreshold(raw: number | null | undefined): number {
    if (typeof raw !== "number" || !Number.isFinite(raw) || raw < 0) {
      return 50;
    }
    return raw;
  }

  private round(v: number, decimals: number): number {
    const p = Math.pow(10, decimals);
    return Math.round(v * p) / p;
  }

  private formatNumber(v: number): string {
    return String(v);
  }

  private escape(v: string): string {
    const s = v ?? "";
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  private injectStyles(host: HTMLElement): void {
    const style = document.createElement("style");
    style.innerHTML = `
      .pcf-table { width: 100%; font-family: Segoe UI, Arial; font-size: 13px; }

      /* 7 columns: Indicator, Unit, Previous Year, Pre-filled, Current Year, % N/N-1, Actions */
      .pcf-row {
        display: grid;
        grid-template-columns: 3fr 1fr 1fr 1fr 1fr 1fr 1fr;
        gap: 12px;
        align-items: center;
        padding: 14px 10px;
        border-bottom: 1px solid #e5e7eb;
      }

      .pcf-row.header {
        font-weight: 700;
     
        background: #f9fafb;
        font-size: 12px;
      }

      .indicator .title { font-weight: 700; }
      .indicator .desc { font-size: 11px; color: #6b7280; margin-top: 3px; }

      .unitSelect {
        width: 100%;
        padding: 6px 8px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        background: white;
      }

      .unitSelect:disabled {
        background: #f3f4f6;
        color: #6b7280;
        cursor: not-allowed;
      }

      .valueInput {
        width: 100%;
        padding: 6px 8px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
      }

      .valueInputReadonly {
        background: #f3f4f6;
        color: #6b7280;
        cursor: not-allowed;
      }

      .valueInputError {
        border-color: #dc2626;
      }

      .valueCell {
        display: flex;
        flex-direction: column;
        gap: 4px;
        padding-right: 6px;
      }

      .valueError {
        display: none;
        font-size: 11px;
        line-height: 1.2;
        color: #dc2626;
      }

      .valueInput:focus { border-color: #2563eb; outline: none; }

      .pctCell {
        font-weight: 700;
        white-space: nowrap;
        padding-left: 10px;
      }

      .actions .action { cursor: pointer; margin-right: 10px; color: #2563eb; user-select: none; }
      .pcf-empty { padding: 14px 10px; color: #6b7280; }
    `;
    host.appendChild(style);
  }
}