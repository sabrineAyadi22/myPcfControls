import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class DataverseCaller implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private container!: HTMLDivElement;
  private notifyOutputChanged!: () => void;

  private root!: HTMLDivElement;
  private sellingInput!: HTMLInputElement;
  private costInput!: HTMLInputElement;
  private indiceInput!: HTMLInputElement;

  private resultEl!: HTMLDivElement;
  private debugEl!: HTMLDivElement;

  private _resultValue?: number;
  private _debugText?: string;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;

    // Root UI
    this.root = document.createElement("div");
    this.root.style.fontFamily = "Segoe UI, Arial, sans-serif";
    this.root.style.border = "1px solid #D0D0D0";
    this.root.style.borderRadius = "10px";
    this.root.style.padding = "12px";
    this.root.style.background = "#FFFFFF";

    const title = document.createElement("div");
    title.style.fontWeight = "600";
    title.style.marginBottom = "8px";
    title.innerText = "✅ DataverseCaller (PCF)";

    const grid = document.createElement("div");
    grid.style.display = "grid";
    grid.style.gridTemplateColumns = "1fr 1fr 1fr";
    grid.style.gap = "8px";
    grid.style.marginBottom = "10px";

    this.sellingInput = this.createTextBox("Selling Price");
    this.costInput = this.createTextBox("Cost Price");
    this.indiceInput = this.createTextBox("Indice (name)");

    this.sellingInput.addEventListener("input", () => this.recomputeFromUI());
    this.costInput.addEventListener("input", () => this.recomputeFromUI());
    this.indiceInput.addEventListener("input", () => this.recomputeFromUI());

    grid.appendChild(this.wrapLabeled("Selling Price", this.sellingInput));
    grid.appendChild(this.wrapLabeled("Cost Price", this.costInput));
    grid.appendChild(this.wrapLabeled("Indice", this.indiceInput));

    this.resultEl = document.createElement("div");
    this.resultEl.style.padding = "10px";
    this.resultEl.style.borderRadius = "8px";
    this.resultEl.style.background = "#F6F8FA";
    this.resultEl.style.border = "1px solid #E5E7EB";
    this.resultEl.style.marginBottom = "8px";
    this.resultEl.style.whiteSpace = "pre-line";
    this.resultEl.innerText = "Result: (waiting for inputs)";

    this.debugEl = document.createElement("div");
    this.debugEl.style.fontSize = "12px";
    this.debugEl.style.color = "#555";
    this.debugEl.style.whiteSpace = "pre-line";

    this.root.appendChild(title);
    this.root.appendChild(grid);
    this.root.appendChild(this.resultEl);
    this.root.appendChild(this.debugEl);

    this.container.appendChild(this.root);

    // Initial sync from Canvas -> UI
    this.syncInputsFromCanvas(context);
    this.recomputeFromUI();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Canvas -> UI
    this.syncInputsFromCanvas(context);

    // UI -> compute outputs
    this.recomputeFromUI();

    // Debug (size + current values)
    const w = context.mode.allocatedWidth;
    const h = context.mode.allocatedHeight;
    const selling = this.parseNumber(this.sellingInput.value);
    const cost = this.parseNumber(this.costInput.value);
    const indice = (this.indiceInput.value ?? "").trim();

    this._debugText =
      `updateView called\n` +
      `W=${w}, H=${h}\n` +
      `selling=${selling ?? "null"} | cost=${cost ?? "null"} | indice="${indice}"`;

    this.debugEl.innerText = this._debugText;
  }

  public getOutputs(): IOutputs {
    return {
      resultValue: this._resultValue,
      debugText: this._debugText
    };
  }

  public destroy(): void {
    if (this.root && this.root.parentElement) {
      this.root.parentElement.removeChild(this.root);
    }
  }

  // ---------------- Helpers ----------------

  private createTextBox(placeholder: string): HTMLInputElement {
    const input = document.createElement("input");
    input.type = "text";
    input.placeholder = placeholder;
    input.style.width = "100%";
    input.style.boxSizing = "border-box";
    input.style.padding = "10px";
    input.style.borderRadius = "8px";
    input.style.border = "1px solid #D1D5DB";
    input.style.fontSize = "14px";
    return input;
  }

  private wrapLabeled(labelText: string, input: HTMLInputElement): HTMLDivElement {
    const wrap = document.createElement("div");
    const label = document.createElement("div");
    label.style.fontSize = "12px";
    label.style.color = "#444";
    label.style.marginBottom = "4px";
    label.innerText = labelText;
    wrap.appendChild(label);
    wrap.appendChild(input);
    return wrap;
  }

  private syncInputsFromCanvas(context: ComponentFramework.Context<IInputs>): void {
    const selling = context.parameters.sellingPrice.raw;
    const cost = context.parameters.costPrice.raw;
    const indice = context.parameters.indiceName.raw;

    if (selling !== null && selling !== undefined) this.sellingInput.value = String(selling);
    if (cost !== null && cost !== undefined) this.costInput.value = String(cost);
    if (indice !== null && indice !== undefined) this.indiceInput.value = String(indice);
  }

  private recomputeFromUI(): void {
    const selling = this.parseNumber(this.sellingInput.value) ?? 0;
    const cost = this.parseNumber(this.costInput.value) ?? 0;
    const indice = (this.indiceInput.value ?? "").trim().toLowerCase();

    // Demo logic
    let result: number;
    if (indice.includes("ratio")) {
      result = cost === 0 ? 0 : selling / cost;
    } else {
      result = selling - cost;
    }

    this._resultValue = this.round(result, 6);

    this.resultEl.innerText =
      `Result: ${this._resultValue}\n` +
      `Mode: ${indice.includes("ratio") ? "ratio (selling/cost)" : "margin (selling-cost)"}`;

    // Notify framework so Canvas reads getOutputs()
    this.notifyOutputChanged();
  }

  private parseNumber(v: string): number | null {
    if (!v) return null;
    const normalized = v.replace(",", ".").trim();
    const n = Number(normalized);
    return Number.isFinite(n) ? n : null;
  }

  private round(v: number, decimals: number): number {
    const p = Math.pow(10, decimals);
    return Math.round(v * p) / p;
  }
}