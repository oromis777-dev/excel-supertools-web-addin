import { ConfigService } from "./configService.js";

let selectionEventHandler = null;
let activeWorksheetId = null;

export function getBlendedColor(hex, transparencyPercent) {
  let h = (hex || "#417EBC").replace("#", "");
  if (h.length === 3) h = h.split("").map(c => c + c).join("");
  const r = parseInt(h.substring(0, 2), 16) || 65;
  const g = parseInt(h.substring(2, 4), 16) || 126;
  const b = parseInt(h.substring(4, 6), 16) || 188;

  const t = Math.max(0, Math.min(90, Number(transparencyPercent) || 0)) / 100;
  // Змішування кольору з білим фоном для створення ефекту прозорості в клітинках Excel
  const nr = Math.round(r * (1 - t) + 255 * t);
  const ng = Math.round(g * (1 - t) + 255 * t);
  const nb = Math.round(b * (1 - t) + 255 * t);

  const toHex = (n) => Math.min(255, Math.max(0, n)).toString(16).padStart(2, "0");
  return `#${toHex(nr)}${toHex(ng)}${toHex(nb)}`;
}

export const FocusGridService = {
  async toggle() {
    const config = await ConfigService.loadConfig();
    config.focusGrid.enabled = !config.focusGrid.enabled;
    await ConfigService.saveConfig(config);

    if (config.focusGrid.enabled) {
      await this.enable();
    } else {
      await this.disable();
    }
    return config.focusGrid.enabled;
  },

  async enable() {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load("id");
      await context.sync();

      activeWorksheetId = worksheet.id;

      if (!selectionEventHandler) {
        selectionEventHandler = worksheet.onSelectionChanged.add(this.handleSelectionChange.bind(this));
      }

      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();

      await this._updateHighlight(context, worksheet, activeCell.rowIndex, activeCell.columnIndex);
      await context.sync();
    });
  },

  async disable() {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      if (selectionEventHandler) {
        selectionEventHandler.remove();
        selectionEventHandler = null;
      }
      activeWorksheetId = null;

      await this._clearHighlight(context, worksheet);
      await context.sync();
    });
  },

  async applyCurrentSettings() {
    const config = (await ConfigService.loadConfig()).focusGrid;
    if (!config.enabled) return;

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();

      await this._updateHighlight(context, worksheet, activeCell.rowIndex, activeCell.columnIndex);
      await context.sync();
    });
  },

  async handleSelectionChange(event) {
    const config = (await ConfigService.loadConfig()).focusGrid;
    if (!config.enabled) return;

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();

      await this._updateHighlight(context, worksheet, activeCell.rowIndex, activeCell.columnIndex);
      await context.sync();
    });
  },

  async _updateHighlight(context, worksheet, rowIndex, columnIndex) {
    const config = (await ConfigService.loadConfig()).focusGrid;
    const rowNum = rowIndex + 1;
    const colNum = columnIndex + 1;

    // Формула для фігури (0: Перехрестя, 1: Рядок, 2: Стовпець)
    let formula = "";
    if (config.shapeType === 1) {
      formula = `=ROW()=${rowNum}`;
    } else if (config.shapeType === 2) {
      formula = `=COLUMN()=${colNum}`;
    } else {
      formula = `=OR(ROW()=${rowNum}, COLUMN()=${colNum})`;
    }

    const usedRange = worksheet.getUsedRangeOrNullObject(true);
    await context.sync();

    const targetRange = usedRange.isNullObject ? worksheet.getRange("A1:Z100") : usedRange;
    const condFormats = targetRange.conditionalFormats;
    condFormats.load(["items", "type"]);
    await context.sync();

    let existingRule = null;
    let activeCellRule = null;

    if (condFormats.items && condFormats.items.length > 0) {
      for (const cf of condFormats.items) {
        if (cf.type === "Custom") {
          cf.custom.rule.load("formula");
        }
      }
      await context.sync();

      for (const cf of condFormats.items) {
        if (cf.type === "Custom" && cf.custom?.rule?.formula) {
          const f = cf.custom.rule.formula;
          if (f.includes("ROW()=") && f.includes("COLUMN()=") && f.startsWith("=AND(")) {
            activeCellRule = cf;
          } else if (f.includes("ROW()=") || f.includes("COLUMN()=")) {
            if (!existingRule) {
              existingRule = cf;
            } else {
              cf.delete();
            }
          }
        }
      }
    }

    // Розрахунок змішаного кольору з урахуванням відсотка прозорості (0% - 90%)
    const blendedColor = getBlendedColor(config.colorHex, config.transparency);
    const borderWeight = config.lineThickness || 2;
    const styleType = config.styleType !== undefined ? config.styleType : 0; // 0: Fill, 1: Border, 2: Line

    if (existingRule) {
      existingRule.custom.rule.formula = formula;
      if (styleType === 0) {
        existingRule.custom.format.fill.color = blendedColor;
      } else {
        // Якщо вибрано стиль Рамка або Тонка лінія
        existingRule.custom.format.fill.color = getBlendedColor(config.colorHex, Math.max(70, config.transparency));
        existingRule.custom.format.borders.top.color = config.colorHex;
        existingRule.custom.format.borders.bottom.color = config.colorHex;
        existingRule.custom.format.borders.left.color = config.colorHex;
        existingRule.custom.format.borders.right.color = config.colorHex;
      }
    } else {
      const newCf = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      newCf.custom.rule.formula = formula;
      if (styleType === 0) {
        newCf.custom.format.fill.color = blendedColor;
      } else {
        newCf.custom.format.fill.color = getBlendedColor(config.colorHex, Math.max(70, config.transparency));
        newCf.custom.format.borders.top.color = config.colorHex;
        newCf.custom.format.borders.bottom.color = config.colorHex;
        newCf.custom.format.borders.left.color = config.colorHex;
        newCf.custom.format.borders.right.color = config.colorHex;
      }
      newCf.stopIfTrue = false;
    }

    // Підсвічування активної клітинки, якщо ввімкнено відповідний прапорець
    if (config.highlightActiveCell) {
      const activeFormula = `=AND(ROW()=${rowNum}, COLUMN()=${colNum})`;
      const activeFillColor = getBlendedColor(config.colorHex, Math.max(0, config.transparency - 20));
      if (activeCellRule) {
        activeCellRule.custom.rule.formula = activeFormula;
        activeCellRule.custom.format.fill.color = activeFillColor;
      } else {
        const actCf = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        actCf.custom.rule.formula = activeFormula;
        actCf.custom.format.fill.color = activeFillColor;
        actCf.custom.format.borders.top.color = config.colorHex;
        actCf.custom.format.borders.bottom.color = config.colorHex;
        actCf.custom.format.borders.left.color = config.colorHex;
        actCf.custom.format.borders.right.color = config.colorHex;
        actCf.stopIfTrue = false;
      }
    } else if (activeCellRule) {
      activeCellRule.delete();
    }
  },

  async _clearHighlight(context, worksheet) {
    const usedRange = worksheet.getUsedRangeOrNullObject(true);
    await context.sync();

    const targetRange = usedRange.isNullObject ? worksheet.getRange("A1:Z100") : usedRange;
    const condFormats = targetRange.conditionalFormats;
    condFormats.load(["items", "type"]);
    await context.sync();

    if (condFormats.items && condFormats.items.length > 0) {
      for (const cf of condFormats.items) {
        if (cf.type === "Custom") {
          cf.custom.rule.load("formula");
        }
      }
      await context.sync();

      for (const cf of condFormats.items) {
        if (cf.type === "Custom" && cf.custom?.rule?.formula) {
          const f = cf.custom.rule.formula;
          if (f.includes("ROW()=") || f.includes("COLUMN()=")) {
            cf.delete();
          }
        }
      }
    }
  }
};
