import { ConfigService } from "./configService.js";

let selectionEventHandler = null;
let activeWorksheetId = null;
let debounceTimer = null;
let isUpdating = false;
let cachedGridRule = null;
let cachedActiveCellRule = null;

export function getBlendedColor(hex, transparencyPercent) {
  let h = (hex || "#417EBC").replace("#", "");
  if (h.length === 3) h = h.split("").map(c => c + c).join("");
  const r = parseInt(h.substring(0, 2), 16) || 65;
  const g = parseInt(h.substring(2, 4), 16) || 126;
  const b = parseInt(h.substring(4, 6), 16) || 188;

  const t = Math.max(0, Math.min(90, Number(transparencyPercent) || 0)) / 100;
  // Змішування кольору з білим фоном для створення ефекту напівпрозорості в клітинках Excel
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
    if (debounceTimer) {
      clearTimeout(debounceTimer);
      debounceTimer = null;
    }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load(["id", "name"]);
      await context.sync();

      activeWorksheetId = worksheet.id;
      cachedGridRule = null;
      cachedActiveCellRule = null;

      if (!selectionEventHandler) {
        selectionEventHandler = worksheet.onSelectionChanged.add(this.handleSelectionChange.bind(this));
      }

      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();

      await this._fastUpdateHighlight(context, worksheet, activeCell.rowIndex, activeCell.columnIndex, true);
      await context.sync();
    });
  },

  async disable() {
    if (debounceTimer) {
      clearTimeout(debounceTimer);
      debounceTimer = null;
    }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      if (selectionEventHandler) {
        selectionEventHandler.remove();
        selectionEventHandler = null;
      }
      activeWorksheetId = null;
      cachedGridRule = null;
      cachedActiveCellRule = null;

      await this._clearAllHighlightRules(context, worksheet);
      await context.sync();
    });
  },

  async applyCurrentSettings() {
    const config = (await ConfigService.loadConfig()).focusGrid;
    if (!config.enabled) return;

    // Скидаємо кешовані правила, щоб застосувати нові кольори/стилі
    cachedGridRule = null;
    cachedActiveCellRule = null;

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();

      await this._fastUpdateHighlight(context, worksheet, activeCell.rowIndex, activeCell.columnIndex, true);
      await context.sync();
    });
  },

  /**
   * Оптимізований обробник переміщення виділення з Throttling / Debounce (60 мс).
   * Запобігає перевантаженню Excel сотнями паралельних запитів при швидкому скролі стрілками.
   */
  handleSelectionChange(event) {
    if (debounceTimer) {
      clearTimeout(debounceTimer);
    }

    debounceTimer = setTimeout(() => {
      this._executeDebouncedUpdate();
    }, 60);
  },

  async _executeDebouncedUpdate() {
    if (isUpdating) {
      // Якщо попередня синхронізація ще триває, запланувати повтор після її завершення
      setTimeout(() => this._executeDebouncedUpdate(), 40);
      return;
    }

    const config = (await ConfigService.loadConfig()).focusGrid;
    if (!config.enabled) return;

    isUpdating = true;
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const activeCell = context.workbook.getActiveCell();
        activeCell.load(["rowIndex", "columnIndex"]);
        worksheet.load("id");
        await context.sync();

        if (activeWorksheetId !== worksheet.id) {
          activeWorksheetId = worksheet.id;
          cachedGridRule = null;
          cachedActiveCellRule = null;
        }

        await this._fastUpdateHighlight(context, worksheet, activeCell.rowIndex, activeCell.columnIndex, false);
        await context.sync();
      });
    } catch (err) {
      console.warn("[ExcelSuperTools] Помилка фокусної сітки:", err);
    } finally {
      isUpdating = false;
    }
  },

  /**
   * Швидке оновлення підсвітки:
   * Замість видалення/створення нових правил на кожен клік (що викликало фантомні підсвітки та блокування буфера),
   * ми мутуємо формулу ОДНОГО постійного правила.
   */
  async _fastUpdateHighlight(context, worksheet, rowIndex, columnIndex, forceReloadRules = false) {
    const config = (await ConfigService.loadConfig()).focusGrid;
    const rowNum = rowIndex + 1;
    const colNum = columnIndex + 1;

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

    // Обмежуємо діапазон лише областю з даними + невеликим запасом (запобігає зависанню на мільйоні порожніх клітинок)
    const targetRange = (!usedRange.isNullObject) ? usedRange : worksheet.getRange("A1:AZ200");
    const condFormats = targetRange.conditionalFormats;

    const blendedColor = getBlendedColor(config.colorHex, config.transparency);
    const styleType = config.styleType !== undefined ? config.styleType : 0;

    // Якщо правило вже закешоване і не вимагається повне перезавантаження
    if (!forceReloadRules && cachedGridRule) {
      try {
        cachedGridRule.custom.rule.formula = formula;
        if (config.highlightActiveCell && cachedActiveCellRule) {
          cachedActiveCellRule.custom.rule.formula = `=AND(ROW()=${rowNum}, COLUMN()=${colNum})`;
        }
        return;
      } catch (cacheErr) {
        // Якщо кеш застарів (наприклад, аркуш змінено), скидаємо та перескановуємо
        cachedGridRule = null;
        cachedActiveCellRule = null;
      }
    }

    // Пошук або ініціалізація єдиного Master-правила
    condFormats.load(["items", "type"]);
    await context.sync();

    let gridRule = null;
    let activeRule = null;
    const duplicateRulesToDelete = [];

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
            if (!activeRule) activeRule = cf;
            else duplicateRulesToDelete.push(cf);
          } else if (f.includes("ROW()=") || f.includes("COLUMN()=")) {
            if (!gridRule) gridRule = cf;
            else duplicateRulesToDelete.push(cf);
          }
        }
      }
    }

    // Видаляємо дублікати правил, якщо вони виникли раніше
    for (const dup of duplicateRulesToDelete) {
      dup.delete();
    }

    // Оновлюємо або створюємо рівно ОДНЕ правило сітки
    if (gridRule) {
      gridRule.custom.rule.formula = formula;
      this._applyStyleToRule(gridRule, styleType, blendedColor, config.colorHex, config.transparency);
    } else {
      gridRule = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      gridRule.custom.rule.formula = formula;
      this._applyStyleToRule(gridRule, styleType, blendedColor, config.colorHex, config.transparency);
      gridRule.stopIfTrue = false;
    }
    cachedGridRule = gridRule;

    // Обробка підсвітки активної комірки
    if (config.highlightActiveCell) {
      const activeFormula = `=AND(ROW()=${rowNum}, COLUMN()=${colNum})`;
      const activeFillColor = getBlendedColor(config.colorHex, Math.max(0, config.transparency - 20));

      if (activeRule) {
        activeRule.custom.rule.formula = activeFormula;
        activeRule.custom.format.fill.color = activeFillColor;
      } else {
        activeRule = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        activeRule.custom.rule.formula = activeFormula;
        activeRule.custom.format.fill.color = activeFillColor;
        activeRule.custom.format.borders.top.color = config.colorHex;
        activeRule.custom.format.borders.bottom.color = config.colorHex;
        activeRule.custom.format.borders.left.color = config.colorHex;
        activeRule.custom.format.borders.right.color = config.colorHex;
        activeRule.stopIfTrue = false;
      }
      cachedActiveCellRule = activeRule;
    } else if (activeRule) {
      activeRule.delete();
      cachedActiveCellRule = null;
    }
  },

  _applyStyleToRule(cf, styleType, blendedColor, mainColorHex, transparency) {
    if (styleType === 0) {
      cf.custom.format.fill.color = blendedColor;
    } else {
      cf.custom.format.fill.color = getBlendedColor(mainColorHex, Math.max(70, transparency));
      cf.custom.format.borders.top.color = mainColorHex;
      cf.custom.format.borders.bottom.color = mainColorHex;
      cf.custom.format.borders.left.color = mainColorHex;
      cf.custom.format.borders.right.color = mainColorHex;
    }
  },

  async _clearAllHighlightRules(context, worksheet) {
    const usedRange = worksheet.getUsedRangeOrNullObject(true);
    await context.sync();

    const targetRange = (!usedRange.isNullObject) ? usedRange : worksheet.getRange("A1:AZ200");
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
