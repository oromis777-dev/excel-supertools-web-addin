import { ConfigService } from "./configService.js";

export const CopyRangesService = {
  async copyAreas() {
    return await Excel.run(async (context) => {
      const config = (await ConfigService.loadConfig()).copyRanges || {};
      const skipHidden = config.skipHidden !== false;

      // Отримуємо виділені області (підтримка несуміжних діапазонів)
      let selectedRanges = null;
      try {
        selectedRanges = context.workbook.getSelectedRanges();
        selectedRanges.load("areas");
        await context.sync();
      } catch (e) {
        // Fallback для одного діапазону
        const singleRange = context.workbook.getSelectedRange();
        selectedRanges = { areas: { items: [singleRange] } };
      }

      if (!selectedRanges.areas || selectedRanges.areas.items.length === 0) {
        throw new Error("Жодного діапазону не виділено.");
      }

      const clipboardBuffer = [];
      let totalRowsCopied = 0;

      for (let area of selectedRanges.areas.items) {
        area.load(["values", "formulas", "address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        area.format.load("columnWidth");
        await context.sync();

        let valuesToSave = area.values;
        let formulasToSave = area.formulas;

        if (skipHidden && area.rowCount > 1) {
          // Завантажуємо стан видимості для кожного рядка
          const rowRanges = [];
          for (let r = 0; r < area.rowCount; r++) {
            const row = area.getRow(r);
            row.load("isRowHidden");
            rowRanges.push(row);
          }
          await context.sync();

          const filteredValues = [];
          const filteredFormulas = [];

          for (let r = 0; r < area.rowCount; r++) {
            if (!rowRanges[r].isRowHidden) {
              filteredValues.push(area.values[r]);
              filteredFormulas.push(area.formulas[r]);
            }
          }

          valuesToSave = filteredValues;
          formulasToSave = filteredFormulas;
        }

        if (valuesToSave.length > 0) {
          clipboardBuffer.push({
            address: area.address,
            values: valuesToSave,
            formulas: formulasToSave,
            rowCount: valuesToSave.length,
            columnCount: area.columnCount,
            columnWidth: area.format.columnWidth
          });
          totalRowsCopied += valuesToSave.length;
        }
      }

      if (clipboardBuffer.length === 0) {
        throw new Error("Не знайдено видимих клітинок для копіювання.");
      }

      window._excelSuperToolsClipboard = clipboardBuffer;
      try {
        localStorage.setItem("excel_supertools_clipboard", JSON.stringify(clipboardBuffer));
      } catch (e) {}

      return `Скопійовано ${clipboardBuffer.length} ${clipboardBuffer.length === 1 ? "область" : "області"} (${totalRowsCopied} рядків).`;
    });
  },

  async pasteToActiveCell(overrideMode = null) {
    return await Excel.run(async (context) => {
      const config = (await ConfigService.loadConfig()).copyRanges || {};
      const pasteMode = overrideMode || config.pasteMode || "Values";
      const preserveDims = config.preserveDimensions !== false;

      let buffer = window._excelSuperToolsClipboard;
      if (!buffer || buffer.length === 0) {
        try {
          const saved = localStorage.getItem("excel_supertools_clipboard");
          if (saved) buffer = JSON.parse(saved);
        } catch (e) {}
      }

      if (!buffer || buffer.length === 0) {
        throw new Error("Буфер копіювання порожній. Спочатку виділіть дані та натисніть 'Копіювати'.");
      }

      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();

      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      let currentOffsetRow = 0;
      let totalPastedCells = 0;

      for (let item of buffer) {
        if (!item.values || item.values.length === 0) continue;

        const targetRange = worksheet.getRangeByIndexes(
          activeCell.rowIndex + currentOffsetRow,
          activeCell.columnIndex,
          item.rowCount,
          item.columnCount
        );

        if (pasteMode === "Formulas") {
          targetRange.formulas = item.formulas;
        } else {
          targetRange.values = item.values;
        }

        if (preserveDims && item.columnWidth) {
          try {
            targetRange.format.columnWidth = item.columnWidth;
          } catch (e) {}
        }

        totalPastedCells += item.rowCount * item.columnCount;
        currentOffsetRow += item.rowCount + 1; // Відступ в 1 порожній рядок між областями
      }

      await context.sync();
      return `Вставлено ${buffer.length} ${buffer.length === 1 ? "область" : "області"} (${totalPastedCells} клітинок) у режимі '${pasteMode}'.`;
    });
  },

  getClipboardInfo() {
    let buffer = window._excelSuperToolsClipboard;
    if (!buffer || buffer.length === 0) {
      try {
        const saved = localStorage.getItem("excel_supertools_clipboard");
        if (saved) buffer = JSON.parse(saved);
      } catch (e) {}
    }
    if (!buffer || buffer.length === 0) return null;
    return {
      areasCount: buffer.length,
      totalRows: buffer.reduce((acc, it) => acc + (it.rowCount || 0), 0)
    };
  }
};
