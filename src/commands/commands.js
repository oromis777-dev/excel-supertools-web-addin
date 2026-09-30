import { FocusGridService } from "../services/focusGridService.js";
import { AIService } from "../services/aiService.js";
import { CopyRangesService } from "../services/copyRangesService.js";

Office.onReady((info) => {
  // Команди стрічки ініціалізовано
});

// 1. Дія: Увімкнення / вимкнення фокусної сітки
export async function actionToggleFocusGrid(event) {
  try {
    const isEnabled = await FocusGridService.toggle();
    console.log(`[ExcelSuperTools] Фокусна сітка: ${isEnabled ? "УВІМКНЕНО" : "ВИМКНЕНО"}`);
  } catch (error) {
    console.error("[ExcelSuperTools] Помилка перемикання сітки:", error);
  } finally {
    event.completed();
  }
}

// 2. Дія: AI Аналіз виділеної таблиці даних
export async function actionAnalyzeData(event) {
  try {
    console.log("[ExcelSuperTools] Запуск AI аналізу даних...");
    await AIService.analyzeSelectedRange();
    console.log("[ExcelSuperTools] AI аналіз успішно завершено та збережено на аркуші 'Аналіз_AI'.");
  } catch (error) {
    console.error("[ExcelSuperTools] Помилка AI аналізу:", error);
  } finally {
    event.completed();
  }
}

// 3. Дія: Генерація формули Excel
export async function actionGenerateFormula(event) {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load(["values", "address"]);
      await context.sync();

      const userText = cell.values[0][0];
      if (userText && typeof userText === "string" && userText.trim().length > 0 && !userText.startsWith("=")) {
        const formula = await AIService.generateFormula(userText);
        const adjacentCell = cell.getOffsetRange(0, 1);
        adjacentCell.formulas = [[formula]];
        adjacentCell.select();
        await context.sync();
        console.log(`[ExcelSuperTools] Формулу згенеровано у сусідню клітинку: ${formula}`);
      } else {
        console.log("[ExcelSuperTools] Введіть опис формули в активну клітинку або скористайтеся AI Панеллю.");
      }
    });
  } catch (error) {
    console.error("[ExcelSuperTools] Помилка генерації формули:", error);
  } finally {
    event.completed();
  }
}

// 4. Дія: Копіювання виділених діапазонів
export async function actionShowCopyRangesDialog(event) {
  try {
    const resultMsg = await CopyRangesService.copyAreas();
    console.log(`[ExcelSuperTools] ${resultMsg}`);
  } catch (error) {
    console.error("[ExcelSuperTools] Помилка копіювання:", error);
  } finally {
    event.completed();
  }
}

// Прив'язка функцій до ідентифікаторів у manifest.xml (<FunctionName>)
Office.actions.associate("actionToggleFocusGrid", actionToggleFocusGrid);
Office.actions.associate("actionAnalyzeData", actionAnalyzeData);
Office.actions.associate("actionGenerateFormula", actionGenerateFormula);
Office.actions.associate("actionShowCopyRangesDialog", actionShowCopyRangesDialog);

if (typeof window !== "undefined") {
  window.actionToggleFocusGrid = actionToggleFocusGrid;
  window.actionAnalyzeData = actionAnalyzeData;
  window.actionGenerateFormula = actionGenerateFormula;
  window.actionShowCopyRangesDialog = actionShowCopyRangesDialog;
}
