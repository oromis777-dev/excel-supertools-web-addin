import { ConfigService } from "../services/configService.js";
import { AIService } from "../services/aiService.js";
import { FocusGridService, getBlendedColor } from "../services/focusGridService.js";
import { CopyRangesService } from "../services/copyRangesService.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initApp();
  }
});

async function initApp() {
  await loadUserSettings();
  setupNavigation();
  setupFocusGridUI();
  setupAIUI();
  setupCopyRangesUI();
  setupSettingsUI();
  updateLivePreview();
}

/* ==========================================================================
   1. НАВІГАЦІЯ ПО ВКЛАДКАХ
   ========================================================================== */
function setupNavigation() {
  const tabs = document.querySelectorAll(".nav-tab");
  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => {
        t.classList.remove("active");
        t.setAttribute("aria-selected", "false");
      });
      document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));

      tab.classList.add("active");
      tab.setAttribute("aria-selected", "true");
      const targetId = tab.getAttribute("data-tab");
      const targetEl = document.getElementById(targetId);
      if (targetEl) targetEl.classList.add("active");
    });
  });
}

/* ==========================================================================
   2. ФОКУСНА СІТКА (KUTOOLS STYLE + LIVE PREVIEW)
   ========================================================================== */
let localGridState = {
  enabled: false,
  shapeType: 0,
  styleType: 0,
  colorHex: "#417EBC",
  transparency: 40,
  lineThickness: 2,
  highlightActiveCell: true,
  autoFollow: true
};

function setupFocusGridUI() {
  const chkEnable = document.getElementById("chk-grid-enable");
  const colorPicker = document.getElementById("grid-color");
  const rngTransparency = document.getElementById("rng-transparency");
  const lblTransparency = document.getElementById("lbl-transparency");
  const chkActive = document.getElementById("chk-highlight-active");
  const chkAutoFollow = document.getElementById("chk-auto-follow");
  const btnApply = document.getElementById("btn-apply-grid");
  const btnReset = document.getElementById("btn-reset-grid");

  // Перемикач стану
  chkEnable.addEventListener("change", async () => {
    localGridState.enabled = chkEnable.checked;
    await syncGridConfig();
    if (chkEnable.checked) {
      await FocusGridService.enable();
      showToast("Фокусну сітку увімкнено");
    } else {
      await FocusGridService.disable();
      showToast("Фокусну сітку вимкнено");
    }
  });

  // Вибір фігури (радіокнопки)
  document.querySelectorAll("input[name='rad-grid-shape']").forEach(rad => {
    rad.addEventListener("change", () => {
      localGridState.shapeType = parseInt(rad.value, 10);
      updateLivePreview();
    });
  });

  // Вибір стилю (радіокнопки)
  document.querySelectorAll("input[name='rad-grid-style']").forEach(rad => {
    rad.addEventListener("change", () => {
      localGridState.styleType = parseInt(rad.value, 10);
      updateLivePreview();
    });
  });

  // Color picker
  colorPicker.addEventListener("input", (e) => {
    localGridState.colorHex = e.target.value;
    updateActivePresetDot(e.target.value);
    updateLivePreview();
  });

  // Швидкі пресети кольорів (Kutools palette)
  document.querySelectorAll(".preset-dot").forEach(dot => {
    dot.addEventListener("click", () => {
      const col = dot.getAttribute("data-color");
      colorPicker.value = col;
      localGridState.colorHex = col;
      updateActivePresetDot(col);
      updateLivePreview();
    });
  });

  // Повзунок прозорості (0 - 90%)
  rngTransparency.addEventListener("input", (e) => {
    const val = parseInt(e.target.value, 10);
    localGridState.transparency = val;
    lblTransparency.textContent = `${val}%`;
    updateLivePreview();
  });

  // Товщина лінії
  document.querySelectorAll("input[name='rad-thickness']").forEach(rad => {
    rad.addEventListener("change", () => {
      localGridState.lineThickness = parseInt(rad.value, 10);
      updateLivePreview();
    });
  });

  // Прапорці
  chkActive.addEventListener("change", () => {
    localGridState.highlightActiveCell = chkActive.checked;
    updateLivePreview();
  });
  chkAutoFollow.addEventListener("change", () => {
    localGridState.autoFollow = chkAutoFollow.checked;
  });

  // Кнопка застосувати до Excel
  btnApply.addEventListener("click", async () => {
    await syncGridConfig();
    if (localGridState.enabled) {
      await FocusGridService.applyCurrentSettings();
      showToast("Параметри фокусної сітки застосовано!");
    } else {
      localGridState.enabled = true;
      chkEnable.checked = true;
      await syncGridConfig();
      await FocusGridService.enable();
      showToast("Фокусну сітку увімкнено та застосовано!");
    }
  });

  // Кнопка скинути до типових
  btnReset.addEventListener("click", async () => {
    localGridState.shapeType = 0;
    localGridState.styleType = 0;
    localGridState.colorHex = "#417EBC";
    localGridState.transparency = 40;
    localGridState.lineThickness = 2;
    localGridState.highlightActiveCell = true;
    localGridState.autoFollow = true;

    // Синхронізація UI
    colorPicker.value = "#417EBC";
    rngTransparency.value = 40;
    lblTransparency.textContent = "40%";
    chkActive.checked = true;
    chkAutoFollow.checked = true;

    const radShape = document.querySelector("input[name='rad-grid-shape'][value='0']");
    if (radShape) radShape.checked = true;
    const radStyle = document.querySelector("input[name='rad-grid-style'][value='0']");
    if (radStyle) radStyle.checked = true;
    const radThick = document.querySelector("input[name='rad-thickness'][value='2']");
    if (radThick) radThick.checked = true;

    updateActivePresetDot("#417EBC");
    updateLivePreview();
    await syncGridConfig();
    if (localGridState.enabled) {
      await FocusGridService.applyCurrentSettings();
    }
    showToast("Налаштування скинуто до типових");
  });
}

function updateActivePresetDot(hex) {
  document.querySelectorAll(".preset-dot").forEach(d => {
    if (d.getAttribute("data-color").toLowerCase() === hex.toLowerCase()) {
      d.classList.add("active");
    } else {
      d.classList.remove("active");
    }
  });
}

/* Оновлення міні-таблиці попереднього перегляду (Live Interactive Preview) */
function updateLivePreview() {
  const table = document.getElementById("preview-table");
  const codeBadge = document.getElementById("preview-color-code");
  if (!table) return;

  const { shapeType, styleType, colorHex, transparency, lineThickness, highlightActiveCell } = localGridState;
  const blendedColor = getBlendedColor(colorHex, transparency);
  const activeCellColor = getBlendedColor(colorHex, Math.max(0, transparency - 25));

  if (codeBadge) codeBadge.textContent = colorHex;

  const rows = table.querySelectorAll("tbody tr");
  rows.forEach((tr, rIdx) => {
    const tds = tr.querySelectorAll("td");
    tds.forEach((td, cIdx) => {
      const isTargetRow = (rIdx === 2); // Рядок 3
      const isTargetCol = (cIdx === 2); // Стовпець C
      const isActiveCell = isTargetRow && isTargetCol;

      // Скидання попередніх стилів
      td.style.backgroundColor = "";
      td.style.border = "1px solid #E5E5E5";
      td.style.boxShadow = "";

      let shouldHighlight = false;
      if (shapeType === 0) shouldHighlight = isTargetRow || isTargetCol;
      else if (shapeType === 1) shouldHighlight = isTargetRow;
      else if (shapeType === 2) shouldHighlight = isTargetCol;

      if (shouldHighlight) {
        if (styleType === 0) {
          // Заливка (Fill)
          td.style.backgroundColor = blendedColor;
        } else if (styleType === 1) {
          // Рамка (Border)
          td.style.backgroundColor = getBlendedColor(colorHex, Math.max(75, transparency));
          const w = `${lineThickness}px`;
          td.style.border = `${w} solid ${colorHex}`;
        } else if (styleType === 2) {
          // Тонка лінія
          td.style.backgroundColor = "transparent";
          td.style.borderTop = `1px dashed ${colorHex}`;
          td.style.borderBottom = `1px dashed ${colorHex}`;
        }
      }

      // Окремий акцент для активної клітинки (C3)
      if (isActiveCell && highlightActiveCell) {
        td.style.backgroundColor = activeCellColor;
        td.style.border = `2px solid ${colorHex}`;
        td.style.boxShadow = `0 0 4px ${colorHex}88`;
      }
    });
  });
}

async function syncGridConfig() {
  const config = await ConfigService.loadConfig();
  config.focusGrid = { ...config.focusGrid, ...localGridState };
  await ConfigService.saveConfig(config);
}

/* ==========================================================================
   3. AI ПОМІЧНИК (GEMINI 3.6 FLASH + ACTIONS)
   ========================================================================== */
function setupAIUI() {
  const promptInput = document.getElementById("ai-prompt");
  const outputBox = document.getElementById("ai-output");
  const btnSubmit = document.getElementById("btn-submit-ai");
  const btnClear = document.getElementById("btn-clear-ai");
  const btnCopy = document.getElementById("btn-copy-ai");
  const btnExport = document.getElementById("btn-export-sheet");

  const btnQuickAnalyze = document.getElementById("btn-quick-analyze");
  const btnQuickFormula = document.getElementById("btn-quick-formula");
  const btnQuickExplain = document.getElementById("btn-quick-explain");
  const btnQuickErrors = document.getElementById("btn-quick-errors");

  // Виконати довільний запит
  btnSubmit.addEventListener("click", async () => {
    const text = promptInput.value.trim();
    if (!text) {
      showToast("Будь ласка, введіть опис запиту");
      promptInput.focus();
      return;
    }
    await runAIWithContext(text);
  });

  btnClear.addEventListener("click", () => {
    promptInput.value = "";
    outputBox.textContent = "Тут з'явиться структурована відповідь AI або аналітичний звіт...";
  });

  btnCopy.addEventListener("click", () => {
    const txt = outputBox.innerText;
    if (!txt || txt.includes("Тут з'явиться")) return;
    navigator.clipboard.writeText(txt).then(() => {
      showToast("Відповідь скопійовано в буфер обміну!");
    });
  });

  btnExport.addEventListener("click", async () => {
    const text = outputBox.innerText;
    if (!text || text.includes("Тут з'явиться")) {
      showToast("Спочатку виконайте AI аналіз");
      return;
    }
    await exportToAnalysisSheet(text);
  });

  // Швидкі дії
  btnQuickAnalyze.addEventListener("click", async () => {
    promptInput.value = "Проаналізуй дані виділеного діапазону та надай ключові бізнес-висновки з практичними рекомендаціями.";
    await runAIWithContext(promptInput.value);
  });

  btnQuickFormula.addEventListener("click", () => {
    promptInput.value = "Створи формулу для: ";
    promptInput.focus();
    promptInput.setSelectionRange(promptInput.value.length, promptInput.value.length);
  });

  btnQuickExplain.addEventListener("click", async () => {
    try {
      await Excel.run(async (context) => {
        const activeCell = context.workbook.getActiveCell();
        activeCell.load(["formulas", "values", "address"]);
        await context.sync();

        const formula = activeCell.formulas[0][0];
        if (!formula || !formula.startsWith("=")) {
          showToast(`У клітинці ${activeCell.address} немає формули`);
          return;
        }
        promptInput.value = `Поясни логіку формули ${formula} у клітинці ${activeCell.address}`;
        await runAIWithContext(promptInput.value, `Клітинка: ${activeCell.address}, Значення: ${activeCell.values[0][0]}`);
      });
    } catch (err) {
      showToast(`Помилка: ${err.message}`);
    }
  });

  btnQuickErrors.addEventListener("click", async () => {
    promptInput.value = "Знайди можливі помилки (#N/A, #VALUE!, дублікати або аномалії) у виділених даних та запропонуй як їх виправити.";
    await runAIWithContext(promptInput.value);
  });
}

async function runAIWithContext(promptText, extraContext = "") {
  const outputBox = document.getElementById("ai-output");
  outputBox.innerHTML = `<span style="color:var(--primary); font-weight:600;">⏳ AI виконує аналіз... Зачекайте...</span>`;

  try {
    let sheetContext = extraContext;
    if (!sheetContext) {
      sheetContext = await getActiveRangeContext();
    }

    const result = await AIService.executePrompt(promptText, sheetContext);
    outputBox.innerText = result;
    showToast("AI відповідь успішно сформовано!");
  } catch (err) {
    outputBox.innerHTML = `<span style="color:var(--error); font-weight:600;">❌ Помилка: ${err.message}</span>`;
  }
}

async function getActiveRangeContext() {
  try {
    return await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount", "address"]);
      await context.sync();

      if (range.rowCount === 1 && range.columnCount === 1 && !range.values[0][0]) {
        return "";
      }

      const rows = range.values.slice(0, 30).map(r => r.join("\t")).join("\n");
      return `Діапазон ${range.address} (${range.rowCount} рядків x ${range.columnCount} колонок):\n${rows}`;
    });
  } catch (e) {
    return "";
  }
}

async function exportToAnalysisSheet(text) {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetName = "Аналіз_AI";
      let targetSheet = sheets.items.find(s => s.name === sheetName);
      if (!targetSheet) {
        targetSheet = sheets.add(sheetName);
      }
      targetSheet.activate();

      const lines = text.split("\n");
      const values = lines.map(line => [line]);

      const writeRange = targetSheet.getRange(`A1:A${lines.length}`);
      writeRange.values = values;
      targetSheet.getRange("A1").format.font.bold = true;
      targetSheet.getRange("A:A").format.autofitColumns();

      await context.sync();
      showToast(`Результат успішно додано на аркуш '${sheetName}'!`);
    });
  } catch (err) {
    showToast(`Помилка експорту: ${err.message}`);
  }
}

/* ==========================================================================
   4. КОПІЮВАННЯ ТА ВСТАВКА ДІАПАЗОНІВ
   ========================================================================== */
function setupCopyRangesUI() {
  const btnCopy = document.getElementById("btn-copy-ranges");
  const btnPaste = document.getElementById("btn-paste-ranges");
  const statusText = document.getElementById("clipboard-status-text");

  btnCopy.addEventListener("click", async () => {
    try {
      const res = await CopyRangesService.copySelectedRanges();
      statusText.innerHTML = `<b>У буфері:</b> ${res.rangeCount} діапазон(ів), всього ${res.totalCells} клітинок`;
      showToast(`Скопійовано ${res.rangeCount} діапазон(ів)`);
    } catch (err) {
      showToast(`Помилка копіювання: ${err.message}`);
    }
  });

  btnPaste.addEventListener("click", async () => {
    try {
      const skipHidden = document.getElementById("chk-skip-hidden").checked;
      const preserveWidths = document.getElementById("chk-preserve-widths").checked;
      const pasteMode = document.getElementById("sel-paste-mode").value;

      await CopyRangesService.pasteToActiveCell({ skipHidden, preserveWidths, pasteMode });
      showToast("Діапазони успішно вставлено!");
    } catch (err) {
      showToast(`Помилка вставки: ${err.message}`);
    }
  });
}

/* ==========================================================================
   5. НАЛАШТУВАННЯ ТА ЗБЕРЕЖЕННЯ КОНФІГУРАЦІЇ
   ========================================================================== */
function setupSettingsUI() {
  const selProvider = document.getElementById("sel-ai-provider");
  const btnSave = document.getElementById("btn-save-settings");

  selProvider.addEventListener("change", () => {
    toggleProviderInputs(selProvider.value);
  });

  btnSave.addEventListener("click", async () => {
    const config = await ConfigService.loadConfig();
    const provider = selProvider.value;

    config.ai.provider = provider;
    config.ai.geminiApiKey = document.getElementById("txt-gemini-key").value.trim();
    config.ai.openaiApiKey = document.getElementById("txt-openai-key").value.trim();
    config.ai.deepseekApiKey = document.getElementById("txt-deepseek-key").value.trim();
    config.ai.customEndpoint = document.getElementById("txt-custom-endpoint").value.trim();
    config.ai.offlineMode = document.getElementById("chk-offline-mode").checked;

    if (provider === "gemini") {
      config.ai.model = document.getElementById("sel-gemini-model").value;
    } else if (provider === "openai") {
      config.ai.model = document.getElementById("sel-openai-model").value;
    }

    await ConfigService.saveConfig(config);
    showToast("Налаштування успішно збережено!");
  });
}

function toggleProviderInputs(provider) {
  document.querySelectorAll(".provider-group").forEach(el => el.classList.add("hidden"));
  const group = document.getElementById(`group-${provider}`);
  if (group) group.classList.remove("hidden");
}

async function loadUserSettings() {
  const config = await ConfigService.loadConfig();

  // Сітка
  localGridState = { ...localGridState, ...config.focusGrid };
  document.getElementById("chk-grid-enable").checked = localGridState.enabled;
  document.getElementById("grid-color").value = localGridState.colorHex || "#417EBC";
  document.getElementById("rng-transparency").value = localGridState.transparency || 40;
  document.getElementById("lbl-transparency").textContent = `${localGridState.transparency || 40}%`;
  document.getElementById("chk-highlight-active").checked = localGridState.highlightActiveCell !== false;
  document.getElementById("chk-auto-follow").checked = localGridState.autoFollow !== false;

  const shapeRad = document.querySelector(`input[name='rad-grid-shape'][value='${localGridState.shapeType || 0}']`);
  if (shapeRad) shapeRad.checked = true;
  const styleRad = document.querySelector(`input[name='rad-grid-style'][value='${localGridState.styleType || 0}']`);
  if (styleRad) styleRad.checked = true;
  const thickRad = document.querySelector(`input[name='rad-thickness'][value='${localGridState.lineThickness || 2}']`);
  if (thickRad) thickRad.checked = true;

  updateActivePresetDot(localGridState.colorHex || "#417EBC");

  // AI
  const ai = config.ai;
  const selProvider = document.getElementById("sel-ai-provider");
  selProvider.value = ai.provider || "gemini";
  toggleProviderInputs(selProvider.value);

  document.getElementById("txt-gemini-key").value = ai.geminiApiKey || "";
  document.getElementById("txt-openai-key").value = ai.openaiApiKey || "";
  document.getElementById("txt-deepseek-key").value = ai.deepseekApiKey || "";
  document.getElementById("txt-custom-endpoint").value = ai.customEndpoint || "";
  document.getElementById("chk-offline-mode").checked = !!ai.offlineMode;

  const selGeminiModel = document.getElementById("sel-gemini-model");
  // Якщо збережено стару 2.0, перемикаємо на 3.6
  if (!ai.model || ai.model.includes("2.0")) {
    selGeminiModel.value = "gemini-3.6-flash";
  } else {
    selGeminiModel.value = ai.model;
  }
}

function showToast(message) {
  const toast = document.getElementById("toast");
  if (!toast) return;
  toast.textContent = message;
  toast.classList.remove("hidden");
  setTimeout(() => {
    toast.classList.add("hidden");
  }, 2600);
}
