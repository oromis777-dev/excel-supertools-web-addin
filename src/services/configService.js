export const ConfigService = {
  defaults: {
    focusGrid: {
      enabled: false,
      shapeType: 0, // 0: Crosshair (Перехрестя), 1: Row (Рядок), 2: Col (Стовпець)
      styleType: 0, // 0: Fill (Заливка), 1: Border (Рамка), 2: Line (Тонка лінія)
      colorHex: "#417EBC",
      transparency: 40, // 0% до 90%
      lineThickness: 2, // 1: Тонка, 2: Середня, 3: Товста
      highlightActiveCell: true,
      autoFollow: true
    },
    ai: {
      provider: "gemini",
      geminiApiKey: "",
      openaiApiKey: "",
      deepseekApiKey: "",
      customEndpoint: "http://localhost:11434/api/generate",
      model: "gemini-3.6-flash",
      offlineMode: false,
      systemLanguage: "uk"
    },
    copyRanges: {
      skipHidden: true,
      pasteMode: "Values", // "Values" | "Formulas" | "All"
      preserveDimensions: true
    }
  },

  async loadConfig() {
    try {
      if (typeof localStorage !== "undefined") {
        const stored = localStorage.getItem("excel_supertools_config");
        if (stored) {
          const parsed = JSON.parse(stored);
          const aiConfig = { ...this.defaults.ai, ...(parsed.ai || {}) };
          // Автоматично оновлюємо застарілу модель 2.0 на 3.6
          if (aiConfig.model && aiConfig.model.includes("2.0")) {
            aiConfig.model = "gemini-3.6-flash";
          }
          return {
            focusGrid: { ...this.defaults.focusGrid, ...(parsed.focusGrid || {}) },
            ai: aiConfig,
            copyRanges: { ...this.defaults.copyRanges, ...(parsed.copyRanges || {}) }
          };
        }
      }
    } catch (e) {
      console.warn("Could not load config from localStorage", e);
    }
    return JSON.parse(JSON.stringify(this.defaults));
  },

  async saveConfig(config) {
    try {
      if (typeof localStorage !== "undefined") {
        localStorage.setItem("excel_supertools_config", JSON.stringify(config));
      }
    } catch (e) {
      console.error("Could not save config to localStorage", e);
    }
  }
};
