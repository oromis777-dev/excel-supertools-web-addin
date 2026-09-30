import { ConfigService } from "./configService.js";

export const AIService = {
  async executePrompt(promptText, contextData = "") {
    const config = (await ConfigService.loadConfig()).ai;

    if (config.offlineMode) {
      return this._generateOfflineResponse(promptText, contextData);
    }

    const fullPrompt = contextData 
      ? `Контекст даних з таблиці Excel:\n${contextData}\n\nЗавдання користувача:\n${promptText}`
      : promptText;

    switch (config.provider) {
      case "gemini":
        return await this._callGeminiAPI(fullPrompt, config.geminiApiKey, config.model || "gemini-3.6-flash");
      case "openai":
        return await this._callOpenAI(fullPrompt, config.openaiApiKey, config.model || "gpt-4o-mini");
      case "deepseek":
        return await this._callDeepSeek(fullPrompt, config.deepseekApiKey);
      case "custom":
        return await this._callCustomEndpoint(fullPrompt, config.customEndpoint);
      default:
        throw new Error(`Непідтримуваний AI провайдер: ${config.provider}`);
    }
  },

  async _callGeminiAPI(prompt, apiKey, model = "gemini-3.6-flash") {
    if (!apiKey) {
      throw new Error("Вкажіть API-ключ Google Gemini у налаштуваннях надбудови (вкладка ⚙️ Налаштування).");
    }

    // Автоматична міграція з відключених моделей 2.0 на 3.6
    let targetModel = model;
    if (!targetModel || targetModel.includes("2.0") || targetModel.includes("2.5")) {
      targetModel = "gemini-3.6-flash";
    }

    const requestBody = {
      contents: [{
        parts: [{ text: prompt }]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 2048
      }
    };

    const tryModel = async (modelName) => {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
      return await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
      });
    };

    let response;
    try {
      response = await tryModel(targetModel);
      // Якщо обрана модель повернула 404, виконуємо автоматичний відкат на стабільні моделі
      if (response.status === 404 && targetModel !== "gemini-3.6-flash") {
        console.warn(`[AI Service] Модель ${targetModel} недоступна (404). Спроба через gemini-3.6-flash...`);
        response = await tryModel("gemini-3.6-flash");
      }
      if (response.status === 404 && targetModel !== "gemini-1.5-flash") {
        console.warn(`[AI Service] Модель gemini-3.6-flash повернула 404. Резервний відкат на gemini-1.5-flash...`);
        response = await tryModel("gemini-1.5-flash");
      }
    } catch (netErr) {
      throw new Error(`Мережева помилка підключення до Gemini API: ${netErr.message}`);
    }

    if (!response.ok) {
      const errText = await response.text();
      let msg = errText;
      try {
        const errJson = JSON.parse(errText);
        msg = errJson.error?.message || errText;
      } catch (e) {}
      throw new Error(`Gemini API помилка [${response.status}]: ${msg}`);
    }

    const data = await response.json();
    if (data.candidates && data.candidates[0]?.content?.parts[0]?.text) {
      return data.candidates[0].content.parts[0].text;
    }
    return "Отримано порожню відповідь від моделі Gemini.";
  },

  async _callOpenAI(prompt, apiKey, model = "gpt-4o-mini") {
    if (!apiKey) throw new Error("Вкажіть API-ключ OpenAI у налаштуваннях надбудови.");
    const url = "https://api.openai.com/v1/chat/completions";
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0.3
      })
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`OpenAI API помилка [${response.status}]: ${err}`);
    }
    const data = await response.json();
    return data.choices?.[0]?.message?.content || "Порожня відповідь від OpenAI.";
  },

  async _callDeepSeek(prompt, apiKey) {
    if (!apiKey) throw new Error("Вкажіть API-ключ DeepSeek у налаштуваннях надбудови.");
    const url = "https://api.deepseek.com/chat/completions";
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: "deepseek-chat",
        messages: [{ role: "user", content: prompt }]
      })
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`DeepSeek API помилка [${response.status}]: ${err}`);
    }
    const data = await response.json();
    return data.choices?.[0]?.message?.content || "Порожня відповідь від DeepSeek.";
  },

  async _callCustomEndpoint(prompt, endpoint) {
    if (!endpoint) throw new Error("Вкажіть Endpoint URL для власного/локального AI.");
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        prompt: prompt,
        stream: false
      })
    });

    if (!response.ok) {
      throw new Error(`Custom AI Endpoint помилка [${response.status}]`);
    }
    const data = await response.json();
    return data.response || data.text || JSON.stringify(data);
  },

  _generateOfflineResponse(prompt, context) {
    return `[ОФЛАЙН РЕЖИМ]\n` +
           `Використано вбудований шаблон правил без підключення до Інтернету.\n\n` +
           `Аналіз запиту: "${prompt}"\n` +
           `Кількість переданих даних: ${context ? context.length + " символів" : "контекст порожній"}.\n\n` +
           `Порада: Для формул обчислення підсумків використовуйте: =SUM(діапазон) або =AGGREGATE(9, 5, діапазон).\n` +
           `Для пошуку значень рекомендується сучасна функція =XLOOKUP(шукане, де_шукати, що_повернути).`;
  },

  async generateFormula(description) {
    const prompt = `Ти експерт з формул Microsoft Excel. На основі наступного запиту українською або англійською мовою:\n` +
                   `"${description}"\n\n` +
                   `Сформулюй готову формулу Excel (починаючи з =). Назви функцій використовуй англійською мовою (наприклад, SUM, VLOOKUP, INDEX, MATCH, XLOOKUP, IF, COUNTIF).\n` +
                   `Надай коротке і зрозуміле пояснення аргументів українською мовою.`;
    return await this.executePrompt(prompt);
  },

  async explainFormula(formulaText) {
    const prompt = `Поясни простими словами логіку роботи наступної формули Excel:\n` +
                   `${formulaText}\n\n` +
                   `1. Що вона робить.\n` +
                   `2. Які функції та діапазони задіяні.\n` +
                   `3. За яких умов вона може повернути помилку (#N/A, #VALUE! тощо) і як їй запобігти.`;
    return await this.executePrompt(prompt);
  },

  async analyzeData(sheetDataSummary) {
    const prompt = `Ти провідний аналітик даних для бізнесу. Проаналізуй наступну таблицю даних з Excel:\n` +
                   `${sheetDataSummary}\n\n` +
                   `Структуруй відповідь наступним чином:\n` +
                   `1. 📊 Загальний огляд даних (структура, обсяг, ключові стовпці).\n` +
                   `2. 🔍 Виявлені аномалії, пропуски або помилки.\n` +
                   `3. 💡 Ключові бізнес-висновки та патерни.\n` +
                   `4. 🎯 Практичні рекомендації для керівництва.`;
    return await this.executePrompt(prompt);
  }
};
