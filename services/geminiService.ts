
import { GoogleGenAI, Modality, Type, GenerateContentResponse } from "@google/genai";
import { LabData, CalculationResult } from "../types";

const getAI = () => new GoogleGenAI({ apiKey: process.env.API_KEY });

export async function extractLabDataFromImage(base64Image: string): Promise<Partial<LabData>> {
  const ai = getAI();
  const response = await ai.models.generateContent({
    model: 'gemini-3-pro-image-preview',
    contents: {
      parts: [
        { inlineData: { data: base64Image, mimeType: 'image/jpeg' } },
        { text: 'Analyze this image and extract laboratory data for clay testing. Look for: Smectite content (содержание смектита), Cation Exchange Capacity (обменная емкость/КОЕ), Humidity/Water content (влажность), and rheometer readings at 300 and 600 RPM (Фи 300, Фи 600). Return ONLY a JSON object with keys "m", "q", "w", "f300", "f600" and their numeric values. If a value is missing, use null.' }
      ]
    },
    config: {
      responseMimeType: "application/json"
    }
  });

  try {
    return JSON.parse(response.text || '{}');
  } catch (e) {
    console.error("Failed to parse image analysis result", e);
    return {};
  }
}

export async function getBentoniteConclusions(results: CalculationResult): Promise<string[]> {
  const ai = getAI();
  const prompt = `Проведи экспертный анализ бентонита для органомодификации на основе данных и критериев из тех. регламента:
  ДАННЫЕ:
  - Содержание смектита (m): ${results.m}%
  - Обменная емкость (q/КОЕ): ${results.q} мг-экв/100г
  - YP/PV отношение: ${results.ypPvRatio.toFixed(2)}
  - Изотропия: ${results.isotropy.toFixed(4)}
  - Критерий генерации: ${results.generation.toFixed(2)}
  - Критерий загущения: ${results.thickening.toFixed(2)}
  - Критерий полноты: ${results.completeness.toFixed(2)}

  ЭКСПЕРТНЫЕ КРИТЕРИИ ИЗ ДОКУМЕНТА:
  1. YP/PV: >6 - вероятно активирован содой (сомнительно для органо), 3-6 - класс OCMA (ограниченно), 1.5-3 - класс Drilling grade (подходит), <1.5 - non treated.
  2. Изотропия: Приемлемо >= 0.24.
  3. Генерация: Приемлемо > 8.5.
  4. Загущение: Хорошо < 1, Приемлемо 1-1.3.
  5. Полнота: Приемлемо 100-115, Хорошо > 115.

  ЗАДАЧА: Сформулируй ровно 5 экспертных выводов о качестве и пригодности этой глины для органомодификации катионными ПАВ. Обоснуй каждый пункт конкретными значениями.
  Верни ответ в формате JSON: {"conclusions": ["вывод 1", "вывод 2", "вывод 3", "вывод 4", "вывод 5"]}`;

  const response = await ai.models.generateContent({
    model: 'gemini-3-flash-preview',
    contents: prompt,
    config: {
      responseMimeType: "application/json"
    }
  });

  try {
    const data = JSON.parse(response.text || '{"conclusions":[]}');
    return data.conclusions || [];
  } catch (e) {
    return ["Ошибка анализа данных. Пожалуйста, проверьте ввод."];
  }
}

export async function generateThematicImage(topic: string): Promise<string | null> {
  const ai = getAI();
  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash-image',
      contents: {
        parts: [{ text: `A professional, high-quality laboratory setting image showing ${topic} related to clay mineralogy, bentonite, and chemical analysis. Photorealistic style, scientific atmosphere, clean lab equipment.` }]
      },
      config: {
        imageConfig: {
          aspectRatio: "16:9"
        }
      }
    });

    for (const part of response.candidates?.[0]?.content?.parts || []) {
      if (part.inlineData) {
        return `data:image/png;base64,${part.inlineData.data}`;
      }
    }
  } catch (e) {
    console.error("Image generation failed", e);
  }
  return null;
}
