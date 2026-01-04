
import { GoogleGenAI, Type } from "@google/genai";
import { LabData, CalculationResult, Conclusion } from "../types";

// Extract laboratory data from an image of a lab report
export async function extractLabDataFromImage(base64Image: string): Promise<Partial<LabData>> {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  try {
    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: {
        parts: [
          { inlineData: { data: base64Image, mimeType: 'image/jpeg' } },
          { text: 'Analyze this image and extract laboratory data for clay testing. Look for: Smectite content (содержание смектита), Cation Exchange Capacity (обменная емкость/КОЕ), Humidity/Water content (влажность), and rheometer readings at 300 and 600 RPM (Фи 300, Фи 600). Return ONLY a JSON object with keys "m", "q", "w", "f300", "f600" and their numeric values. If a value is missing, use null.' }
        ]
      },
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            m: { type: Type.NUMBER },
            q: { type: Type.NUMBER },
            w: { type: Type.NUMBER },
            f300: { type: Type.NUMBER },
            f600: { type: Type.NUMBER }
          }
        }
      }
    });

    // Use .text property directly and trim it
    return JSON.parse(response.text?.trim() || '{}');
  } catch (e) {
    console.error("Failed to extract lab data from image:", e);
    return {};
  }
}

// Get expert conclusions based on calculated lab results
export async function getBentoniteConclusions(results: CalculationResult): Promise<Conclusion[]> {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const prompt = `Проведи строгий экспертный анализ бентонита.
  ДАННЫЕ:
  - Содержание смектита (m): ${results.m}%
  - Обменная емкость (q/КОЕ): ${results.q} мг-экв/100г
  - YP/PV отношение: ${results.ypPvRatio.toFixed(2)}
  - Изотропия: ${results.isotropy.toFixed(4)}
  - Критерий генерации: ${results.generation.toFixed(2)}
  - Критерий загущения: ${results.thickening.toFixed(2)}
  - Критерий полноты: ${results.completeness.toFixed(2)}

  КРИТЕРИИ ОЦЕНКИ (будь критичен!):
  1. YP/PV: <1.5 или >6 часто указывает на проблемы (необработанный или пересоленный).
  2. Изотропия: < 0.24 — это ПЛОХО (negative).
  3. Генерация: < 8.5 — это НИЗКИЙ потенциал (negative).
  4. Загущение: > 1.3 — слишком высокое, возможны трудности (negative).
  5. Полнота: < 100 — недостаточное качество (negative).

  ЗАДАЧА: Сформулируй 5 выводов. 
  Для каждого вывода ОПАСНОСТЬ или НЕДОСТАТОК классифицируй ТОЛЬКО как "negative".
  Хорошие показатели — "positive". Спорные — "neutral".
  Используй только эти три слова для sentiment.`;

  try {
    const response = await ai.models.generateContent({
      // Using pro model for complex expert reasoning task
      model: 'gemini-3-pro-preview',
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            conclusions: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  text: { type: Type.STRING },
                  sentiment: { type: Type.STRING, description: "Strictly 'positive', 'neutral', or 'negative'" }
                },
                required: ["text", "sentiment"]
              }
            }
          },
          required: ["conclusions"]
        }
      }
    });

    // Use .text property directly and trim it
    const data = JSON.parse(response.text?.trim() || '{"conclusions":[]}');
    return data.conclusions || [];
  } catch (e) {
    console.error("Failed to get conclusions from Gemini:", e);
    throw e;
  }
}

export async function generateThematicImage(topic: string): Promise<string | null> {
  return null;
}
