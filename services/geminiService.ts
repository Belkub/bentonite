import { GoogleGenAI, Type } from "@google/genai";
import { LabData, CalculationResult, Conclusion } from "../types.ts";

// Extract laboratory data from an image of a lab report
export async function extractLabDataFromImage(base64Image: string): Promise<Partial<LabData>> {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  try {
    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: {
        parts: [
          { inlineData: { data: base64Image, mimeType: 'image/jpeg' } },
          { text: 'Analyze this image and extract laboratory data for clay testing. Return ONLY a JSON object.' }
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
  - m: ${results.m}%
  - q: ${results.q}
  - YP/PV: ${results.ypPvRatio.toFixed(2)}
  - Изотропия: ${results.isotropy.toFixed(4)}
  - Генерация: ${results.generation.toFixed(2)}
  - Загущение: ${results.thickening.toFixed(2)}
  - Полнота: ${results.completeness.toFixed(2)}`;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview',
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
                  sentiment: { type: Type.STRING }
                },
                required: ["text", "sentiment"]
              }
            }
          },
          required: ["conclusions"]
        }
      }
    });

    const data = JSON.parse(response.text?.trim() || '{"conclusions":[]}');
    return data.conclusions || [];
  } catch (e) {
    console.error("Failed to get conclusions from Gemini:", e);
    throw e;
  }
}