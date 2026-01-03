
import { GoogleGenAI, Modality, Type, GenerateContentResponse } from "@google/genai";
import { LabData } from "../types";

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

export async function getSearchGrounding(query: string) {
  const ai = getAI();
  const response = await ai.models.generateContent({
    model: 'gemini-3-flash-preview',
    contents: query,
    config: {
      tools: [{ googleSearch: {} }],
    },
  });

  return {
    text: response.text,
    sources: response.candidates?.[0]?.groundingMetadata?.groundingChunks || []
  };
}

export async function generateSpeech(text: string): Promise<string | undefined> {
  const ai = getAI();
  const response = await ai.models.generateContent({
    model: "gemini-2.5-flash-preview-tts",
    contents: [{ parts: [{ text }] }],
    config: {
      responseModalities: [Modality.AUDIO],
      speechConfig: {
        voiceConfig: {
          prebuiltVoiceConfig: { voiceName: 'Kore' },
        },
      },
    },
  });

  return response.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data;
}

export async function transcribeAudio(base64Audio: string): Promise<string> {
  const ai = getAI();
  const response = await ai.models.generateContent({
    model: 'gemini-3-flash-preview',
    contents: {
      parts: [
        { inlineData: { data: base64Audio, mimeType: 'audio/wav' } },
        { text: 'Transcribe this audio. Return only the numbers found or clear text.' }
      ]
    }
  });
  return response.text || '';
}
