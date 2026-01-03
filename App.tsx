
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { GoogleGenAI, Modality, LiveServerMessage } from "@google/genai";
import { LabData, LabKey, LAB_LABELS, LAB_ORDER, CalculationResult } from './types';
import { 
  extractLabDataFromImage 
} from './services/geminiService';
import { decode, decodeAudioData, createBlob } from './utils/audio';
import { parseSpokenNumber } from './utils/voiceParser';

const App: React.FC = () => {
  const [labData, setLabData] = useState<LabData>({
    m: '', q: '', w: '', f300: '', f600: '',
  });
  const [currentStep, setCurrentStep] = useState(0);
  const [showResults, setShowResults] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  
  // Храним актуальное состояние для Live API коллбэков
  const stateRef = useRef({ currentStep, labData, showResults });
  useEffect(() => {
    stateRef.current = { currentStep, labData, showResults };
  }, [currentStep, labData, showResults]);

  // Состояния для Live API
  const [isLiveActive, setIsLiveActive] = useState(false);
  const [liveTranscription, setLiveTranscription] = useState('');
  
  const sessionRef = useRef<any>(null);
  const audioContextInRef = useRef<AudioContext | null>(null);
  const audioContextOutRef = useRef<AudioContext | null>(null);
  const nextStartTimeRef = useRef(0);
  const sourcesRef = useRef<Set<AudioBufferSourceNode>>(new Set());

  // Накопитель транскрипции для текущего "хода" (turn)
  const turnTranscriptionRef = useRef<string>('');

  const handleInputChange = (key: LabKey, value: string) => {
    const standardized = value.replace(',', '.');
    if (/^[0-9.]*$/.test(standardized)) {
      setLabData(prev => ({ ...prev, [key]: standardized }));
    }
  };

  const results = useMemo((): CalculationResult | null => {
    const { m, q, w, f300, f600 } = labData;
    const parse = (val: string) => parseFloat(val) || 0;
    const isValid = [m, q, w, f300, f600].every(val => val.trim() !== '' && !isNaN(parseFloat(val)));
    if (!isValid) return null;

    const m_val = parse(m);
    const q_val = parse(q);
    const w_val = parse(w);
    const f3_val = parse(f300);
    const f6_val = parse(f600);
    
    const pv = f6_val - f3_val;
    const yp = f3_val - pv;
    const f = f6_val / 2;

    // Advanced Calculations
    const poe = q_val / (1 - 0.01 * w_val);
    const ypPvRatio = pv !== 0 ? yp / pv : 0;
    const s = f6_val !== 0 ? yp / f6_val : 0;
    
    // Технологические критерии
    
    // 1. Изотропия: (0.5 - ((YP/f600) - 0.5)^2) * m * m * 0.01 * 0.01
    const isotropy = (f6_val !== 0) 
      ? (0.5 - Math.pow(((yp / f6_val) - 0.5), 2)) * m_val * m_val * 0.01 * 0.01 
      : 0;

    // 2. Генерация: f600 * (1 - (YP/f600)) * m / q
    const generation = (q_val !== 0 && f6_val !== 0) 
      ? f6_val * (1 - (yp / f6_val)) * m_val / q_val 
      : 0;
    
    // 3. Загущение: f600 / (0.01 * 0.01 * m^2 * q)
    const thickening = (m_val !== 0 && q_val !== 0) ? f6_val / (0.01 * 0.01 * m_val * m_val * q_val) : 0;
    
    // 4. Полнота: q * 0.01 * m * 0.01 * m * LOG10(f600 * 0.001)^2 / (YP / f600)
    // s = YP / f600
    const logArg = f6_val * 0.001;
    const logVal = logArg > 0 ? Math.log10(logArg) : 0;
    const completeness = (s !== 0) 
      ? (q_val * 0.01 * m_val * 0.01 * m_val * Math.pow(logVal, 2)) / s 
      : 0;
    
    return { 
      m: m_val, q: q_val, w: w_val, f, pv, yp,
      poe, ypPvRatio, s, isotropy, generation, thickening, completeness
    };
  }, [labData]);

  const getSystemInstruction = () => {
    let instr = `Вы — эксперт GeoLab. `;
    if (showResults && results) {
      instr += `АНАЛИЗ ГОТОВ. 
      Результаты: Смектит=${results.m}, КОЕ=${results.q}, ПОЕ=${results.poe.toFixed(2)}, YP/PV=${results.ypPvRatio.toFixed(2)}, s=${results.s.toFixed(2)}.
      Критерии: Изотропия=${results.isotropy.toFixed(4)}, Генерация=${results.generation.toFixed(2)}, Загущение=${results.thickening.toFixed(2)}, Полнота=${results.completeness.toFixed(2)}.
      Обсудите эти показатели, дайте экспертную оценку качеству глины и рекомендации.`;
    } else {
      instr += `Ждем ввод для: ${LAB_LABELS[LAB_ORDER[stateRef.current.currentStep]]}. Принимайте числа полностью.`;
    }
    return instr;
  };

  const startLiveSession = async () => {
    if (isLiveActive) { stopLiveSession(); return; }
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
      audioContextInRef.current = new AudioContext({ sampleRate: 16000 });
      audioContextOutRef.current = new AudioContext({ sampleRate: 24000 });
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      
      const sessionPromise = ai.live.connect({
        model: 'gemini-2.5-flash-native-audio-preview-09-2025',
        config: {
          responseModalities: [Modality.AUDIO],
          inputAudioTranscription: {},
          tools: [{ googleSearch: {} }],
          systemInstruction: getSystemInstruction(),
          speechConfig: { voiceConfig: { prebuiltVoiceConfig: { voiceName: 'Kore' } } }
        },
        callbacks: {
          onopen: () => {
            setIsLiveActive(true);
            const source = audioContextInRef.current!.createMediaStreamSource(stream);
            const scriptProcessor = audioContextInRef.current!.createScriptProcessor(4096, 1, 1);
            scriptProcessor.onaudioprocess = (e) => {
              const inputData = e.inputBuffer.getChannelData(0);
              sessionPromise.then(s => s.sendRealtimeInput({ media: createBlob(inputData) }));
            };
            source.connect(scriptProcessor);
            scriptProcessor.connect(audioContextInRef.current!.destination);
          },
          onmessage: async (msg: LiveServerMessage) => {
            const audioData = msg.serverContent?.modelTurn?.parts[0]?.inlineData?.data;
            if (audioData && audioContextOutRef.current) {
              const ctx = audioContextOutRef.current;
              nextStartTimeRef.current = Math.max(nextStartTimeRef.current, ctx.currentTime);
              const buffer = await decodeAudioData(decode(audioData), ctx, 24000, 1);
              const source = ctx.createBufferSource();
              source.buffer = buffer;
              source.connect(ctx.destination);
              source.start(nextStartTimeRef.current);
              nextStartTimeRef.current += buffer.duration;
              sourcesRef.current.add(source);
            }

            if (msg.serverContent?.inputTranscription) {
              const text = msg.serverContent.inputTranscription.text;
              turnTranscriptionRef.current += " " + text;
              setLiveTranscription(turnTranscriptionRef.current);
              
              const parsed = parseSpokenNumber(turnTranscriptionRef.current);
              if (parsed !== null) {
                const key = LAB_ORDER[stateRef.current.currentStep];
                setLabData(prev => ({ ...prev, [key]: parsed.toString() }));
              }
            }

            if (msg.serverContent?.turnComplete) {
              const finalVal = parseSpokenNumber(turnTranscriptionRef.current);
              if (finalVal !== null) {
                const idx = stateRef.current.currentStep;
                if (idx < 4) setCurrentStep(idx + 1);
                else setShowResults(true);
              }
              turnTranscriptionRef.current = '';
              setLiveTranscription('');
            }

            if (msg.serverContent?.interrupted) {
              sourcesRef.current.forEach(s => s.stop());
              sourcesRef.current.clear();
              nextStartTimeRef.current = 0;
            }
          },
          onclose: () => setIsLiveActive(false),
          onerror: () => setIsLiveActive(false)
        }
      });
      sessionRef.current = await sessionPromise;
    } catch (e) { setIsLiveActive(false); }
  };

  const stopLiveSession = () => {
    sessionRef.current?.close();
    sessionRef.current = null;
    setIsLiveActive(false);
  };

  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);
    const reader = new FileReader();
    reader.onloadend = async () => {
      const base64 = (reader.result as string).split(',')[1];
      const ext = await extractLabDataFromImage(base64);
      const normalized: Partial<LabData> = {};
      Object.keys(ext).forEach(k => {
        const val = ext[k as keyof LabData];
        if (val !== null && val !== undefined) normalized[k as keyof LabData] = val.toString();
      });
      setLabData(prev => ({ ...prev, ...normalized }));
      setIsUploading(false);
    };
    reader.readAsDataURL(file);
  };

  const isDataValid = !!results;

  return (
    <div className="min-h-screen bg-slate-50 p-3 md:p-5 text-slate-900 font-sans">
      <div className="max-w-4xl mx-auto space-y-3">
        
        {/* Компактный хедер */}
        <header className="flex items-center justify-between bg-white px-4 py-3 rounded-xl shadow-sm border border-slate-100">
          <div className="flex items-center gap-2">
            <div className={`w-2 h-2 rounded-full ${isLiveActive ? 'bg-red-500 animate-pulse' : 'bg-indigo-600'}`}></div>
            <h1 className="text-lg font-black tracking-tighter">GeoLab<span className="text-indigo-600">Lite</span></h1>
          </div>
          <div className="flex gap-2">
            <button onClick={startLiveSession} className={`px-3 py-1.5 rounded-lg font-bold text-[10px] uppercase tracking-wider transition-all ${isLiveActive ? 'bg-red-500 text-white shadow-lg' : 'bg-indigo-600 text-white shadow-md'}`}>
              {isLiveActive ? 'Стоп Голос' : 'Голос'}
            </button>
            <label className="cursor-pointer bg-slate-100 text-slate-600 px-3 py-1.5 rounded-lg font-bold text-[10px] uppercase transition-all">
              {isUploading ? '...' : 'Фото'}
              <input type="file" className="hidden" accept="image/*" onChange={handleImageUpload} />
            </label>
          </div>
        </header>

        {/* Статус транскрипции */}
        {isLiveActive && (
          <div className="bg-slate-900 text-white px-4 py-2 rounded-lg shadow-inner animate-in fade-in slide-in-from-top-1">
            <p className="text-[11px] text-indigo-300 font-medium italic truncate">
              {liveTranscription || "Слушаю... Назовите число целиком"}
            </p>
          </div>
        )}

        {/* Компактная сетка параметров */}
        <div className="grid grid-cols-5 gap-1.5">
          {LAB_ORDER.map((key, index) => (
            <button
              key={key}
              onClick={() => setCurrentStep(index)}
              className={`p-2 rounded-lg border text-center transition-all ${
                currentStep === index ? 'bg-indigo-600 border-indigo-600 text-white shadow-md' : 
                labData[key] ? 'bg-white border-emerald-200 text-slate-800' : 'bg-white border-slate-200 opacity-60'
              }`}
            >
              <div className={`text-[7px] font-black uppercase mb-0.5 truncate ${currentStep === index ? 'text-indigo-100' : 'text-slate-400'}`}>
                {LAB_LABELS[key].split(' ')[0]}
              </div>
              <div className="text-sm font-mono font-bold leading-none">{labData[key] || '0'}</div>
            </button>
          ))}
        </div>

        {/* Окно активного ввода */}
        {!showResults && (
          <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 relative overflow-hidden">
            <div className="absolute top-0 left-0 w-full h-1 bg-slate-100">
              <div className="h-full bg-indigo-600 transition-all duration-500" style={{ width: `${((currentStep + 1) / 5) * 100}%` }}></div>
            </div>
            <div className="text-center mb-4">
              <h2 className="text-base font-black text-slate-500 uppercase tracking-widest">{LAB_LABELS[LAB_ORDER[currentStep]]}</h2>
            </div>
            <div className="max-w-[200px] mx-auto relative mb-4">
              <input
                type="text"
                inputMode="decimal"
                className="w-full text-center text-4xl font-mono font-black py-3 bg-slate-50 border-b-4 border-slate-100 focus:border-indigo-600 outline-none transition-all rounded-xl text-slate-900"
                value={labData[LAB_ORDER[currentStep]]}
                onChange={(e) => handleInputChange(LAB_ORDER[currentStep], e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && (currentStep < 4 ? setCurrentStep(currentStep + 1) : setShowResults(true))}
                autoFocus
              />
            </div>
            <div className="flex justify-center gap-4">
              {currentStep > 0 && <button onClick={() => setCurrentStep(currentStep - 1)} className="text-[10px] font-bold text-slate-400 uppercase">Назад</button>}
              <button onClick={() => (currentStep < 4 ? setCurrentStep(currentStep + 1) : setShowResults(true))} className="bg-indigo-600 text-white px-8 py-2 rounded-xl font-black text-xs uppercase shadow-lg">
                {currentStep === 4 ? 'Готово' : 'Далее'}
              </button>
            </div>
          </div>
        )}

        {/* Компактные результаты */}
        {showResults && (
          <div className={`rounded-2xl border-2 overflow-hidden animate-in zoom-in-95 duration-300 ${isDataValid ? 'border-emerald-500/30' : 'border-amber-500/30'}`}>
            <div className={`px-4 py-3 text-white flex justify-between items-center ${isDataValid ? 'bg-emerald-600' : 'bg-amber-600'}`}>
              <h3 className="text-sm font-black uppercase tracking-widest">Результаты анализа</h3>
              <div className="flex gap-2">
                <button onClick={() => setShowResults(false)} className="text-[9px] bg-black/20 px-2 py-1 rounded-md font-black">ПРАВИТЬ ШАГИ</button>
                {isLiveActive && <button onClick={() => { stopLiveSession(); setTimeout(startLiveSession, 100); }} className="text-[9px] bg-white/20 px-2 py-1 rounded-md border border-white/20 font-black">ОБСУДИТЬ</button>}
              </div>
            </div>
            
            <div className="bg-white p-4 space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* Editable Inputs */}
                <div className="space-y-1.5">
                  <p className="text-[8px] font-black text-slate-400 uppercase mb-1">Исходные данные (редактируемые)</p>
                  {LAB_ORDER.map((key) => (
                    <div key={key} className="flex justify-between items-center px-3 py-1.5 bg-slate-50 rounded-lg border border-slate-100 focus-within:border-indigo-300 transition-colors">
                      <span className="text-[10px] font-bold text-slate-500">{LAB_LABELS[key]}</span>
                      <input 
                        type="text"
                        className="text-xs font-mono font-black text-slate-800 bg-transparent text-right outline-none w-20"
                        value={labData[key]}
                        onChange={(e) => handleInputChange(key, e.target.value)}
                      />
                    </div>
                  ))}
                  {/* Additional Derived Calculation POE */}
                  <div className="flex justify-between items-center px-3 py-1.5 bg-indigo-50/50 rounded-lg border border-indigo-100">
                    <span className="text-[10px] font-bold text-indigo-600">ПОЕ (полная обменная емк.)</span>
                    <span className="text-xs font-mono font-black text-indigo-700">{results?.poe.toFixed(2) || '---'}</span>
                  </div>
                </div>

                {/* Main Results */}
                <div className="space-y-2">
                  <p className="text-[8px] font-black text-slate-400 uppercase mb-1">Реологические параметры</p>
                  {[
                    { label: 'F (Конст)', val: results?.f.toFixed(2), color: 'text-emerald-600', formula: 'f600/2' },
                    { label: 'PV (Вязк)', val: results?.pv.toFixed(2), color: 'text-indigo-600', formula: 'f600-f300' },
                    { label: 'YP (Текуч)', val: results?.yp.toFixed(2), color: 'text-rose-600', formula: 'f300-PV' },
                    { label: 'YP/PV', val: results?.ypPvRatio.toFixed(2), color: 'text-slate-600', formula: 'YP/PV' },
                    { label: 's (Степень зам.)', val: results?.s.toFixed(2), color: 'text-slate-600', formula: 'YP/f600' },
                  ].map((item, i) => (
                    <div key={i} className="flex justify-between items-center p-2.5 bg-white rounded-xl border border-slate-100 shadow-sm">
                      <div>
                        <span className="font-black text-[10px] text-slate-900 block leading-none">{item.label}</span>
                        <span className="text-[7px] font-mono text-slate-400">{item.formula}</span>
                      </div>
                      <span className={`font-mono text-lg font-black ${item.color}`}>{item.val || '---'}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Criteria Section - Light Blue Background */}
              <div className="bg-sky-50 p-4 rounded-2xl border border-sky-100 space-y-3">
                <p className="text-[9px] font-black text-sky-600 uppercase tracking-widest text-center mb-2">Технологические критерии</p>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3">
                  {[
                    { label: 'Изотропия', val: results?.isotropy.toFixed(4), desc: 'Структурная стабильность' },
                    { label: 'Генерация', val: results?.generation.toFixed(2), desc: 'Эффективность фаз' },
                    { label: 'Загущение', val: results?.thickening.toFixed(2), desc: 'Реакция на электролиты' },
                    { label: 'Полнота', val: results?.completeness.toFixed(2), desc: 'Качество очистки' },
                  ].map((crit, idx) => (
                    <div key={idx} className="bg-white p-3 rounded-xl border border-sky-200 shadow-sm flex flex-col items-center">
                      <span className="text-[9px] font-black text-slate-500 uppercase">{crit.label}</span>
                      <span className="text-xl font-mono font-black text-sky-700 my-1">{crit.val || '---'}</span>
                      <span className="text-[7px] text-slate-400 text-center leading-tight">{crit.desc}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default App;
