
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { GoogleGenAI, Modality, LiveServerMessage } from "@google/genai";
import { LabData, LabKey, LAB_LABELS, LAB_ORDER, CalculationResult, Conclusion, SavedChart } from './types';
import { 
  extractLabDataFromImage,
  getBentoniteConclusions
} from './services/geminiService';
import { decode, decodeAudioData, createBlob } from './utils/audio';
import { parseSpokenNumber } from './utils/voiceParser';

// Библиотеки для экспорта
import * as docx from 'docx';
import PptxGenJS from 'pptxgenjs';

declare var Plotly: any;

// The environment already provides aistudio on the global window object.
// We use type casting to (window as any) to access its members to avoid redeclaring conflicts.

type PPTStyle = {
  id: string;
  name: string;
  primary: string;
  secondary: string;
  text: string;
  bg: string;
};

const PPT_STYLES: PPTStyle[] = [
  { id: 'indigo', name: 'Geo Blue', primary: '#4f46e5', secondary: '#818cf8', text: '#ffffff', bg: '#f8fafc' },
  { id: 'emerald', name: 'Lab Green', primary: '#059669', secondary: '#34d399', text: '#ffffff', bg: '#f0fdf4' },
  { id: 'dark', name: 'Industrial', primary: '#1e293b', secondary: '#f97316', text: '#ffffff', bg: '#0f172a' },
  { id: 'academic', name: 'Classic', primary: '#334155', secondary: '#cbd5e1', text: '#000000', bg: '#ffffff' },
];

const CHART_VARIABLES = {
  f600: { label: 'Фи 600', min: 15, max: 120 },
  yp: { label: 'YP', min: 5, max: 60 },
  m: { label: 'Смектит (m)', min: 40, max: 100 },
  q: { label: 'КОЕ (q)', min: 40, max: 120 },
  yp_f600: { label: 'YP/f600', min: 0.05, max: 1.0 }
};

const App: React.FC = () => {
  const [labData, setLabData] = useState<LabData>({
    m: '', q: '', w: '', f300: '', f600: '', s_equiv: '', mm: ''
  });
  const [currentStep, setCurrentStep] = useState(0);
  const [showResults, setShowResults] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [conclusions, setConclusions] = useState<Conclusion[]>([]);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisError, setAnalysisError] = useState<string | null>(null);
  const [showStylePicker, setShowStylePicker] = useState(false);
  const [isGeneratingDoc, setIsGeneratingDoc] = useState(false);
  
  const [showChartModal, setShowChartModal] = useState(false);
  const [axisX, setAxisX] = useState<keyof typeof CHART_VARIABLES>('f600');
  const [axisY, setAxisY] = useState<keyof typeof CHART_VARIABLES>('m');
  const [currentChartImage, setCurrentChartImage] = useState<string | null>(null);
  const [savedCharts, setSavedCharts] = useState<SavedChart[]>([]);
  const chartRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const checkKey = async () => {
      const aistudio = (window as any).aistudio;
      if (aistudio && typeof aistudio.hasSelectedApiKey === 'function') {
        try {
          const hasKey = await aistudio.hasSelectedApiKey();
          if (!hasKey && typeof aistudio.openSelectKey === 'function') {
            await aistudio.openSelectKey();
          }
        } catch (err) {
          console.warn("AI Studio key check skipped:", err);
        }
      }
    };
    checkKey();
  }, []);

  const results = useMemo((): CalculationResult | null => {
    const { m, q, w, f300, f600, s_equiv, mm } = labData;
    const parse = (val: string | number) => {
        const cleaned = (val ?? '').toString().replace(',', '.').trim();
        return cleaned === '' ? NaN : parseFloat(cleaned);
    };
    
    const m_val = parse(m);
    const q_val = parse(q);
    const w_val = parse(w);
    const f3_val = parse(f300);
    const f6_val = parse(f600);
    const mm_val = parse(mm);
    
    const isValid = [m_val, q_val, w_val, f3_val, f6_val].every(v => !isNaN(v));
    if (!isValid) return null;

    const pv = f6_val - f3_val;
    const yp = f3_val - pv;
    const f = f6_val / 2;
    const poe = q_val / (1 - 0.01 * w_val);
    const ypPvRatio = pv !== 0 ? yp / pv : 0;
    const s_ratio = f6_val !== 0 ? yp / f6_val : 0;
    
    // Изотропия: (0,5 - (YP/f600 - 0,5)^2) * MM * MM * 0,01 * 0,01
    const mm_effective = isNaN(mm_val) ? 1.0 : mm_val;
    const isotropy = (f6_val !== 0) 
      ? (0.5 - Math.pow(((yp / f6_val) - 0.5), 2)) * mm_effective * mm_effective * 0.01 * 0.01 
      : 0;
      
    // Генерация: f600 * (1 - YP/f600) * MM / KOE
    const generation = (q_val !== 0 && f6_val !== 0) 
      ? f6_val * (1 - (yp / f6_val)) * mm_effective / q_val 
      : 0;
      
    // Загущение: f600 / (KOE * MM^2 * 0.01^2)
    const thickening = (q_val !== 0 && f6_val !== 0) 
      ? f6_val / (q_val * mm_effective * mm_effective * 0.01 * 0.01) 
      : 0;
      
    const logArg = f6_val !== 0 ? f6_val * 0.001 : 0.06;
    const logVal = logArg > 0 ? Math.log10(logArg) : 0;
    
    // Полнота: (KOE * 0.01 * MM * 0.01 * MM * LOG(f600*0.001)^2) / (YP/f600)
    const completeness = (s_ratio !== 0) 
      ? (q_val * 0.01 * mm_effective * 0.01 * mm_effective * Math.pow(logVal, 2)) / s_ratio 
      : 0;
    
    // Equivalent Calculation
    let equivalent: number | undefined = undefined;
    const s_equiv_val = parse(s_equiv);
    if (!isNaN(s_equiv_val) && !isNaN(mm_val)) {
      equivalent = (q_val / (1 - w_val * 0.01)) * 10 * 0.001 * mm_val * (1 - s_equiv_val * 0.01);
    }

    return { 
      m: m_val, q: q_val, w: w_val, f, pv, yp,
      poe, ypPvRatio, s: s_ratio, isotropy, generation, thickening, completeness,
      equivalent
    };
  }, [labData]);

  const getCriterionStyle = (type: string, value: number) => {
    switch (type) {
      case 'YP/PV':
        if (value > 6) return 'bg-rose-100 border-rose-200 text-rose-900';
        if (value > 3) return 'bg-amber-100 border-amber-200 text-amber-900';
        return 'bg-emerald-100 border-emerald-200 text-emerald-900';
      case 'Изотропия':
        if (value < 0.2) return 'bg-rose-100 border-rose-200 text-rose-900';
        if (value < 0.24) return 'bg-amber-100 border-amber-200 text-amber-900';
        return 'bg-emerald-100 border-emerald-200 text-emerald-900';
      case 'Генерация':
        if (value < 4) return 'bg-rose-100 border-rose-200 text-rose-900';
        if (value < 8.5) return 'bg-amber-100 border-amber-200 text-amber-900';
        return 'bg-emerald-100 border-emerald-200 text-emerald-900';
      case 'Загущение':
        if (value > 1.3) return 'bg-rose-100 border-rose-200 text-rose-900';
        if (value >= 1.0) return 'bg-amber-100 border-amber-200 text-amber-900';
        return 'bg-emerald-100 border-emerald-200 text-emerald-900';
      case 'Полнота':
        if (value < 100) return 'bg-rose-100 border-rose-200 text-rose-900';
        if (value <= 115) return 'bg-amber-100 border-amber-200 text-amber-900';
        return 'bg-emerald-100 border-emerald-200 text-emerald-900';
      default:
        return 'bg-slate-50 border-slate-100 text-slate-700';
    }
  };

  const calculateCompleteness = (m: number, q: number, f600: number, yp: number, manualS?: number) => {
    const s = manualS !== undefined ? manualS : (f600 !== 0 ? yp / f600 : 0.0001);
    const logArg = f600 * 0.001;
    const logVal = logArg > 0 ? Math.log10(logArg) : 0;
    const mm_val = parseFloat(labData.mm) || 1.0;
    let res = (s !== 0) ? (q * 0.01 * mm_val * 0.01 * mm_val * Math.pow(logVal, 2)) / s : 0;
    if (res < 30) res = 30;
    if (res > 200) res = 200;
    return res;
  };

  const generateChart = async () => {
    if (!results || !chartRef.current) return;
    await new Promise(r => setTimeout(r, 100));
    const xVar = CHART_VARIABLES[axisX];
    const yVar = CHART_VARIABLES[axisY];
    const xSteps = 20, ySteps = 20;
    const xData: number[] = [], yData: number[] = [], zData: number[][] = [];
    for (let i = 0; i <= xSteps; i++) xData.push(xVar.min + (xVar.max - xVar.min) * (i / xSteps));
    for (let j = 0; j <= ySteps; j++) yData.push(yVar.min + (yVar.max - yVar.min) * (j / ySteps));
    for (let j = 0; j <= ySteps; j++) {
      const row: number[] = [];
      for (let i = 0; i <= xSteps; i++) {
        let current_m = results.m, current_q = results.q, current_f600 = parseFloat(labData.f600.toString()) || 60, current_yp = results.yp, current_s: number | undefined = undefined;
        if (axisX === 'm') current_m = xData[i]; else if (axisX === 'q') current_q = xData[i]; else if (axisX === 'f600') current_f600 = xData[i]; else if (axisX === 'yp') current_yp = xData[i]; else if (axisX === 'yp_f600') current_s = xData[i];
        if (axisY === 'm') current_m = yData[j]; else if (axisY === 'q') current_q = yData[j]; else if (axisY === 'f600') current_f600 = yData[j]; else if (axisY === 'yp') current_yp = yData[j]; else if (axisY === 'yp_f600') current_s = yData[j];
        row.push(calculateCompleteness(current_m, current_q, current_f600, current_yp, current_s));
      }
      zData.push(row);
    }
    const data = [{ z: zData, x: xData, y: yData, type: 'surface', colorscale: 'Viridis', colorbar: { title: 'Полнота' } }];
    const layout = { title: `Полнота от ${xVar.label} и ${yVar.label}`, autosize: true, margin: { l: 0, r: 0, b: 0, t: 50 }, scene: { xaxis: { title: xVar.label }, yaxis: { title: yVar.label }, zaxis: { title: 'Полнота', range: [30, 200] } } };
    try {
      await Plotly.newPlot(chartRef.current, data, layout, { responsive: true, displayModeBar: false });
      const img = await Plotly.toImage(chartRef.current, { format: 'png', width: 1000, height: 800 });
      setCurrentChartImage(img);
    } catch (err) { console.error(err); }
  };

  useEffect(() => { if (showChartModal && results) generateChart(); }, [showChartModal, axisX, axisY, results]);

  const toggleChartInReport = () => {
    if (!currentChartImage) return;
    const chartId = `${axisX}-${axisY}`;
    const exists = savedCharts.find(c => c.id === chartId);
    if (exists) setSavedCharts(savedCharts.filter(c => c.id !== chartId));
    else setSavedCharts([...savedCharts, { id: chartId, axisX: CHART_VARIABLES[axisX].label, axisY: CHART_VARIABLES[axisY].label, imageData: currentChartImage }]);
  };

  const handleGetConclusions = async () => {
    if (!results) { setAnalysisError("Введите данные."); return; }
    setAnalysisError(null); setIsAnalyzing(true);
    const aistudio = (window as any).aistudio;
    try {
      if (!process.env.API_KEY && aistudio && typeof aistudio.openSelectKey === 'function') {
        await aistudio.openSelectKey();
      }
      const cons = await getBentoniteConclusions(results);
      setConclusions(cons);
    } catch (err: any) {
      console.error("Gemini Error:", err);
      const errorMessage = err.message || "";
      if ((errorMessage.includes("Requested entity was not found") || 
           errorMessage.includes("API key not valid") || 
           errorMessage.includes("403") || 
           errorMessage.includes("401")) && aistudio) {
        await aistudio.openSelectKey();
        setAnalysisError("Пожалуйста, выберите корректный API ключ.");
      } else {
        setAnalysisError("Ошибка связи с ИИ.");
      }
    } finally { setIsAnalyzing(false); }
  };

  const exportWord = async () => {
    if (!results) return;
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, VerticalAlign, ImageRun } = docx;
    const base64ToUint8Array = (base64: string) => {
      const bin = atob(base64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return bytes;
    };
    const createCell = (text: string, bold = false) => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold, size: 22 })], alignment: "center" })], verticalAlign: VerticalAlign.CENTER });
    const reportRows = [ 
      ["Содержание смектита (m)", `${results.m}%`], 
      ["Обменная емкость (q)", String(results.q)], 
      ["Влажность (w)", `${results.w}%`], 
      ["PV (Пл. вязкость)", results.pv.toFixed(2)], 
      ["YP (Пред. текучести)", results.yp.toFixed(2)], 
      ["YP/PV", results.ypPvRatio.toFixed(2)], 
      ["Степень замещения (YP/f600)", results.s.toFixed(3)], 
      ["Критерий полноты", results.completeness.toFixed(2)] 
    ];
    if (results.equivalent !== undefined) {
      reportRows.push(["Эквивалент ПАВ, г/кг", results.equivalent.toFixed(3)]);
    }

    const children: any[] = [
      new Paragraph({ alignment: "center", spacing: { after: 400 }, children: [new TextRun({ text: "ОТЧЕТ ОБ ИСПЫТАНИИ БЕНТОНИТА", bold: true, size: 32, color: "4f46e5" })] }),
      new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ children: [createCell("Параметр", true), createCell("Значение", true)] }), ...reportRows.map(row => new TableRow({ children: [createCell(row[0], false), createCell(row[1], false)] })) ] }),
      new Paragraph({ text: "", spacing: { before: 400 } }),
      new Paragraph({ children: [new TextRun({ text: "ЭКСПЕРТНЫЕ ВЫВОДЫ:", bold: true, size: 28, color: "10b981" })] }),
      ...conclusions.map((c, i) => new Paragraph({ spacing: { before: 120 }, children: [new TextRun({ text: `${i + 1}. `, bold: true }), new TextRun({ text: c.text })] }))
    ];
    if (savedCharts.length > 0) {
      children.push(new Paragraph({ text: "", spacing: { before: 400 } }));
      savedCharts.forEach(chart => {
        children.push(new Paragraph({ alignment: "center", children: [new ImageRun({ data: base64ToUint8Array(chart.imageData.split(',')[1]), transformation: { width: 500, height: 375 } } as any)] }));
      });
    }
    const blob = await Packer.toBlob(new Document({ sections: [{ children }] }));
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'geolab_report.docx'; a.click();
  };

  const createPPTX = async (style: PPTStyle) => {
    if (!results) return; setIsGeneratingDoc(true); setShowStylePicker(false);
    try {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      pptx.defineSlideMaster({ title: 'MASTER', background: { color: style.bg }, objects: [{ rect: { x: 0, y: 0, w: '100%', h: 0.8, fill: { color: style.primary } } }, { text: { text: 'GeoLab Pro Report', options: { x: 0.5, y: 0.2, color: 'ffffff', fontSize: 24, bold: true } } }] });
      const s1 = pptx.addSlide({ masterName: 'MASTER' });
      s1.addText('Технический отчет по качеству бентонита', { x: 1, y: 2, w: 11, fontSize: 36, bold: true, color: style.primary });
      const s2 = pptx.addSlide({ masterName: 'MASTER' });
      const tableData = [[{ text: 'Параметр' }, { text: 'Значение' }], [{ text: 'Смектит' }, { text: `${results.m}%` }], [{ text: 'КОЕ' }, { text: String(results.q) }], [{ text: 'PV' }, { text: results.pv.toFixed(2) }], [{ text: 'YP' }, { text: results.yp.toFixed(2) }], [{ text: 'Замещение' }, { text: results.s.toFixed(3) }], [{ text: 'Полнота' }, { text: results.completeness.toFixed(2) }]];
      if (results.equivalent !== undefined) tableData.push([{ text: 'Эквивалент ПАВ' }, { text: results.equivalent.toFixed(3) }]);
      s2.addTable(tableData, { x: 0.5, y: 1.5, w: 12, border: { pt: 1, color: style.secondary }, fill: { color: 'ffffff' } });
      savedCharts.forEach(chart => { const sc = pptx.addSlide({ masterName: 'MASTER' }); sc.addImage({ data: chart.imageData, x: 1, y: 1, w: 11, h: 5.5 }); });
      await pptx.writeFile({ fileName: 'Geolab_Analysis.pptx' });
    } finally { setIsGeneratingDoc(false); }
  };

  const [isLiveActive, setIsLiveActive] = useState(false);
  const sessionRef = useRef<any>(null);
  const sourcesRef = useRef<Set<AudioBufferSourceNode>>(new Set());
  const nextStartTimeRef = useRef<number>(0);

  const startLiveSession = async () => {
    if (isLiveActive) { sessionRef.current?.close(); setIsLiveActive(false); return; }
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const ctxIn = new (window.AudioContext || (window as any).webkitAudioContext)({ sampleRate: 16000 });
      const ctxOut = new (window.AudioContext || (window as any).webkitAudioContext)({ sampleRate: 24000 });
      
      const systemInstruction = `Вы — эксперт GeoLab Pro по реологии бентонита. Помогайте анализировать результаты, опираясь на научные данные:
1. Реология: Пластическая вязкость (PV) — скорость разрушения структуры. Чем крупнее слоистый пакет, тем выше PV и ниже YP.
2. YP/PV (хрупкость): > 6 — бентонит активирован содой, для органомодификации сомнителен. 1.5 - 3 — класс 'Drilling grade'.
3. Изотропия: Приемлемо >= 0.24. Генерация: Приемлемо > 8.5. Загущение: Норма 1.0 - 1.3. Полнота: Приемлемо 100-115, высокое качество > 115.
4. PV ~ КОЕ(Ca)/КОЕ(Na). YP ~ ММ * (КОЕ/КОЕo).
Текущие данные пользователя: ${results ? `PV:${results.pv.toFixed(2)}, YP:${results.yp.toFixed(2)}, Изотропия:${results.isotropy.toFixed(4)}, Полнота:${results.completeness.toFixed(2)}` : 'Данные не введены.'}`;

      const sessionPromise = ai.live.connect({
        model: 'gemini-2.5-flash-native-audio-preview-09-2025',
        config: { 
          responseModalities: [Modality.AUDIO], 
          inputAudioTranscription: {}, 
          systemInstruction: systemInstruction 
        },
        callbacks: {
          onopen: () => {
            setIsLiveActive(true);
            const source = ctxIn.createMediaStreamSource(stream);
            const proc = ctxIn.createScriptProcessor(4096, 1, 1);
            proc.onaudioprocess = (e) => { 
              sessionPromise.then(s => s.sendRealtimeInput({ media: createBlob(e.inputBuffer.getChannelData(0)) })); 
            };
            source.connect(proc); proc.connect(ctxIn.destination);
          },
          onmessage: async (msg: LiveServerMessage) => {
            const audioData = msg.serverContent?.modelTurn?.parts?.[0]?.inlineData?.data;
            if (audioData) {
              nextStartTimeRef.current = Math.max(nextStartTimeRef.current, ctxOut.currentTime);
              const buffer = await decodeAudioData(decode(audioData), ctxOut, 24000, 1);
              const source = ctxOut.createBufferSource(); 
              source.buffer = buffer; source.connect(ctxOut.destination);
              source.addEventListener('ended', () => sourcesRef.current.delete(source));
              source.start(nextStartTimeRef.current); 
              nextStartTimeRef.current += buffer.duration;
              sourcesRef.current.add(source);
            }
            if (msg.serverContent?.interrupted) {
              for (const source of sourcesRef.current.values()) { try { source.stop(); } catch (e) {} }
              sourcesRef.current.clear(); nextStartTimeRef.current = 0;
            }
            if (msg.serverContent?.inputTranscription) {
              const val = parseSpokenNumber(msg.serverContent.inputTranscription.text || '');
              if (val !== null) setLabData(prev => ({ ...prev, [LAB_ORDER[currentStep]]: val.toString() }));
            }
          },
          onclose: () => { setIsLiveActive(false); nextStartTimeRef.current = 0; },
          onerror: () => { setIsLiveActive(false); nextStartTimeRef.current = 0; }
        }
      });
      sessionRef.current = await sessionPromise;
    } catch (e) { setIsLiveActive(false); }
  };

  const gridResults = results ? [
    { label: 'Изотропия', value: results.isotropy, type: 'Изотропия', formatted: results.isotropy.toFixed(4) },
    { label: 'Генерация', value: results.generation, type: 'Генерация', formatted: results.generation.toFixed(2) },
    { label: 'Загущение', value: results.thickening, type: 'Загущение', formatted: results.thickening.toFixed(2) },
    { label: 'YP/PV', value: results.ypPvRatio, type: 'YP/PV', formatted: results.ypPvRatio.toFixed(2) },
    { label: 'Степень замещения', value: results.s, type: 'Замещение', formatted: results.s.toFixed(3) }
  ] : [];

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-10 text-slate-900 flex flex-col items-center">
      <div className="w-full max-w-4xl space-y-6">
        <header className="flex justify-between items-center bg-white p-6 rounded-[2rem] shadow-sm border border-slate-100">
          <div className="flex items-center gap-3">
            <div className={`w-3 h-3 rounded-full ${isLiveActive ? 'bg-red-500 animate-pulse' : 'bg-indigo-600'}`}></div>
            <div className="flex items-baseline gap-2">
              <h1 className="text-2xl font-black tracking-tighter">GeoLab<span className="text-indigo-600">Pro</span></h1>
              <span className="text-red-600 text-[10px] font-black uppercase tracking-widest hidden sm:inline">Bentonite Co</span>
            </div>
          </div>
          <div className="flex gap-2">
            <button onClick={startLiveSession} className={`px-5 py-2.5 rounded-2xl font-black uppercase text-[10px] tracking-widest shadow-md transition-all ${isLiveActive ? 'bg-red-500 text-white' : 'bg-indigo-600 text-white'}`}>
              {isLiveActive ? 'Стоп' : 'Голос'}
            </button>
            <label className="cursor-pointer bg-slate-100 hover:bg-slate-200 text-slate-600 px-5 py-2.5 rounded-2xl font-black uppercase text-[10px] tracking-widest shadow-sm">
              {isUploading ? '...' : 'Фото'}
              <input type="file" className="hidden" accept="image/*" onChange={async (e) => {
                const file = e.target.files?.[0]; if (!file) return;
                setIsUploading(true);
                const reader = new FileReader();
                reader.onloadend = async () => {
                  try {
                    const ext = await extractLabDataFromImage((reader.result as string).split(',')[1]);
                    setLabData(prev => ({ ...prev, ...ext as any }));
                  } finally { setIsUploading(false); }
                };
                reader.readAsDataURL(file);
              }} />
            </label>
          </div>
        </header>

        <div className="flex flex-col gap-4">
           <div className="grid grid-cols-5 gap-3">
            {LAB_ORDER.slice(0, 5).map((key, index) => (
              <button key={key} onClick={() => { setCurrentStep(index); setShowResults(false); }}
                className={`p-4 rounded-[1.5rem] border-2 text-center transition-all ${currentStep === index ? 'bg-indigo-600 border-indigo-600 text-white shadow-xl scale-105' : labData[key] ? 'bg-white border-emerald-100 text-slate-800' : 'bg-white border-slate-200 opacity-60'}`}>
                <div className={`text-[8px] font-black uppercase mb-1 truncate ${currentStep === index ? 'text-indigo-100' : 'text-slate-400'}`}>{LAB_LABELS[key]}</div>
                <div className="text-xl font-mono font-black">{labData[key] || '0'}</div>
              </button>
            ))}
          </div>
          <div className="space-y-2">
            <h4 className="text-[10px] font-black uppercase tracking-widest text-slate-400 ml-2">Расчет эквивалента</h4>
            <div className="grid grid-cols-2 gap-3">
              {LAB_ORDER.slice(5).map((key, index) => {
                const stepIdx = index + 5;
                return (
                  <button key={key} onClick={() => { setCurrentStep(stepIdx); setShowResults(false); }}
                    className={`p-4 rounded-[1.5rem] border-2 text-center transition-all ${currentStep === stepIdx ? 'bg-indigo-600 border-indigo-600 text-white shadow-xl scale-105' : labData[key] ? 'bg-white border-emerald-100 text-slate-800' : 'bg-white border-slate-200 opacity-60'}`}>
                    <div className={`text-[8px] font-black uppercase mb-1 truncate ${currentStep === stepIdx ? 'text-indigo-100' : 'text-slate-400'}`}>{LAB_LABELS[key]}</div>
                    <div className="text-xl font-mono font-black">{labData[key] || '0'}</div>
                  </button>
                );
              })}
            </div>
          </div>
        </div>

        {!showResults ? (
          <div className="bg-white p-12 rounded-[2.5rem] shadow-2xl border border-slate-100 text-center relative overflow-hidden animate-slide-up">
            <h2 className="text-[10px] font-black text-slate-400 uppercase tracking-[0.3em] mb-10">
               {currentStep >= 5 ? 'Расчет эквивалента: ' : ''}{LAB_LABELS[LAB_ORDER[currentStep]]}
            </h2>
            <input type="text" inputMode="decimal" className="w-full text-center text-7xl font-mono font-black py-6 bg-slate-50 border-b-8 border-slate-100 focus:border-indigo-600 outline-none rounded-3xl transition-all"
              value={labData[LAB_ORDER[currentStep]]} onChange={(e) => setLabData(p => ({ ...p, [LAB_ORDER[currentStep]]: e.target.value }))}
              onKeyDown={(e) => e.key === 'Enter' && (currentStep < 6 ? setCurrentStep(currentStep + 1) : setShowResults(true))} autoFocus />
            <button onClick={() => (currentStep < 6 ? setCurrentStep(currentStep + 1) : setShowResults(true))} className="mt-12 bg-indigo-600 text-white px-16 py-5 rounded-3xl font-black uppercase text-sm tracking-widest shadow-2xl hover:bg-indigo-700 active:scale-95 transition-all">
              {currentStep === 6 ? 'Рассчитать' : 'Продолжить'}
            </button>
          </div>
        ) : (
          <div className="space-y-6 animate-slide-up">
            <div className="bg-white rounded-[2.5rem] overflow-hidden shadow-2xl border border-slate-100">
              <div className="bg-indigo-600 p-8 text-white flex justify-between items-center">
                <h3 className="text-lg font-black uppercase tracking-widest">Протокол</h3>
                <div className="flex gap-2">
                  <button onClick={() => setShowChartModal(true)} className="bg-emerald-500 hover:bg-emerald-600 text-white px-4 py-2.5 rounded-2xl text-[10px] font-black uppercase shadow-lg">Графики</button>
                  <button onClick={exportWord} className="bg-white/20 px-4 py-2.5 rounded-2xl text-[10px] font-black uppercase">Word</button>
                  <button onClick={() => setShowStylePicker(true)} className="bg-white text-indigo-700 px-4 py-2.5 rounded-2xl text-[10px] font-black uppercase shadow-lg">PPTX</button>
                </div>
              </div>
              <div className="p-8 grid grid-cols-1 md:grid-cols-2 gap-10">
                <div className="space-y-3">
                  {LAB_ORDER.map(k => (
                    <div key={k} className="flex justify-between border-b border-slate-50 pb-2 text-sm">
                      <span className="text-slate-500 font-bold">{LAB_LABELS[k]}</span>
                      <span className="font-mono font-black">{labData[k] || '-'}</span>
                    </div>
                  ))}
                  <div className="grid grid-cols-1 gap-4 mt-6">
                    {results?.equivalent !== undefined && (
                      <div className="p-5 rounded-2xl bg-sky-100 border-2 border-sky-200 text-sky-900 flex justify-between items-center shadow-sm">
                        <span className="text-[10px] font-black uppercase">Эквивалент ПАВ, г/кг</span>
                        <span className="text-2xl font-mono font-black">{results.equivalent.toFixed(3)}</span>
                      </div>
                    )}
                    {results && (
                      <div className={`p-5 rounded-2xl flex justify-between items-center border-2 transition-colors ${getCriterionStyle('Полнота', results.completeness)}`}>
                        <span className="text-[10px] font-black uppercase">Критерий полноты</span>
                        <span className="text-2xl font-mono font-black">{results.completeness.toFixed(2)}</span>
                      </div>
                    )}
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                   {gridResults.map(c => (
                     <div key={c.label} className={`p-5 rounded-3xl border-2 flex flex-col items-center justify-center transition-colors ${getCriterionStyle(c.type, c.value)}`}>
                       <span className="text-[9px] font-black uppercase mb-2 text-center leading-tight opacity-70">{c.label}</span>
                       <span className="text-xl font-mono font-black">{c.formatted}</span>
                     </div>
                   ))}
                </div>
              </div>
            </div>

            <div className="bg-white rounded-[2.5rem] p-8 shadow-2xl border border-slate-100">
              <div className="flex justify-between items-center mb-8">
                <h4 className="text-xl font-black text-slate-900 uppercase tracking-widest">Выводы ИИ</h4>
                <button onClick={handleGetConclusions} disabled={isAnalyzing} className="text-[10px] bg-emerald-600 text-white px-6 py-3 rounded-2xl font-black uppercase shadow-lg hover:bg-emerald-700">
                   {isAnalyzing ? 'Изучаю...' : 'Обновить'}
                </button>
              </div>
              {analysisError && <div className="p-4 bg-red-50 text-red-700 text-xs rounded-xl mb-4 font-bold">{analysisError}</div>}
              <div className="space-y-4">
                {conclusions.map((c, i) => {
                  const isNeg = c.sentiment === 'negative';
                  return (
                    <div key={i} className={`flex gap-5 p-6 rounded-3xl border-2 transition-all shadow-sm ${isNeg ? 'bg-amber-100 border-amber-300' : 'bg-emerald-100 border-emerald-300'}`}>
                      <span className={`flex-shrink-0 w-10 h-10 rounded-2xl text-white flex items-center justify-center font-black ${isNeg ? 'bg-amber-600' : 'bg-emerald-600'}`}>{i+1}</span>
                      <p className={`text-sm font-bold leading-relaxed ${isNeg ? 'text-amber-900' : 'text-emerald-900'}`}>{c.text}</p>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>

      {showChartModal && (
        <div className="fixed inset-0 bg-slate-900/90 backdrop-blur-xl flex items-center justify-center z-[70] p-4">
          <div className="bg-white w-full max-w-5xl rounded-[3rem] overflow-hidden shadow-2xl flex flex-col h-[85vh]">
            <div className="flex-grow p-4 min-h-0"><div ref={chartRef} className="w-full h-full"></div></div>
            <div className="p-8 bg-slate-50 border-t flex flex-wrap gap-4 justify-between items-center">
               <div className="flex gap-3">
                  <select value={axisX} onChange={(e) => setAxisX(e.target.value as any)} className="bg-white border-2 border-slate-200 px-4 py-2 rounded-xl text-xs font-black">
                    {Object.entries(CHART_VARIABLES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                  </select>
                  <select value={axisY} onChange={(e) => setAxisY(e.target.value as any)} className="bg-white border-2 border-slate-200 px-4 py-2 rounded-xl text-xs font-black">
                    {Object.entries(CHART_VARIABLES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                  </select>
               </div>
               <div className="flex gap-3">
                 <button onClick={toggleChartInReport} className={`px-6 py-3 rounded-xl font-black uppercase text-[10px] tracking-widest ${savedCharts.some(c => c.id === `${axisX}-${axisY}`) ? 'bg-rose-500 text-white' : 'bg-emerald-600 text-white'}`}>
                   {savedCharts.some(c => c.id === `${axisX}-${axisY}`) ? "Убрать" : "В отчет"}
                 </button>
                 <button onClick={() => setShowChartModal(false)} className="px-6 py-3 bg-slate-900 text-white rounded-xl font-black uppercase text-[10px]">Закрыть</button>
               </div>
            </div>
          </div>
        </div>
      )}

      {showStylePicker && (
        <div className="fixed inset-0 bg-slate-900/70 flex items-center justify-center z-[80] p-4">
          <div className="bg-white w-full max-sm:w-full max-w-sm rounded-[2.5rem] p-8 shadow-2xl">
            <h5 className="text-lg font-black uppercase mb-6 text-center">Стиль презентации</h5>
            <div className="grid gap-3">
              {PPT_STYLES.map(s => (
                <button key={s.id} onClick={() => createPPTX(s)} className="p-4 rounded-2xl border-2 border-slate-100 hover:border-indigo-600 flex items-center gap-4 transition-all">
                  <div className="w-8 h-8 rounded-lg" style={{ background: s.primary }}></div>
                  <span className="font-black text-sm text-slate-700">{s.name}</span>
                </button>
              ))}
            </div>
            <button onClick={() => setShowStylePicker(false)} className="w-full mt-6 py-2 font-black text-slate-400 uppercase text-[10px]">Отмена</button>
          </div>
        </div>
      )}

      {isGeneratingDoc && (
        <div className="fixed inset-0 bg-indigo-600/95 flex flex-col items-center justify-center z-[100] text-white">
          <div className="w-12 h-12 border-4 border-white/20 border-t-white rounded-full animate-spin mb-4"></div>
          <p className="text-xs font-black uppercase tracking-widest animate-pulse">Генерация документа...</p>
        </div>
      )}
    </div>
  );
};

export default App;
