
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
    m: '', q: '', w: '', f300: '', f600: '',
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

  const results = useMemo((): CalculationResult | null => {
    const { m, q, w, f300, f600 } = labData;
    const parse = (val: string) => {
        const cleaned = val.replace(',', '.').trim();
        return cleaned === '' ? NaN : parseFloat(cleaned);
    };
    
    const m_val = parse(m);
    const q_val = parse(q);
    const w_val = parse(w);
    const f3_val = parse(f300);
    const f6_val = parse(f600);
    
    const isValid = [m_val, q_val, w_val, f3_val, f6_val].every(v => !isNaN(v));
    if (!isValid) return null;

    const pv = f6_val - f3_val;
    const yp = f3_val - pv;
    const f = f6_val / 2;
    const poe = q_val / (1 - 0.01 * w_val);
    const ypPvRatio = pv !== 0 ? yp / pv : 0;
    const s = f6_val !== 0 ? yp / f6_val : 0;
    
    const isotropy = (f6_val !== 0) 
      ? (0.5 - Math.pow(((yp / f6_val) - 0.5), 2)) * m_val * m_val * 0.01 * 0.01 
      : 0;
    const generation = (q_val !== 0 && f6_val !== 0) 
      ? f6_val * (1 - (yp / f6_val)) * m_val / q_val 
      : 0;
    const thickening = (m_val !== 0 && q_val !== 0 && f6_val !== 0) ? f6_val / (0.01 * 0.01 * m_val * m_val * q_val) : 0;
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

  const calculateCompleteness = (m: number, q: number, f600: number, yp: number, manualS?: number) => {
    const s = manualS !== undefined ? manualS : (f600 !== 0 ? yp / f600 : 0.0001);
    const logArg = f600 * 0.001;
    const logVal = logArg > 0 ? Math.log10(logArg) : 0;
    let res = (s !== 0) ? (q * 0.01 * m * 0.01 * m * Math.pow(logVal, 2)) / s : 0;
    if (res < 30) res = 30;
    if (res > 200) res = 200;
    return res;
  };

  const generateChart = async () => {
    if (!results || !chartRef.current) return;
    await new Promise(r => setTimeout(r, 100));

    const xVar = CHART_VARIABLES[axisX];
    const yVar = CHART_VARIABLES[axisY];
    
    const xSteps = 20;
    const ySteps = 20;
    const xData: number[] = [];
    const yData: number[] = [];
    const zData: number[][] = [];

    for (let i = 0; i <= xSteps; i++) xData.push(xVar.min + (xVar.max - xVar.min) * (i / xSteps));
    for (let j = 0; j <= ySteps; j++) yData.push(yVar.min + (yVar.max - yVar.min) * (j / ySteps));

    for (let j = 0; j <= ySteps; j++) {
      const row: number[] = [];
      for (let i = 0; i <= xSteps; i++) {
        let current_m = results.m;
        let current_q = results.q;
        let current_f600 = parseFloat(labData.f600.replace(',', '.')) || 60;
        let current_yp = results.yp;
        let current_s: number | undefined = undefined;

        const updateParams = (key: keyof typeof CHART_VARIABLES, val: number) => {
          if (key === 'm') current_m = val;
          else if (key === 'q') current_q = val;
          else if (key === 'f600') current_f600 = val;
          else if (key === 'yp') current_yp = val;
          else if (key === 'yp_f600') current_s = val;
        };

        updateParams(axisX, xData[i]);
        updateParams(axisY, yData[j]);
        row.push(calculateCompleteness(current_m, current_q, current_f600, current_yp, current_s));
      }
      zData.push(row);
    }

    const data = [{
      z: zData,
      x: xData,
      y: yData,
      type: 'surface',
      colorscale: 'Viridis',
      colorbar: { title: 'Полнота', thickness: 15 }
    }];

    const layout = {
      title: `Зависимость полноты от ${xVar.label} и ${yVar.label}`,
      autosize: true,
      margin: { l: 0, r: 0, b: 0, t: 50 },
      scene: {
        xaxis: { title: xVar.label },
        yaxis: { title: yVar.label },
        zaxis: { title: 'Полнота', range: [30, 200] }
      },
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
    };

    try {
      await Plotly.newPlot(chartRef.current, data, layout, { responsive: true, displayModeBar: false });
      const img = await Plotly.toImage(chartRef.current, { format: 'png', width: 1000, height: 800 });
      setCurrentChartImage(img);
    } catch (err) {
      console.error("Plotly error:", err);
    }
  };

  useEffect(() => {
    if (showChartModal && results) {
      generateChart();
    }
  }, [showChartModal, axisX, axisY, results]);

  const toggleChartInReport = () => {
    if (!currentChartImage) return;
    const chartId = `${axisX}-${axisY}`;
    const exists = savedCharts.find(c => c.id === chartId);
    if (exists) {
      setSavedCharts(savedCharts.filter(c => c.id !== chartId));
    } else {
      setSavedCharts([...savedCharts, {
        id: chartId,
        axisX: CHART_VARIABLES[axisX].label,
        axisY: CHART_VARIABLES[axisY].label,
        imageData: currentChartImage
      }]);
    }
  };

  const isCurrentChartInReport = savedCharts.some(c => c.id === `${axisX}-${axisY}`);

  const handleGetConclusions = async () => {
    if (!results) {
        setAnalysisError("Пожалуйста, заполните все лабораторные данные для проведения анализа.");
        return;
    }
    setAnalysisError(null);
    setIsAnalyzing(true);
    try {
      const cons = await getBentoniteConclusions(results);
      setConclusions(cons);
    } catch (err) { 
        console.error("Analysis failed:", err); 
        setAnalysisError("Не удалось получить экспертный анализ. Проверьте подключение к интернету или API ключ.");
    } finally { 
        setIsAnalyzing(false); 
    }
  };

  const exportWord = async () => {
    if (!results) return;
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, VerticalAlign, ImageRun } = docx;

    const base64ToUint8Array = (base64: string) => {
      const binaryString = window.atob(base64);
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) { bytes[i] = binaryString.charCodeAt(i); }
      return bytes;
    };

    const createCell = (text: string, bold = false, align: any = "center") => new TableCell({
      children: [new Paragraph({ 
        children: [new TextRun({ text: String(text), bold, size: 22 })],
        alignment: align
      })],
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 100, bottom: 100, left: 100, right: 100 }
    });

    const reportRows = [
      ["Содержание смектита (m)", `${results.m}%`],
      ["Обменная емкость (q)", String(results.q)],
      ["Влажность (w)", `${results.w}%`],
      ["PV (Пластическая вязкость)", results.pv.toFixed(2)],
      ["YP (Предел текучести)", results.yp.toFixed(2)],
      ["Эфф. вязкость (f600/2)", results.f.toFixed(2)],
      ["YP/PV", results.ypPvRatio.toFixed(2)],
      ["ПОЕ (полная емкость)", results.poe.toFixed(2)],
      ["Критерий изотропии", results.isotropy.toFixed(4)],
      ["Критерий генерации", results.generation.toFixed(2)],
      ["Критерий загущения", results.thickening.toFixed(2)],
      ["Критерий полноты", results.completeness.toFixed(2)],
    ];

    const children: any[] = [
      new Paragraph({
        alignment: "center",
        spacing: { after: 400 },
        children: [new TextRun({ text: "ОТЧЕТ ОБ ИСПЫТАНИИ БЕНТОНИТА", bold: true, size: 32, color: "4f46e5" })],
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ children: [createCell("Параметр", true), createCell("Значение", true)] }),
          ...reportRows.map(row => new TableRow({ 
            children: [ createCell(row[0], false, "left"), createCell(row[1], false, "center") ] 
          }))
        ],
      }),
      new Paragraph({ text: "", spacing: { before: 400 } }),
      new Paragraph({ children: [new TextRun({ text: "ЭКСПЕРТНЫЕ ВЫВОДЫ:", bold: true, size: 28, color: "10b981" })], spacing: { after: 200 } }),
      ...conclusions.map((c, i) => {
        const sRaw = c.sentiment?.toLowerCase().trim() || 'neutral';
        const isNeg = sRaw.includes('neg') || sRaw.includes('нег') || sRaw.includes('плох') || sRaw.includes('bad');
        return new Paragraph({ 
          spacing: { before: 120 }, 
          children: [
            new TextRun({ text: `${i + 1}. `, bold: true, color: isNeg ? "b45309" : "10b981" }), 
            new TextRun({ text: c.text })
          ] 
        });
      }),
    ];

    if (savedCharts.length > 0) {
      children.push(new Paragraph({ text: "", spacing: { before: 400 } }));
      children.push(new Paragraph({ alignment: "center", children: [new TextRun({ text: "ГРАФИЧЕСКИЙ АНАЛИЗ ПОЛНОТЫ", bold: true, size: 28, color: "4f46e5" })] }));
      savedCharts.forEach(chart => {
        const base64Data = chart.imageData.split(',')[1];
        const bytes = base64ToUint8Array(base64Data);
        children.push(new Paragraph({ spacing: { before: 200 }, alignment: "center", children: [new TextRun({ text: `Зависимость от ${chart.axisX} и ${chart.axisY}`, bold: true, size: 20 })] }));
        children.push(new Paragraph({
          alignment: "center",
          children: [new ImageRun({ data: bytes, transformation: { width: 500, height: 375 } })]
        }));
      });
    }

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'geolab_report.docx';
    a.click();
    URL.revokeObjectURL(url);
  };

  const createPPTX = async (style: PPTStyle) => {
    if (!results) return;
    setIsGeneratingDoc(true);
    setShowStylePicker(false);
    
    try {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      pptx.defineSlideMaster({
        title: 'MASTER_SLIDE',
        background: { color: style.bg },
        objects: [
          { rect: { x: 0, y: 0, w: '100%', h: 0.8, fill: { color: style.primary } } },
          { text: { text: 'GeoLab Pro: Технический отчет', options: { x: 0.5, y: 0.2, color: 'ffffff', fontSize: 24, bold: true } } }
        ]
      });

      const s1 = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
      s1.addText('Анализ качества и пригодности\nбентонита', { x: 1, y: 1.5, w: 8, fontSize: 36, bold: true, color: style.primary });

      const s2 = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
      s2.addText('Результаты измерений', { x: 0.5, y: 1.2, fontSize: 28, bold: true, color: style.primary });
      const rows = [['Параметр', 'Значение'], ['Смектит (m)', `${results.m}%`], ['КОЕ (q)', String(results.q)], ['PV', results.pv.toFixed(2)], ['YP', results.yp.toFixed(2)], ['YP/PV', results.ypPvRatio.toFixed(2)]];
      s2.addTable(rows as any, { x: 0.5, y: 2, w: 10, fill: { color: 'FFFFFF' }, border: { pt: 1, color: style.secondary } });

      savedCharts.forEach(chart => {
        const sChart = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
        sChart.addText(`3D Анализ: ${chart.axisX} vs ${chart.axisY}`, { x: 0.5, y: 1.2, fontSize: 28, bold: true, color: style.primary });
        sChart.addImage({ data: chart.imageData, x: 1, y: 1.8, w: 11, h: 5.2 });
      });

      const s3 = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
      s3.addText('Критерии качества', { x: 0.5, y: 1.2, fontSize: 28, bold: true, color: style.primary });
      const critRows = [['Критерий', 'Значение', 'Оценка'], ['Изотропия', results.isotropy.toFixed(4), results.isotropy >= 0.24 ? 'Ок' : 'Низкая'], ['Генерация', results.generation.toFixed(2), results.generation > 8.5 ? 'Высокая' : 'Низкая'], ['Загущение', results.thickening.toFixed(2), results.thickening < 1.3 ? 'Оптимально' : 'Высокое'], ['Полнота', results.completeness.toFixed(2), results.completeness > 115 ? 'Идеально' : 'Средне']];
      s3.addTable(critRows as any, { x: 0.5, y: 2, w: 12, fill: { color: 'FFFFFF' }, border: { pt: 1, color: style.secondary } });

      const s4 = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
      s4.addText('Экспертные выводы', { x: 0.5, y: 1.2, fontSize: 28, bold: true, color: style.primary });
      conclusions.forEach((c, i) => {
        const sVal = c.sentiment?.toLowerCase().trim() || 'neutral';
        const isN = sVal.includes('neg') || sVal.includes('нег') || sVal.includes('bad');
        const color = isN ? 'b45309' : style.primary;
        s4.addText(`${i+1}. ${c.text}`, { x: 0.5, y: 2 + i * 0.7, w: 9, fontSize: 14, color });
      });

      await pptx.writeFile({ fileName: 'Geolab_Analysis.pptx' });
    } catch (err) { console.error(err); }
    finally { setIsGeneratingDoc(false); }
  };

  const startLiveSession = async () => {
    if (isLiveActive) { sessionRef.current?.close(); setIsLiveActive(false); return; }
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
          systemInstruction: 'Вы эксперт GeoLab. Помогайте вводить данные и анализировать бентонит.',
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
            const parts = msg.serverContent?.modelTurn?.parts;
            const audioData = parts && parts[0]?.inlineData?.data;
            if (audioData && audioContextOutRef.current) {
              const ctx = audioContextOutRef.current;
              nextStartTimeRef.current = Math.max(nextStartTimeRef.current, ctx.currentTime);
              const buffer = await decodeAudioData(decode(audioData), ctx, 24000, 1);
              const source = ctx.createBufferSource();
              source.buffer = buffer;
              source.connect(ctx.destination);
              source.addEventListener('ended', () => { sourcesRef.current.delete(source); });
              source.start(nextStartTimeRef.current);
              nextStartTimeRef.current += buffer.duration;
              sourcesRef.current.add(source);
            }
            if (msg.serverContent?.interrupted) {
              for (const source of sourcesRef.current.values()) { source.stop(); sourcesRef.current.delete(source); }
              nextStartTimeRef.current = 0;
            }
            if (msg.serverContent?.inputTranscription) {
              const text = msg.serverContent.inputTranscription.text;
              setLiveTranscription(text ?? '');
              const parsed = parseSpokenNumber(text ?? '');
              if (parsed !== null) setLabData(prev => ({ ...prev, [LAB_ORDER[currentStep]]: parsed.toString() }));
            }
          },
          onclose: () => setIsLiveActive(false),
          onerror: () => setIsLiveActive(false)
        }
      });
      sessionRef.current = await sessionPromise;
    } catch (e) { setIsLiveActive(false); }
  };

  const [isLiveActive, setIsLiveActive] = useState(false);
  const [liveTranscription, setLiveTranscription] = useState('');
  const sessionRef = useRef<any>(null);
  const audioContextInRef = useRef<AudioContext | null>(null);
  const audioContextOutRef = useRef<AudioContext | null>(null);
  const nextStartTimeRef = useRef(0);
  const sourcesRef = useRef<Set<AudioBufferSourceNode>>(new Set());

  return (
    <div className="min-h-screen bg-slate-50 p-3 md:p-5 text-slate-900 font-sans text-left">
      <div className="max-w-4xl mx-auto space-y-4">
        <header className="flex items-center justify-between bg-white px-5 py-4 rounded-2xl shadow-sm border border-slate-100">
          <div className="flex items-center gap-3">
            <div className={`w-3 h-3 rounded-full ${isLiveActive ? 'bg-red-500 animate-pulse' : 'bg-indigo-600'}`}></div>
            <h1 className="text-xl font-black tracking-tighter">GeoLab<span className="text-indigo-600">Pro</span></h1>
          </div>
          <div className="flex gap-2">
            <button onClick={startLiveSession} className={`px-4 py-2 rounded-xl font-bold text-xs uppercase shadow-md ${isLiveActive ? 'bg-red-500 text-white' : 'bg-indigo-600 text-white'}`}>
              {isLiveActive ? 'Стоп' : 'Голос'}
            </button>
            <label className="cursor-pointer bg-slate-100 hover:bg-slate-200 text-slate-600 px-4 py-2 rounded-xl font-bold text-xs uppercase shadow-sm">
              {isUploading ? '...' : 'Фото'}
              <input type="file" className="hidden" accept="image/*" onChange={async (e) => {
                const file = e.target.files?.[0]; if (!file) return;
                setIsUploading(true);
                const reader = new FileReader();
                reader.onloadend = async () => {
                  try {
                    const base64 = (reader.result as string).split(',')[1];
                    const ext = await extractLabDataFromImage(base64);
                    setLabData(prev => ({ ...prev, ...ext }));
                  } catch (err) { console.error(err); } 
                  finally { setIsUploading(false); }
                };
                reader.readAsDataURL(file);
              }} />
            </label>
          </div>
        </header>

        {isLiveActive && (
          <div className="bg-slate-900 text-white px-5 py-3 rounded-xl shadow-lg border-l-4 border-indigo-500">
            <p className="text-xs text-indigo-300 font-bold uppercase mb-1">Распознано:</p>
            <p className="text-sm italic">{liveTranscription || "Слушаю..."}</p>
          </div>
        )}

        <div className="grid grid-cols-5 gap-2">
          {LAB_ORDER.map((key, index) => (
            <button key={key} onClick={() => { setCurrentStep(index); setShowResults(false); }}
              className={`p-3 rounded-2xl border-2 text-center transition-all ${currentStep === index ? 'bg-indigo-600 border-indigo-600 text-white shadow-lg' : labData[key] ? 'bg-white border-emerald-100 text-slate-800' : 'bg-white border-slate-200 opacity-60'}`}>
              <div className={`text-[8px] font-black uppercase mb-1 truncate ${currentStep === index ? 'text-indigo-100' : 'text-slate-400'}`}>{LAB_LABELS[key]}</div>
              <div className="text-lg font-mono font-black">{labData[key] || '0'}</div>
            </button>
          ))}
        </div>

        {!showResults && (
          <div className="bg-white p-8 rounded-3xl shadow-xl border border-slate-100 text-center">
            <h2 className="text-lg font-black text-slate-400 uppercase tracking-widest mb-6">{LAB_LABELS[LAB_ORDER[currentStep]]}</h2>
            <input type="text" inputMode="decimal" className="w-full text-center text-6xl font-mono font-black py-4 bg-slate-50 border-b-8 border-slate-100 focus:border-indigo-600 outline-none rounded-2xl"
              value={labData[LAB_ORDER[currentStep]]} onChange={(e) => setLabData(p => ({ ...p, [LAB_ORDER[currentStep]]: e.target.value }))}
              onKeyDown={(e) => e.key === 'Enter' && (currentStep < 4 ? setCurrentStep(currentStep + 1) : setShowResults(true))} autoFocus />
            <div className="flex justify-center mt-8 gap-4">
              <button onClick={() => (currentStep < 4 ? setCurrentStep(currentStep + 1) : setShowResults(true))} className="bg-indigo-600 text-white px-12 py-4 rounded-2xl font-black uppercase shadow-xl hover:bg-indigo-700">
                {currentStep === 4 ? 'Рассчитать' : 'Продолжить'}
              </button>
            </div>
          </div>
        )}

        {showResults && results && (
          <div className="space-y-4 animate-in fade-in zoom-in-95">
            <div className="bg-white rounded-3xl overflow-hidden shadow-2xl border border-slate-100">
              <div className="bg-indigo-600 p-6 text-white flex justify-between items-center">
                <h3 className="text-lg font-black uppercase">Протокол испытаний</h3>
                <div className="flex gap-2">
                  <button onClick={() => setShowChartModal(true)} className="bg-emerald-500 hover:bg-emerald-600 text-white px-5 py-2 rounded-lg text-xs font-black uppercase transition-all shadow-lg flex items-center gap-2">
                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M7 12l3-3 3 3 4-4M8 21l4-4 4 4M3 4h18M4 4h16v12a1 1 0 01-1 1H5a1 1 0 01-1-1V4z"></path></svg>
                    Графики ({savedCharts.length})
                  </button>
                  <button onClick={exportWord} className="bg-white/20 hover:bg-white/40 px-3 py-2 rounded-lg text-[10px] font-black uppercase transition-all">Word</button>
                  <button onClick={() => setShowStylePicker(true)} className="bg-white text-indigo-700 px-4 py-2 rounded-lg text-[10px] font-black uppercase shadow-lg hover:bg-slate-50 transition-all">Презентация</button>
                </div>
              </div>
              <div className="p-6 grid grid-cols-1 md:grid-cols-2 gap-8">
                <div className="space-y-3">
                  {LAB_ORDER.map(k => (
                    <div key={k} className="flex justify-between border-b pb-1 text-sm"><span className="text-slate-500">{LAB_LABELS[k]}</span><span className="font-mono font-black">{labData[k]}</span></div>
                  ))}
                  <div className="space-y-1 pt-2">
                    <div className="flex justify-between border-b pb-1 text-sm"><span className="text-slate-500">PV (Пластическая вязкость)</span><span className="font-mono font-black">{results.pv.toFixed(2)}</span></div>
                    <div className="flex justify-between border-b pb-1 text-sm"><span className="text-slate-500">YP (Предел текучести)</span><span className="font-mono font-black">{results.yp.toFixed(2)}</span></div>
                    <div className="flex justify-between border-b pb-1 text-sm"><span className="text-slate-500">Эфф. вязкость (f600/2)</span><span className="font-mono font-black">{results.f.toFixed(2)}</span></div>
                    <div className="flex justify-between border-b pb-1 text-sm"><span className="text-slate-500">YP/PV</span><span className="font-mono font-black">{results.ypPvRatio.toFixed(2)}</span></div>
                    <div className="flex justify-between font-black text-indigo-600 bg-indigo-50 p-2 rounded mt-2"><span>ПОЕ (полная емкость)</span><span>{results.poe.toFixed(2)}</span></div>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                   {[{ l: 'Изотропия', v: results.isotropy.toFixed(4) }, { l: 'Генерация', v: results.generation.toFixed(2) }, { l: 'Загущение', v: results.thickening.toFixed(2) }, { l: 'Полнота', v: results.completeness.toFixed(2) }].map(c => (
                     <div key={c.l} className="bg-slate-50 p-4 rounded-xl border flex flex-col items-center shadow-inner">
                       <span className="text-[10px] font-black text-slate-400 uppercase">{c.l}</span>
                       <span className="text-xl font-mono font-black text-indigo-700">{c.v}</span>
                     </div>
                   ))}
                </div>
              </div>
            </div>

            <div className="bg-white rounded-3xl p-8 shadow-2xl border border-slate-100">
              <div className="flex justify-between items-center mb-6">
                <h4 className="text-xl font-black text-slate-900">Выводы по качеству</h4>
                <div className="flex flex-col items-end">
                  <button onClick={handleGetConclusions} disabled={isAnalyzing} className="text-[10px] bg-emerald-600 text-white px-4 py-2 rounded-lg font-black uppercase hover:bg-emerald-700 disabled:opacity-50">
                    {isAnalyzing ? 'Анализ...' : 'Обновить анализ'}
                  </button>
                  {analysisError && <span className="text-[10px] text-rose-500 font-bold mt-2 text-right">{analysisError}</span>}
                </div>
              </div>
              <div className="space-y-3">
                {conclusions.length === 0 && !isAnalyzing && <p className="text-center text-slate-400 py-10 italic">Нажмите кнопку выше для формирования экспертных выводов</p>}
                {conclusions.map((c, i) => {
                  const s = c.sentiment?.toLowerCase().trim() || 'neutral';
                  const isNegative = s.includes('neg') || s.includes('нег') || s.includes('bad') || s.includes('плох');
                  
                  // ПРИНУДИТЕЛЬНОЕ ИСПОЛЬЗОВАНИЕ ЯРКИХ ЦВЕТОВ: amber (коричневый оттенок) и emerald (зеленый)
                  const bg = isNegative ? 'bg-amber-100 border-amber-300' : 'bg-emerald-100 border-emerald-300';
                  const dot = isNegative ? 'bg-amber-600' : 'bg-emerald-600';
                  const text = isNegative ? 'text-amber-900' : 'text-emerald-900';
                  
                  return (
                    <div key={i} className={`flex gap-4 p-5 rounded-2xl border-2 transition-all shadow-md ${bg}`}>
                      <span className={`flex-shrink-0 w-8 h-8 rounded-full text-white flex items-center justify-center font-black shadow-sm ${dot}`}>{i+1}</span>
                      <p className={`text-sm font-bold leading-relaxed ${text}`}>{c.text}</p>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>

      {showChartModal && (
        <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md flex items-center justify-center z-[70] p-4">
          <div className="bg-white w-full max-w-5xl rounded-3xl overflow-hidden shadow-2xl flex flex-col md:flex-row h-[90vh] animate-slide-up">
            <div className="p-6 bg-slate-50 border-r w-full md:w-72 shrink-0 flex flex-col">
              <h5 className="font-black text-lg uppercase text-slate-800 mb-6 flex items-center gap-2">
                 <svg className="w-5 h-5 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
                 Анализ 3D
              </h5>
              <div className="space-y-6 flex-grow overflow-y-auto pr-2">
                <div>
                  <label className="text-[10px] font-black uppercase text-slate-500 mb-2 block tracking-widest">Ось X</label>
                  <select value={axisX} onChange={(e) => setAxisX(e.target.value as any)} className="w-full bg-white border-2 border-slate-200 rounded-xl px-4 py-3 text-sm font-bold outline-none focus:border-indigo-600 transition-all">
                    {Object.entries(CHART_VARIABLES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black uppercase text-slate-500 mb-2 block tracking-widest">Ось Y</label>
                  <select value={axisY} onChange={(e) => setAxisY(e.target.value as any)} className="w-full bg-white border-2 border-slate-200 rounded-xl px-4 py-3 text-sm font-bold outline-none focus:border-indigo-600 transition-all">
                    {Object.entries(CHART_VARIABLES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                  </select>
                </div>
                <div className="space-y-3 pt-4">
                  <button onClick={toggleChartInReport} className={`w-full py-4 rounded-xl font-black uppercase text-xs tracking-widest shadow-lg transition-all flex items-center justify-center gap-2 ${isCurrentChartInReport ? 'bg-rose-500 text-white hover:bg-rose-600' : 'bg-emerald-600 text-white hover:bg-emerald-700'}`}>
                    {isCurrentChartInReport ? "Исключить" : "В отчет"}
                  </button>
                  <p className="text-[10px] text-center text-slate-400 font-bold uppercase tracking-tighter">Сохранено: {savedCharts.length}</p>
                </div>
              </div>
              <button onClick={() => setShowChartModal(false)} className="mt-8 w-full py-4 bg-slate-900 text-white rounded-2xl font-black uppercase text-xs tracking-widest shadow-xl hover:bg-black transition-all">Закрыть</button>
            </div>
            <div className="flex-grow flex items-center justify-center bg-white p-4 relative min-h-0 overflow-hidden">
               <div ref={chartRef} className="w-full h-full"></div>
            </div>
          </div>
        </div>
      )}

      {showStylePicker && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-end md:items-center justify-center z-50 p-4">
          <div className="bg-white w-full max-w-lg rounded-3xl p-6 shadow-2xl animate-slide-up">
            <h5 className="text-lg font-black uppercase mb-4 text-center">Стиль презентации</h5>
            <div className="grid grid-cols-2 gap-3">
              {PPT_STYLES.map(s => (
                <button key={s.id} onClick={() => createPPTX(s)} className="p-4 rounded-2xl border-2 flex items-center gap-3 transition-all hover:border-indigo-600 text-left group">
                  <div className="w-8 h-8 rounded-lg shadow-inner" style={{ background: s.primary }}></div>
                  <span className="font-bold text-sm text-slate-700 group-hover:text-indigo-600">{s.name}</span>
                </button>
              ))}
            </div>
            <button onClick={() => setShowStylePicker(false)} className="w-full mt-6 py-3 font-black text-slate-400 uppercase tracking-widest text-xs">Отмена</button>
          </div>
        </div>
      )}

      {isGeneratingDoc && (
        <div className="fixed inset-0 bg-indigo-600/90 flex flex-col items-center justify-center z-[80] text-white">
          <div className="w-16 h-16 border-4 border-white/30 border-t-white rounded-full animate-spin mb-4"></div>
          <p className="text-lg font-black uppercase tracking-widest">Обработка...</p>
        </div>
      )}
    </div>
  );
};

export default App;
