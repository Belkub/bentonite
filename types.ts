
export interface LabData {
  m: string; // содержание смектита
  q: string; // обменная емкость
  w: string; // влажность
  f300: string; // фи 300
  f600: string; // фи 600
  s_equiv: string; // влага глины
  mm: string; // ММ ПАВ
  t: string; // Доля ПАВ, %
}

export type LabKey = keyof LabData;

export interface Conclusion {
  text: string;
  sentiment: 'positive' | 'neutral' | 'negative';
}

export interface SavedChart {
  id: string;
  axisX: string;
  axisY: string;
  imageData: string;
}

export interface CalculationResult {
  m: number;
  q: number;
  w: number;
  f: number;
  f600: number;
  pv: number;
  yp: number;
  poe: number;
  ypPvRatio: number;
  s: number;
  isotropy: number;
  generation: number;
  thickening: number;
  completeness: number;
  equivalent?: number;
}

export const LAB_LABELS: Record<LabKey, string> = {
  m: 'Содержание смектита',
  q: 'Обменная емкость (КОЕ)',
  w: 'Влажность (%)',
  f300: 'Фи 300',
  f600: 'Фи 600',
  s_equiv: 'Влага глины, %',
  mm: 'Мол масса ПАВ, г/моль',
  t: 'Доля ПАВ, %'
};

export const LAB_ORDER: LabKey[] = ['m', 'q', 'w', 'f300', 'f600', 's_equiv', 'mm', 't'];
