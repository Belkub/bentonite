
/**
 * Извлекает число из строки, объединяя разрозненные цифры.
 * Решает проблему, когда "8 6" превращается в "6" или "8".
 */
export function parseSpokenNumber(text: string): number | null {
  if (!text) return null;

  // 1. Очистка и нормализация
  let clean = text.toLowerCase().trim();
  
  // 2. Обработка разделителей дробной части
  clean = clean.replace(/\s*(и|точка|запятая)\s*/g, '.');
  clean = clean.replace(/(и|точка|запятая)/g, '.');
  
  // 3. Замена запятых на точки
  clean = clean.replace(/,/g, '.');

  // 4. ОБЪЕДИНЕНИЕ ЦИФР
  // Ищем все группы цифр и точечный разделитель
  // Если пользователь сказал "три пять ноль", в транскрипции будет "3 5 0"
  // Мы убираем пробелы МЕЖДУ цифрами, но сохраняем структуру числа
  const digitsAndDots = clean.match(/[\d.]+/g);
  
  if (digitsAndDots) {
    // Соединяем все части, которые выглядят как компоненты одного числа
    // Например ["3", "5", "0"] -> "350"
    const combined = digitsAndDots.join('');
    
    // Пытаемся распарсить результат
    const num = parseFloat(combined);
    return isNaN(num) ? null : num;
  }
  
  return null;
}
