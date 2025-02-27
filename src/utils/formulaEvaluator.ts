import { SheetState } from '../types/sheet';

export function evaluateFormula(formula: string, state: SheetState): string | number {
  if (!formula.startsWith('=')) return formula;

  const cleanFormula = formula.substring(1).toUpperCase();
  
  // Handle basic mathematical functions
  if (cleanFormula.startsWith('SUM(')) {
    return evaluateSum(cleanFormula, state);
  } else if (cleanFormula.startsWith('AVERAGE(')) {
    return evaluateAverage(cleanFormula, state);
  } else if (cleanFormula.startsWith('MAX(')) {
    return evaluateMax(cleanFormula, state);
  } else if (cleanFormula.startsWith('MIN(')) {
    return evaluateMin(cleanFormula, state);
  } else if (cleanFormula.startsWith('COUNT(')) {
    return evaluateCount(cleanFormula, state);
  }

  // Handle data quality functions
  if (cleanFormula.startsWith('TRIM(')) {
    return evaluateTrim(cleanFormula, state);
  } else if (cleanFormula.startsWith('UPPER(')) {
    return evaluateUpper(cleanFormula, state);
  } else if (cleanFormula.startsWith('LOWER(')) {
    return evaluateLower(cleanFormula, state);
  }

  return formula;
}

function getRangeValues(range: string, state: SheetState): number[] {
  const values: number[] = [];
  const [start, end] = range.split(':');
  
  // Convert column letters to numbers (e.g., 'A' -> 0, 'B' -> 1)
  const startColMatch = start.match(/[A-Z]+/);
  const startRowMatch = start.match(/\d+/);
  const endColMatch = end.match(/[A-Z]+/);
  const endRowMatch = end.match(/\d+/);

  if (!startColMatch || !startRowMatch || !endColMatch || !endRowMatch) {
    throw new Error('Invalid range format');
  }

  const startCol = startColMatch[0];
  const startRow = parseInt(startRowMatch[0]) - 1;
  const endCol = endColMatch[0];
  const endRow = parseInt(endRowMatch[0]) - 1;
  
  const startColIndex = columnToIndex(startCol);
  const endColIndex = columnToIndex(endCol);
  
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startColIndex; col <= endColIndex; col++) {
      const cellId = `${indexToColumn(col)}${row + 1}`;
      const cell = state.data[cellId];
      if (cell) {
        const value = cell.formula ? evaluateFormula(cell.formula, state) : cell.value;
        values.push(Number(value) || 0);
      }
    }
  }
  
  return values;
}

// Helper functions for column conversion
function columnToIndex(col: string): number {
  return col.split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function indexToColumn(index: number): string {
  let column = '';
  index++;
  while (index > 0) {
    const remainder = (index - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    index = Math.floor((index - 1) / 26);
  }
  return column;
}

// Mathematical Functions
function evaluateSum(formula: string, state: SheetState): number {
  const range = extractRange(formula);
  const values = getRangeValues(range, state);
  return values.reduce((sum, val) => sum + val, 0);
}

function evaluateAverage(formula: string, state: SheetState): number {
  const range = extractRange(formula);
  const values = getRangeValues(range, state);
  return values.length ? values.reduce((sum, val) => sum + val, 0) / values.length : 0;
}

function evaluateMax(formula: string, state: SheetState): number {
  const range = extractRange(formula);
  const values = getRangeValues(range, state);
  return values.length ? Math.max(...values) : 0;
}

function evaluateMin(formula: string, state: SheetState): number {
  const range = extractRange(formula);
  const values = getRangeValues(range, state);
  return values.length ? Math.min(...values) : 0;
}

function evaluateCount(formula: string, state: SheetState): number {
  const range = extractRange(formula);
  const values = getRangeValues(range, state);
  return values.filter(val => typeof val === 'number').length;
}

// Data Quality Functions
function evaluateTrim(formula: string, state: SheetState): string {
  const value = extractSingleValue(formula, state);
  return String(value).trim();
}

function evaluateUpper(formula: string, state: SheetState): string {
  const value = extractSingleValue(formula, state);
  return String(value).toUpperCase();
}

function evaluateLower(formula: string, state: SheetState): string {
  const value = extractSingleValue(formula, state);
  return String(value).toLowerCase();
}

// Helper Functions
function extractRange(formula: string): string {
  const match = formula.match(/\((.*?)\)/);
  return match ? match[1] : '';
}

function extractSingleValue(formula: string, state: SheetState): string {
  const cellId = extractRange(formula);
  const cell = state.data[cellId];
  return cell ? cell.value : '';
}