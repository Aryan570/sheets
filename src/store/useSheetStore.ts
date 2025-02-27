import { create } from 'zustand';
import { produce } from 'immer';
import { SheetState, CellData, CellStyle } from '../types/sheet';

const DEFAULT_ROWS = 100;
const DEFAULT_COLS = 26;

const createEmptyCell = (): CellData => ({
  value: '',
  formula: '',
  style: {
    bold: false,
    italic: false,
    fontSize: 12,
    color: '#000000'
  }
});

const initialState: SheetState = {
  data: {},
  selectedCell: null,
  selectedRange: null,
  columnWidths: {},
  rowHeights: {},
  numRows: DEFAULT_ROWS,
  numCols: DEFAULT_COLS,
};

export const useSheetStore = create<{
  state: SheetState;
  setCellValue: (cellId: string, value: string) => void;
  setCellFormula: (cellId: string, formula: string) => void;
  setCellStyle: (cellId: string, style: Partial<CellStyle>) => void;
  setSelectedCell: (cellId: string | null) => void;
  setSelectedRange: (range: string[] | null) => void;
  setColumnWidth: (colId: string, width: number) => void;
  setRowHeight: (rowId: number, height: number) => void;
}>((set) => ({
  state: initialState,
  
  setCellValue: (cellId, value) =>
    set(
      produce((state) => {
        if (!state.state.data[cellId]) {
          state.state.data[cellId] = createEmptyCell();
        }
        state.state.data[cellId].value = value;
      })
    ),

  setCellFormula: (cellId, formula) =>
    set(
      produce((state) => {
        if (!state.state.data[cellId]) {
          state.state.data[cellId] = createEmptyCell();
        }
        state.state.data[cellId].formula = formula;
      })
    ),

  setCellStyle: (cellId, style) =>
    set(
      produce((state) => {
        if (!state.state.data[cellId]) {
          state.state.data[cellId] = createEmptyCell();
        }
        state.state.data[cellId].style = {
          ...state.state.data[cellId].style,
          ...style,
        };
      })
    ),

  setSelectedCell: (cellId) =>
    set(
      produce((state) => {
        state.state.selectedCell = cellId;
      })
    ),

  setSelectedRange: (range) =>
    set(
      produce((state) => {
        state.state.selectedRange = range;
      })
    ),

  setColumnWidth: (colId, width) =>
    set(
      produce((state) => {
        state.state.columnWidths[colId] = width;
      })
    ),

  setRowHeight: (rowId, height) =>
    set(
      produce((state) => {
        state.state.rowHeights[rowId] = height;
      })
    ),
}));