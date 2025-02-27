import React, { useState, useCallback } from 'react';
import { useSheetStore } from '../store/useSheetStore';
import { evaluateFormula } from '../utils/formulaEvaluator';
import clsx from 'clsx';

interface CellProps {
  id: string;
  row: number;
  col: number;
}

export const Cell: React.FC<CellProps> = ({ id, row, col }) => {
  const { state, setCellValue, setCellFormula, setSelectedCell } = useSheetStore();
  const [isEditing, setIsEditing] = useState(false);
  
  const cell = state.data[id] || {
    value: '',
    formula: '',
    style: { bold: false, italic: false, fontSize: 12, color: '#000000' }
  };

  const isSelected = state.selectedCell === id;

  const handleDoubleClick = () => {
    setIsEditing(true);
  };

  const handleBlur = () => {
    setIsEditing(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      setIsEditing(false);
      const value = (e.target as HTMLInputElement).value;
      if (value.startsWith('=')) {
        setCellFormula(id, value);
      } else {
        setCellValue(id, value);
      }
    }
  };

  const displayValue = useCallback(() => {
    if (cell.formula) {
      return evaluateFormula(cell.formula, state);
    }
    return cell.value;
  }, [cell, state]);

  return (
    <div
      className={clsx(
        'border border-gray-200 p-1 min-w-[100px] h-[25px] relative',
        isSelected && 'border-2 border-blue-500',
        cell.style.bold && 'font-bold',
        cell.style.italic && 'italic'
      )}
      style={{
        fontSize: `${cell.style.fontSize}px`,
        color: cell.style.color
      }}
      onClick={() => setSelectedCell(id)}
      onDoubleClick={handleDoubleClick}
    >
      {isEditing ? (
        <input
          className="absolute inset-0 w-full h-full p-1"
          defaultValue={cell.formula || cell.value}
          onBlur={handleBlur}
          onKeyDown={handleKeyDown}
          autoFocus
        />
      ) : (
        <div className="truncate">{displayValue()}</div>
      )}
    </div>
  );
};