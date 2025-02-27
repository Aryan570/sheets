import React, { useState, useEffect } from 'react';
import { useSheetStore } from '../store/useSheetStore';

export const FormulaBar: React.FC = () => {
  const { state, setCellFormula } = useSheetStore();
  const [formula, setFormula] = useState('');

  useEffect(() => {
    if (state.selectedCell) {
      setFormula(state.data[state.selectedCell]?.formula || '');
    }
  }, [state.selectedCell, state.data]);

  const handleFormulaChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFormula(e.target.value);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && state.selectedCell) {
      setCellFormula(state.selectedCell, formula);
    }
  };

  return (
    <div className="flex items-center gap-2 p-2 border-b border-gray-200 bg-white">
      <span className="font-mono">=</span>
      <input
        type="text"
        value={formula}
        onChange={handleFormulaChange}
        onKeyDown={handleKeyDown}
        placeholder="Enter formula"
        className="flex-1 px-2 py-1 border rounded"
      />
    </div>
  );
};