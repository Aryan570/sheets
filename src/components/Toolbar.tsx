import React from 'react';
import { useSheetStore } from '../store/useSheetStore';
import { Bold, Italic, Type, Palette } from 'lucide-react';

export const Toolbar: React.FC = () => {
  const { state, setCellStyle } = useSheetStore();
  const selectedCell = state.selectedCell;

  const handleBold = () => {
    if (!selectedCell) return;
    const currentStyle = state.data[selectedCell]?.style || { bold: false };
    setCellStyle(selectedCell, { bold: !currentStyle.bold });
  };

  const handleItalic = () => {
    if (!selectedCell) return;
    const currentStyle = state.data[selectedCell]?.style || { italic: false };
    setCellStyle(selectedCell, { italic: !currentStyle.italic });
  };

  const handleFontSize = (e: React.ChangeEvent<HTMLSelectElement>) => {
    if (!selectedCell) return;
    setCellStyle(selectedCell, { fontSize: parseInt(e.target.value) });
  };

  const handleColor = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!selectedCell) return;
    setCellStyle(selectedCell, { color: e.target.value });
  };

  return (
    <div className="flex items-center gap-2 p-2 border-b border-gray-200 bg-white">
      <button
        onClick={handleBold}
        className="p-2 hover:bg-gray-100 rounded"
      >
        <Bold size={18} />
      </button>
      <button
        onClick={handleItalic}
        className="p-2 hover:bg-gray-100 rounded"
      >
        <Italic size={18} />
      </button>
      <div className="flex items-center gap-1">
        <Type size={18} />
        <select
          onChange={handleFontSize}
          className="border rounded p-1"
          defaultValue="12"
        >
          {[8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24].map(size => (
            <option key={size} value={size}>{size}</option>
          ))}
        </select>
      </div>
      <div className="flex items-center gap-1">
        <Palette size={18} />
        <input
          type="color"
          onChange={handleColor}
          className="w-6 h-6"
        />
      </div>
    </div>
  );
};