import React from 'react';
import { Cell } from './Cell';
import { useSheetStore } from '../store/useSheetStore';

export const Grid: React.FC = () => {
  const { state } = useSheetStore();
  const { numRows, numCols } = state;

  const getColumnLabel = (index: number) => {
    let label = '';
    while (index >= 0) {
      label = String.fromCharCode(65 + (index % 26)) + label;
      index = Math.floor(index / 26) - 1;
    }
    return label;
  };

  return (
    <div className="overflow-auto flex-1">
      <div className="inline-block">
        {/* Header row */}
        <div className="flex">
          <div className="w-[50px] h-[25px] bg-gray-100 border-b border-r border-gray-300" />
          {Array.from({ length: numCols }).map((_, col) => (
            <div
              key={col}
              className="min-w-[100px] h-[25px] bg-gray-100 border-b border-r border-gray-300 flex items-center justify-center font-semibold"
            >
              {getColumnLabel(col)}
            </div>
          ))}
        </div>

        {/* Grid rows */}
        {Array.from({ length: numRows }).map((_, row) => (
          <div key={row} className="flex">
            <div className="w-[50px] h-[25px] bg-gray-100 border-b border-r border-gray-300 flex items-center justify-center font-semibold">
              {row + 1}
            </div>
            {Array.from({ length: numCols }).map((_, col) => (
              <Cell
                key={`${row}-${col}`}
                id={`${getColumnLabel(col)}${row + 1}`}
                row={row}
                col={col}
              />
            ))}
          </div>
        ))}
      </div>
    </div>
  );
};