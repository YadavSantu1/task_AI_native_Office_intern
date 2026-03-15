import { useState, useRef, useCallback, useMemo, useEffect } from "react";
import "./App.css";
import { createEngine, sortRows, applyFilters } from "./engine/core.js";

const TOTAL_ROWS = 50;
const TOTAL_COLS = 50;

export default function App() {
  const [engine] = useState(() => {
    const saved = localStorage.getItem("spreadsheet_state");
    if (!saved) return createEngine(TOTAL_ROWS, TOTAL_COLS);

    try {
      const data = JSON.parse(saved);
      const rows = data.rows ?? TOTAL_ROWS;
      const cols = data.cols ?? TOTAL_COLS;
      const eng = createEngine(rows, cols);
      if (Array.isArray(data.cells) && data.cells.length) {
        eng.updateCells(
          data.cells.map((cell) => ({ r: cell.r, c: cell.c, raw: cell.value }))
        );
      }
      return eng;
    } catch {
      localStorage.removeItem("spreadsheet_state");
      return createEngine(TOTAL_ROWS, TOTAL_COLS);
    }
  });

  const [cellStyles, _setCellStyles] = useState(() => {
    const saved = localStorage.getItem("spreadsheet_state");
    if (!saved) return {};
    try {
      const data = JSON.parse(saved);
      return data.styles || {};
    } catch {
      return {};
    }
  });

  const [version, setVersion] = useState(0);
  const forceRerender = useCallback(() => setVersion((v) => v + 1), []);

  const [selectedCell, setSelectedCell] = useState(null);
  const [selectionAnchor, setSelectionAnchor] = useState(null);
  const [selectionEnd, setSelectionEnd] = useState(null);

  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState("");

  const [sortConfig, setSortConfig] = useState({ column: null, direction: null });
  const [filters, setFilters] = useState({});
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);

  const historyRef = useRef([]);
  const historyIndexRef = useRef(0);
  const [, setHistoryVersion] = useState(0);

  const internalClipboardRef = useRef(null);
  const lastCopiedInternalRef = useRef(false);

  const cellInputRef = useRef(null);

  const getSelectionRange = useCallback(() => {
    if (!selectionAnchor || !selectionEnd) return null;
    const startRow = Math.min(selectionAnchor.r, selectionEnd.r);
    const endRow = Math.max(selectionAnchor.r, selectionEnd.r);
    const startCol = Math.min(selectionAnchor.c, selectionEnd.c);
    const endCol = Math.max(selectionAnchor.c, selectionEnd.c);
    return { startRow, endRow, startCol, endCol };
  }, [selectionAnchor, selectionEnd]);

  const isCellInSelection = (r, c) => {
    const range = getSelectionRange();
    if (!range) return false;
    return (
      r >= range.startRow &&
      r <= range.endRow &&
      c >= range.startCol &&
      c <= range.endCol
    );
  };

  const pushHistory = useCallback((changes) => {
    if (!changes || changes.length === 0) return;
    const past = historyRef.current.slice(0, historyIndexRef.current);
    past.push(changes);
    historyRef.current = past;
    historyIndexRef.current = past.length;
    setHistoryVersion((v) => v + 1);
  }, []);

  const applyChanges = useCallback(
    (changes, useAfter = true) => {
      if (!changes || changes.length === 0) return;
      const payload = changes.map((change) => ({
        r: change.r,
        c: change.c,
        raw: useAfter ? change.after : change.before,
      }));
      engine.updateCells(payload);
      forceRerender();
    },
    [engine, forceRerender]
  );

  const undo = useCallback(() => {
    if (historyIndexRef.current === 0) return;
    const idx = historyIndexRef.current - 1;
    const changes = historyRef.current[idx];
    applyChanges(changes, false);
    historyIndexRef.current = idx;
    setHistoryVersion((v) => v + 1);
  }, [applyChanges]);

  const redo = useCallback(() => {
    if (historyIndexRef.current >= historyRef.current.length) return;
    const changes = historyRef.current[historyIndexRef.current];
    applyChanges(changes, true);
    historyIndexRef.current += 1;
    setHistoryVersion((v) => v + 1);
  }, [applyChanges]);

  const getVisibleRows = useMemo(() => {
    // Use `version` to ensure view updates when engine content changes.
    void version;

    const allRows = Array.from({ length: engine.rows }, (_, i) => i);
    const filtered = applyFilters(allRows, filters, (r, c) =>
      engine.getDisplayValue(r, c)
    );
    return sortRows(
      filtered,
      sortConfig.column,
      sortConfig.direction,
      (r, c) => engine.getDisplayValue(r, c)
    );
  }, [engine, filters, sortConfig, version]);

  const getColumnLabel = (col) => {
    let label = "";
    let num = col + 1;
    while (num > 0) {
      num--;
      label = String.fromCharCode(65 + (num % 26)) + label;
      num = Math.floor(num / 26);
    }
    return label;
  };

  const getFilterOptions = (col) => {
    const values = new Set();
    for (let r = 0; r < engine.rows; r++) {
      values.add(String(engine.getDisplayValue(r, col)));
    }
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  };

  const toggleFilterValue = (col, value) => {
    setFilters((prev) => {
      const prevValues = new Set(prev[col] || []);
      if (prevValues.has(value)) {
        prevValues.delete(value);
      } else {
        prevValues.add(value);
      }
      return {
        ...prev,
        [col]: Array.from(prevValues),
      };
    });
  };

  const clearFilter = (col) => {
    setFilters((prev) => {
      const next = { ...prev };
      delete next[col];
      return next;
    });
  };

  const handleSortColumn = (colIndex) => {
    let direction = "asc";
    if (sortConfig.column === colIndex && sortConfig.direction === "asc") {
      direction = "desc";
    } else if (sortConfig.column === colIndex && sortConfig.direction === "desc") {
      direction = null;
    }
    setSortConfig({ column: colIndex, direction });
  };

  const startEditing = (row, col) => {
    const cell = engine.getCell(row, col);
    setSelectedCell({ r: row, c: col });
    setSelectionAnchor({ r: row, c: col });
    setSelectionEnd({ r: row, c: col });
    setEditingCell({ r: row, c: col });
    setEditValue(cell.raw);
    setTimeout(() => cellInputRef.current?.focus(), 0);
  };

  const commitEdit = (row, col, value) => {
    const cell = engine.getCell(row, col);
    const newValue = value ?? editValue;
    if (cell.raw !== newValue) {
      const changes = [{ r: row, c: col, before: cell.raw, after: newValue }];
      pushHistory(changes);
      applyChanges(changes, true);
    }
    setEditingCell(null);
  };

  const handleCellClick = (row, col, event) => {
    if (event.shiftKey && selectedCell) {
      // Shift + click extends selection range
      setSelectionEnd({ r: row, c: col });
      setSelectedCell({ r: row, c: col });
      setEditingCell(null);
      setEditValue(engine.getCell(row, col).raw);
      return;
    }

    // Single click enters edit mode immediately
    startEditing(row, col);
  };

  const getSelectionData = useCallback(() => {
    const range = getSelectionRange();
    if (!range) return null;
    const data = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(engine.getDisplayValue(r, c));
      }
      data.push(row);
    }
    return data;
  }, [engine, getSelectionRange]);

  const handleCopy = useCallback(() => {
    const selection = getSelectionData();
    if (!selection) return;
    const text = selection.map((row) => row.join("\t")).join("\n");
    navigator.clipboard.writeText(text).catch(() => {});
    internalClipboardRef.current = selection;
    lastCopiedInternalRef.current = true;
  }, [getSelectionData]);

  const handlePaste = useCallback(
    (event) => {
      if (!selectedCell) return;

      event.preventDefault();

      const data =
        lastCopiedInternalRef.current && internalClipboardRef.current
          ? internalClipboardRef.current
          : event.clipboardData
              .getData("text")
              .replace(/\r/g, "")
              .split("\n")
              .map((row) => row.split("\t"));

      const changes = [];

      data.forEach((row, i) => {
        row.forEach((value, j) => {
          const r = selectedCell.r + i;
          const c = selectedCell.c + j;
          if (r < engine.rows && c < engine.cols) {
            const before = engine.getCell(r, c).raw;
            if (before !== value) {
              changes.push({ r, c, before, after: value });
            }
          }
        });
      });

      if (changes.length) {
        pushHistory(changes);
        applyChanges(changes, true);
      }

      lastCopiedInternalRef.current = false;
    },
    [selectedCell, engine, pushHistory, applyChanges]
  );

  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.ctrlKey && (e.key === "z" || e.key === "Z")) {
        e.preventDefault();
        if (editingCell) {
          // Cancel inline editing (undo the in-progress edit)
          const original = engine.getCell(editingCell.r, editingCell.c).raw;
          setEditValue(original);
          setEditingCell(null);
          return;
        }
        undo();
        return;
      }

      if (e.ctrlKey && (e.key === "y" || e.key === "Y")) {
        e.preventDefault();
        redo();
        return;
      }

      if (editingCell) return;

      if (e.ctrlKey && (e.key === "c" || e.key === "C")) {
        e.preventDefault();
        handleCopy();
      }
    };

    const handlePasteInternal = (e) => {
      if (editingCell) return;
      handlePaste(e);
    };

    window.addEventListener("keydown", handleKeyDown);
    window.addEventListener("paste", handlePasteInternal);

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      window.removeEventListener("paste", handlePasteInternal);
    };
  }, [handleCopy, handlePaste, undo, redo, editingCell, engine]);

  useEffect(() => {
    const timer = setTimeout(() => {
      const data = {
        rows: engine.rows,
        cols: engine.cols,
        cells: [],
        styles: cellStyles,
      };

      for (let r = 0; r < engine.rows; r++) {
        for (let c = 0; c < engine.cols; c++) {
          const cell = engine.getCell(r, c);
          if (cell.raw) {
            data.cells.push({ r, c, value: cell.raw });
          }
        }
      }

      try {
        localStorage.setItem("spreadsheet_state", JSON.stringify(data));
      } catch (err) {
        console.warn("Unable to persist spreadsheet state", err);
      }
    }, 500);

    return () => clearTimeout(timer);
  }, [version, cellStyles, engine]);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (!e.target.closest(".filter-menu") && !e.target.closest(".filter-btn")) {
        setActiveFilterColumn(null);
      }
    };
    window.addEventListener("mousedown", handleClickOutside);
    return () => window.removeEventListener("mousedown", handleClickOutside);
  }, []);

  return (
    <div className="app-wrapper">
      <h2>Spreadsheet</h2>

      <div className="grid-scroll">
        <table className="grid-table">
          <thead>
            <tr>
              <th className="col-header-blank"></th>
              {Array.from({ length: engine.cols }, (_, colIndex) => {
                const isActive = activeFilterColumn === colIndex;
                const filterOptions = getFilterOptions(colIndex);
                const selectedValues = filters[colIndex] || [];
                const sortArrow =
                  sortConfig.column === colIndex
                    ? sortConfig.direction === "asc"
                      ? " ▲"
                      : sortConfig.direction === "desc"
                      ? " ▼"
                      : ""
                    : "";

                return (
                  <th
                    key={colIndex}
                    className="col-header"
                    style={{ position: "relative" }}
                  >
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: 4,
                      }}
                    >
                      <span
                        style={{ cursor: "pointer" }}
                        onClick={() => handleSortColumn(colIndex)}
                      >
                        {getColumnLabel(colIndex)}{sortArrow}
                      </span>
                      <button
                        type="button"
                        className="filter-btn"
                        onClick={(e) => {
                          e.stopPropagation();
                          setActiveFilterColumn((prev) =>
                            prev === colIndex ? null : colIndex
                          );
                        }}
                        aria-label="Filter"
                      >
                        ⏷
                      </button>
                    </div>

                    {isActive ? (
                      <div className="filter-menu">
                        <div className="filter-menu-header">
                          <div style={{ fontWeight: 600 }}>Filter</div>
                          <button
                            type="button"
                            className="toolbar-btn"
                            onClick={() => clearFilter(colIndex)}
                          >
                            Clear
                          </button>
                        </div>
                        <div className="filter-menu-options">
                          {filterOptions.map((opt) => {
                            const value = opt === "" ? "(blank)" : opt;
                            const checked = selectedValues.includes(opt);
                            return (
                              <label key={opt} className="filter-option">
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => toggleFilterValue(colIndex, opt)}
                                />
                                <span>{value}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    ) : null}
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody>
            {getVisibleRows.map((rowIndex) => (
              <tr key={rowIndex}>
                <td className="row-header">{rowIndex + 1}</td>
                {Array.from({ length: engine.cols }, (_, colIndex) => {
                  const isSelected = isCellInSelection(rowIndex, colIndex);
                  const isEditing =
                    editingCell?.r === rowIndex &&
                    editingCell?.c === colIndex;
                  const cellData = engine.getCell(rowIndex, colIndex);
                  const displayValue =
                    cellData.error || cellData.computed || cellData.raw || "";
                  const cellStyle = cellStyles[`${rowIndex},${colIndex}`] || {};

                  return (
                    <td
                      key={colIndex}
                      className={isSelected ? "cell selected" : "cell"}
                      style={cellStyle}
                      onClick={(e) => handleCellClick(rowIndex, colIndex, e)}
                      onDoubleClick={() => startEditing(rowIndex, colIndex)}
                    >
                      {isEditing ? (
                        <input
                          ref={cellInputRef}
                          className="cell-input"
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          onBlur={() => commitEdit(rowIndex, colIndex)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") {
                              commitEdit(rowIndex, colIndex);
                              setEditingCell(null);
                            }
                          }}
                        />
                      ) : (
                        <div
                          className={
                            "cell-display" + (cellData.error ? " error" : "")
                          }
                        >
                          {displayValue}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="footer-hint">
        Tip: Use <b>Ctrl+C</b> and <b>Ctrl+V</b> to copy/paste. <b>Ctrl+Z</b> to undo.
      </div>
    </div>
  );
}
