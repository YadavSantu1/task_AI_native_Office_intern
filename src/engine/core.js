// ------------------------------
// Spreadsheet Engine
// ------------------------------

function colLabelToIndex(label) {
  let index = 0;
  for (let i = 0; i < label.length; i++) {
    const charCode = label.charCodeAt(i);
    if (charCode < 65 || charCode > 90) return -1;
    index = index * 26 + (charCode - 65 + 1);
  }
  return index - 1;
}

function toJsValue(value) {
  if (value === undefined || value === null || value === "") return 0;
  if (typeof value === "number") return value;
  const trimmed = String(value).trim();
  if (trimmed === "") return 0;
  const num = Number(trimmed);
  if (!Number.isNaN(num)) return num;
  return JSON.stringify(trimmed);
}

function safeEvaluate(expression) {
  // Only allow a limited subset of characters to avoid arbitrary code execution.
  const safe = expression.replace(/[^0-9+\-*/().,\s"'\\]/g, "");
  return Function(`"use strict"; return (${safe});`)();
}

export function createEngine(rows = 50, cols = 50) {
  const cells = new Map();

  const clamp = (value, max) => Math.max(0, Math.min(max - 1, value));
  const key = (r, c) => `${r},${c}`;

  const getRaw = (r, c) => {
    const cell = cells.get(key(r, c));
    return cell?.raw ?? "";
  };

  const setRaw = (r, c, raw) => {
    if (raw === "" || raw === null || raw === undefined) {
      cells.delete(key(r, c));
    } else {
      cells.set(key(r, c), { raw: String(raw) });
    }
  };

  const computeAll = () => {
    const memo = new Map();
    const visiting = new Set();

    const evaluateCell = (r, c) => {
      const k = key(r, c);
      if (memo.has(k)) return memo.get(k);
      if (visiting.has(k)) {
        const result = { computed: undefined, error: "#CYCLE" };
        memo.set(k, result);
        return result;
      }

      visiting.add(k);
      const raw = getRaw(r, c);
      let computed = raw;
      let error = null;

      if (typeof raw === "string" && raw.startsWith("=")) {
        const expr = raw.slice(1);
        try {
          const replaced = expr.replace(/([A-Z]+)(\d+)/g, (match, colLabel, rowStr) => {
            const refCol = colLabelToIndex(colLabel.toUpperCase());
            const refRow = parseInt(rowStr, 10) - 1;

            if (refCol < 0 || refRow < 0) return 0;

            const rClamped = clamp(refRow, rows);
            const cClamped = clamp(refCol, cols);
            const ref = evaluateCell(rClamped, cClamped);

            if (ref.error) throw new Error(ref.error);

            return toJsValue(ref.computed ?? getRaw(rClamped, cClamped));
          });

          computed = safeEvaluate(replaced);
        } catch (e) {
          computed = undefined;
          error = e.message || "#ERROR";
        }
      }

      const result = { computed, error };
      memo.set(k, result);
      visiting.delete(k);
      return result;
    };

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const k = key(r, c);
        const raw = getRaw(r, c);
        if (!raw) {
          memo.set(k, { computed: "", error: null });
          continue;
        }
        evaluateCell(r, c);
      }
    }

    for (const [k, { computed, error }] of memo.entries()) {
      const [r, c] = k.split(",").map(Number);
      const raw = getRaw(r, c);
      if (!cells.has(k) && raw === "") continue;
      const existing = cells.get(k) ?? { raw: "" };
      cells.set(k, { raw: existing.raw, computed, error });
    }
  };

  const getCell = (r, c) => {
    const row = clamp(r, rows);
    const col = clamp(c, cols);
    const cell = cells.get(key(row, col)) ?? { raw: "" };
    return {
      raw: cell.raw ?? "",
      computed: cell.computed ?? cell.raw ?? "",
      error: cell.error ?? null,
    };
  };

  const setCell = (r, c, raw) => {
    const row = clamp(r, rows);
    const col = clamp(c, cols);
    setRaw(row, col, raw);
    computeAll();
  };

  const updateCells = (changes) => {
    changes.forEach(({ r, c, raw }) => {
      const row = clamp(r, rows);
      const col = clamp(c, cols);
      setRaw(row, col, raw);
    });
    computeAll();
  };

  const getDisplayValue = (r, c) => {
    const cell = getCell(r, c);
    if (cell.error) return cell.error;
    return cell.computed ?? cell.raw ?? "";
  };

  // Trigger initial compute to populate computed values.
  computeAll();

  return {
    rows,
    cols,
    getCell,
    setCell,
    updateCells,
    getDisplayValue,
  };
}

// ------------------------------
// Column Sorting
// ------------------------------

export function sortRows(rows, columnIndex, direction, getValue) {
  if (!direction || columnIndex === null) return rows;

  const sorted = [...rows].sort((a, b) => {
    const valA = getValue(a, columnIndex);
    const valB = getValue(b, columnIndex);

    const numA = Number(valA);
    const numB = Number(valB);
    const isNumberA = !Number.isNaN(numA);
    const isNumberB = !Number.isNaN(numB);

    if (isNumberA && isNumberB) {
      if (numA === numB) return 0;
      return direction === "asc" ? numA - numB : numB - numA;
    }

    if (valA === valB) return 0;

    if (direction === "asc") {
      return valA > valB ? 1 : -1;
    }

    return valA < valB ? 1 : -1;
  });

  return sorted;
}

// ------------------------------
// Column Filtering
// ------------------------------

export function applyFilters(rows, filters, getValue) {
  return rows.filter((row) => {
    return Object.keys(filters).every((col) => {
      const allowed = filters[col];
      if (!allowed || allowed.length === 0) return true;
      const value = getValue(row, Number(col));
      return allowed.includes(value);
    });
  });
}
