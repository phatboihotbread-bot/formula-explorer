let state = {
  active: false,
  workbookName: "",
  originSheet: "",
  originAddress: "",
  formula: "",
  refs: [],
  pos: 0
};

Office.onReady(() => {
  document.getElementById("left").onclick = () => navigate(-1);
  document.getElementById("right").onclick = () => navigate(1);

  // Complete the Office command event explicitly after each shortcut.
  // This prevents Excel from leaving the ExecuteFunction command pending,
  // which can produce the macOS invalid-action beep even though navigation ran.
  Office.actions.associate("FormulaExplorer.GoLeft", async (event) => {
    try {
      await ensureExplorerVisible();
      await navigate(-1);
    } finally {
      if (event && typeof event.completed === "function") event.completed();
    }
  });
  Office.actions.associate("FormulaExplorer.GoRight", async (event) => {
    try {
      await ensureExplorerVisible();
      await navigate(1);
    } finally {
      if (event && typeof event.completed === "function") event.completed();
    }
  });

Office.actions.associate("FormulaExplorer.Close", async (event) => {
  try {
    await closeExplorer();
  } finally {
    if (event && typeof event.completed === "function") event.completed();
  }
});
  const exitBtn = document.getElementById("exit");
  if (exitBtn) exitBtn.onclick = () => exitNavigation();
  // If Excel for Mac leaves keyboard focus in the add-in webview after an
  // ExecuteFunction shortcut, consume navigation keys here instead of letting
  // macOS play the disabled/beep sound. Arrow keys move the Excel selection
  // normally and end the trace; Enter/Escape simply accept the current cell.
  document.addEventListener("keydown", async (e) => {
    if (e.key === "Escape" || e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      await exitNavigation();
      releaseWebFocus();
      return;
    }
    const moves = {ArrowUp:[-1,0], ArrowDown:[1,0], ArrowLeft:[0,-1], ArrowRight:[0,1]};
    if (moves[e.key] && !e.metaKey && !e.altKey && !e.ctrlKey) {
      e.preventDefault();
      e.stopPropagation();
      const [dr, dc] = moves[e.key];
      await moveSelection(dr, dc);
      await exitNavigation();
      releaseWebFocus();
    }
  }, true);
});

function esc(s) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

function parseFormula(formula) {
  const refs = [];
  // Direct A1 references, optional quoted/unquoted sheet prefix, optional range.
  // Negative lookbehind avoids grabbing pieces embedded in identifiers.
  const re = /(?<![A-Za-z0-9_])(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)(?![A-Za-z0-9_])/g;
  let m;
  while ((m = re.exec(formula)) !== null) {
    refs.push({
      token: m[0],
      sheet: m[1] ? m[1].replace(/''/g,"'") : (m[2] || ""),
      address: m[3],
      start: m.index,
      length: m[0].length
    });
  }
  return refs;
}

function normalizeAddress(address) {
  return String(address || "").replace(/\$/g, "").toUpperCase();
}

function expectedLocation() {
  if (!state.active) return null;
  if (state.pos === 0) {
    return { sheet: state.originSheet, address: state.originAddress };
  }
  const r = state.refs[state.pos - 1];
  return { sheet: r.sheet || state.originSheet, address: r.address };
}

async function selectionStillBelongsToSession() {
  if (!state.active) return false;
  return Excel.run(async context => {
    // IMPORTANT: compare the WHOLE selected range, not getActiveCell().
    // When Formula Explorer selects J14:J17, getActiveCell() is only J14.
    // The old comparison therefore killed the original trace and accidentally
    // started tracing J14's own formula. That was the range-loop bug.
    const selected = context.workbook.getSelectedRange();
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    selected.load("address");
    sheet.load("name");
    await context.sync();

    const expected = expectedLocation();
    const selectedAddress = selected.address.includes("!") ? selected.address.split("!").slice(1).join("!") : selected.address;
    return sheet.name === expected.sheet && normalizeAddress(selectedAddress) === normalizeAddress(expected.address);
  });
}

function clearSession() {
  state.active = false;
  state.workbookName = "";
  state.originSheet = "";
  state.originAddress = "";
  state.formula = "";
  state.refs = [];
  state.pos = 0;
}

async function exitNavigation() {
  clearSession();
  document.getElementById("position").textContent = "Navigation ended";
  document.getElementById("formula").textContent = "Current cell kept. Press ⌥⌘J or ⌥⌘K to trace from it if it contains a formula.";
  setStatus("Formula Explorer is idle.");
}


async function resolveStructuredReferences(formula, originSheet, originAddress) {
  // Supports the most useful table-reference forms for precedent navigation:
  //   [@[Column]]              current row, current table
  //   Table1[@[Column]]        current row, named table
  //   Table1[Column]           table data column
  //   Table1[[#Data],[Column]] table data column
  //
  // The current-row form is the key case: it resolves to the actual cell in
  // the same table row as the formula cell.
  const candidates = [];

  // Current-row references, with optional table name.
  const currentRowRe = /(?:([A-Za-z_][A-Za-z0-9_.]*)\s*)?\[@\[([^\]]+)\]\]/g;
  let m;
  while ((m = currentRowRe.exec(formula)) !== null) {
    candidates.push({
      raw: m[0],
      tableName: m[1] || null,
      columnName: m[2],
      currentRow: true,
      pos: m.index
    });
  }

  // Explicit named-table column references not already consumed above.
  const tableColumnRe = /([A-Za-z_][A-Za-z0-9_.]*)\s*\[([^\[\]#@][^\]]*)\]/g;
  while ((m = tableColumnRe.exec(formula)) !== null) {
    candidates.push({
      raw: m[0],
      tableName: m[1],
      columnName: m[2],
      currentRow: false,
      pos: m.index
    });
  }

  // Explicit #Data form.
  const tableDataRe = /([A-Za-z_][A-Za-z0-9_.]*)\s*\[\[#Data\],\[([^\]]+)\]\]/g;
  while ((m = tableDataRe.exec(formula)) !== null) {
    candidates.push({
      raw: m[0],
      tableName: m[1],
      columnName: m[2],
      currentRow: false,
      pos: m.index
    });
  }

  if (!candidates.length) return [];

  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(originSheet);
    const origin = sheet.getRange(originAddress);
    origin.load(["rowIndex", "columnIndex"]);
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();

    const results = [];

    for (const c of candidates.sort((a,b) => a.pos - b.pos)) {
      try {
        let table = null;

        if (c.tableName) {
          // A named table may be on another sheet, so resolve workbook-wide.
          table = context.workbook.tables.getItemOrNullObject(c.tableName);
          table.load(["name"]);
          await context.sync();
          if (table.isNullObject) continue;
        } else {
          // Find the table containing the formula cell.
          for (const t of tables.items) {
            const tr = t.getRange();
            tr.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
            await context.sync();
            const inside =
              origin.rowIndex >= tr.rowIndex &&
              origin.rowIndex < tr.rowIndex + tr.rowCount &&
              origin.columnIndex >= tr.columnIndex &&
              origin.columnIndex < tr.columnIndex + tr.columnCount;
            if (inside) {
              table = t;
              break;
            }
          }
          if (!table) continue;
        }

        const column = table.columns.getItemOrNullObject(c.columnName.trim());
        column.load(["name"]);
        await context.sync();
        if (column.isNullObject) continue;

        let target;
        if (c.currentRow) {
          const body = table.getDataBodyRange();
          body.load(["rowIndex", "rowCount"]);
          const colBody = column.getDataBodyRange();
          colBody.load(["address", "rowIndex", "rowCount"]);
          await context.sync();

          const relativeRow = origin.rowIndex - body.rowIndex;
          if (relativeRow < 0 || relativeRow >= body.rowCount) continue;
          target = colBody.getCell(relativeRow, 0);
        } else {
          target = column.getDataBodyRange();
        }

        target.load("address");
        target.worksheet.load("name");
        await context.sync();

        let address = target.address || "";
        if (address.includes("!")) address = address.substring(address.indexOf("!") + 1);

        results.push({
          sheet: target.worksheet.name,
          address,
          _formulaPos: c.pos
        });
      } catch (_) {
        // If a structured reference cannot be resolved, leave existing
        // A1-reference navigation untouched.
      }
    }
    return results;
  });
}

async function startSession(direction) {
  return Excel.run(async context => {
    const cell = context.workbook.getActiveCell();
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    cell.load(["address","formulas"]);
    sheet.load("name");
    await context.sync();

    const formula = cell.formulas[0][0];
    if (typeof formula !== "string" || !formula.startsWith("=")) {
      setStatus("Select a single formula cell first.");
      return false;
    }

    let refs = parseFormula(formula);
    const structuredRefs = await resolveStructuredReferences(
      formula,
      sheet.name,
      active.address
    );

    // Merge structured references with existing A1 references in formula order.
    for (const sr of structuredRefs) refs.push(sr);

    const lowerFormula = formula.toLowerCase();
    for (const r of refs) {
      if (r._formulaPos == null) {
        const qualified = r.sheet ? `${r.sheet}!${r.address}` : r.address;
        let p = lowerFormula.indexOf(qualified.toLowerCase());
        if (p < 0) p = lowerFormula.indexOf(r.address.toLowerCase());
        r._formulaPos = p < 0 ? Number.MAX_SAFE_INTEGER : p;
      }
    }
    refs.sort((a, b) => a._formulaPos - b._formulaPos);

    // Avoid duplicate navigation targets.
    const seenTargets = new Set();
    refs = refs.filter(r => {
      const key = `${(r.sheet || sheet.name).toLowerCase()}!${r.address.toLowerCase()}`;
      if (seenTargets.has(key)) return false;
      seenTargets.add(key);
      return true;
    });
    if (!refs.length) {
      setStatus("No direct A1-style references found.");
      return false;
    }

    state.active = true;
    state.originSheet = sheet.name;
    state.originAddress = cell.address.includes("!") ? cell.address.split("!").slice(1).join("!") : cell.address;
    state.formula = formula;
    state.refs = refs;
    state.pos = direction > 0 ? 1 : refs.length;
    return true;
  });
}

async function ensureExplorerVisible() {
  try {
    if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
      await Office.addin.showAsTaskpane();
    }
  } catch (_) {}
}

async function closeExplorer() {
  clearSession();
  render();
  try {
    if (Office.addin && typeof Office.addin.hide === "function") {
      await Office.addin.hide();
    }
  } catch (_) {}
}

async function navigate(direction) {
  try {
    if (state.active) {
      const sameSession = await selectionStillBelongsToSession();
      if (!sameSession) clearSession();
    }

    if (!state.active) {
      const ok = await startSession(direction);
      if (!ok) return;
    } else {
      state.pos += direction;
      if (state.pos > state.refs.length) state.pos = 0;
      if (state.pos < 0) state.pos = state.refs.length;
    }

    await Excel.run(async context => {
      let sheet, range;

      if (state.pos === 0) {
        sheet = context.workbook.worksheets.getItem(state.originSheet);
        range = sheet.getRange(state.originAddress);
      } else {
        const r = state.refs[state.pos - 1];
        sheet = context.workbook.worksheets.getItem(r.sheet || state.originSheet);
        range = sheet.getRange(r.address);
      }

      sheet.activate();
      range.select();
      await context.sync();
    });

    render();
  } catch (e) {
    clearSession();
    setStatus("Navigation error: " + (e.message || e));
  }
}

function releaseWebFocus() {
  // There is no supported Office.js API that explicitly focuses Excel's grid,
  // but blurring the webview prevents controls in the pane from retaining focus.
  try {
    const el = document.activeElement;
    if (el && typeof el.blur === "function") el.blur();
    if (typeof window.blur === "function") window.blur();
  } catch (_) {}
}

async function moveSelection(rowDelta, colDelta) {
  return Excel.run(async context => {
    const cell = context.workbook.getActiveCell();
    cell.load(["rowIndex", "columnIndex"]);
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    await context.sync();
    const r = Math.max(0, cell.rowIndex + rowDelta);
    const c = Math.max(0, cell.columnIndex + colDelta);
    sheet.getCell(r, c).select();
    await context.sync();
  });
}

function render() {
  const box = document.getElementById("formula");
  const position = document.getElementById("position");

  if (!state.active) return;

  if (state.pos === 0) {
    box.innerHTML = esc(state.formula);
    position.textContent = "ORIGINAL";
    setStatus(state.originSheet + "!" + state.originAddress);
    return;
  }

  const r = state.refs[state.pos - 1];
  const before = state.formula.slice(0, r.start);
  const active = state.formula.slice(r.start, r.start + r.length);
  const after = state.formula.slice(r.start + r.length);

  box.innerHTML = esc(before) + '<span class="ref active">' + esc(active) + '</span>' + esc(after);
  position.textContent = state.pos + "/" + state.refs.length;
  setStatus((r.sheet || state.originSheet) + "!" + r.address);
}

function setStatus(s) {
  document.getElementById("status").textContent = s;
}
