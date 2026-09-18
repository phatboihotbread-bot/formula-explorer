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

  Office.actions.associate("FormulaExplorer.GoLeft", () => navigate(-1));
  Office.actions.associate("FormulaExplorer.GoRight", () => navigate(1));
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

    const refs = parseFormula(formula);
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

async function navigate(direction) {
  try {
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
    setStatus("Navigation error: " + (e.message || e));
  }
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
