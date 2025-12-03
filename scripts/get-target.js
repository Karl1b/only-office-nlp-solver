(function () {
  "use strict";

  let parameterCount = 0;

  // SVG icon for cell picker
  const pickerIconSVG =
    '<svg viewBox="0 0 24 24"><path d="M3 5v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2zm2 0h6v6H5V5zm8 0h6v6h-6V5zm-8 8h6v6H5v-6zm8 0h6v6h-6v-6z"/></svg>';

  // Helper: Convert column letter to number (A=1, B=2, ..., Z=26, AA=27, etc.)
  function colToNum(col) {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.charCodeAt(i) - 64);
    }
    return num;
  }

  // Helper: Convert number to column letter
  function numToCol(num) {
    let col = "";
    while (num > 0) {
      const remainder = (num - 1) % 26;
      col = String.fromCharCode(65 + remainder) + col;
      num = Math.floor((num - 1) / 26);
    }
    return col;
  }

  // Expand a range like "A1:A5" into individual cells ["A1", "A2", "A3", "A4", "A5"]
  function expandRange(rangeStr) {
    const normalized = rangeStr.replace(/\$/g, "").toUpperCase();
    const parts = normalized.split(":");
    if (parts.length !== 2) return [normalized];

    const startMatch = parts[0].match(/([A-Z]+)(\d+)/);
    const endMatch = parts[1].match(/([A-Z]+)(\d+)/);
    if (!startMatch || !endMatch) return [normalized];

    const startCol = colToNum(startMatch[1]);
    const endCol = colToNum(endMatch[1]);
    const startRow = parseInt(startMatch[2]);
    const endRow = parseInt(endMatch[2]);

    const cells = [];
    for (
      let c = Math.min(startCol, endCol);
      c <= Math.max(startCol, endCol);
      c++
    ) {
      for (
        let r = Math.min(startRow, endRow);
        r <= Math.max(startRow, endRow);
        r++
      ) {
        cells.push(numToCol(c) + r);
      }
    }
    return cells;
  }

  window.getIterations = function () {
    const input = document.getElementById("iterationsInput");
    if (!input) return 1000;
    const val = parseInt(input.value, 10);
    return isNaN(val) || val <= 0 ? 1000 : val;
  };

  // Pick target cell from current selection
  window.pickTargetCell = function () {
    window.Asc.plugin.callCommand(
      function () {
        const sheet = Api.GetActiveSheet();
        const selection = sheet.GetSelection();
        if (!selection) return null;
        const addr = selection.GetAddress(true, true, "xlA1", false);
        return addr ? addr.replace(/\$/g, "").split(":")[0] : null;
      },
      false,
      false,
      function (result) {
        if (result) {
          document.getElementById("targetCell").value = result.toUpperCase();
        }
      }
    );
  };

  // Pick parameter cells from current selection (replaces all existing parameters)
  window.pickParameterRange = function () {
    window.Asc.plugin.callCommand(
      function () {
        const sheet = Api.GetActiveSheet();
        const selection = sheet.GetSelection();
        if (!selection) return null;
        const addr = selection.GetAddress(true, true, "xlA1", false);
        return addr ? addr.replace(/\$/g, "") : null;
      },
      false,
      false,
      function (result) {
        if (result) {
          const cells = expandRange(result);
          // Clear all existing parameter rows
          const container = document.getElementById("parametersList");
          container.innerHTML = "";
          parameterCount = 0;
          // Add a row for each cell in the selection
          cells.forEach(function (cell) {
            addParameterRowWithValues(cell, "", "");
          });
        }
      }
    );
  };

  window.addParameterRow = function () {
    addParameterRowWithValues("", "", "");
  };

  function addParameterRowWithValues(cellValue, minValue, maxValue) {
    const container = document.getElementById("parametersList");
    const rowId = `param-${parameterCount++}`;

    const row = document.createElement("div");
    row.className = "parameter-row";
    row.id = rowId;
    row.innerHTML = `
      <div class="param-top-row">
        <span class="param-label">Cell</span>
        <input type="text" class="param-cell" placeholder="B1" value="${cellValue}" />
        <button class="remove-btn" onclick="removeParameterRow('${rowId}')">×</button>
      </div>
<div class="param-bounds-row">
  <div class="limit-check"><input type="checkbox" class="param-limit-min" /><label>StrictMin</label></div>
  <input type="text" class="param-minmax param-min" placeholder="-1" value="${minValue}" />
  <div class="limit-check"><input type="checkbox" class="param-limit-max" /><label>StrictMax</label></div>
  <input type="text" class="param-minmax param-max" placeholder="1" value="${maxValue}" />
</div>
    `;
    container.appendChild(row);
  }

  window.removeParameterRow = function (rowId) {
    const row = document.getElementById(rowId);
    if (row) row.remove();
  };

  window.getParameterCells = function () {
    const rows = document.querySelectorAll(".parameter-row");
    const parameters = [];
    rows.forEach((row) => {
      const cell = row.querySelector(".param-cell").value.trim().toUpperCase();
      const minStr = row.querySelector(".param-min").value.trim();
      const maxStr = row.querySelector(".param-max").value.trim();
      if (cell) {
        parameters.push({
          cell: cell,
          min_value: minStr !== "" ? parseFloat(minStr) : null,
          max_value: maxStr !== "" ? parseFloat(maxStr) : null,
          limit_min: row.querySelector(".param-limit-min").checked,
          limit_max: row.querySelector(".param-limit-max").checked,
        });
      }
    });
    return parameters;
  };

  window.startSolver = function () {
    const addr = document
      .getElementById("targetCell")
      .value.trim()
      .toUpperCase();
    if (!addr) {
      alert("Please set a target cell first.");
      return;
    }

    const parameterCells = getParameterCells();
    if (parameterCells.length === 0) {
      alert("Please add at least one parameter cell.");
      return;
    }

    document.getElementById("startBtn").disabled = true;
    document.getElementById("statusText").textContent = "Starting...";
    document.getElementById("statusText").className = "status-value running";

    Asc.scope.parameterCells = parameterCells;
    Asc.scope.targetCell = addr;

    window.Asc.plugin.callCommand(
      function () {
        function getSimpleFormula(addr) {
          var sheet = Api.GetActiveSheet();
          var rng = sheet.GetRange(addr);
          if (!rng) return null;
          return rng.GetFormula();
        }

        function getCellValue(addr) {
          var sheet = Api.GetActiveSheet();
          var rng = sheet.GetRange(addr);
          if (!rng) return null;
          return rng.GetValue();
        }

        function normalizeRef(ref) {
          return ref.replace(/\$/g, "");
        }

        function expandRange(rangeStr) {
          var parts = rangeStr.replace(/\$/g, "").split(":");
          if (parts.length !== 2) return [rangeStr.replace(/\$/g, "")];
          var start = parts[0],
            end = parts[1];
          var startMatch = start.match(/([A-Z]+)(\d+)/);
          var endMatch = end.match(/([A-Z]+)(\d+)/);
          if (!startMatch || !endMatch) return [rangeStr.replace(/\$/g, "")];

          var startCol = startMatch[1],
            startRow = parseInt(startMatch[2]);
          var endCol = endMatch[1],
            endRow = parseInt(endMatch[2]);

          function colToNum(col) {
            var num = 0;
            for (var i = 0; i < col.length; i++) {
              num = num * 26 + (col.charCodeAt(i) - 64);
            }
            return num;
          }
          function numToCol(num) {
            var col = "";
            while (num > 0) {
              var remainder = (num - 1) % 26;
              col = String.fromCharCode(65 + remainder) + col;
              num = Math.floor((num - 1) / 26);
            }
            return col;
          }

          var startColNum = colToNum(startCol),
            endColNum = colToNum(endCol);
          var cells = [];
          for (var c = startColNum; c <= endColNum; c++) {
            for (var r = startRow; r <= endRow; r++) {
              cells.push(numToCol(c) + r);
            }
          }
          return cells;
        }

        function getFullFormula(cell, parameterCells) {
          var maxIterations = 20;
          var parameterCellSet = new Set();
          if (parameterCells && Array.isArray(parameterCells)) {
            parameterCells.forEach(function (param) {
              if (param.cell) parameterCellSet.add(normalizeRef(param.cell));
            });
          }

          var result = getSimpleFormula(cell);
          if (!result || !result.startsWith("=")) return result;

          var expandedCells = new Set();
          expandedCells.add(normalizeRef(cell));

          for (var iteration = 0; iteration < maxIterations; iteration++) {
            var hasReplacement = false;
            var cellRefPattern = /\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/g;
            var matches = result.match(cellRefPattern);
            if (!matches) break;

            var uniqueMatches = [];
            var seen = new Set();
            for (var i = 0; i < matches.length; i++) {
              var normalized = normalizeRef(matches[i]);
              if (!seen.has(normalized)) {
                seen.add(normalized);
                uniqueMatches.push(matches[i]);
              }
            }
            uniqueMatches.sort(function (a, b) {
              return b.length - a.length;
            });

            for (var j = 0; j < uniqueMatches.length; j++) {
              var match = uniqueMatches[j];
              var normalizedMatch = normalizeRef(match);

              if (match.includes(":")) {
                var cells = expandRange(match);
                var expandedFormulas = [];
                for (var k = 0; k < cells.length; k++) {
                  var cellAddr = cells[k];
                  if (parameterCellSet.has(cellAddr)) {
                    expandedFormulas.push(cellAddr);
                  } else {
                    var cellFormula = getSimpleFormula(cellAddr);
                    if (cellFormula && cellFormula.startsWith("=")) {
                      expandedFormulas.push(
                        "(" + cellFormula.substring(1) + ")"
                      );
                    } else {
                      expandedFormulas.push(getCellValue(cellAddr));
                    }
                  }
                }
                var replacement = "(" + expandedFormulas.join("+") + ")";
                var escapedMatch = match
                  .replace(/\$/g, "\\$?")
                  .replace(/:/g, "\\s*:\\s*");
                var rangePattern = new RegExp(
                  "SUM\\s*\\(\\s*" + escapedMatch + "\\s*\\)",
                  "gi"
                );
                var newResult = result.replace(rangePattern, replacement);
                if (newResult !== result) {
                  result = newResult;
                  hasReplacement = true;
                }
                continue;
              }

              if (parameterCellSet.has(normalizedMatch)) continue;
              if (expandedCells.has(normalizedMatch)) continue;

              var refFormula = getSimpleFormula(normalizedMatch);
              if (refFormula && refFormula.startsWith("=")) {
                expandedCells.add(normalizedMatch);
                var formulaContent = "(" + refFormula.substring(1) + ")";
                var col = normalizedMatch.match(/[A-Z]+/)[0];
                var row = normalizedMatch.match(/\d+/)[0];
                var refPattern = new RegExp(
                  "\\$?" + col + "\\$?" + row + "(?![0-9])",
                  "g"
                );
                var newResult = result.replace(refPattern, formulaContent);
                if (newResult !== result) {
                  result = newResult;
                  hasReplacement = true;
                }
              }
            }
            if (!hasReplacement) break;
          }

          var cellRefPattern = /\$?[A-Z]+\$?\d+/g;
          var matches = result.match(cellRefPattern);
          if (matches) {
            var uniqueMatches = [];
            var seen = new Set();
            for (var i = 0; i < matches.length; i++) {
              var normalized = normalizeRef(matches[i]);
              if (!seen.has(normalized)) {
                seen.add(normalized);
                uniqueMatches.push({
                  original: matches[i],
                  normalized: normalized,
                });
              }
            }
            uniqueMatches.sort(function (a, b) {
              return b.normalized.length - a.normalized.length;
            });

            for (var j = 0; j < uniqueMatches.length; j++) {
              var item = uniqueMatches[j];
              var normalizedRef = item.normalized;
              if (parameterCellSet.has(normalizedRef)) {
                var col = normalizedRef.match(/[A-Z]+/)[0];
                var row = normalizedRef.match(/\d+/)[0];
                var refPattern = new RegExp(
                  "\\$?" + col + "\\$?" + row + "(?![0-9])",
                  "g"
                );
                result = result.replace(refPattern, normalizedRef);
                continue;
              }
              var cellValue = getCellValue(normalizedRef);
              if (cellValue !== null) {
                var col = normalizedRef.match(/[A-Z]+/)[0];
                var row = normalizedRef.match(/\d+/)[0];
                var refPattern = new RegExp(
                  "\\$?" + col + "\\$?" + row + "(?![0-9])",
                  "g"
                );
                result = result.replace(refPattern, cellValue);
              }
            }
          }
          return result;
        }

        return getFullFormula(Asc.scope.targetCell, Asc.scope.parameterCells);
      },
      false,
      false,
      function (res) {
        console.log("[RESULT]", res);
        if (res && window.goParseInput) {
          var parameterCells = getParameterCells();
          var paramsJSON = JSON.stringify(parameterCells);
          try {
            var iterations = getIterations();
            window.goParseInput(res, paramsJSON, iterations.toString());
          } catch (err) {
            console.error("[WASM] Error:", err);
            document.getElementById("statusText").textContent = "Error";
            document.getElementById("statusText").className =
              "status-value error";
            document.getElementById("startBtn").disabled = false;
          }
        } else {
          document.getElementById("statusText").textContent =
            "No formula found";
          document.getElementById("statusText").className =
            "status-value error";
          document.getElementById("startBtn").disabled = false;
        }
      }
    );
  };
})();
