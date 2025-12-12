/**
 * [Tanaike](https://tanaikech.github.io/)
 * [GitHub](https://github.com/tanaikech/UtlApp)
*/

/**
 * @namespace ConvText
 * 
 * ### Description
 * Converting text as unicode.
 * Ref: https://tanaikech.github.io/2020/11/13/converting-texts-to-bold-italic-and-bold-italic-types-of-unicode-using-google-apps-script/
 * 
 * ### Sample script
 * ```
 * const text = "Sample value 123\nSample value 456\nSample value 789";
 * console.log(text); // Original text
 * console.log(UtlApp.ConvText.bold(text)); // Bold type
 * console.log(UtlApp.ConvText.italic(text)); // Italic type
 * console.log(UtlApp.ConvText.boldItalic(text)); // Bold-italic type
 * console.log(UtlApp.ConvText.underLine(text)); // Underline
 * console.log(UtlApp.ConvText.strikethrough(text)); // Strikethrough
 * ```
 * 
 */
var ConvText = {
  /**
   * Method for processing
   * @param {String} text
   * @param {Object} obj
   * @return {String} Converted text.
   */
  c_: function (text, obj) {
    return text.replace(
      new RegExp(`[${obj.reduce((s, { r }) => (s += r), "")}]`, "g"),
      (e) => {
        const t = e.codePointAt(0);
        if ((t >= 48 && t <= 57) || (t >= 65 && t <= 90) || (t >= 97 && t <= 122)) {
          return obj.reduce((s, { r, d }) => {
            if (new RegExp(`[${r}]`).test(e)) s = String.fromCodePoint(e.codePointAt(0) + d);
            return s;
          }, "");
        }
        return e;
      }
    );
  },

  /**
   * Convert text to bold type.
   * @param {String} text
   * @return {String} Converted text.
   */
  bold: function (text) {
    return this.c_(text, [{ r: "0-9", d: 120734 }, { r: "A-Z", d: 120211 }, { r: "a-z", d: 120205 }]);
  },

  /**
   * Convert text to italic type.
   * Ref: https://tanaikech.github.io/2020/11/13/converting-texts-to-bold-italic-and-bold-italic-types-of-unicode-using-google-apps-script/
   * @param {String} text
   * @return {String} Converted text.
   */
  italic: function (text) {
    return this.c_(text, [{ r: "A-Z", d: 120263 }, { r: "a-z", d: 120257 }]);
  },

  /**
   * Convert text to bold and italic type.
   * @param {String} text
   * @return {String} Converted text.
   */
  boldItalic: function (text) {
    return this.c_(text, [{ r: "A-Z", d: 120315 }, { r: "a-z", d: 120309 }]);
  },

  /**
   * Convert text to underline type.
   * @param {String} text
   * @return {String} Converted text.
   */
  underLine: function (text) {
    return text.length > 0 ? [...text].join("\u0332") + "\u0332" : "";
  },

  /**
   * Convert text to strikethrough type.
   * @param {String} text
   * @return {String} Converted text.
   */
  strikethrough: function (text) {
    return text.length > 0 ? [...text].join("\u0336") + "\u0336" : "";
  },
};

/**
 * ### Description
 * Converting colum letter to column index. Start of column index is 0.
 * Ref: https://tanaikech.github.io/2022/05/01/increasing-column-letter-by-one-using-google-apps-script/
 * 
 * ### Sample script
 * ```
 * const res1 = UtlApp.columnLetterToIndex("Z");
 * console.log(res1); // 25
 * const res2 = UtlApp.columnLetterToIndex("AA");
 * console.log(res2); // 26
 * ```
 * 
 * @param {String} letter Column letter.
 * @return {Number} Column index.
 */
function columnLetterToIndex(letter = null) {
  if (letter === null || typeof letter != "string") {
    throw new Error("Please give the column letter as a string.");
  }
  letter = letter.toUpperCase();
  return [...letter].reduce((c, e, i, a) => (c += (e.charCodeAt(0) - 64) * Math.pow(26, a.length - i - 1)), -1);
}

/**
 * ### Description
 * Converting colum index to column letter. Start of column index is 0.
 * Ref: https://stackoverflow.com/a/53678158/7108653
 * 
 * ### Sample script
 * ```
 * const res1 = UtlApp.columnIndexToLetter(25);
 * console.log(res1); // Z
 * const res2 = UtlApp.columnIndexToLetter(26);
 * console.log(res2); // AA
 * ```
 * 
 * @param {Number} index Column index.
 * @return {String} Column letter.
 */
function columnIndexToLetter(index = null) {
  if (index === null || isNaN(index)) {
    throw new Error("Please give the column indexr as a number. In this case, 1st number is 0.");
  }
  let a;
  return (a = Math.floor(index / 26)) >= 0 ? columnIndexToLetter(a - 1) + String.fromCharCode(65 + (index % 26)) : "";
}

/**
 * ### Description
 * Converting a1Notation to gridrange. This will be useful for using Sheets API.
 * Ref: https://tanaikech.github.io/2017/07/31/converting-a1notation-to-gridrange-for-google-sheets-api/
 * 
 * ### Sample script
 * ```
 * const a1Notation = "AB25:AD51";
 * const sheetId = 0;
 * const res = UtlApp.convA1NotationToGridRange(a1Notation, sheetId);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * {
 *   sheetId: 0,
 *   startRowIndex: 24,
 *   endRowIndex: 51,
 *   startColumnIndex: 27,
 *   endColumnIndex: 30
 * }
 * ```
 * 
 * @param {String} a1Notation A1Notation of range.
 * @param {Number} sheetId Sheet ID of the range.
 * @return {Object} Gridrange.
 */
function convA1NotationToGridRange(a1Notation = null, sheetId = null) {
  if (a1Notation === null || sheetId === null || typeof a1Notation != "string" || isNaN(sheetId)) {
    throw new Error("Please give a1Notation (string) and sheet ID (integer).");
  }
  const { col, row } = a1Notation.toUpperCase().split("!").map(f => f.split(":")).pop().reduce((o, g) => {
    var [r1, r2] = ["[A-Z]+", "[0-9]+"].map(h => g.match(new RegExp(h)));
    o.col.push(r1 && columnLetterToIndex(r1[0]))
    o.row.push(r2 && Number(r2[0]))
    return o;
  }, { col: [], row: [] });
  col.sort((a, b) => a > b ? 1 : -1);
  row.sort((a, b) => a > b ? 1 : -1);
  const [start, end] = col.map((e, i) => ({ col: e, row: row[i] }));
  const obj = {
    sheetId,
    startRowIndex: start?.row && start.row - 1,
    endRowIndex: end?.row ? end.row : start.row,
    startColumnIndex: start && start.col,
    endColumnIndex: end ? end.col + 1 : start.col + 1,
  };
  if (obj.startRowIndex === null) {
    obj.startRowIndex = 0;
    delete obj.endRowIndex;
  }
  if (obj.startColumnIndex === null) {
    obj.startColumnIndex = 0;
    delete obj.endColumnIndex;
  }
  return obj;
}

/**
 * ### Description
 * Converting gridrange to a1Notation. This will be useful for using Sheets API.
 * Ref: https://tanaikech.github.io/2017/07/31/converting-a1notation-to-gridrange-for-google-sheets-api/
 * 
 * ### Sample script 1
 * ```
 * const gridRange = {
 *   sheetId: 0,
 *   startRowIndex: 24,
 *   endRowIndex: 51,
 *   startColumnIndex: 27,
 *   endColumnIndex: 30
 * };
 * const sheetName = "Sheet1";
 * const res = UtlApp.convGridRangeToA1Notation(gridRange, sheetName);
 * console.log(res); // 'Sheet1'!AB25:AD51
 * ```
 * 
 * ### Sample script 2
 * ```
 * const gridRange = {
 *   sheetId: 0,
 *   startRowIndex: 24,
 *   endRowIndex: 51,
 *   startColumnIndex: 27,
 *   endColumnIndex: 30
 * };
 * const res = UtlApp.convGridRangeToA1Notation(gridRange);
 * console.log(res); // AB25:AD51
 * ```
 * 
 * @param {Object} gridrange Gridrange of range.
 * @param {String} sheetName Sheet name of the range.
 * @return {String} A1Notation.
 */
function convGridRangeToA1Notation(gridrange, sheetName = "") {
  if (gridrange === null) {
    throw new Error("Please give gridRange (JSON object).");
  }
  const start = {};
  const end = {};
  if (gridrange.hasOwnProperty("startColumnIndex")) {
    start.col = columnIndexToLetter(gridrange.startColumnIndex);
  } else if (
    !gridrange.hasOwnProperty("startColumnIndex") &&
    gridrange.hasOwnProperty("endColumnIndex")
  ) {
    start.col = "A";
  }
  if (gridrange.hasOwnProperty("startRowIndex")) {
    start.row = gridrange.startRowIndex + 1;
  } else if (
    !gridrange.hasOwnProperty("startRowIndex") &&
    gridrange.hasOwnProperty("endRowIndex")
  ) {
    start.row = gridrange.endRowIndex;
  }
  if (gridrange.hasOwnProperty("endColumnIndex")) {
    end.col = columnIndexToLetter(gridrange.endColumnIndex - 1);
  } else if (!gridrange.hasOwnProperty("endColumnIndex")) {
    end.col = "{Here, please set the max column letter.}";
  }
  if (gridrange.hasOwnProperty("endRowIndex")) {
    end.row = gridrange.endRowIndex;
  }
  const k = ["col", "row"];
  const st = k.map((e) => start[e]).join("");
  const en = k.map((e) => end[e]).join("");
  if (sheetName) {
    return st == en ? `'${sheetName}'!${st}` : `'${sheetName}'!${st}:${en}`;
  }
  return st == en ? `${st}` : `${st}:${en}`;
}

/**
 * ### Description
 * This method is used for adding the query parameters to the URL.
 * Ref: https://tanaikech.github.io/2018/07/12/adding-query-parameters-to-url-using-google-apps-script/
 * 
 * ### Sample script 1
 * ```
 * const url = "https://sampleUrl";
 * const query = {
 *   query1: ["value1A", "value1B", "value1C"],
 *   query2: "value2A, value2B",
 *   query3: "value3A/value3B",
 * };
 * const endpoint = UtlApp.addQueryParameters(url, query);
 * console.log(endpoint); // https://sampleUrl?query1=value1A&query1=value1B&query1=value1C&query2=value2A%2C%20value2B&query3=value3A%2Fvalue3B
 * ```
 * 
 * ### Sample script 2
 * ```
 * const url = "";
 * const query = {
 *   query1: ["value1A", "value1B", "value1C"],
 *   query2: "value2A, value2B",
 *   query3: "value3A/value3B",
 * };
 * const endpoint = UtlApp.addQueryParameters(url, query);
 * console.log(endpoint); // query1=value1A&query1=value1B&query1=value1C&query2=value2A%2C%20value2B&query3=value3A%2Fvalue3B
 * ```
 * 
 * @param {String} url The base URL for adding the query parameters.
 * @param {Object} obj JSON object including query parameters.
 * @return {String} URL including the query parameters.
 */
function addQueryParameters(url, obj) {
  if (url === null || obj === null || typeof url != "string") {
    throw new Error("Please give URL (String) and query parameter (JSON object).");
  }
  const o = Object.entries(obj);
  return (url == "" ? "" : `${url}${o.length > 0 ? "?" : ""}`) + o.flatMap(([k, v]) => Array.isArray(v) ? v.map(e => `${k}=${encodeURIComponent(e)}`) : `${k}=${encodeURIComponent(v)}`).join("&");
}

/**
 * ### Description
 * This method is used for parsing the URL including the query parameters.
 * Ref: https://tanaikech.github.io/2018/07/12/adding-query-parameters-to-url-using-google-apps-script/
 * 
 * ### Sample script
 * ```
 * const url = "https://sampleUrl.com/sample?key1=value1&key2=value2&key1=value3&key3=value4&key2=value5";
 * const res = UtlApp.parseQueryParameters(url);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * {
 *   url: 'https://sampleUrl.com/sample',
 *   queryParameters:
 *   {
 *     key1: ['value1', 'value3'],
 *     key2: ['value2', 'value5'],
 *     key3: ['value4']
 *   }
 * }
 * ```
 * 
 * @param {String} url The URL including the query parameters.
 * @return {Object} JSON object including the base url and the query parameters.
 */
function parseQueryParameters(url) {
  if (url === null || typeof url != "string") {
    throw new Error("Please give URL (String) including the query parameters.");
  }
  const s = url.split("?");
  if (s.length == 1) {
    return { url: s[0], queryParameters: null };
  }
  const [baseUrl, query] = s;
  if (query) {
    const queryParameters = query.split("&").reduce(function (o, e) {
      const temp = e.split("=");
      const key = temp[0].trim();
      let value = temp[1].trim();
      value = isNaN(value) ? value : Number(value);
      if (o[key]) {
        o[key].push(value);
      } else {
        o[key] = [value];
      }
      return o;
    }, {});
    return { url: baseUrl, queryParameters };
  }
  return null;
}

/**
 * ### Description
 * This method is used for expanding A1Notations.
 * Ref: https://tanaikech.github.io/2020/04/04/updated-expanding-a1notations-using-google-apps-script/
 * 
 * ### Sample script
 * ```
 * const a1Notations = ["A1:E3", "B10:W13", "EZ5:FA8", "AAA1:AAB3"];
 * const res = UtlApp.expandA1Notations(a1Notations);
 * console.log(JSON.stringify(res));
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * [
 *   ["A1","B1","C1","D1","E1","A2","B2","C2","D2","E2","A3","B3","C3","D3","E3"],
 *   ["B10","C10","D10","E10","F10","G10","H10","I10","J10","K10","L10","M10","N10","O10","P10","Q10","R10","S10","T10","U10","V10","W10","B11","C11","D11","E11","F11","G11","H11","I11","J11","K11","L11","M11","N11","O11","P11","Q11","R11","S11","T11","U11","V11","W11","B12","C12","D12","E12","F12","G12","H12","I12","J12","K12","L12","M12","N12","O12","P12","Q12","R12","S12","T12","U12","V12","W12","B13","C13","D13","E13","F13","G13","H13","I13","J13","K13","L13","M13","N13","O13","P13","Q13","R13","S13","T13","U13","V13","W13"],
 *   ["EZ5","FA5","EZ6","FA6","EZ7","FA7","EZ8","FA8"],
 *   ["AAA1","AAB1","AAA2","AAB2","AAA3","AAB3"]
 * ]
 * ```
 * 
 * @param {Array} a1Notations Array including A1Notations.
 * @return {Array} Array including the expanded A1Notations.
 */
function expandA1Notations(a1Notations, maxRow = "1000", maxColumn = "Z") {
  if (!Array.isArray(a1Notations) || a1Notations.length == 0) {
    throw new Error("Please give a1Notations (Array).");
  }
  const reg1 = new RegExp("^([A-Z]+)([0-9]+)$");
  const reg2 = new RegExp("^([A-Z]+)$");
  const reg3 = new RegExp("^([0-9]+)$");
  return a1Notations.map(e => {
    const a1 = e.split("!");
    const r = a1.length > 1 ? a1[1] : a1[0];
    const [r1, r2] = r.split(":");
    if (!r2) return [r1];
    let rr;
    if (reg1.test(r1) && reg1.test(r2)) {
      rr = [r1.toUpperCase().match(reg1), r2.toUpperCase().match(reg1)];
    } else if (reg2.test(r1) && reg2.test(r2)) {
      rr = [[null, r1, 1], [null, r2, maxRow]];
    } else if (reg1.test(r1) && reg2.test(r2)) {
      rr = [r1.toUpperCase().match(reg1), [null, r2, maxRow]];
    } else if (reg2.test(r1) && reg1.test(r2)) {
      rr = [[null, r1, maxRow], r2.toUpperCase().match(reg1)];
    } else if (reg3.test(r1) && reg3.test(r2)) {
      rr = Number(r1) > Number(r2) ? [[null, "A", r2], [null, maxColumn, r1]] : [[null, "A", r1], [null, maxColumn, r2]];
    } else if (reg1.test(r1) && reg3.test(r2)) {
      rr = [r1.toUpperCase().match(reg1), [null, maxColumn, r2]];
    } else if (reg3.test(r1) && reg1.test(r2)) {
      let temp = r2.toUpperCase().match(reg1);
      rr = Number(temp[2]) > Number(r1) ? [[null, temp[1], r1], [null, maxColumn, temp[2]]] : [temp, [null, maxColumn, r1]];
    } else {
      throw new Error(`Wrong a1Notation: ${r}`);
    }
    const obj = {
      startRowIndex: Number(rr[0][2]),
      endRowIndex: rr.length == 1 ? Number(rr[0][2]) + 1 : Number(rr[1][2]) + 1,
      startColumnIndex: columnLetterToIndex(rr[0][1]),
      endColumnIndex: rr.length == 1 ? columnLetterToIndex(rr[0][1]) + 1 : columnLetterToIndex(rr[1][1]) + 1
    };
    let temp = [];
    for (let i = obj.startRowIndex; i < obj.endRowIndex; i++) {
      for (let j = obj.startColumnIndex; j < obj.endColumnIndex; j++) {
        temp.push(columnIndexToLetter(j) + i);
      }
    }
    return temp;
  });
}

/**
 * ### Description
 * This method is used for consolidating the scattered A1Notations.
 * 
 * ### Sample script
 * ```
 * const a1Notations = ["C1", "I1", "B2", "C2", "D2", "E2", "F2", "H2", "I2", "C3", "D3", "F3", "H3", "I3", "C4", "D4", "E4", "H4", "K4", "C5", "E5", "F5", "I5", "J5", "K5", "F6", "I6", "K6", "D7", "E7", "I7", "J7", "K7", "D8", "I8", "J8", "D9", "J9", "K9", "D10"];
 * const res = consolidateA1Notations(a1Notations);
 * if (res.length > 0) {
 *   console.log(JSON.stringify(res[0]));
 * }
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * ["B2:F2","K4:K7","C3:D4","I5:I8","I1:I3","H2:H4","D7:D10","E4:E5","F5:F6","J7:J9","C5","J5","F3","E7","K9","C1"]
 * ```
 * 
 * Wrapper function for external calls.
 * @param {Array} a1Notations Scattered A1Notations.
 * @return {Array} Array of arrays, where the 0th element is the optimal solution.
 */
function consolidateA1Notations(a1Notations) {
  return new ConsolidateA1Notations(a1Notations).run();
}

/**
 * Class definition: Logic for consolidating A1Notations (Randomized Greedy Approach).
 */
class ConsolidateA1Notations {
  constructor(a1Notations) {
    this.a1Notations = a1Notations;
    this.solutions = []; // Store valid solutions
  }

  run() {
    // 1. Convert input to a Set of coordinates "row,col"
    const initialCells = this.getUniqueCells_(this.a1Notations);
    
    // 2. Execute Search
    // First, create a baseline using a deterministic method (Greedy)
    this.performIteration_(initialCells, false);

    // Next, perform random search until the time limit is reached to find better solutions
    const startTime = Date.now();
    const timeLimit = 1000; // Search for 1 second (Adjust based on data size)

    while (Date.now() - startTime < timeLimit) {
      this.performIteration_(initialCells, true);
    }

    // 3. Organize and Sort Results
    // Remove duplicates (compare via JSON string)
    const uniqueSolutions = Array.from(new Set(this.solutions.map(JSON.stringify))).map(JSON.parse);

    // Sort criteria:
    // 1. Fewest number of ranges (Ascending)
    // 2. Largest sum of squared areas (Descending) - prefers larger chunks
    const sortedSolutions = uniqueSolutions.sort((a, b) => {
      // Compare by length
      if (a.length !== b.length) {
        return a.length - b.length;
      }
      
      // If lengths are equal, calculate score based on area size
      const scoreA = this.calculateScore_(a);
      const scoreB = this.calculateScore_(b);
      return scoreB - scoreA;
    });

    return sortedSolutions;
  }

  /**
   * Performs a single search iteration.
   * @param {Set} allCells All cells to be processed.
   * @param {boolean} isRandom Whether to use random selection.
   */
  performIteration_(allCells, isRandom) {
    let currentCells = new Set(allCells);
    let ranges = [];

    while (currentCells.size > 0) {
      // Find all maximal rectangles creatable from current cells
      let candidates = this.findMaximalRectangles_(currentCells);
      
      if (candidates.length === 0) break; // Should not happen

      // Sort by area (descending)
      candidates.sort((a, b) => b.area - a.area);

      let selectedRect;
      if (!isRandom) {
        // Greedy method: Always choose the largest one
        selectedRect = candidates[0];
      } else {
        // Random method: Randomly select from the top 3 largest
        // (If fewer than 3, select from available candidates)
        const limit = Math.min(candidates.length, 3);
        const index = Math.floor(Math.random() * limit);
        selectedRect = candidates[index];
      }

      // Add selected rectangle to results
      ranges.push(this.gridRangeToA1_(selectedRect));

      // Remove cells included in the selected rectangle
      for (let r = selectedRect.startRow; r <= selectedRect.endRow; r++) {
        for (let c = selectedRect.startCol; c <= selectedRect.endCol; c++) {
          currentCells.delete(`${r},${c}`);
        }
      }
    }

    // Save the completed solution
    this.solutions.push(ranges);
  }

  /**
   * Calculate score (Sum of areas squared).
   * Solutions with larger chunks are rated higher.
   */
  calculateScore_(a1List) {
    let score = 0;
    a1List.forEach(a1 => {
        const range = this.a1ToGridRange_(a1);
        const area = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
        score += area * area;
    });
    return score;
  }

  /**
   * Converts A1Notation string ("C1:C5") to GridRange object (for score calculation).
   */
  a1ToGridRange_(a1) {
    if (a1.includes(':')) {
        const [start, end] = a1.split(':');
        const s = this.a1ToGrid_(start);
        const e = this.a1ToGrid_(end);
        return { startRow: s.row, startCol: s.col, endRow: e.row, endCol: e.col };
    } else {
        const s = this.a1ToGrid_(a1);
        return { startRow: s.row, startCol: s.col, endRow: s.row, endCol: s.col };
    }
  }

  /**
   * Finds all "Maximal Rectangles" from the remaining cells.
   */
  findMaximalRectangles_(cellsSet) {
    // Convert Set to coordinates array
    const coords = [];
    for (const s of cellsSet) {
        const [r, c] = s.split(',').map(Number);
        coords.push({r, c});
    }

    const results = [];
    const seenSignatures = new Set(); 

    // Try every cell as the "Top-Left" of a rectangle
    for (const { r: r0, c: c0 } of coords) {
      // 1. Measure max continuous width from this point
      let maxWidth = 0;
      while (cellsSet.has(`${r0},${c0 + maxWidth}`)) maxWidth++;

      // 2. For each width from 1 to maxWidth, measure max height
      for (let w = 1; w <= maxWidth; w++) {
        let h = 1;
        let valid = true;
        while (valid) {
          // Check next row
          for (let k = 0; k < w; k++) {
            if (!cellsSet.has(`${r0 + h},${c0 + k}`)) {
              valid = false;
              break;
            }
          }
          if (valid) h++;
        }
        
        // Rectangle (r0, c0) Width w, Height h
        const area = w * h;
        // Key for duplicate check
        const signature = `${r0},${c0},${w},${h}`;
        
        if (!seenSignatures.has(signature)) {
            seenSignatures.add(signature);
            results.push({
                startRow: r0,
                startCol: c0,
                endRow: r0 + h - 1,
                endCol: c0 + w - 1,
                area: area
            });
        }
      }
    }
    
    // Remove small rectangles completely contained within other candidates (Efficiency optimization)
    return results.filter((r1, i, arr) => {
        return !arr.some((r2, j) => 
            i !== j && 
            r1.startRow >= r2.startRow && r1.endRow <= r2.endRow &&
            r1.startCol >= r2.startCol && r1.endCol <= r2.endCol
        );
    });
  }

  // --- Helper Methods ---

  getUniqueCells_(a1Notations) {
    const expanded = this.expandA1Notations_(a1Notations).flat();
    const cells = new Set();
    expanded.forEach(a1 => {
      const { row, col } = this.a1ToGrid_(a1);
      cells.add(`${row},${col}`);
    });
    return cells;
  }

  expandA1Notations_(a1Notations) {
    return a1Notations.map(a1 => {
       if (a1.includes(':')) {
           const [start, end] = a1.split(':');
           const s = this.a1ToGrid_(start);
           const e = this.a1ToGrid_(end);
           const res = [];
           for(let r=s.row; r<=e.row; r++) {
               for(let c=s.col; c<=e.col; c++) {
                   res.push(this.gridToA1_(r, c));
               }
           }
           return res;
       }
       return a1;
    });
  }

  a1ToGrid_(a1) {
    const match = a1.match(/([A-Z]+)([0-9]+)/);
    if (!match) return null;
    return {
      col: this.columnLetterToIndex_(match[1]),
      row: parseInt(match[2], 10)
    };
  }

  gridRangeToA1_(range) {
    if (range.startRow === range.endRow && range.startCol === range.endCol) {
      return this.gridToA1_(range.startRow, range.startCol);
    }
    return `${this.gridToA1_(range.startRow, range.startCol)}:${this.gridToA1_(range.endRow, range.endCol)}`;
  }

  gridToA1_(row, col) {
    return `${this.columnIndexToLetter_(col)}${row}`;
  }

  columnLetterToIndex_(letter) {
    let column = 0;
    for (let i = 0; i < letter.length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
    }
    return column;
  }

  columnIndexToLetter_(index) {
    let temp, letter = '';
    while (index > 0) {
      temp = (index - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      index = (index - temp - 1) / 26;
    }
    return letter;
  }
}

/**
 * ### Description
 * This method is used for converting Blob to the data URL.
 * 
 * ### Sample script
 * ```
 * const blob = Utilities.newBlob("sample", MimeType.PLAIN_TEXT);
 * const res = UtlApp.blobToDataUrl(blob);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * data:text/plain;base64,c2FtcGxl
 * ```
 * 
 * @param {Blob} blob Blob
 * @return {String} Data URL converted from Blob.
 */
function blobToDataUrl(blob) {
  if (typeof blob != "object" || blob.toString() != "Blob") {
    throw new Error("Please give Blob as an argument.");
  }
  const mimeType = blob.getContentType();
  if (!mimeType) {
    throw new Error("Given Blob has no mimeType. Please set the mimeType to Blob.");
  }
  return `data:${mimeType};base64,${Utilities.base64Encode(blob.getBytes())}`;
}

/**
 * ### Description
 * This method is used for converting a string of the snake case to the camel case.
 * 
 * ### Sample script
 * ```
 * const res = UtlApp.snake_caseToCamelCase("sample1_sample2_sample3", true);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * Sample1Sample2Sample3
 * ```
 * 
 * @param {String} value String value of the snake case.
 * @param {Boolean} upperCaseForTopCharacter When this is true, the top character is converted to upper case. The default is false.
 * @return {String} String value converted from the snake case to the camel case.
 */
function snake_caseToCamelCase(value, upperCaseForTopCharacter = false) {
  if (!value || typeof value != "string") {
    throw new Error("Please set string value of the snake case.");
  }
  if (upperCaseForTopCharacter) {
    value = value.replace(/^./, ([a]) => a.toUpperCase());
  }
  return value.replace(/_./g, ([, a]) => a.toUpperCase());
}

/**
 * ### Description
 * This method is used for converting a string of the camel case to the snake case.
 * 
 * ### Sample script
 * ```
 * const res = UtlApp.camelCaseTosnake_case("Sample1Sample2Sample3");
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * sample1_sample2_sample3
 * ```
 * 
 * @param {String} value String value of the snake case.
 * @return {String} String value converted from the snake case to the camel case.
 */
function camelCaseTosnake_case(value) {
  if (!value || typeof value != "string") {
    throw new Error("Please set string value of the camel case.");
  }
  return value.replace(/.[A-Z]/g, ([a, b]) => `${a}_${b}`).toLocaleLowerCase();
}

/**
 * ### Description
 * This method is used for creating the form data to HTTP request from an object.
 * 
 * ### Sample script
 * ```
 * const obj = {
 *   key0: "value0",
 *   key1: {
 *     key1a: "value1a",
 *     key1b: "value1b",
 *   },
 *   key2: {
 *     key2a: {
 *       key2aa: "value2aa",
 *       key2ab: "value2ab",
 *     },
 *     key1b: "value1b",
 *   },
 *   key3: ["ar1", "ar2", "ar3"],
 * };
 * const res = UtlApp.createFormDataObject(obj, false);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * {
 *   "key0": "value0",
 *   "key1[key1a]": "value1a",
 *   "key1[key1b]": "value1b",
 *   "key2[key2a][key2aa]": "value2aa",
 *   "key2[key2a][key2ab]": "value2ab",
 *   "key2[key1b]": "value1b",
 *   "key3[0]": "ar1",
 *   "key3[1]": "ar2",
 *   "key3[2]": "ar3"
 * }
 * ```
 * 
 * When the 2nd argument is true, the followine result is obtained.
 * 
 * ```
 * key0=value0&key1[key1a]=value1a&key1[key1b]=value1b&key2[key2a][key2aa]=value2aa&key2[key2a][key2ab]=value2ab&key2[key1b]=value1b&key3[0]=ar1&key3[1]=ar2&key3[2]=ar3
 * ```
 * 
 * @param {Object} object Object for converting to the form data.
 * @param {Boolean} asQueryParameters When this is true, the result is returned as the query parameter. The default is false.
 * @return {Object|String} 
 */
function createFormDataObject(object, asQueryParameters = false) {
  if (!object || typeof object != "object") {
    throw new Error("Please set an object.");
  }

  // ref: https://stackoverflow.com/a/19101235
  Object.flatten = function (data) {
    var result = {};
    function recurse(cur, prop) {
      if (Object(cur) !== cur) {
        result[prop] = cur;
      } else if (Array.isArray(cur)) {
        for (var i = 0, l = cur.length; i < l; i++)
          recurse(cur[i], prop + "[" + i + "]");
        if (l == 0)
          result[prop] = [];
      } else {
        var isEmpty = true;
        for (var p in cur) {
          isEmpty = false;
          recurse(cur[p], prop ? prop + "___" + p : p);
        }
        if (isEmpty && prop)
          result[prop] = {};
      }
    }
    recurse(data, "");
    return result;
  }

  const obj = Object.flatten(object);
  const res = Object.entries(obj).map(([k, v]) => {
    const [t1, ...t2] = k.split("___");
    return [`${t1}${t2.map(e => `[${e}]`).join("")}`, v];
  });
  if (asQueryParameters) {
    return res.map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join("&");
  }
  return Object.fromEntries(res);
}
