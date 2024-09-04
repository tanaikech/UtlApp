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
 * const res1 = UtlApp.columnIndexToLetter(res1);
 * console.log(res1); // Z
 * const res2 = UtlApp.columnIndexToLetter(res2);
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
  return (url == "" ? "" : `${url}?`) + Object.entries(obj).flatMap(([k, v]) => Array.isArray(v) ? v.map(e => `${k}=${encodeURIComponent(e)}`) : `${k}=${encodeURIComponent(v)}`).join("&");
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
 * const res = UtlApp.consolidateA1Notations(a1Notations);
 * console.log(JSON.stringify(res));
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * ["C1:C5","K4:K7","I5:I8","D7:D10","I1:I3","D2:F2","H2:H4","J7:J9","D3:D4","E4:E5","F5:F6","B2","F3","J5","E7","K9"]
 * ```
 * 
 * @param {Array} a1Notations Scattered A1Notations.
 * @return {Array} Array including the consolidated A1Notations.
 */
function consolidateA1Notations(array) {
  return new ConsolidateA1Notations(array).run();
}

/**
 * Consolidate Scattered A1Notations into Continuous Ranges on Google Spreadsheet.
 */
class ConsolidateA1Notations {
  /**
   *
   * @param {Array} a1Notations Scattered A1Notations.
   */
  constructor(a1Notations) {
    this.a1Notations = a1Notations;
  }

  /**
   * ### Description
   * Main method.
   *
   * @returns {Array} Array including consolidated A1Notations.
   */
  run() {
    return this.getConsolidateA1Notations_(this.a1Notations);
  }

  /**
   * ### Description
   * Processing to consolidate A1Notations.
   *
   * @param {Array} a1Notations Array including the inputted A1Notations.
   * @returns {Array} Array including consolidated A1Notations.
   */
  getConsolidateA1Notations_(a1Notations) {
    const expandedA1Notations = expandA1Notations(a1Notations).flat();
    const orgGridRanges = expandedA1Notations.map(e => convA1NotationToGridRange(e, 0));
    const gridRanges = orgGridRanges
      .sort((a, b) => a.startColumnIndex < b.startColumnIndex ? 1 : -1)
      .sort((a, b) => a.startRowIndex > b.startRowIndex ? 1 : -1);
    const maxRow = Math.max(...gridRanges.map(({ endRowIndex }) => endRowIndex));
    const maxCol = Math.max(...gridRanges.map(({ endColumnIndex }) => endColumnIndex));
    const obj = gridRanges.reduce((o, e) => {
      const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = e;
      const key = `sr@${startRowIndex}_er@${endRowIndex}_sc@${startColumnIndex}_ec@${endColumnIndex}`;
      o[key] = e;
      return o;
    }, {});
    let copiedGridranges = gridRanges.slice();
    let copiedObj = { ...obj };
    const res = [];
    do {
      const resObj = this.getMaxAreas_({ obj: copiedObj, gridRanges: copiedGridranges, maxRow, maxCol });
      res.push(resObj);
      if (resObj.length > 0) {
        const removeKeys = resObj.flatMap(({ removeKeys }) => removeKeys.flatMap(({ remove }) => remove));
        copiedObj = Object.fromEntries(Object.entries(copiedObj).filter(([k]) => !removeKeys.includes(k)));
        copiedGridranges = Object.entries(copiedObj).map(([, v]) => v);
      }
    } while (Object.keys(copiedObj).length > 0);
    const result = res.reduce((ar, e) => {
      if (e.length > 0) {
        e.forEach(f => {
          f.removeKeys.forEach(g => {
            if (g.topLeft == g.bottomRight) {
              ar.push(g.topLeft)
            } else {
              ar.push(`${g.topLeft}:${g.bottomRight}`);
            }
          });
        });
      }
      return ar;
    }, []);
    return result;
  }

  /**
   * ### Description
   * Calculated the maximum rectangles from the inputted A1Notations.
   *
   * @param {Object} object GridRange converted from the inputted A1Notations, the numbers of maximum row and column. 
   * @returns {Object} Object including the calculated maximum rectangles.
   */
  getMaxAreas_(object) {
    let { obj, gridRanges, maxRow, maxCol } = object;
    const areas = gridRanges.map(o => {
      const { startRowIndex, startColumnIndex } = o;
      const cc = [];
      const rr = [];
      for (let r = startRowIndex; r < maxRow; r++) {
        for (let c = startColumnIndex; c < maxCol; c++) {
          const key = `sr@${r}_er@${r + 1}_sc@${c}_ec@${c + 1}`;
          if (!obj[key]) {
            cc.push(c - startColumnIndex);
            break;
          } else if (c == maxCol - 1) {
            cc.push(c - startColumnIndex + 1);
            break;
          }
        }
      }
      for (let c = startColumnIndex; c < maxCol; c++) {
        for (let r = startRowIndex; r < maxRow; r++) {
          const key = `sr@${r}_er@${r + 1}_sc@${c}_ec@${c + 1}`;
          if (!obj[key]) {
            rr.push(r - startRowIndex);
            break;
          } else if (r == maxRow - 1) {
            rr.push(r - startRowIndex + 1);
            break;
          }
        }
      }
      const r1 = rr.reduce((oo, e) => {
        if (e < oo.temp) {
          oo.temp = e;
          oo.v.push(oo.temp);
        } else {
          oo.v.push(oo.temp);
        }
        return oo;
      }, { temp: rr[0], v: [] });
      const c1 = cc.reduce((oo, e) => {
        if (e < oo.temp) {
          oo.temp = e;
          oo.v.push(oo.temp);
        } else {
          oo.v.push(oo.temp);
        }
        return oo;
      }, { temp: cc[0], v: [] });
      const r0 = r1.v.indexOf(0);
      const c0 = c1.v.indexOf(0);
      if (r0 > -1) {
        r1.v.splice(r0);
      }
      if (c0 > -1) {
        c1.v.splice(c0);
      }
      const areas = [];
      let c1v = c1.v.slice();
      for (let x = 1; x <= r1.v[0]; x++) {
        const y = c1v.shift();
        const key = `sr@${o.startRowIndex + x - 1}_er@${o.startRowIndex + x}_sc@${o.startColumnIndex + y - 1}_ec@${o.startColumnIndex + y}`;
        obj = Object.fromEntries(Object.entries(obj).filter(([k]) => k != key));
        areas.push({ r: x, c: y, area: x * y });
      }
      const maxArea = Math.max(...areas.map(e => e.area));
      const tempMaxAreaObj = areas.filter(e => e.area == maxArea);
      if (tempMaxAreaObj.length == 0) {
        return [{ areas: [], maxAreaObj: [], o, removeKeys: [{ remove: [], topLeft: "", bottomRight: "" }] }];
      }
      const maxAreaObj = [tempMaxAreaObj[0]];
      const removeKeys = maxAreaObj.map(e => {
        const temp = [];
        for (let i = o.startRowIndex; i < o.startRowIndex + e.r; i++) {
          for (let j = o.startColumnIndex; j < o.startColumnIndex + e.c; j++) {
            const key = `sr@${i}_er@${i + 1}_sc@${j}_ec@${j + 1}`;
            temp.push(key);
          }
        }
        return {
          remove: temp,
          topLeft: `${columnIndexToLetter(o.startColumnIndex)}${o.startRowIndex + 1}`,
          bottomRight: `${columnIndexToLetter(o.startColumnIndex + e.c - 1)}${o.startRowIndex + e.r}`
        };
      });
      return { areas, maxAreaObj, o, removeKeys };
    });
    const a = Math.max(...areas.map(({ maxAreaObj }) => (maxAreaObj && maxAreaObj.length > 0 && maxAreaObj[0]?.area) ? maxAreaObj[0].area : -1));
    const result = areas.filter(({ maxAreaObj }) => maxAreaObj && maxAreaObj.length > 0 && maxAreaObj[0]?.area && maxAreaObj[0]?.area == a);
    return result;
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
