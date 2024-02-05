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
 * ### Sample script
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
 * @param {Object} gridrange Gridrange of range.
 * @param {String} sheetName Sheet name of the range.
 * @return {String} A1Notation.
 */
function convGridRangeToA1Notation(gridrange, sheetName) {
  if (gridrange === null || sheetName === null || typeof sheetName != "string") {
    throw new Error("Please give gridRange (JSON object) and sheet name (String).");
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
  const a1Notation = st == en ? `'${sheetName}'!${st}` : `'${sheetName}'!${st}:${en}`;
  return a1Notation;
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
