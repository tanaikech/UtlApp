/**
 * [Tanaike](https://tanaikech.github.io/)
 * [GitHub](https://github.com/tanaikech/UtlApp)
*/

/**
 * ### Description
 * When the inputted array is 2 dimensional array, true is returned.
 * 
 * ### Sample script
 * ```
 * const array1 = ["", [1, 2, 3], [1, 2, 3], [1, 2, 3]];
 * const res1 = UtlApp.is2DimensionalArray(array1); // false
 * 
 * const array2 = [[1, 2, 3], [1, 2], [1]];
 * const res2 = UtlApp.is2DimensionalArray(array2); // true
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @return {Boolean} When the inputted array is 2 dimensional array, true is returned.
 */
function is2DimensionalArray(array) {
  return array.every(r => Array.isArray(r));
}

/**
 * ### Description
 * When the inputted 2 dimensional array is the uniformed array, true is returned.
 * 
 * ### Sample script
 * ```
 * const array1 = [[1, 2, 3], [1, 2, 3], [1, 2, 3]];
 * const res1 = UtlApp.isUniform2DArray(array1); // true
 * 
 * const array2 = [[1, 2, 3], [1, 2], [1]];
 * const res2 = UtlApp.isUniform2DArray(array2); // false
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Boolean} When the inputted 2 dimensional array is the uniformed array, true is returned.
 */
function isUniform2DArray(array, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  return new Set(array.map(r => r.length)).size == 1 ? true : false;
}

/**
 * ### Description
 * Make all array in 2 dimensional array uniforming the same length.
 * 
 * ### Sample script
 * ```
 * const array1 = [[1, 2, 3], [1, 2, 3], [1, 2, 3]];
 * const res1 = UtlApp.uniform2DArray(array1);
 * console.log(res1); // [ [ 1, 2, 3 ], [ 1, 2, 3 ], [ 1, 2, 3 ] ]
 * 
 * const array2 = [[1, 2, 3], [1, 2], [1]];
 * const res2 = UtlApp.uniform2DArray(array2);
 * console.log(res2); // [ [ 1, 2, 3 ], [ 1, 2, null ], [ 1, null, null ] ]
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @param {*} empty Value used in the added element. Default is null.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Array} Uniformed array.
 */
function uniform2DArray(array, empty = null, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  const maxLen = Math.max(...array.map(r => r.length));
  return array.map(r => [...r, ...Array(maxLen - r.length).fill(empty)]);
}

/**
 * ### Description
 * Transpose 2 dimensional array.
 * 
 * ### Sample script
 * ```
 * const array1 = [[1, 2, 3], [1, 2, 3], [1, 2, 3]];
 * const res1 = UtlApp.transpose(array1);
 * console.log(res1); // [ [ 1, 1, 1 ], [ 2, 2, 2 ], [ 3, 3, 3 ] ]
 * 
 * const array2 = [[1, 2, 3], [1, 2], [1]];
 * const res2 = UtlApp.transpose(array2);
 * console.log(res2); // [ [ 1, 1, 1 ], [ 2, 2, null ], [ 3, null, null ] ]
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Array} Transposed array.
 */
function transpose(array, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  return array[0].map((_, col) => array.map(row => row[col] || null));
}

/**
 * ### Description
 * Split array every n length.
 * Ref: https://tanaikech.github.io/2022/05/21/splitting-and-processing-an-array-every-n-length-using-google-apps-script/
 *  
 * ### Sample script
 * ```
 * const size = 3;
 * const array1 = ["a1", "b1", "c1", "d1", "e1", "f1", "g1", "h1", "i1", "j1"];
 * const res = UtlApp.splitArray(array1, size);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * [
 *   [ 'a1', 'b1', 'c1' ],
 *   [ 'd1', 'e1', 'f1' ],
 *   [ 'g1', 'h1', 'i1' ],
 *   [ 'j1' ]
 * ]
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Array} Transposed array.
 */
function splitArray(array, size) {
  if (!array || !size || !Array.isArray(array)) {
    throw new Error("Please give an array and split size.");
  }
  return [...Array(Math.ceil(array.length / size))].map(_ => array.splice(0, size));
}

/**
 * ### Description
 * Retrieve the specific columns from 2 dimensional array.
 *  
 * ### Sample script
 * ```
 * const array1 = [[1, 2, 3], [1, 2, 3], [1, 2, 3]];
 * const columns = [2];
 * const res = UtlApp.getSpecificColumns(array1, columns);
 * console.log(res); // [ [ 2 ], [ 2 ], [ 2 ] ]
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @param {Array} columns 1 dimensional array. Give the specific column numbers you want to get. The 1st number is 1.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Array} Retrieved specific columns.
 */
function getSpecificColumns(array, columns, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  if (!columns || !Array.isArray(columns) || columns.length == 0) {
    throw new Error("Please set column numbers you want to retrieve.");
  }
  return array.map(r => columns.map(e => r[e - 1] || null));
}

/**
 * ### Description
 * Delete specific columns from 2 dimensional array.
 *  
 * ### Sample script
 * ```
 * const array1 = [[1, 2, 3], [1, 2, 3], [1, 2, 3]];
 * const columns = [2];
 * const res = UtlApp.deleteSpecificColumns(array1, columns);
 * console.log(res); // [ [ 1, 3 ], [ 1, 3 ], [ 1, 3 ] ]
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @param {Array} columns 1 dimensional array. Give the specific column numbers you want to delete. The 1st number is 1.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Array} Array deleted the specific columns.
 */
function deleteSpecificColumns(array, columns, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  if (!columns || !Array.isArray(columns) || columns.length == 0) {
    throw new Error("Please set column numbers you want to retrieve.");
  }
  return array.map(r => r.filter((_, j) => !columns.includes(j + 1)));
}

/**
 * ### Description
 * Insert columns to 2 dimensional array.
 *  
 * ### Sample script
 * ```
 * const array1 = [[1, 2, 3], [1, 2, 3], [1, 2, 3]];
 * const array2 = [[4, 5, 6], [7, 8, 9]];
 * const column = 2;
 * const res1 = UtlApp.insertColumns(array1, array2, column);
 * console.log(res1); // [[1,4,5,6,2,3],[1,7,8,9,2,3],[1,null,null,null,2,3]]
 * 
 * const res2 = UtlApp.insertColumns(array1, array2, column, true);
 * console.log(res1); // [[1,4,5,6,2,3],[1,7,8,9,2,3],[1,null,null,null,2,3]]
 * ```
 * 
 * @param {Array} array 2 dimensional array.
 * @param {Array} insertArray 2 dimensional array you want to insert.
 * @param {Number} column Column number you want to insert the array. The 1st number is 1. Default is 1.
 * @param {Boolean} rep When you want to replace the column, please use true. Default is false. In this case, the columns are inserted.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Array} Array inserted the columns.
 */
function insertColumns(array, insertArray, column = 1, rep = false, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  if (!insertArray || !Array.isArray(insertArray) || insertArray.length == 0 || !is2DimensionalArray(insertArray)) {
    throw new Error("Please set columns you want to insert.");
  }
  const t = transpose(array, false);
  t.splice(column - 1, rep ? 1 : 0, ...transpose(insertArray, false));
  return transpose(t, false);
}

/**
 * ### Description
 * Remove duplicated values from 1 dimensional array.
 *  
 * ### Sample script
 * ```
 * const array1 = ["a1", "b1", "c1", "b1", "c1"];
 * const res = UtlApp.removeDuplicatedValues(array1);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * {
 * "removeDuplicatedValues":["a1","b1","c1"],
 * "duplicatedValues":["b1","c1"],
 * "numberOfDuplicate":{"a1":1,"b1":2,"c1":2}
 * }
 * ```
 *  
 * @param {Array} array 1 dimensional array.
 * @return {Object} Object including removeDuplicatedValues, duplicatedValues and numberOfDuplicate.
 */
function removeDuplicatedValues(array) {
  if (!Array.isArray(array)) {
    throw new Error("Please use 1 dimensional array.");
  }
  const obj = array.reduce((m, e) => m.set(e, m.has(e) ? m.get(e) + 1 : 1), new Map());
  const e = [...obj.entries()];
  return {
    removeDuplicatedValues: [...obj.keys()],
    duplicatedValues: e.reduce((ar, [k, v]) => {
      if (v != 1) ar.push(k);
      return ar;
    }, []),
    numberOfDuplicate: Object.fromEntries(e)
  }
}

/**
 * ### Description
 * Retrieve empty row index.
 *  
 * ### Sample script
 * ```
 * const array1 = [["a1"], [2], [3], [4], [], [5], [], ["z"], [""], [7], [8]];
 * const res = UtlApp.get1stEmptyRow(array1);
 * console.log(res); // { topEmptyRow: 4, lastEmptyRow: 8 }
 * ```
 *  
 * @param {Array} array 2 dimensional array.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Object} Object including the top empty row index and the last empty row index.
 */
function get1stEmptyRow(array, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  return {
    topEmptyRow: array.findIndex(r => r.join("") == ""),
    lastEmptyRow: array.length - 1 - array.reverse().findIndex(r => r.join("") == "")
  }
}

/**
 * ### Description
 * Retrieve empty column index.
 *  
 * ### Sample script
 * ```
 * const array1 = [["a1", "b1", "f", "d1", "aa", "h", "aa"], [2, 3, 4, 5], [3, 1, 1, , , 3], [4, 1, 1, 1], [1, 2, 3], [5, 1, 1], [1, 1, 1], ["z", 1, 1], [1, "b", 1, 1, ""], [7, 1, 1, 1], [8, 6, 5, 4, 3]];
 * const res = UtlApp.get1stEmptyColumn(array1);
 * console.log(res); // { topEmptyColumn: 3, lastEmptyColumn: 7 }
 * ```
 *  
 * @param {Array} array 2 dimensional array.
 * @param {Boolean} check Check whether the inputted array is 2 dimensional array. Default is true.
 * @return {Object} Object including the top empty column index and the last empty column index.
 */
function get1stEmptyColumn(array, check = true) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  return array.reduce((o, r) => {
    let top = r.findIndex(c => typeof c === "undefined" || c.toString() == "");
    let last = r.length - 1 - r.reverse().findIndex(c => typeof c === "undefined" || c.toString() == "");
    top = top == -1 ? r.length : top;
    last = last == -1 ? 0 : last;
    o.topEmptyColumn = o.topEmptyColumn < top ? o.topEmptyColumn : top;
    o.lastEmptyColumn = o.lastEmptyColumn > last ? o.lastEmptyColumn : last;
    return o;
  }, { topEmptyColumn: Infinity, lastEmptyColumn: 0 });
}

/**
 * ### Description
 * Sum numbers in an array.
 *  
 * ### Sample script
 * ```
 * const array1 = [1, 2, 3, 4, 5];
 * const res = UtlApp.sum(array1);
 * console.log(res); // 15
 * ```
 * 
 * @param {Array} array Input 1 dimensional array including number values.
 * @param {Boolean} check Check whether the values of the inputted array is number. Default is true.
 * @return {String} int8array.
 */
function sum(array, check = true) {
  if (check && !(Array.isArray(array) && array.every(e => !isNaN(e)))) {
    throw new Error("Please give an array including numbers.");
  }
  return array.reduce((n, e) => n += e, 0);
}

/**
 * ### Description
 * Compiling Continuous Numbers using Google Apps Script.
 * Ref: https://tanaikech.github.io/2021/10/08/compiling-continuous-numbers-using-google-apps-script/
 *  
 * ### Sample script
 * ```
 * const ar = [4, 5, 9, 3, 10, 5, 11, 7, 7, 13, 1];
 * const res = UtlApp.compilingNumbers(ar);
 * console.log(res)
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * [
 *   { start: 1, end: 1 },
 *   { start: 3, end: 5 },
 *   { start: 7, end: 7 },
 *   { start: 9, end: 11 },
 *   { start: 13, end: 13 }
 * ]
 * ```
 *  
 * @param {Array} array Input array.
 * @return {Array} Array including object like [{"start":1,"end":1},{"start":3,"end":5},{"start":7,"end":7},{"start":9,"end":11},{"start":13,"end":13}].
 */
function compilingNumbers(array) {
  if (!(Array.isArray(array) && array.every(e => !isNaN(e)))) {
    throw new Error("Please give an array including numbers.");
  }
  const { values } = [...new Set(array.sort((a, b) => a - b))].reduce((o, e, i, a) => {
    if (o.temp.length == 0 || (o.temp.length > 0 && e == o.temp[o.temp.length - 1] + 1)) {
      o.temp.push(e);
    } else {
      if (o.temp.length > 0) {
        o.values.push({ start: o.temp[0], end: o.temp[o.temp.length - 1] });
      }
      o.temp = [e];
    }
    if (i == a.length - 1) {
      o.values.push(o.temp.length > 1 ? { start: o.temp[0], end: o.temp[o.temp.length - 1] } : { start: e, end: e });
    }
    return o;
  }, { temp: [], values: [] });
  return values;
}

/**
 * ### Description
 * Converting 2 dimensional array to JSON object.
 * Ref: https://tanaikech.github.io/2021/10/24/converting-values-of-google-spreadsheet-to-object-using-google-apps-script/
 *  
 * ### Sample script
 * ```
 * const headers = ["header1", "header2", "header3"];
 * const rows = [["a2", "b2", "c2"], ["a3", "b3", "c3"]];
 * const res = UtlApp.convArrayToObject(headers, rows);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * [
 *   { header1: 'a2', header2: 'b2', header3: 'c2' },
 *   { header1: 'a3', header2: 'b3', header3: 'c3' }
 * ]
 * ```
 *  
 * @param {Array} headers Header array (1 dimensional array).
 * @param {Array} rows Row array (2 dimensional array).
 * @return {Object} JSON object.
 */
function convArrayToObject(headers, rows) {
  if (!Array.isArray(headers) || !Array.isArray(rows) || !is2DimensionalArray(rows)) {
    throw new Error("Please give an array of header and values.");
  }
  return rows.map(r => headers.reduce((o, h, j) => (o[h] = r[j], o), {}));
}

/**
 * ### Description
 * Converting 2-dimensional array as unpivot (reverse pivot).
 * Ref: https://tanaikech.github.io/2023/05/11/unpivot-on-google-spreadsheet-using-google-apps-script/
 *
 * ### Sample script
 * ```
 * const values = [
 *   ["", "b1", "c1", "d1"],
 *   ["a2", 1, 2, 3],
 *   ["a3", 4, 5, 6],
 *   ["a4", 7, 8, 9],
 *   ["a5", 10, 11, 12]
 * ];
 * const res = UtlApp.unpivot(values);
 * console.log(res);
 * ```
 *
 * Result is as follows.
 *
 * ```
 * [
 *   ['b1', 'a2', 1],
 *   ['b1', 'a3', 4],
 *   ['b1', 'a4', 7],
 *   ['b1', 'a5', 10],
 *   ['c1', 'a2', 2],
 *   ['c1', 'a3', 5],
 *   ['c1', 'a4', 8],
 *   ['c1', 'a5', 11],
 *   ['d1', 'a2', 3],
 *   ['d1', 'a3', 6],
 *   ['d1', 'a4', 9],
 *   ['d1', 'a5', 12]
 * ]
 * ```
 *
 * @param {Array} values 2 dimensional array.
 * @return {Array} 2 dimensional array converted as unpivot (reverse pivot).
 */
function unpivot(values) {
  if (!Array.isArray(values) || !is2DimensionalArray(values)) {
    throw new Error("Please give an array of values.");
  }
  const [[, ...h], ...v] = values;
  return h.flatMap((hh, i) => v.map(t => [hh, t[0], t[i + 1]]));
}

/**
 * ### Description
 * Reversing 2-dimensional array with unpivot.
 * Ref: https://tanaikech.github.io/2023/05/11/unpivot-on-google-spreadsheet-using-google-apps-script/
 *
 * ### Sample script
 * ```
 * const values = [
 *   ["b1", "a2", 1],
 *   ["b1", "a3", 4],
 *   ["b1", "a4", 7],
 *   ["b1", "a5", 10],
 *   ["c1", "a2", 2],
 *   ["c1", "a3", 5],
 *   ["c1", "a4", 8],
 *   ["c1", "a5", 11],
 *   ["d1", "a2", 3],
 *   ["d1", "a3", 6],
 *   ["d1", "a4", 9],
 *   ["d1", "a5", 12]
 * ];
 * const res = UtlApp.reverseUnpivot(values);
 * console.log(res);
 * ```
 *
 * Result is as follows.
 *
 * ```
 * [
 *   [null, 'b1', 'c1', 'd1'],
 *   ['a2', 1, 2, 3],
 *   ['a3', 4, 5, 6],
 *   ['a4', 7, 8, 9],
 *   ['a5', 10, 11, 12]
 * ]
 * ```
 *
 * @param {Array} values 2 dimensional array.
 * @return {Array} Reversed 2 dimensional array with unpivot.
 */
function reverseUnpivot(values) {
  if (!Array.isArray(values) || !is2DimensionalArray(values)) {
    throw new Error("Please give an array of values.");
  }
  const [a, b, c] = values[0].map((_, c) => values.map(r => r[c]));
  const ch = [...new Set(a)];
  const rh = [...new Set(b)];
  const size = rh.length;
  const temp = [[null, ...rh], ...[...Array(Math.ceil(c.length / size))].map(_ => c.splice(0, size)).map((vv, i) => [ch[i], ...vv])];
  return temp[0].map((_, c) => temp.map(r => r[c]));
}

/**
 * ### Description
 * Calculate dot product from 2 arrays.
 *
 * ### Sample script
 * ```
 * const array1 = [1, 2, 3, 4, 5];
 * const array2 = [3, 4, 5, 6, 7];
 * const res = UtlApp.dotProduct(array1, array2);
 * ```
 *
 * @param {Array} array1 1-dimensional array including numbers.
 * @param {Array} array2 1-dimensional array including numbers.
 * @return {Number} Calculated result of dot product.
 */
function dotProduct(array1, array2) {
  if (!Array.isArray(array1) || !Array.isArray(array2)) {
    throw new Error("Please give 2 arrays.");
  }
  return array1.reduce((t, e, i) => t += e * array2[i], 0);
}

/**
 * ### Description
 * Calculate cosine similarity from 2 arrays.
 *
 * ### Sample script
 * ```
 * const array1 = [1, 2, 3, 4, 5];
 * const array2 = [3, 4, 5, 6, 7];
 * const res = UtlApp.cosineSimilarity(array1, array2);
 * ```
 *
 * @param {Array} array1 1-dimensional array including numbers.
 * @param {Array} array2 1-dimensional array including numbers.
 * @return {Number} Calculated result of cosine similarity.
 */
function cosineSimilarity(array1, array2) {
  if (!Array.isArray(array1) || !Array.isArray(array2)) {
    throw new Error("Please give 2 arrays.");
  }
  const dotProduct = array1.reduce((t, e, i) => t += e * array2[i], 0);
  const magnitudes = [array1, array2].map(e => Math.sqrt(e.reduce((t, f) => t += f * f, 0))).reduce((t, f) => t *= f, 1);
  return dotProduct / magnitudes;
}
