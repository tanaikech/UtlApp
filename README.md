# UtlApp

<a name="top"></a>
[MIT License](LICENCE)

<a name="overview"></a>

# Overview

This is a Google Apps Script library including useful scripts for supporting to development of applications by Google Apps Script. In the current stage, the 3 categories "For array processing", "For binary processing", and "For string processing" are included in this library.

![](images/fig1.png)

<a name="description"></a>

# Description

When I create applications using Google Apps Script, there are useful scripts for often use. At that time, I thought that when those scripts can be simply used, they will be useful not only to me but also to other users. From this motivation, I created a Google Apps Script library including those scripts. But, I have been using these useful scripts only in my development before.

I sometimes answer questions related to Google Apps Script at Stackoverflow. From my experience that I have answered those questions, I also added useful scripts, that I believe, in the library. Some of these scripts will be also the direct answer to the questions. If this was useful for your situation, I'm glad.

# Policy of this library

**The policy of this library is not to use any scopes.**

# Library's project key

```
1idMI9-WtPMbYvbK5D7KH2_GWh62Dny9RG8NzjwjHI5whGIAPXEtTJmeC
```

<a name="usage"></a>

# Usage

## 1. Install library

In order to use this library, please install the library as follows.

1. Create a GAS project.

   - You can use this library for the GAS project of both the standalone type and the container-bound script type.

1. [Install this library](https://developers.google.com/apps-script/guides/libraries).

   - Library's project key is **`1idMI9-WtPMbYvbK5D7KH2_GWh62Dny9RG8NzjwjHI5whGIAPXEtTJmeC`**.

# Methods

## For array processing

| Methods                                           | Description                                                                     |
| :------------------------------------------------ | :------------------------------------------------------------------------------ |
| [is2DimensionalArray](#is2dimensionalarray)       | When the inputted array is 2 dimensional array, true is returned.               |
| [isUniform2DArray](#isuniform2darray)             | When the inputted 2 dimensional array is the uniformed array, true is returned. |
| [uniform2DArray](#uniform2darray)                 | Make all array in 2 dimensional array uniforming the same length.               |
| [transpose](#transpose)                           | Transpose 2 dimensional array.                                                  |
| [splitArray](#splitarray)                         | Split array every n length.                                                     |
| [getSpecificColumns](#getspecificcolumns)         | Retrieve the specific columns from 2 dimensional array.                         |
| [deleteSpecificColumns](#deletespecificcolumns)   | Delete specific columns from 2 dimensional array.                               |
| [insertColumns](#insertcolumns)                   | Insert columns to 2 dimensional array.                                          |
| [removeDuplicatedValues](#removeduplicatedvalues) | Remove duplicated values from 1 dimensional array.                              |
| [get1stEmptyRow](#get1stemptyrow)                 | Retrieve empty row index.                                                       |
| [get1stEmptyColumn](#get1stemptycolumn)           | Retrieve empty column index.                                                    |
| [sum](#sum)                                       | Sum numbers in an array.                                                        |
| [compilingNumbers](#compilingnumbers)             | Compiling Continuous Numbers using Google Apps Script.                          |
| [convArrayToObject](#convarraytoobject)           | Converting 2 dimensional array to JSON object.                                  |
| [unpivot](#unpivot)                               | Converting 2-dimensional array as unpivot (reverse pivot).                      |
| [reverseUnpivot](#reverseunpivot)                 | Reversing 2-dimensional array with unpivot.                                     |

## For binary processing

| Methods                                                   | Description                                      |
| :-------------------------------------------------------- | :----------------------------------------------- |
| [convInt8ArrayToHexAr](#convint8arraytohexar)             | Convert Int8Array to hex string array.           |
| [convStrToHex](#convStrToHex)                             | Convert string to hex.                           |
| [convHexToInt8Ar](#convhextoint8ar)                       | Convert string to hex.                           |
| [convInt8ArToStr](#convint8artostr)                       | Convert Int8Array to string value.               |
| [convInt8ArToUint8Ar](#convint8artouint8ar)               | Convert Int8Array to Uint8Array.                 |
| [convUint8ArToInt8Ar](#convuint8artoint8ar)               | Convert Uint8Array to Int8Array.                 |
| [searchIndexFromDataByData](#searchindexfromdatabydata)   | Search index from base data using a search data. |
| [splitByteArrayBySearchData](#splitbytearraybysearchdata) | Split byteArray by a search data.                |

## For string processing

| Methods                                                 | Description                                                                   |
| :------------------------------------------------------ | :---------------------------------------------------------------------------- |
| [ConvText](#convtext)                                   | Converting text as unicode.                                                   |
| [columnLetterToIndex](#columnlettertoindex)             | Converting colum letter to column index. Start of column index is 0.          |
| [columnIndexToLetter](#columnindextoletter)             | Converting colum index to column letter. Start of column index is 0.          |
| [convA1NotationToGridRange](#conva1notationtogridrange) | Converting a1Notation to gridrange. This will be useful for using Sheets API. |
| [convGridRangeToA1Notation](#convgridrangetoa1notation) | Converting gridrange to a1Notation. This will be useful for using Sheets API. |
| [addQueryParameters](#addqueryparameters)               | This method is used for adding the query parameters to the URL.               |
| [parseQueryParameters](#parsequeryparameters)           | This method is used for parsing the URL including the query parameters.       |
| [expandA1Notations](#expandA1Notations)                 | This method is used for expanding A1Notations.                                |

## Show document in script editor

When you use this library with the script editor of Google Apps Script, you can see the document of each method by the autocompletion of the script editor. You can see the following demonstration.

![](images/fig2.gif)

# Scripts of all methods

---

## For array processing

<a name="is2dimensionalarray"></a>

### is2DimensionalArray

When the inputted array is 2 dimensional array, true is returned.

````javascript
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
  return array.every((r) => Array.isArray(r));
}
````

<a name="isuniform2darray"></a>

### isUniform2DArray

When the inputted 2 dimensional array is the uniformed array, true is returned.

````javascript
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
  return new Set(array.map((r) => r.length)).size == 1 ? true : false;
}
````

<a name="uniform2darray"></a>

### uniform2DArray

Make all array in 2 dimensional array uniforming the same length.

````javascript
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
  const maxLen = Math.max(...array.map((r) => r.length));
  return array.map((r) => [...r, ...Array(maxLen - r.length).fill(empty)]);
}
````

<a name="transpose"></a>

### transpose

Transpose 2 dimensional array.

````javascript
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
  return array[0].map((_, col) => array.map((row) => row[col] || null));
}
````

<a name="splitarray"></a>

### splitArray

Split array every n length. [Ref](https://tanaikech.github.io/2022/05/21/splitting-and-processing-an-array-every-n-length-using-google-apps-script/)

````javascript
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
  return [...Array(Math.ceil(array.length / size))].map((_) =>
    array.splice(0, size)
  );
}
````

<a name="getspecificcolumns"></a>

### getSpecificColumns

Retrieve the specific columns from 2 dimensional array.

````javascript
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
  return array.map((r) => columns.map((e) => r[e - 1] || null));
}
````

<a name="deletespecificcolumns"></a>

### deleteSpecificColumns

Delete specific columns from 2 dimensional array.

````javascript
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
  return array.map((r) => r.filter((_, j) => !columns.includes(j + 1)));
}
````

<a name="insertcolumns"></a>

### insertColumns

Insert columns to 2 dimensional array.

````javascript
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
function insertColumns(
  array,
  insertArray,
  column = 1,
  rep = false,
  check = true
) {
  if (check && !is2DimensionalArray(array)) {
    throw new Error("Please use 2 dimensional array.");
  }
  if (
    !insertArray ||
    !Array.isArray(insertArray) ||
    insertArray.length == 0 ||
    !is2DimensionalArray(insertArray)
  ) {
    throw new Error("Please set columns you want to insert.");
  }
  const t = transpose(array, false);
  t.splice(column - 1, rep ? 1 : 0, ...transpose(insertArray, false));
  return transpose(t, false);
}
````

<a name="removeduplicatedvalues"></a>

### removeDuplicatedValues

Remove duplicated values from 1 dimensional array.

````javascript
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
  const obj = array.reduce(
    (m, e) => m.set(e, m.has(e) ? m.get(e) + 1 : 1),
    new Map()
  );
  const e = [...obj.entries()];
  return {
    removeDuplicatedValues: [...obj.keys()],
    duplicatedValues: e.reduce((ar, [k, v]) => {
      if (v != 1) ar.push(k);
      return ar;
    }, []),
    numberOfDuplicate: Object.fromEntries(e),
  };
}
````

<a name="get1stemptyrow"></a>

### get1stEmptyRow

Retrieve empty row index.

````javascript
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
    topEmptyRow: array.findIndex((r) => r.join("") == ""),
    lastEmptyRow:
      array.length - 1 - array.reverse().findIndex((r) => r.join("") == ""),
  };
}
````

<a name="get1stemptycolumn"></a>

### get1stEmptyColumn

Retrieve empty column index.

````javascript
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
  return array.reduce(
    (o, r) => {
      let top = r.findIndex(
        (c) => typeof c === "undefined" || c.toString() == ""
      );
      let last =
        r.length -
        1 -
        r
          .reverse()
          .findIndex((c) => typeof c === "undefined" || c.toString() == "");
      top = top == -1 ? r.length : top;
      last = last == -1 ? 0 : last;
      o.topEmptyColumn = o.topEmptyColumn < top ? o.topEmptyColumn : top;
      o.lastEmptyColumn = o.lastEmptyColumn > last ? o.lastEmptyColumn : last;
      return o;
    },
    { topEmptyColumn: Infinity, lastEmptyColumn: 0 }
  );
}
````

<a name="sum"></a>

### sum

Sum numbers in an array.

````javascript
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
  if (check && !(Array.isArray(array) && array.every((e) => !isNaN(e)))) {
    throw new Error("Please give an array including numbers.");
  }
  return array.reduce((n, e) => (n += e), 0);
}
````

<a name="compilingnumbers"></a>

### compilingNumbers

Compiling Continuous Numbers using Google Apps Script. [Ref](https://tanaikech.github.io/2021/10/08/compiling-continuous-numbers-using-google-apps-script/)

````javascript
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
  if (!(Array.isArray(array) && array.every((e) => !isNaN(e)))) {
    throw new Error("Please give an array including numbers.");
  }
  const { values } = [...new Set(array.sort((a, b) => a - b))].reduce(
    (o, e, i, a) => {
      if (
        o.temp.length == 0 ||
        (o.temp.length > 0 && e == o.temp[o.temp.length - 1] + 1)
      ) {
        o.temp.push(e);
      } else {
        if (o.temp.length > 0) {
          o.values.push({ start: o.temp[0], end: o.temp[o.temp.length - 1] });
        }
        o.temp = [e];
      }
      if (i == a.length - 1) {
        o.values.push(
          o.temp.length > 1
            ? { start: o.temp[0], end: o.temp[o.temp.length - 1] }
            : { start: e, end: e }
        );
      }
      return o;
    },
    { temp: [], values: [] }
  );
  return values;
}
````

<a name="convarraytoobject"></a>

### convArrayToObject

Converting 2 dimensional array to JSON object. [Ref](https://tanaikech.github.io/2021/10/24/converting-values-of-google-spreadsheet-to-object-using-google-apps-script/)

````javascript
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
  if (
    !Array.isArray(headers) ||
    !Array.isArray(rows) ||
    !is2DimensionalArray(rows)
  ) {
    throw new Error("Please give an array of header and values.");
  }
  return rows.map((r) => headers.reduce((o, h, j) => ((o[h] = r[j]), o), {}));
}
````

<a name="unpivot"></a>

### unpivot

Converting 2-dimensional array as unpivot (reverse pivot). [Ref](https://tanaikech.github.io/2023/05/11/unpivot-on-google-spreadsheet-using-google-apps-script/)

````javascript
/**
 * ### Description
 * Converting 2-dimensional array as unpivot (reverse pivot).
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
  return h.flatMap((hh, i) => v.map((t) => [hh, t[0], t[i + 1]]));
}
````

<a name="reverseunpivot"></a>

### reverseUnpivot

Reversing 2-dimensional array with unpivot. [Ref](https://tanaikech.github.io/2023/05/11/unpivot-on-google-spreadsheet-using-google-apps-script/)

````javascript
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
  const [a, b, c] = values[0].map((_, c) => values.map((r) => r[c]));
  const ch = [...new Set(a)];
  const rh = [...new Set(b)];
  const size = rh.length;
  const temp = [
    [null, ...rh],
    ...[...Array(Math.ceil(c.length / size))]
      .map((_) => c.splice(0, size))
      .map((vv, i) => [ch[i], ...vv]),
  ];
  return temp[0].map((_, c) => temp.map((r) => r[c]));
}
````

---

## For binary processing

<a name="convInt8arraytohexar"></a>

### convInt8ArrayToHexAr

Convert Int8Array to hex string array. [Ref](https://tanaikech.github.io/2021/12/18/converting-from-string-to-hex-from-hex-to-bytes-from-bytes-to-string-using-google-apps-script/)

````javascript
/**
 * ### Description
 * Convert Int8Array to hex string array.
 * Ref: https://tanaikech.github.io/2021/12/18/converting-from-string-to-hex-from-hex-to-bytes-from-bytes-to-string-using-google-apps-script/
 *
 * ### Sample script
 * ```
 * const ar = Utilities.newBlob("さんぷる").getBytes();
 * const res = UtlApp.convInt8ArrayToHexAr(ar);
 * console.log(res); // [ 'e3', '81', '95', 'e3', '82', '93', 'e3', '81', 'b7', 'e3', '82', '8b' ]
 * ```
 *
 * @param {Array} array int8 array.
 * @return {Array} Hex string array.
 */
function convInt8ArrayToHexAr(array) {
  if (!Array.isArray(array)) {
    throw new Error("Please give Int8Array.");
  }
  return array.map((byte) => ("0" + (byte & 0xff).toString(16)).slice(-2));
}
````

<a name="convstrtohex"></a>

### convStrToHex

Convert string to hex. [Ref](https://tanaikech.github.io/2021/12/18/converting-from-string-to-hex-from-hex-to-bytes-from-bytes-to-string-using-google-apps-script/)

````javascript
/**
 * ### Description
 * Convert string to hex.
 * Ref: https://tanaikech.github.io/2021/12/18/converting-from-string-to-hex-from-hex-to-bytes-from-bytes-to-string-using-google-apps-script/
 *
 * ### Sample script
 * ```
 * const str = "sample string";
 * const res = UtlApp.convStrToHex(str);
 * console.log(res); // 73616d706c6520737472696e67
 * ```
 *
 * @param {String} str Input string value.
 * @return {String} Hex string value.
 */
function convStrToHex(str) {
  if (!str) {
    throw new Error("Please give a string value.");
  }
  return convInt8ArrayToHexAr(Utilities.newBlob(str).getBytes()).join("");
}
````

<a name="convhextoInt8ar"></a>

### convHexToInt8Ar

Convert string to hex. [Ref](https://tanaikech.github.io/2021/12/18/converting-from-string-to-hex-from-hex-to-bytes-from-bytes-to-string-using-google-apps-script/)

````javascript
/**
 * ### Description
 * Convert string to hex.
 * Ref: https://tanaikech.github.io/2021/12/18/converting-from-string-to-hex-from-hex-to-bytes-from-bytes-to-string-using-google-apps-script/
 *
 * ### Sample script
 * ```
 * const hex = "73616d706c6520737472696e67";
 * const res = UtlApp.convHexToInt8Ar(hex);
 * console.log(res); // [ 115, 97, 109, 112, 108, 101, 32, 115, 116, 114, 105, 110, 103 ]
 *
 * console.log(UtlApp.convInt8ArToStr(res)); // sample string
 * ```
 *
 * @param {String} hex Input hex string value.
 * @return {Array} Converted int8array.
 */
function convHexToInt8Ar(hex) {
  if (!hex) {
    throw new Error("Please give a hex string value.");
  }
  return hex
    .match(/.{2}/g)
    .map((e) =>
      parseInt(e[0], 16).toString(2).length == 4
        ? parseInt(e, 16) - 256
        : parseInt(e, 16)
    );
}
````

<a name="convint8artostr"></a>

### convInt8ArToStr

Convert Int8Array to string value.

````javascript
/**
 * ### Description
 * Convert Int8Array to string value.
 *
 * ### Sample script
 * ```
 * const hex = "73616d706c6520737472696e67";
 * const res = UtlApp.convHexToInt8Ar(hex);
 * console.log(res); // [ 115, 97, 109, 112, 108, 101, 32, 115, 116, 114, 105, 110, 103 ]
 *
 * console.log(UtlApp.convInt8ArToStr(res)); // sample string
 * ```
 *
 * @param {Array} array Input int8array.
 * @return {String} Converted string value.
 */
function convInt8ArToStr(array) {
  if (!Array.isArray(array)) {
    throw new Error("Please give Int8Array.");
  }
  return Utilities.newBlob(array).getDataAsString();
}
````

<a name="convint8artouint8ar"></a>

### convInt8ArToUint8Ar

Convert Int8Array to Uint8Array.

````javascript
/**
 * ### Description
 * Convert Int8Array to Uint8Array.
 *
 * ### Sample script
 * ```
 * const str = "さんぷる";
 * const int8Ar = Utilities.newBlob(str).getBytes();
 * console.log(int8Ar); // [ -29, -127, -107, -29, -126, -109, -29, -127, -73, -29, -126, -117 ]
 *
 * const res1 = UtlApp.convInt8ArToUint8Ar(int8Ar);
 * console.log(res1); // [ 227, 129, 149, 227, 130, 147, 227, 129, 183, 227, 130, 139 ]
 *
 * const res2 = UtlApp.convUint8ArToInt8Ar(res1);
 * console.log(res2); // [ -29, -127, -107, -29, -126, -109, -29, -127, -73, -29, -126, -117 ]
 * ```
 *
 * @param {Array} array Input int8array.
 * @return {Array} unit8array.
 */
function convInt8ArToUint8Ar(array) {
  if (!Array.isArray(array)) {
    throw new Error("Please give Int8Array.");
  }
  return [...Uint8Array.from(array)];
}
````

<a name="convuint8artoInt8ar"></a>

### convUint8ArToInt8Ar

Convert Uint8Array to Int8Array.

````javascript
/**
 * ### Description
 * Convert Uint8Array to Int8Array.
 *
 * ### Sample script
 * ```
 * const str = "さんぷる";
 * const int8Ar = Utilities.newBlob(str).getBytes();
 * console.log(int8Ar); // [ -29, -127, -107, -29, -126, -109, -29, -127, -73, -29, -126, -117 ]
 *
 * const res1 = UtlApp.convInt8ArToUint8Ar(int8Ar);
 * console.log(res1); // [ 227, 129, 149, 227, 130, 147, 227, 129, 183, 227, 130, 139 ]
 *
 * const res2 = UtlApp.convUint8ArToInt8Ar(res1);
 * console.log(res2); // [ -29, -127, -107, -29, -126, -109, -29, -127, -73, -29, -126, -117 ]
 * ```
 *
 * @param {Array} array Input unit8array.
 * @return {Array} int8array.
 */
function convUint8ArToInt8Ar(array) {
  if (!Array.isArray(array)) {
    throw new Error("Please give Uint8Array.");
  }
  return [...Int8Array.from(array)];
}
````

<a name="searchIndexfromdatabydata"></a>

### searchIndexFromDataByData

Search index from base data using a search data.

````javascript
/**
 * ### Description
 * Search index from base data using a search data.
 *
 * ### Sample script
 * ```
 * const sampleString = "abc123def123ghi123jkl";
 * const splitValue = "123";
 * const baseData = Utilities.newBlob(sampleString).getBytes();
 * const searchData = Utilities.newBlob(splitValue).getBytes();
 * const res = UtlApp.searchIndexFromDataByData(baseData, searchData);
 * console.log(res); // 3
 * ```
 *
 * @param {Array} baseData Input byteArray of base data.
 * @param {Array} searchData Input byteArray of search data using split.
 * @return {Number} Index of search data. If the search data is not found, -1 is returned.
 */
function searchIndexFromDataByData(baseData, searchData) {
  if (!Array.isArray(baseData) || !Array.isArray(searchData)) {
    throw new Error("Please give Int8Array.");
  }
  const search = searchData.join("");
  const bLen = searchData.length;
  return baseData.findIndex(
    (_, i, a) => [...Array(bLen)].map((_, j) => a[j + i]).join("") == search
  );
}
````

<a name="splitbytearraybysearchdata"></a>

### splitByteArrayBySearchData

Split byteArray by a search data. [Ref](https://tanaikech.github.io/2023/03/08/split-binary-data-with-search-data-using-google-apps-script/)

````javascript
/**
 * ### Description
 * Split byteArray by a search data.
 *
 * ### Sample script
 * ```
 * const sampleString = "abc123def123ghi123jkl";
 * const splitValue = "123";
 * const baseData = Utilities.newBlob(sampleString).getBytes();
 * const searchData = Utilities.newBlob(splitValue).getBytes();
 * const res = UtlApp.splitByteArrayBySearchData(baseData, searchData);
 * console.log(res);
 * ```
 *
 * Result is as follows.
 *
 * ```
 * [
 *   [ 97, 98, 99 ],
 *   [ 100, 101, 102 ],
 *   [ 103, 104, 105 ],
 *   [ 106, 107, 108 ]
 * ]
 * ```
 *
 * @param {Array} baseData Input byteArray of base data.
 * @param {Array} searchData Input byteArray of search data using split.
 * @return {Array} An array including byteArray.
 */
function splitByteArrayBySearchData(baseData, searchData) {
  if (!Array.isArray(baseData) || !Array.isArray(searchData)) {
    throw new Error("Please give Int8Array.");
  }
  const search = searchData.join("");
  const bLen = searchData.length;
  const res = [];
  let idx = 0;
  do {
    idx = baseData.findIndex(
      (_, i, a) => [...Array(bLen)].map((_, j) => a[j + i]).join("") == search
    );
    if (idx != -1) {
      res.push(baseData.splice(0, idx));
      baseData.splice(0, bLen);
    } else {
      res.push(baseData.splice(0));
    }
  } while (idx != -1);
  return res;
}
````

---

## For string processing

<a name="convtext"></a>

### ConvText

Converting text as unicode. [Ref](https://tanaikech.github.io/2020/11/13/converting-texts-to-bold-italic-and-bold-italic-types-of-unicode-using-google-apps-script/)

````javascript
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
        if (
          (t >= 48 && t <= 57) ||
          (t >= 65 && t <= 90) ||
          (t >= 97 && t <= 122)
        ) {
          return obj.reduce((s, { r, d }) => {
            if (new RegExp(`[${r}]`).test(e))
              s = String.fromCodePoint(e.codePointAt(0) + d);
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
    return this.c_(text, [
      { r: "0-9", d: 120734 },
      { r: "A-Z", d: 120211 },
      { r: "a-z", d: 120205 },
    ]);
  },

  /**
   * Convert text to italic type.
   * Ref: https://tanaikech.github.io/2020/11/13/converting-texts-to-bold-italic-and-bold-italic-types-of-unicode-using-google-apps-script/
   * @param {String} text
   * @return {String} Converted text.
   */
  italic: function (text) {
    return this.c_(text, [
      { r: "A-Z", d: 120263 },
      { r: "a-z", d: 120257 },
    ]);
  },

  /**
   * Convert text to bold and italic type.
   * @param {String} text
   * @return {String} Converted text.
   */
  boldItalic: function (text) {
    return this.c_(text, [
      { r: "A-Z", d: 120315 },
      { r: "a-z", d: 120309 },
    ]);
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
````

<a name="columnlettertoIndex"></a>

### columnLetterToIndex

Converting colum letter to column index. Start of column index is 0. [Ref](https://tanaikech.github.io/2022/05/01/increasing-column-letter-by-one-using-google-apps-script/)

````javascript
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
  return [...letter].reduce(
    (c, e, i, a) =>
      (c += (e.charCodeAt(0) - 64) * Math.pow(26, a.length - i - 1)),
    -1
  );
}
````

<a name="columnIndextoletter"></a>

### columnIndexToLetter

Converting colum index to column letter. Start of column index is 0. [Ref](https://stackoverflow.com/a/53678158/7108653)

````javascript
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
    throw new Error(
      "Please give the column indexr as a number. In this case, 1st number is 0."
    );
  }
  return (a = Math.floor(index / 26)) >= 0
    ? columnIndexToLetter(a - 1) + String.fromCharCode(65 + (index % 26))
    : "";
}
````

<a name="conva1notationtogridrange"></a>

### convA1NotationToGridRange

Converting a1Notation to gridrange. This will be useful for using Sheets API. [Ref](https://tanaikech.github.io/2017/07/31/converting-a1notation-to-gridrange-for-google-sheets-api/)

````javascript
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
  if (
    a1Notation === null ||
    sheetId === null ||
    typeof a1Notation != "string" ||
    isNaN(sheetId)
  ) {
    throw new Error("Please give a1Notation (string) and sheet ID (integer).");
  }
  const { col, row } = a1Notation
    .toUpperCase()
    .split("!")
    .map((f) => f.split(":"))
    .pop()
    .reduce(
      (o, g) => {
        var [r1, r2] = ["[A-Z]+", "[0-9]+"].map((h) => g.match(new RegExp(h)));
        o.col.push(r1 && columnLetterToIndex(r1[0]));
        o.row.push(r2 && Number(r2[0]));
        return o;
      },
      { col: [], row: [] }
    );
  col.sort((a, b) => (a > b ? 1 : -1));
  row.sort((a, b) => (a > b ? 1 : -1));
  const [start, end] = col.map((e, i) => ({ col: e, row: row[i] }));
  const obj = {
    sheetId,
    startRowIndex: start?.row && start.row - 1,
    endRowIndex: end?.row ? end.row : start.row,
    startColumnIndex: start && start.col,
    endColumnIndex: end ? end.col + 1 : 1,
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
````

<a name="convgridrangetoa1Notation"></a>

### convGridRangeToA1Notation

Converting gridrange to a1Notation. This will be useful for using Sheets API. [Ref](https://tanaikech.github.io/2017/07/31/converting-a1notation-to-gridrange-for-google-sheets-api/)

````javascript
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
  const a1Notation =
    st == en ? `'${sheetName}'!${st}` : `'${sheetName}'!${st}:${en}`;
  return a1Notation;
}
````

<a name="addqueryparameters"></a>

### addQueryParameters

This method is used for adding the query parameters to the URL. [Ref](https://tanaikech.github.io/2018/07/12/adding-query-parameters-to-url-using-google-apps-script/)

````javascript
/**
 * ### Description
 * This method is used for adding the query parameters to the URL.
 * Ref: https://tanaikech.github.io/2018/07/12/adding-query-parameters-to-url-using-google-apps-script/
 *
 * ### Sample script
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
 * @param {String} url The base URL for adding the query parameters.
 * @param {Object} obj JSON object including query parameters.
 * @return {String} URL including the query parameters.
 */
function addQueryParameters(url, obj) {
  if (url === null || obj === null || typeof url != "string") {
    throw new Error(
      "Please give URL (String) and query parameter (JSON object)."
    );
  }
  return (
    url +
    "?" +
    Object.entries(obj)
      .flatMap(([k, v]) =>
        Array.isArray(v)
          ? v.map((e) => `${k}=${encodeURIComponent(e)}`)
          : `${k}=${encodeURIComponent(v)}`
      )
      .join("&")
  );
}
````

<a name="parsequeryparameters"></a>

### parseQueryParameters

This method is used for parsing the URL including the query parameters. [Ref](https://tanaikech.github.io/2018/07/12/adding-query-parameters-to-url-using-google-apps-script/)

````javascript
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
````

<a name="expanda1notations"></a>

### expandA1Notations

This method is used for expanding A1Notations. [Ref](https://tanaikech.github.io/2020/04/04/updated-expanding-a1notations-using-google-apps-script/)

````javascript
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
  return a1Notations.map((e) => {
    const a1 = e.split("!");
    const r = a1.length > 1 ? a1[1] : a1[0];
    const [r1, r2] = r.split(":");
    if (!r2) return [r1];
    let rr;
    if (reg1.test(r1) && reg1.test(r2)) {
      rr = [r1.toUpperCase().match(reg1), r2.toUpperCase().match(reg1)];
    } else if (reg2.test(r1) && reg2.test(r2)) {
      rr = [
        [null, r1, 1],
        [null, r2, maxRow],
      ];
    } else if (reg1.test(r1) && reg2.test(r2)) {
      rr = [r1.toUpperCase().match(reg1), [null, r2, maxRow]];
    } else if (reg2.test(r1) && reg1.test(r2)) {
      rr = [[null, r1, maxRow], r2.toUpperCase().match(reg1)];
    } else if (reg3.test(r1) && reg3.test(r2)) {
      rr =
        Number(r1) > Number(r2)
          ? [
              [null, "A", r2],
              [null, maxColumn, r1],
            ]
          : [
              [null, "A", r1],
              [null, maxColumn, r2],
            ];
    } else if (reg1.test(r1) && reg3.test(r2)) {
      rr = [r1.toUpperCase().match(reg1), [null, maxColumn, r2]];
    } else if (reg3.test(r1) && reg1.test(r2)) {
      let temp = r2.toUpperCase().match(reg1);
      rr =
        Number(temp[2]) > Number(r1)
          ? [
              [null, temp[1], r1],
              [null, maxColumn, temp[2]],
            ]
          : [temp, [null, maxColumn, r1]];
    } else {
      throw new Error(`Wrong a1Notation: ${r}`);
    }
    const obj = {
      startRowIndex: Number(rr[0][2]),
      endRowIndex: rr.length == 1 ? Number(rr[0][2]) + 1 : Number(rr[1][2]) + 1,
      startColumnIndex: columnLetterToIndex(rr[0][1]),
      endColumnIndex:
        rr.length == 1
          ? columnLetterToIndex(rr[0][1]) + 1
          : columnLetterToIndex(rr[1][1]) + 1,
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
````

---

# If you want to use only one method in the above methods without using a library

If you want to use only one method in the above methods without a library, you can do this simply.

In that case, please copy and paste one of the methods from the above scripts. At this time, there are several methods using the other methods. Please copy and paste all dependence methods.

If you want to run the method directly, please remove the identification of the library `UtlApp`. By this, the script is run.

For example, when you want to directly use the following method `is2DimensionalArray` without the library, please copy and paste the following script to the script editor.

````javascript
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
  return array.every((r) => Array.isArray(r));
}
````

When you run this method, please use it with the following script.

```javascript
function sample() {
  const array1 = ["", [1, 2, 3], [1, 2, 3], [1, 2, 3]];
  const res1 = is2DimensionalArray(array1); // false
}
```

# Contribution

I believe that these methods will help to develop the applications created by Google Apps Script. So, I would like to grow this. If you want to add more methods, I welcome your contributions. At that time, please do "pull request" or email me ( tanaike@hotmail.com ). By testing your method, I would like to add it. Your name is also added as a contributor.

**The policy of these methods is not to use the scopes.** So, when you add your method, **please confirm that your proposed method uses no scopes**. And also, please check your JSDoc.

---

<a name="licence"></a>

# Licence

[MIT](LICENCE)

<a name="author"></a>

# Author

[Tanaike](https://tanaikech.github.io/about/)

[Donate](https://tanaikech.github.io/donate/)

<a name="updatehistory"></a>

# Update History

- v1.0.0 (May 13, 2023)

  1. Initial release.

[TOP](#top)
