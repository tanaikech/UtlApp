/**
 * [Tanaike](https://tanaikech.github.io/)
 * [GitHub](https://github.com/tanaikech/UtlApp)
*/

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
  return array.map(byte => ("0" + (byte & 0xff).toString(16)).slice(-2));
}

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
  return hex.match(/.{2}/g).map((e) => parseInt(e[0], 16).toString(2).length == 4 ? parseInt(e, 16) - 256 : parseInt(e, 16));
}

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
  return baseData.findIndex((_, i, a) => [...Array(bLen)].map((_, j) => a[j + i]).join("") == search);
}

/**
 * ### Description
 * Split byteArray by a search data.
 * Ref: https://tanaikech.github.io/2023/03/08/split-binary-data-with-search-data-using-google-apps-script/
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
    idx = baseData.findIndex((_, i, a) => [...Array(bLen)].map((_, j) => a[j + i]).join("") == search);
    if (idx != -1) {
      res.push(baseData.splice(0, idx));
      baseData.splice(0, bLen);
    } else {
      res.push(baseData.splice(0));
    }
  } while (idx != -1);
  return res;
}
