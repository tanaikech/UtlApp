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

/**
 * ### Description
 * Converts a byte data of "audio/L16" to a byte data of "audio/wav".
 * L16 assumes 16-bit PCM.
 * This audio format is often used with Text-To-Speech (TTS).
 * Ref: https://datatracker.ietf.org/doc/html/rfc2586
 * Ref: https://medium.com/google-cloud/text-to-speech-tts-using-gemini-api-with-google-apps-script-6ece50a617fd
 *
 * ### Sample script
 * ```
 * const fileId = "###"; // file ID of "audio/L16";
 * const data = DriveApp.getFileById(fileId).getBlob().getBytes();
 * const sampleRate = 24000;
 * const numChannels = 1;
 * const convertedData = UtlApp.convertL16ToWav(data, sampleRate, numChannels); // "audio/wav"
 * ```
 * 
 * When this sample script is run, when the data is "audio/L16", the data is converted to "audio/wav".
 * 
 * @param {Byte[]} inputData Input data (audio/L16).
 * @param {number} sampleRate Ex. 8000, 11025, 16000, 22050, 24000, 32000, 44100, and 48000. Default is 24000.
 * @param {number} numChannels - Mono and stereo are 1 and 2, respectively. Default is 1.
 * @return {Byte[]} Converted data as byte data.
 */
function convertL16ToWav(inputData, sampleRate = 24000, numChannels = 1) {
  if (!Array.isArray(inputData)) {
    throw new Error("Invalid data.");
  }
  const bitsPerSample = 16;
  const blockAlign = numChannels * bitsPerSample / 8;
  const byteRate = Number(sampleRate) * blockAlign;
  const dataSize = inputData.length;
  const fileSize = 36 + dataSize;
  const header = new ArrayBuffer(44);
  const view = new DataView(header);
  const data = [
    { method: "setUint8", value: [..."RIFF"].map(e => e.charCodeAt(0)), add: [0, 1, 2, 3] },
    { method: "setUint32", value: [fileSize], add: [4], littleEndian: true },
    { method: "setUint8", value: [..."WAVE"].map(e => e.charCodeAt(0)), add: [8, 9, 10, 11] },
    { method: "setUint8", value: [..."fmt "].map(e => e.charCodeAt(0)), add: [12, 13, 14, 15] },
    { method: "setUint32", value: [16], add: [16], littleEndian: true },
    { method: "setUint16", value: [1, numChannels], add: [20, 22], littleEndian: true },
    { method: "setUint32", value: [Number(sampleRate), byteRate], add: [24, 28], littleEndian: true },
    { method: "setUint16", value: [blockAlign, bitsPerSample], add: [32, 34], littleEndian: true },
    { method: "setUint8", value: [..."data"].map(e => e.charCodeAt(0)), add: [36, 37, 38, 39] },
    { method: "setUint32", value: [dataSize], add: [40], littleEndian: true },
  ];
  data.forEach(({ method, value, add, littleEndian }) =>
    add.forEach((a, i) => view[method](a, value[i], littleEndian || false))
  );
  return [...new Uint8Array(header), ...inputData];
}

/**
 * ### Description
 * This method is used for retrieving the MP3 tag information.
 *
 * ### Sample script
 * ```
 * const blob = ###; // Please set your MP3 blob.
 * const res = UtlApp.getMP3Tag(blob);
 * ```
 * 
 * When this script is run, the MP3 tag information is returned as JSON.
 *
 * @param {Blob} blob Blob of MP3 data.
 * @return {Object} Object including MP3 tag information.
 */
function getMP3Tag(blob) {
  return new GetMP3Tag().run(blob);
}

/**
 * Class object for GetMP3Tag.
 * 
 * version 1.0.0
 * @class
 */
class GetMP3Tag {
  constructor() {
    /** @private */
    this.FRAME_ID_MAP = {
      'TIT2': 'title', 'TPE1': 'artist', 'TALB': 'album',
      'TRCK': 'track', 'TYER': 'year', 'TCON': 'genre',
      'COMM': 'comment', 'TPE2': 'albumArtist', 'TLEN': 'lengthMs',
      'TDRC': 'recordingTime'
    };

    /** @private */
    this.BITRATES = {
      'V1L1': [0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448],
      'V1L2': [0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384],
      'V1L3': [0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320],
    };

    /** @private */
    this.SAMPLING_RATES = {
      'MPEG1': [44100, 48000, 32000],
      'MPEG2': [22050, 24000, 16000],
      'MPEG2.5': [11025, 12000, 8000]
    };
  }

  /**
   * Get metadata and duration for an MP3 file in Google Drive by specifying its ID.
   *
   * @param {Blob} blob
   * @return {Object}
   */
  run(blob) {
    if (!blob) {
      throw new Error("Set MP3 blob.");
    }
    const mimeType = blob.getContentType();
    if (mimeType != "audio/mpeg") {
      throw new Error(`The mimeType is "${mimeType}". This data is not MP3.`);
    }
    const bytes = blob.getBytes();
    const fileSize = blob.getBytes().length;
    let metadata = {
      fileName: blob.getName(),
      fileSize: fileSize,
      mimeType: blob.getContentType(),
      id3v2: null,
      id3v1: null,
      duration: null,
      info: {}
    };

    // For ID3v2
    const id3v2Data = this.parseId3v2Tag_(bytes);
    if (id3v2Data) {
      metadata.id3v2 = id3v2Data.tags;
      Object.assign(metadata.info, id3v2Data.tags);
    }
    const id3v2TagSize = id3v2Data ? id3v2Data.size : 0;

    // For ID3v1
    const id3v1Data = this.parseId3v1Tag_(bytes, fileSize);
    if (id3v1Data) {
      metadata.id3v1 = id3v1Data;
      for (const key in id3v1Data) {
        if (!metadata.info[key]) {
          metadata.info[key] = id3v1Data[key];
        }
      }
    }

    // For duration
    metadata.duration = this.calculateDuration_(bytes, fileSize, id3v2TagSize);
    if (metadata.duration) {
      metadata.info.durationSeconds = metadata.duration.durationSeconds;
      metadata.info.durationFormatted = metadata.duration.formatted;
    }

    return metadata;
  }

  /**
   * Parsing ID3v2 tags
   * 
   * @param {Byte[]} bytes
   * @return {Object}
   * @private
   */
  parseId3v2Tag_(bytes) {
    if (bytes[0] !== 0x49 || bytes[1] !== 0x44 || bytes[2] !== 0x33) {
      return null;
    }
    const size = (bytes[6] << 21) | (bytes[7] << 14) | (bytes[8] << 7) | bytes[9];
    const totalTagSize = size + 10;
    let offset = 10;
    const tags = {};
    while (offset < totalTagSize) {
      const frameId = String.fromCharCode(bytes[offset], bytes[offset + 1], bytes[offset + 2], bytes[offset + 3]);
      if (frameId.charCodeAt(0) === 0) {
        break;
      }
      const frameSize = (bytes[offset + 4] << 24) | (bytes[offset + 5] << 16) | (bytes[offset + 6] << 8) | bytes[offset + 7];
      const frameDataStart = offset + 10;
      const frameDataEnd = frameDataStart + frameSize;
      if (frameDataEnd > totalTagSize) break;
      if (this.FRAME_ID_MAP[frameId]) {
        const encodingByte = bytes[frameDataStart];
        let charset = 'ISO-8859-1';
        let textStart = frameDataStart + 1;
        if (encodingByte === 1) {
          charset = 'UTF-16';
          textStart += 2;
        } else if (encodingByte === 3) {
          charset = 'UTF-8';
        }
        const textBytes = bytes.slice(textStart, frameDataEnd);
        let text = Utilities.newBlob(textBytes).getDataAsString(charset).replace(/\0/g, '').trim();
        if (frameId === 'TDRC' && !tags['year']) {
          tags['year'] = text.substring(0, 4);
        } else {
          tags[this.FRAME_ID_MAP[frameId]] = text;
        }
      }
      offset = frameDataEnd;
    }
    return { tags, size: totalTagSize };
  }

  /**
   * Parsing ID3v1 tags
   * 
   * @param {Byte[]} bytes
   * @param {Number} fileSize
   * @return {Object}
   * @private
   */
  parseId3v1Tag_(bytes, fileSize) {
    if (fileSize < 128) return null;
    const tagOffset = fileSize - 128;
    const tagBytes = bytes.slice(tagOffset);
    if (String.fromCharCode(tagBytes[0], tagBytes[1], tagBytes[2]) !== 'TAG') {
      return null;
    }
    const getString = (start, length) => {
      return Utilities.newBlob(tagBytes.slice(start, start + length)).getDataAsString('ISO-8859-1').replace(/\0/g, '').trim();
    };
    const tags = {};
    tags.title = getString(3, 30);
    tags.artist = getString(33, 30);
    tags.album = getString(63, 30);
    tags.year = getString(93, 4);
    tags.comment = getString(97, 30);
    if (tagBytes[125] === 0 && tagBytes[126] !== 0) {
      tags.track = tagBytes[126];
      tags.comment = getString(97, 28);
    }
    return tags;
  }

  /**
   * Calculating duration
   * 
   * @param {Byte[]} bytes
   * @param {Number} fileSize
   * @param {Number} id3v2TagSize
   * @return {Object}
   * @private
   */
  calculateDuration_(bytes, fileSize, id3v2TagSize) {
    let offset = id3v2TagSize;
    let headerFound = false;
    while (offset < bytes.length - 4 && !headerFound) {
      if ((bytes[offset] & 0xFF) === 0xFF && (bytes[offset + 1] & 0xE0) === 0xE0) {
        headerFound = true;
      } else {
        offset++;
      }
    }
    if (!headerFound) {
      return null;
    }
    const header_b1 = bytes[offset + 1] & 0xFF;
    const header_b2 = bytes[offset + 2] & 0xFF;
    const mpegVersionId = (header_b1 >> 3) & 0x03;
    const mpegVersion = mpegVersionId == 3 ? "MPEG1" : mpegVersionId == 2 ? "MPEG2" : mpegVersionId == 0 ? "MPEG2.5" : null;
    const layerId = (header_b1 >> 1) & 0x03;
    const layer = layerId == 3 ? "L1" : layerId == 2 ? "L2" : layerId == 1 ? "L1" : null;
    let bitrateKey;
    if (mpegVersion === 'MPEG1') {
      bitrateKey = 'V1' + layer;
    } else {
      console.warn(`MPEG Version ${mpegVersion} cannot be used.`);
      return null;
    }
    const bitrateIndex = (header_b2 >> 4) & 0x0F;
    if (bitrateIndex === 0 || bitrateIndex === 15) {
      return null;
    }
    const bitrate = this.BITRATES[bitrateKey][bitrateIndex] * 1000;
    const samplingRateIndex = (header_b2 >> 2) & 0x03;
    if (samplingRateIndex === 3) {
      return null;
    }
    const samplingRate = this.SAMPLING_RATES[mpegVersion][samplingRateIndex];
    if (!samplingRate) {
      return null;
    }
    const audioDataSize = fileSize - id3v2TagSize;
    const durationSeconds = Math.floor((audioDataSize * 8) / bitrate);
    const minutes = Math.floor(durationSeconds / 60);
    const seconds = durationSeconds % 60;
    const formatted = `${minutes}:${seconds < 10 ? '0' : ''}${seconds}`;
    return { mpegVersion, layer, bitrate, samplingRate, durationSeconds, formatted };
  }
}

/**
 * ### Description
 * Return information of blob.
 * The available mimeTypes are image/png, image/jpeg, image/gif, application/pdf, image/svg.
 * 
 * ### Sample script
 * ```
 * const blob = "###"; // Please set your blob.
 * const res = UtlApp.getInfFromBlob(blob);
 * console.log(res);
 * ```
 * 
 * Result is as follows.
 * 
 * ```
 * {"identification":"PNG","width":123,"height":456,"filesize":12345}
 * ```
 * 
 * @param {Blob} blob
 * @return {Object}
 */
function getInfFromBlob(blob) {
  const bytes = blob.getBytes();
  const filesize = bytes.length;
  if (filesize < 16) {
    throw new Error("File size is too small.");
  }
  const b = bytes.map(byte => byte & 0xFF);

  // For image/png
  if ([0x89, 0x50, 0x4E, 0x47].every((e, i) => b[i] == e)) {
    const width = (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19];
    const height = (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23];
    return { identification: "PNG", width, height, filesize };
  }

  // For image/jpeg
  else if ([0xFF, 0xD8].every((e, i) => b[i] == e)) {
    let pos = 2;
    while (pos < b.length - 8) {
      if (b[pos] !== 0xFF) {
        pos++; continue;
      }
      const marker = b[pos + 1];
      if ([0xC0, 0xC1, 0xC2].includes(marker)) {
        const height = (b[pos + 5] << 8) | b[pos + 6];
        const width = (b[pos + 7] << 8) | b[pos + 8];
        return { identification: "JPEG", width, height, filesize };
      }
      const segmentLength = (b[pos + 2] << 8) | b[pos + 3];
      pos += 2 + segmentLength;
    }
    return { identification: "JPEG", filesize };
  }

  // For image/gif
  else if ([0x47, 0x49, 0x46, 0x38].every((e, i) => b[i] == e)) {
    const width = b[6] | (b[7] << 8);
    const height = b[8] | (b[9] << 8);
    return { identification: "GIF", width, height, filesize };
  }

  // For application/pdf 
  else if ([0x25, 0x50, 0x44, 0x46].every((e, i) => b[i] == e)) {
    let pageCount = "N/A";
    const content = blob.getDataAsString("ISO-8859-1");
    const pagesMatch = content.match(/\/Type\s*\/Pages\s*[\s\S]*?\/Count\s+(\d+)/);
    if (pagesMatch && pagesMatch[1]) {
      pageCount = parseInt(pagesMatch[1], 10);
    } else {
      const countMatches = content.match(/\/Count\s+(\d+)/g);
      if (countMatches) {
        const counts = countMatches.map(match => parseInt(match.match(/\d+/)[0], 10));
        pageCount = Math.max(...counts);
      }
    }
    return { identification: "PDF", pageCount, filesize };
  }

  // For image/svg
  const content = blob.getDataAsString();
  if (content.trim().match(/^<\?xml[^>]*>\s*<svg|^<svg/i)) {
    const xml = blob.getDataAsString();
    const doc = XmlService.parse(xml);
    const root = doc.getRootElement();
    let width = root.getAttribute("width")?.getValue();
    let height = root.getAttribute("height")?.getValue();
    if (!width || !height) {
      const viewBox = root.getAttribute("viewBox")?.getValue();
      if (viewBox) {
        const parts = viewBox.split(/[\s,]+/);
        if (parts.length === 4) {
          width = parts[2];
          height = parts[3];
        }
      }
    }
    return { identification: "SVG", width: width || "N/A", height: height || "N/A", filesize };
  }

  return null;
}
