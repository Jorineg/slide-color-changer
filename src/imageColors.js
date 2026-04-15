import JSZip from 'jszip';

const MAX_DOMINANT_COLORS = 3;
const MIN_COLOR_PIXELS_PCT = 1.5;
const HUE_BUCKET_SIZE = 12;
const SATURATION_THRESHOLD = 15;
const ACHROMATIC_VALUE_MIN = 20;
const ACHROMATIC_VALUE_MAX = 90;

function rgbToHsv(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const d = max - min;
  const v = max;
  const s = max === 0 ? 0 : d / max;
  let h = 0;
  if (d !== 0) {
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  return [h * 360, s * 100, v * 100];
}

function hsvToRgb(h, s, v) {
  h /= 360; s /= 100; v /= 100;
  let r, g, b;
  const i = Math.floor(h * 6);
  const f = h * 6 - i;
  const p = v * (1 - s);
  const q = v * (1 - f * s);
  const t = v * (1 - (1 - f) * s);
  switch (i % 6) {
    case 0: r = v; g = t; b = p; break;
    case 1: r = q; g = v; b = p; break;
    case 2: r = p; g = v; b = t; break;
    case 3: r = p; g = q; b = v; break;
    case 4: r = t; g = p; b = v; break;
    case 5: r = v; g = p; b = q; break;
  }
  return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}

function hexToRgb(hex) {
  return [
    parseInt(hex.substring(0, 2), 16),
    parseInt(hex.substring(2, 4), 16),
    parseInt(hex.substring(4, 6), 16),
  ];
}

function rgbToHex(r, g, b) {
  return [r, g, b].map(c => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function loadImageFromBytes(bytes, mime) {
  return new Promise((resolve, reject) => {
    const blob = new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Failed to load image')); };
    img.src = url;
  });
}

function getImagePixels(img) {
  const canvas = document.createElement('canvas');
  canvas.width = img.width;
  canvas.height = img.height;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(img, 0, 0);
  return ctx.getImageData(0, 0, canvas.width, canvas.height);
}

/**
 * Find dominant chromatic colors in an image using hue bucketing.
 * Returns up to MAX_DOMINANT_COLORS hex colors, or empty array if
 * the image is too complex (photos/screenshots).
 */
function findDominantColors(imageData) {
  const { data, width, height } = imageData;
  const totalPixels = width * height;
  const hueBuckets = new Array(Math.ceil(360 / HUE_BUCKET_SIZE)).fill(0);
  const bucketRgbSum = hueBuckets.map(() => [0, 0, 0]);
  let chromaticPixels = 0;

  for (let i = 0; i < data.length; i += 4) {
    const r = data[i], g = data[i + 1], b = data[i + 2], a = data[i + 3];
    if (a < 30) continue;

    const [h, s, v] = rgbToHsv(r, g, b);

    if (s < SATURATION_THRESHOLD) continue;
    if (v < ACHROMATIC_VALUE_MIN || v > ACHROMATIC_VALUE_MAX * 1.11) {
      // keep going, this is still chromatic if sat is high enough
    }
    if (s < SATURATION_THRESHOLD) continue;

    chromaticPixels++;
    const bucket = Math.floor(h / HUE_BUCKET_SIZE) % hueBuckets.length;
    hueBuckets[bucket]++;
    bucketRgbSum[bucket][0] += r;
    bucketRgbSum[bucket][1] += g;
    bucketRgbSum[bucket][2] += b;
  }

  if (chromaticPixels < totalPixels * 0.01) return [];

  // Merge adjacent buckets into clusters
  const clusters = [];
  let i = 0;
  while (i < hueBuckets.length) {
    if (hueBuckets[i] === 0) { i++; continue; }

    let count = hueBuckets[i];
    let rSum = bucketRgbSum[i][0];
    let gSum = bucketRgbSum[i][1];
    let bSum = bucketRgbSum[i][2];
    let j = i + 1;

    while (j < hueBuckets.length && hueBuckets[j] > 0) {
      count += hueBuckets[j];
      rSum += bucketRgbSum[j][0];
      gSum += bucketRgbSum[j][1];
      bSum += bucketRgbSum[j][2];
      j++;
    }

    clusters.push({
      count,
      r: Math.round(rSum / count),
      g: Math.round(gSum / count),
      b: Math.round(bSum / count),
    });
    i = j;
  }

  // Also merge the wrap-around (last cluster with first if both exist)
  if (clusters.length >= 2) {
    const first = clusters[0];
    const last = clusters[clusters.length - 1];
    const firstHue = rgbToHsv(first.r, first.g, first.b)[0];
    const lastHue = rgbToHsv(last.r, last.g, last.b)[0];
    if ((360 - lastHue + firstHue) < HUE_BUCKET_SIZE * 3) {
      const merged = first.count + last.count;
      first.r = Math.round((first.r * first.count + last.r * last.count) / merged);
      first.g = Math.round((first.g * first.count + last.g * last.count) / merged);
      first.b = Math.round((first.b * first.count + last.b * last.count) / merged);
      first.count = merged;
      clusters.pop();
    }
  }

  clusters.sort((a, b) => b.count - a.count);

  const minPixels = totalPixels * MIN_COLOR_PIXELS_PCT / 100;
  const significant = clusters.filter(c => c.count >= minPixels);

  if (significant.length === 0 || significant.length > MAX_DOMINANT_COLORS) {
    return [];
  }

  return significant.map(c => ({
    hex: rgbToHex(c.r, c.g, c.b),
    pixelCount: c.count,
  }));
}

/**
 * Scan all images in ppt/media/ and extract dominant colors.
 * Returns a Map of imagePath -> [{ hex, pixelCount }]
 */
export async function extractImageColors(zipOrBuffer) {
  const zip = zipOrBuffer instanceof JSZip
    ? zipOrBuffer
    : await JSZip.loadAsync(zipOrBuffer);

  const imageFiles = zip.file(/ppt\/media\/.*\.(png|jpg|jpeg|gif)$/i);
  const imageColorMap = new Map();

  for (const file of imageFiles) {
    try {
      const bytes = await file.async('uint8array');
      const ext = file.name.split('.').pop().toLowerCase();
      const mime = ext === 'png' ? 'image/png'
        : ext === 'gif' ? 'image/gif'
        : 'image/jpeg';

      const img = await loadImageFromBytes(bytes, mime);
      const imageData = getImagePixels(img);
      const colors = findDominantColors(imageData);

      if (colors.length > 0) {
        imageColorMap.set(file.name, colors);
      }
    } catch {
      // Skip images that fail to decode
    }
  }

  return imageColorMap;
}

/**
 * Apply HSV hue-shift to an image and return the modified PNG as Uint8Array.
 *
 * @param {Uint8Array} imageBytes - Original image bytes
 * @param {string} mime - MIME type
 * @param {Map<string, string>} colorReplacements - Map of origHex -> newHex (uppercase, no #)
 * @returns {Promise<Uint8Array|null>} Modified PNG bytes, or null if no changes needed
 */
export async function recolorImage(imageBytes, mime, colorReplacements) {
  if (colorReplacements.size === 0) return null;

  const img = await loadImageFromBytes(imageBytes, mime);
  const canvas = document.createElement('canvas');
  canvas.width = img.width;
  canvas.height = img.height;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(img, 0, 0);
  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const { data } = imageData;

  // Build hue-shift rules from the color replacements
  const hueShifts = [];
  for (const [origHex, newHex] of colorReplacements) {
    const [or, og, ob] = hexToRgb(origHex);
    const [origH, origS] = rgbToHsv(or, og, ob);
    const [nr, ng, nb] = hexToRgb(newHex);
    const [newH, newS] = rgbToHsv(nr, ng, nb);

    if (origS < 10) continue;

    hueShifts.push({
      origHue: origH,
      hueDelta: newH - origH,
      satRatio: newS / Math.max(origS, 1),
      tolerance: 30,
    });
  }

  if (hueShifts.length === 0) return null;

  let changed = false;
  for (let i = 0; i < data.length; i += 4) {
    const a = data[i + 3];
    if (a < 10) continue;

    const r = data[i], g = data[i + 1], b = data[i + 2];
    const [h, s, v] = rgbToHsv(r, g, b);

    if (s < 8) continue;

    for (const shift of hueShifts) {
      let hueDiff = Math.abs(h - shift.origHue);
      if (hueDiff > 180) hueDiff = 360 - hueDiff;

      if (hueDiff <= shift.tolerance) {
        let newH = (h + shift.hueDelta + 360) % 360;
        let newS = Math.min(100, s * shift.satRatio);
        const [nr, ng, nb] = hsvToRgb(newH, newS, v);
        data[i] = nr;
        data[i + 1] = ng;
        data[i + 2] = nb;
        changed = true;
        break;
      }
    }
  }

  if (!changed) return null;

  ctx.putImageData(imageData, 0, 0);

  const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

export { rgbToHsv, hsvToRgb, hexToRgb, rgbToHex };
