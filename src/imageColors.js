import JSZip from 'jszip';

const QUANT_STEP = 16;
const LAB_MERGE_DIST = 15;
const ENTROPY_THRESHOLD = 5.0;
const COVERAGE85_THRESHOLD = 30;
const CHROMATIC_SAT_MIN = 12;
const MIN_CLUSTER_PIXEL_PCT = 0.3;

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

function rgbToLab(r, g, b) {
  let rl = r / 255, gl = g / 255, bl = b / 255;
  rl = rl > 0.04045 ? Math.pow((rl + 0.055) / 1.055, 2.4) : rl / 12.92;
  gl = gl > 0.04045 ? Math.pow((gl + 0.055) / 1.055, 2.4) : gl / 12.92;
  bl = bl > 0.04045 ? Math.pow((bl + 0.055) / 1.055, 2.4) : bl / 12.92;
  let x = (rl * 0.4124564 + gl * 0.3575761 + bl * 0.1804375) / 0.95047;
  let y = (rl * 0.2126729 + gl * 0.7151522 + bl * 0.0721750);
  let z = (rl * 0.0193339 + gl * 0.1191920 + bl * 0.9503041) / 1.08883;
  const f = t => t > 0.008856 ? Math.cbrt(t) : 7.787 * t + 16 / 116;
  x = f(x); y = f(y); z = f(z);
  return [116 * y - 16, 500 * (x - y), 200 * (y - z)];
}

function labDistance(a, b) {
  return Math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2);
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
 * Classify an image as graphic (few distinct colors, recolorable) vs photo
 * (continuous color distribution, not recolorable).
 *
 * Uses two metrics that together give robust separation:
 * 1. Entropy of quantized color histogram — graphics < 5.0, photos >= 5.0
 * 2. Colors for 85% coverage — graphics need <= 30 quantized bins,
 *    photos need 39+
 *
 * Returns dominant chromatic colors for graphics, empty array for photos.
 */
function findDominantColors(imageData) {
  const { data, width, height } = imageData;
  const totalPixels = width * height;

  const q16Map = new Map();
  let opaquePixels = 0;

  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] < 30) continue;
    opaquePixels++;
    const r = Math.round(data[i] / QUANT_STEP) * QUANT_STEP;
    const g = Math.round(data[i + 1] / QUANT_STEP) * QUANT_STEP;
    const b = Math.round(data[i + 2] / QUANT_STEP) * QUANT_STEP;
    const key = (r << 16) | (g << 8) | b;
    q16Map.set(key, (q16Map.get(key) || 0) + 1);
  }

  if (opaquePixels < totalPixels * 0.01) return [];

  let entropy = 0;
  for (const count of q16Map.values()) {
    const p = count / opaquePixels;
    if (p > 0) entropy -= p * Math.log2(p);
  }

  const q16Sorted = [...q16Map.entries()]
    .map(([key, count]) => ({
      r: (key >> 16) & 0xFF,
      g: (key >> 8) & 0xFF,
      b: key & 0xFF,
      count,
    }))
    .sort((a, b) => b.count - a.count);

  let cumulative = 0;
  let colorsFor85 = 0;
  for (let i = 0; i < q16Sorted.length; i++) {
    cumulative += q16Sorted[i].count;
    if (cumulative >= opaquePixels * 0.85) { colorsFor85 = i + 1; break; }
  }

  if (entropy >= ENTROPY_THRESHOLD || colorsFor85 > COVERAGE85_THRESHOLD) {
    return [];
  }

  const entries = q16Sorted.map(e => ({
    ...e,
    lab: rgbToLab(e.r, e.g, e.b),
    rSum: e.r * e.count,
    gSum: e.g * e.count,
    bSum: e.b * e.count,
  }));

  const clusters = [];
  for (const entry of entries) {
    let merged = false;
    for (const cl of clusters) {
      if (labDistance(entry.lab, cl.lab) < LAB_MERGE_DIST) {
        cl.rSum += entry.r * entry.count;
        cl.gSum += entry.g * entry.count;
        cl.bSum += entry.b * entry.count;
        cl.count += entry.count;
        const total = cl.count;
        const avgR = Math.round(cl.rSum / total);
        const avgG = Math.round(cl.gSum / total);
        const avgB = Math.round(cl.bSum / total);
        cl.lab = rgbToLab(avgR, avgG, avgB);
        merged = true;
        break;
      }
    }
    if (!merged) {
      clusters.push({
        count: entry.count,
        lab: entry.lab,
        rSum: entry.rSum,
        gSum: entry.gSum,
        bSum: entry.bSum,
      });
    }
  }

  clusters.sort((a, b) => b.count - a.count);

  const minPixels = opaquePixels * MIN_CLUSTER_PIXEL_PCT / 100;
  const results = [];

  for (const cl of clusters) {
    if (cl.count < minPixels) continue;
    const r = Math.round(cl.rSum / cl.count);
    const g = Math.round(cl.gSum / cl.count);
    const b = Math.round(cl.bSum / cl.count);
    const [, s] = rgbToHsv(r, g, b);
    if (s < CHROMATIC_SAT_MIN) continue;

    results.push({
      hex: rgbToHex(r, g, b),
      pixelCount: cl.count,
    });
  }

  return results;
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

const SVG_SATURATION_THRESHOLD = 10;

function findSvgColors(svgContent) {
  const hexPattern = /#([0-9a-fA-F]{6})\b/g;
  const colorCounts = new Map();

  let match;
  while ((match = hexPattern.exec(svgContent)) !== null) {
    const hex = match[1].toUpperCase();
    const [r, g, b] = hexToRgb(hex);
    const [, s] = rgbToHsv(r, g, b);
    if (s < SVG_SATURATION_THRESHOLD) continue;
    colorCounts.set(hex, (colorCounts.get(hex) || 0) + 1);
  }

  return [...colorCounts.entries()].map(([hex, count]) => ({ hex, count }));
}

/**
 * Scan all SVG files in ppt/media/ and extract hex colors.
 * Returns a Map of svgPath -> [{ hex, count }]
 */
export async function extractSvgColors(zipOrBuffer) {
  const zip = zipOrBuffer instanceof JSZip
    ? zipOrBuffer
    : await JSZip.loadAsync(zipOrBuffer);

  const svgFiles = zip.file(/ppt\/media\/.*\.svg$/i);
  const svgColorMap = new Map();

  for (const file of svgFiles) {
    try {
      const content = await file.async('string');
      const colors = findSvgColors(content);
      if (colors.length > 0) {
        svgColorMap.set(file.name, colors);
      }
    } catch {
      // Skip SVGs that fail to read
    }
  }

  return svgColorMap;
}

/**
 * Replace hex colors inside an SVG string.
 * @param {string} svgContent - Original SVG markup
 * @param {Map<string, string>} colorReplacements - origHex -> newHex (uppercase, no #)
 * @returns {string|null} Modified SVG, or null if unchanged
 */
export function recolorSvg(svgContent, colorReplacements) {
  let modified = svgContent;
  let changed = false;

  for (const [origHex, newHex] of colorReplacements) {
    const pattern = new RegExp(`#${origHex}`, 'gi');
    const result = modified.replace(pattern, `#${newHex}`);
    if (result !== modified) {
      modified = result;
      changed = true;
    }
  }

  return changed ? modified : null;
}

export { rgbToHsv, hsvToRgb, hexToRgb, rgbToHex };
