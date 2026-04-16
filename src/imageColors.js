import JSZip from 'jszip';

const QUANT_STEP = 16;
const LAB_MERGE_DIST = 15;
const ENTROPY_THRESHOLD = 5.0;
const COVERAGE85_THRESHOLD = 30;
const CHROMATIC_SAT_MIN = 12;
const MIN_CLUSTER_PIXEL_PCT = 0.3;
const CROSS_IMAGE_LAB_DIST = 8;
const BIDIR_GROUP_THRESHOLD = 0.5;
const THUMB_MAX_DIM = 64;

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

function classifyColor(r, g, b) {
  const [, s] = rgbToHsv(r, g, b);
  if (r > 230 && g > 230 && b > 230) return 'white';
  if (r < 30 && g < 30 && b < 30) return 'black';
  if (s < CHROMATIC_SAT_MIN) return 'gray';
  return 'chromatic';
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

function makeThumbnail(img) {
  const scale = Math.min(1, THUMB_MAX_DIM / Math.max(img.width, img.height));
  const w = Math.round(img.width * scale);
  const h = Math.round(img.height * scale);
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(img, 0, 0, w, h);
  return canvas.toDataURL('image/png');
}

/**
 * Analyze an image and return all dominant color clusters with metadata.
 * Returns { isPhoto, colors: [{hex, pixelCount, kind, lab}] }
 */
function analyzeImage(imageData) {
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

  if (opaquePixels < totalPixels * 0.01) {
    return { isPhoto: false, colors: [] };
  }

  let entropy = 0;
  for (const count of q16Map.values()) {
    const p = count / opaquePixels;
    if (p > 0) entropy -= p * Math.log2(p);
  }

  const q16Sorted = [...q16Map.entries()]
    .map(([key, count]) => ({
      r: (key >> 16) & 0xFF, g: (key >> 8) & 0xFF, b: key & 0xFF, count,
    }))
    .sort((a, b) => b.count - a.count);

  let cumulative = 0;
  let colorsFor85 = 0;
  for (let i = 0; i < q16Sorted.length; i++) {
    cumulative += q16Sorted[i].count;
    if (cumulative >= opaquePixels * 0.85) { colorsFor85 = i + 1; break; }
  }

  if (entropy >= ENTROPY_THRESHOLD || colorsFor85 > COVERAGE85_THRESHOLD) {
    return { isPhoto: true, colors: [] };
  }

  const entries = q16Sorted.map(e => ({
    ...e,
    lab: rgbToLab(e.r, e.g, e.b),
    rSum: e.r * e.count, gSum: e.g * e.count, bSum: e.b * e.count,
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
        const avgR = Math.round(cl.rSum / cl.count);
        const avgG = Math.round(cl.gSum / cl.count);
        const avgB = Math.round(cl.bSum / cl.count);
        cl.lab = rgbToLab(avgR, avgG, avgB);
        merged = true;
        break;
      }
    }
    if (!merged) {
      clusters.push({
        count: entry.count, lab: entry.lab,
        rSum: entry.rSum, gSum: entry.gSum, bSum: entry.bSum,
      });
    }
  }

  clusters.sort((a, b) => b.count - a.count);
  const minPixels = opaquePixels * MIN_CLUSTER_PIXEL_PCT / 100;

  const colors = [];
  for (const cl of clusters) {
    if (cl.count < minPixels) continue;
    const r = Math.round(cl.rSum / cl.count);
    const g = Math.round(cl.gSum / cl.count);
    const b = Math.round(cl.bSum / cl.count);
    colors.push({
      hex: rgbToHex(r, g, b),
      pixelCount: cl.count,
      kind: classifyColor(r, g, b),
      lab: rgbToLab(r, g, b),
    });
  }

  return { isPhoto: false, colors };
}

/**
 * Compute bidirectional chromatic match score between two color sets.
 * Returns min(A→B fraction, B→A fraction) considering only chromatic colors.
 * Score of 1.0 = identical chromatic palettes, 0.5 = one is subset of other
 * (with at most 2x size difference), 0 = no overlap.
 */
function bidirectionalChromaticScore(colorsA, colorsB) {
  const chrA = colorsA.filter(c => c.kind === 'chromatic');
  const chrB = colorsB.filter(c => c.kind === 'chromatic');
  if (chrA.length === 0 && chrB.length === 0) return 1.0;
  if (chrA.length === 0 || chrB.length === 0) return 0;

  let matchedA = 0;
  for (const ca of chrA) {
    for (const cb of chrB) {
      if (labDistance(ca.lab, cb.lab) < CROSS_IMAGE_LAB_DIST) { matchedA++; break; }
    }
  }
  let matchedB = 0;
  for (const cb of chrB) {
    for (const ca of chrA) {
      if (labDistance(ca.lab, cb.lab) < CROSS_IMAGE_LAB_DIST) { matchedB++; break; }
    }
  }

  return Math.min(matchedA / chrA.length, matchedB / chrB.length);
}

/**
 * Group graphic images by chromatic color similarity using bidirectional
 * matching with union-find.
 *
 * Two images are linked if the minimum of (fraction of A's chromatic colors
 * found in B, fraction of B's in A) >= BIDIR_GROUP_THRESHOLD.
 * This prevents a many-colored image (e.g. world map) from absorbing
 * single-color icons, since the map→icon direction scores near 0.
 *
 * @param {Map} analysisMap - path -> { colors, isPhoto }
 * @returns {Array<{paths: string[], colors: Array<{hex, kind, lab}>}>}
 */
function groupImages(analysisMap) {
  const chromatics = [];
  for (const [path, analysis] of analysisMap) {
    if (analysis.isPhoto) continue;
    const chr = analysis.colors.filter(c => c.kind === 'chromatic');
    if (chr.length === 0) continue;
    chromatics.push({ path, colors: analysis.colors });
  }

  const parent = new Map();
  for (const img of chromatics) parent.set(img.path, img.path);
  function find(x) {
    while (parent.get(x) !== x) {
      parent.set(x, parent.get(parent.get(x)));
      x = parent.get(x);
    }
    return x;
  }
  function union(a, b) {
    const ra = find(a), rb = find(b);
    if (ra !== rb) parent.set(ra, rb);
  }

  for (let i = 0; i < chromatics.length; i++) {
    for (let j = i + 1; j < chromatics.length; j++) {
      const score = bidirectionalChromaticScore(
        chromatics[i].colors,
        chromatics[j].colors,
      );
      if (score >= BIDIR_GROUP_THRESHOLD) {
        union(chromatics[i].path, chromatics[j].path);
      }
    }
  }

  const groupMap = new Map();
  for (const img of chromatics) {
    const root = find(img.path);
    if (!groupMap.has(root)) groupMap.set(root, []);
    groupMap.get(root).push(img);
  }

  const groups = [];
  for (const members of groupMap.values()) {
    const paths = members.map(m => m.path);

    const unionColors = [];
    for (const member of members) {
      for (const c of member.colors) {
        let found = false;
        for (const uc of unionColors) {
          if (labDistance(c.lab, uc.lab) < CROSS_IMAGE_LAB_DIST) {
            found = true;
            break;
          }
        }
        if (!found) unionColors.push({ ...c });
      }
    }

    groups.push({ paths, colors: unionColors });
  }

  return groups;
}

/**
 * Scan all raster images in ppt/media/ and return:
 * - imageColorMap: Map<path, [{hex, pixelCount}]> (chromatic only, for recoloring)
 * - imageGroups: Array of {paths, colors, thumbnail} (for UI display)
 */
export async function extractImageColors(zipOrBuffer) {
  const zip = zipOrBuffer instanceof JSZip
    ? zipOrBuffer
    : await JSZip.loadAsync(zipOrBuffer);

  const imageFiles = zip.file(/ppt\/media\/.*\.(png|jpg|jpeg|gif)$/i);
  const analysisMap = new Map();
  const thumbnails = new Map();

  for (const file of imageFiles) {
    try {
      const bytes = await file.async('uint8array');
      const ext = file.name.split('.').pop().toLowerCase();
      const mime = ext === 'png' ? 'image/png'
        : ext === 'gif' ? 'image/gif'
        : 'image/jpeg';

      const img = await loadImageFromBytes(bytes, mime);
      const imageData = getImagePixels(img);
      const analysis = analyzeImage(imageData);
      analysisMap.set(file.name, analysis);

      if (!analysis.isPhoto && analysis.colors.some(c => c.kind === 'chromatic')) {
        thumbnails.set(file.name, makeThumbnail(img));
      }
    } catch {
      // Skip images that fail to decode
    }
  }

  const imageColorMap = new Map();
  for (const [path, analysis] of analysisMap) {
    if (analysis.isPhoto) continue;
    const chromatic = analysis.colors
      .filter(c => c.kind === 'chromatic')
      .map(c => ({ hex: c.hex, pixelCount: c.pixelCount }));
    if (chromatic.length > 0) {
      imageColorMap.set(path, chromatic);
    }
  }

  const imageGroups = groupImages(analysisMap);
  for (const group of imageGroups) {
    group.thumbnail = thumbnails.get(group.paths[0]) || null;
    group.imageCount = group.paths.length;
  }

  return { imageColorMap, imageGroups };
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
