import JSZip from 'jszip';
import { recolorImage, recolorSvg } from './imageColors.js';

/**
 * Apply color replacements to the PPTX and return a new Blob.
 */
export async function replaceColors(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap, svgColorMap) {
  const zip = await JSZip.loadAsync(originalBuffer);
  const { directReplacements, themeReplacements } = buildReplacementMaps(colorMap, themeNameToOrigHex);
  const imageOnlyMap = buildTypeRestrictedMap(colorMap, imageColorMap);
  const svgOnlyMap = buildTypeRestrictedMap(colorMap, svgColorMap);

  await applyToSlideFiles(zip, directReplacements, themeReplacements);
  await applyToThemeFiles(zip, themeReplacements);
  await applyToImages(zip, imageOnlyMap, imageColorMap);
  await applyToSvgs(zip, svgOnlyMap, svgColorMap);
  await applyToSvgFallbackPngs(zip, svgOnlyMap, svgColorMap);

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
}

/**
 * Build a modified PPTX as ArrayBuffer for re-rendering preview.
 */
export async function buildModifiedBuffer(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap, svgColorMap) {
  const { directReplacements, themeReplacements } = buildReplacementMaps(colorMap, themeNameToOrigHex);
  const imageOnlyMap = buildTypeRestrictedMap(colorMap, imageColorMap);
  const svgOnlyMap = buildTypeRestrictedMap(colorMap, svgColorMap);
  const hasImageChanges = imageColorMap && hasActiveReplacements(imageOnlyMap, imageColorMap);
  const hasSvgChanges = svgColorMap && hasActiveReplacements(svgOnlyMap, svgColorMap);

  if (directReplacements.size === 0 && themeReplacements.size === 0 && !hasImageChanges && !hasSvgChanges) {
    return originalBuffer;
  }

  const zip = await JSZip.loadAsync(originalBuffer);
  await applyToSlideFiles(zip, directReplacements, themeReplacements);
  await applyToThemeFiles(zip, themeReplacements);
  await applyToImages(zip, imageOnlyMap, imageColorMap);
  await applyToSvgs(zip, svgOnlyMap, svgColorMap);
  await applyToSvgFallbackPngs(zip, svgOnlyMap, svgColorMap);

  return zip.generateAsync({ type: 'arraybuffer' });
}

function buildReplacementMaps(colorMap, themeNameToOrigHex) {
  const directReplacements = new Map();
  const themeReplacements = new Map();

  for (const [origHex, newHex] of colorMap) {
    if (origHex === newHex) continue;
    directReplacements.set(origHex, newHex);
  }

  for (const [themeName, origHex] of themeNameToOrigHex) {
    const newHex = colorMap.get(origHex);
    if (newHex && newHex !== origHex) {
      themeReplacements.set(themeName, newHex);
    }
  }

  return { directReplacements, themeReplacements };
}

async function applyToSlideFiles(zip, directReplacements, themeReplacements) {
  if (directReplacements.size === 0 && themeReplacements.size === 0) return;

  const slideFiles = zip.file(/ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/);

  for (const file of slideFiles) {
    let xml = await file.async('string');
    let modified = false;

    if (directReplacements.size > 0) {
      const replaced = replaceDirectColors(xml, directReplacements);
      if (replaced !== xml) { xml = replaced; modified = true; }
    }

    if (themeReplacements.size > 0) {
      const replaced = replaceSchemeColors(xml, themeReplacements);
      if (replaced !== xml) { xml = replaced; modified = true; }
    }

    if (modified) {
      zip.file(file.name, xml);
    }
  }
}

async function applyToThemeFiles(zip, themeReplacements) {
  if (themeReplacements.size === 0) return;

  const themeFiles = zip.file(/ppt\/theme\/theme\d*\.xml/);
  for (const file of themeFiles) {
    const xml = await file.async('string');
    const replaced = replaceThemeDefinitions(xml, themeReplacements);
    if (replaced !== xml) {
      zip.file(file.name, replaced);
    }
  }
}

function replaceDirectColors(xml, replacements) {
  for (const [orig, replacement] of replacements) {
    const pattern = new RegExp(`(<a:srgbClr\\s+val=")${orig}(")`, 'gi');
    xml = xml.replace(pattern, `$1${replacement}$2`);
  }
  return xml;
}

/**
 * Replace schemeClr references with srgbClr, preserving child modifier
 * elements (alpha, tint, shade, satMod, lumMod, etc).
 */
function replaceSchemeColors(xml, themeReplacements) {
  for (const [themeName, newHex] of themeReplacements) {
    // Self-closing: <a:schemeClr val="accent6"/>
    const selfClosing = new RegExp(
      `<a:schemeClr\\s+val="${themeName}"\\s*/>`,
      'g'
    );
    xml = xml.replace(selfClosing, `<a:srgbClr val="${newHex}"/>`);

    // With children: preserve child modifier elements
    const withChildren = new RegExp(
      `<a:schemeClr\\s+val="${themeName}"\\s*>(.*?)</a:schemeClr>`,
      'gs'
    );
    xml = xml.replace(withChildren, (_, children) => {
      return `<a:srgbClr val="${newHex}">${children}</a:srgbClr>`;
    });
  }
  return xml;
}

function replaceThemeDefinitions(xml, themeReplacements) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  const colorScheme = doc.getElementsByTagName('a:clrScheme')[0];
  if (!colorScheme) return xml;

  let changed = false;
  for (const [themeName, newHex] of themeReplacements) {
    const el = colorScheme.getElementsByTagName(`a:${themeName}`)[0];
    if (!el) continue;

    while (el.firstChild) el.removeChild(el.firstChild);
    const srgb = doc.createElementNS(
      'http://schemas.openxmlformats.org/drawingml/2006/main',
      'a:srgbClr'
    );
    srgb.setAttribute('val', newHex);
    el.appendChild(srgb);
    changed = true;
  }

  if (!changed) return xml;
  return new XMLSerializer().serializeToString(doc);
}

function hexColorDistance(hex1, hex2) {
  const r1 = parseInt(hex1.substring(0, 2), 16);
  const g1 = parseInt(hex1.substring(2, 4), 16);
  const b1 = parseInt(hex1.substring(4, 6), 16);
  const r2 = parseInt(hex2.substring(0, 2), 16);
  const g2 = parseInt(hex2.substring(2, 4), 16);
  const b2 = parseInt(hex2.substring(4, 6), 16);
  return Math.sqrt((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2);
}

const IMAGE_COLOR_MATCH_THRESHOLD = 45;

/**
 * Build a colorMap restricted to only colors that were extracted from a
 * specific media type (imageColorMap or svgColorMap). This prevents SVG
 * color changes from bleeding into raster image recoloring and vice versa.
 */
function buildTypeRestrictedMap(colorMap, mediaColorMap) {
  if (!mediaColorMap || mediaColorMap.size === 0) return new Map();

  const mediaHexes = new Set();
  for (const [, colors] of mediaColorMap) {
    for (const c of colors) mediaHexes.add(c.hex);
  }

  const restricted = new Map();
  for (const [origHex, newHex] of colorMap) {
    if (origHex === newHex) continue;
    if (mediaHexes.has(origHex)) {
      restricted.set(origHex, newHex);
      continue;
    }
    for (const mHex of mediaHexes) {
      if (hexColorDistance(origHex, mHex) < IMAGE_COLOR_MATCH_THRESHOLD) {
        restricted.set(origHex, newHex);
        break;
      }
    }
  }
  return restricted;
}

function findClosestChangedColor(rawHex, colorMap) {
  const exact = colorMap.get(rawHex);
  if (exact && exact !== rawHex) return exact;

  let bestTarget = null;
  let bestDist = IMAGE_COLOR_MATCH_THRESHOLD;
  for (const [origHex, newHex] of colorMap) {
    if (origHex === newHex) continue;
    const dist = hexColorDistance(rawHex, origHex);
    if (dist < bestDist) {
      bestDist = dist;
      bestTarget = newHex;
    }
  }
  return bestTarget;
}

function hasActiveReplacements(restrictedMap, mediaColorMap) {
  if (!mediaColorMap) return false;
  for (const [, colors] of mediaColorMap) {
    for (const { hex } of colors) {
      if (findClosestChangedColor(hex, restrictedMap)) return true;
    }
  }
  return false;
}

async function applyToImages(zip, colorMap, imageColorMap) {
  if (!imageColorMap || imageColorMap.size === 0) return;

  const imagesToProcess = [];

  for (const [path, colors] of imageColorMap) {
    const replacements = new Map();
    for (const { hex } of colors) {
      const target = findClosestChangedColor(hex, colorMap);
      if (target) {
        replacements.set(hex, target);
      }
    }
    if (replacements.size > 0) {
      imagesToProcess.push({ path, replacements });
    }
  }

  if (imagesToProcess.length === 0) return;

  for (const { path, replacements } of imagesToProcess) {
    const file = zip.file(path);
    if (!file) continue;

    const bytes = await file.async('uint8array');
    const ext = path.split('.').pop().toLowerCase();
    const mime = ext === 'png' ? 'image/png'
      : ext === 'gif' ? 'image/gif'
      : 'image/jpeg';

    const recolored = await recolorImage(bytes, mime, replacements);
    if (recolored) {
      zip.file(path, recolored);
    }
  }
}

async function applyToSvgs(zip, svgOnlyMap, svgColorMap) {
  if (!svgColorMap || svgColorMap.size === 0) return;

  for (const [path, colors] of svgColorMap) {
    const replacements = new Map();
    for (const { hex } of colors) {
      const target = findClosestChangedColor(hex, svgOnlyMap);
      if (target) {
        replacements.set(hex, target);
      }
    }
    if (replacements.size === 0) continue;

    const file = zip.file(path);
    if (!file) continue;

    const svgContent = await file.async('string');
    const recolored = recolorSvg(svgContent, replacements);
    if (recolored) {
      zip.file(path, recolored);
    }
  }
}

/**
 * PowerPoint stores SVGs with a PNG fallback. Many renderers (including
 * pptx-to-html and Apryse) display the PNG, not the SVG. When we recolor
 * an SVG, we also need to recolor its paired PNG fallback.
 *
 * The pairing is found in slide .rels files where a blip embeds a PNG
 * and has an asvg:svgBlip extension pointing to the SVG.
 */
async function applyToSvgFallbackPngs(zip, svgOnlyMap, svgColorMap) {
  if (!svgColorMap || svgColorMap.size === 0) return;

  const changedSvgPaths = new Set();
  for (const [path, colors] of svgColorMap) {
    for (const { hex } of colors) {
      if (findClosestChangedColor(hex, svgOnlyMap)) {
        changedSvgPaths.add(path);
        break;
      }
    }
  }
  if (changedSvgPaths.size === 0) return;

  const svgToPngFallback = await buildSvgToPngMap(zip);

  for (const svgPath of changedSvgPaths) {
    const pngPath = svgToPngFallback.get(svgPath);
    if (!pngPath) continue;

    const svgColors = svgColorMap.get(svgPath);
    if (!svgColors) continue;

    const replacements = new Map();
    for (const { hex } of svgColors) {
      const target = findClosestChangedColor(hex, svgOnlyMap);
      if (target) replacements.set(hex, target);
    }
    if (replacements.size === 0) continue;

    const file = zip.file(pngPath);
    if (!file) continue;

    const bytes = await file.async('uint8array');
    const ext = pngPath.split('.').pop().toLowerCase();
    const mime = ext === 'png' ? 'image/png'
      : ext === 'gif' ? 'image/gif'
      : 'image/jpeg';

    const recolored = await recolorImage(bytes, mime, replacements);
    if (recolored) {
      zip.file(pngPath, recolored);
    }
  }
}

/**
 * Scan slide .rels files for SVG+PNG blip pairs and return a map
 * of svgPath -> pngFallbackPath.
 */
async function buildSvgToPngMap(zip) {
  const svgToPng = new Map();
  const relsFiles = zip.file(/ppt\/slides\/_rels\/slide\d+\.xml\.rels$/);

  for (const relsFile of relsFiles) {
    const xml = await relsFile.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    const idToTarget = new Map();
    for (const rel of doc.getElementsByTagName('Relationship')) {
      const id = rel.getAttribute('Id');
      let target = rel.getAttribute('Target') || '';
      if (target.startsWith('..')) target = 'ppt' + target.substring(2);
      else if (!target.startsWith('ppt/')) target = 'ppt/slides/' + target;
      idToTarget.set(id, target);
    }

    const slideName = relsFile.name.replace('_rels/', '').replace('.rels', '');
    const slideFile = zip.file(slideName);
    if (!slideFile) continue;

    const slideXml = await slideFile.async('string');
    const svgBlipPattern = /r:embed="(rId\d+)"[^>]*>[\s\S]*?svgBlip[^>]*r:embed="(rId\d+)"/g;
    let m;
    while ((m = svgBlipPattern.exec(slideXml)) !== null) {
      const pngId = m[1];
      const svgId = m[2];
      const pngPath = idToTarget.get(pngId);
      const svgPath = idToTarget.get(svgId);
      if (pngPath && svgPath && !svgToPng.has(svgPath)) {
        svgToPng.set(svgPath, pngPath);
      }
    }
  }

  return svgToPng;
}
