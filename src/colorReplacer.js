import JSZip from 'jszip';
import { recolorImage } from './imageColors.js';

/**
 * Apply color replacements to the PPTX and return a new Blob.
 */
export async function replaceColors(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap) {
  const zip = await JSZip.loadAsync(originalBuffer);
  const { directReplacements, themeReplacements } = buildReplacementMaps(colorMap, themeNameToOrigHex);

  await applyToSlideFiles(zip, directReplacements, themeReplacements);
  await applyToThemeFiles(zip, themeReplacements);
  await applyToImages(zip, colorMap, imageColorMap);

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
}

/**
 * Build a modified PPTX as ArrayBuffer for re-rendering preview.
 */
export async function buildModifiedBuffer(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap) {
  const { directReplacements, themeReplacements } = buildReplacementMaps(colorMap, themeNameToOrigHex);
  const hasImageChanges = imageColorMap && hasActiveImageReplacements(colorMap, imageColorMap);

  if (directReplacements.size === 0 && themeReplacements.size === 0 && !hasImageChanges) {
    return originalBuffer;
  }

  const zip = await JSZip.loadAsync(originalBuffer);
  await applyToSlideFiles(zip, directReplacements, themeReplacements);
  await applyToThemeFiles(zip, themeReplacements);
  await applyToImages(zip, colorMap, imageColorMap);

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

function hasActiveImageReplacements(colorMap, imageColorMap) {
  if (!imageColorMap) return false;
  for (const [, colors] of imageColorMap) {
    for (const { hex } of colors) {
      const target = colorMap.get(hex);
      if (target && target !== hex) return true;
    }
  }
  return false;
}

async function applyToImages(zip, colorMap, imageColorMap) {
  if (!imageColorMap || imageColorMap.size === 0) return;

  // Build replacement map for image colors only
  const imageReplacements = new Map();
  for (const [, colors] of imageColorMap) {
    for (const { hex } of colors) {
      const target = colorMap.get(hex);
      if (target && target !== hex && !imageReplacements.has(hex)) {
        imageReplacements.set(hex, target);
      }
    }
  }
  if (imageReplacements.size === 0) return;

  // Find all image files that need recoloring
  const pathsToRecolor = new Set();
  for (const [path, colors] of imageColorMap) {
    for (const { hex } of colors) {
      if (imageReplacements.has(hex)) {
        pathsToRecolor.add(path);
        break;
      }
    }
  }

  for (const path of pathsToRecolor) {
    const file = zip.file(path);
    if (!file) continue;

    const bytes = await file.async('uint8array');
    const ext = path.split('.').pop().toLowerCase();
    const mime = ext === 'png' ? 'image/png'
      : ext === 'gif' ? 'image/gif'
      : 'image/jpeg';

    const recolored = await recolorImage(bytes, mime, imageReplacements);
    if (recolored) {
      zip.file(path, recolored);
    }
  }
}
