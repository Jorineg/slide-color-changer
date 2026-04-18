import JSZip from 'jszip';
import { recolorImage, recolorSvg } from './imageColors.js';

/**
 * Build structured replacement data from the UI's colorMap (id → newHex).
 *
 * @param {Array} colorList - full color list with type info
 * @param {Map<string, string>} colorMap - entry.id → newHex
 * @param {Map<string, string>} themeNameToOrigHex - themeName → origHex
 * @param {Map} svgColorMap - path → [{hex, count}]
 * @param {Array} imageGroups - [{paths, colors, ...}]
 * @param {Map} imageAnalysis - path → {palette: [{hex, r, g, b}]}
 */
export function buildReplacementPlan(colorList, colorMap, themeNameToOrigHex, svgColorMap, imageGroups, imageAnalysis) {
  const xmlReplacements = new Map();
  const themeReplacements = new Map();
  const svgReplacements = new Map();
  const imageGroupPlans = [];

  for (const entry of colorList) {
    const newHex = colorMap.get(entry.id);
    if (!newHex || newHex === entry.hex) continue;

    if (entry.type === 'image') {
      continue; // handled below per-group
    }

    // Unified entry (direct/theme/SVG): apply to XML and SVGs where the hex appears
    xmlReplacements.set(entry.hex, newHex);

    if (entry.sources?.includes('svg')) {
      svgReplacements.set(entry.hex, newHex);
    }
  }

  // Theme replacements: derived from xmlReplacements
  for (const [themeName, origHex] of themeNameToOrigHex) {
    const newHex = xmlReplacements.get(origHex);
    if (newHex) {
      themeReplacements.set(themeName, newHex);
    }
  }

  // Image group replacements
  if (imageGroups && imageAnalysis) {
    const groupReplacements = new Map();

    for (const entry of colorList) {
      if (entry.type !== 'image') continue;
      const newHex = colorMap.get(entry.id);
      if (!newHex || newHex === entry.hex) continue;

      if (!groupReplacements.has(entry.groupIndex)) {
        groupReplacements.set(entry.groupIndex, new Map());
      }
      groupReplacements.get(entry.groupIndex).set(entry.hex, newHex);
    }

    for (const [gi, replacements] of groupReplacements) {
      const group = imageGroups[gi];
      if (!group) continue;

      imageGroupPlans.push({
        paths: group.paths,
        replacements,
        getAnalysis: (path) => imageAnalysis.get(path),
      });
    }
  }

  return { xmlReplacements, themeReplacements, svgReplacements, imageGroupPlans };
}

/**
 * Apply color replacements to the PPTX and return a new Blob.
 */
export async function replaceColors(originalBuffer, plan, svgColorMap) {
  const zip = await JSZip.loadAsync(originalBuffer);

  await applyToSlideFiles(zip, plan.xmlReplacements, plan.themeReplacements);
  await applyToThemeFiles(zip, plan.themeReplacements);
  await applyToSvgFiles(zip, plan.svgReplacements, svgColorMap);
  await applyToSvgFallbackPngs(zip, plan);
  await applyToImageGroups(zip, plan.imageGroupPlans);

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
}

/**
 * Build a modified PPTX as ArrayBuffer for re-rendering preview.
 */
export async function buildModifiedBuffer(originalBuffer, plan, svgColorMap) {
  const hasXml = plan.xmlReplacements.size > 0 || plan.themeReplacements.size > 0;
  const hasSvg = plan.svgReplacements.size > 0;
  const hasImages = plan.imageGroupPlans.length > 0;

  if (!hasXml && !hasSvg && !hasImages) {
    return originalBuffer;
  }

  const zip = await JSZip.loadAsync(originalBuffer);
  await applyToSlideFiles(zip, plan.xmlReplacements, plan.themeReplacements);
  await applyToThemeFiles(zip, plan.themeReplacements);
  await applyToSvgFiles(zip, plan.svgReplacements, svgColorMap);
  await applyToSvgFallbackPngs(zip, plan);
  await applyToImageGroups(zip, plan.imageGroupPlans);

  return zip.generateAsync({ type: 'arraybuffer' });
}

// --- XML replacement (exact match) ---

async function applyToSlideFiles(zip, xmlReplacements, themeReplacements) {
  if (xmlReplacements.size === 0 && themeReplacements.size === 0) return;

  const slideFiles = zip.file(/ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/);

  for (const file of slideFiles) {
    let xml = await file.async('string');
    let modified = false;

    if (xmlReplacements.size > 0) {
      const replaced = replaceDirectColors(xml, xmlReplacements);
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

function replaceSchemeColors(xml, themeReplacements) {
  for (const [themeName, newHex] of themeReplacements) {
    const selfClosing = new RegExp(
      `<a:schemeClr\\s+val="${themeName}"\\s*/>`,
      'g'
    );
    xml = xml.replace(selfClosing, `<a:srgbClr val="${newHex}"/>`);

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

// --- SVG replacement (exact hex match in SVG file content) ---

async function applyToSvgFiles(zip, svgReplacements, svgColorMap) {
  if (svgReplacements.size === 0 || !svgColorMap) return;

  for (const [path, colors] of svgColorMap) {
    const pathReplacements = new Map();
    for (const { hex } of colors) {
      const newHex = svgReplacements.get(hex);
      if (newHex) pathReplacements.set(hex, newHex);
    }
    if (pathReplacements.size === 0) continue;

    const file = zip.file(path);
    if (!file) continue;

    const svgContent = await file.async('string');
    const recolored = recolorSvg(svgContent, pathReplacements);
    if (recolored) {
      zip.file(path, recolored);
    }
  }
}

// --- SVG fallback PNGs ---

async function applyToSvgFallbackPngs(zip, plan) {
  if (plan.svgReplacements.size === 0) return;

  const svgToPng = await buildSvgToPngMap(zip);

  for (const [svgPath, pngPath] of svgToPng) {
    const file = zip.file(pngPath);
    if (!file) continue;

    const bytes = await file.async('uint8array');
    const ext = pngPath.split('.').pop().toLowerCase();
    const mime = ext === 'png' ? 'image/png'
      : ext === 'gif' ? 'image/gif'
      : 'image/jpeg';

    const recolored = await recolorImage(bytes, mime, [], plan.svgReplacements);
    if (recolored) {
      zip.file(pngPath, recolored);
    }
  }
}

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
      const pngPath = idToTarget.get(m[1]);
      const svgPath = idToTarget.get(m[2]);
      if (pngPath && svgPath && !svgToPng.has(svgPath)) {
        svgToPng.set(svgPath, pngPath);
      }
    }
  }

  return svgToPng;
}

// --- Image group replacement (palette-delta) ---

async function applyToImageGroups(zip, imageGroupPlans) {
  if (!imageGroupPlans || imageGroupPlans.length === 0) return;

  for (const plan of imageGroupPlans) {
    for (const path of plan.paths) {
      const analysis = plan.getAnalysis(path);
      if (!analysis) continue;

      const file = zip.file(path);
      if (!file) continue;

      const bytes = await file.async('uint8array');
      const ext = path.split('.').pop().toLowerCase();
      const mime = ext === 'png' ? 'image/png'
        : ext === 'gif' ? 'image/gif'
        : 'image/jpeg';

      const recolored = await recolorImage(bytes, mime, analysis.palette, plan.replacements);
      if (recolored) {
        zip.file(path, recolored);
      }
    }
  }
}
