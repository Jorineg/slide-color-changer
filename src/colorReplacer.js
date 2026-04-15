import JSZip from 'jszip';

/**
 * Apply color replacements to the PPTX and return a new Blob.
 *
 * @param {ArrayBuffer} originalBuffer - The original .pptx file
 * @param {Map<string, string>} colorMap - Map of original hex -> new hex (both uppercase, no #)
 * @param {Map<string, string>} themeNameToOrigHex - Map of theme name -> original hex
 * @returns {Promise<Blob>} Modified .pptx as a Blob
 */
export async function replaceColors(originalBuffer, colorMap, themeNameToOrigHex) {
  const zip = await JSZip.loadAsync(originalBuffer);

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

  const slideFiles = zip.file(/ppt\/(slides|slideLayouts|slideMasters)\/.*\.xml$/);
  for (const file of slideFiles) {
    let xml = await file.async('string');
    let modified = false;

    if (directReplacements.size > 0) {
      const replaced = replaceDirectColors(xml, directReplacements);
      if (replaced !== xml) {
        xml = replaced;
        modified = true;
      }
    }

    if (themeReplacements.size > 0) {
      const replaced = replaceSchemeColorsWithDirect(xml, themeReplacements);
      if (replaced !== xml) {
        xml = replaced;
        modified = true;
      }
    }

    if (modified) {
      zip.file(file.name, xml);
    }
  }

  if (themeReplacements.size > 0) {
    const themeFiles = zip.file(/ppt\/theme\/theme\d*\.xml/);
    for (const file of themeFiles) {
      let xml = await file.async('string');
      const replaced = replaceThemeDefinitions(xml, themeReplacements);
      if (replaced !== xml) {
        zip.file(file.name, replaced);
      }
    }
  }

  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}

function replaceDirectColors(xml, replacements) {
  for (const [orig, replacement] of replacements) {
    const pattern = new RegExp(`(<a:srgbClr\\s+val=")${orig}(")`, 'gi');
    xml = xml.replace(pattern, `$1${replacement}$2`);
  }
  return xml;
}

/**
 * For theme colors that are being remapped, convert schemeClr references
 * to direct srgbClr with the new color. This ensures the visual change
 * takes effect even if the theme is also used elsewhere.
 */
function replaceSchemeColorsWithDirect(xml, themeReplacements) {
  for (const [themeName, newHex] of themeReplacements) {
    const pattern = new RegExp(
      `<a:schemeClr\\s+val="${themeName}"\\s*/>`,
      'g'
    );
    xml = xml.replace(pattern, `<a:srgbClr val="${newHex}"/>`);

    const patternWithChildren = new RegExp(
      `<a:schemeClr\\s+val="${themeName}"\\s*>(.*?)</a:schemeClr>`,
      'gs'
    );
    xml = xml.replace(patternWithChildren, `<a:srgbClr val="${newHex}"/>`);
  }
  return xml;
}

function replaceThemeDefinitions(xml, themeReplacements) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  const colorScheme = doc.getElementsByTagName('a:clrScheme')[0];
  if (!colorScheme) return xml;

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
  }

  const serializer = new XMLSerializer();
  return serializer.serializeToString(doc);
}

/**
 * Build a modified PPTX as ArrayBuffer for re-rendering preview.
 */
export async function buildModifiedBuffer(originalBuffer, colorMap, themeNameToOrigHex) {
  const zip = await JSZip.loadAsync(originalBuffer);

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

  if (directReplacements.size === 0 && themeReplacements.size === 0) {
    return originalBuffer;
  }

  const slideFiles = zip.file(/ppt\/(slides|slideLayouts|slideMasters)\/.*\.xml$/);
  for (const file of slideFiles) {
    let xml = await file.async('string');
    let modified = false;

    if (directReplacements.size > 0) {
      const replaced = replaceDirectColors(xml, directReplacements);
      if (replaced !== xml) { xml = replaced; modified = true; }
    }
    if (themeReplacements.size > 0) {
      const replaced = replaceSchemeColorsWithDirect(xml, themeReplacements);
      if (replaced !== xml) { xml = replaced; modified = true; }
    }

    if (modified) zip.file(file.name, xml);
  }

  if (themeReplacements.size > 0) {
    const themeFiles = zip.file(/ppt\/theme\/theme\d*\.xml/);
    for (const file of themeFiles) {
      let xml = await file.async('string');
      const replaced = replaceThemeDefinitions(xml, themeReplacements);
      if (replaced !== xml) zip.file(file.name, replaced);
    }
  }

  return zip.generateAsync({ type: 'arraybuffer' });
}
