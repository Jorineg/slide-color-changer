import JSZip from 'jszip';

const THEME_COLOR_NAMES = {
  dk1: 'Dark 1',
  dk2: 'Dark 2',
  lt1: 'Light 1',
  lt2: 'Light 2',
  accent1: 'Accent 1',
  accent2: 'Accent 2',
  accent3: 'Accent 3',
  accent4: 'Accent 4',
  accent5: 'Accent 5',
  accent6: 'Accent 6',
  hlink: 'Hyperlink',
  folHlink: 'Followed Link',
};

/**
 * Parse the theme XML to extract theme color definitions.
 * Returns a map of schemeClr name -> hex color (uppercase, no #).
 */
function parseThemeColors(themeXml) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(themeXml, 'application/xml');

  const colorScheme = doc.getElementsByTagName('a:clrScheme')[0];
  if (!colorScheme) return {};

  const colors = {};
  for (const child of colorScheme.children) {
    const tag = child.localName;
    const srgb = child.getElementsByTagName('a:srgbClr')[0];
    const sys = child.getElementsByTagName('a:sysClr')[0];

    if (srgb) {
      colors[tag] = srgb.getAttribute('val').toUpperCase();
    } else if (sys) {
      const lastClr = sys.getAttribute('lastClr');
      if (lastClr) colors[tag] = lastClr.toUpperCase();
    }
  }
  return colors;
}

/**
 * Walk all XML files in the PPTX looking for color references.
 * Returns { directColors: Map<hex, count>, themeColors: Map<schemeName, { hex, count }> }
 */
export async function extractColors(zipOrBuffer) {
  const zip = zipOrBuffer instanceof JSZip
    ? zipOrBuffer
    : await JSZip.loadAsync(zipOrBuffer);

  let themeColorMap = {};
  const themeFile = zip.file(/ppt\/theme\/theme\d*\.xml/);
  if (themeFile.length > 0) {
    const xml = await themeFile[0].async('string');
    themeColorMap = parseThemeColors(xml);
  }

  const directColors = new Map();
  const themeColorUsage = new Map();

  const xmlFiles = zip.file(/ppt\/(slides|slideLayouts|slideMasters)\/.*\.xml$/);

  for (const file of xmlFiles) {
    const xml = await file.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    const srgbElements = doc.getElementsByTagName('a:srgbClr');
    for (const el of srgbElements) {
      const val = el.getAttribute('val')?.toUpperCase();
      if (val && val.length === 6 && /^[0-9A-F]{6}$/.test(val)) {
        directColors.set(val, (directColors.get(val) || 0) + 1);
      }
    }

    const schemeElements = doc.getElementsByTagName('a:schemeClr');
    for (const el of schemeElements) {
      const val = el.getAttribute('val');
      if (val && themeColorMap[val]) {
        const existing = themeColorUsage.get(val);
        if (existing) {
          existing.count++;
        } else {
          themeColorUsage.set(val, { hex: themeColorMap[val], count: 1 });
        }
      }
    }
  }

  return { directColors, themeColorUsage, themeColorMap };
}

/**
 * Build a unified sorted color list for the UI.
 * Each entry: { id, hex, count, type: 'theme'|'direct', themeName?, themeLabel? }
 */
export function buildColorList(directColors, themeColorUsage) {
  const list = [];

  for (const [hex, count] of directColors) {
    const isDuplicate = [...themeColorUsage.values()].some(t => t.hex === hex);
    if (!isDuplicate) {
      list.push({ id: `direct-${hex}`, hex, count, type: 'direct' });
    }
  }

  for (const [name, { hex, count }] of themeColorUsage) {
    const directCount = directColors.get(hex) || 0;
    list.push({
      id: `theme-${name}`,
      hex,
      count: count + directCount,
      type: 'theme',
      themeName: name,
      themeLabel: THEME_COLOR_NAMES[name] || name,
    });
  }

  list.sort((a, b) => b.count - a.count);
  return list;
}
