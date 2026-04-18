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

async function resolveThemeColors(zip) {
  const allThemeColors = {};

  const themeFiles = zip.file(/ppt\/theme\/theme\d*\.xml/);
  const themeColorsByFile = {};
  for (const f of themeFiles) {
    themeColorsByFile[f.name] = parseThemeColors(await f.async('string'));
  }

  const masterRelsFiles = zip.file(/ppt\/slideMasters\/_rels\/slideMaster\d+\.xml\.rels$/);
  const mastersProcessed = new Set();

  for (const relsFile of masterRelsFiles) {
    const relsXml = await relsFile.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(relsXml, 'application/xml');

    for (const rel of doc.getElementsByTagName('Relationship')) {
      const type = rel.getAttribute('Type') || '';
      if (type.includes('/theme')) {
        let target = rel.getAttribute('Target') || '';
        if (target.startsWith('..')) {
          target = 'ppt' + target.substring(2);
        } else if (!target.startsWith('ppt/')) {
          target = 'ppt/slideMasters/' + target;
        }

        if (themeColorsByFile[target] && !mastersProcessed.has(target)) {
          mastersProcessed.add(target);
          const colors = themeColorsByFile[target];
          for (const [name, hex] of Object.entries(colors)) {
            if (!allThemeColors[name]) {
              allThemeColors[name] = hex;
            }
          }
        }
      }
    }
  }

  if (Object.keys(allThemeColors).length === 0) {
    const sorted = Object.keys(themeColorsByFile).sort();
    for (const file of sorted) {
      for (const [name, hex] of Object.entries(themeColorsByFile[file])) {
        if (!allThemeColors[name]) {
          allThemeColors[name] = hex;
        }
      }
    }
  }

  return allThemeColors;
}

export async function buildMediaSlideMap(zipOrBuffer) {
  const zip = zipOrBuffer instanceof JSZip
    ? zipOrBuffer
    : await JSZip.loadAsync(zipOrBuffer);

  const mediaSlideMap = new Map();
  const relsFiles = zip.file(/ppt\/slides\/_rels\/slide\d+\.xml\.rels$/);

  for (const relsFile of relsFiles) {
    const match = relsFile.name.match(/slide(\d+)\.xml\.rels$/);
    if (!match) continue;
    const slideNum = parseInt(match[1], 10);

    const xml = await relsFile.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    for (const rel of doc.getElementsByTagName('Relationship')) {
      const type = rel.getAttribute('Type') || '';
      if (!type.includes('/image')) continue;

      let target = rel.getAttribute('Target') || '';
      if (target.startsWith('..')) {
        target = 'ppt' + target.substring(2);
      } else if (!target.startsWith('ppt/')) {
        target = 'ppt/slides/' + target;
      }

      if (!mediaSlideMap.has(target)) {
        mediaSlideMap.set(target, []);
      }
      const slides = mediaSlideMap.get(target);
      if (!slides.includes(slideNum)) {
        slides.push(slideNum);
      }
    }
  }

  for (const slides of mediaSlideMap.values()) {
    slides.sort((a, b) => a - b);
  }

  return mediaSlideMap;
}

export async function extractColors(zipOrBuffer) {
  const zip = zipOrBuffer instanceof JSZip
    ? zipOrBuffer
    : await JSZip.loadAsync(zipOrBuffer);

  const themeColorMap = await resolveThemeColors(zip);

  const directColors = new Map();
  const themeColorUsage = new Map();

  const xmlFiles = zip.file(/ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/);

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
 * Build a unified color list for the UI.
 *
 * Direct, theme, and SVG colors with the same hex are merged into one entry.
 * Image palette colors are always separate entries (per group, per color).
 */
export function buildColorList(directColors, themeColorUsage, imageGroups, svgColorMap, mediaSlideMap) {
  const list = [];

  // Collect all SVG hex → {totalCount, paths, slides}
  const svgHexInfo = new Map();
  if (svgColorMap && svgColorMap.size > 0) {
    for (const [path, colors] of svgColorMap) {
      const pathSlides = mediaSlideMap?.get(path) || [];
      for (const { hex, count } of colors) {
        if (!svgHexInfo.has(hex)) {
          svgHexInfo.set(hex, { totalCount: 0, paths: [], slides: new Set() });
        }
        const info = svgHexInfo.get(hex);
        info.totalCount += count;
        if (!info.paths.includes(path)) info.paths.push(path);
        for (const s of pathSlides) info.slides.add(s);
      }
    }
  }

  // Build unified entries: merge direct + theme + SVG by hex
  const unifiedByHex = new Map();

  for (const [hex, count] of directColors) {
    if (!unifiedByHex.has(hex)) {
      unifiedByHex.set(hex, { hex, count: 0, sources: [] });
    }
    const entry = unifiedByHex.get(hex);
    entry.count += count;
    entry.sources.push('direct');
  }

  for (const [name, { hex, count }] of themeColorUsage) {
    if (!unifiedByHex.has(hex)) {
      unifiedByHex.set(hex, { hex, count: 0, sources: [] });
    }
    const entry = unifiedByHex.get(hex);
    entry.count += count;
    entry.sources.push('theme');
    entry.themeName = name;
    entry.themeLabel = THEME_COLOR_NAMES[name] || name;
  }

  for (const [hex, info] of svgHexInfo) {
    if (!unifiedByHex.has(hex)) {
      unifiedByHex.set(hex, { hex, count: 0, sources: [] });
    }
    const entry = unifiedByHex.get(hex);
    entry.count += info.totalCount;
    entry.sources.push('svg');
    entry.svgPaths = info.paths;
    entry.svgSlides = [...info.slides].sort((a, b) => a - b);
  }

  for (const [hex, data] of unifiedByHex) {
    const primaryType = data.sources.includes('theme') ? 'theme'
      : data.sources.includes('direct') ? 'direct'
      : 'svg';

    list.push({
      id: `color-${hex}`,
      hex,
      count: data.count,
      type: primaryType,
      sources: data.sources,
      themeName: data.themeName || null,
      themeLabel: data.themeLabel || null,
      svgPaths: data.svgPaths || null,
      slides: data.svgSlides || null,
    });
  }

  // Image entries: per group, per palette color (separate from unified)
  if (imageGroups && imageGroups.length > 0) {
    for (let gi = 0; gi < imageGroups.length; gi++) {
      const group = imageGroups[gi];
      const groupId = `imggrp-${gi}`;
      const slides = resolveGroupSlides(group.paths, mediaSlideMap);

      for (const color of group.colors) {
        list.push({
          id: `${groupId}-${color.hex}`,
          hex: color.hex,
          count: group.imageCount,
          type: 'image',
          kind: color.kind,
          groupId,
          groupIndex: gi,
          imagePaths: group.paths,
          imageCount: group.imageCount,
          thumbnail: group.thumbnail || null,
          groupColors: group.colors.map(c => ({ hex: c.hex, kind: c.kind })),
          slides,
        });
      }
    }
  }

  list.sort((a, b) => b.count - a.count);
  return list;
}

function resolveGroupSlides(paths, mediaSlideMap) {
  if (!mediaSlideMap) return [];
  const slides = new Set();
  for (const path of paths) {
    for (const s of mediaSlideMap.get(path) || []) slides.add(s);
  }
  return [...slides].sort((a, b) => a - b);
}
