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

/**
 * Resolve the theme used by each slide master by following the .rels files.
 * Returns a merged color map from all themes referenced by slide masters,
 * with the first slide master's theme taking priority.
 */
async function resolveThemeColors(zip) {
  const allThemeColors = {};

  const themeFiles = zip.file(/ppt\/theme\/theme\d*\.xml/);
  const themeColorsByFile = {};
  for (const f of themeFiles) {
    themeColorsByFile[f.name] = parseThemeColors(await f.async('string'));
  }

  // Try to find slide master -> theme mapping via .rels files
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
        // Resolve relative path: ../theme/theme1.xml -> ppt/theme/theme1.xml
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

  // Fallback: if no masters resolved, merge all themes (theme1 preferred)
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

/**
 * Build a map from media file path (e.g. "ppt/media/image1.png") to an array
 * of 1-based slide numbers where the media appears.
 */
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

/**
 * Walk all XML files in the PPTX looking for color references.
 */
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
 * Build a unified sorted color list for the UI.
 *
 * @param {Map} directColors
 * @param {Map} themeColorUsage
 * @param {Array} [imageGroups] - Optional array from extractImageColors()
 * @param {Map} [svgColorMap] - Optional map from extractSvgColors()
 * @param {Map} [mediaSlideMap] - Optional map from buildMediaSlideMap()
 */
export function buildColorList(directColors, themeColorUsage, imageGroups, svgColorMap, mediaSlideMap) {
  const list = [];

  const themeHexSet = new Set([...themeColorUsage.values()].map(t => t.hex));

  for (const [hex, count] of directColors) {
    if (!themeHexSet.has(hex)) {
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

  if (imageGroups && imageGroups.length > 0) {
    const existingHexes = new Set(list.map(e => e.hex));
    for (let gi = 0; gi < imageGroups.length; gi++) {
      const group = imageGroups[gi];
      const groupId = `imggrp-${gi}`;
      const slides = resolveGroupSlides(group.paths, mediaSlideMap);
      const chromaticColors = group.colors.filter(c => c.kind === 'chromatic');

      for (const color of chromaticColors) {
        if (existingHexes.has(color.hex)) continue;
        if (closeToAny(color.hex, existingHexes)) continue;

        list.push({
          id: `${groupId}-${color.hex}`,
          hex: color.hex,
          count: group.imageCount,
          type: 'image',
          groupId,
          imagePaths: group.paths,
          imageCount: group.imageCount,
          thumbnail: group.thumbnail || null,
          groupColors: chromaticColors.map(c => c.hex),
          slides,
        });
      }
    }
  }

  // Add SVG-derived colors
  if (svgColorMap && svgColorMap.size > 0) {
    const svgGroups = groupSvgColors(svgColorMap, list, mediaSlideMap);
    for (const group of svgGroups) {
      list.push({
        id: `svg-${group.hex}`,
        hex: group.hex,
        count: group.totalCount,
        type: 'svg',
        svgPaths: group.paths,
        slides: group.slides,
      });
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

function closeToAny(hex, hexSet) {
  for (const existing of hexSet) {
    if (hexColorDistance(hex, existing) < 40) return true;
  }
  return false;
}

/**
 * Group SVG colors, deduplicating against existing list entries and across files.
 */
function groupSvgColors(svgColorMap, existingList, mediaSlideMap) {
  const existingHexes = new Set(existingList.map(e => e.hex));
  const groups = new Map();

  for (const [path, colors] of svgColorMap) {
    const pathSlides = mediaSlideMap?.get(path) || [];
    for (const { hex, count } of colors) {
      if (existingHexes.has(hex)) continue;

      let matchedHex = null;
      for (const existing of existingHexes) {
        if (hexColorDistance(hex, existing) < 40) {
          matchedHex = existing;
          break;
        }
      }
      if (matchedHex) continue;

      let groupKey = hex;
      for (const [key] of groups) {
        if (hexColorDistance(hex, key) < 40) {
          groupKey = key;
          break;
        }
      }

      if (groups.has(groupKey)) {
        const g = groups.get(groupKey);
        if (!g.paths.includes(path)) g.paths.push(path);
        for (const s of pathSlides) {
          if (!g.slides.includes(s)) g.slides.push(s);
        }
        g.totalCount += count;
      } else {
        groups.set(groupKey, { hex: groupKey, paths: [path], totalCount: count, slides: [...pathSlides] });
      }
    }
  }

  for (const g of groups.values()) {
    g.slides.sort((a, b) => a - b);
  }

  return [...groups.values()];
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
