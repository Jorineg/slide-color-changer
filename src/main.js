import './style.css';
import { pptxToHtml } from '@jvmr/pptx-to-html';
import html2canvas from 'html2canvas';
import { extractColors, buildColorList, buildMediaSlideMap } from './colorExtractor.js';
import { extractImageColors, extractSvgColors } from './imageColors.js';
import { replaceColors, buildModifiedBuffer } from './colorReplacer.js';
import { openColorPicker, useEyeDropper } from './colorPicker.js';
import { loadApryseCore, isApryseLoaded, renderSlidesHD } from './apryseRenderer.js';

const RENDER_WIDTH = 960;
const RENDER_HEIGHT = 540;
const CANVAS_SCALE = 2;

const HD_PREF_KEY = 'slide-color-changer-hd-mode';

let originalBuffer = null;
let currentBuffer = null;
let fileName = '';
let colorList = [];
let colorMap = new Map();
let themeNameToOrigHex = new Map();
let imageColorMap = new Map();
let imageGroups = [];
let svgColorMap = new Map();
let mediaSlideMap = new Map();
let directColors = new Map();
let themeColorUsage = new Map();
let slideHtmls = [];
let slideCanvases = [];
let hdMode = localStorage.getItem(HD_PREF_KEY) === 'true';
let hdLoading = false;
let previewDebounceTimer = null;

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// --- Drop zone ---
const dropzone = $('#dropzone');
const fileInput = $('#file-input');

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') fileInput.click();
});

dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.classList.add('border-indigo-500', 'bg-gray-900');
});
dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('border-indigo-500', 'bg-gray-900');
});
dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('border-indigo-500', 'bg-gray-900');
  const file = e.dataTransfer.files[0];
  if (file && file.name.endsWith('.pptx')) {
    handleFile(file);
  }
});

fileInput.addEventListener('change', () => {
  const file = fileInput.files[0];
  if (file) handleFile(file);
});

// --- File handling ---
async function handleFile(file) {
  fileName = file.name;
  showLoading('Reading file...');

  try {
    originalBuffer = await file.arrayBuffer();

    showLoading('Extracting colors...');
    ({ directColors, themeColorUsage } = await extractColors(originalBuffer));

    showLoading('Scanning media...');
    ({ imageColorMap, imageGroups } = await extractImageColors(originalBuffer));
    svgColorMap = await extractSvgColors(originalBuffer);
    mediaSlideMap = await buildMediaSlideMap(originalBuffer);

    themeNameToOrigHex = new Map();
    for (const [name, { hex }] of themeColorUsage) {
      themeNameToOrigHex.set(name, hex);
    }

    rebuildColorList();

    currentBuffer = originalBuffer;

    if (hdMode) {
      slideHtmls = [];
      hideLoading();
      $('#file-name').textContent = fileName;
      renderColorTable();
      updateHdButton();
      triggerHdRender(currentBuffer);
    } else {
      showLoading('Rendering slides...');
      slideHtmls = await pptxToHtml(originalBuffer, {
        width: RENDER_WIDTH,
        height: RENDER_HEIGHT,
        scaleToFit: true,
      });

      hideLoading();
      updateHdButton();
      renderUI();
    }
  } catch (err) {
    console.error('Failed to process file:', err);
    hideLoading();
    dropzone.classList.remove('hidden');
    showError(`Failed to process file: ${err.message}`);
  }
}

function showError(msg) {
  let toast = $('#error-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'error-toast';
    toast.className = 'fixed bottom-6 left-1/2 -translate-x-1/2 bg-red-900/90 text-red-200 px-6 py-3 rounded-xl shadow-lg z-50 text-sm backdrop-blur';
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 5000);
}

// --- Loading ---
function showLoading(text) {
  $('#loading').classList.remove('hidden');
  $('#loading-text').textContent = text;
  $('#main-content').classList.add('hidden');
  dropzone.classList.add('hidden');
}

function hideLoading() {
  $('#loading').classList.add('hidden');
  dropzone.classList.add('hidden');
  $('#main-content').classList.remove('hidden');
}

function rebuildColorList() {
  colorList = buildColorList(
    directColors,
    themeColorUsage,
    imageGroups,
    svgColorMap,
    mediaSlideMap,
  );

  const oldMap = colorMap;
  colorMap = new Map();
  for (const entry of colorList) {
    colorMap.set(entry.hex, oldMap.get(entry.hex) || entry.hex);
  }
}

$('#chk-include-images').addEventListener('change', () => {
  renderColorTable();
});

// --- Render ---
function renderUI() {
  $('#file-name').textContent = fileName;
  const count = slideCanvases.length || slideHtmls.length;
  $('#slide-count').textContent = `${count} slide${count !== 1 ? 's' : ''}`;

  renderColorTable();
  renderSlides();
}

function renderColorTable() {
  const container = $('#color-table');
  container.innerHTML = '';

  const showImages = $('#chk-include-images')?.checked;
  const visibleList = showImages ? colorList : colorList.filter(e => e.type !== 'image');

  if (visibleList.length === 0) {
    container.innerHTML = '<div class="px-4 py-8 text-center text-gray-500 text-sm">No colors found in this presentation.</div>';
    return;
  }

  for (const entry of visibleList) {
    const currentHex = colorMap.get(entry.hex) || entry.hex;
    const isModified = currentHex !== entry.hex;

    const row = document.createElement('div');
    row.className = 'color-row grid grid-cols-[1.75rem_1fr_0.75rem_1.75rem_1fr] items-center gap-x-2 px-3 py-2';
    row.dataset.colorId = entry.id;
    row.dataset.origHex = entry.hex;

    row.innerHTML = `
      <div class="color-swatch original-swatch" style="background:#${entry.hex}" title="#${entry.hex}"></div>
      <div class="min-w-0">
        <div class="text-[0.65rem] font-mono text-gray-400 leading-tight truncate">#${entry.hex}</div>
        <div class="flex items-center gap-1">
          <span class="badge ${entry.type === 'theme' ? 'badge-theme' : entry.type === 'image' ? 'badge-image' : entry.type === 'svg' ? 'badge-svg' : 'badge-direct'}">
            ${entry.type === 'theme' ? entry.themeLabel : entry.type === 'image' ? 'Image' : entry.type === 'svg' ? 'SVG' : 'Direct'}
          </span>
          <span class="text-[0.6rem] text-gray-600">${entry.count}x</span>
          ${entry.slides && entry.slides.length > 0 ? `<span class="text-[0.55rem] text-gray-500" title="Appears on slide${entry.slides.length > 1 ? 's' : ''} ${entry.slides.join(', ')}">S${entry.slides.join(',')}</span>` : ''}
        </div>
      </div>
      <svg class="arrow-icon h-3 w-3 justify-self-center" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5">
        <path stroke-linecap="round" stroke-linejoin="round" d="M13 7l5 5m0 0l-5 5m5-5H6" />
      </svg>
      <div class="color-swatch target-swatch${isModified ? ' ring-2 ring-indigo-500' : ''}" style="background:#${currentHex}" title="#${currentHex}"></div>
      <div class="min-w-0">
        <div class="text-[0.65rem] font-mono text-gray-400 leading-tight truncate target-hex">#${currentHex}</div>
        <div class="flex items-center gap-1">
          ${isModified ? `<button class="reset-btn text-[0.6rem] text-indigo-400 hover:text-indigo-300 cursor-pointer leading-tight">Reset</button>` : ''}
          ${'EyeDropper' in window ? `<button class="eyedropper-btn text-gray-500 hover:text-gray-300 cursor-pointer leading-tight" title="Pick from screen">
            <svg class="h-3 w-3" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
              <path stroke-linecap="round" stroke-linejoin="round" d="M15.042 21.672L13.684 16.6m0 0l-2.51 2.225.569-9.47 5.227 7.917-3.286-.672zM12 2.25V4.5m5.834.166l-1.591 1.591M20.25 10.5H18M7.757 14.743l-1.59 1.59M6 10.5H3.75m4.007-4.243l-1.59-1.59" />
            </svg>
          </button>` : ''}
        </div>
      </div>
    `;

    const targetSwatch = row.querySelector('.target-swatch');
    targetSwatch.addEventListener('click', (e) => {
      e.stopPropagation();
      openColorPicker(targetSwatch, currentHex, (newHex) => {
        colorMap.set(entry.hex, newHex);
        renderColorTable();
        schedulePreviewUpdate();
      });
    });

    const resetBtn = row.querySelector('.reset-btn');
    if (resetBtn) {
      resetBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        colorMap.set(entry.hex, entry.hex);
        renderColorTable();
        schedulePreviewUpdate();
      });
    }

    const eyedropperBtn = row.querySelector('.eyedropper-btn');
    if (eyedropperBtn) {
      eyedropperBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const hex = await useEyeDropper();
        if (hex) {
          colorMap.set(entry.hex, hex);
          renderColorTable();
          schedulePreviewUpdate();
        }
      });
    }

    container.appendChild(row);
  }
}

const SLIDE_FONTS_CSS = `@import url('https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&family=Caladea:ital,wght@0,400;0,700;1,400;1,700&family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap');`;

const SLIDE_HARDENING_CSS = `
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    margin: 0;
    font-family: 'Carlito', 'Calibri', 'Lato', 'Arial', 'Helvetica Neue', sans-serif;
    line-height: 1.15;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
    text-rendering: optimizeLegibility;
    background: #ffffff;
    overflow: hidden;
  }
  img { object-fit: contain; image-rendering: auto; }
  span, p, div { word-break: break-word; overflow-wrap: break-word; }
  table { border-collapse: collapse; }
  svg { overflow: visible; }
`;

function createIsolatedFrame() {
  const iframe = document.createElement('iframe');
  iframe.style.cssText = `position:fixed;left:-9999px;top:0;width:${RENDER_WIDTH}px;height:${RENDER_HEIGHT}px;border:none;visibility:hidden;`;
  document.body.appendChild(iframe);
  return iframe;
}

async function rasterizeSlide(html) {
  const iframe = createIsolatedFrame();
  const doc = iframe.contentDocument;

  doc.open();
  doc.write(`<!DOCTYPE html><html><head>
    <style>${SLIDE_FONTS_CSS}</style>
    <style>${SLIDE_HARDENING_CSS}</style>
  </head><body>${html}</body></html>`);
  doc.close();

  await new Promise((r) => {
    if (iframe.contentWindow.document.fonts) {
      iframe.contentWindow.document.fonts.ready.then(r);
    } else {
      r();
    }
  });
  await new Promise((r) => setTimeout(r, 100));

  const canvas = await html2canvas(doc.body, {
    width: RENDER_WIDTH,
    height: RENDER_HEIGHT,
    scale: CANVAS_SCALE,
    useCORS: true,
    allowTaint: true,
    backgroundColor: '#ffffff',
    logging: false,
    windowWidth: RENDER_WIDTH,
    windowHeight: RENDER_HEIGHT,
  });

  document.body.removeChild(iframe);
  return canvas;
}

function renderSlides() {
  const container = $('#slides-container');
  container.innerHTML = '';
  slideCanvases = new Array(slideHtmls.length).fill(null);

  slideHtmls.forEach((html, i) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'slide-wrapper clickable-slide';
    wrapper.dataset.slideIndex = i;

    const label = document.createElement('div');
    label.className = 'slide-label';
    label.textContent = `Slide ${i + 1}`;

    const placeholder = document.createElement('div');
    placeholder.className = 'slide-placeholder';
    placeholder.innerHTML = '<div class="spinner"></div>';

    wrapper.appendChild(label);
    wrapper.appendChild(placeholder);

    wrapper.addEventListener('click', (e) => {
      handleSlideClick(e, i);
    });

    container.appendChild(wrapper);
  });

  rasterizeAllSlides();
}

async function rasterizeAllSlides() {
  for (let i = 0; i < slideHtmls.length; i++) {
    try {
      const canvas = await rasterizeSlide(slideHtmls[i]);
      slideCanvases[i] = canvas;

      const wrapper = $(`.slide-wrapper[data-slide-index="${i}"]`);
      if (!wrapper) continue;

      const placeholder = wrapper.querySelector('.slide-placeholder');
      if (placeholder) placeholder.remove();

      const img = document.createElement('img');
      img.className = 'slide-canvas-content';
      img.src = canvas.toDataURL('image/png');
      img.alt = `Slide ${i + 1}`;
      img.draggable = false;

      const label = wrapper.querySelector('.slide-label');
      if (label) {
        label.after(img);
      } else {
        wrapper.appendChild(img);
      }
    } catch (err) {
      console.error(`Failed to rasterize slide ${i + 1}:`, err);
      const wrapper = $(`.slide-wrapper[data-slide-index="${i}"]`);
      if (wrapper) {
        const placeholder = wrapper.querySelector('.slide-placeholder');
        if (placeholder) {
          placeholder.innerHTML = `<span class="text-gray-400 text-sm">Render failed</span>`;
        }
      }
    }
  }
}

// --- Slide click -> read pixel from canvas ---
function handleSlideClick(event, slideIndex) {
  const canvas = slideCanvases[slideIndex];
  if (!canvas) return;

  const img = event.currentTarget.querySelector('.slide-canvas-content');
  if (!img) return;

  const rect = img.getBoundingClientRect();
  const scaleX = canvas.width / rect.width;
  const scaleY = canvas.height / rect.height;
  const x = Math.round((event.clientX - rect.left) * scaleX);
  const y = Math.round((event.clientY - rect.top) * scaleY);

  const ctx = canvas.getContext('2d');
  const pixel = ctx.getImageData(x, y, 1, 1).data;
  const hex = [pixel[0], pixel[1], pixel[2]]
    .map((c) => c.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();

  if (hex.length === 6) {
    highlightColorRow(hex);
  }
}

function highlightColorRow(clickedHex) {
  $$('.color-row').forEach(row => row.classList.remove('highlighted'));

  let bestMatch = null;
  let bestDistance = Infinity;

  for (const entry of colorList) {
    const currentHex = colorMap.get(entry.hex) || entry.hex;
    const distOrig = colorDistance(clickedHex, entry.hex);
    const distCurrent = colorDistance(clickedHex, currentHex);
    const dist = Math.min(distOrig, distCurrent);
    if (dist < bestDistance) {
      bestDistance = dist;
      bestMatch = entry;
    }
  }

  if (bestMatch && bestDistance < 60) {
    const row = $(`.color-row[data-color-id="${bestMatch.id}"]`);
    if (row) {
      row.classList.add('highlighted');
      row.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }
}

function colorDistance(hex1, hex2) {
  if (hex1.length !== 6 || hex2.length !== 6) return Infinity;
  const r1 = parseInt(hex1.substring(0, 2), 16);
  const g1 = parseInt(hex1.substring(2, 4), 16);
  const b1 = parseInt(hex1.substring(4, 6), 16);
  const r2 = parseInt(hex2.substring(0, 2), 16);
  const g2 = parseInt(hex2.substring(2, 4), 16);
  const b2 = parseInt(hex2.substring(4, 6), 16);
  return Math.sqrt((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2);
}

// --- HD renderer ---
function updateHdButton() {
  const btn = $('#btn-hd-preview');
  const label = $('#hd-label');
  if (!btn) return;

  if (hdLoading) {
    btn.className = 'hd-toggle hd-loading px-3 py-1.5 text-sm rounded-lg transition-colors flex items-center gap-1.5';
    label.textContent = 'Loading HD...';
  } else if (hdMode) {
    btn.className = 'hd-toggle hd-active px-3 py-1.5 text-sm rounded-lg transition-colors flex items-center gap-1.5';
    label.textContent = 'HD On';
  } else {
    btn.className = 'hd-toggle px-3 py-1.5 text-sm rounded-lg transition-colors flex items-center gap-1.5';
    label.textContent = 'HD Preview';
  }
}

function createSlideplaceholders(count) {
  const container = $('#slides-container');
  container.innerHTML = '';
  slideCanvases = new Array(count).fill(null);

  $('#slide-count').textContent = `${count} slide${count !== 1 ? 's' : ''}`;

  for (let i = 0; i < count; i++) {
    const wrapper = document.createElement('div');
    wrapper.className = 'slide-wrapper clickable-slide';
    wrapper.dataset.slideIndex = i;

    const label = document.createElement('div');
    label.className = 'slide-label';
    label.textContent = `Slide ${i + 1}`;

    const placeholder = document.createElement('div');
    placeholder.className = 'slide-placeholder';
    placeholder.innerHTML = '<div class="spinner"></div>';

    wrapper.appendChild(label);
    wrapper.appendChild(placeholder);

    wrapper.addEventListener('click', (e) => {
      handleSlideClick(e, i);
    });

    container.appendChild(wrapper);
  }
}

function placeRenderedSlide(canvas, index) {
  slideCanvases[index] = canvas;

  const wrapper = $(`.slide-wrapper[data-slide-index="${index}"]`);
  if (!wrapper) return;

  const placeholder = wrapper.querySelector('.slide-placeholder');
  if (placeholder) placeholder.remove();

  const existing = wrapper.querySelector('.slide-canvas-content');
  if (existing) existing.remove();

  const img = document.createElement('img');
  img.className = 'slide-canvas-content';
  img.src = canvas.toDataURL('image/png');
  img.alt = `Slide ${index + 1}`;
  img.draggable = false;

  const label = wrapper.querySelector('.slide-label');
  if (label) {
    label.after(img);
  } else {
    wrapper.appendChild(img);
  }
}

async function triggerHdRender(buffer) {
  if (hdLoading) return;
  hdLoading = true;
  updateHdButton();

  try {
    const canvases = await renderSlidesHD(buffer, {
      onProgress: (msg) => {
        const label = $('#hd-label');
        if (label) label.textContent = msg;
      },
      onSlideCount: (count) => {
        createSlideplaceholders(count);
      },
      onSlide: (canvas, index) => {
        placeRenderedSlide(canvas, index);
      },
    });

    slideCanvases = canvases;
    hdMode = true;
    localStorage.setItem(HD_PREF_KEY, 'true');
  } catch (err) {
    console.error('HD render failed:', err);
    showError(`HD render failed: ${err.message}`);
    hdMode = false;
    localStorage.removeItem(HD_PREF_KEY);
  } finally {
    hdLoading = false;
    updateHdButton();
  }
}

$('#btn-hd-preview').addEventListener('click', async () => {
  if (hdLoading) return;

  if (hdMode) {
    hdMode = false;
    localStorage.removeItem(HD_PREF_KEY);
    updateHdButton();

    if (currentBuffer) {
      if (slideHtmls.length === 0) {
        slideHtmls = await pptxToHtml(currentBuffer, {
          width: RENDER_WIDTH,
          height: RENDER_HEIGHT,
          scaleToFit: true,
        });
      }
      renderSlides();
    }
    return;
  }

  if (!currentBuffer) return;
  triggerHdRender(currentBuffer);
});

// --- Preview update ---
function schedulePreviewUpdate() {
  clearTimeout(previewDebounceTimer);
  previewDebounceTimer = setTimeout(updateModifiedPreview, 400);
}

async function updateModifiedPreview() {
  try {
    const modifiedBuffer = await buildModifiedBuffer(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap, svgColorMap);
    currentBuffer = modifiedBuffer;

    if (hdMode) {
      triggerHdRender(modifiedBuffer);
      return;
    }

    slideHtmls = await pptxToHtml(modifiedBuffer, {
      width: RENDER_WIDTH,
      height: RENDER_HEIGHT,
      scaleToFit: true,
    });
    renderSlides();
  } catch (err) {
    console.error('Preview update failed:', err);
  }
}

// --- Download ---
$('#btn-download').addEventListener('click', async () => {
  if (!originalBuffer) return;

  const btn = $('#btn-download');
  const originalContent = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '<span class="animate-pulse">Building...</span>';

  try {
    const blob = await replaceColors(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap, svgColorMap);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.replace('.pptx', '_recolored.pptx');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('Download failed:', err);
    showError(`Download failed: ${err.message}`);
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalContent;
  }
});

// --- Reset all ---
$('#btn-reset-all').addEventListener('click', () => {
  for (const entry of colorList) {
    colorMap.set(entry.hex, entry.hex);
  }
  renderColorTable();
  schedulePreviewUpdate();
});

// --- Upload new file button ---
$('#btn-new-file').addEventListener('click', () => {
  fileInput.value = '';
  fileInput.click();
});
