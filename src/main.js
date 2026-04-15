import './style.css';
import { pptxToHtml } from '@jvmr/pptx-to-html';
import { extractColors, buildColorList } from './colorExtractor.js';
import { extractImageColors } from './imageColors.js';
import { replaceColors, buildModifiedBuffer } from './colorReplacer.js';
import { openColorPicker, useEyeDropper } from './colorPicker.js';

let originalBuffer = null;
let fileName = '';
let colorList = [];
let colorMap = new Map();
let themeNameToOrigHex = new Map();
let imageColorMap = new Map();
let slideHtmls = [];
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
    const { directColors, themeColorUsage } = await extractColors(originalBuffer);

    showLoading('Scanning images...');
    imageColorMap = await extractImageColors(originalBuffer);

    colorList = buildColorList(directColors, themeColorUsage, imageColorMap);

    themeNameToOrigHex = new Map();
    for (const [name, { hex }] of themeColorUsage) {
      themeNameToOrigHex.set(name, hex);
    }

    colorMap = new Map();
    for (const entry of colorList) {
      colorMap.set(entry.hex, entry.hex);
    }

    showLoading('Rendering slides...');
    slideHtmls = await pptxToHtml(originalBuffer, {
      width: 960,
      height: 540,
      scaleToFit: true,
    });

    hideLoading();
    renderUI();
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

// --- Render ---
function renderUI() {
  $('#file-name').textContent = fileName;
  $('#slide-count').textContent = `${slideHtmls.length} slide${slideHtmls.length !== 1 ? 's' : ''}`;

  renderColorTable();
  renderSlides();
}

function renderColorTable() {
  const container = $('#color-table');
  container.innerHTML = '';

  if (colorList.length === 0) {
    container.innerHTML = '<div class="px-4 py-8 text-center text-gray-500 text-sm">No colors found in this presentation.</div>';
    return;
  }

  for (const entry of colorList) {
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
          <span class="badge ${entry.type === 'theme' ? 'badge-theme' : entry.type === 'image' ? 'badge-image' : 'badge-direct'}">
            ${entry.type === 'theme' ? entry.themeLabel : entry.type === 'image' ? 'Image' : 'Direct'}
          </span>
          <span class="text-[0.6rem] text-gray-600">${entry.count}x</span>
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

function renderSlides() {
  const container = $('#slides-container');
  container.innerHTML = '';

  slideHtmls.forEach((html, i) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'slide-wrapper clickable-slide';

    const label = document.createElement('div');
    label.className = 'slide-label';
    label.textContent = `Slide ${i + 1}`;

    const content = document.createElement('div');
    content.className = 'slide-html-content';
    content.innerHTML = html;

    wrapper.appendChild(label);
    wrapper.appendChild(content);

    wrapper.addEventListener('click', (e) => {
      handleSlideClick(e);
    });

    container.appendChild(wrapper);
  });
}

// --- Slide click -> color pick ---
function handleSlideClick(event) {
  const el = document.elementFromPoint(event.clientX, event.clientY);
  if (!el) return;

  const color = getElementColor(el);
  if (!color) return;

  const hex = rgbToHex(color).toUpperCase();
  if (hex.length === 6) {
    highlightColorRow(hex);
  }
}

function getElementColor(el) {
  const style = window.getComputedStyle(el);

  if (style.backgroundColor && style.backgroundColor !== 'rgba(0, 0, 0, 0)' && style.backgroundColor !== 'transparent') {
    return style.backgroundColor;
  }
  if (style.color && style.color !== 'rgba(0, 0, 0, 0)') {
    return style.color;
  }
  return null;
}

function rgbToHex(rgb) {
  const match = rgb.match(/\d+/g);
  if (!match || match.length < 3) return '';
  const [r, g, b] = match.map(Number);
  return [r, g, b].map(c => c.toString(16).padStart(2, '0')).join('');
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

// --- Preview update ---
function schedulePreviewUpdate() {
  clearTimeout(previewDebounceTimer);
  previewDebounceTimer = setTimeout(updateModifiedPreview, 400);
}

async function updateModifiedPreview() {
  try {
    const modifiedBuffer = await buildModifiedBuffer(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap);
    slideHtmls = await pptxToHtml(modifiedBuffer, {
      width: 960,
      height: 540,
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
    const blob = await replaceColors(originalBuffer, colorMap, themeNameToOrigHex, imageColorMap);
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
