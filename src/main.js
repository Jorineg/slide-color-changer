import './style.css';
import { pptxToHtml } from '@jvmr/pptx-to-html';
import { extractColors, buildColorList } from './colorExtractor.js';
import { replaceColors, buildModifiedBuffer } from './colorReplacer.js';
import { openColorPicker, useEyeDropper } from './colorPicker.js';

let originalBuffer = null;
let fileName = '';
let colorList = [];
let colorMap = new Map();
let themeNameToOrigHex = new Map();
let originalSlideHtmls = [];
let modifiedSlideHtmls = [];
let activeTab = 'original';
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
  if (file && file.name.endsWith('.pptx')) handleFile(file);
});

fileInput.addEventListener('change', () => {
  const file = fileInput.files[0];
  if (file) handleFile(file);
});

// --- File handling ---
async function handleFile(file) {
  fileName = file.name;
  showLoading('Reading file...');

  originalBuffer = await file.arrayBuffer();

  showLoading('Extracting colors...');
  const { directColors, themeColorUsage, themeColorMap } = await extractColors(originalBuffer);
  colorList = buildColorList(directColors, themeColorUsage);

  themeNameToOrigHex = new Map();
  for (const [name, { hex }] of themeColorUsage) {
    themeNameToOrigHex.set(name, hex);
  }

  colorMap = new Map();
  for (const entry of colorList) {
    colorMap.set(entry.hex, entry.hex);
  }

  showLoading('Rendering slides...');
  originalSlideHtmls = await pptxToHtml(originalBuffer, {
    width: 960,
    height: 540,
    scaleToFit: true,
  });
  modifiedSlideHtmls = [...originalSlideHtmls];

  hideLoading();
  renderUI();
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
  $('#slide-count').textContent = `${originalSlideHtmls.length} slide${originalSlideHtmls.length !== 1 ? 's' : ''}`;

  renderColorTable();
  renderSlides();
  setupTabs();
}

function renderColorTable() {
  const container = $('#color-table');
  container.innerHTML = '';

  for (const entry of colorList) {
    const currentHex = colorMap.get(entry.hex) || entry.hex;
    const isModified = currentHex !== entry.hex;

    const row = document.createElement('div');
    row.className = 'color-row flex items-center gap-3 px-4 py-3';
    row.dataset.colorId = entry.id;
    row.dataset.origHex = entry.hex;

    row.innerHTML = `
      <div class="flex items-center gap-3 flex-1 min-w-0">
        <div class="color-swatch original-swatch" style="background:#${entry.hex}" title="#${entry.hex}"></div>
        <div class="flex flex-col min-w-0">
          <span class="text-xs font-mono text-gray-400">#${entry.hex}</span>
          <div class="flex items-center gap-1.5 mt-0.5">
            <span class="badge ${entry.type === 'theme' ? 'badge-theme' : 'badge-direct'}">
              ${entry.type === 'theme' ? entry.themeLabel : 'Direct'}
            </span>
            <span class="text-[0.65rem] text-gray-600">${entry.count}x</span>
          </div>
        </div>
      </div>
      <svg class="arrow-icon h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
        <path stroke-linecap="round" stroke-linejoin="round" d="M13 7l5 5m0 0l-5 5m5-5H6" />
      </svg>
      <div class="flex items-center gap-2">
        <div class="color-swatch target-swatch" style="background:#${currentHex}" title="#${currentHex}"></div>
        <div class="flex flex-col items-end">
          <span class="text-xs font-mono text-gray-400 target-hex">#${currentHex}</span>
          <div class="flex items-center gap-1 mt-0.5">
            ${isModified ? `<button class="reset-btn text-[0.65rem] text-indigo-400 hover:text-indigo-300 cursor-pointer">Reset</button>` : ''}
            ${'EyeDropper' in window ? `<button class="eyedropper-btn text-[0.65rem] text-gray-500 hover:text-gray-300 cursor-pointer" title="Pick from screen">
              <svg class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.042 21.672L13.684 16.6m0 0l-2.51 2.225.569-9.47 5.227 7.917-3.286-.672zM12 2.25V4.5m5.834.166l-1.591 1.591M20.25 10.5H18M7.757 14.743l-1.59 1.59M6 10.5H3.75m4.007-4.243l-1.59-1.59" />
              </svg>
            </button>` : ''}
          </div>
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

  const htmls = activeTab === 'original' ? originalSlideHtmls : modifiedSlideHtmls;

  htmls.forEach((html, i) => {
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
      handleSlideClick(e, wrapper);
    });

    container.appendChild(wrapper);
  });
}

function setupTabs() {
  const tabOriginal = $('#tab-original');
  const tabModified = $('#tab-modified');

  tabOriginal.onclick = () => {
    activeTab = 'original';
    tabOriginal.className = 'px-4 py-1.5 rounded-md text-sm font-medium bg-gray-700 text-white transition-colors';
    tabModified.className = 'px-4 py-1.5 rounded-md text-sm font-medium text-gray-400 hover:text-white transition-colors';
    renderSlides();
  };

  tabModified.onclick = () => {
    activeTab = 'modified';
    tabModified.className = 'px-4 py-1.5 rounded-md text-sm font-medium bg-gray-700 text-white transition-colors';
    tabOriginal.className = 'px-4 py-1.5 rounded-md text-sm font-medium text-gray-400 hover:text-white transition-colors';
    renderSlides();
  };
}

// --- Slide click → color pick ---
function handleSlideClick(event, slideWrapper) {
  const el = document.elementFromPoint(event.clientX, event.clientY);
  if (!el) return;

  const color = getElementColor(el);
  if (!color) return;

  const hex = rgbToHex(color).toUpperCase();
  highlightColorRow(hex);
}

function getElementColor(el) {
  const style = window.getComputedStyle(el);

  if (style.backgroundColor && style.backgroundColor !== 'rgba(0, 0, 0, 0)' && style.backgroundColor !== 'transparent') {
    return style.backgroundColor;
  }
  if (style.color) {
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
    const distance = colorDistance(clickedHex, entry.hex);
    if (distance < bestDistance) {
      bestDistance = distance;
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
  previewDebounceTimer = setTimeout(updateModifiedPreview, 300);
}

async function updateModifiedPreview() {
  const modifiedBuffer = await buildModifiedBuffer(originalBuffer, colorMap, themeNameToOrigHex);
  modifiedSlideHtmls = await pptxToHtml(modifiedBuffer, {
    width: 960,
    height: 540,
    scaleToFit: true,
  });

  if (activeTab === 'modified') {
    renderSlides();
  }
}

// --- Download ---
$('#btn-download').addEventListener('click', async () => {
  const btn = $('#btn-download');
  btn.disabled = true;
  btn.textContent = 'Building...';

  try {
    const blob = await replaceColors(originalBuffer, colorMap, themeNameToOrigHex);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.replace('.pptx', '_recolored.pptx');
    a.click();
    URL.revokeObjectURL(url);
  } finally {
    btn.disabled = false;
    btn.innerHTML = `
      <svg class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
        <path stroke-linecap="round" stroke-linejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
      </svg>
      Download Modified .pptx
    `;
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

// --- Upload new file ---
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && $('#main-content').classList.contains('hidden') === false) {
    // Could add reset functionality here
  }
});
