let coreLoaded = false;
let coreLoading = null;

const WORKER_PATH = '/webviewer-core';

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.onload = resolve;
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(script);
  });
}

export async function loadApryseCore(onProgress) {
  if (coreLoaded) return;
  if (coreLoading) return coreLoading;

  coreLoading = (async () => {
    onProgress?.('Loading HD renderer...');
    await loadScript(`${WORKER_PATH}/webviewer-core.min.js`);

    const Core = window.Core;
    if (!Core) throw new Error('Core failed to initialize');

    Core.setWorkerPath(WORKER_PATH);

    onProgress?.('HD renderer ready');
    coreLoaded = true;
  })();

  return coreLoading;
}

export function isApryseLoaded() {
  return coreLoaded;
}

function waitForAllPages(doc) {
  return new Promise((resolve) => {
    let settled = false;
    let debounceTimer = null;

    const checkDone = () => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        if (!settled) {
          settled = true;
          doc.removeEventListener('pagesUpdated', onPagesUpdated);
          resolve(doc.getPageCount());
        }
      }, 500);
    };

    const onPagesUpdated = () => {
      checkDone();
    };

    doc.addEventListener('pagesUpdated', onPagesUpdated);

    // If all pages are already available (e.g. small doc), resolve after a short wait
    setTimeout(() => {
      if (!settled && doc.getPageCount() > 0) {
        checkDone();
      }
    }, 200);
  });
}

export async function renderSlidesHD(arrayBuffer, onProgress) {
  if (!coreLoaded) {
    await loadApryseCore(onProgress);
  }

  const Core = window.Core;
  onProgress?.('Opening presentation...');

  const doc = await Core.createDocument(arrayBuffer, {
    extension: 'pptx',
    filename: 'presentation.pptx',
  });

  onProgress?.('Converting slides...');
  const pageCount = await waitForAllPages(doc);

  const canvases = [];

  for (let i = 1; i <= pageCount; i++) {
    onProgress?.(`Rendering slide ${i}/${pageCount}...`);

    const canvas = await new Promise((resolve) => {
      doc.loadCanvas({
        pageNumber: i,
        zoom: 2,
        pageRotation: Core.PageRotation.e_0,
        drawComplete: (c) => resolve(c),
      });
    });

    canvases.push(canvas);
  }

  return canvases;
}
