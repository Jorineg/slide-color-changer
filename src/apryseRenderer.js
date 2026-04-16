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
    Core.enableFullPDF();

    onProgress?.('HD renderer ready');
    coreLoaded = true;
  })();

  return coreLoading;
}

export function isApryseLoaded() {
  return coreLoaded;
}

export async function renderSlidesHD(arrayBuffer, onProgress) {
  if (!coreLoaded) {
    await loadApryseCore(onProgress);
  }

  const Core = window.Core;
  onProgress?.('Opening presentation...');

  const blob = new Blob([arrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });

  const doc = await Core.createDocument(blob, {
    extension: 'pptx',
    officeOptions: {
      templateValues: undefined,
    },
  });

  const pageCount = doc.getPageCount();
  const canvases = [];

  for (let i = 1; i <= pageCount; i++) {
    onProgress?.(`Rendering slide ${i}/${pageCount}...`);

    const canvas = await new Promise((resolve, reject) => {
      doc.loadCanvas({
        pageNumber: i,
        zoom: 2,
        pageRotation: Core.PageRotation.e_0,
        drawComplete: (c) => resolve(c),
        renderError: (err) => reject(err || new Error(`Render failed for slide ${i}`)),
      });
    });

    canvases.push(canvas);
  }

  return canvases;
}
