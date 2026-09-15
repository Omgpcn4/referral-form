const MAX_DIMENSION_PX = 2000;
const PDF_RENDER_DPI = 150;

let nextId = 1;

function canvasToPngBytes(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Failed to encode image"));
        return;
      }
      blob.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
    }, "image/png");
  });
}

function scaledSize(width, height, maxDimension) {
  const scale = Math.min(1, maxDimension / Math.max(width, height));
  return { width: Math.round(width * scale), height: Math.round(height * scale) };
}

async function processImageFile(file) {
  const bitmap = await createImageBitmap(file);
  const { width, height } = scaledSize(bitmap.width, bitmap.height, MAX_DIMENSION_PX);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(bitmap, 0, 0, width, height);
  const bytes = await canvasToPngBytes(canvas);
  const dataUrl = canvas.toDataURL("image/png");
  return [
    {
      id: nextId++,
      label: file.name,
      width,
      height,
      bytes,
      thumbnail: dataUrl,
    },
  ];
}

let pdfjsLibPromise;
async function loadPdfjs() {
  if (!pdfjsLibPromise) {
    pdfjsLibPromise = Promise.all([
      import("pdfjs-dist"),
      import("pdfjs-dist/build/pdf.worker.min.mjs?url"),
    ]).then(([pdfjsLib, workerUrlModule]) => {
      pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrlModule.default;
      return pdfjsLib;
    });
  }
  return pdfjsLibPromise;
}

async function processPdfFile(file) {
  const pdfjsLib = await loadPdfjs();
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const results = [];
  const scale = PDF_RENDER_DPI / 72;
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    const ctx = canvas.getContext("2d");
    await page.render({ canvasContext: ctx, viewport }).promise;

    const { width, height } = scaledSize(canvas.width, canvas.height, MAX_DIMENSION_PX);
    let finalCanvas = canvas;
    if (width !== canvas.width || height !== canvas.height) {
      finalCanvas = document.createElement("canvas");
      finalCanvas.width = width;
      finalCanvas.height = height;
      finalCanvas.getContext("2d").drawImage(canvas, 0, 0, width, height);
    }

    const bytes = await canvasToPngBytes(finalCanvas);
    const dataUrl = finalCanvas.toDataURL("image/png");
    const pageLabel = pdf.numPages > 1 ? `${file.name} (page ${pageNum}/${pdf.numPages})` : file.name;
    results.push({ id: nextId++, label: pageLabel, width, height, bytes, thumbnail: dataUrl });
  }
  return results;
}

export async function processAttachmentFile(file) {
  if (file.type === "application/pdf") {
    return processPdfFile(file);
  }
  if (file.type.startsWith("image/")) {
    return processImageFile(file);
  }
  throw new Error(`Unsupported file type: ${file.type || file.name}`);
}
