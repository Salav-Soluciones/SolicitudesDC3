// Referencias a los elementos
const fileInput = document.getElementById("fileInput");
const logoInput = document.getElementById("logoInput");
const tableHead = document.getElementById("tableHead");
const tableBody = document.getElementById("tableBody");
const btnPDF = document.getElementById("btnPDF");
const batchSizeInput = document.getElementById("batchSizeInput");
const loader = document.getElementById("loader");
const loaderText = document.getElementById("loaderText");
const progressBar = document.getElementById("progressBar");
const progressDetails = document.getElementById("progressDetails");
const useFolderCheckbox = document.getElementById("useFolder");

let excelData = [];
let logoImageBytes = null; // ArrayBuffer de la imagen (si se sube)

// Leer logo (opcional) al seleccionarlo
logoInput.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) {
    logoImageBytes = null;
    return;
  }
  logoImageBytes = await f.arrayBuffer();
});

// Leer archivo Excel
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    mostrarTabla(excelData);
    btnPDF.disabled = false;
  };
  reader.readAsArrayBuffer(file);
});

// Mostrar tabla HTML
function mostrarTabla(data) {
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";

  if (data.length === 0) return;

  const headers = data[0];
  const headRow = document.createElement("tr");
  headers.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    headRow.appendChild(th);
  });
  tableHead.appendChild(headRow);

  for (let i = 1; i < data.length; i++) {
    const row = document.createElement("tr");
    data[i].forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell;
      row.appendChild(td);
    });
    tableBody.appendChild(row);
  }
}

/* Helpers UI */
function showLoader(text = "Procesando...", total = 0) {
  loaderText.textContent = text;
  progressBar.style.width = "0%";
  progressDetails.textContent = `0 / ${total}`;
  loader.style.display = "block";
}
function hideLoader() {
  loader.style.display = "none";
}
function updateProgress(done, total) {
  const pct = total === 0 ? 0 : Math.round((done / total) * 100);
  progressBar.style.width = pct + "%";
  progressDetails.textContent = `${done} / ${total} (${pct}%)`;
}

/* Función principal: paginado con opción de carpeta o ZIP por lotes, progreso, logo incrustado */
btnPDF.addEventListener("click", async () => {
  if (!excelData || excelData.length <= 1) {
    alert("No hay filas de datos para generar PDFs.");
    return;
  }

  // Validar batch size
  let BATCH_SIZE = parseInt(batchSizeInput.value, 10);
  if (!Number.isFinite(BATCH_SIZE) || BATCH_SIZE <= 0) BATCH_SIZE = 50;

  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const headers = excelData[0];
  const totalRows = excelData.length - 1;

  // Deshabilitar UI mientras procesa
  btnPDF.disabled = true;
  fileInput.disabled = true;
  logoInput.disabled = true;
  batchSizeInput.disabled = true;
  useFolderCheckbox.disabled = true;

  showLoader("Iniciando generación de PDFs...", totalRows);

  // Preparar logo embebible (si existe)
  let logoImageObject = null; // será un objeto embebido de PDFLib
  let logoIsPng = false;
  if (logoImageBytes) {
    // Intentamos detectar formato por encabezado de bytes (PNG o JPEG)
    const header = new Uint8Array(logoImageBytes).subarray(0, 8);
    const isPng = header[0] === 0x89 && header[1] === 0x50 && header[2] === 0x4E;
    const isJpeg = header[0] === 0xFF && header[1] === 0xD8 && header[2] === 0xFF;
    logoIsPng = isPng;
    try {
      // Para crear logo embebido, necesitamos convertir cuando creemos el pdf: no se puede reusar entre documentos sin re-embed en pdf-lib actual
      // Guardamos logoImageBytes y marcar formato; lo embebemos dentro de cada PDF más abajo.
    } catch (err) {
      console.warn("No se pudo preparar logo:", err);
      logoImageBytes = null;
    }
  }

  const PAUSE_MS = 30; // pausa entre archivos para mantener UI responsiva

  // Opción 1: File System Access API (escribe directamente en carpeta)
  if (useFolderCheckbox.checked && window.showDirectoryPicker) {
    try {
      const dirHandle = await window.showDirectoryPicker();
      let saved = 0;
      for (let r = 1; r < excelData.length; r++) {
        const row = excelData[r];

        const pdfDoc = await PDFDocument.create();
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const titleSize = 16;
        const fontSize = 12;

        // Si hay logo, embed dentro de ESTE documento
        if (logoImageBytes) {
          try {
            if (logoIsPng) {
              logoImageObject = await pdfDoc.embedPng(logoImageBytes);
            } else {
              logoImageObject = await pdfDoc.embedJpg(logoImageBytes);
            }
          } catch (err) {
            console.warn("Error embebiendo logo en pdf:", err);
            logoImageObject = null;
          }
        }

        let page = pdfDoc.addPage([600, 800]);

        // Dibujar logo en la esquina superior derecha si existe
        if (logoImageObject) {
          const { width: iw, height: ih } = logoImageObject.scale(1);
          // Queremos logo de ancho máximo 120 px manteniendo proporción
          const maxW = 120;
          const scale = Math.min(1, maxW / iw);
          const w = iw * scale;
          const h = ih * scale;
          page.drawImage(logoImageObject, { x: 600 - 50 - w, y: 800 - 50 - h, width: w, height: h });
        }

        page.drawText("Datos del Excel:", { x: 50, y: 760, size: titleSize, font, color: rgb(0, 0, 0) });

        let y = 730;
        for (let c = 0; c < Math.max(headers.length, row.length); c++) {
          const label = headers[c] !== undefined ? String(headers[c]) : `Col ${c + 1}`;
          const value = row[c] !== undefined ? String(row[c]) : "";
          const line = `${label}: ${value}`;

          if (y < 50) {
            page = pdfDoc.addPage([600, 800]);
            y = 760;
          }

          page.drawText(line, { x: 50, y, size: fontSize, font });
          y -= 18;
        }

        const pdfBytes = await pdfDoc.save(); // Uint8Array
        const blob = new Blob([pdfBytes], { type: "application/pdf" });

        const firstCell = row[0] ? String(row[0]).replace(/[^\w\-\.]/g, "_") : `row${r}`;
        const filename = `${firstCell}_${r}.pdf`;

        // Guardar archivo en la carpeta seleccionada
        const fileHandle = await dirHandle.getFileHandle(filename, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();

        saved++;
        updateProgress(saved, totalRows);
        await new Promise(res => setTimeout(res, PAUSE_MS));
      }

      alert("Se han guardado los PDFs en la carpeta seleccionada.");
      hideLoader();
      // Rehabilitar UI
      btnPDF.disabled = false;
      fileInput.disabled = false;
      logoInput.disabled = false;
      batchSizeInput.disabled = false;
      useFolderCheckbox.disabled = false;
      return;
    } catch (err) {
      console.error("Error usando File System Access API:", err);
      // Si falla o usuario cancela, seguimos al fallback por lotes
    }
  }

  // Fallback: Generar zips por lotes con JSZip
  if (typeof JSZip === "undefined") {
    alert('JSZip no está cargado. Añade: <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script> en tu HTML para el modo por lotes.');
    hideLoader();
    btnPDF.disabled = false;
    fileInput.disabled = false;
    logoInput.disabled = false;
    batchSizeInput.disabled = false;
    useFolderCheckbox.disabled = false;
    return;
  }

  const totalBatches = Math.ceil(totalRows / BATCH_SIZE);
  let globalDone = 0;

  for (let b = 0; b < totalBatches; b++) {
    const zip = new JSZip();
    const startIndex = 1 + b * BATCH_SIZE;
    const endIndex = Math.min(startIndex + BATCH_SIZE - 1, excelData.length - 1);
    loaderText.textContent = `Generando batch ${b + 1} de ${totalBatches} (filas ${startIndex}..${endIndex})`;

    for (let r = startIndex; r <= endIndex; r++) {
      const row = excelData[r];

      const pdfDoc = await PDFDocument.create();
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const titleSize = 16;
      const fontSize = 12;

      // Embebemos logo en este pdf si existe
      if (logoImageBytes) {
        try {
          if (logoIsPng) {
            logoImageObject = await pdfDoc.embedPng(logoImageBytes);
          } else {
            logoImageObject = await pdfDoc.embedJpg(logoImageBytes);
          }
        } catch (err) {
          console.warn("Error embebiendo logo en pdf del batch:", err);
          logoImageObject = null;
        }
      }

      let page = pdfDoc.addPage([600, 800]);

      if (logoImageObject) {
        const { width: iw, height: ih } = logoImageObject.scale(1);
        const maxW = 120;
        const scale = Math.min(1, maxW / iw);
        const w = iw * scale;
        const h = ih * scale;
        page.drawImage(logoImageObject, { x: 600 - 50 - w, y: 800 - 50 - h, width: w, height: h });
      }

      page.drawText("Datos del Excel:", { x: 50, y: 760, size: titleSize, font, color: rgb(0, 0, 0) });

      let y = 730;
      for (let c = 0; c < Math.max(headers.length, row.length); c++) {
        const label = headers[c] !== undefined ? String(headers[c]) : `Col ${c + 1}`;
        const value = row[c] !== undefined ? String(row[c]) : "";
        const line = `${label}: ${value}`;

        if (y < 50) {
          page = pdfDoc.addPage([600, 800]);
          y = 760;
        }

        page.drawText(line, { x: 50, y, size: fontSize, font });
        y -= 18;
      }

      const pdfBytes = await pdfDoc.save();
      const firstCell = row[0] ? String(row[0]).replace(/[^\w\-\.]/g, "_") : `row${r}`;
      const filename = `${firstCell}_${r}.pdf`;

      zip.file(filename, pdfBytes);

      globalDone++;
      updateProgress(globalDone, totalRows);

      // pequeña pausa
      await new Promise(res => setTimeout(res, PAUSE_MS));
    }

    // Generar y descargar ZIP del lote
    loaderText.textContent = `Comprimiendo batch ${b + 1} de ${totalBatches}...`;
    try {
      const zipBlob = await zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      }, (metadata) => {
        // metadata.percent puede usarse para mostrar progreso de compresión local si se desea
        // pero cuidado con actualizar demasiado frecuentemente; aquí podríamos mostrar la mitad del progreso del batch
      });

      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `pdfs_por_fila_batch_${b + 1}_of_${totalBatches}.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Error generando ZIP del lote", b + 1, err);
      alert("Error generando ZIP del lote. Revisa la consola.");
      hideLoader();
      btnPDF.disabled = false;
      fileInput.disabled = false;
      logoInput.disabled = false;
      batchSizeInput.disabled = false;
      useFolderCheckbox.disabled = false;
      return;
    }

    // pequeña pausa entre batches
    await new Promise(res => setTimeout(res, 200));
  }

  hideLoader();
  alert(`Se han generado ${totalBatches} zip(s) con los PDFs (total: ${totalRows} PDFs).`);

  // Rehabilitar UI
  btnPDF.disabled = false;
  fileInput.disabled = false;
  logoInput.disabled = false;
  batchSizeInput.disabled = false;
  useFolderCheckbox.disabled = false;
});
