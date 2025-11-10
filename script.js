// Referencias a los elementos
const fileInput = document.getElementById("fileInput");
const logoInput = document.getElementById("logoInput");
const tableHead = document.getElementById("tableHead");
const tableBody = document.getElementById("tableBody");
const btnPDF = document.getElementById("btnPDF");
const batchSizeInput = document.getElementById("batchSizeInput");
const rowsPerPdfInput = document.getElementById("rowsPerPdfInput");
const linesPerPageInput = document.getElementById("linesPerPageInput");
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

/* Función principal: paginado/segmentado con opciones de agrupación y páginas internas */
btnPDF.addEventListener("click", async () => {
  if (!excelData || excelData.length <= 1) {
    alert("No hay filas de datos para generar PDFs.");
    return;
  }

  // Parámetros de UI
  let BATCH_SIZE = parseInt(batchSizeInput.value, 10);
  if (!Number.isFinite(BATCH_SIZE) || BATCH_SIZE <= 0) BATCH_SIZE = 50;

  let ROWS_PER_PDF = parseInt(rowsPerPdfInput.value, 10);
  if (!Number.isFinite(ROWS_PER_PDF) || ROWS_PER_PDF <= 0) ROWS_PER_PDF = 1;

  let LINES_PER_PAGE = parseInt(linesPerPageInput.value, 10);
  if (!Number.isFinite(LINES_PER_PAGE) || LINES_PER_PAGE <= 5) LINES_PER_PAGE = 30;

  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const headers = excelData[0];
  const totalRows = excelData.length - 1;
  const totalPdfs = Math.ceil(totalRows / ROWS_PER_PDF);

  // Deshabilitar UI mientras procesa
  btnPDF.disabled = true;
  fileInput.disabled = true;
  logoInput.disabled = true;
  batchSizeInput.disabled = true;
  rowsPerPdfInput.disabled = true;
  linesPerPageInput.disabled = true;
  useFolderCheckbox.disabled = true;

  showLoader("Iniciando generación de PDFs...", totalPdfs);

  // Preparar logo (detectar tipo)
  let logoIsPng = false;
  if (logoImageBytes) {
    const header = new Uint8Array(logoImageBytes).subarray(0, 8);
    logoIsPng = header[0] === 0x89 && header[1] === 0x50 && header[2] === 0x4E;
  }

  const PAUSE_MS = 25;

  // Try File System Access API if requested
  if (useFolderCheckbox.checked && window.showDirectoryPicker) {
    try {
      const dirHandle = await window.showDirectoryPicker();
      let pdfSaved = 0;

      for (let pdfIndex = 0; pdfIndex < totalPdfs; pdfIndex++) {
        const startRowIndex = 1 + pdfIndex * ROWS_PER_PDF;
        const endRowIndex = Math.min(startRowIndex + ROWS_PER_PDF - 1, totalRows);

        const pdfDoc = await PDFDocument.create();
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const titleSize = 16;
        const fontSize = 12;

        // embed logo per-document if provided
        let logoImageObject = null;
        if (logoImageBytes) {
          try {
            logoImageObject = logoIsPng ? await pdfDoc.embedPng(logoImageBytes) : await pdfDoc.embedJpg(logoImageBytes);
          } catch (err) {
            console.warn("Error embebiendo logo:", err);
            logoImageObject = null;
          }
        }

        let page = pdfDoc.addPage([600, 800]);
        let y = 760;

        // Si dibujamos logo en la primera página
        if (logoImageObject) {
          const { width: iw, height: ih } = logoImageObject.scale(1);
          const maxW = 120;
          const scale = Math.min(1, maxW / iw);
          const w = iw * scale;
          const h = ih * scale;
          page.drawImage(logoImageObject, { x: 600 - 50 - w, y: 800 - 50 - h, width: w, height: h });
        }

        page.drawText("Datos del Excel:", { x: 50, y, size: titleSize, font, color: rgb(0, 0, 0) });
        y -= 30;

        // Para cada fila incluida en este PDF
        for (let r = startRowIndex; r <= endRowIndex; r++) {
          const row = excelData[r];

          // Escribir un subtítulo con fila/índice
          const rowTitle = `Fila ${r}`;
          if (y < 60) { page = pdfDoc.addPage([600, 800]); y = 760; }
          page.drawText(rowTitle, { x: 50, y, size: 13, font });
          y -= 18;

          // Para cada columna/valor de la fila, formatear en líneas: "Header: value"
          let lineCount = 0;
          for (let c = 0; c < Math.max(headers.length, row.length); c++) {
            const label = headers[c] !== undefined ? String(headers[c]) : `Col ${c + 1}`;
            const value = row[c] !== undefined ? String(row[c]) : "";
            const line = `${label}: ${value}`;

            // Si alcanza la capacidad de líneas en la página, crear una nueva página
            if (lineCount >= LINES_PER_PAGE || y < 60) {
              page = pdfDoc.addPage([600, 800]);
              y = 760;
              lineCount = 0;
            }

            page.drawText(line, { x: 50, y, size: fontSize, font });
            y -= 18;
            lineCount++;
          }

          // separación entre filas dentro del mismo PDF
          y -= 8;
          if (y < 60) { page = pdfDoc.addPage([600, 800]); y = 760; }
        }

        const pdfBytes = await pdfDoc.save();
        const blob = new Blob([pdfBytes], { type: "application/pdf" });

        // Nombre del archivo: si agrupa varias filas, usar rango; si 1 fila, usar primera celda
        let filename;
        if (startRowIndex === endRowIndex) {
          const row = excelData[startRowIndex];
          const firstCell = row[0] ? String(row[0]).replace(/[^\w\-\.]/g, "_") : `row${startRowIndex}`;
          filename = `${firstCell}_${startRowIndex}.pdf`;
        } else {
          const firstRow = excelData[startRowIndex];
          const firstCell = firstRow[0] ? String(firstRow[0]).replace(/[^\w\-\.]/g, "_") : `rows${startRowIndex}`;
          filename = `${firstCell}_${startRowIndex}_to_${endRowIndex}.pdf`;
        }

        const fileHandle = await dirHandle.getFileHandle(filename, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();

        pdfSaved++;
        updateProgress(pdfSaved, totalPdfs);
        await new Promise(res => setTimeout(res, PAUSE_MS));
      }

      alert("Se han guardado los PDFs en la carpeta seleccionada.");
      hideLoader();
      // Rehabilitar UI
      btnPDF.disabled = false;
      fileInput.disabled = false;
      logoInput.disabled = false;
      batchSizeInput.disabled = false;
      rowsPerPdfInput.disabled = false;
      linesPerPageInput.disabled = false;
      useFolderCheckbox.disabled = false;
      return;
    } catch (err) {
      console.error("Error con File System Access API:", err);
      // si falla, continuamos con fallback por lotes
    }
  }

  // Fallback: zip por lotes (JSZip)
  if (typeof JSZip === "undefined") {
    alert('JSZip no está cargado. Añade: <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script> en tu HTML para el modo por lotes.');
    hideLoader();
    btnPDF.disabled = false;
    fileInput.disabled = false;
    logoInput.disabled = false;
    batchSizeInput.disabled = false;
    rowsPerPdfInput.disabled = false;
    linesPerPageInput.disabled = false;
    useFolderCheckbox.disabled = false;
    return;
  }

  const totalBatches = Math.ceil(totalPdfs / BATCH_SIZE);
  let globalDone = 0;

  for (let b = 0; b < totalBatches; b++) {
    const zip = new JSZip();
    const startPdfIndex = b * BATCH_SIZE;
    const endPdfIndex = Math.min(startPdfIndex + BATCH_SIZE - 1, totalPdfs - 1);
    loaderText.textContent = `Generando batch ${b + 1} de ${totalBatches} (PDFs ${startPdfIndex + 1}..${endPdfIndex + 1})`;

    for (let pdfIndex = startPdfIndex; pdfIndex <= endPdfIndex; pdfIndex++) {
      const startRowIndex = 1 + pdfIndex * ROWS_PER_PDF;
      const endRowIndex = Math.min(startRowIndex + ROWS_PER_PDF - 1, totalRows);

      const pdfDoc = await PDFDocument.create();
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const titleSize = 16;
      const fontSize = 12;

      // embed logo per-document if provided
      let logoImageObject = null;
      if (logoImageBytes) {
        try {
          logoImageObject = logoIsPng ? await pdfDoc.embedPng(logoImageBytes) : await pdfDoc.embedJpg(logoImageBytes);
        } catch (err) {
          console.warn("Error embebiendo logo en pdf:", err);
          logoImageObject = null;
        }
      }

      let page = pdfDoc.addPage([600, 800]);
      let y = 760;

      if (logoImageObject) {
        const { width: iw, height: ih } = logoImageObject.scale(1);
        const maxW = 120;
        const scale = Math.min(1, maxW / iw);
        const w = iw * scale;
        const h = ih * scale;
        page.drawImage(logoImageObject, { x: 600 - 50 - w, y: 800 - 50 - h, width: w, height: h });
      }

      page.drawText("Datos del Excel:", { x: 50, y, size: titleSize, font, color: rgb(0, 0, 0) });
      y -= 30;

      for (let r = startRowIndex; r <= endRowIndex; r++) {
        const row = excelData[r];
        const rowTitle = `Fila ${r}`;
        if (y < 60) { page = pdfDoc.addPage([600, 800]); y = 760; }
        page.drawText(rowTitle, { x: 50, y, size: 13, font });
        y -= 18;

        let lineCount = 0;
        for (let c = 0; c < Math.max(headers.length, row.length); c++) {
          const label = headers[c] !== undefined ? String(headers[c]) : `Col ${c + 1}`;
          const value = row[c] !== undefined ? String(row[c]) : "";
          const line = `${label}: ${value}`;

          if (lineCount >= LINES_PER_PAGE || y < 60) {
            page = pdfDoc.addPage([600, 800]);
            y = 760;
            lineCount = 0;
          }

          page.drawText(line, { x: 50, y, size: fontSize, font });
          y -= 18;
          lineCount++;
        }

        y -= 8;
        if (y < 60) { page = pdfDoc.addPage([600, 800]); y = 760; }
      }

      const pdfBytes = await pdfDoc.save();
      let filename;
      if (startRowIndex === endRowIndex) {
        const row = excelData[startRowIndex];
        const firstCell = row[0] ? String(row[0]).replace(/[^\w\-\.]/g, "_") : `row${startRowIndex}`;
        filename = `${firstCell}_${startRowIndex}.pdf`;
      } else {
        const firstRow = excelData[startRowIndex];
        const firstCell = firstRow[0] ? String(firstRow[0]).replace(/[^\w\-\.]/g, "_") : `rows${startRowIndex}`;
        filename = `${firstCell}_${startRowIndex}_to_${endRowIndex}.pdf`;
      }

      zip.file(filename, pdfBytes);
      globalDone++;
      updateProgress(globalDone, totalPdfs);
      await new Promise(res => setTimeout(res, PAUSE_MS));
    }

    // Generar y descargar ZIP del lote
    loaderText.textContent = `Comprimiendo batch ${b + 1} de ${totalBatches}...`;
    try {
      const zipBlob = await zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
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
      rowsPerPdfInput.disabled = false;
      linesPerPageInput.disabled = false;
      useFolderCheckbox.disabled = false;
      return;
    }

    await new Promise(res => setTimeout(res, 200));
  }

  hideLoader();
  alert(`Se han generado ${totalPdfs} PDF(s) (en ${totalBatches} zip(s)).`);

  // Rehabilitar UI
  btnPDF.disabled = false;
  fileInput.disabled = false;
  logoInput.disabled = false;
  batchSizeInput.disabled = false;
  rowsPerPdfInput.disabled = false;
  linesPerPageInput.disabled = false;
  useFolderCheckbox.disabled = false;
});
