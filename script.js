// Referencias a los elementos
const fileInput = document.getElementById("fileInput");
const tableHead = document.getElementById("tableHead");
const tableBody = document.getElementById("tableBody");
const btnPDF = document.getElementById("btnPDF");

let excelData = [];

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

// Generar PDF
btnPDF.addEventListener("click", async () => {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([600, 800]);
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontSize = 12;

  page.drawText("Datos del Excel cargado:", { x: 50, y: 760, size: 16, font, color: rgb(0, 0, 0) });

  let y = 730;
  excelData.forEach((row) => {
    page.drawText(row.join(" | "), { x: 50, y, size: fontSize, font });
    y -= 20;
  });

  const pdfBytes = await pdfDoc.save();
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "resultado.pdf";
  a.click();
});
