/* Construye y descarga el anexo fotográfico .docx usando la copia local de docx. */
(function () {
  "use strict";

  const FONT = "Poppins";
  const CONTENT_WIDTH = 9026;
  const LABEL_WIDTH = 1450;
  const IMAGE_MAX_WIDTH = 430;
  const IMAGE_MAX_HEIGHT = 255;

  function paragraph(text, options) {
    return new docx.Paragraph({
      children: [new docx.TextRun({ text, font: FONT, size: 16, color: "233C35" })],
      spacing: { before: 0, after: 40 },
      ...options
    });
  }

  function cell(children, options) {
    return new docx.TableCell({
      children,
      margins: { top: 70, bottom: 70, left: 100, right: 100 },
      borders: {
        top: { style: docx.BorderStyle.SINGLE, size: 4, color: "D8E2DC" },
        bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: "D8E2DC" },
        left: { style: docx.BorderStyle.SINGLE, size: 4, color: "D8E2DC" },
        right: { style: docx.BorderStyle.SINGLE, size: 4, color: "D8E2DC" }
      },
      ...options
    });
  }

  function renderDescription(description, project) {
    return description
      .replace(/\{RAZON_SOCIAL\}/gi, project.companyName || project.establishment || "")
      .replace(/\{PLACAS\}/gi, project.vehiclePlates || "");
  }

  function imageRow(image, pageBreakBefore) {
    const imageParagraph = new docx.Paragraph({
      children: [image],
      alignment: docx.AlignmentType.CENTER,
      spacing: { before: 20, after: 20 },
      pageBreakBefore: Boolean(pageBreakBefore),
      keepNext: true
    });
    return new docx.TableRow({
      cantSplit: true,
      children: [cell([imageParagraph], {
        columnSpan: 2,
        width: { size: CONTENT_WIDTH, type: docx.WidthType.DXA }
      })]
    });
  }

  function captionRow(photo, index, project) {
    return new docx.TableRow({
      cantSplit: true,
      children: [
        cell([paragraph(`Foto N.º ${String(index + 1).padStart(2, "0")}`, {
          alignment: docx.AlignmentType.CENTER
        })], {
          width: { size: LABEL_WIDTH, type: docx.WidthType.DXA },
          verticalAlign: docx.VerticalAlign.CENTER
        }),
        cell([paragraph(renderDescription(photo.description.trim(), project))], {
          width: { size: CONTENT_WIDTH - LABEL_WIDTH, type: docx.WidthType.DXA },
          verticalAlign: docx.VerticalAlign.CENTER
        })
      ]
    });
  }

  function formatVisitDate(value) {
    const [year, month, day] = value.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    return new Intl.DateTimeFormat("es-PE", {
      day: "2-digit",
      month: "long",
      year: "numeric"
    }).format(date).toLocaleUpperCase("es-PE");
  }

  async function imageRun(photo) {
    const widthPx = Number(photo.width) || IMAGE_MAX_WIDTH;
    const heightPx = Number(photo.height) || IMAGE_MAX_HEIGHT;
    const scale = Math.min(IMAGE_MAX_WIDTH / widthPx, IMAGE_MAX_HEIGHT / heightPx, 1);
    const width = Math.max(1, Math.round(widthPx * scale));
    const height = Math.max(1, Math.round(heightPx * scale));
    const bitmap = await createImageBitmap(photo.blob);
    try {
      const rasterScale = Math.min(1, 1000 / Math.max(bitmap.width, bitmap.height));
      const rasterWidth = Math.max(1, Math.round(bitmap.width * rasterScale));
      const rasterHeight = Math.max(1, Math.round(bitmap.height * rasterScale));
      const canvas = document.createElement("canvas");
      canvas.width = rasterWidth;
      canvas.height = rasterHeight;
      const context = canvas.getContext("2d", { alpha: false });
      if (!context) throw new Error("No se pudo preparar una imagen para el documento Word.");
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, rasterWidth, rasterHeight);
      context.drawImage(bitmap, 0, 0, rasterWidth, rasterHeight);
      const png = await new Promise((resolve, reject) => {
        canvas.toBlob((result) => {
          if (result) resolve(result);
          else reject(new Error("No se pudo convertir una imagen al formato compatible con Word."));
        }, "image/png");
      });
      const bytes = new Uint8Array(await png.arrayBuffer());
      canvas.width = 0;
      canvas.height = 0;
      return new docx.ImageRun({ data: bytes, transformation: { width, height } });
    } finally {
      bitmap.close();
    }
  }

  async function createDocument(project, photos) {
    const headerText = `FOTOGRAFÍAS DE FECHA ${formatVisitDate(project.date)}`;
    const rows = [
      new docx.TableRow({
        cantSplit: true,
        children: [cell([paragraph(headerText, {
          alignment: docx.AlignmentType.CENTER,
          children: [new docx.TextRun({ text: headerText, bold: true, font: FONT, size: 16, color: "233C35" })]
        })], {
          columnSpan: 2,
          width: { size: CONTENT_WIDTH, type: docx.WidthType.DXA },
          shading: { fill: "EAF2ED" }
        })]
      })
    ];

    for (let index = 0; index < photos.length; index += 1) {
      rows.push(imageRow(await imageRun(photos[index]), index > 0 && index % 2 === 0));
      rows.push(captionRow(photos[index], index, project));
    }

    const table = new docx.Table({
      width: { size: CONTENT_WIDTH, type: docx.WidthType.DXA },
      columnWidths: [LABEL_WIDTH, CONTENT_WIDTH - LABEL_WIDTH],
      rows,
      margins: { top: 50, bottom: 50, left: 70, right: 70 }
    });
    const appendixTitle = project.appendixTitle || "ANEXO N.° 1 - FOTOGRAFÍAS";
    const metadata = [
      paragraph(`Código: ${project.code}`),
      paragraph(`Establecimiento / unidad: ${project.establishment || project.companyName}`),
      ...(project.vehiclePlates ? [paragraph(`Placa(s): ${project.vehiclePlates}`)] : []),
      paragraph(`Fiscalizador: ${project.inspector}`),
      paragraph(`Lugar: ${project.place}`, { spacing: { before: 0, after: 120 } })
    ];

    const document = new docx.Document({
      styles: {
        default: {
          document: { run: { font: FONT, size: 16, color: "233C35" },
            paragraph: { spacing: { before: 0, after: 40 } } }
        }
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
          }
        },
        children: [
          paragraph(appendixTitle, {
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 0, after: 150 },
            children: [new docx.TextRun({ text: appendixTitle, bold: true, font: FONT, size: 24, color: "173A35" })]
          }),
          ...metadata,
          table
        ]
      }]
    });
    return docx.Packer.toBlob(document);
  }

  async function download(project, photos) {
    const blob = await createDocument(project, photos);
    const code = project.code.normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9_-]+/g, "");
    const filename = `Anexo2_${code || "proyecto"}.docx`;
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  window.RegistroExportar = { createDocument, download, formatVisitDate };
})();
