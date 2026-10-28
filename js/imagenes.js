/* Normaliza la orientación, reduce a 1600 px y prepara JPEG y vista previa local. */
(function () {
  "use strict";

  const MAX_EDGE = 1600;
  const JPEG_QUALITY = 0.88;

  function loadImage(file) {
    return new Promise((resolve, reject) => {
      const source = URL.createObjectURL(file);
      const image = new Image();
      image.onload = () => resolve({
        width: image.naturalWidth,
        height: image.naturalHeight,
        draw: (context, width, height) => context.drawImage(image, 0, 0, width, height),
        close: () => URL.revokeObjectURL(source)
      });
      image.onerror = () => {
        URL.revokeObjectURL(source);
        reject(new Error("El archivo seleccionado no es una imagen compatible."));
      };
      image.src = source;
    });
  }

  async function decodeImage(file) {
    if (typeof createImageBitmap === "function") {
      try {
        const bitmap = await createImageBitmap(file, { imageOrientation: "from-image" });
        return {
          width: bitmap.width,
          height: bitmap.height,
          draw: (context, width, height) => context.drawImage(bitmap, 0, 0, width, height),
          close: () => bitmap.close()
        };
      } catch (error) {
        try {
          const bitmap = await createImageBitmap(file);
          return {
            width: bitmap.width,
            height: bitmap.height,
            draw: (context, width, height) => context.drawImage(bitmap, 0, 0, width, height),
            close: () => bitmap.close()
          };
        } catch {
          return loadImage(file);
        }
      }
    }
    return loadImage(file);
  }

  async function prepareImage(file) {
    if (!file || !file.type || !file.type.startsWith("image/")) {
      throw new Error("El archivo seleccionado no tiene un formato de imagen compatible.");
    }

    const image = await decodeImage(file);
    try {
      const scale = Math.min(1, MAX_EDGE / Math.max(image.width, image.height));
      const width = Math.max(1, Math.round(image.width * scale));
      const height = Math.max(1, Math.round(image.height * scale));
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d", { alpha: false });
      if (!context) throw new Error("El navegador no pudo preparar esta imagen.");
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, width, height);
      image.draw(context, width, height);
      const blob = await new Promise((resolve, reject) => {
        canvas.toBlob((result) => {
          if (result) resolve(result);
          else reject(new Error("No se pudo convertir la imagen a JPEG."));
        }, "image/jpeg", JPEG_QUALITY);
      });
      return {
        blob,
        previewUrl: URL.createObjectURL(blob),
        width,
        height,
        originalName: file.name || "foto.jpg"
      };
    } finally {
      image.close();
    }
  }

  window.RegistroImagenes = { prepareImage, MAX_EDGE };
})();
