/* Eventos de captura, banco de descripciones, proyectos y vista previa del Word. */
(function () {
  "use strict";

  const elements = {
    message: document.getElementById("mensaje"),
    projectForm: document.getElementById("project-form"),
    projectId: document.getElementById("project-id"),
    projectSelect: document.getElementById("project-select"),
    appendixTitle: document.getElementById("appendix-title"),
    companyName: document.getElementById("company-name"),
    establishment: document.getElementById("project-establishment"),
    vehiclePlates: document.getElementById("vehicle-plates"),
    projectCode: document.getElementById("project-code"),
    projectDate: document.getElementById("project-date"),
    projectPlace: document.getElementById("project-place"),
    projectInspector: document.getElementById("project-inspector"),
    saveProject: document.getElementById("save-project"),
    newProject: document.getElementById("new-project"),
    deleteProject: document.getElementById("delete-project"),
    takePhoto: document.getElementById("take-photo"),
    cameraInput: document.getElementById("camera-input"),
    choosePhotos: document.getElementById("choose-photos"),
    imagesInput: document.getElementById("images-input"),
    photoRows: document.getElementById("photo-rows"),
    photoTemplate: document.getElementById("photo-row-template"),
    emptyList: document.getElementById("empty-list"),
    photoCount: document.getElementById("photo-count"),
    projectSummary: document.getElementById("project-summary"),
    assignmentCount: document.getElementById("assignment-count"),
    photoFilter: document.getElementById("photo-filter"),
    moveUp: document.getElementById("move-up"),
    moveDown: document.getElementById("move-down"),
    photoPlaceholder: document.getElementById("photo-placeholder"),
    photoDetail: document.getElementById("photo-detail"),
    selectedImage: document.getElementById("selected-image"),
    selectedFile: document.getElementById("selected-file"),
    descriptionHelp: document.getElementById("description-help"),
    descriptionSearch: document.getElementById("description-search"),
    clearSearch: document.getElementById("clear-search"),
    descriptionCount: document.getElementById("description-count"),
    descriptionSelect: document.getElementById("description-select"),
    photoDescription: document.getElementById("photo-description"),
    addDescription: document.getElementById("add-description"),
    descriptionDialog: document.getElementById("description-dialog"),
    descriptionFile: document.getElementById("description-file"),
    descriptionLines: document.getElementById("description-lines"),
    saveDescriptionBank: document.getElementById("save-description-bank"),
    returnPending: document.getElementById("return-pending"),
    deletePhoto: document.getElementById("delete-photo"),
    appendixProjectSummary: document.getElementById("appendix-project-summary"),
    appendixPreview: document.getElementById("appendix-preview"),
    exportButton: document.getElementById("export-docx"),
    pendingDialog: document.getElementById("pending-dialog"),
    pendingCount: document.getElementById("pending-count"),
    pendingGrid: document.getElementById("pending-grid"),
    discardPending: document.getElementById("discard-pending"),
    confirmPhotos: document.getElementById("confirm-photos")
  };

  const projectFields = [
    elements.appendixTitle,
    elements.companyName,
    elements.establishment,
    elements.vehiclePlates,
    elements.projectCode,
    elements.projectDate,
    elements.projectPlace,
    elements.projectInspector
  ];

  let activeProject = null;
  let photos = [];
  let pendingPhotos = [];
  let activePhotoId = null;
  let descriptionCatalog = [];
  let messageTimer;
  const photoUrls = new Map();
  let invalidPhotoIds = [];
  let projectSaveQueue = Promise.resolve();
  let exporting = false;

  function createId() {
    if (typeof crypto.randomUUID === "function") return crypto.randomUUID();
    const random = new Uint32Array(2);
    crypto.getRandomValues(random);
    return `${Date.now().toString(36)}-${random[0].toString(36)}${random[1].toString(36)}`;
  }

  function showMessage(text, kind) {
    window.clearTimeout(messageTimer);
    elements.message.replaceChildren(document.createTextNode(text));
    elements.message.className = `notice notice-${kind}`;
    elements.message.hidden = false;
    messageTimer = window.setTimeout(() => { elements.message.hidden = true; }, 6500);
  }

  function showValidationErrors(errors) {
    window.clearTimeout(messageTimer);
    const heading = document.createElement("strong");
    heading.textContent = "Corrige lo siguiente antes de generar el Word:";
    const list = document.createElement("ul");
    errors.forEach((error) => {
      const item = document.createElement("li");
      item.textContent = error.message;
      list.append(item);
    });
    elements.message.replaceChildren(heading, list);
    elements.message.className = "notice notice-error";
    elements.message.hidden = false;
  }

  function reportError(error, context) {
    console.error(context, error);
    showMessage(`${context} ${error && error.message ? error.message : "Inténtalo nuevamente."}`, "error");
  }

  function projectFromForm() {
    return {
      id: elements.projectId.value || createId(),
      appendixTitle: elements.appendixTitle.value.trim(),
      companyName: elements.companyName.value.trim(),
      establishment: elements.establishment.value.trim(),
      vehiclePlates: elements.vehiclePlates.value.trim(),
      code: elements.projectCode.value.trim(),
      date: elements.projectDate.value,
      place: elements.projectPlace.value.trim(),
      inspector: elements.projectInspector.value.trim(),
      descriptionCatalog: [...descriptionCatalog],
      updatedAt: new Date().toISOString()
    };
  }

  function projectHasData(project) {
    return Boolean(project.code || project.date || project.place || project.inspector
      || project.companyName || project.establishment || project.vehiclePlates);
  }

  function fillProjectForm(project) {
    elements.projectId.value = project ? project.id : "";
    elements.appendixTitle.value = project ? (project.appendixTitle || "ANEXO N.° 1 - FOTOGRAFÍAS") : "ANEXO N.° 1 - FOTOGRAFÍAS";
    elements.companyName.value = project ? (project.companyName || "") : "";
    elements.establishment.value = project ? (project.establishment || "") : "";
    elements.vehiclePlates.value = project ? (project.vehiclePlates || "") : "";
    elements.projectCode.value = project ? (project.code || "") : "";
    elements.projectDate.value = project ? (project.date || "") : "";
    elements.projectPlace.value = project ? (project.place || "") : "";
    elements.projectInspector.value = project ? (project.inspector || "") : "";
    projectFields.forEach((field) => field.classList.remove("invalid"));
  }

  async function refreshProjectOptions(selectedId) {
    const projects = await RegistroDB.getProjects();
    projects.sort((first, second) => first.code.localeCompare(second.code, "es"));
    elements.projectSelect.replaceChildren(new Option("Nuevo proyecto", ""));
    projects.forEach((project) => {
      elements.projectSelect.add(new Option(`${project.code || "Sin código"} · ${project.establishment || project.companyName || "Proyecto"}`, project.id));
    });
    elements.projectSelect.value = selectedId || "";
    elements.deleteProject.disabled = !selectedId;
  }

  function activePhoto() {
    return photos.find((photo) => photo.id === activePhotoId) || null;
  }

  function photoUrl(photo) {
    if (!photoUrls.has(photo.id)) photoUrls.set(photo.id, URL.createObjectURL(photo.blob));
    return photoUrls.get(photo.id);
  }

  function revokePhotoUrls() {
    photoUrls.forEach((url) => URL.revokeObjectURL(url));
    photoUrls.clear();
  }

  function visiblePhotos() {
    if (elements.photoFilter.value === "pending") return photos.filter((photo) => !photo.description.trim());
    if (elements.photoFilter.value === "assigned") return photos.filter((photo) => photo.description.trim());
    return photos;
  }

  function renderedDescription(description, project) {
    const currentProject = project || projectFromForm();
    return description
      .replace(/\{RAZON_SOCIAL\}/gi, currentProject.companyName || currentProject.establishment || "")
      .replace(/\{PLACAS\}/gi, currentProject.vehiclePlates || "");
  }

  function updateExportButton() {
    elements.exportButton.disabled = exporting;
  }

  function renderList() {
    elements.photoRows.replaceChildren();
    const visible = visiblePhotos();
    const assigned = photos.filter((photo) => photo.description.trim());
    elements.photoCount.textContent = String(photos.length);
    elements.assignmentCount.textContent = `${photos.length - assigned.length} pendientes · ${assigned.length} asignadas`;
    elements.emptyList.hidden = visible.length > 0;

    visible.forEach((photo) => {
      const row = elements.photoTemplate.content.firstElementChild.cloneNode(true);
      row.dataset.photoId = photo.id;
      row.classList.toggle("selected", photo.id === activePhotoId);
      row.classList.toggle("invalid", invalidPhotoIds.includes(photo.id));
      const image = row.querySelector("img");
      const imageUrl = photoUrl(photo);
      image.src = imageUrl;
      image.dataset.objectUrl = imageUrl;
      image.alt = `Miniatura ${photo.originalName || "de fotografía"}`;
      row.querySelector(".filename").textContent = photo.originalName || "Fotografía";
      row.querySelector(".table-file-date").textContent = photo.addedAt
        ? new Intl.DateTimeFormat("es-PE", { dateStyle: "short", timeStyle: "short" }).format(new Date(photo.addedAt))
        : "—";
      const number = photos.indexOf(photo) + 1;
      row.querySelector(".word-number").textContent = `Foto N.º ${String(number).padStart(2, "0")}`;
      const status = row.querySelector(".photo-status");
      status.className = `photo-status ${photo.description.trim() ? "assigned" : "pending"}`;
      status.textContent = photo.description.trim() ? "Asignada" : "Pendiente";
      elements.photoRows.append(row);
    });

    const index = photos.findIndex((photo) => photo.id === activePhotoId);
    elements.moveUp.disabled = index <= 0;
    elements.moveDown.disabled = index < 0 || index >= photos.length - 1;
    renderDetail();
    renderAppendix();
    updateExportButton();
  }

  function availableDescriptions() {
    const used = new Set(photos.map((photo) => photo.description.trim()).filter(Boolean));
    return descriptionCatalog.filter((description) => !used.has(description));
  }

  function renderDescriptionOptions() {
    const query = elements.descriptionSearch.value.trim().toLocaleLowerCase("es");
    const available = availableDescriptions();
    const filtered = available.filter((description) =>
      description.toLocaleLowerCase("es").includes(query));
    elements.descriptionSelect.replaceChildren();
    filtered.forEach((description) => elements.descriptionSelect.add(new Option(
      renderedDescription(description),
      description
    )));
    elements.descriptionCount.textContent = `${available.length} disponible${available.length === 1 ? "" : "s"}`;
    elements.descriptionHelp.textContent = descriptionCatalog.length
      ? "Al elegir una descripción se asigna a esta fotografía."
      : "Agrega descripciones al banco del proyecto para asignarlas aquí.";
  }

  function renderDetail() {
    const photo = activePhoto();
    elements.photoPlaceholder.hidden = Boolean(photo);
    elements.photoDetail.hidden = !photo;
    if (!photo) {
      elements.selectedImage.removeAttribute("src");
      elements.photoDescription.value = "";
      renderDescriptionOptions();
      return;
    }

    elements.selectedImage.src = photoUrl(photo);
    const added = new Date(photo.addedAt);
    elements.selectedFile.textContent = `${photo.originalName || "Fotografía"} · ${new Intl.DateTimeFormat("es-PE", { dateStyle: "medium", timeStyle: "short" }).format(added)} · Fecha de incorporación; no certifica la fecha original.`;
    elements.photoDescription.value = photo.description;
    elements.returnPending.disabled = !photo.description.trim();
    renderDescriptionOptions();
  }

  function renderAppendix() {
    const project = projectFromForm();
    elements.appendixProjectSummary.textContent =
      `${project.code || "Sin código"} · ${project.establishment || project.companyName || "Sin establecimiento"}`;
    elements.appendixPreview.replaceChildren();
    if (!photos.length) {
      const empty = document.createElement("p");
      empty.className = "appendix-empty";
      empty.textContent = "Agrega fotografías para ver aquí la secuencia del anexo.";
      elements.appendixPreview.append(empty);
      return;
    }
    photos.forEach((photo, index) => {
      const item = document.createElement("article");
      item.className = "appendix-item";
      const image = document.createElement("img");
      image.alt = `Foto N.º ${String(index + 1).padStart(2, "0")}`;
      const url = photoUrl(photo);
      image.src = url;
      image.dataset.objectUrl = url;
      const copy = document.createElement("div");
      const number = document.createElement("strong");
      number.textContent = `Foto N.º ${String(index + 1).padStart(2, "0")}`;
      const description = document.createElement("p");
      description.textContent = photo.description.trim()
        ? renderedDescription(photo.description, project)
        : "Pendiente: asigna una descripción para incluir esta foto en el Word.";
      copy.append(number, description);
      item.append(image, copy);
      elements.appendixPreview.append(item);
    });
  }

  async function persistProject(showSuccess) {
    const project = projectFromForm();
    if (!projectHasData(project)) {
      if (showSuccess) showMessage("Completa algún dato antes de guardar el proyecto.", "info");
      return null;
    }
    elements.projectId.value = project.id;
    projectSaveQueue = projectSaveQueue.catch(() => {}).then(async () => {
      await RegistroDB.saveProject(project);
      if (elements.projectId.value === project.id) {
        activeProject = project;
        descriptionCatalog = [...project.descriptionCatalog];
        localStorage.setItem("registro-fotografico-ultimo-proyecto", project.id);
        await refreshProjectOptions(project.id);
        updateProjectSummary();
      }
      return project;
    });
    try {
      const saved = await projectSaveQueue;
      if (showSuccess) showMessage("Proyecto guardado correctamente.", "success");
      return saved;
    } catch (error) {
      reportError(error, "No se pudo guardar el proyecto.");
      throw error;
    }
  }

  function updateProjectSummary() {
    const project = activeProject;
    elements.projectSummary.textContent = project
      ? `${project.code || "Sin código"} · ${project.establishment || project.companyName || "Proyecto activo"}`
      : "Crea o abre un proyecto para comenzar.";
    elements.deleteProject.disabled = !project;
    renderAppendix();
  }

  async function openProject(project) {
    revokePhotoUrls();
    activeProject = project;
    descriptionCatalog = project && Array.isArray(project.descriptionCatalog) ? [...project.descriptionCatalog] : [];
    activePhotoId = null;
    invalidPhotoIds = [];
    if (project) localStorage.setItem("registro-fotografico-ultimo-proyecto", project.id);
    else localStorage.removeItem("registro-fotografico-ultimo-proyecto");
    elements.photoFilter.value = "all";
    photos = project ? await RegistroDB.getPhotos(project.id) : [];
    if (photos.length) activePhotoId = photos[0].id;
    updateProjectSummary();
    renderList();
  }

  function clearPending() {
    pendingPhotos.forEach((photo) => URL.revokeObjectURL(photo.previewUrl));
    pendingPhotos = [];
    elements.pendingGrid.replaceChildren();
    elements.cameraInput.value = "";
    elements.imagesInput.value = "";
    if (elements.pendingDialog.open) elements.pendingDialog.close();
  }

  async function ensureActiveProject() {
    if (activeProject) return activeProject;
    const saved = await persistProject(false);
    if (!saved) throw new Error("Completa al menos un dato del proyecto antes de añadir fotografías.");
    return saved;
  }

  async function prepareSelection(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) {
      showMessage("No se seleccionó ninguna foto. Puedes continuar y agregarla después.", "info");
      return;
    }
    const prepared = [];
    try {
      await ensureActiveProject();
      showMessage(`Preparando ${files.length} fotografía(s)…`, "info");
      for (let index = 0; index < files.length; index += 3) {
        prepared.push(...await Promise.all(files.slice(index, index + 3)
          .map((file) => RegistroImagenes.prepareImage(file))));
      }
      pendingPhotos = prepared.map((photo) => ({ ...photo, description: "" }));
      renderPending();
      elements.pendingDialog.showModal();
      showMessage("Revisa las imágenes y confirma o descarta la selección.", "info");
    } catch (error) {
      prepared.forEach((photo) => URL.revokeObjectURL(photo.previewUrl));
      pendingPhotos.forEach((photo) => URL.revokeObjectURL(photo.previewUrl));
      pendingPhotos = [];
      reportError(error, "No se pudieron preparar las imágenes.");
    }
  }

  function renderPending() {
    elements.pendingGrid.replaceChildren();
    elements.pendingCount.textContent = `${pendingPhotos.length} fotografía(s). Escribe una descripción inicial si lo necesitas.`;
    pendingPhotos.forEach((photo, index) => {
      const item = document.createElement("article");
      item.className = "pending-item";
      const image = document.createElement("img");
      image.src = photo.previewUrl;
      image.alt = `Vista previa ${photo.originalName}`;
      const filename = document.createElement("span");
      filename.textContent = photo.originalName;
      const description = document.createElement("textarea");
      description.rows = 2;
      description.maxLength = 1000;
      description.placeholder = `Descripción inicial para foto ${index + 1} (opcional)`;
      description.value = photo.description;
      description.dataset.pendingIndex = String(index);
      item.append(image, filename, description);
      elements.pendingGrid.append(item);
    });
  }

  async function confirmPendingPhotos() {
    if (!pendingPhotos.length) return;
    elements.confirmPhotos.disabled = true;
    try {
      await ensureActiveProject();
      const start = photos.length;
      const records = pendingPhotos.map((pending, index) => ({
        id: createId(),
        projectId: activeProject.id,
        blob: pending.blob,
        description: pending.description.trim(),
        position: start + index,
        originalName: pending.originalName,
        width: pending.width,
        height: pending.height,
        addedAt: new Date().toISOString()
      }));
      await RegistroDB.savePhotos(records);
      pendingPhotos.forEach((photo) => URL.revokeObjectURL(photo.previewUrl));
      pendingPhotos = [];
      elements.pendingDialog.close();
      photos = await RegistroDB.getPhotos(activeProject.id);
      activePhotoId = records[0].id;
      renderList();
      showMessage(`${records.length} fotografía(s) guardada(s) correctamente.`, "success");
    } catch (error) {
      reportError(error, "No se pudieron guardar las fotografías.");
    } finally {
      elements.confirmPhotos.disabled = false;
    }
  }

  async function savePhotoDescription(photo, description) {
    const previous = photo.description;
    photo.description = description.trim();
    try {
      await RegistroDB.savePhoto(photo);
      if (photo.description.trim()) {
        invalidPhotoIds = invalidPhotoIds.filter((id) => id !== photo.id);
      }
      renderList();
      return true;
    } catch (error) {
      photo.description = previous;
      renderList();
      reportError(error, "No se pudo guardar la descripción.");
      return false;
    }
  }

  async function moveSelectedPhoto(direction) {
    const index = photos.findIndex((photo) => photo.id === activePhotoId);
    const next = index + direction;
    if (index < 0 || next < 0 || next >= photos.length) return;
    const reordered = [...photos];
    [reordered[index], reordered[next]] = [reordered[next], reordered[index]];
    reordered.forEach((photo, position) => { photo.position = position; });
    try {
      await RegistroDB.savePhotos(reordered);
      photos = reordered;
      renderList();
      showMessage("Orden del Word actualizado.", "success");
    } catch (error) {
      reportError(error, "No se pudo cambiar el orden de las fotos.");
    }
  }

  async function deleteSelectedPhoto() {
    const photo = activePhoto();
    if (!photo || !window.confirm(`¿Eliminar "${photo.originalName}"? Esta acción no se puede deshacer.`)) return;
    const reordered = photos.filter((entry) => entry.id !== photo.id)
      .map((entry, position) => ({ ...entry, position }));
    try {
      await RegistroDB.deletePhoto(photo.id, reordered);
      const url = photoUrls.get(photo.id);
      if (url) URL.revokeObjectURL(url);
      photoUrls.delete(photo.id);
      photos = reordered;
      activePhotoId = photos.length ? photos[Math.min(photo.position, photos.length - 1)].id : null;
      renderList();
      showMessage("Fotografía eliminada y orden actualizado.", "success");
    } catch (error) {
      reportError(error, "No se pudo eliminar la fotografía.");
    }
  }

  function getValidationErrors() {
    const project = projectFromForm();
    const errors = [];
    const required = [
      [elements.appendixTitle, project.appendixTitle, "el título del anexo"],
      [elements.projectCode, project.code, "el código"],
      [elements.projectDate, project.date, "la fecha de visita"],
      [elements.establishment, project.establishment || project.companyName, "el establecimiento o razón social"],
      [elements.projectPlace, project.place, "el lugar"],
      [elements.projectInspector, project.inspector, "el fiscalizador"]
    ];
    required.forEach(([field, value, label]) => {
      field.classList.toggle("invalid", !value);
      if (!value) errors.push({ message: `Completa ${label}.`, field });
    });
    if (!photos.length) errors.push({ message: "Agrega al menos una fotografía.", field: null });
    invalidPhotoIds = [];
    photos.forEach((photo, index) => {
      if (!photo.description.trim()) {
        invalidPhotoIds.push(photo.id);
        errors.push({ message: `Asigna una descripción a "${photo.originalName || `Foto ${index + 1}`}"; está pendiente.`, field: null, photoId: photo.id });
      }
    });
    return errors;
  }

  async function generateWord() {
    if (exporting) return;
    const errors = getValidationErrors();
    if (errors.length) {
      showValidationErrors(errors);
      const firstPending = photos.find((photo) => !photo.description.trim());
      if (firstPending) {
        elements.photoFilter.value = "all";
        activePhotoId = firstPending.id;
      }
      renderList();
      const firstError = errors.find((error) => error.field)?.field;
      if (firstError) firstError.focus();
      else if (firstPending) {
        elements.photoRows.querySelector(`[data-photo-id="${CSS.escape(firstPending.id)}"]`)?.scrollIntoView({ block: "nearest" });
      }
      return;
    }

    exporting = true;
    updateExportButton();
    try {
      const project = await persistProject(false);
      const orderedAssigned = photos.filter((photo) => photo.description.trim());
      await RegistroExportar.download(project, orderedAssigned);
      showMessage("Documento Word generado correctamente.", "success");
    } catch (error) {
      reportError(error, "No se pudo generar el documento Word.");
    } finally {
      exporting = false;
      updateExportButton();
    }
  }

  function addDescriptionLines(lines) {
    return [...new Set(lines.split(/\r?\n/).map((line) => line.trim()).filter(Boolean))];
  }

  async function saveDescriptionBank() {
    try {
      await ensureActiveProject();
      await projectSaveQueue;
      const additions = addDescriptionLines(elements.descriptionLines.value);
      if (!additions.length) throw new Error("Escribe o carga al menos una descripción.");
      const existing = new Set(descriptionCatalog);
      const newDescriptions = additions.filter((description) => !existing.has(description));
      if (!newDescriptions.length) throw new Error("Todas esas descripciones ya están en el banco.");
      descriptionCatalog.push(...newDescriptions);
      const project = projectFromForm();
      project.id = activeProject.id;
      project.descriptionCatalog = [...descriptionCatalog];
      await RegistroDB.saveProject(project);
      activeProject = project;
      await refreshProjectOptions(project.id);
      elements.descriptionLines.value = "";
      elements.descriptionDialog.close();
      renderDescriptionOptions();
      showMessage(`${newDescriptions.length} descripción(es) añadida(s) al banco.`, "success");
    } catch (error) {
      reportError(error, "No se pudo guardar el banco de descripciones.");
    }
  }

  elements.projectForm.addEventListener("submit", (event) => {
    event.preventDefault();
    persistProject(true).catch(() => {});
  });

  elements.projectForm.addEventListener("change", (event) => {
    if (!projectFields.includes(event.target)) return;
    event.target.classList.remove("invalid");
    renderAppendix();
    persistProject(true).catch(() => {});
  });

  elements.projectForm.addEventListener("input", (event) => {
    if (!projectFields.includes(event.target)) return;
    event.target.classList.remove("invalid");
    renderAppendix();
    renderDescriptionOptions();
  });

  elements.projectSelect.addEventListener("change", async () => {
    try {
      await projectSaveQueue;
      const project = elements.projectSelect.value ? await RegistroDB.getProject(elements.projectSelect.value) : null;
      fillProjectForm(project);
      await openProject(project);
      showMessage(project ? `Proyecto ${project.code || ""} abierto.` : "Formulario de nuevo proyecto listo.", "success");
    } catch (error) {
      reportError(error, "No se pudo abrir el proyecto.");
    }
  });

  elements.newProject.addEventListener("click", async () => {
    try {
      await projectSaveQueue;
      fillProjectForm(null);
      await openProject(null);
      elements.projectSelect.value = "";
      elements.projectCode.focus();
      showMessage("Completa los datos; se guardarán automáticamente al salir de cada campo.", "info");
    } catch (error) {
      reportError(error, "No se pudo iniciar un proyecto nuevo.");
    }
  });

  elements.deleteProject.addEventListener("click", async () => {
    if (!activeProject || !window.confirm(`¿Eliminar el proyecto "${activeProject.code || activeProject.establishment}" y todas sus fotos? Esta acción no se puede deshacer.`)) return;
    try {
      await RegistroDB.deleteProject(activeProject.id);
      localStorage.removeItem("registro-fotografico-ultimo-proyecto");
      fillProjectForm(null);
      await refreshProjectOptions("");
      await openProject(null);
      showMessage("Proyecto y fotografías eliminados.", "success");
    } catch (error) {
      reportError(error, "No se pudo eliminar el proyecto.");
    }
  });

  elements.takePhoto.addEventListener("click", () => {
    elements.cameraInput.click();
  });
  elements.choosePhotos.addEventListener("click", () => elements.imagesInput.click());
  elements.cameraInput.addEventListener("change", () => {
    prepareSelection(elements.cameraInput.files);
    elements.cameraInput.value = "";
  });
  elements.imagesInput.addEventListener("change", () => {
    prepareSelection(elements.imagesInput.files);
    elements.imagesInput.value = "";
  });
  [elements.cameraInput, elements.imagesInput].forEach((input) => {
    input.addEventListener("cancel", () => showMessage("Captura o selección cancelada. Puedes continuar sin agregar fotos.", "info"));
  });

  elements.pendingGrid.addEventListener("input", (event) => {
    if (!event.target.matches("textarea[data-pending-index]")) return;
    const photo = pendingPhotos[Number(event.target.dataset.pendingIndex)];
    if (photo) photo.description = event.target.value;
  });
  elements.confirmPhotos.addEventListener("click", confirmPendingPhotos);
  elements.discardPending.addEventListener("click", clearPending);
  elements.pendingDialog.addEventListener("close", () => {
    if (pendingPhotos.length) clearPending();
  });

  elements.photoRows.addEventListener("click", (event) => {
    const row = event.target.closest("tr[data-photo-id]");
    if (!row) return;
    activePhotoId = row.dataset.photoId;
    renderList();
  });
  elements.photoRows.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    const row = event.target.closest("tr[data-photo-id]");
    if (!row) return;
    event.preventDefault();
    activePhotoId = row.dataset.photoId;
    renderList();
  });
  elements.photoFilter.addEventListener("change", () => {
    if (!visiblePhotos().some((photo) => photo.id === activePhotoId)) activePhotoId = visiblePhotos()[0]?.id || null;
    renderList();
  });
  elements.moveUp.addEventListener("click", () => moveSelectedPhoto(-1));
  elements.moveDown.addEventListener("click", () => moveSelectedPhoto(1));

  elements.descriptionSearch.addEventListener("input", renderDescriptionOptions);
  elements.clearSearch.addEventListener("click", () => {
    elements.descriptionSearch.value = "";
    renderDescriptionOptions();
    elements.descriptionSearch.focus();
  });
  elements.descriptionSelect.addEventListener("change", async () => {
    const photo = activePhoto();
    const description = elements.descriptionSelect.value;
    if (!photo || !description) return;
    elements.photoDescription.value = description;
    if (await savePhotoDescription(photo, description)) {
      showMessage("Descripción asignada a la fotografía.", "success");
    }
  });
  elements.photoDescription.addEventListener("change", async () => {
    const photo = activePhoto();
    if (!photo || photo.description === elements.photoDescription.value.trim()) return;
    if (await savePhotoDescription(photo, elements.photoDescription.value)) {
      showMessage(photo.description ? "Descripción guardada; foto asignada." : "La fotografía quedó pendiente.", "success");
    }
  });
  elements.photoDescription.addEventListener("input", renderAppendix);
  elements.returnPending.addEventListener("click", async () => {
    const photo = activePhoto();
    if (!photo) return;
    if (await savePhotoDescription(photo, "")) showMessage("Fotografía devuelta a pendientes.", "success");
  });
  elements.deletePhoto.addEventListener("click", deleteSelectedPhoto);

  elements.addDescription.addEventListener("click", async () => {
    try {
      await ensureActiveProject();
      elements.descriptionDialog.showModal();
      elements.descriptionLines.focus();
    } catch (error) {
      reportError(error, "No se pudo abrir el banco de descripciones.");
    }
  });
  elements.descriptionFile.addEventListener("change", async () => {
    const file = elements.descriptionFile.files && elements.descriptionFile.files[0];
    if (!file) return;
    try {
      elements.descriptionLines.value = await file.text();
      showMessage("Archivo de texto cargado. Revisa las líneas y confirma.", "info");
    } catch (error) {
      reportError(error, "No se pudo leer el archivo de texto.");
    } finally {
      elements.descriptionFile.value = "";
    }
  });
  elements.saveDescriptionBank.addEventListener("click", saveDescriptionBank);
  elements.exportButton.addEventListener("click", generateWord);

  window.addEventListener("beforeunload", () => {
    revokePhotoUrls();
    pendingPhotos.forEach((photo) => URL.revokeObjectURL(photo.previewUrl));
  });

  (async function initialize() {
    try {
      await RegistroDB.openDatabase();
      const lastProjectId = localStorage.getItem("registro-fotografico-ultimo-proyecto");
      await refreshProjectOptions(lastProjectId || "");
      const project = lastProjectId ? await RegistroDB.getProject(lastProjectId) : null;
      if (lastProjectId && !project) localStorage.removeItem("registro-fotografico-ultimo-proyecto");
      fillProjectForm(project || null);
      await openProject(project || null);
    } catch (error) {
      reportError(error, "No se pudo iniciar el almacenamiento local.");
    }
  })();
})();
