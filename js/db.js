/* Persistencia local: proyectos con su banco y fotos Blob en transacciones IndexedDB. */
(function () {
  "use strict";

  const DB_NAME = "registro-fotografico-fiscalizacion-tarea03";
  const LEGACY_DB_NAME = "registro-fotografico-fiscalizacion";
  const DB_VERSION = 1;
  let connection;

  async function migrateLegacyData(database) {
    const migratedKey = "registro-fotografico-tarea03-datos-migrados";
    if (localStorage.getItem(migratedKey)) return;

    const legacyRequest = indexedDB.open(LEGACY_DB_NAME);
    const legacy = await new Promise((resolve, reject) => {
      legacyRequest.onsuccess = () => resolve(legacyRequest.result);
      legacyRequest.onerror = () => reject(legacyRequest.error || new Error("No se pudo consultar la base de datos anterior."));
      legacyRequest.onblocked = () => reject(new Error("La lectura de los proyectos anteriores está bloqueada por otra pestaña."));
    });

    try {
      if (!legacy.objectStoreNames.contains("projects") || !legacy.objectStoreNames.contains("photos")) {
        localStorage.setItem(migratedKey, "true");
        return;
      }
      const readTransaction = legacy.transaction(["projects", "photos"], "readonly");
      const readDone = transactionDone(readTransaction);
      const projectsRequest = requestResult(readTransaction.objectStore("projects").getAll());
      const photosRequest = requestResult(readTransaction.objectStore("photos").getAll());
      const [projects, photos] = await Promise.all([projectsRequest, photosRequest]);
      await readDone;
      if (!projects.length && !photos.length) {
        localStorage.setItem(migratedKey, "true");
        return;
      }

      const writeTransaction = database.transaction(["proyectos", "fotos"], "readwrite");
      const writeDone = transactionDone(writeTransaction);
      const projectStore = writeTransaction.objectStore("proyectos");
      projects.forEach((project) => projectStore.put({
        ...project,
        descriptionCatalog: Array.isArray(project.descriptionCatalog) ? project.descriptionCatalog : []
      }));
      const photoStore = writeTransaction.objectStore("fotos");
      photos.filter((photo) => photo.blob).forEach((photo) => photoStore.put(photo));
      await writeDone;
      localStorage.setItem(migratedKey, "true");
    } finally {
      legacy.close();
    }
  }

  function openDatabase() {
    if (connection) return connection;

    connection = new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains("proyectos")) {
          database.createObjectStore("proyectos", { keyPath: "id" });
        }
        if (!database.objectStoreNames.contains("fotos")) {
          const photoStore = database.createObjectStore("fotos", { keyPath: "id" });
          photoStore.createIndex("projectId", "projectId", { unique: false });
        }
      };
      request.onsuccess = () => {
        connection = request.result;
        connection.onversionchange = () => {
          connection.close();
          connection = null;
        };
        migrateLegacyData(connection).then(() => resolve(connection)).catch((error) => {
          connection.close();
          connection = null;
          reject(error);
        });
      };
      request.onerror = () => {
        connection = null;
        reject(request.error || new Error("No se pudo abrir la base de datos local."));
      };
      request.onblocked = () => {
        connection = null;
        reject(new Error("La actualización de la base de datos está bloqueada. Cierra las demás pestañas de la aplicación."));
      };
    });

    return connection;
  }

  function transactionDone(transaction) {
    return new Promise((resolve, reject) => {
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("La operación de almacenamiento falló."));
      transaction.onabort = () => reject(transaction.error || new Error("La operación de almacenamiento fue cancelada."));
    });
  }

  function requestResult(request) {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("No se pudo leer el almacenamiento local."));
    });
  }

  async function saveProject(project) {
    const database = await openDatabase();
    const transaction = database.transaction("proyectos", "readwrite");
    const done = transactionDone(transaction);
    transaction.objectStore("proyectos").put({
      descriptionCatalog: [],
      ...project,
      updatedAt: new Date().toISOString()
    });
    await done;
    return project;
  }

  async function getProjects() {
    const database = await openDatabase();
    const transaction = database.transaction("proyectos", "readonly");
    const done = transactionDone(transaction);
    const result = await requestResult(transaction.objectStore("proyectos").getAll());
    await done;
    return result;
  }

  async function getProject(id) {
    const database = await openDatabase();
    const transaction = database.transaction("proyectos", "readonly");
    const done = transactionDone(transaction);
    const result = await requestResult(transaction.objectStore("proyectos").get(id));
    await done;
    return result;
  }

  async function getPhotos(projectId) {
    const database = await openDatabase();
    const transaction = database.transaction("fotos", "readonly");
    const done = transactionDone(transaction);
    const photos = await requestResult(transaction.objectStore("fotos").index("projectId").getAll(projectId));
    await done;
    return photos.sort((first, second) => first.position - second.position);
  }

  async function savePhotos(photos) {
    if (!photos.length) return photos;
    const database = await openDatabase();
    const transaction = database.transaction("fotos", "readwrite");
    const done = transactionDone(transaction);
    const store = transaction.objectStore("fotos");
    photos.forEach((photo) => store.put(photo));
    await done;
    return photos;
  }

  async function savePhoto(photo) {
    return savePhotos([photo]);
  }

  async function deletePhoto(id, reorderedPhotos) {
    const database = await openDatabase();
    const transaction = database.transaction("fotos", "readwrite");
    const done = transactionDone(transaction);
    const store = transaction.objectStore("fotos");
    store.delete(id);
    reorderedPhotos.forEach((photo) => store.put(photo));
    await done;
  }

  async function deleteProject(id) {
    const database = await openDatabase();
    const transaction = database.transaction(["proyectos", "fotos"], "readwrite");
    const done = transactionDone(transaction);
    const photoStore = transaction.objectStore("fotos");
    const request = photoStore.index("projectId").openCursor(id);
    request.onsuccess = () => {
      const cursor = request.result;
      if (cursor) {
        cursor.delete();
        cursor.continue();
      }
    };
    transaction.objectStore("proyectos").delete(id);
    await done;
  }

  window.RegistroDB = {
    openDatabase,
    saveProject,
    getProjects,
    getProject,
    getPhotos,
    savePhotos,
    savePhoto,
    deletePhoto,
    deleteProject
  };
})();
