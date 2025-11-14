(function () {
  "use strict";

  // Datos de ejemplo. Puedes reemplazar/expandir según tu catálogo real.
  /**
   * Estructura por producto
   * {
   *   id: string,
   *   nombre: string,
   *   procesos: Array<{
   *     codigo: string,
   *     analisis: string,
   *     metodo: string,
   *     cMtra_g: number,
   *     cantidad: number,
   *     vrUnit1: number,
   *     vrUnit2: number,
   *     vrUnit3: number,
   *     vrUnitUSD: number
   *   }>
   * }
   */
  let productos = [
    {
      id: "A-ACEITE-EUCALIPTO-USP",
      nombre: "ACEITE DE EUCALIPTO USP",
      procesos: [
        { codigo: "AEU-001", analisis: "Identificación", metodo: "USP <197>", cMtra_g: 5, cantidad: 1, vrUnit1: 190000, vrUnit2: 180000, vrUnit3: 165000, vrUnitUSD: 190000 },
        { codigo: "AEU-002", analisis: "Índice de refracción", metodo: "USP <831>", cMtra_g: 3, cantidad: 1, vrUnit1: 240000, vrUnit2: 225000, vrUnit3: 210000, vrUnitUSD: 240000 },
        { codigo: "AEU-003", analisis: "Cromatografía GC", metodo: "USP <621>", cMtra_g: 2, cantidad: 1, vrUnit1: 880000, vrUnit2: 850000, vrUnit3: 820000, vrUnitUSD: 880000 }
      ]
    },
    {
      id: "A-ALCOHOL-ISOPROPILICO",
      nombre: "ALCOHOL ISOPROPÍLICO",
      procesos: [
        { codigo: "AIS-010", analisis: "Pureza", metodo: "GC", cMtra_g: 10, cantidad: 1, vrUnit1: 420000, vrUnit2: 400000, vrUnit3: 380000, vrUnitUSD: 420000 },
        { codigo: "AIS-011", analisis: "Identificación", metodo: "IR", cMtra_g: 4, cantidad: 1, vrUnit1: 170000, vrUnit2: 160000, vrUnit3: 150000, vrUnitUSD: 170000 }
      ]
    },
    {
      id: "B-BENZOATO-SODIO",
      nombre: "BENZOATO DE SODIO",
      procesos: [
        { codigo: "BZS-020", analisis: "Ensayo", metodo: "HPLC", cMtra_g: 2, cantidad: 1, vrUnit1: 720000, vrUnit2: 690000, vrUnit3: 650000, vrUnitUSD: 720000 },
        { codigo: "BZS-021", analisis: "Impurezas", metodo: "HPLC", cMtra_g: 2, cantidad: 1, vrUnit1: 1150000, vrUnit2: 1100000, vrUnit3: 1050000, vrUnitUSD: 1150000 }
      ]
    },
    {
      id: "C-CAFEINA",
      nombre: "CAFEÍNA",
      procesos: [
        { codigo: "CAF-030", analisis: "Ensayo", metodo: "UV-Vis", cMtra_g: 1, cantidad: 1, vrUnit1: 260000, vrUnit2: 245000, vrUnit3: 230000, vrUnitUSD: 260000 },
        { codigo: "CAF-031", analisis: "Identificación", metodo: "IR", cMtra_g: 1, cantidad: 1, vrUnit1: 180000, vrUnit2: 170000, vrUnit3: 160000, vrUnitUSD: 180000 }
      ]
    },
    {
      id: "E-ETANOL",
      nombre: "ETANOL",
      procesos: [
        { codigo: "ETA-050", analisis: "Pureza", metodo: "GC", cMtra_g: 8, cantidad: 1, vrUnit1: 400000, vrUnit2: 390000, vrUnit3: 370000, vrUnitUSD: 400000 }
      ]
    }
  ];

  const $alphabetNav = document.getElementById("alphabetNav");
  const $productsContainer = document.getElementById("productsContainer");
  const $noResults = document.getElementById("noResults");
  const $productsPagination = document.getElementById("productsPagination");
  const $selectedCount = document.getElementById("selectedCount");
  const $cartCount = document.getElementById("cartCount");
  const $openCartBtn = document.getElementById("openCartBtn");
  const $btnGeneratePDF = document.getElementById("btnGeneratePDF");
  const $clientName = document.getElementById("clientName");
  const $clientEmail = document.getElementById("clientEmail");
  const $quoteDate = document.getElementById("quoteDate");
  const $quoteClientId = document.getElementById("quoteClientId"); // ID del cliente en el modal de cotización
  const $quoteClientNit = document.getElementById("quoteClientNit");
  const $quoteClientContacto = document.getElementById("quoteClientContacto");
  const $contactoDropdown = document.getElementById("contactoDropdown");
  const $selectedContactos = document.getElementById("selectedContactos");
  const $quoteClientContactosSelected = document.getElementById("quoteClientContactosSelected");
  const $quoteClientCelular = document.getElementById("quoteClientCelular");
  const $quoteClientFormaPago = document.getElementById("quoteClientFormaPago");
  const $quoteDuracionAnalisis = document.getElementById("quoteDuracionAnalisis");
  const $quoteNota = document.getElementById("quoteNota");
  const $clientDropdown = document.getElementById("clientDropdown");
  
  // Variable para almacenar los contactos disponibles del cliente seleccionado
  let availableContactos = [];
  const $historyTableBody = document.querySelector("#historyTable tbody");
  const $noHistory = document.getElementById("noHistory");

  const $filterConsecutivo = document.getElementById("filterConsecutivo");
  const $filterClient = document.getElementById("filterClient");
  const $filterProduct = document.getElementById("filterProduct");
  const $filterFrom = document.getElementById("filterFrom");
  const $filterTo = document.getElementById("filterTo");
  const $btnApplyFilters = document.getElementById("btnApplyFilters");
  const $btnClearFilters = document.getElementById("btnClearFilters");
  const $btnExportXlsx = document.getElementById("btnExportXlsx");
  const $btnLoadExcel = document.getElementById("btnLoadExcel");
  const $excelInput = document.getElementById("excelInput");
  const $btnDeleteProducts = document.getElementById("btnDeleteProducts");
  const $productSearch = document.getElementById("productSearch");
  const $scrollTopBtn = document.getElementById("scrollTopBtn");
  
  // Modal eliminar productos
  const $deleteProductsModal = document.getElementById("deleteProductsModal");
  const $deleteProductsList = document.getElementById("deleteProductsList");
  const $btnSelectAllProducts = document.getElementById("btnSelectAllProducts");
  const $btnDeselectAllProducts = document.getElementById("btnDeselectAllProducts");
  const $btnConfirmDeleteProducts = document.getElementById("btnConfirmDeleteProducts");
  const $btnDeleteAllProducts = document.getElementById("btnDeleteAllProducts");
  const $deleteProductsCount = document.getElementById("deleteProductsCount");
  
  // Modal eliminar clientes
  const $deleteClientsModal = document.getElementById("deleteClientsModal");
  const $deleteClientsList = document.getElementById("deleteClientsList");
  const $btnSelectAllClients = document.getElementById("btnSelectAllClients");
  const $btnDeselectAllClients = document.getElementById("btnDeselectAllClients");
  const $btnConfirmDeleteClients = document.getElementById("btnConfirmDeleteClients");
  const $deleteClientsCount = document.getElementById("deleteClientsCount");
  const $btnDeleteClients = document.getElementById("btnDeleteClients");
  const $statsCanvas = document.getElementById("statsTopProducts");
  const $noStatsEl = document.getElementById("noStats");
  
  // Clientes
  const $clientsTableBody = document.getElementById("clientsTableBody");
  const $noClients = document.getElementById("noClients");
  const $btnAddClient = document.getElementById("btnAddClient");
  const $btnImportClients = document.getElementById("btnImportClients");
  const $btnExportClients = document.getElementById("btnExportClients");
  const $clientsExcelInput = document.getElementById("clientsExcelInput");
  const $clientModalTitle = document.getElementById("clientModalTitle");
  const $clientId = document.getElementById("clientId");
  const $clientNombre = document.getElementById("clientNombre");
  const $clientNit = document.getElementById("clientNit");
  const $clientCorreo = document.getElementById("clientCorreo");
  const $clientCelular = document.getElementById("clientCelular");
  const $clientFormaPago = document.getElementById("clientFormaPago");
  const $contactosContainer = document.getElementById("contactosContainer");
  const $btnAddContacto = document.getElementById("btnAddContacto");
  const $btnSaveClient = document.getElementById("btnSaveClient");
  const $clientFormModal = document.getElementById("clientFormModal");
  
  // Contactos
  const $contactosTableBody = document.getElementById("contactosTableBody");
  const $noContactos = document.getElementById("noContactos");
  const $btnAddContactoNew = document.getElementById("btnAddContacto");
  const $btnImportContactos = document.getElementById("btnImportContactos");
  const $btnExportContactos = document.getElementById("btnExportContactos");
  const $contactosExcelInput = document.getElementById("contactosExcelInput");
  const $contactoModalTitle = document.getElementById("contactoModalTitle");
  const $contactoId = document.getElementById("contactoId");
  const $contactoNombreCliente = document.getElementById("contactoNombreCliente");
  const $contactoNombre = document.getElementById("contactoNombre");
  const $contactoCorreo = document.getElementById("contactoCorreo");
  const $btnSaveContacto = document.getElementById("btnSaveContacto");
  const $contactoFormModal = document.getElementById("contactoFormModal");
  const $btnDeleteContactos = document.getElementById("btnDeleteContactos");
  const $deleteContactosModal = document.getElementById("deleteContactosModal");
  const $deleteContactosList = document.getElementById("deleteContactosList");
  const $btnSelectAllContactos = document.getElementById("btnSelectAllContactos");
  const $btnDeselectAllContactos = document.getElementById("btnDeselectAllContactos");
  const $btnConfirmDeleteContactos = document.getElementById("btnConfirmDeleteContactos");
  const $deleteContactosCount = document.getElementById("deleteContactosCount");
  
  // Login y usuarios
  const $loginPage = document.getElementById("loginPage");
  const $mainContent = document.getElementById("mainContent");
  const $loginForm = document.getElementById("loginForm");
  const $loginUsername = document.getElementById("loginUsername");
  const $loginPassword = document.getElementById("loginPassword");
  const $loginError = document.getElementById("loginError");
  const $logoutBtn = document.getElementById("logoutBtn");
  const $userInfo = document.getElementById("userInfo");
  
  // Gestión de usuarios
  const $usersTableBody = document.getElementById("usersTableBody");
  const $noUsers = document.getElementById("noUsers");
  const $btnAddUser = document.getElementById("btnAddUser");
  const $btnExportUsers = document.getElementById("btnExportUsers");
  const $userModalTitle = document.getElementById("userModalTitle");
  const $userId = document.getElementById("userId");
  const $userUsername = document.getElementById("userUsername");
  const $userName = document.getElementById("userName");
  const $userPassword = document.getElementById("userPassword");
  const $userRole = document.getElementById("userRole");
  const $userActive = document.getElementById("userActive");
  const $btnSaveUser = document.getElementById("btnSaveUser");
  const $userFormModal = document.getElementById("userFormModal");
  const $passwordRequired = document.getElementById("passwordRequired");
  const $passwordHint = document.getElementById("passwordHint");
  
  let logoAsset = null; // { el?: HTMLImageElement, dataUrl?: string }
  let topProductsChart = null;
  let currentUser = null;

  let selectedProductIds = new Set();
  // Map para almacenar el valor unitario seleccionado por producto
  // Clave: productId, Valor: 'vrUnit1' | 'vrUnit2' | 'vrUnit3' | 'vrUnitUSD'
  let selectedUnitValues = new Map();
  let currentLetter = "A";
  let currentPage = 1;
  const productsPerPage = 10;
  let currentSearchQuery = "";

  // ==================== SISTEMA DE AUTENTICACIÓN ====================

  function initDefaultUsers() {
    const users = getUsers();
    if (users.length === 0) {
      // Crear usuarios por defecto
      const defaultUsers = [
        {
          id: "USER-ADMIN-001",
          username: "admin",
          password: hashPassword("admin"),
          name: "Administrador",
          role: "admin",
          active: true,
          createdAt: new Date().toISOString()
        },
        {
          id: "USER-VENDEDOR-001",
          username: "vendedor",
          password: hashPassword("vendedor"),
          name: "Vendedor",
          role: "vendedor",
          active: true,
          createdAt: new Date().toISOString()
        }
      ];
      saveUsers(defaultUsers);
    }
  }

  function hashPassword(password) {
    // Hash simple (en producción usar bcrypt o similar)
    let hash = 0;
    for (let i = 0; i < password.length; i++) {
      const char = password.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32bit integer
    }
    return hash.toString();
  }

  function getUsers() {
    try {
      return JSON.parse(localStorage.getItem("olgroup_users") || "[]");
    } catch {
      return [];
    }
  }

  function saveUsers(users) {
    try {
      localStorage.setItem("olgroup_users", JSON.stringify(users || []));
    } catch {
      showAlert("No se pudo guardar los usuarios. Verifica el espacio disponible.", "error");
    }
  }

  function getCurrentUser() {
    try {
      const userStr = sessionStorage.getItem("olgroup_current_user");
      return userStr ? JSON.parse(userStr) : null;
    } catch {
      return null;
    }
  }

  function setCurrentUser(user) {
    try {
      if (user) {
        sessionStorage.setItem("olgroup_current_user", JSON.stringify(user));
      } else {
        sessionStorage.removeItem("olgroup_current_user");
      }
    } catch {}
  }

  function login(username, password) {
    const users = getUsers();
    const hashedPassword = hashPassword(password);
    
    // Primero verificar si el usuario existe y la contraseña es correcta
    const user = users.find(
      (u) => u.username.toLowerCase() === username.toLowerCase() && 
             u.password === hashedPassword
    );

    if (user) {
      // Si el usuario existe pero está inactivo
      if (user.active === false) {
        return { success: false, inactive: true };
      }
      // Si el usuario existe y está activo
      const { password: _, ...userWithoutPassword } = user;
      setCurrentUser(userWithoutPassword);
      return { success: true };
    }
    
    // Usuario no encontrado o contraseña incorrecta
    return { success: false, inactive: false };
  }

  function logout() {
    setCurrentUser(null);
    showLogin();
  }

  function showLogin() {
    if ($loginPage) $loginPage.classList.remove("d-none");
    if ($mainContent) $mainContent.classList.add("d-none");
    if ($loginForm) $loginForm.reset();
    if ($loginError) $loginError.classList.add("d-none");
  }

  function showMainContent() {
    currentUser = getCurrentUser();
    if (!currentUser) {
      showLogin();
      return;
    }

    if ($loginPage) $loginPage.classList.add("d-none");
    if ($mainContent) $mainContent.classList.remove("d-none");
    
    // Actualizar información del usuario
    if ($userInfo) {
      $userInfo.textContent = `${currentUser.name} (${currentUser.role === "admin" ? "Administrador" : "Vendedor"})`;
      // Agregar clase según el rol para diferenciar visualmente
      $userInfo.className = `small me-2 user-role-${currentUser.role}`;
    }

    // Aplicar clase de rol al body
    document.body.className = currentUser.role === "admin" ? "admin" : "";

    // Inicializar aplicación
    initApp();
  }

  function checkAuth() {
    currentUser = getCurrentUser();
    if (currentUser) {
      showMainContent();
    } else {
      showLogin();
    }
  }

  function init() {
    // Inicializar usuarios por defecto
    initDefaultUsers();

    // Verificar autenticación
    checkAuth();

    // Event listeners para login
    if ($loginForm) {
      $loginForm.addEventListener("submit", (e) => {
        e.preventDefault();
        const username = $loginUsername.value.trim();
        const password = $loginPassword.value.trim();

        if (!username || !password) {
          if ($loginError) {
            $loginError.textContent = "Por favor ingresa usuario y contraseña.";
            $loginError.classList.remove("d-none");
          }
          return;
        }

        const loginResult = login(username, password);
        if (loginResult.success) {
          showMainContent();
        } else {
          if ($loginError) {
            if (loginResult.inactive) {
              $loginError.textContent = "Usuario inactivo por administrador.";
            } else {
              $loginError.textContent = "Usuario o contraseña incorrectos.";
            }
            $loginError.classList.remove("d-none");
          }
        }
      });
    }

    if ($logoutBtn) {
      $logoutBtn.addEventListener("click", logout);
    }
  }

  function initApp() {
    // Fecha por defecto
    if (!$quoteDate.value) {
      $quoteDate.valueAsDate = new Date();
    }
    // Cargar catálogo previo si existe
    const persistedCatalog = loadCatalog();
    if (persistedCatalog && Array.isArray(persistedCatalog) && persistedCatalog.length > 0) {
      productos = persistedCatalog;
    }
    // Solo inicializar si estamos autenticados
    if (!currentUser) return;

    buildAlphabetNav();
    renderProductsByLetter(currentLetter);
    renderHistory();
    renderClients();
    
    // Renderizar usuarios solo si es admin
    if (currentUser.role === "admin") {
      renderUsers();
    }

    $btnGeneratePDF.addEventListener("click", onGeneratePDF);
    // Filtrado en tiempo real - debounce para campos de texto
    let filterTimeout;
    const applyFiltersDebounced = () => {
      clearTimeout(filterTimeout);
      filterTimeout = setTimeout(() => {
        renderHistory(getHistoryFilters());
      }, 300); // 300ms de delay para evitar demasiadas llamadas
    };

    // Event listeners para filtrado en tiempo real
    if ($filterConsecutivo) {
      $filterConsecutivo.addEventListener("input", applyFiltersDebounced);
    }
    if ($filterClient) {
      $filterClient.addEventListener("input", applyFiltersDebounced);
    }
    if ($filterProduct) {
      $filterProduct.addEventListener("input", applyFiltersDebounced);
    }
    if ($filterFrom) {
      $filterFrom.addEventListener("change", () => renderHistory(getHistoryFilters()));
    }
    if ($filterTo) {
      $filterTo.addEventListener("change", () => renderHistory(getHistoryFilters()));
    }

    // Mantener el botón "Aplicar filtros" por compatibilidad (aunque ya no es necesario)
    $btnApplyFilters.addEventListener("click", () => renderHistory(getHistoryFilters()));

    // Validar que el campo Duración de Análisis solo acepte números
    if ($quoteDuracionAnalisis) {
      // Prevenir escribir caracteres no numéricos
      $quoteDuracionAnalisis.addEventListener("keydown", function(e) {
        // Permitir teclas especiales: backspace, delete, tab, escape, enter, y teclas de navegación
        if ([8, 9, 27, 13, 46, 35, 36, 37, 38, 39, 40].indexOf(e.keyCode) !== -1 ||
            (e.keyCode === 65 && e.ctrlKey === true) || // Ctrl+A
            (e.keyCode >= 35 && e.keyCode <= 40)) { // Home, End, Left, Right, Up, Down
          return;
        }
        // Permitir solo números
        if ((e.shiftKey || (e.keyCode < 48 || e.keyCode > 57)) && (e.keyCode < 96 || e.keyCode > 105)) {
          e.preventDefault();
        }
      });
      
      // Remover cualquier carácter no numérico al escribir
      $quoteDuracionAnalisis.addEventListener("input", function(e) {
        // Remover cualquier carácter que no sea número
        this.value = this.value.replace(/[^0-9]/g, '');
      });
      
      // Prevenir pegar texto que contenga caracteres no numéricos
      $quoteDuracionAnalisis.addEventListener("paste", function(e) {
        e.preventDefault();
        const paste = (e.clipboardData || window.clipboardData).getData('text');
        const numbersOnly = paste.replace(/[^0-9]/g, '');
        this.value = numbersOnly;
      });
    }
    if ($btnClearFilters) $btnClearFilters.addEventListener("click", clearFilters);
    if ($btnExportXlsx) $btnExportXlsx.addEventListener("click", () => exportHistoryToXlsx(getHistoryFilters()));

    // Event listeners para clientes
    if ($btnAddClient) $btnAddClient.addEventListener("click", () => openClientModal());
    if ($btnSaveClient) $btnSaveClient.addEventListener("click", saveClient);
    if ($btnAddContacto) $btnAddContacto.addEventListener("click", () => addContactoInput());
    if ($btnImportClients) $btnImportClients.addEventListener("click", () => $clientsExcelInput.click());
    if ($btnDeleteClients) {
      $btnDeleteClients.addEventListener("click", openDeleteClientsModal);
    }
    if ($btnSelectAllClients) {
      $btnSelectAllClients.addEventListener("click", selectAllClientsToDelete);
    }
    if ($btnDeselectAllClients) {
      $btnDeselectAllClients.addEventListener("click", deselectAllClientsToDelete);
    }
    if ($btnConfirmDeleteClients) {
      $btnConfirmDeleteClients.addEventListener("click", deleteSelectedClients);
    }
    if ($clientsExcelInput) $clientsExcelInput.addEventListener("change", onClientsExcelSelected);
    if ($btnExportClients) $btnExportClients.addEventListener("click", exportClientsToExcel);
    
    // Resetear formulario de cliente al cerrar el modal
    const clientModalEl = document.getElementById("clientModal");
    if (clientModalEl && typeof bootstrap !== "undefined") {
      clientModalEl.addEventListener("hidden.bs.modal", () => {
        if ($clientFormModal) $clientFormModal.reset();
        if ($clientId) $clientId.value = "";
      });
    }

    // Event listeners para contactos
    if ($btnAddContactoNew) $btnAddContactoNew.addEventListener("click", () => openContactoModal());
    if ($btnSaveContacto) $btnSaveContacto.addEventListener("click", saveContacto);
    if ($btnImportContactos) $btnImportContactos.addEventListener("click", () => $contactosExcelInput.click());
    if ($btnDeleteContactos) {
      $btnDeleteContactos.addEventListener("click", openDeleteContactosModal);
    }
    if ($btnSelectAllContactos) {
      $btnSelectAllContactos.addEventListener("click", selectAllContactosToDelete);
    }
    if ($btnDeselectAllContactos) {
      $btnDeselectAllContactos.addEventListener("click", deselectAllContactosToDelete);
    }
    if ($btnConfirmDeleteContactos) {
      $btnConfirmDeleteContactos.addEventListener("click", deleteSelectedContactos);
    }
    if ($contactosExcelInput) $contactosExcelInput.addEventListener("change", onContactosExcelSelected);
    if ($btnExportContactos) $btnExportContactos.addEventListener("click", exportContactosToExcel);
    
    // Resetear formulario de contacto al cerrar el modal
    const contactoModalEl = document.getElementById("contactoModal");
    if (contactoModalEl && typeof bootstrap !== "undefined") {
      contactoModalEl.addEventListener("hidden.bs.modal", () => {
        if ($contactoFormModal) $contactoFormModal.reset();
        if ($contactoId) $contactoId.value = "";
      });
    }

    // Event listeners para usuarios (solo admin)
    if (currentUser.role === "admin") {
      if ($btnAddUser) $btnAddUser.addEventListener("click", () => openUserModal());
      if ($btnSaveUser) $btnSaveUser.addEventListener("click", saveUser);
      if ($btnExportUsers) $btnExportUsers.addEventListener("click", exportUsersToExcel);
      
      // Resetear formulario de usuario al cerrar el modal
      const userModalEl = document.getElementById("userModal");
      if (userModalEl && typeof bootstrap !== "undefined") {
        userModalEl.addEventListener("hidden.bs.modal", () => {
          if ($userFormModal) $userFormModal.reset();
          if ($userId) $userId.value = "";
          if ($passwordRequired) $passwordRequired.classList.remove("d-none");
          if ($passwordHint) $passwordHint.classList.add("d-none");
        });
      }
    }

    // Ocultar todas las secciones primero y luego mostrar solo la activa
    hideAllSections();
    showSection("viewProducts");
    
    // Asegurar que las secciones admin-only estén ocultas si no es admin
    // Esto se hace después de mostrar la sección activa para no interferir
    enforceRoleBasedVisibility();
    
    // Asegurar que la pestaña "Productos y procesos" esté activa
    const nav = document.getElementById("mainNav");
    if (nav) {
      const links = nav.querySelectorAll(".nav-link");
      links.forEach((l) => l.classList.remove("active"));
      const productsLink = Array.from(links).find((l) => l.getAttribute("data-target") === "#viewProducts");
      if (productsLink) {
        productsLink.classList.add("active");
      }
    }
    
    setupMainNav();

    if ($openCartBtn) {
      $openCartBtn.addEventListener("click", () => {
        const modalEl = document.getElementById("quoteModal");
        if (modalEl && typeof bootstrap !== "undefined") {
          const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
          // Inicializar fecha actual al abrir el modal
          if ($quoteDate) {
            const today = new Date();
            $quoteDate.valueAsDate = today;
          }
          modal.show();
        }
      });
    }

    // Inicializar autocompletado de clientes
    setupClientAutocomplete();
    setupContactoAutocomplete();

    if ($btnLoadExcel && $excelInput) {
      $btnLoadExcel.addEventListener("click", () => $excelInput.click());
      $excelInput.addEventListener("change", onExcelSelected);
    }

    // Event listeners para eliminar productos (solo admin)
    if ($btnDeleteProducts && currentUser && currentUser.role === "admin") {
      $btnDeleteProducts.addEventListener("click", openDeleteProductsModal);
      // El CSS se encarga de mostrar/ocultar según la clase admin del body
    }
    if ($btnSelectAllProducts) {
      $btnSelectAllProducts.addEventListener("click", selectAllProductsToDelete);
    }
    if ($btnDeselectAllProducts) {
      $btnDeselectAllProducts.addEventListener("click", deselectAllProductsToDelete);
    }
    if ($btnConfirmDeleteProducts) {
      $btnConfirmDeleteProducts.addEventListener("click", deleteSelectedProducts);
    }
    if ($btnDeleteAllProducts) {
      $btnDeleteAllProducts.addEventListener("click", deleteAllProducts);
    }

    if ($productSearch) {
      $productSearch.addEventListener("input", () => {
        currentSearchQuery = ($productSearch.value || "").trim().toLowerCase();
        renderProductsByLetter(currentLetter);
      });
    }

    // Scroll-to-top
    if ($scrollTopBtn) {
      $scrollTopBtn.addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
      window.addEventListener("scroll", onScrollToggleScrollTop);
      onScrollToggleScrollTop();
    }
  }

  function onExcelSelected(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: "array" });
        productos = buildProductsFromWorkbook(wb);
        saveCatalog(productos);
        selectedProductIds.clear();
        updateSelectionStateUI();
        renderProductsByLetter(currentLetter);
      } catch (err) {
        showAlert("No se pudo leer el Excel. Verifica el formato.", "error");
        console.error(err);
      } finally {
        e.target.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function buildProductsFromWorkbook(wb) {
    const products = [];
    (wb.SheetNames || []).forEach((sheetName) => {
      const sheet = wb.Sheets[sheetName];
      if (!sheet) return;
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
      if (!rows || rows.length === 0) return;

      let i = 0;
      let lastProductName = null;
      let tableIndex = 0;
      while (i < rows.length) {
        const row = ensureArray(rows[i]);
        const maybeName = extractProductNameFromRow(row);
        if (maybeName) {
          lastProductName = maybeName;
          i += 1;
          continue;
        }
        if (!isHeaderRow(row)) { i += 1; continue; }

        const colIndex = mapHeaderIndices(row);
        if (!colIndex) { i += 1; continue; }
        i += 1; // avanzar a la primera fila de datos

        const procesos = [];
        for (; i < rows.length; i++) {
          const r = ensureArray(rows[i]);
          if (extractProductNameFromRow(r) || isHeaderRow(r)) break; // nueva tabla o nuevo producto
          const textRow = r.map((x) => String(x || "").trim());
          const isEmpty = textRow.every((c) => c === "");
          const hasTotal = textRow.some((c) => normalizeText(c) === "total");
          if (isEmpty || hasTotal) { i += 1; break; }

          const get = (idx) => (idx != null && idx >= 0 ? r[idx] : "");
          const proceso = {
            codigo: String(get(colIndex.codigo) || "").trim(),
            analisis: String(get(colIndex.analisis) || "").trim(),
            metodo: String(get(colIndex.metodo) || "").trim(),
            cMtra: String(get(colIndex.cmtra_g) || "").trim(),
            cantidad: String(get(colIndex.cantidad) || "").trim(),
            vrUnit1: String(get(colIndex.vrunit1) || "").trim(),
            vrUnit2: String(get(colIndex.vrunit2) || "").trim(),
            vrUnit3: String(get(colIndex.vrunit3) || "").trim(),
            vrUnitUSD: String(get(colIndex.vrunit) || "").trim(),
            vrTotal: String(get(colIndex.vrtotal) || "").trim()
          };
          const meaningful = proceso.analisis || proceso.metodo || proceso.vrUnitUSD > 0 || proceso.cMtra_g > 0;
          if (meaningful) procesos.push(proceso);
        }

        const nameForThisTable = lastProductName || `${sheetName} ${++tableIndex}`;
        lastProductName = null; // reset para la próxima tabla

        if (procesos.length > 0) {
          products.push({
            id: `${nameForThisTable.charAt(0).toUpperCase()}-${sanitizeFilename(nameForThisTable).toUpperCase()}`,
            nombre: String(nameForThisTable).toUpperCase(),
            procesos
          });
        }
      }
    });
    return products;
  }

  function productoMatchRegex() { return /^producto\s*:/i; }

  function ensureArray(r) { return Array.isArray(r) ? r : []; }

  function isHeaderRow(r) {
    const cells = ensureArray(r).map((c) => slug(String(c || "")));
    const hasCodigo = cells.includes("codigo");
    const hasAnalisis = cells.includes("analisis");
    const hasMetodo = cells.includes("metodo");
    return (hasCodigo && hasAnalisis) || (hasAnalisis && hasMetodo);
  }

  function mapHeaderIndices(r) {
    const cells = ensureArray(r).map((c) => slug(String(c || "")));
    const findIdx = (...keys) => {
      const set = new Set(keys.map((k) => slug(k)));
      let idx = -1;
      cells.forEach((c, i) => { if (set.has(c)) idx = i; });
      return idx >= 0 ? idx : null;
    };
    const map = {
      codigo: findIdx("codigo", "code"),
      analisis: findIdx("analisis", "analysis"),
      metodo: findIdx("metodo", "method"),
      cmtra_g: findIdx("cmtrag", "cmtra", "cmtra g", "cmtra[g]"),
      cantidad: findIdx("cant", "cantidad"),
      vrunit1: findIdx("vrunitario1", "vrunit1"),
      vrunit2: findIdx("vrunitario2", "vrunit2"),
      vrunit3: findIdx("vrunitario3", "vrunit3"),
      vrunit: findIdx("vrunitariousd", "vrunitario", "vrunitariocop", "vrunitariousd"),
      vrtotal: findIdx("vrtotal")
    };
    if (map.analisis == null && map.metodo == null) return null;
    return map;
  }

  function extractProductNameFromRow(r) {
    const cells = ensureArray(r).map((c) => String(c || "").trim());
    // Caso 1: una sola celda tipo "Producto: Nombre"
    for (let c of cells) {
      const m = /^\s*producto\s*:\s*(.+)$/i.exec(c);
      if (m && m[1]) return m[1].trim();
    }
    // Caso 2: celda "Producto" y nombre en la siguiente celda no vacía
    const idx = cells.findIndex((c) => normalizeText(c) === "producto");
    if (idx >= 0) {
      for (let j = idx + 1; j < cells.length; j++) {
        if (String(cells[j]).trim() !== "") return String(cells[j]).trim();
      }
    }
    return null;
  }

  function normalizeText(s) {
    return String(s)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function slug(s) {
    return String(s)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "");
  }

  function toNumber(v, fallback = 0) {
    if (v == null) return fallback;
    if (typeof v === "number" && isFinite(v)) return v;
    const s = String(v).replace(/[^0-9,.-]/g, "").replace(/\.(?=.*\.)/g, "").replace(/,(?=\d{3}\b)/g, "");
    const val = Number(s.replace(",", "."));
    return isNaN(val) ? fallback : val;
  }

  function enforceRoleBasedVisibility() {
    const isAdmin = currentUser && currentUser.role === "admin";
    const viewStats = document.getElementById("viewStats");
    const viewUsers = document.getElementById("viewUsers");
    
    // Controlar visibilidad del botón de eliminar productos
    // El CSS se encarga de mostrar/ocultar el botón según la clase admin del body
    // No necesitamos establecer display manualmente aquí
    
    // Solo para vendedores: asegurar que las secciones admin estén ocultas
    // Esto es una medida de seguridad adicional, pero el sistema de pestañas
    // ya maneja la visibilidad correctamente
    if (!isAdmin) {
      if (viewStats) {
        // Solo ocultar si no es la sección activa
        const activeSection = document.querySelector('section:not(.d-none)');
        if (activeSection && activeSection.id !== 'viewStats') {
          viewStats.classList.add("d-none");
          viewStats.style.display = "none";
          viewStats.style.visibility = "hidden";
        }
      }
      if (viewUsers) {
        // Solo ocultar si no es la sección activa
        const activeSection = document.querySelector('section:not(.d-none)');
        if (activeSection && activeSection.id !== 'viewUsers') {
          viewUsers.classList.add("d-none");
          viewUsers.style.display = "none";
          viewUsers.style.visibility = "hidden";
        }
      }
    }
    // Para admin, el sistema de pestañas ya maneja todo correctamente
  }

  function hideAllSections() {
    const viewProducts = document.getElementById("viewProducts");
    const viewHistory = document.getElementById("viewHistory");
    const viewClients = document.getElementById("viewClients");
    const viewContactos = document.getElementById("viewContactos");
    const viewStats = document.getElementById("viewStats");
    const viewUsers = document.getElementById("viewUsers");
    
    // Ocultar TODAS las secciones sin excepción, usando múltiples métodos para asegurar
    const sections = [viewProducts, viewHistory, viewClients, viewContactos, viewStats, viewUsers];
    sections.forEach(section => {
      if (section) {
        section.classList.add("d-none");
        section.style.display = "none";
        section.style.visibility = "hidden";
      }
    });
  }

  function showSection(sectionId) {
    const section = document.getElementById(sectionId);
    if (!section) return;
    
    const isAdmin = currentUser && currentUser.role === "admin";
    const isAdminSection = sectionId === "viewStats" || sectionId === "viewUsers";
    
    // Si es una sección admin y el usuario no es admin, no mostrar
    if (isAdminSection && !isAdmin) {
      return;
    }
    
    // Mostrar SOLO esta sección, asegurando que esté visible
    section.classList.remove("d-none");
    section.style.display = "block";
    section.style.visibility = "visible";
  }

  function setupMainNav() {
    const nav = document.getElementById("mainNav");
    if (!nav) return;
    
    const isAdmin = currentUser && currentUser.role === "admin";
    
    const links = nav.querySelectorAll(".nav-link");
    links.forEach((link) => {
      link.addEventListener("click", () => {
        links.forEach((l) => l.classList.remove("active"));
        link.classList.add("active");
        const target = link.getAttribute("data-target");
        
        // Ocultar TODAS las secciones primero
        hideAllSections();
        
        // Mostrar solo la sección seleccionada
        if (target === "#viewHistory") {
          showSection("viewHistory");
        } else if (target === "#viewClients") {
          showSection("viewClients");
          renderClients();
        } else if (target === "#viewContactos") {
          showSection("viewContactos");
          renderContactos();
        } else if (target === "#viewStats" && isAdmin) {
          showSection("viewStats");
          renderStats();
        } else if (target === "#viewUsers" && isAdmin) {
          showSection("viewUsers");
          renderUsers();
        } else {
          // Productos por defecto
          showSection("viewProducts");
        }
        
        // Aplicar visibilidad basada en roles como medida de seguridad adicional
        // Solo para asegurar que las secciones admin no sean accesibles por vendedores
        // pero sin interferir con el sistema de pestañas
        if (!isAdmin) {
          enforceRoleBasedVisibility();
        }
      });
    });
  }

  function clearFilters() {
    $filterConsecutivo.value = "";
    $filterClient.value = "";
    $filterProduct.value = "";
    $filterFrom.value = "";
    $filterTo.value = "";
    renderHistory();
  }

  function buildAlphabetNav() {
    // Limpiar el contenedor antes de construir la navegación
    if ($alphabetNav) {
      $alphabetNav.innerHTML = "";
    }
    
    const fragment = document.createDocumentFragment();
    const letters = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));

    // Botón "Todos" primero
    const allBtn = document.createElement("button");
    allBtn.type = "button";
    allBtn.className = `btn btn-sm btn-outline-secondary alphabet-btn`;
    allBtn.textContent = "Todos";
    allBtn.addEventListener("click", () => {
      currentLetter = "*";
      updateAlphabetActive();
      renderProductsByLetter(currentLetter);
    });
    fragment.appendChild(allBtn);

    // Botones de letras A-Z en orden
    letters.forEach((letter) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.dataset.letter = letter;
      btn.className = `btn btn-sm btn-outline-primary alphabet-btn`;
      btn.textContent = letter;
      btn.addEventListener("click", () => {
        currentLetter = letter;
        updateAlphabetActive();
        renderProductsByLetter(letter);
      });
      fragment.appendChild(btn);
    });
    
    // Agregar todos los botones al contenedor
    if ($alphabetNav) {
      $alphabetNav.appendChild(fragment);
    }
    
    updateAlphabetActive();
    // Asegurar inicio al principio en móviles
    try { 
      if ($alphabetNav) {
        $alphabetNav.scrollLeft = 0;
      }
    } catch {}
  }

  function updateAlphabetActive() {
    const buttons = $alphabetNav.querySelectorAll("button");
    buttons.forEach((b) => b.classList.remove("active"));
    const active = Array.from(buttons).find((b) => (currentLetter === "*" ? b.textContent === "Todos" : b.textContent === currentLetter));
    if (active) active.classList.add("active");
  }

  // Función para normalizar texto quitando tildes y acentos
  function normalizeText(text) {
    return text
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase();
  }

  function renderProductsByLetter(letter, resetPage = true) {
    // Resetear a la primera página solo cuando cambia el filtro, no cuando cambia la página
    if (resetPage) {
      currentPage = 1;
    }
    
    $productsContainer.innerHTML = "";
    const filteredList = productos
      .slice()
      .sort((a, b) => a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" }))
      .filter((p) => {
        const matchesSearch = currentSearchQuery ? p.nombre.toLowerCase().includes(currentSearchQuery) : true;
        let matchesLetter = true;
        if (!currentSearchQuery && letter !== "*") {
          // Normalizar el nombre del producto y comparar con la letra normalizada
          const normalizedProductName = normalizeText(p.nombre.trim());
          const normalizedLetter = normalizeText(letter);
          matchesLetter = normalizedProductName.startsWith(normalizedLetter);
        }
        return matchesSearch && matchesLetter;
      });

    if (filteredList.length === 0) {
      $noResults.classList.remove("d-none");
      $productsPagination.classList.add("d-none");
      return;
    }
    $noResults.classList.add("d-none");

    // Calcular paginación
    const totalPages = Math.ceil(filteredList.length / productsPerPage);
    const startIndex = (currentPage - 1) * productsPerPage;
    const endIndex = startIndex + productsPerPage;
    const paginatedList = filteredList.slice(startIndex, endIndex);

    // Renderizar productos de la página actual
    const fragment = document.createDocumentFragment();
    paginatedList.forEach((prod) => {
      fragment.appendChild(renderProductCard(prod));
    });
    $productsContainer.appendChild(fragment);
    updateSelectionStateUI();

    // Renderizar controles de paginación
    renderPagination(totalPages, filteredList.length);
  }

  function renderPagination(totalPages, totalProducts) {
    if (totalPages <= 1) {
      $productsPagination.classList.add("d-none");
      return;
    }

    $productsPagination.classList.remove("d-none");
    const $paginationList = $productsPagination.querySelector("ul.pagination");
    $paginationList.innerHTML = "";
    
    // Limpiar información de paginación anterior si existe
    const existingInfo = $productsPagination.querySelector(".pagination-info");
    if (existingInfo) {
      existingInfo.remove();
    }

    // Botón Anterior
    const prevLi = document.createElement("li");
    prevLi.className = `page-item ${currentPage === 1 ? "disabled" : ""}`;
    prevLi.innerHTML = `<a class="page-link" href="#" data-page="${currentPage - 1}">Anterior</a>`;
    if (currentPage > 1) {
      prevLi.querySelector("a").addEventListener("click", (e) => {
        e.preventDefault();
        goToPage(currentPage - 1);
      });
    }
    $paginationList.appendChild(prevLi);

    // Números de página
    const maxVisiblePages = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
    
    if (endPage - startPage < maxVisiblePages - 1) {
      startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }

    if (startPage > 1) {
      const firstLi = document.createElement("li");
      firstLi.className = "page-item";
      firstLi.innerHTML = `<a class="page-link" href="#" data-page="1">1</a>`;
      firstLi.querySelector("a").addEventListener("click", (e) => {
        e.preventDefault();
        goToPage(1);
      });
      $paginationList.appendChild(firstLi);

      if (startPage > 2) {
        const ellipsisLi = document.createElement("li");
        ellipsisLi.className = "page-item disabled";
        ellipsisLi.innerHTML = `<span class="page-link">...</span>`;
        $paginationList.appendChild(ellipsisLi);
      }
    }

    for (let i = startPage; i <= endPage; i++) {
      const pageLi = document.createElement("li");
      pageLi.className = `page-item ${i === currentPage ? "active" : ""}`;
      pageLi.innerHTML = `<a class="page-link" href="#" data-page="${i}">${i}</a>`;
      if (i !== currentPage) {
        pageLi.querySelector("a").addEventListener("click", (e) => {
          e.preventDefault();
          goToPage(i);
        });
      }
      $paginationList.appendChild(pageLi);
    }

    if (endPage < totalPages) {
      if (endPage < totalPages - 1) {
        const ellipsisLi = document.createElement("li");
        ellipsisLi.className = "page-item disabled";
        ellipsisLi.innerHTML = `<span class="page-link">...</span>`;
        $paginationList.appendChild(ellipsisLi);
      }

      const lastLi = document.createElement("li");
      lastLi.className = "page-item";
      lastLi.innerHTML = `<a class="page-link" href="#" data-page="${totalPages}">${totalPages}</a>`;
      lastLi.querySelector("a").addEventListener("click", (e) => {
        e.preventDefault();
        goToPage(totalPages);
      });
      $paginationList.appendChild(lastLi);
    }

    // Botón Siguiente
    const nextLi = document.createElement("li");
    nextLi.className = `page-item ${currentPage === totalPages ? "disabled" : ""}`;
    nextLi.innerHTML = `<a class="page-link" href="#" data-page="${currentPage + 1}">Siguiente</a>`;
    if (currentPage < totalPages) {
      nextLi.querySelector("a").addEventListener("click", (e) => {
        e.preventDefault();
        goToPage(currentPage + 1);
      });
    }
    $paginationList.appendChild(nextLi);

    // Información de paginación
    const infoDiv = document.createElement("div");
    infoDiv.className = "text-center text-muted small mt-2 pagination-info";
    const startProduct = (currentPage - 1) * productsPerPage + 1;
    const endProduct = Math.min(currentPage * productsPerPage, totalProducts);
    infoDiv.textContent = `Mostrando ${startProduct}-${endProduct} de ${totalProducts} productos`;
    $productsPagination.appendChild(infoDiv);
  }

  function goToPage(page) {
    currentPage = page;
    renderProductsByLetter(currentLetter, false); // No resetear la página, solo cambiar de página
    // Scroll al inicio de los productos
    $productsContainer.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderProductCard(product) {
    const col = document.createElement("div");
    col.className = "col-12"; // ancho completo por fila

    const card = document.createElement("div");
    card.className = "card product-card h-100";

    // Obtener el valor unitario seleccionado o usar 'vrUnitUSD' por defecto
    const selectedUnit = selectedUnitValues.get(product.id) || 'vrUnitUSD';
    
    const header = document.createElement("div");
    header.className = "card-header d-flex justify-content-between align-items-start gap-2 flex-wrap";
    header.innerHTML = `
      <div class="product-title">
        <input class="form-check-input me-2 product-select" type="checkbox" data-product-id="${product.id}">
        <span class="fw-semibold">${product.nombre}</span>
      </div>
      <div class="d-flex align-items-center gap-2">
        <label class="small text-muted mb-0">Valor unitario:</label>
        <select class="form-select form-select-sm unit-selector" data-product-id="${product.id}" style="width: auto; min-width: 140px;">
          <option value="vrUnit1" ${selectedUnit === 'vrUnit1' ? 'selected' : ''}>Vr. Unitario 1</option>
          <option value="vrUnit2" ${selectedUnit === 'vrUnit2' ? 'selected' : ''}>Vr. Unitario 2</option>
          <option value="vrUnit3" ${selectedUnit === 'vrUnit3' ? 'selected' : ''}>Vr. Unitario 3</option>
          <option value="vrUnitUSD" ${selectedUnit === 'vrUnitUSD' ? 'selected' : ''}>Vr. Unitario USD</option>
        </select>
      </div>
    `;

    const body = document.createElement("div");
    body.className = "card-body p-0";
    body.appendChild(renderProductTable(product));

    card.appendChild(header);
    card.appendChild(body);
    col.appendChild(card);

    const checkbox = header.querySelector(".product-select");
    checkbox.checked = selectedProductIds.has(product.id);
    checkbox.addEventListener("change", (e) => {
      if (e.target.checked) {
        selectedProductIds.add(product.id);
      } else {
        selectedProductIds.delete(product.id);
      }
      updateSelectionStateUI();
    });

    // Event listener para el selector de valor unitario
    const unitSelector = header.querySelector(".unit-selector");
    if (unitSelector) {
      unitSelector.addEventListener("change", (e) => {
        const productId = e.target.getAttribute("data-product-id");
        const selectedValue = e.target.value;
        selectedUnitValues.set(productId, selectedValue);
        
        // Recalcular totales del producto y actualizar la tabla
        const productData = productos.find(p => p.id === productId);
        if (productData) {
          // Actualizar la columna Vr. Total en cada fila y el subtotal
          updateProductTableTotals(productId, productData);
        }
      });
    }

    return col;
  }

  function updateProductTableTotals(productId, product) {
    // Recalcular el subtotal usando el valor unitario seleccionado
    const selectedUnit = selectedUnitValues.get(productId) || 'vrUnitUSD';
    let sumTotal = 0;
    
    // Actualizar la columna Vr. Total en cada fila de la tabla
    const unitSelector = document.querySelector(`.unit-selector[data-product-id="${productId}"]`);
    if (unitSelector) {
      const card = unitSelector.closest('.product-card');
      if (card) {
        const tbody = card.querySelector('tbody');
        if (tbody) {
          const rows = tbody.querySelectorAll('tr');
          rows.forEach((tr, index) => {
            if (product.procesos[index]) {
              const row = product.procesos[index];
              const qty = parseNumStrict(row.cantidad);
              const unitValue = parseNumStrict(row[selectedUnit]);
              let calculatedTotal = 0;
              
              // Solo calcular si el valor unitario seleccionado tiene un valor válido
              if (unitValue > 0 && qty > 0) {
                calculatedTotal = unitValue * qty;
                sumTotal += calculatedTotal;
              }
              
              // Actualizar la celda de Vr. Total (última columna)
              const totalCell = tr.querySelector('td:last-child');
              if (totalCell) {
                totalCell.textContent = calculatedTotal > 0 ? formatMoney(calculatedTotal) : "";
              }
            }
          });
        }
        
        // Actualizar el subtotal en el footer de la tabla
        const tfoot = card.querySelector('tfoot');
        if (tfoot) {
          const subtotalCell = tfoot.querySelector('td:last-child');
          if (subtotalCell) {
            subtotalCell.textContent = formatMoney(sumTotal);
          }
        }
      }
    }
  }

  function renderProductTable(product) {
    const tableWrapper = document.createElement("div");
    tableWrapper.className = "table-responsive";
    const table = document.createElement("table");
    table.className = "table table-sm table-striped table-hover mb-0";

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th>Código</th>
        <th>Análisis</th>
        <th>Método</th>
        <th>C. Mtra. [g]</th>
        <th>Cant.</th>
        <th>Vr. Unitario 1</th>
        <th>Vr. Unitario 2</th>
        <th>Vr. Unitario 3</th>
        <th>Vr. Unitario USD</th>
        <th>Vr. Total</th>
      </tr>
    `;

    const tbody = document.createElement("tbody");
    let sumCMtra = 0;
    let sumTotal = 0;
    
    // Obtener el valor unitario seleccionado para este producto
    const selectedUnit = selectedUnitValues.get(product.id) || 'vrUnitUSD';
    
    product.procesos.forEach((row) => {
      const getText = (v) => (v == null ? "" : String(v));
      const valueOrFormat = (text, num, isMoney = false) => {
        if (text !== undefined && text !== null && String(text) !== "") return String(text);
        if (num === undefined || num === null || String(num) === "") return "";
        return isMoney ? formatMoney(num) : formatNumber(num);
      };

      // Sumas usando el valor unitario seleccionado
      const cMtraForSum = parseNumStrict(row.cMtra ?? row.cMtra_g);
      const qty = parseNumStrict(row.cantidad);
      const unitValue = parseNumStrict(row[selectedUnit]);
      
      sumCMtra += cMtraForSum;
      
      // Calcular el total usando el valor unitario seleccionado
      // Solo sumar si el valor unitario tiene un valor válido
      if (unitValue > 0 && qty > 0) {
        sumTotal += unitValue * qty;
      }
      // Si el valor unitario está vacío o es 0, no sumar nada

      // Calcular el Vr. Total usando el valor unitario seleccionado
      const qtyForDisplay = parseNumStrict(row.cantidad);
      const unitValueForDisplay = parseNumStrict(row[selectedUnit]);
      let calculatedTotal = 0;
      
      // Solo calcular si el valor unitario seleccionado tiene un valor válido
      if (unitValueForDisplay > 0 && qtyForDisplay > 0) {
        calculatedTotal = unitValueForDisplay * qtyForDisplay;
      }
      
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${getText(row.codigo)}</td>
        <td>${getText(row.analisis)}</td>
        <td>${getText(row.metodo)}</td>
        <td class=\"text-center\">${valueOrFormat(row.cMtra, row.cMtra_g)}</td>
        <td class=\"text-center\">${valueOrFormat(row.cantidad, row.cantidad)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnit1)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnit2)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnit3)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnitUSD)}</td>
        <td class=\"text-center\">${calculatedTotal > 0 ? formatMoney(calculatedTotal) : ""}</td>
      `;
      tbody.appendChild(tr);
    });

    const tfoot = document.createElement("tfoot");
    tfoot.innerHTML = `
      <tr>
        <td colspan="3" class="text-end">Totales:</td>
        <td class="text-center">${formatNumber(sumCMtra)}</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="text-center fw-semibold">Subtotal</td>
        <td class="text-center fw-semibold">${formatMoney(sumTotal)}</td>
      </tr>
    `;

    table.appendChild(thead);
    table.appendChild(tbody);
    table.appendChild(tfoot);
    tableWrapper.appendChild(table);

    // Guardar totales calculados en el objeto para uso en PDF
    product._sumCMtra = sumCMtra;
    product._sumTotal = sumTotal;
    return tableWrapper;
  }

  function updateSelectionStateUI() {
    const count = selectedProductIds.size;
    $selectedCount.textContent = `${count} seleccionados`;
    $btnGeneratePDF.disabled = count === 0;
    if ($cartCount) $cartCount.textContent = `${count}`;
    // Sincronizar checkboxes visibles
    document.querySelectorAll(".product-select").forEach((cb) => {
      const id = cb.getAttribute("data-product-id");
      cb.checked = selectedProductIds.has(id);
    });
  }

  async function onGeneratePDF() {
    const clientName = $clientName.value.trim();
    const clientEmail = $clientEmail.value.trim();
    const clientNit = $quoteClientNit ? $quoteClientNit.value.trim() : "";
    const clientCelular = $quoteClientCelular ? $quoteClientCelular.value.trim() : "";
    const clientFormaPago = $quoteClientFormaPago ? $quoteClientFormaPago.value.trim() : "";
    const duracionAnalisis = $quoteDuracionAnalisis ? $quoteDuracionAnalisis.value.trim() : "";
    const nota = $quoteNota ? $quoteNota.value.trim() : "";
    const clientId = $quoteClientId ? $quoteClientId.value.trim() : "";
    
    // Obtener contactos seleccionados
    let clientContactos = [];
    let contactoEmail = ""; // Correo del contacto seleccionado
    if ($quoteClientContactosSelected) {
      const value = $quoteClientContactosSelected.value.trim();
      if (value) {
        try {
          const contactosArray = JSON.parse(value);
          clientContactos = contactosArray.map(c => c.nombre);
          // Usar el correo del primer contacto seleccionado, o del que tenga correo
          const contactoConCorreo = contactosArray.find(c => c.correo && c.correo.trim());
          if (contactoConCorreo) {
            contactoEmail = contactoConCorreo.correo;
          } else if (contactosArray.length > 0 && contactosArray[0].correo) {
            contactoEmail = contactosArray[0].correo;
          }
        } catch {
          // Fallback para formato antiguo
          clientContactos = value.split(",");
        }
      }
    }
    const dateStr = $quoteDate.value ? $quoteDate.value : new Date().toISOString().slice(0, 10);

    if (!clientName) {
      showAlert("Por favor ingresa el nombre del cliente.", "warning");
      return;
    }

    const selectedProducts = productos
      .filter((p) => selectedProductIds.has(p.id))
      .map(recomputeTotalsForProduct);

    if (selectedProducts.length === 0) {
      showAlert("Selecciona al menos un producto.", "warning");
      return;
    }

    // Totales
    const totalGeneral = selectedProducts.reduce((acc, p) => acc + (p._sumTotal || 0), 0);

    // Cargar logo (cache) y generar PDF
    if (!logoAsset) {
      try { logoAsset = await loadLogo(); } catch {}
    }
    // Generar número de consecutivo (solo una vez)
    const quoteNumber = generateQuoteNumber();
    
    await generatePDF({ 
        clientName, 
        clientEmail: contactoEmail || clientEmail, // Usar correo del contacto si está disponible
      clientNit, 
      clientContactos: clientContactos,
      clientCelular,
      clientFormaPago,
      duracionAnalisis,
      nota,
      dateStr, 
      quoteNumber,
      products: selectedProducts, 
      totalGeneral, 
      logo: logoAsset 
    });

    // Guardar historial con el número de consecutivo
    const quote = {
      id: `Q-${Date.now()}`,
      quoteNumber,
      date: dateStr,
      clientName,
      clientEmail,
      clientNit,
      clientContactos: clientContactos,
      clientCelular,
      clientFormaPago,
      duracionAnalisis,
      nota,
      clientId,
      products: selectedProducts.map((p) => ({ id: p.id, nombre: p.nombre, subtotal: p._sumTotal })),
      totalCOP: totalGeneral
    };
    saveQuote(quote);
    renderHistory();
    renderStats();

    // Cerrar modal y resetear formulario
    closeQuoteModal();
    resetClientForm();
    clearCartSelection();
  }

  function recomputeTotalsForProduct(product) {
    let sumCMtra = 0;
    let sumTotal = 0;
    
    // Obtener el valor unitario seleccionado para este producto
    const selectedUnit = selectedUnitValues.get(product.id) || 'vrUnitUSD';
    
    product.procesos.forEach((row) => {
      const cMtraNum = parseNumStrict(row.cMtra_g ?? row.cMtra);
      const qty = parseNumStrict(row.cantidad);
      // Usar el valor unitario seleccionado
      const unit = parseNumStrict(row[selectedUnit]);
      if (isFinite(cMtraNum)) sumCMtra += cMtraNum;
      
      // Solo calcular si el valor unitario seleccionado tiene un valor válido
      if (unit > 0 && qty > 0) {
        sumTotal += unit * qty;
      }
      // Si el valor unitario está vacío o es 0, no sumar nada
    });
    return { ...product, _sumCMtra: sumCMtra, _sumTotal: sumTotal, _selectedUnit: selectedUnit };
  }

  function closeQuoteModal() {
    const modalEl = document.getElementById("quoteModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.hide();
    }
  }

  function resetClientForm() {
    const form = document.getElementById("clientForm");
    if (form) form.reset();
    if ($quoteDate) {
      $quoteDate.value = "";
      $quoteDate.valueAsDate = new Date();
    }
    // Limpiar datos del cliente seleccionado
    if ($quoteClientId) $quoteClientId.value = "";
    if ($quoteClientNit) $quoteClientNit.value = "";
    if ($quoteClientContacto) $quoteClientContacto.value = "";
    if ($selectedContactos) $selectedContactos.innerHTML = "";
    if ($quoteClientContactosSelected) $quoteClientContactosSelected.value = "";
    if ($contactoDropdown) $contactoDropdown.style.display = "none";
    availableContactos = [];
    if ($quoteClientCelular) $quoteClientCelular.value = "";
    if ($quoteClientFormaPago) $quoteClientFormaPago.value = "";
    if ($quoteDuracionAnalisis) $quoteDuracionAnalisis.value = "";
    if ($quoteNota) $quoteNota.value = "";
    if ($clientDropdown) $clientDropdown.style.display = "none";
  }

  // Funciones para el autocompletado de clientes en el modal de cotización
  function searchClients(query) {
    if (!query || query.trim().length < 1) {
      return [];
    }
    const clients = getClients();
    const searchTerm = query.toLowerCase().trim();
    return clients.filter(client => {
      const nombre = (client.nombre || "").toLowerCase();
      const nit = (client.nit || "").toLowerCase();
      const contactos = Array.isArray(client.contactos) 
        ? client.contactos.join(" ").toLowerCase()
        : (client.contacto || "").toLowerCase();
      const correo = (client.correo || "").toLowerCase();
      return nombre.includes(searchTerm) || 
             nit.includes(searchTerm) || 
             contactos.includes(searchTerm) || 
             correo.includes(searchTerm);
    });
  }

  function renderClientDropdown(clients) {
    if (!$clientDropdown) return;
    
    $clientDropdown.innerHTML = "";
    
    if (clients.length === 0) {
      $clientDropdown.style.display = "none";
      return;
    }

    clients.slice(0, 10).forEach(client => {
      const item = document.createElement("a");
      item.className = "dropdown-item";
      item.href = "#";
      item.style.cursor = "pointer";
      item.innerHTML = `
        <div class="fw-semibold">${escapeHtml(client.nombre || "")}</div>
        <small class="text-muted">${escapeHtml(client.nit || "Sin NIT")} ${client.correo ? "• " + escapeHtml(client.correo) : ""}</small>
      `;
      item.addEventListener("click", (e) => {
        e.preventDefault();
        selectClient(client);
      });
      $clientDropdown.appendChild(item);
    });

    $clientDropdown.style.display = "block";
  }

  function selectClient(client) {
    if (!$clientName || !$quoteClientId) return;
    
    // Llenar los campos con los datos del cliente
    $clientName.value = client.nombre || "";
    $quoteClientId.value = client.id || "";
    $clientEmail.value = client.correo || "";
    $quoteClientNit.value = client.nit || "";
    if ($quoteClientCelular) $quoteClientCelular.value = client.celular || "";
    if ($quoteClientFormaPago) $quoteClientFormaPago.value = client.formaPago || "";
    
    // Buscar contactos del cliente en la pestaña Contactos
    const allContactos = getContactos();
    const clientNameNormalized = normalizeText(client.nombre || "");
    const clientNitNormalized = client.nit ? normalizeText(client.nit) : null;
    
    // Filtrar contactos que pertenezcan a este cliente (por nombre o NIT)
    availableContactos = allContactos
      .filter(contacto => {
        const contactoClienteNormalized = normalizeText(contacto.nombreCliente || "");
        return contactoClienteNormalized === clientNameNormalized ||
               (clientNitNormalized && contactoClienteNormalized === clientNitNormalized);
      })
      .map(contacto => ({
        nombre: contacto.nombre || '',
        correo: contacto.correo || ''
      }));
    
    // Limpiar el campo de contacto y los seleccionados
    if ($quoteClientContacto) {
      $quoteClientContacto.value = "";
    }
    if ($selectedContactos) {
      $selectedContactos.innerHTML = "";
    }
    if ($quoteClientContactosSelected) {
      $quoteClientContactosSelected.value = "";
    }
    if ($contactoDropdown) {
      $contactoDropdown.style.display = "none";
    }
    
    // Ocultar el dropdown
    if ($clientDropdown) {
      $clientDropdown.style.display = "none";
    }
    
    // Validar el formulario
    if ($clientName.form) {
      $clientName.form.reportValidity();
    }
  }

  function setupClientAutocomplete() {
    if (!$clientName || !$clientDropdown) return;

    let searchTimeout = null;

    // Event listener para cuando se escribe en el campo
    $clientName.addEventListener("input", (e) => {
      const query = e.target.value.trim();
      
      // Limpiar el timeout anterior
      if (searchTimeout) {
        clearTimeout(searchTimeout);
      }

      // Si se borra el texto, limpiar los campos
      if (query.length === 0) {
        if ($quoteClientId) $quoteClientId.value = "";
        if ($clientEmail) $clientEmail.value = "";
        if ($quoteClientNit) $quoteClientNit.value = "";
        if ($quoteClientContacto) $quoteClientContacto.value = "";
        if ($selectedContactos) $selectedContactos.innerHTML = "";
        if ($quoteClientContactosSelected) $quoteClientContactosSelected.value = "";
        if ($contactoDropdown) $contactoDropdown.style.display = "none";
        availableContactos = [];
        if ($clientDropdown) $clientDropdown.style.display = "none";
        return;
      }

      // Si hay un cliente seleccionado, verificar si el texto coincide
      // Si no coincide, limpiar el ID y los campos relacionados
      if ($quoteClientId && $quoteClientId.value.trim()) {
        const clients = getClients();
        const selectedClient = clients.find(c => c.id === $quoteClientId.value.trim());
        // Si el texto escrito no coincide con el cliente seleccionado, limpiar
        if (!selectedClient || selectedClient.nombre !== query) {
          $quoteClientId.value = "";
          if ($quoteClientNit) $quoteClientNit.value = "";
          if ($quoteClientContacto) $quoteClientContacto.value = "";
          if ($selectedContactos) $selectedContactos.innerHTML = "";
          if ($quoteClientContactosSelected) $quoteClientContactosSelected.value = "";
          if ($contactoDropdown) $contactoDropdown.style.display = "none";
          availableContactos = [];
          // Mantener el email solo si el usuario no lo cambió manualmente
          // (esto es opcional, puedes limpiarlo también si prefieres)
        }
      } else {
        // Si no hay cliente seleccionado, asegurar que los campos estén limpios
        if ($quoteClientNit) $quoteClientNit.value = "";
        if ($quoteClientContacto) $quoteClientContacto.value = "";
        if ($selectedContactos) $selectedContactos.innerHTML = "";
        if ($quoteClientContactosSelected) $quoteClientContactosSelected.value = "";
        if ($contactoDropdown) $contactoDropdown.style.display = "none";
        availableContactos = [];
      }

      // Buscar clientes con un pequeño delay para evitar búsquedas excesivas
      searchTimeout = setTimeout(() => {
        const results = searchClients(query);
        renderClientDropdown(results);
        
        // Si hay exactamente un resultado que coincide exactamente con el texto escrito,
        // podríamos auto-seleccionarlo, pero es mejor dejar que el usuario lo seleccione
      }, 150);
    });

    // Cerrar el dropdown cuando se hace clic fuera
    document.addEventListener("click", (e) => {
      if ($clientDropdown && $clientName) {
        const isClickInside = $clientName.contains(e.target) || $clientDropdown.contains(e.target);
        if (!isClickInside) {
          $clientDropdown.style.display = "none";
        }
      }
    });

    // Cerrar el dropdown cuando se presiona Escape
    $clientName.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        if ($clientDropdown) {
          $clientDropdown.style.display = "none";
        }
      }
    });
  }

  // Funciones para el autocompletado de contactos
  function getSelectedContactos() {
    if (!$quoteClientContactosSelected) return [];
    const value = $quoteClientContactosSelected.value.trim();
    if (!value) return [];
    try {
      return JSON.parse(value);
    } catch {
      // Fallback para formato antiguo (strings separados por comas)
      return value.split(",").map(c => ({ nombre: c.trim(), correo: '' }));
    }
  }

  function addContacto(contactoObj) {
    if (!contactoObj) return;
    
    // Asegurar que contactoObj es un objeto
    const contacto = typeof contactoObj === 'object' ? contactoObj : { nombre: String(contactoObj || ''), correo: '' };
    if (!contacto.nombre || !contacto.nombre.trim()) return;
    
    const selected = getSelectedContactos();
    // Verificar si ya está seleccionado (comparar por nombre)
    const alreadySelected = selected.some(c => normalizeText(c.nombre) === normalizeText(contacto.nombre));
    if (alreadySelected) {
      return; // Ya está seleccionado
    }
    
    selected.push({ nombre: contacto.nombre.trim(), correo: contacto.correo ? contacto.correo.trim() : '' });
    updateSelectedContactos(selected);
    
    // Actualizar el campo Email con el correo del contacto seleccionado
    if (contacto.correo && contacto.correo.trim() && $clientEmail) {
      $clientEmail.value = contacto.correo.trim();
    }
    
    // Limpiar el input
    if ($quoteClientContacto) {
      $quoteClientContacto.value = "";
    }
    if ($contactoDropdown) {
      $contactoDropdown.style.display = "none";
    }
  }

  function removeContacto(contactoNombre) {
    const selected = getSelectedContactos();
    const index = selected.findIndex(c => normalizeText(c.nombre) === normalizeText(contactoNombre));
    if (index > -1) {
      selected.splice(index, 1);
      updateSelectedContactos(selected);
    }
  }

  function updateSelectedContactos(contactos) {
    if (!$quoteClientContactosSelected || !$selectedContactos) return;
    
    // Guardar como JSON string
    $quoteClientContactosSelected.value = JSON.stringify(contactos);
    
    // Renderizar los tags de contactos seleccionados
    $selectedContactos.innerHTML = "";
    contactos.forEach((contacto) => {
      const tag = document.createElement("span");
      tag.className = "badge bg-primary d-inline-flex align-items-center gap-1";
      tag.innerHTML = `
        ${contacto.nombre}${contacto.correo ? ` (${contacto.correo})` : ''}
        <button type="button" class="btn-close btn-close-white" style="font-size: 0.65em;" aria-label="Eliminar"></button>
      `;
      tag.querySelector("button").addEventListener("click", () => {
        removeContacto(contacto.nombre);
      });
      $selectedContactos.appendChild(tag);
    });
  }

  function filterContactos(query) {
    const selected = getSelectedContactos();
    const selectedNombres = selected.map(c => normalizeText(c.nombre));
    
    if (!query || !query.trim()) {
      return availableContactos.filter(contacto => {
        return !selectedNombres.includes(normalizeText(contacto.nombre));
      });
    }
    
    const queryLower = query.toLowerCase().trim();
    return availableContactos.filter(contacto => {
      const nombreMatch = contacto.nombre.toLowerCase().includes(queryLower);
      const correoMatch = contacto.correo && contacto.correo.toLowerCase().includes(queryLower);
      return !selectedNombres.includes(normalizeText(contacto.nombre)) && (nombreMatch || correoMatch);
    });
  }

  function renderContactoDropdown(contactos) {
    if (!$contactoDropdown) return;
    
    $contactoDropdown.innerHTML = "";
    
    if (contactos.length === 0) {
      const item = document.createElement("div");
      item.className = "dropdown-item-text text-muted";
      item.textContent = "No hay contactos disponibles";
      $contactoDropdown.appendChild(item);
      $contactoDropdown.style.display = "block";
      return;
    }
    
    contactos.forEach((contacto) => {
      const item = document.createElement("a");
      item.className = "dropdown-item";
      item.href = "#";
      item.innerHTML = `
        <div>
          <strong>${contacto.nombre}</strong>
          ${contacto.correo ? `<br><small class="text-muted">${contacto.correo}</small>` : ''}
        </div>
      `;
      item.addEventListener("click", (e) => {
        e.preventDefault();
        addContacto(contacto);
      });
      $contactoDropdown.appendChild(item);
    });
    
    $contactoDropdown.style.display = "block";
  }

  function setupContactoAutocomplete() {
    if (!$quoteClientContacto || !$contactoDropdown) return;

    let searchTimeout = null;

    // Event listener para cuando se hace clic o focus en el campo
    $quoteClientContacto.addEventListener("focus", (e) => {
      // Si no hay contactos disponibles o no hay cliente seleccionado, no mostrar dropdown
      if (availableContactos.length === 0 || !$quoteClientId || !$quoteClientId.value.trim()) {
        if ($contactoDropdown) {
          $contactoDropdown.style.display = "none";
        }
        return;
      }

      const query = e.target.value.trim();
      const results = filterContactos(query);
      renderContactoDropdown(results);
    });

    // Event listener para cuando se escribe en el campo
    $quoteClientContacto.addEventListener("input", (e) => {
      const query = e.target.value.trim();
      
      // Limpiar el timeout anterior
      if (searchTimeout) {
        clearTimeout(searchTimeout);
      }

      // Si no hay contactos disponibles o no hay cliente seleccionado, no mostrar dropdown
      if (availableContactos.length === 0 || !$quoteClientId || !$quoteClientId.value.trim()) {
        if ($contactoDropdown) {
          $contactoDropdown.style.display = "none";
        }
        return;
      }

      // Buscar contactos con un pequeño delay
      searchTimeout = setTimeout(() => {
        const results = filterContactos(query);
        renderContactoDropdown(results);
      }, 150);
    });

    // Cerrar el dropdown cuando se hace clic fuera
    document.addEventListener("click", (e) => {
      if ($contactoDropdown && $quoteClientContacto) {
        const isClickInside = $quoteClientContacto.contains(e.target) || $contactoDropdown.contains(e.target) || $selectedContactos.contains(e.target);
        if (!isClickInside) {
          $contactoDropdown.style.display = "none";
        }
      }
    });

    // Cerrar el dropdown cuando se presiona Escape
    $quoteClientContacto.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        if ($contactoDropdown) {
          $contactoDropdown.style.display = "none";
        }
      } else if (e.key === "Enter") {
        e.preventDefault();
        // Seleccionar el primer resultado si hay
        const firstItem = $contactoDropdown.querySelector(".dropdown-item");
        if (firstItem) {
          firstItem.click();
        }
      }
    });
  }

  function clearCartSelection() {
    selectedProductIds.clear();
    // No limpiar selectedUnitValues para mantener las selecciones del usuario
    // selectedUnitValues.clear();
    updateSelectionStateUI();
  }

  // Función para convertir imagen a dataURL (más confiable para PDFs)
  function imageToDataUrl(img, format = 'image/png') {
    try {
      if (!img) return null;
      const canvas = document.createElement('canvas');
      const width = img.naturalWidth || img.width || img.clientWidth;
      const height = img.naturalHeight || img.height || img.clientHeight;
      if (width <= 0 || height <= 0) {
        console.error('Dimensiones inválidas para la imagen:', width, height);
        return null;
      }
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);
      return canvas.toDataURL(format);
    } catch (e) {
      console.error('Error al convertir imagen a dataURL:', e);
      return null;
    }
  }

  // Función para cargar imagen como dataURL (compatible con file:// y http://)
  async function loadImageAsDataUrl(src) {
    // Primero intentar con loadImageElement (funciona mejor con file://)
    try {
      const img = await loadImageElement(src);
      if (img) {
        const dataUrl = imageToDataUrl(img, 'image/png');
        if (dataUrl) {
          return dataUrl;
        }
      }
    } catch (e) {
      // Si falla, intentar con fetch (solo funciona con http:// o https://)
      console.warn(`loadImageElement falló para ${src}, intentando con fetch...`);
    }
    
    // Si loadImageElement falla, intentar con fetch (solo funciona con servidor HTTP)
    const isFileProtocol = window.location.protocol === 'file:';
    const basePath = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1);
    
    const candidates = [
      src,
      `./${src}`,
      `${basePath}${src}`,
      ...(isFileProtocol ? [] : [`/${src}`]) // Solo usar rutas absolutas si no es file://
    ];
    
    for (const candidate of candidates) {
      try {
        const response = await fetch(candidate);
        if (response.ok) {
          const blob = await response.blob();
          return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
          });
        }
      } catch (e) {
        // Continuar con el siguiente candidato
        continue;
      }
    }
    
    console.error(`No se pudo cargar ${src} desde ninguna ruta`);
    return null;
  }

  // Función helper para justificar texto en jsPDF
  function justifyText(doc, text, x, y, maxWidth, fontSize = 9, pageHeight = null, bottomMargin = 40) {
    if (!text || text.trim().length === 0) {
      return y;
    }
    
    const lines = doc.splitTextToSize(text, maxWidth);
    const lineHeight = fontSize * 1.2;
    let currentY = y;
    
    for (let index = 0; index < lines.length; index++) {
      const line = lines[index];
      const trimmedLine = line.trim();
      
      // Verificar si necesitamos una nueva página antes de agregar esta línea
      if (pageHeight && currentY + lineHeight > pageHeight - bottomMargin) {
        doc.addPage();
        currentY = 50; // Posición inicial en la nueva página con margen superior
      }
      
      // Última línea o línea vacía: alinear a la izquierda (no justificar)
      if (index === lines.length - 1 || trimmedLine.length === 0) {
        doc.text(trimmedLine || line, x, currentY);
        currentY += lineHeight;
        continue;
      }
      
      // Calcular espacios necesarios para justificar
      const words = trimmedLine.split(/\s+/).filter(w => w.length > 0);
      
      if (words.length <= 1) {
        // Una sola palabra: alinear a la izquierda
        doc.text(trimmedLine, x, currentY);
        currentY += lineHeight;
        continue;
      }
      
      // Calcular ancho del texto sin espacios extra
      let textWidth = 0;
      for (let i = 0; i < words.length; i++) {
        textWidth += doc.getTextWidth(words[i]);
        if (i < words.length - 1) {
          textWidth += doc.getTextWidth(' ');
        }
      }
      
      const totalSpaces = words.length - 1;
      const extraSpace = totalSpaces > 0 ? (maxWidth - textWidth) / totalSpaces : 0;
      
      // Si el espacio extra es muy grande o negativo, simplemente alinear a la izquierda
      if (extraSpace < 0 || extraSpace > maxWidth / 2) {
        doc.text(trimmedLine, x, currentY);
        currentY += lineHeight;
        continue;
      }
      
      // Justificar: distribuir el espacio extra entre palabras
      let currentX = x;
      for (let i = 0; i < words.length; i++) {
        doc.text(words[i], currentX, currentY);
        currentX += doc.getTextWidth(words[i]);
        if (i < words.length - 1) {
          currentX += doc.getTextWidth(' ') + extraSpace;
        }
      }
      
      currentY += lineHeight;
    }
    
    return currentY;
  }

  async function generatePDF({ clientName, clientEmail, clientNit, clientContactos = [], clientCelular = "", clientFormaPago = "", duracionAnalisis = "", nota = "", dateStr, quoteNumber, products, totalGeneral, logo }) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    // Colores de marca (alineados con la web)
    const brandPrimary = [68, 194, 196]; // #44c2c4
    const brandSecondary = [243, 192, 44]; // #f3c02c
    const dark = [33, 37, 41];

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const left = 40;
    const right = 555; // ancho útil para alinear a la derecha
    const bottomMargin = 40; // margen inferior
    const topMargin = 50; // margen superior para nuevas páginas
    
    // Función helper para verificar y crear nueva página si es necesario
    function checkPageBreak(y, requiredSpace = 20) {
      if (y + requiredSpace > pageHeight - bottomMargin) {
        doc.addPage();
        return topMargin; // Retornar posición inicial de la nueva página con margen superior
      }
      return y;
    }

    // Logo en la parte superior izquierda (logoolg.png)
    const logoW = 220; // ancho fijo en px
    const logoH = 80;  // alto fijo en px
    const xLogo = left; // Posición izquierda
    const yLogo = 20; // Posición superior
    try {
      const logoolgDataUrl = await loadImageAsDataUrl("assets/images/logoolg.png");
      if (logoolgDataUrl) {
        doc.addImage(logoolgDataUrl, "PNG", xLogo, yLogo, logoW, logoH);
      } else {
        console.error("No se pudo cargar logoolg.png");
      }
    } catch (e) {
      console.error("Error al cargar logoolg.png:", e);
    }
    
    // Cargar y mostrar C1.png y C2.png en la parte superior derecha
    const certW = 70; // ancho para cada certificado
    const certH = 70; // alto para cada certificado (circular)
    const spacing = 8; // espacio entre certificados
    const marginRight = left; // Margen derecho igual al izquierdo
    const totalCertWidth = (certW * 2) + spacing; // Ancho total de ambos certificados
    const xC1 = pageWidth - totalCertWidth - marginRight; // Posición de C1 (desde la derecha)
    const xC2 = xC1 + certW + spacing; // Posición de C2 (al lado de C1)
    const yCert = yLogo + (logoH - certH) / 2; // Centrado verticalmente con el logo
    
    // Cargar C1.png
    try {
      const c1DataUrl = await loadImageAsDataUrl("assets/images/C1.png");
      if (c1DataUrl) {
        doc.addImage(c1DataUrl, "PNG", xC1, yCert, certW, certH);
      } else {
        console.error("No se pudo cargar C1.png");
      }
    } catch (e) {
      console.error("Error al cargar C1.png:", e);
    }
    
    // Cargar C2.png
    try {
      const c2DataUrl = await loadImageAsDataUrl("assets/images/C2.png");
      if (c2DataUrl) {
        doc.addImage(c2DataUrl, "PNG", xC2, yCert, certW, certH);
      } else {
        console.error("No se pudo cargar C2.png");
      }
    } catch (e) {
      console.error("Error al cargar C2.png:", e);
    }

    // Información de la empresa (lado izquierdo)
    let yInfo = yLogo + logoH + 12; // Posición debajo del logo
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(0, 0, 0);
    doc.text("OLGROUP SAS", left, yInfo);
    yInfo += 12;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text("NIT 900.840.594-1", left, yInfo);
    yInfo += 12;
    doc.text("CR 51 7 SUR 30 BRR GUAYABAL", left, yInfo);
    yInfo += 12;
    doc.text("Tel: +57 3232267375", left, yInfo);
    yInfo += 12;
    doc.text("Medellín - Colombia", left, yInfo);

    // Tabla con datos del cliente (debajo de la información de la empresa)
    const tableStartY = yInfo + 16;
    const tableRowHeight = 18;
    const tableLeftColWidth = 100;
    const tableRightColWidth = 200;
    const tableMiddleColWidth = 150;
    const tableTotalWidth = tableLeftColWidth + tableMiddleColWidth + tableRightColWidth;
    const tableX = left;
    
    // Número de cotización: ubicar el título en la mitad vertical entre logos y el cuadro del cliente
    {
      const bottomOfLogos = Math.max(yLogo + logoH, yCert + certH);
      const cotizacionY = bottomOfLogos + (tableStartY - bottomOfLogos) / 2;
      doc.setFont("helvetica", "bold");
      doc.setFontSize(14);
      doc.setTextColor(...brandPrimary);
      doc.text("COTIZACIÓN", right, cotizacionY, { align: "right" });
      if (quoteNumber) {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11);
        doc.setTextColor(0, 0, 0);
        doc.text(`No. ${quoteNumber}`, right, cotizacionY + 16, { align: "right" });
      }
    }
    
    // Dibujar borde de la tabla
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.5);
    const tableHeight = tableRowHeight * 3;
    doc.rect(tableX, tableStartY, tableTotalWidth, tableHeight);
    
    // Línea vertical que separa las columnas
    const verticalLineX = tableX + tableLeftColWidth + tableMiddleColWidth;
    doc.line(verticalLineX, tableStartY, verticalLineX, tableStartY + tableHeight);
    
    // Líneas horizontales entre filas
    doc.line(tableX, tableStartY + tableRowHeight, tableX + tableLeftColWidth + tableMiddleColWidth, tableStartY + tableRowHeight);
    doc.line(tableX, tableStartY + tableRowHeight * 2, tableX + tableLeftColWidth + tableMiddleColWidth, tableStartY + tableRowHeight * 2);
    
    // Líneas horizontales en la columna derecha (para Contacto, email, Celular)
    const rightColStartY = tableStartY;
    doc.line(verticalLineX, rightColStartY + tableRowHeight, tableX + tableTotalWidth, rightColStartY + tableRowHeight);
    doc.line(verticalLineX, rightColStartY + tableRowHeight * 2, tableX + tableTotalWidth, rightColStartY + tableRowHeight * 2);
    
    // Texto de la tabla - Columna izquierda
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("Cliente:", tableX + 4, tableStartY + 12);
    doc.text("Nit:", tableX + 4, tableStartY + tableRowHeight + 12);
    doc.text("Fecha:", tableX + 4, tableStartY + tableRowHeight * 2 + 12);
    
    // Valores de la columna izquierda
    doc.setFont("helvetica", "normal");
    doc.text(clientName || "", tableX + tableLeftColWidth + 4, tableStartY + 12);
    doc.text(clientNit || "", tableX + tableLeftColWidth + 4, tableStartY + tableRowHeight + 12);
    doc.text(formatDateHuman(dateStr), tableX + tableLeftColWidth + 4, tableStartY + tableRowHeight * 2 + 12);
    
    // Texto de la columna derecha (Contacto, Email, Celular)
    const rightColX = verticalLineX + 4;
    const contactoNombre = clientContactos && clientContactos.length > 0 ? clientContactos[0] : "";
    const labelWidth = 55; // Ancho para las etiquetas
    const valueX = rightColX + labelWidth; // Posición X para los valores
    const maxValueWidth = tableRightColWidth - labelWidth - 8; // Ancho máximo para los valores (restar padding)
    
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("Contacto:", rightColX, rightColStartY + 12);
    doc.text("Email:", rightColX, rightColStartY + tableRowHeight + 12);
    doc.text("Celular:", rightColX, rightColStartY + tableRowHeight * 2 + 12);
    
    // Valores de la columna derecha (con ajuste de texto para que quepa)
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    
    // Contacto
    const contactoLines = doc.splitTextToSize(contactoNombre || "", maxValueWidth);
    doc.text(contactoLines[0] || "", valueX, rightColStartY + 12);
    
    // Email
    const emailLines = doc.splitTextToSize(clientEmail || "", maxValueWidth);
    doc.text(emailLines[0] || "", valueX, rightColStartY + tableRowHeight + 12);
    
    // Celular
    const celularLines = doc.splitTextToSize(clientCelular || "", maxValueWidth);
    doc.text(celularLines[0] || "", valueX, rightColStartY + tableRowHeight * 2 + 12);
    
    // Ajustar la posición del separador después de la tabla
    const separatorY = tableStartY + tableHeight + 18;

    // Separador sutil con color de marca
    doc.setDrawColor(...brandPrimary);
    doc.line(left, separatorY, 555, separatorY);
    let y = separatorY + 18;

    // Productos
    products.forEach((p, idx) => {
      if (idx > 0) {
        y += 8;
      }
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11);
      doc.setTextColor(...brandPrimary);
      doc.text(p.nombre, left, y);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(0, 0, 0);
      y += 12;

      // Obtener el valor unitario seleccionado para este producto
      const selectedUnit = p._selectedUnit || 'vrUnitUSD';
      
      // Mapear el nombre de la columna según el valor seleccionado
      const unitColumnNames = {
        'vrUnit1': 'Vr. Unitario 1',
        'vrUnit2': 'Vr. Unitario 2',
        'vrUnit3': 'Vr. Unitario 3',
        'vrUnitUSD': 'Vr. Unitario USD'
      };
      const selectedUnitName = unitColumnNames[selectedUnit] || 'Vr. Unitario USD';

      const tableData = p.procesos.map((row) => {
        // Calcular el valor total usando el valor unitario seleccionado
        const qty = parseNumStrict(row.cantidad);
        const unitValue = parseNumStrict(row[selectedUnit]);
        let finalTotal = 0;
        
        // Calcular usando el valor unitario seleccionado
        // Solo si el valor unitario tiene un valor válido
        if (unitValue > 0 && qty > 0) {
          finalTotal = unitValue * qty;
        }
        // Si el valor unitario está vacío o es 0, el total será 0
        
        return [
          String(row.codigo ?? ""),
          String(row.analisis ?? ""),
          String(row.metodo ?? ""),
          String((row.cMtra ?? row.cMtra_g) ?? ""),
          String(row.cantidad ?? ""),
          String(row[selectedUnit] ?? ""),
          String(finalTotal > 0 ? formatMoney(finalTotal) : "")
        ];
      });

      doc.autoTable({
        head: [["Código", "Análisis", "Método", "C. Mtra. [g]", "Cant.", selectedUnitName, "Vr. Total"]],
        body: tableData,
        startY: y,
        margin: { top: topMargin, right: left, bottom: bottomMargin, left: left },
        styles: { font: "helvetica", fontStyle: "normal", fontSize: 8, textColor: [0, 0, 0] },
        headStyles: { font: "helvetica", fontStyle: "bold", fontSize: 8, fillColor: brandPrimary, textColor: [255, 255, 255] },
        columnStyles: {
          3: { halign: "right" },
          4: { halign: "right" },
          5: { halign: "right" },
          6: { halign: "right" }
        },
        didDrawPage: (data) => {},
        willDrawCell: (data) => {}
      });
      y = doc.lastAutoTable.finalY + 16;

      // Subtotal por producto (alineado a la derecha bajo la tabla)
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...brandSecondary);
      doc.text(`Subtotal: ${formatMoney(p._sumTotal)}`, right, y, { align: "right" });
      doc.setTextColor(0, 0, 0);
      doc.setFont("helvetica", "normal");
      y += 12;
    });

    // Total general
    y += 4;
    doc.setDrawColor(...brandPrimary);
    doc.line(left, y, 555, y);
    y += 16;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.setTextColor(...brandSecondary);
    doc.text(`Total general (COP): ${formatMoney(totalGeneral)}`, left, y);
    doc.setTextColor(0, 0, 0);

    // Nuevo contenido después del precio total
    y += 24;
    
    const textWidth = pageWidth - 2 * left;
    
    // CONDICIONES DEL SERVICIO
    y = checkPageBreak(y, 50);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("CONDICIONES DEL SERVICIO:", left, y);
    y += 12;
    if (nota) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      y = justifyText(doc, nota, left, y, textWidth, 9, pageHeight, bottomMargin);
      y += 20;
    } else {
      y += 20;
    }

    // ENVÍO DE LAS MUESTRAS
    y = checkPageBreak(y, 80);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("ENVÍO DE LAS MUESTRAS", left, y);
    y += 12;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    const envioTexto = "Las muestras se deberán enviar en un empaque apropiado en lo posible que evite el paso de la luz para garantizar la integridad de la muestra. Las muestras se deberán enviar en el formato de solicitud de servicio \"adjunto\" adecuadamente diligenciado a la dirección Carrera 51 7 sur 30, barrio Guayabal, Medellín (Ant.)";
    y = justifyText(doc, envioTexto, left, y, textWidth, 9, pageHeight, bottomMargin);
    y += 20;

    // DURACIÓN DE ANALISIS
    y = checkPageBreak(y, 50);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("DURACIÓN DE ANALISIS:", left, y);
    y += 12;
    if (duracionAnalisis) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      const duracionConDias = `${duracionAnalisis} Días`;
      y = justifyText(doc, duracionConDias, left, y, textWidth, 9, pageHeight, bottomMargin);
      y += 20;
    } else {
      y += 20;
    }

    // CONDICIÓN DE PAGO
    y = checkPageBreak(y, 50);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("CONDICIÓN DE PAGO:", left, y);
    y += 12;
    if (clientFormaPago) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      y = justifyText(doc, clientFormaPago, left, y, textWidth, 9, pageHeight, bottomMargin);
      y += 20;
    } else {
      y += 20;
    }

    // Texto sobre transferencia bancaria
    y = checkPageBreak(y, 100);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    const transferenciaTexto = "Para realizar la solicitud del servicio se deberá cancelar la totalidad de lo cotizado y solicitado por medio de una transferencia bancaria a la cuenta de ahorros: 01943248679, Bancolombia a nombre de OLGROUP S.A.S con Nit: 900840594-1. Una vez realizada la consignación puede hacerla llegar al correo asesorcomercial1@olgroup.co con copia al correo facturacion@olgroup.co. Para estudio de crédito hacer la solicitud a facturacion@olgroup.co.";
    y = justifyText(doc, transferenciaTexto, left, y, textWidth, 9, pageHeight, bottomMargin);
    y += 20;

    // TIEMPO DE RETENCIÓN Y DISPOSICIÓN FINAL DE LAS MUESTRAS
    y = checkPageBreak(y, 50);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("TIEMPO DE RETENCIÓN Y DISPOSICIÓN FINAL DE LAS MUESTRAS", left, y);
    y += 12;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    const retencionTexto = "Perecederos y aguas 8 días, no perecederos 1 mes.";
    y = justifyText(doc, retencionTexto, left, y, textWidth, 9, pageHeight, bottomMargin);
    y += 20;

    // VIGENCIA
    y = checkPageBreak(y, 50);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("VIGENCIA", left, y);
    y += 12;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    const vigenciaTexto = "La lista de precios adjunta es de carácter CONFIDENCIAL y tendrá una vigencia de 2025.";
    y = justifyText(doc, vigenciaTexto, left, y, textWidth, 9, pageHeight, bottomMargin);
    y += 9; // Reducir espacio después de VIGENCIA

    // Texto final
    y = checkPageBreak(y, 80);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    const contactoTexto = "Cualquier inquietud será atendida en el +57 3232267375.";
    doc.text(contactoTexto, left, y);
    y += 9; // Espacio mínimo entre líneas
    const correoTexto = "Si desea responder a este correo, por favor hacerlo a comercial@olgroup.co";
    doc.text(correoTexto, left, y);
    y += 9; // Espacio mínimo entre líneas
    const agradecimientoTexto = "Le agradecemos por utilizar los servicios de OLGROUP S.A.S.";
    doc.text(agradecimientoTexto, left, y);
    y += 24;
    y += 60; // Agregar 5 espacios adicionales (12 puntos cada uno = 60 puntos) antes de la firma

    // Firma
    y = checkPageBreak(y, 30);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text("FRANCISCO JAVIER OLAYA RAMIREZ", left, y);
    y += 12;
    doc.setFont("helvetica", "bold");
    doc.text("Gerente Comercial", left, y);

    // Nombre del archivo incluyendo el número de consecutivo si está disponible
    const filename = quoteNumber 
      ? `Cotizacion_${quoteNumber}_${sanitizeFilename(clientName)}_${dateStr}.pdf`
      : `Cotizacion_${sanitizeFilename(clientName)}_${dateStr}.pdf`;
    doc.save(filename);
  }

  // ==================== GESTIÓN DE CONSECUTIVOS DE COTIZACIONES ====================

  function getQuoteCounter() {
    try {
      const stored = localStorage.getItem("olgroup_quote_counter");
      if (stored) {
        const data = JSON.parse(stored);
        const currentYear = new Date().getFullYear().toString().slice(-2);
        // Si el año cambió, reiniciar el contador
        if (data.year !== currentYear) {
          return { year: currentYear, counter: 0 };
        }
        return data;
      }
      return { year: new Date().getFullYear().toString().slice(-2), counter: 0 };
    } catch {
      return { year: new Date().getFullYear().toString().slice(-2), counter: 0 };
    }
  }

  function incrementQuoteCounter() {
    const data = getQuoteCounter();
    data.counter += 1;
    try {
      localStorage.setItem("olgroup_quote_counter", JSON.stringify(data));
    } catch {
      console.error("No se pudo guardar el contador de cotizaciones");
    }
    return data;
  }

  function generateQuoteNumber() {
    const data = incrementQuoteCounter();
    const year = data.year;
    const number = data.counter.toString().padStart(3, "0");
    return `COT-${year}-${number}`;
  }

  function saveQuote(quote) {
    const key = "olgroup_quotes";
    const list = JSON.parse(localStorage.getItem(key) || "[]");
    list.unshift(quote);
    localStorage.setItem(key, JSON.stringify(list));
  }

  function getQuotes() {
    try {
      return JSON.parse(localStorage.getItem("olgroup_quotes") || "[]");
    } catch {
      return [];
    }
  }

  function deleteQuote(id) {
    const list = getQuotes().filter((q) => q.id !== id);
    localStorage.setItem("olgroup_quotes", JSON.stringify(list));
  }

  function getHistoryFilters() {
    return {
      consecutivo: $filterConsecutivo.value.trim().toUpperCase(),
      client: $filterClient.value.trim().toLowerCase(),
      product: $filterProduct.value.trim().toLowerCase(),
      from: $filterFrom.value,
      to: $filterTo.value
    };
  }

  function renderHistory(filters) {
    const all = getQuotes();
    const list = (filters ? applyFilters(all, filters) : all).slice();
    $historyTableBody.innerHTML = "";

    if (list.length === 0) {
      $noHistory.classList.remove("d-none");
      return;
    }
    $noHistory.classList.add("d-none");

    list.forEach((q, idx) => {
      const tr = document.createElement("tr");
      const productsText = q.products.map((p) => p.nombre).join(", ");
      // Mostrar el consecutivo de la cotización o el número secuencial si no existe
      const displayNumber = q.quoteNumber || (idx + 1);
      tr.innerHTML = `
        <td>${escapeHtml(displayNumber)}</td>
        <td>${formatDateHuman(q.date)}</td>
        <td>${escapeHtml(q.clientName)}</td>
        <td>${escapeHtml(productsText)}</td>
        <td class="text-center">${formatMoney(q.totalCOP != null ? q.totalCOP : q.totalUSD)}</td>
        <td class="text-nowrap">
          <button class="btn btn-sm btn-outline-primary me-1" data-action="view" data-id="${q.id}">PDF</button>
          <button class="btn btn-sm btn-outline-danger" data-action="delete" data-id="${q.id}">Borrar</button>
        </td>
      `;
      $historyTableBody.appendChild(tr);
    });

    $historyTableBody.querySelectorAll("button").forEach((btn) => {
      btn.addEventListener("click", () => onHistoryAction(btn.dataset.action, btn.dataset.id));
    });

    // Actualizar estadísticas al renderizar historial completo
    try { renderStats(); } catch {}
  }

  async function exportHistoryToXlsx(filters) {
    if (typeof XLSX === "undefined") {
      showAlert("No se pudo cargar la librería de Excel. Verifica tu conexión.", "error");
      return;
    }
    const all = getQuotes();
    const list = (filters ? applyFilters(all, filters) : all).slice();
    if (list.length === 0) {
      showAlert("No hay registros para exportar.", "info");
      return;
    }
    
    // Preparar datos como arrays (igual que clientes, contactos y usuarios)
    const historyRows = list.map((q, idx) => [
      q.quoteNumber || (idx + 1),
      formatDateHuman(q.date),
      q.clientName || "",
      (q.products || []).map((p) => p.nombre).join(", "),
      Number((q.totalCOP != null ? q.totalCOP : (q.totalUSD != null ? q.totalUSD : 0)))
    ]);

    const headers = ["Consecutivo", "Fecha", "Cliente", "Productos", "Total COP"];
    const filename = `Historial_Cotizaciones_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // Intentar usar ExcelJS con estilos (igual que clientes, contactos y usuarios)
    const success = await applyHeaderStylesWithExcelJS(
      historyRows, 
      headers, 
      filename, 
      "Historial",
      [12, 15, 30, 50, 15]
    );
    if (success) {
      return; // Si ExcelJS funcionó, salir
    }

    // Si ExcelJS no está disponible, usar XLSX sin estilos (igual que las otras funciones)
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      headers,
      ...historyRows
    ]);
    
    // Ajustar ancho de columnas
    ws['!cols'] = [
      { wch: 12 }, // Consecutivo
      { wch: 15 }, // Fecha
      { wch: 30 }, // Cliente
      { wch: 50 }, // Productos
      { wch: 15 }  // Total COP
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, "Historial");
    XLSX.writeFile(wb, filename);
  }

  // Estadísticas: productos más cotizados
  function computeTopProducts(limit = 10) {
    const quotes = getQuotes();
    const countByName = new Map();
    quotes.forEach((q) => {
      (q.products || []).forEach((p) => {
        const name = String(p.nombre || "").trim();
        if (!name) return;
        countByName.set(name, (countByName.get(name) || 0) + 1);
      });
    });
    const sorted = Array.from(countByName.entries()).sort((a, b) => b[1] - a[1]).slice(0, limit);
    return { labels: sorted.map((x) => x[0]), data: sorted.map((x) => x[1]) };
  }

  function renderStats() {
    if (!$statsCanvas) return;
    const { labels, data } = computeTopProducts(10);
    const hasData = labels.length > 0;
    if ($noStatsEl) $noStatsEl.classList.toggle("d-none", hasData);
    $statsCanvas.classList.toggle("d-none", !hasData);
    if (!hasData) {
      if (topProductsChart) { try { topProductsChart.destroy(); } catch {} topProductsChart = null; }
      return;
    }
    const ctx = $statsCanvas.getContext("2d");
    if (topProductsChart) { try { topProductsChart.destroy(); } catch {} }
    topProductsChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          label: "Veces cotizado",
          data,
          backgroundColor: "rgba(68, 194, 196, 0.6)",
          borderColor: "rgb(68, 194, 196)",
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: { beginAtZero: true, ticks: { precision: 0, stepSize: 1 } }
        },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (ctx) => ` ${ctx.parsed.y} cotizaciones` } }
        }
      }
    });
  }

  async function loadLogo() {
    // 1) Intentar usar la imagen existente en el DOM para evitar CORS
    try {
      const el = document.querySelector(".brand-logo");
      if (el && el.complete && el.naturalWidth > 0) {
        try { return { el, dataUrl: elementToDataUrl(el) }; } catch { return { el }; }
      }
    } catch {}
    // 2) Probar una lista de rutas comunes para el logo
    const candidates = [
      "assets/images/logoolg.png",
      "logoolg.png"
    ];
    for (const src of candidates) {
      try {
        const img = await loadImageElement(src);
        try { return { el: img, dataUrl: elementToDataUrl(img) }; } catch { return { el: img }; }
      } catch {}
    }
    return null;
  }

  function getLogoDataUrlSync(logo) {
    if (!logo) return null;
    if (logo.dataUrl) return logo.dataUrl;
    try {
      if (logo.el) return elementToDataUrl(logo.el);
    } catch {}
    return null;
  }

  function elementToDataUrl(imgEl) {
    const w = Math.max(1, imgEl.naturalWidth || 64);
    const h = Math.max(1, imgEl.naturalHeight || 64);
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(imgEl, 0, 0, w, h);
    return canvas.toDataURL("image/png");
  }


  function loadImageElement(src) {
    return new Promise((resolve, reject) => {
      // Para file://, no usar crossOrigin ya que causa problemas de CORS
      const isFileProtocol = window.location.protocol === 'file:';
      
      // Obtener la ruta base del archivo HTML
      const basePath = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1);
      
      const candidates = [
        src,
        `./${src}`,
        `${basePath}${src}`,
        ...(isFileProtocol ? [] : [`/${src}`]) // Solo usar rutas absolutas si no es file://
      ];
      
      let currentIndex = 0;
      const tryNext = () => {
        if (currentIndex >= candidates.length) {
          reject(new Error(`No se pudo cargar ${src} desde ninguna ruta`));
          return;
        }
        
        const img = new Image();
        // Solo usar crossOrigin si no es file://
        if (!isFileProtocol) {
          try { img.crossOrigin = "anonymous"; } catch {}
        }
        
        img.onload = () => {
          if (img.complete && img.naturalWidth > 0 && img.naturalHeight > 0) {
            resolve(img);
          } else {
            currentIndex++;
            tryNext();
          }
        };
        
        img.onerror = () => {
          currentIndex++;
          tryNext();
        };
        
        img.src = candidates[currentIndex];
      };
      
      tryNext();
    });
  }

  // Calcula tamaño de render conservando proporción dentro de un límite
  function getLogoRenderSize(logo, maxW, maxH) {
    let naturalW = 64;
    let naturalH = 64;
    try {
      if (logo && logo.el) {
        naturalW = Math.max(1, Number(logo.el.naturalWidth || logo.el.width || 64));
        naturalH = Math.max(1, Number(logo.el.naturalHeight || logo.el.height || 64));
      }
    } catch {}
    const scale = Math.max(0.0001, Math.min(maxW / naturalW, maxH / naturalH));
    const w = Math.round(naturalW * scale);
    const h = Math.round(naturalH * scale);
    return { w, h };
  }

  async function onHistoryAction(action, id) {
    const list = getQuotes();
    const q = list.find((x) => x.id === id);
    if (!q) return;
    if (action === "view") {
      // Reconstruir productos desde nombres y subtotales. Usamos catálogo actual para filas.
      const selected = (q.products || [])
        .map((qp) => productos.find((p) => p.id === qp.id) || null)
        .filter(Boolean)
        .map(recomputeTotalsForProduct);
      // Asegurar que el logo esté disponible también al abrir desde historial
      if (!logoAsset) {
        try { logoAsset = await loadLogo(); } catch {}
      }
      await generatePDF({ 
        clientName: q.clientName, 
        clientEmail: q.clientEmail || "", 
        clientNit: q.clientNit || "",
        clientContactos: q.clientContactos || (q.clientContacto ? [q.clientContacto] : []),
        clientCelular: q.clientCelular || "",
        clientFormaPago: q.clientFormaPago || "",
        duracionAnalisis: q.duracionAnalisis || "",
        nota: q.nota || "",
        dateStr: q.date, 
        quoteNumber: q.quoteNumber || "",
        products: selected, 
        totalGeneral: (q.totalCOP != null ? q.totalCOP : q.totalUSD), 
        logo: logoAsset 
      });
    } else if (action === "delete") {
      showConfirm("¿Borrar esta cotización del historial?", "Borrar").then((ok) => {
        if (ok) {
          deleteQuote(id);
          renderHistory(getHistoryFilters());
          renderStats();
        }
      });
    }
  }

  function showAlert(message, type = "info", title = null) {
    const modalEl = document.getElementById("alertModal");
    const titleEl = document.getElementById("alertModalTitle");
    const textEl = document.getElementById("alertModalText");
    const iconEl = document.getElementById("alertModalIcon");
    const headerEl = document.getElementById("alertModalHeader");
    const acceptBtn = document.getElementById("alertModalAcceptBtn");
    
    if (!modalEl || !textEl || typeof bootstrap === "undefined") {
      // Fallback seguro si no carga Bootstrap
      window.alert(message);
      return;
    }

    // Configurar título
    if (titleEl) {
      titleEl.textContent = title || (type === "error" ? "Error" : type === "success" ? "Éxito" : type === "warning" ? "Advertencia" : "Información");
    }

    // Configurar mensaje (soporta saltos de línea)
    if (textEl) {
      // Convertir saltos de línea a <br> para HTML
      const messageWithBreaks = message.replace(/\n/g, '<br>');
      textEl.innerHTML = messageWithBreaks;
    }

    // Configurar icono y colores según el tipo
    const typeConfig = {
      success: {
        icon: "bi-check-circle-fill",
        iconColor: "text-success",
        headerClass: "bg-success text-white",
        btnClass: "btn-success"
      },
      error: {
        icon: "bi-exclamation-triangle-fill",
        iconColor: "text-danger",
        headerClass: "bg-danger text-white",
        btnClass: "btn-danger"
      },
      warning: {
        icon: "bi-exclamation-circle-fill",
        iconColor: "text-warning",
        headerClass: "bg-warning text-dark",
        btnClass: "btn-warning"
      },
      info: {
        icon: "bi-info-circle-fill",
        iconColor: "text-info",
        headerClass: "bg-info text-white",
        btnClass: "btn-info"
      }
    };

    const config = typeConfig[type] || typeConfig.info;

    // Aplicar estilos
    if (iconEl) {
      iconEl.className = `bi ${config.icon} fs-3 ${config.iconColor}`;
    }
    if (headerEl) {
      headerEl.className = `modal-header ${config.headerClass}`;
    }
    if (acceptBtn) {
      acceptBtn.className = `btn ${config.btnClass}`;
    }

    // Mostrar modal
    const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
    modal.show();
  }

  function showConfirm(message, acceptLabel) {
    return new Promise((resolve) => {
      const modalEl = document.getElementById("confirmModal");
      const msgEl = document.getElementById("confirmMessage");
      const acceptBtn = document.getElementById("confirmAcceptBtn");
      if (!modalEl || !msgEl || !acceptBtn || typeof bootstrap === "undefined") {
        // Fallback seguro si no carga Bootstrap
        resolve(window.confirm(message || "¿Confirmar?"));
        return;
      }
      // Convertir saltos de línea a <br> para HTML
      const messageWithBreaks = (message || "¿Confirmar?").replace(/\n/g, '<br>');
      msgEl.innerHTML = messageWithBreaks;
      if (acceptLabel) acceptBtn.textContent = acceptLabel;
      
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl, {
        backdrop: true,
        keyboard: true
      });
      
      // Cuando el modal se muestra, asegurar z-index mayor que otros modales
      const onShown = () => {
        // Forzar z-index mayor para el modal de confirmación
        modalEl.style.zIndex = '1065';
        
        // Ajustar backdrop solo si hay múltiples modales abiertos
        // Esto solo debe afectar al backdrop del modal de confirmación
        const backdrops = document.querySelectorAll('.modal-backdrop');
        if (backdrops.length > 1) {
          // El último backdrop (del modal de confirmación) debe tener z-index mayor
          const lastBackdrop = backdrops[backdrops.length - 1];
          lastBackdrop.style.zIndex = '1060';
          // Asegurar que el backdrop anterior no bloquee la interacción
          if (backdrops.length > 1) {
            backdrops[backdrops.length - 2].style.zIndex = '1050';
          }
        }
        
        modalEl.removeEventListener('shown.bs.modal', onShown);
      };
      
      let resolved = false;
      const onAccept = () => {
        resolved = true;
        cleanup();
        modal.hide();
        resolve(true);
      };
      const onHidden = () => {
        if (!resolved) {
          cleanup();
          resolve(false);
        }
        // Restaurar z-index por defecto
        modalEl.style.zIndex = '';
      };
      function cleanup() {
        acceptBtn.removeEventListener("click", onAccept);
        modalEl.removeEventListener("hidden.bs.modal", onHidden);
      }
      acceptBtn.addEventListener("click", onAccept, { once: true });
      modalEl.addEventListener("hidden.bs.modal", onHidden, { once: true });
      modalEl.addEventListener("shown.bs.modal", onShown, { once: true });
      modal.show();
    });
  }

  function applyFilters(list, { consecutivo, client, product, from, to }) {
    return list.filter((q) => {
      const quoteNumber = (q.quoteNumber || "").toUpperCase();
      const matchesConsecutivo = consecutivo ? quoteNumber.includes(consecutivo) : true;
      const matchesClient = client ? (q.clientName || "").toLowerCase().includes(client) : true;
      const matchesProduct = product ? (q.products || []).some((p) => (p.nombre || "").toLowerCase().includes(product)) : true;
      const date = q.date || "";
      const matchesFrom = from ? date >= from : true;
      const matchesTo = to ? date <= to : true;
      return matchesConsecutivo && matchesClient && matchesProduct && matchesFrom && matchesTo;
    });
  }

  function onScrollToggleScrollTop() {
    if (!$scrollTopBtn) return;
    if (window.scrollY > 200) {
      $scrollTopBtn.classList.add("show");
    } else {
      $scrollTopBtn.classList.remove("show");
    }
  }

  // Utilidades
  function formatMoney(n) {
    const num = Number(n) || 0;
    return num.toLocaleString("es-CO", { style: "currency", currency: "COP", minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }
  function formatNumber(n) {
    const num = Number(n) || 0;
    return num.toLocaleString("es-CO", { maximumFractionDigits: 4 });
  }
  function parseNumStrict(v) {
    if (v == null) return 0;
    if (typeof v === "number" && isFinite(v)) return v;
    const s = String(v).replace(/[^0-9,.-]/g, "").replace(/\.(?=.*\.)/g, "").replace(/,(?=\d{3}\b)/g, "");
    const val = Number(s.replace(",", "."));
    return isNaN(val) ? 0 : val;
  }

  function saveCatalog(list) {
    try { localStorage.setItem("olgroup_catalog", JSON.stringify(list || [])); } catch {}
  }
  function loadCatalog() {
    try { return JSON.parse(localStorage.getItem("olgroup_catalog") || "null"); } catch { return null; }
  }

  // ==================== ELIMINAR PRODUCTOS ====================

  function openDeleteProductsModal() {
    // Validar que solo administradores puedan acceder
    if (!currentUser || currentUser.role !== "admin") {
      showAlert("No tienes permisos para realizar esta acción.", "error");
      return;
    }
    if (!$deleteProductsModal || !$deleteProductsList) return;
    
    // Renderizar lista de productos con checkboxes
    $deleteProductsList.innerHTML = "";
    
    if (productos.length === 0) {
      $deleteProductsList.innerHTML = '<div class="text-center text-muted py-3">No hay productos para eliminar.</div>';
      if ($btnConfirmDeleteProducts) $btnConfirmDeleteProducts.disabled = true;
      if ($btnDeleteAllProducts) $btnDeleteAllProducts.disabled = true;
    } else {
      productos.forEach((product) => {
        const item = document.createElement("div");
        item.className = "list-group-item";
        item.innerHTML = `
          <div class="form-check">
            <input class="form-check-input product-delete-checkbox" type="checkbox" value="${product.id}" id="delete-${product.id}">
            <label class="form-check-label" for="delete-${product.id}">
              ${escapeHtml(product.nombre || "")}
            </label>
          </div>
        `;
        $deleteProductsList.appendChild(item);
      });
      
      if ($btnDeleteAllProducts) $btnDeleteAllProducts.disabled = false;
      updateDeleteProductsButtonState();
    }

    // Agregar event listeners a los checkboxes
    $deleteProductsList.querySelectorAll(".product-delete-checkbox").forEach((checkbox) => {
      checkbox.addEventListener("change", updateDeleteProductsButtonState);
    });

    // Mostrar modal
    if (typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance($deleteProductsModal);
      modal.show();
    }
  }

  function selectAllProductsToDelete() {
    if (!$deleteProductsList) return;
    $deleteProductsList.querySelectorAll(".product-delete-checkbox").forEach((checkbox) => {
      checkbox.checked = true;
    });
    updateDeleteProductsButtonState();
  }

  function deselectAllProductsToDelete() {
    if (!$deleteProductsList) return;
    $deleteProductsList.querySelectorAll(".product-delete-checkbox").forEach((checkbox) => {
      checkbox.checked = false;
    });
    updateDeleteProductsButtonState();
  }

  function updateDeleteProductsButtonState() {
    if (!$deleteProductsList || !$btnConfirmDeleteProducts) return;
    const checked = $deleteProductsList.querySelectorAll(".product-delete-checkbox:checked");
    const count = checked.length;
    $btnConfirmDeleteProducts.disabled = count === 0;
    if ($deleteProductsCount) {
      $deleteProductsCount.textContent = count;
    }
  }

  async function deleteSelectedProducts() {
    // Validar que solo administradores puedan eliminar
    if (!currentUser || currentUser.role !== "admin") {
      showAlert("No tienes permisos para realizar esta acción.", "error");
      return;
    }
    if (!$deleteProductsList) return;
    
    const checked = $deleteProductsList.querySelectorAll(".product-delete-checkbox:checked");
    if (checked.length === 0) {
      showAlert("Por favor selecciona al menos un producto para eliminar.", "warning");
      return;
    }

    const idsToDelete = Array.from(checked).map((cb) => cb.value);
    const count = idsToDelete.length;
    
    const confirmed = await showConfirm(`¿Estás seguro de que deseas eliminar ${count} producto(s)?\n\nEsta acción no se puede deshacer.`, "Eliminar");
    if (!confirmed) {
      return;
    }

    // Eliminar productos seleccionados
    productos = productos.filter((p) => !idsToDelete.includes(p.id));
    
    // Limpiar selecciones de productos eliminados
    idsToDelete.forEach((id) => {
      selectedProductIds.delete(id);
      selectedUnitValues.delete(id);
    });

    // Guardar catálogo actualizado
    saveCatalog(productos);
    
    // Actualizar UI
    updateSelectionStateUI();
    renderProductsByLetter(currentLetter);
    
    // Cerrar modal
    if (typeof bootstrap !== "undefined" && $deleteProductsModal) {
      const modal = bootstrap.Modal.getOrCreateInstance($deleteProductsModal);
      modal.hide();
    }
    
    showAlert(`${count} producto(s) eliminado(s) exitosamente.`, "success");
  }

  async function deleteAllProducts() {
    // Validar que solo administradores puedan eliminar
    if (!currentUser || currentUser.role !== "admin") {
      showAlert("No tienes permisos para realizar esta acción.", "error");
      return;
    }
    if (productos.length === 0) {
      showAlert("No hay productos para eliminar.", "info");
      return;
    }

    const count = productos.length;
    const confirmed = await showConfirm(`¿Estás seguro de que deseas eliminar TODOS los ${count} producto(s)?\n\nEsta acción no se puede deshacer.`, "Eliminar Todos");
    if (!confirmed) {
      return;
    }

    // Eliminar todos los productos
    productos = [];
    
    // Limpiar todas las selecciones
    selectedProductIds.clear();
    selectedUnitValues.clear();

    // Guardar catálogo vacío
    saveCatalog(productos);
    
    // Actualizar UI
    updateSelectionStateUI();
    renderProductsByLetter(currentLetter);
    
    // Cerrar modal
    if (typeof bootstrap !== "undefined" && $deleteProductsModal) {
      const modal = bootstrap.Modal.getOrCreateInstance($deleteProductsModal);
      modal.hide();
    }
    
    showAlert(`Todos los ${count} producto(s) han sido eliminados exitosamente.`, "success");
  }

  function formatDateHuman(d) {
    const [y, m, da] = d.split("-").map((x) => Number(x));
    if (!y || !m || !da) return d;
    return `${da.toString().padStart(2, "0")}/${m.toString().padStart(2, "0")}/${y}`;
  }
  function sanitizeFilename(s) {
    return (s || "").replace(/[^a-z0-9\-_]+/gi, "_");
  }
  function escapeHtml(str) {
    return String(str)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  // ==================== FUNCIONES DE CLIENTES ====================

  function getClients() {
    try {
      return JSON.parse(localStorage.getItem("olgroup_clients") || "[]");
    } catch {
      return [];
    }
  }

  function saveClients(clients) {
    try {
      localStorage.setItem("olgroup_clients", JSON.stringify(clients || []));
    } catch {
      showAlert("No se pudo guardar los clientes. Verifica el espacio disponible.", "error");
    }
  }

  function saveClient() {
    if (!$clientFormModal || !$clientFormModal.checkValidity()) {
      $clientFormModal.reportValidity();
      return;
    }

    const nombre = $clientNombre.value.trim();
    const nit = $clientNit.value.trim();
    const correo = $clientCorreo.value.trim();
    const celular = $clientCelular ? $clientCelular.value.trim() : "";
    const formaPago = $clientFormaPago ? $clientFormaPago.value.trim() : "";
    const id = $clientId.value.trim();

    // Obtener todos los contactos
    const contactosInputs = $contactosContainer ? $contactosContainer.querySelectorAll(".contacto-input") : [];
    const contactos = Array.from(contactosInputs)
      .map(input => input.value.trim())
      .filter(contacto => contacto !== "");

    if (!nombre) {
      showAlert("El nombre es requerido.", "warning");
      return;
    }

    const clients = getClients();
    let updatedClients;

    if (id) {
      // Editar cliente existente
      updatedClients = clients.map((c) =>
        c.id === id
          ? { 
              ...c, 
              nombre, 
              nit, 
              contactos: contactos.length > 0 ? contactos : (c.contactos || []), 
              correo, 
              celular,
              formaPago,
              updatedAt: new Date().toISOString() 
            }
          : c
      );
    } else {
      // Crear nuevo cliente
      const newClient = {
        id: `CLIENT-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        nombre,
        nit,
        contactos: contactos,
        correo,
        celular,
        formaPago,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      updatedClients = [...clients, newClient];
    }

    saveClients(updatedClients);
    renderClients();
    closeClientModal();
  }

  function deleteClient(id) {
    showConfirm("¿Estás seguro de que deseas eliminar este cliente?", "Eliminar").then((ok) => {
      if (ok) {
        const clients = getClients().filter((c) => c.id !== id);
        saveClients(clients);
        renderClients();
      }
    });
  }

  // ==================== ELIMINAR CLIENTES (MÚLTIPLES) ====================

  function openDeleteClientsModal() {
    if (!$deleteClientsModal || !$deleteClientsList) return;
    
    const clients = getClients();
    
    // Renderizar lista de clientes con checkboxes
    $deleteClientsList.innerHTML = "";
    
    if (clients.length === 0) {
      $deleteClientsList.innerHTML = '<div class="text-center text-muted py-3">No hay clientes para eliminar.</div>';
      if ($btnConfirmDeleteClients) $btnConfirmDeleteClients.disabled = true;
    } else {
      clients.forEach((client) => {
        const item = document.createElement("div");
        item.className = "list-group-item";
        // Obtener contactos (puede ser array o string legacy)
        const contactos = Array.isArray(client.contactos) 
          ? client.contactos.map(c => typeof c === 'object' ? c.nombre : c).join(", ")
          : (client.contacto || "");
        item.innerHTML = `
          <div class="form-check">
            <input class="form-check-input client-delete-checkbox" type="checkbox" value="${client.id}" id="delete-client-${client.id}">
            <label class="form-check-label" for="delete-client-${client.id}">
              <strong>${escapeHtml(client.nombre || "")}</strong>
              ${client.nit ? `<br><small class="text-muted">NIT/CC: ${escapeHtml(client.nit)}</small>` : ''}
              ${contactos ? `<br><small class="text-muted">Contacto(s): ${escapeHtml(contactos)}</small>` : ''}
            </label>
          </div>
        `;
        $deleteClientsList.appendChild(item);
      });
      
      updateDeleteClientsButtonState();
    }

    // Agregar event listeners a los checkboxes
    $deleteClientsList.querySelectorAll(".client-delete-checkbox").forEach((checkbox) => {
      checkbox.addEventListener("change", updateDeleteClientsButtonState);
    });

    // Mostrar modal
    if (typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance($deleteClientsModal);
      modal.show();
    }
  }

  function selectAllClientsToDelete() {
    if (!$deleteClientsList) return;
    $deleteClientsList.querySelectorAll(".client-delete-checkbox").forEach((checkbox) => {
      checkbox.checked = true;
    });
    updateDeleteClientsButtonState();
  }

  function deselectAllClientsToDelete() {
    if (!$deleteClientsList) return;
    $deleteClientsList.querySelectorAll(".client-delete-checkbox").forEach((checkbox) => {
      checkbox.checked = false;
    });
    updateDeleteClientsButtonState();
  }

  function updateDeleteClientsButtonState() {
    if (!$deleteClientsList || !$btnConfirmDeleteClients || !$deleteClientsCount) return;
    const checked = $deleteClientsList.querySelectorAll(".client-delete-checkbox:checked");
    const count = checked.length;
    $btnConfirmDeleteClients.disabled = count === 0;
    $deleteClientsCount.textContent = count;
  }

  async function deleteSelectedClients() {
    if (!$deleteClientsList) return;
    
    const checked = $deleteClientsList.querySelectorAll(".client-delete-checkbox:checked");
    if (checked.length === 0) {
      showAlert("Por favor selecciona al menos un cliente para eliminar.", "warning");
      return;
    }

    const idsToDelete = Array.from(checked).map((cb) => cb.value);
    const count = idsToDelete.length;
    
    const confirmed = await showConfirm(`¿Estás seguro de que deseas eliminar ${count} cliente(s)?\n\nEsta acción no se puede deshacer.`, "Eliminar");
    if (!confirmed) {
      return;
    }

    // Eliminar clientes seleccionados
    const clients = getClients();
    const remainingClients = clients.filter((c) => !idsToDelete.includes(c.id));
    saveClients(remainingClients);
    
    // Actualizar UI
    renderClients();
    
    // Cerrar modal
    if (typeof bootstrap !== "undefined" && $deleteClientsModal) {
      const modal = bootstrap.Modal.getOrCreateInstance($deleteClientsModal);
      modal.hide();
    }
    
    showAlert(`${count} cliente(s) eliminado(s) exitosamente.`, "success");
  }

  function openClientModal(client = null) {
    if (!$clientModalTitle || !$clientId || !$clientNombre || !$clientNit || !$clientCorreo) return;

    if (client) {
      $clientModalTitle.textContent = "Editar Cliente";
      $clientId.value = client.id;
      $clientNombre.value = client.nombre || "";
      $clientNit.value = client.nit || "";
      $clientCorreo.value = client.correo || "";
      if ($clientCelular) $clientCelular.value = client.celular || "";
      if ($clientFormaPago) $clientFormaPago.value = client.formaPago || "";
      
      // Cargar contactos
      if ($contactosContainer) {
        $contactosContainer.innerHTML = "";
        const contactos = client.contactos || (client.contacto ? [client.contacto] : []);
        if (contactos.length === 0) {
          addContactoInput();
        } else {
          contactos.forEach((contacto, index) => {
            addContactoInput(contacto, index === 0);
          });
        }
      }
    } else {
      $clientModalTitle.textContent = "Nuevo Cliente";
      $clientId.value = "";
      $clientNombre.value = "";
      $clientNit.value = "";
      $clientCorreo.value = "";
      if ($clientCelular) $clientCelular.value = "";
      if ($clientFormaPago) $clientFormaPago.value = "";
      
      // Limpiar contactos
      if ($contactosContainer) {
        $contactosContainer.innerHTML = "";
        addContactoInput();
      }
    }

    const modalEl = document.getElementById("clientModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.show();
    }
  }

  function addContactoInput(value = "", isFirst = false) {
    if (!$contactosContainer) return;
    
    const div = document.createElement("div");
    div.className = "input-group mb-2";
    div.innerHTML = `
      <input type="text" class="form-control contacto-input" placeholder="Nombre de contacto" value="${escapeHtml(value)}">
      <button type="button" class="btn btn-outline-danger btn-remove-contacto" ${isFirst ? 'style="display: none;"' : ''}>
        <i class="bi bi-trash"></i>
      </button>
    `;
    $contactosContainer.appendChild(div);
    
    // Agregar evento al botón de eliminar
    const removeBtn = div.querySelector(".btn-remove-contacto");
    if (removeBtn) {
      removeBtn.addEventListener("click", () => {
        div.remove();
        updateRemoveButtons();
      });
    }
    
    updateRemoveButtons();
  }

  function updateRemoveButtons() {
    if (!$contactosContainer) return;
    const contactosInputs = $contactosContainer.querySelectorAll(".contacto-input");
    const removeButtons = $contactosContainer.querySelectorAll(".btn-remove-contacto");
    
    contactosInputs.forEach((input, index) => {
      const removeBtn = input.closest(".input-group").querySelector(".btn-remove-contacto");
      if (removeBtn) {
        removeBtn.style.display = contactosInputs.length > 1 ? "" : "none";
      }
    });
  }

  function closeClientModal() {
    const modalEl = document.getElementById("clientModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.hide();
    }
    if ($clientFormModal) $clientFormModal.reset();
  }

  function renderClients() {
    if (!$clientsTableBody || !$noClients) return;

    const clients = getClients();
    $clientsTableBody.innerHTML = "";

    if (clients.length === 0) {
      $noClients.classList.remove("d-none");
      return;
    }

    $noClients.classList.add("d-none");

    const isAdmin = currentUser && currentUser.role === "admin";
    
    clients.forEach((client, idx) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td>${escapeHtml(client.nombre || "")}</td>
        <td>${escapeHtml(client.nit || "")}</td>
        <td>${escapeHtml(client.correo || "")}</td>
        <td>${escapeHtml(client.celular || "")}</td>
        <td>${escapeHtml(client.formaPago || "")}</td>
        <td class="text-nowrap">
          <button class="btn btn-sm btn-outline-primary me-1" data-action="edit" data-id="${client.id}">
            <i class="bi bi-pencil"></i> Editar
          </button>
        </td>
      `;
      $clientsTableBody.appendChild(tr);
    });

    $clientsTableBody.querySelectorAll("button").forEach((btn) => {
      btn.addEventListener("click", () => {
        const action = btn.dataset.action;
        const id = btn.dataset.id;
        const client = clients.find((c) => c.id === id);
        if (action === "edit" && client) {
          openClientModal(client);
        }
      });
    });
  }

  function onClientsExcelSelected(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const clients = buildClientsFromWorkbook(wb);
        
        if (clients.length === 0) {
          showAlert("No se encontraron clientes válidos en el archivo Excel.", "warning");
          return;
        }

        // Confirmar si se deben agregar o reemplazar
        const existingClients = getClients();
        
        if (existingClients.length > 0) {
          const shouldAppend = await showConfirm(
            `Ya existen ${existingClients.length} cliente(s) registrados.\n\n` +
            `¿Deseas AGREGAR los ${clients.length} nuevo(s) cliente(s) a los existentes?\n\n` +
            `- Aceptar: Agregar nuevos (se evitarán duplicados por nombre)\n` +
            `- Cancelar: Reemplazar todos los clientes existentes`,
            "Agregar"
          );

          if (shouldAppend) {
            // Agregar nuevos clientes (evitando duplicados por nombre)
            const nombres = new Set(existingClients.map((c) => c.nombre.toLowerCase()));
            const nuevos = clients.filter((c) => !nombres.has(c.nombre.toLowerCase()));
            const duplicados = clients.length - nuevos.length;
            if (duplicados > 0) {
              showAlert(`${duplicados} cliente(s) ya existen y no fueron agregados.`, "warning");
            }
            saveClients([...existingClients, ...nuevos]);
            renderClients();
            showAlert(`Se agregaron ${nuevos.length} cliente(s) exitosamente.`, "success");
          } else {
            // Reemplazar todos
            const confirmed = await showConfirm(`¿Estás seguro de que deseas reemplazar todos los ${existingClients.length} cliente(s) existentes con los ${clients.length} nuevo(s)?`, "Reemplazar");
            if (confirmed) {
              saveClients(clients);
              renderClients();
              showAlert(`Se importaron ${clients.length} cliente(s) exitosamente.`, "success");
            }
          }
        } else {
          // No hay clientes existentes, importar directamente
          saveClients(clients);
          renderClients();
          showAlert(`Se importaron ${clients.length} cliente(s) exitosamente.`, "success");
        }
      } catch (err) {
        showAlert("No se pudo leer el archivo Excel. Verifica el formato.\n\nEl formato esperado es:\n- Nombre (requerido)\n- NIT/CC\n- Contacto\n- Correo", "error");
        console.error(err);
      } finally {
        e.target.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function buildClientsFromWorkbook(wb) {
    const clients = [];
    
    // Procesar solo la hoja CLIENTES (ignorar CLIENTES_CONTACTOS)
    const clientesSheetName = wb.SheetNames.find(name => 
      normalizeText(name) === "clientes" || 
      !normalizeText(name).includes("contacto")
    ) || wb.SheetNames[0]; // Si no encuentra, usar la primera hoja
    
    const clientesSheet = wb.Sheets[clientesSheetName];
    if (!clientesSheet) {
      return clients;
    }
    
    // Leer como JSON con encabezados
    const rows = XLSX.utils.sheet_to_json(clientesSheet, { header: 1, raw: false, defval: "" });
    if (!rows || rows.length === 0) return clients;

    // Buscar fila de encabezados en CLIENTES
    let headerRow = -1;
    let nombreIdx = -1;
    let nitIdx = -1;
    let correoIdx = -1;
    let celularIdx = -1;
    let formaPagoIdx = -1;

    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = ensureArray(rows[i]).map((c) => normalizeText(String(c || "")));
      const nombrePos = row.findIndex((c) => c.includes("nombre") && !c.includes("contacto"));
      const nitPos = row.findIndex((c) => c.includes("nit") || c.includes("cc") || c.includes("cedula"));
      const correoPos = row.findIndex((c) => c.includes("correo") || c.includes("email") || c.includes("mail"));
      const celularPos = row.findIndex((c) => c.includes("celular") || c.includes("telefono") || c.includes("movil"));
      const formaPagoPos = row.findIndex((c) => c.includes("forma") && c.includes("pago") || c.includes("pago"));

      if (nombrePos >= 0) {
        headerRow = i;
        nombreIdx = nombrePos;
        nitIdx = nitPos >= 0 ? nitPos : -1;
        correoIdx = correoPos >= 0 ? correoPos : -1;
        celularIdx = celularPos >= 0 ? celularPos : -1;
        formaPagoIdx = formaPagoPos >= 0 ? formaPagoPos : -1;
        break;
      }
    }

    if (headerRow < 0 || nombreIdx < 0) {
      // Sin encabezados: asumir primera fila es encabezado y segunda fila en adelante son datos
      if (rows.length > 1) {
        nombreIdx = 0;
        nitIdx = 1;
        correoIdx = 2;
        celularIdx = 3;
        formaPagoIdx = 4;
        headerRow = 0;
      } else {
        return clients;
      }
    }

    // Procesar filas de datos de clientes (sin contactos)
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = ensureArray(rows[i]);
      const nombre = String(row[nombreIdx] || "").trim();
      if (!nombre) continue; // Saltar filas sin nombre

      const nit = String(row[nitIdx] || "").trim();
      const correo = String(row[correoIdx] || "").trim();
      const celular = String(row[celularIdx] || "").trim();
      const formaPago = String(row[formaPagoIdx] || "").trim();

      clients.push({
        id: `CLIENT-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 9)}`,
        nombre,
        nit,
        contactos: [], // Los contactos se gestionan en la pestaña Contactos
        correo,
        celular,
        formaPago,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
    }
    
    return clients;
  }

  // Función auxiliar para aplicar estilos usando ExcelJS
  async function applyHeaderStylesWithExcelJS(data, headers, filename, sheetName, colWidths) {
    try {
      if (typeof ExcelJS === "undefined") {
        // Si ExcelJS no está disponible, usar XLSX sin estilos
        return false;
      }

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet(sheetName || "Datos");

      // Agregar headers con estilos
      const headerRow = worksheet.addRow(headers);
      headerRow.font = { size: 13, bold: true };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' } // Amarillo
      };
      headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

      // Agregar datos
      data.forEach(row => {
        worksheet.addRow(row);
      });

      // Ajustar ancho de columnas
      if (colWidths && colWidths.length > 0) {
        worksheet.columns.forEach((column, index) => {
          if (colWidths[index]) {
            column.width = colWidths[index];
          }
        });
      }

      // Guardar archivo
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      link.click();
      window.URL.revokeObjectURL(url);
      
      return true;
    } catch (e) {
      console.warn("Error al usar ExcelJS:", e);
      return false;
    }
  }

  async function exportClientsToExcel() {
    if (typeof XLSX === "undefined") {
      showAlert("No se pudo cargar la librería de Excel. Verifica tu conexión.", "error");
      return;
    }

    const clients = getClients();
    if (clients.length === 0) {
      showAlert("No hay clientes para exportar.", "info");
      return;
    }

    // Preparar datos para la hoja CLIENTES (solo datos del cliente, sin contactos)
    const clientesRows = clients.map((c) => [
      c.nombre || "",
      c.nit || "",
      c.correo || "",
      c.celular || "",
      c.formaPago || ""
    ]);

    const headers = ["Nombre", "NIT/CC", "Email", "Celular", "Forma de Pago"];
    const filename = `Clientes_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // Intentar usar ExcelJS con estilos
    const success = await applyHeaderStylesWithExcelJS(
      clientesRows, 
      headers, 
      filename, 
      "CLIENTES",
      [25, 15, 30, 15, 15]
    );
    if (success) {
      return; // Si ExcelJS funcionó, salir
    }

    // Si ExcelJS no está disponible, usar XLSX sin estilos
    const wb = XLSX.utils.book_new();
    const wsClientes = XLSX.utils.aoa_to_sheet([
      headers,
      ...clientesRows
    ]);
    
    // Ajustar ancho de columnas para CLIENTES
    wsClientes['!cols'] = [
      { wch: 25 }, // Nombre
      { wch: 15 }, // NIT/CC
      { wch: 30 }, // Email
      { wch: 15 }, // Celular
      { wch: 15 }  // Forma de Pago
    ];
    
    XLSX.utils.book_append_sheet(wb, wsClientes, "CLIENTES");
    XLSX.writeFile(wb, filename);
  }

  // ==================== GESTIÓN DE CONTACTOS ====================

  function getContactos() {
    try {
      return JSON.parse(localStorage.getItem("olgroup_contactos") || "[]");
    } catch {
      return [];
    }
  }

  function saveContactos(contactos) {
    try {
      localStorage.setItem("olgroup_contactos", JSON.stringify(contactos || []));
    } catch {
      showAlert("No se pudo guardar los contactos. Verifica el espacio disponible.", "error");
    }
  }

  function renderContactos() {
    if (!$contactosTableBody || !$noContactos) return;

    const contactos = getContactos();
    $contactosTableBody.innerHTML = "";

    if (contactos.length === 0) {
      $noContactos.classList.remove("d-none");
      return;
    }

    $noContactos.classList.add("d-none");

    const isAdmin = currentUser && currentUser.role === "admin";

    contactos.forEach((contacto, idx) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td>${escapeHtml(contacto.nombreCliente || "")}</td>
        <td>${escapeHtml(contacto.nombre || "")}</td>
        <td>${escapeHtml(contacto.correo || "")}</td>
        <td class="text-nowrap">
          <button class="btn btn-sm btn-outline-primary me-1" data-action="edit" data-id="${contacto.id}">
            <i class="bi bi-pencil"></i> Editar
          </button>
        </td>
      `;
      $contactosTableBody.appendChild(tr);
    });

    $contactosTableBody.querySelectorAll("button").forEach((btn) => {
      btn.addEventListener("click", () => {
        const action = btn.dataset.action;
        const id = btn.dataset.id;
        const contacto = contactos.find((c) => c.id === id);
        if (action === "edit" && contacto) {
          openContactoModal(contacto);
        }
      });
    });
  }

  function openContactoModal(contacto = null) {
    if (!$contactoModalTitle || !$contactoId || !$contactoNombreCliente || !$contactoNombre || !$contactoCorreo) return;

    if (contacto) {
      $contactoModalTitle.textContent = "Editar Contacto";
      $contactoId.value = contacto.id;
      $contactoNombreCliente.value = contacto.nombreCliente || "";
      $contactoNombre.value = contacto.nombre || "";
      $contactoCorreo.value = contacto.correo || "";
    } else {
      $contactoModalTitle.textContent = "Nuevo Contacto";
      $contactoId.value = "";
      $contactoNombreCliente.value = "";
      $contactoNombre.value = "";
      $contactoCorreo.value = "";
    }

    const modalEl = document.getElementById("contactoModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.show();
    }
  }

  function saveContacto() {
    if (!$contactoFormModal || !$contactoFormModal.checkValidity()) {
      $contactoFormModal.reportValidity();
      return;
    }

    const nombreCliente = $contactoNombreCliente.value.trim();
    const nombre = $contactoNombre.value.trim();
    const correo = $contactoCorreo.value.trim();
    const id = $contactoId.value.trim();

    if (!nombreCliente || !nombre) {
      showAlert("El nombre del cliente y el nombre del contacto son requeridos.", "warning");
      return;
    }

    const contactos = getContactos();
    let updatedContactos;

    if (id) {
      // Editar contacto existente
      updatedContactos = contactos.map((c) =>
        c.id === id
          ? { 
              ...c, 
              nombreCliente, 
              nombre, 
              correo,
              updatedAt: new Date().toISOString()
            }
          : c
      );
    } else {
      // Nuevo contacto
      const newContacto = {
        id: `CONTACTO-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        nombreCliente,
        nombre,
        correo,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      updatedContactos = [...contactos, newContacto];
    }

    saveContactos(updatedContactos);
    renderContactos();

    const modalEl = document.getElementById("contactoModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.hide();
    }
  }

  async function deleteContacto(id) {
    const confirmed = await showConfirm("¿Estás seguro de que deseas eliminar este contacto?", "Eliminar");
    if (!confirmed) return;

    const contactos = getContactos().filter((c) => c.id !== id);
    saveContactos(contactos);
    renderContactos();
  }

  function onContactosExcelSelected(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (typeof XLSX === "undefined") {
      showAlert("No se pudo cargar la librería de Excel. Verifica tu conexión.", "error");
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const contactos = buildContactosFromWorkbook(wb);
        
        if (contactos.length === 0) {
          showAlert("No se encontraron contactos válidos en el archivo.", "warning");
          return;
        }

        const existingContactos = getContactos();
        const nuevos = contactos.filter((c) => {
          // Verificar si ya existe un contacto con el mismo nombre de cliente y nombre de contacto
          return !existingContactos.some((ec) => 
            normalizeText(ec.nombreCliente) === normalizeText(c.nombreCliente) &&
            normalizeText(ec.nombre) === normalizeText(c.nombre)
          );
        });

        if (nuevos.length === 0) {
          showAlert("Todos los contactos ya existen en el sistema.", "warning");
          return;
        }

        saveContactos([...existingContactos, ...nuevos]);
        renderContactos();
        showAlert(`Se importaron ${nuevos.length} contacto(s) exitosamente.`, "success");
      } catch (error) {
        console.error("Error al importar contactos:", error);
        showAlert("Error al importar el archivo. Verifica que sea un archivo Excel válido.", "error");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function buildContactosFromWorkbook(wb) {
    const contactos = [];
    
    // Buscar hoja de contactos
    const contactosSheetName = wb.SheetNames.find(name => {
      const normalized = normalizeText(name);
      return normalized.includes("contacto") || 
             normalized === "clientes_contactos" ||
             normalized === "clientescontactos";
    }) || wb.SheetNames[0];
    
    const contactosSheet = wb.Sheets[contactosSheetName];
    if (!contactosSheet) {
      return contactos;
    }
    
    // Leer como JSON con encabezados
    const rows = XLSX.utils.sheet_to_json(contactosSheet, { header: 1, raw: false, defval: "" });
    if (!rows || rows.length === 0) return contactos;

    // Buscar fila de encabezados
    let headerRow = -1;
    let nombreClienteIdx = -1;
    let nombreContactoIdx = -1;
    let correoContactoIdx = -1;

    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = ensureArray(rows[i]).map((c) => normalizeText(String(c || "")));
      const nombreClientePos = row.findIndex((c) => 
        (c.includes("nombre") && !c.includes("contacto")) || 
        c.includes("cliente")
      );
      const nombreContactoPos = row.findIndex((c) => 
        c.includes("nombre") && c.includes("contacto")
      );
      const correoContactoPos = row.findIndex((c) => 
        (c.includes("correo") || c.includes("email") || c.includes("mail")) && 
        c.includes("contacto")
      );

      if (nombreClientePos >= 0 && nombreContactoPos >= 0) {
        headerRow = i;
        nombreClienteIdx = nombreClientePos;
        nombreContactoIdx = nombreContactoPos;
        correoContactoIdx = correoContactoPos >= 0 ? correoContactoPos : -1;
        break;
      }
    }

    if (headerRow < 0 || nombreClienteIdx < 0 || nombreContactoIdx < 0) {
      // Sin encabezados: asumir estructura por defecto
      if (rows.length > 1) {
        nombreClienteIdx = 0;
        nombreContactoIdx = 1;
        correoContactoIdx = 2;
        headerRow = 0;
      } else {
        return contactos;
      }
    }

    // Procesar filas de datos
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = ensureArray(rows[i]);
      const nombreCliente = String(row[nombreClienteIdx] || "").trim();
      const nombreContacto = String(row[nombreContactoIdx] || "").trim();
      
      if (!nombreCliente || !nombreContacto) continue; // Saltar filas sin datos requeridos

      const correoContacto = correoContactoIdx >= 0 ? String(row[correoContactoIdx] || "").trim() : "";

      contactos.push({
        id: `CONTACTO-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 9)}`,
        nombreCliente,
        nombre: nombreContacto,
        correo: correoContacto,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
    }
    
    return contactos;
  }

  async function exportContactosToExcel() {
    if (typeof XLSX === "undefined") {
      showAlert("No se pudo cargar la librería de Excel. Verifica tu conexión.", "error");
      return;
    }

    const contactos = getContactos();
    if (contactos.length === 0) {
      showAlert("No hay contactos para exportar.", "info");
      return;
    }

    // Preparar datos para la hoja CONTACTOS
    const contactosRows = contactos.map((c) => [
      c.nombreCliente || "",
      c.nombre || "",
      c.correo || ""
    ]);

    const headers = ["Nombre Cliente", "Nombre Contacto", "Correo Contacto"];
    const filename = `Contactos_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // Intentar usar ExcelJS con estilos
    const success = await applyHeaderStylesWithExcelJS(
      contactosRows, 
      headers, 
      filename, 
      "CONTACTOS",
      [25, 25, 30]
    );
    if (success) {
      return; // Si ExcelJS funcionó, salir
    }

    // Si ExcelJS no está disponible, usar XLSX sin estilos
    const wb = XLSX.utils.book_new();
    const wsContactos = XLSX.utils.aoa_to_sheet([
      headers,
      ...contactosRows
    ]);
    
    // Ajustar ancho de columnas
    wsContactos['!cols'] = [
      { wch: 25 }, // Nombre (cliente)
      { wch: 25 }, // Nombre contacto
      { wch: 30 }  // Correo contacto
    ];
    
    XLSX.utils.book_append_sheet(wb, wsContactos, "CONTACTOS");
    XLSX.writeFile(wb, filename);
  }

  function openDeleteContactosModal() {
    if (!$deleteContactosModal || !$deleteContactosList) return;

    const contactos = getContactos();
    if (contactos.length === 0) {
      showAlert("No hay contactos para eliminar.", "info");
      return;
    }

    $deleteContactosList.innerHTML = "";

    contactos.forEach((contacto) => {
      const item = document.createElement("div");
      item.className = "list-group-item";
      item.innerHTML = `
        <div class="form-check">
          <input class="form-check-input contacto-delete-checkbox" type="checkbox" value="${contacto.id}" id="contacto-${contacto.id}">
          <label class="form-check-label" for="contacto-${contacto.id}">
            <strong>${escapeHtml(contacto.nombreCliente || "")}</strong> - ${escapeHtml(contacto.nombre || "")} ${contacto.correo ? `(${escapeHtml(contacto.correo)})` : ''}
          </label>
        </div>
      `;
      $deleteContactosList.appendChild(item);
    });

    // Agregar event listeners a los checkboxes
    $deleteContactosList.querySelectorAll(".contacto-delete-checkbox").forEach((checkbox) => {
      checkbox.addEventListener("change", updateDeleteContactosButtonState);
    });

    updateDeleteContactosButtonState();

    if (typeof bootstrap !== "undefined" && $deleteContactosModal) {
      const modal = bootstrap.Modal.getOrCreateInstance($deleteContactosModal);
      modal.show();
    }
  }

  function selectAllContactosToDelete() {
    if (!$deleteContactosList) return;
    $deleteContactosList.querySelectorAll(".contacto-delete-checkbox").forEach((checkbox) => {
      checkbox.checked = true;
    });
    updateDeleteContactosButtonState();
  }

  function deselectAllContactosToDelete() {
    if (!$deleteContactosList) return;
    $deleteContactosList.querySelectorAll(".contacto-delete-checkbox").forEach((checkbox) => {
      checkbox.checked = false;
    });
    updateDeleteContactosButtonState();
  }

  function updateDeleteContactosButtonState() {
    if (!$deleteContactosList || !$btnConfirmDeleteContactos) return;
    const checked = $deleteContactosList.querySelectorAll(".contacto-delete-checkbox:checked");
    const count = checked.length;
    $btnConfirmDeleteContactos.disabled = count === 0;
    if ($deleteContactosCount) {
      $deleteContactosCount.textContent = count;
    }
  }

  async function deleteSelectedContactos() {
    if (!$deleteContactosList) return;

    const checked = $deleteContactosList.querySelectorAll(".contacto-delete-checkbox:checked");
    if (checked.length === 0) {
      showAlert("Por favor selecciona al menos un contacto para eliminar.", "warning");
      return;
    }

    const idsToDelete = Array.from(checked).map((cb) => cb.value);
    const count = idsToDelete.length;
    
    const confirmed = await showConfirm(`¿Estás seguro de que deseas eliminar ${count} contacto(s)?\n\nEsta acción no se puede deshacer.`, "Eliminar");
    if (!confirmed) {
      return;
    }

    const contactos = getContactos().filter((c) => !idsToDelete.includes(c.id));
    saveContactos(contactos);
    renderContactos();

    if (typeof bootstrap !== "undefined" && $deleteContactosModal) {
      const modal = bootstrap.Modal.getOrCreateInstance($deleteContactosModal);
      modal.hide();
    }

    showAlert(`Se eliminaron ${count} contacto(s) exitosamente.`, "success");
  }

  // ==================== GESTIÓN DE USUARIOS ====================

  async function exportUsersToExcel() {
    if (typeof XLSX === "undefined") {
      showAlert("No se pudo cargar la librería de Excel. Verifica tu conexión.", "error");
      return;
    }

    const users = getUsers();
    if (users.length === 0) {
      showAlert("No hay usuarios para exportar.", "info");
      return;
    }

    // Preparar datos para la hoja USUARIOS
    const usersRows = users.map((u) => [
      u.username || "",
      u.name || "",
      u.role === "admin" ? "Administrador" : "Vendedor",
      u.active ? "Activo" : "Inactivo"
    ]);

    const headers = ["Usuario", "Nombre", "Rol", "Estado"];
    const filename = `Usuarios_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // Intentar usar ExcelJS con estilos
    const success = await applyHeaderStylesWithExcelJS(
      usersRows, 
      headers, 
      filename, 
      "Usuarios",
      [20, 30, 15, 12]
    );
    if (success) {
      return; // Si ExcelJS funcionó, salir
    }

    // Si ExcelJS no está disponible, usar XLSX sin estilos
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      headers,
      ...usersRows
    ]);
    
    // Ajustar ancho de columnas
    ws['!cols'] = [
      { wch: 20 }, // Usuario
      { wch: 30 }, // Nombre
      { wch: 15 }, // Rol
      { wch: 12 }  // Estado
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, "Usuarios");
    XLSX.writeFile(wb, filename);
  }

  function renderUsers() {
    if (!$usersTableBody || !$noUsers) return;

    const users = getUsers();
    $usersTableBody.innerHTML = "";

    if (users.length === 0) {
      $noUsers.classList.remove("d-none");
      return;
    }

    $noUsers.classList.add("d-none");

    users.forEach((user, idx) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td>${escapeHtml(user.username || "")}</td>
        <td>${escapeHtml(user.name || "")}</td>
        <td>
          <span class="badge ${user.role === "admin" ? "bg-primary" : "bg-secondary"}">
            ${user.role === "admin" ? "Administrador" : "Vendedor"}
          </span>
        </td>
        <td>
          <span class="badge ${user.active ? "bg-success" : "bg-danger"}">
            ${user.active ? "Activo" : "Inactivo"}
          </span>
        </td>
        <td class="text-nowrap">
          <button class="btn btn-sm btn-outline-primary me-1" data-action="edit" data-id="${user.id}">
            <i class="bi bi-pencil"></i> Editar
          </button>
          <button class="btn btn-sm btn-outline-danger" data-action="delete" data-id="${user.id}" ${user.id === currentUser.id ? "disabled" : ""}>
            <i class="bi bi-trash"></i> Eliminar
          </button>
        </td>
      `;
      $usersTableBody.appendChild(tr);
    });

    $usersTableBody.querySelectorAll("button").forEach((btn) => {
      btn.addEventListener("click", () => {
        const action = btn.dataset.action;
        const id = btn.dataset.id;
        const user = users.find((u) => u.id === id);
        if (action === "edit" && user) {
          openUserModal(user);
        } else if (action === "delete" && user.id !== currentUser.id) {
          deleteUser(id);
        }
      });
    });
  }

  function openUserModal(user = null) {
    if (!$userModalTitle || !$userId || !$userUsername || !$userName || !$userPassword || !$userRole || !$userActive) return;

    if (user) {
      $userModalTitle.textContent = "Editar Usuario";
      $userId.value = user.id;
      $userUsername.value = user.username || "";
      $userName.value = user.name || "";
      $userPassword.value = "";
      $userRole.value = user.role || "";
      $userActive.checked = user.active !== false;
      
      // Ocultar requerimiento de contraseña al editar
      if ($passwordRequired) $passwordRequired.classList.add("d-none");
      if ($passwordHint) $passwordHint.classList.remove("d-none");
      if ($userPassword) $userPassword.removeAttribute("required");
    } else {
      $userModalTitle.textContent = "Nuevo Usuario";
      $userId.value = "";
      $userUsername.value = "";
      $userName.value = "";
      $userPassword.value = "";
      $userRole.value = "";
      $userActive.checked = true;
      
      // Mostrar requerimiento de contraseña al crear
      if ($passwordRequired) $passwordRequired.classList.remove("d-none");
      if ($passwordHint) $passwordHint.classList.add("d-none");
      if ($userPassword) $userPassword.setAttribute("required", "required");
    }

    const modalEl = document.getElementById("userModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.show();
    }
  }

  function updateUserInfoInNavbar() {
    currentUser = getCurrentUser();
    if (currentUser && $userInfo) {
      $userInfo.textContent = `${currentUser.name} (${currentUser.role === "admin" ? "Administrador" : "Vendedor"})`;
      $userInfo.className = `small me-2 user-role-${currentUser.role}`;
    }
    // Actualizar clase de rol al body
    if (currentUser) {
      document.body.className = currentUser.role === "admin" ? "admin" : "";
    }
  }

  function saveUser() {
    if (!$userFormModal || !$userFormModal.checkValidity()) {
      $userFormModal.reportValidity();
      return;
    }

    const username = $userUsername.value.trim();
    const name = $userName.value.trim();
    const password = $userPassword.value.trim();
    const role = $userRole.value;
    const active = $userActive.checked;
    const id = $userId.value.trim();

    if (!username || !name || !role) {
      showAlert("Por favor completa todos los campos requeridos.", "warning");
      return;
    }

    const users = getUsers();
    let updatedUsers;
    let isCurrentUser = false;

    if (id) {
      // Editar usuario existente
      const existingUser = users.find((u) => u.id === id);
      if (!existingUser) {
        showAlert("Usuario no encontrado.", "error");
        return;
      }

      // Verificar si es el usuario actual
      isCurrentUser = currentUser && currentUser.id === id;

      // Verificar que el username no esté en uso por otro usuario
      const usernameTaken = users.some((u) => u.id !== id && u.username.toLowerCase() === username.toLowerCase());
      if (usernameTaken) {
        showAlert("El nombre de usuario ya está en uso.", "warning");
        return;
      }

      updatedUsers = users.map((u) =>
        u.id === id
          ? {
              ...u,
              username,
              name,
              role,
              active,
              password: password ? hashPassword(password) : u.password,
              updatedAt: new Date().toISOString()
            }
          : u
      );
    } else {
      // Crear nuevo usuario
      if (!password) {
        showAlert("La contraseña es requerida para nuevos usuarios.", "warning");
        return;
      }

      // Verificar que el username no esté en uso
      const usernameTaken = users.some((u) => u.username.toLowerCase() === username.toLowerCase());
      if (usernameTaken) {
        showAlert("El nombre de usuario ya está en uso.", "warning");
        return;
      }

      const newUser = {
        id: `USER-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        username,
        password: hashPassword(password),
        name,
        role,
        active,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      updatedUsers = [...users, newUser];
    }

    saveUsers(updatedUsers);
    
    // Si se editó el usuario actual, actualizar la sesión y el navbar
    if (isCurrentUser) {
      const updatedUser = updatedUsers.find((u) => u.id === id);
      if (updatedUser) {
        const { password: _, ...userWithoutPassword } = updatedUser;
        setCurrentUser(userWithoutPassword);
        updateUserInfoInNavbar();
      }
    }
    
    renderUsers();
    closeUserModal();
  }

  function deleteUser(id) {
    if (id === currentUser.id) {
      showAlert("No puedes eliminar tu propio usuario.", "warning");
      return;
    }

    showConfirm("¿Estás seguro de que deseas eliminar este usuario?", "Eliminar").then((ok) => {
      if (ok) {
        const users = getUsers().filter((u) => u.id !== id);
        saveUsers(users);
        renderUsers();
      }
    });
  }

  function closeUserModal() {
    const modalEl = document.getElementById("userModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.hide();
    }
    if ($userFormModal) $userFormModal.reset();
    if ($userId) $userId.value = "";
    if ($passwordRequired) $passwordRequired.classList.remove("d-none");
    if ($passwordHint) $passwordHint.classList.add("d-none");
  }

  // Inicio robusto: si el DOM ya está listo, inicializa de inmediato
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();


