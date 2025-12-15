(function () {
  "use strict";

  // Manejador de errores global para suprimir errores de extensiones de Chrome
  window.addEventListener('error', function(event) {
    // Suprimir el error común de extensiones de Chrome que no afecta la funcionalidad
    const errorMessage = event.message || event.error?.message || '';
    if (errorMessage.includes('message channel closed') || 
        errorMessage.includes('asynchronous response')) {
      event.preventDefault();
      event.stopPropagation();
      return true;
    }
  }, true);

  // Manejador para promesas rechazadas no capturadas
  window.addEventListener('unhandledrejection', function(event) {
    // Suprimir el error común de extensiones de Chrome
    const errorMessage = event.reason?.message || event.reason?.toString() || '';
    if (errorMessage.includes('message channel closed') || 
        errorMessage.includes('asynchronous response')) {
      event.preventDefault();
      event.stopImmediatePropagation();
      return true;
    }
  });

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
  const $filterUser = document.getElementById("filterUser");
  const $filterFrom = document.getElementById("filterFrom");
  const $filterTo = document.getElementById("filterTo");
  const $btnClearFilters = document.getElementById("btnClearFilters");
  const $btnExportXlsx = document.getElementById("btnExportXlsx");
  const $btnLoadExcel = document.getElementById("btnLoadExcel");
  const $excelInput = document.getElementById("excelInput");
  const $btnDeleteProducts = document.getElementById("btnDeleteProducts");
  const $productSearch = document.getElementById("productSearch");
  const $scrollTopBtn = document.getElementById("scrollTopBtn");
  
  // Sidebar
  const $sidebarToggle = document.getElementById("sidebarToggle");
  const $sidebarToggleInSidebar = document.getElementById("sidebarToggleInSidebar");
  const $sidebar = document.getElementById("sidebar");
  const $sidebarOverlay = document.getElementById("sidebarOverlay");
  
  // Referencias a los handlers del sidebar para poder removerlos
  let sidebarToggleHandler = null;
  let sidebarToggleTouchHandler = null;
  let sidebarToggleInSidebarHandler = null;
  let sidebarOverlayHandler = null;
  let sidebarResizeHandler = null;
  
  // Productos - crear/editar
  const $btnAddProduct = document.getElementById("btnAddProduct");
  const $btnEditProduct = document.getElementById("btnEditProduct");
  const $productModal = document.getElementById("productModal");
  const $productModalTitle = document.getElementById("productModalTitle");
  const $productId = document.getElementById("productId");
  const $productNombre = document.getElementById("productNombre");
  const $procesosContainer = document.getElementById("procesosContainer");
  const $btnAddProceso = document.getElementById("btnAddProceso");
  const $btnSaveProduct = document.getElementById("btnSaveProduct");
  const $productFormModal = document.getElementById("productFormModal");
  
  // Modal seleccionar producto para editar
  const $selectProductModal = document.getElementById("selectProductModal");
  const $selectProductList = document.getElementById("selectProductList");
  const $selectProductSearch = document.getElementById("selectProductSearch");
  
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
  const $loginSuccess = document.getElementById("loginSuccess");
  const $loginSubmitBtn = document.getElementById("loginSubmitBtn");
  const $logoutBtn = document.getElementById("logoutBtn");
  const $userInfo = document.getElementById("userInfo");
  
  // Verificación de código
  const $verificationModal = document.getElementById("verificationModal");
  const $verificationForm = document.getElementById("verificationForm");
  const $verificationCode = document.getElementById("verificationCode");
  const $verificationError = document.getElementById("verificationError");
  const $verificationEmail = document.getElementById("verificationEmail");
  const $resendCodeBtn = document.getElementById("resendCodeBtn");
  const $codeExpiryTime = document.getElementById("codeExpiryTime");
  const $closeVerificationModal = document.getElementById("closeVerificationModal");
  
  // Almacenar información de verificación temporal
  let pendingVerification = null;
  let verificationTimer = null;
  
  // Gestión de usuarios
  const $usersTableBody = document.getElementById("usersTableBody");
  const $noUsers = document.getElementById("noUsers");
  const $btnAddUser = document.getElementById("btnAddUser");
  const $btnExportUsers = document.getElementById("btnExportUsers");
  const $userModalTitle = document.getElementById("userModalTitle");
  const $userId = document.getElementById("userId");
  const $userUsername = document.getElementById("userUsername");
  const $userEmail = document.getElementById("userEmail");
  const $userName = document.getElementById("userName");
  const $userCargo = document.getElementById("userCargo");
  const $userPassword = document.getElementById("userPassword");
  const $userPasswordConfirm = document.getElementById("userPasswordConfirm");
  const $userRole = document.getElementById("userRole");
  const $userActive = document.getElementById("userActive");
  const $btnSaveUser = document.getElementById("btnSaveUser");
  const $userFormModal = document.getElementById("userFormModal");
  const $passwordRequired = document.getElementById("passwordRequired");
  const $passwordHint = document.getElementById("passwordHint");
  const $passwordConfirmRequired = document.getElementById("passwordConfirmRequired");
  const $passwordConfirmHint = document.getElementById("passwordConfirmHint");
  
  let logoAsset = null; // { el?: HTMLImageElement, dataUrl?: string }
  let topProductsChart = null;
  let currentUser = null;

  let selectedProductIds = new Set();
  // Map para almacenar el valor unitario seleccionado por producto
  // Clave: productId, Valor: 'vrUnit1' | 'vrUnit2' | 'vrUnit3' | 'vrUnitUSD'
  let selectedUnitValues = new Map();
  // Map para almacenar qué análisis están seleccionados por producto
  // Clave: productId, Valor: Set<index> donde index es el índice del proceso en product.procesos
  let selectedAnalisis = new Map();
  let currentLetter = "A";
  let currentPage = 1;
  const productsPerPage = 10;
  let currentSearchQuery = "";

  // ==================== SISTEMA DE AUTENTICACIÓN ====================

  function initDefaultUsers() {
    try {
      const users = getUsers();
      
      // Verificar si ya existen usuarios (no solo contar, verificar que sean válidos)
      const hasValidUsers = Array.isArray(users) && users.length > 0 && users.some(u => u.id && u.email);
      
      // Solo crear usuarios por defecto si NO hay ningún usuario en el sistema (primera vez)
      // Si ya hay usuarios, no recrear usuarios por defecto aunque hayan sido eliminados
      if (!hasValidUsers) {
        // Solo crear usuarios por defecto si realmente no hay usuarios válidos
        const defaultUsers = [
          {
            id: "USER-ADMIN-001",
            username: "admin",
            email: "admin@olgroup.com",
            password: hashPassword("admin"),
            name: "Administrador",
            role: "admin",
            active: true,
            createdAt: new Date().toISOString()
          },
          {
            id: "USER-VENDEDOR-001",
            username: "vendedor",
            email: "vendedor@olgroup.com",
            password: hashPassword("vendedor"),
            name: "Vendedor",
            role: "vendedor",
            active: true,
            createdAt: new Date().toISOString()
          }
        ];
        
        saveUsers(defaultUsers);
        console.log("Usuarios por defecto creados (primera inicialización)");
      } else {
        // Ya hay usuarios en el sistema, no recrear usuarios por defecto
        // Esto respeta las eliminaciones intencionales de usuarios
        console.log("Usuarios existentes preservados. Total:", users.length);
      }
    } catch (error) {
      console.error("Error al inicializar usuarios por defecto:", error);
      // En caso de error, intentar recuperar datos del localStorage
      try {
        const rawData = localStorage.getItem("olgroup_users");
        if (rawData) {
          console.log("Datos encontrados en localStorage:", rawData.substring(0, 100));
        }
      } catch (e) {
        console.error("Error al leer localStorage:", e);
      }
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
      const usersData = localStorage.getItem("olgroup_users");
      if (!usersData) {
        return [];
      }
      const parsed = JSON.parse(usersData);
      // Validar que sea un array
      if (!Array.isArray(parsed)) {
        console.warn("Los datos de usuarios no son un array válido. Restaurando...");
        return [];
      }
      // Validar que los usuarios tengan estructura básica
      return parsed.filter(u => u && (u.id || u.email || u.username));
    } catch (error) {
      console.error("Error al leer usuarios del localStorage:", error);
      // Intentar recuperar datos corruptos
      try {
        const rawData = localStorage.getItem("olgroup_users");
        console.warn("Datos corruptos encontrados:", rawData);
      } catch (e) {
        console.error("Error al leer datos corruptos:", e);
      }
      return [];
    }
  }

  function saveUsers(users) {
    try {
      // Validar que users sea un array
      if (!Array.isArray(users)) {
        console.error("Error: saveUsers recibió datos que no son un array:", users);
        showAlert("Error al guardar usuarios: datos inválidos.", "error");
        return false;
      }
      
      // Validar que los usuarios tengan estructura básica
      const validUsers = users.filter(u => u && (u.id || u.email || u.username));
      
      if (validUsers.length !== users.length) {
        console.warn("Algunos usuarios fueron filtrados por ser inválidos");
      }
      
      // Guardar en localStorage
      localStorage.setItem("olgroup_users", JSON.stringify(validUsers));
      
      // Verificar que se guardó correctamente
      const verification = localStorage.getItem("olgroup_users");
      if (!verification) {
        throw new Error("No se pudo verificar el guardado");
      }
      
      console.log(`Usuarios guardados correctamente: ${validUsers.length} usuarios`);
      return true;
    } catch (error) {
      console.error("Error al guardar usuarios:", error);
      showAlert("No se pudo guardar los usuarios. Verifica el espacio disponible en el navegador.", "error");
      return false;
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

  // Funciones de verificación de código
  function generateVerificationCode() {
    return Math.floor(100000 + Math.random() * 900000).toString();
  }

  /**
   * Envía un código de verificación por correo electrónico usando EmailJS
   * 
   * CONFIGURACIÓN DE LA PLANTILLA EN EMAILJS:
   * 
   * 1. Ve a tu cuenta de EmailJS: https://dashboard.emailjs.com/
   * 2. Ve a "Email Templates" y edita la plantilla con ID: template_a5mi2mc
   * 3. Configura el ASUNTO del correo:
   *    Código de verificación - Laboratorio de Soluciones Analíticas
   * 
   * 4. Configura el CUERPO del correo con este contenido:
   * 
   *    Hola {{to_name}},
   *    
   *    Tu código de verificación para iniciar sesión es:
   *    
   *    {{verification_code}}
   *    
   *    Este código expira en 5 minutos.
   *    
   *    Si no solicitaste este código, ignora este mensaje.
   *    
   *    Saludos,
   *    {{from_name}}
   * 
   * 5. IMPORTANTE: Asegúrate de usar EXACTAMENTE estos nombres de variables:
   *    - {{to_email}} - Email del destinatario
   *    - {{to_name}} - Nombre del usuario
   *    - {{verification_code}} - Código de 6 dígitos (ESTE ES EL MÁS IMPORTANTE)
   *    - {{from_name}} - Nombre del remitente
   * 
   * 6. Guarda la plantilla
   * 7. Verifica que la plantilla esté activa
   */
  function sendVerificationEmail(email, code, userName) {
    return new Promise((resolve, reject) => {
      // ============================================
      // CONFIGURACIÓN DE EMAILJS
      // ============================================
      const EMAILJS_SERVICE_ID = 'service_tbwaczv';
      const EMAILJS_TEMPLATE_ID = 'template_a5mi2mc';
      const EMAILJS_PUBLIC_KEY = 'ZUesJyPvBcqMQfcP1';
      // ============================================
      
      // Verificar que EmailJS esté disponible
      if (typeof emailjs === 'undefined' || typeof emailjs.send !== 'function') {
        console.error('❌ EmailJS no está cargado o no está disponible.');
        console.warn('⚠️ Código de verificación (solo para desarrollo):', code);
        console.warn('📝 Asegúrate de que el script de EmailJS esté cargado correctamente.');
        console.warn('🔗 Script esperado: https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js');
        
        // Simular envío exitoso después de 1 segundo para permitir desarrollo
        setTimeout(() => {
          resolve({ success: true, code: code });
        }, 1000);
        return;
      }
      
      console.log('✅ EmailJS está disponible y listo para usar');
      
      // Preparar los parámetros del template
      // IMPORTANTE: Estos nombres deben coincidir EXACTAMENTE con las variables en tu plantilla de EmailJS
      const templateParams = {
        to_email: email,
        to_name: userName || 'Usuario',
        verification_code: code,
        from_name: 'Laboratorio de Soluciones Analíticas',
        // Variables alternativas por si la plantilla usa otros nombres
        email: email,
        name: userName || 'Usuario',
        code: code
      };
      
      console.log('📧 Enviando código de verificación a:', email);
      console.log('📋 Parámetros del template:', templateParams);
      console.log('🔑 Service ID:', EMAILJS_SERVICE_ID);
      console.log('📝 Template ID:', EMAILJS_TEMPLATE_ID);
      console.log('🔐 Public Key:', EMAILJS_PUBLIC_KEY);
      
      // Inicializar EmailJS con la Public Key
      try {
        if (typeof emailjs.init === 'function') {
          emailjs.init(EMAILJS_PUBLIC_KEY);
        }
      } catch (initError) {
        console.warn('⚠️ EmailJS init error (puede estar ya inicializado):', initError);
      }
      
      // Enviar email usando la API de EmailJS v4
      // Nota: En v4, la Public Key se puede pasar como 4to parámetro o usar init()
      emailjs.send(EMAILJS_SERVICE_ID, EMAILJS_TEMPLATE_ID, templateParams)
      .then((response) => {
        console.log('✅ Email enviado exitosamente!');
        console.log('📊 Respuesta:', {
          status: response.status,
          text: response.text
        });
        console.log('✉️ Código de verificación enviado a:', email);
        resolve({ success: true });
      })
      .catch((error) => {
        console.error('❌ Error al enviar código de verificación');
        console.error('📊 Detalles del error:', {
          status: error?.status || 'N/A',
          text: error?.text || 'N/A',
          message: error?.message || 'Error desconocido',
          error: error
        });
        
        // Mostrar el código en consola para debugging
        console.warn('⚠️ Código de verificación (para debugging):', code);
        console.warn('⚠️ Checklist de verificación:');
        console.warn('  1. Service ID correcto?', EMAILJS_SERVICE_ID);
        console.warn('  2. Template ID correcto?', EMAILJS_TEMPLATE_ID);
        console.warn('  3. Public Key correcto?', EMAILJS_PUBLIC_KEY);
        console.warn('  4. Plantilla tiene variables? to_email, to_name, verification_code, from_name');
        console.warn('  5. Servicio de email activo en EmailJS?');
        console.warn('  6. Revisa la consola de EmailJS para más detalles');
        
        // Resolver con el código para permitir continuar en caso de error
        resolve({ success: true, code: code, error: error });
      });
    });
  }

  function startVerificationTimer(expiryMinutes = 5) {
    if (verificationTimer) {
      clearInterval(verificationTimer);
    }
    
    let timeLeft = expiryMinutes * 60; // Convertir a segundos
    
    const updateTimer = () => {
      const minutes = Math.floor(timeLeft / 60);
      const seconds = timeLeft % 60;
      
      if ($codeExpiryTime) {
        $codeExpiryTime.textContent = `${minutes}:${seconds.toString().padStart(2, '0')}`;
      }
      
      if (timeLeft <= 0) {
        clearInterval(verificationTimer);
        if (pendingVerification) {
          pendingVerification = null;
          if ($verificationModal) {
            const modal = bootstrap.Modal.getInstance($verificationModal);
            if (modal) modal.hide();
          }
          if ($verificationError) {
            $verificationError.textContent = 'El código de verificación ha expirado. Por favor, inicia sesión nuevamente.';
            $verificationError.classList.remove('d-none');
          }
        }
        return;
      }
      
      timeLeft--;
    };
    
    updateTimer(); // Actualizar inmediatamente
    verificationTimer = setInterval(updateTimer, 1000);
  }

  function showVerificationModal(email, user) {
    if (!$verificationModal) return;
    
    // Generar código de verificación
    const code = generateVerificationCode();
    const expiryTime = new Date(Date.now() + 5 * 60000); // 5 minutos
    
    // Guardar información de verificación
    pendingVerification = {
      code: code,
      email: email,
      user: user,
      expiryTime: expiryTime
    };
    
    // Mostrar email en el modal
    if ($verificationEmail) {
      $verificationEmail.textContent = email;
    }
    
    // Limpiar campos
    if ($verificationCode) {
      $verificationCode.value = '';
    }
    if ($verificationError) {
      $verificationError.classList.add('d-none');
    }
    
    // Enviar código por email
    sendVerificationEmail(email, code, user.name || user.username)
      .then((result) => {
        // Si el código se muestra en consola (desarrollo), mostrarlo también en el modal
        if (result.code && $verificationError) {
          $verificationError.innerHTML = `<strong>Modo desarrollo:</strong> El código es: <code style="font-size: 1.2rem; font-weight: bold;">${result.code}</code>`;
          $verificationError.classList.remove('d-none', 'alert-danger');
          $verificationError.classList.add('alert-info');
        }
      });
    
    // Iniciar temporizador
    startVerificationTimer(5);
    
    // Mostrar modal
    const modal = new bootstrap.Modal($verificationModal, {
      backdrop: 'static',
      keyboard: false
    });
    modal.show();
    
    // Enfocar el campo de código
    setTimeout(() => {
      if ($verificationCode) {
        $verificationCode.focus();
      }
    }, 500);
  }

  function verifyCode(inputCode) {
    if (!pendingVerification) {
      return { success: false, error: 'No hay verificación pendiente.' };
    }
    
    // Verificar expiración
    if (new Date() > pendingVerification.expiryTime) {
      pendingVerification = null;
      return { success: false, error: 'El código de verificación ha expirado.' };
    }
    
    // Verificar código
    if (inputCode === pendingVerification.code) {
      const user = pendingVerification.user;
      pendingVerification = null;
      
      if (verificationTimer) {
        clearInterval(verificationTimer);
        verificationTimer = null;
      }
      
      return { success: true, user: user };
    }
    
    return { success: false, error: 'Código de verificación incorrecto.' };
  }

  function login(email, password) {
    try {
      const users = getUsers();
      
      // Depuración: mostrar información en consola
      console.log("=== INTENTO DE LOGIN ===");
      console.log("Email/Username:", email);
      console.log("Total de usuarios en sistema:", users.length);
      console.log("Usuarios disponibles:", users.map(u => ({ 
        email: u.email, 
        username: u.username, 
        name: u.name,
        active: u.active 
      })));
      
      if (users.length === 0) {
        console.error("ERROR: No hay usuarios en el sistema. Verifica el localStorage.");
        return { 
          success: false, 
          inactive: false, 
          error: "No hay usuarios registrados en el sistema. Por favor, contacta al administrador." 
        };
      }
      
      const hashedPassword = hashPassword(password);
      console.log("Contraseña hasheada:", hashedPassword);
      
      // PRIMERO: Verificar si el usuario existe (solo por email/username, sin verificar contraseña)
      // Buscar por email primero, si no tiene email buscar por username (retrocompatibilidad)
      const user = users.find(
        (u) => {
          const emailMatch = u.email && u.email.toLowerCase() === email.toLowerCase();
          const usernameMatch = !u.email && u.username && u.username.toLowerCase() === email.toLowerCase();
          return emailMatch || usernameMatch;
        }
      );

      // Si el usuario NO existe
      if (!user) {
        console.warn("Usuario no encontrado:", email);
        return { 
          success: false, 
          inactive: false,
          userNotFound: true,
          error: "El correo electrónico no está registrado en el sistema."
        };
      }

      console.log("Usuario encontrado:", user.email || user.username);
      
      // SEGUNDO: Verificar la contraseña del usuario encontrado
      const passwordMatch = user.password === hashedPassword;
      
      console.log("Verificación de contraseña:", {
        storedPassword: user.password,
        inputPassword: hashedPassword,
        passwordMatch: passwordMatch
      });
      
      // Si la contraseña es incorrecta
      if (!passwordMatch) {
        console.warn("Contraseña incorrecta para usuario:", user.email || user.username);
        return { 
          success: false, 
          inactive: false,
          wrongPassword: true,
          error: "La contraseña es incorrecta."
        };
      }
      
      // Si el usuario existe pero está inactivo
      if (user.active === false) {
        console.warn("Usuario inactivo");
        return { success: false, inactive: true };
      }
      
      // Si el usuario existe, la contraseña es correcta y está activo, mostrar modal de verificación
      const userEmail = user.email || email;
      showVerificationModal(userEmail, user);
      
      return { success: true, requiresVerification: true };
    } catch (error) {
      console.error("Error en función login:", error);
      return { 
        success: false, 
        inactive: false, 
        error: "Error al procesar el login. Por favor, recarga la página." 
      };
    }
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

    // Usar requestAnimationFrame para asegurar que el DOM esté completamente renderizado
    // antes de inicializar la aplicación, especialmente para el sidebar
    requestAnimationFrame(() => {
      // Usar un pequeño delay adicional para asegurar que los elementos estén completamente disponibles
      setTimeout(() => {
        initApp();
      }, 50);
    });
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
            $loginError.textContent = "Por favor ingresa email y contraseña.";
            $loginError.classList.remove("d-none");
          }
          if ($loginSuccess) {
            $loginSuccess.classList.add("d-none");
          }
          return;
        }

        // Ocultar mensajes previos
        if ($loginError) {
          $loginError.classList.add("d-none");
        }
        if ($loginSuccess) {
          $loginSuccess.classList.add("d-none");
        }

        // Deshabilitar botón mientras se procesa
        if ($loginSubmitBtn) {
          $loginSubmitBtn.disabled = true;
          $loginSubmitBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Verificando...';
        }

        const loginResult = login(username, password);
        
        // Rehabilitar botón
        if ($loginSubmitBtn) {
          $loginSubmitBtn.disabled = false;
          $loginSubmitBtn.innerHTML = '<i class="bi bi-box-arrow-in-right"></i> Iniciar Sesión';
        }

        if (loginResult.success) {
          if (loginResult.requiresVerification) {
            // Mostrar mensaje de éxito
            if ($loginSuccess) {
              $loginSuccess.classList.remove("d-none");
            }
            // El modal de verificación se mostrará automáticamente
          } else {
            // Si no requiere verificación (no debería pasar con el nuevo flujo)
            showMainContent();
          }
        } else {
          if ($loginError) {
            if (loginResult.inactive) {
              $loginError.textContent = "Usuario inactivo por administrador.";
            } else if (loginResult.error) {
              // Mostrar el mensaje de error específico (correo no registrado o contraseña incorrecta)
              $loginError.textContent = loginResult.error;
            } else {
              // Mensaje genérico solo si no hay un error específico
              $loginError.textContent = "Error al iniciar sesión. Por favor, intenta nuevamente.";
            }
            $loginError.classList.remove("d-none");
          }
        }
      });
    }

    // Event listeners para verificación de código
    if ($verificationForm) {
      $verificationForm.addEventListener("submit", (e) => {
        e.preventDefault();
        const code = $verificationCode.value.trim().replace(/\s/g, '');

        if (!code || code.length !== 6) {
          if ($verificationError) {
            $verificationError.textContent = "Por favor ingresa un código de 6 dígitos.";
            $verificationError.classList.remove("d-none");
            $verificationError.classList.remove("alert-info");
            $verificationError.classList.add("alert-danger");
          }
          return;
        }

        const verificationResult = verifyCode(code);

        if (verificationResult.success) {
          // Cerrar modal
          const modal = bootstrap.Modal.getInstance($verificationModal);
          if (modal) modal.hide();

          // Establecer usuario actual
          const { password: _, ...userWithoutPassword } = verificationResult.user;
          setCurrentUser(userWithoutPassword);

          // Limpiar verificación pendiente
          pendingVerification = null;
          if (verificationTimer) {
            clearInterval(verificationTimer);
            verificationTimer = null;
          }

          // Mostrar contenido principal
          showMainContent();

          // Limpiar formulario de login
          if ($loginForm) {
            $loginForm.reset();
          }
          if ($loginError) {
            $loginError.classList.add("d-none");
          }
          if ($loginSuccess) {
            $loginSuccess.classList.add("d-none");
          }
        } else {
          if ($verificationError) {
            $verificationError.textContent = verificationResult.error || "Código incorrecto. Por favor, intenta nuevamente.";
            $verificationError.classList.remove("d-none", "alert-info");
            $verificationError.classList.add("alert-danger");
          }
          // Limpiar campo de código
          if ($verificationCode) {
            $verificationCode.value = '';
            $verificationCode.focus();
          }
        }
      });
    }

    // Reenviar código
    if ($resendCodeBtn) {
      $resendCodeBtn.addEventListener("click", (e) => {
        e.preventDefault();
        if (!pendingVerification) {
          return;
        }

        // Generar nuevo código
        const newCode = generateVerificationCode();
        const expiryTime = new Date(Date.now() + 5 * 60000);

        pendingVerification.code = newCode;
        pendingVerification.expiryTime = expiryTime;

        // Limpiar mensajes
        if ($verificationError) {
          $verificationError.classList.add("d-none");
        }
        if ($verificationCode) {
          $verificationCode.value = '';
        }

        // Enviar nuevo código
        sendVerificationEmail(pendingVerification.email, newCode, pendingVerification.user.name || pendingVerification.user.username)
          .then((result) => {
            if (result.code && $verificationError) {
              $verificationError.innerHTML = `<strong>Modo desarrollo:</strong> El nuevo código es: <code style="font-size: 1.2rem; font-weight: bold;">${result.code}</code>`;
              $verificationError.classList.remove("d-none", "alert-danger");
              $verificationError.classList.add("alert-info");
            }
          });

        // Reiniciar temporizador
        startVerificationTimer(5);
      });
    }

    // Cerrar modal de verificación (solo permitir si se cancela explícitamente)
    if ($closeVerificationModal) {
      $closeVerificationModal.addEventListener("click", () => {
        if (confirm("¿Estás seguro de que deseas cancelar la verificación? Deberás iniciar sesión nuevamente.")) {
          pendingVerification = null;
          if (verificationTimer) {
            clearInterval(verificationTimer);
            verificationTimer = null;
          }
          if ($loginForm) {
            $loginForm.reset();
          }
          if ($loginSuccess) {
            $loginSuccess.classList.add("d-none");
          }
        }
      });
    }

    // Limpiar estado cuando se cierra el modal de verificación
    if ($verificationModal) {
      $verificationModal.addEventListener('hidden.bs.modal', () => {
        // Solo limpiar si no se completó la verificación exitosamente
        if (pendingVerification) {
          pendingVerification = null;
          if (verificationTimer) {
            clearInterval(verificationTimer);
            verificationTimer = null;
          }
          if ($verificationCode) {
            $verificationCode.value = '';
          }
          if ($verificationError) {
            $verificationError.classList.add('d-none');
          }
        }
      });
    }

    // Permitir solo números en el campo de código
    if ($verificationCode) {
      $verificationCode.addEventListener("input", (e) => {
        e.target.value = e.target.value.replace(/\D/g, '').slice(0, 6);
      });

      $verificationCode.addEventListener("paste", (e) => {
        e.preventDefault();
        const paste = (e.clipboardData || window.clipboardData).getData("text");
        const numbers = paste.replace(/\D/g, '').slice(0, 6);
        e.target.value = numbers;
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
    if ($filterUser) {
      $filterUser.addEventListener("input", applyFiltersDebounced);
    }
    if ($filterFrom) {
      $filterFrom.addEventListener("change", () => renderHistory(getHistoryFilters()));
    }
    if ($filterTo) {
      $filterTo.addEventListener("change", () => renderHistory(getHistoryFilters()));
    }


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
          if ($userPasswordConfirm) $userPasswordConfirm.value = "";
          if ($passwordRequired) $passwordRequired.classList.remove("d-none");
          if ($passwordHint) $passwordHint.classList.remove("d-none");
          if ($passwordHint) $passwordHint.textContent = "Mínimo 8 caracteres. Dejar en blanco para mantener la actual (solo al editar)";
          if ($passwordConfirmRequired) $passwordConfirmRequired.classList.remove("d-none");
          if ($passwordConfirmHint) $passwordConfirmHint.textContent = "Repite la contraseña para verificar";
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

    // Event listeners para productos - crear/editar
    if ($btnAddProduct) {
      $btnAddProduct.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (!currentUser || currentUser.role !== "admin") {
          showAlert("Solo los administradores pueden crear productos.", "error", "Acceso denegado");
          return;
        }
        openProductModal();
      });
    } else {
      console.warn("Botón btnAddProduct no encontrado");
    }
    if ($btnEditProduct) {
      $btnEditProduct.addEventListener("click", () => {
        if (!currentUser || currentUser.role !== "admin") {
          showAlert("Solo los administradores pueden editar productos.", "error", "Acceso denegado");
          return;
        }
        openProductModal(true);
      });
    }
    
    // Event listener para el buscador del modal de selección de productos
    if ($selectProductSearch) {
      $selectProductSearch.addEventListener("input", () => {
        renderSelectProductList();
      });
    }
    if ($btnAddProceso) {
      $btnAddProceso.addEventListener("click", addProcesoRow);
    }
    if ($btnSaveProduct) {
      $btnSaveProduct.addEventListener("click", saveProduct);
    }
    
    // Resetear formulario de producto al cerrar el modal
    if ($productModal && typeof bootstrap !== "undefined") {
      $productModal.addEventListener("hidden.bs.modal", () => {
        if ($productFormModal) $productFormModal.reset();
        if ($productId) $productId.value = "";
        if ($procesosContainer) $procesosContainer.innerHTML = "";
        if ($productModalTitle) $productModalTitle.textContent = "Nuevo Producto";
      });
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

    // Sidebar toggle
    function toggleSidebar(e) {
      e.preventDefault();
      e.stopPropagation();
      if ($sidebar) {
        const isActive = $sidebar.classList.contains("active");
        if (isActive) {
          closeSidebar();
        } else {
          openSidebar();
        }
      }
    }

    function openSidebar() {
      if ($sidebar) {
        $sidebar.classList.add("active");
        if ($sidebarOverlay) {
          $sidebarOverlay.classList.add("active");
        }
        // Agregar clase al body para ocultar el top-navbar
        document.body.classList.add("sidebar-open");
        // Actualizar icono del botón hamburguesa
        updateSidebarToggleIcon(true);
        // Guardar estado en localStorage
        localStorage.setItem("sidebarOpen", "true");
      }
    }

    function closeSidebar() {
      if ($sidebar) {
        $sidebar.classList.remove("active");
        if ($sidebarOverlay) {
          $sidebarOverlay.classList.remove("active");
        }
        // Remover clase del body para mostrar el top-navbar
        document.body.classList.remove("sidebar-open");
        // Actualizar icono del botón hamburguesa
        updateSidebarToggleIcon(false);
        // Guardar estado en localStorage
        localStorage.setItem("sidebarOpen", "false");
      }
    }

    function updateSidebarToggleIcon(isOpen) {
      if ($sidebarToggle) {
        const icon = $sidebarToggle.querySelector("i");
        if (icon) {
          icon.className = isOpen ? "bi bi-x-lg fs-4" : "bi bi-list fs-4";
        }
      }
      // El botón en el sidebar siempre muestra X porque solo está visible cuando está abierto
      if ($sidebarToggleInSidebar) {
        const icon = $sidebarToggleInSidebar.querySelector("i");
        if (icon) {
          icon.className = "bi bi-x-lg fs-4";
        }
      }
    }

    // Event listeners - Remover listeners previos si existen antes de agregar nuevos
    if ($sidebarToggle) {
      // Remover listeners previos si existen
      if (sidebarToggleHandler) {
        $sidebarToggle.removeEventListener("click", sidebarToggleHandler);
      }
      if (sidebarToggleTouchHandler) {
        $sidebarToggle.removeEventListener("touchend", sidebarToggleTouchHandler);
      }
      
      // Forzar un reflow del DOM para asegurar que los estilos estén aplicados
      void $sidebarToggle.offsetHeight;
      
      // Asegurar que el botón sea clickeable
      $sidebarToggle.style.pointerEvents = "auto";
      $sidebarToggle.style.zIndex = "1053";
      $sidebarToggle.style.cursor = "pointer";
      $sidebarToggle.style.position = "relative";
      
      // Crear nuevos handlers y guardar referencias
      sidebarToggleHandler = function(e) {
        e.preventDefault();
        e.stopPropagation();
        toggleSidebar(e);
      };
      sidebarToggleTouchHandler = function(e) {
        e.preventDefault();
        e.stopPropagation();
        toggleSidebar(e);
      };
      
      // Agregar nuevos listeners con { passive: false } para poder prevenir el comportamiento por defecto
      $sidebarToggle.addEventListener("click", sidebarToggleHandler, { passive: false });
      $sidebarToggle.addEventListener("touchend", sidebarToggleTouchHandler, { passive: false });
    }

    if ($sidebarToggleInSidebar) {
      // Remover listener previo si existe
      if (sidebarToggleInSidebarHandler) {
        $sidebarToggleInSidebar.removeEventListener("click", sidebarToggleInSidebarHandler);
      }
      
      // Crear nuevo handler y guardar referencia
      sidebarToggleInSidebarHandler = function(e) {
        e.preventDefault();
        e.stopPropagation();
        closeSidebar();
      };
      
      // Agregar nuevo listener
      $sidebarToggleInSidebar.addEventListener("click", sidebarToggleInSidebarHandler);
    }

    if ($sidebarOverlay) {
      // Remover listener previo si existe
      if (sidebarOverlayHandler) {
        $sidebarOverlay.removeEventListener("click", sidebarOverlayHandler);
      }
      
      // Crear nuevo handler y guardar referencia
      sidebarOverlayHandler = function(e) {
        e.preventDefault();
        e.stopPropagation();
        closeSidebar();
      };
      
      // Agregar nuevo listener
      $sidebarOverlay.addEventListener("click", sidebarOverlayHandler);
    }

    // Cerrar sidebar al hacer clic en un enlace de navegación en móvil
    const sidebarLinks = document.querySelectorAll(".sidebar-nav .nav-link");
    sidebarLinks.forEach(link => {
      link.addEventListener("click", function() {
        if (window.innerWidth < 992) {
          closeSidebar();
        }
      });
    });

    // Restaurar estado del sidebar desde localStorage (solo en desktop)
    if ($sidebar && window.innerWidth >= 992) {
      const sidebarOpen = localStorage.getItem("sidebarOpen");
      if (sidebarOpen === "true") {
        openSidebar();
      }
    }

    // Cerrar sidebar al cambiar de tamaño de ventana si está en móvil
    // Remover handler previo si existe
    if (sidebarResizeHandler) {
      window.removeEventListener("resize", sidebarResizeHandler);
    }
    
    let resizeTimeout;
    sidebarResizeHandler = () => {
      clearTimeout(resizeTimeout);
      resizeTimeout = setTimeout(() => {
        if (window.innerWidth < 992 && $sidebar) {
          closeSidebar();
        }
      }, 250);
    };
    
    window.addEventListener("resize", sidebarResizeHandler);
  }

  function onExcelSelected(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const newProducts = buildProductsFromWorkbook(wb);
        
        if (newProducts.length === 0) {
          showAlert("No se encontraron productos válidos en el archivo Excel.", "warning");
          return;
        }

        // Confirmar si se deben agregar o reemplazar
        const existingProducts = productos || [];
        
        if (existingProducts.length > 0) {
          const shouldAppend = await showConfirm(
            `Ya existen ${existingProducts.length} producto(s) registrados.\n\n` +
            `¿Deseas AGREGAR los ${newProducts.length} nuevo(s) producto(s) a los existentes?\n\n` +
            `- Aceptar: Agregar nuevos (se evitarán duplicados por nombre)\n` +
            `- Cancelar: Reemplazar todos los productos existentes`,
            "Agregar"
          );

          if (shouldAppend) {
            // Agregar nuevos productos (evitando duplicados por nombre)
            const nombres = new Set(existingProducts.map((p) => p.nombre.toLowerCase()));
            const nuevos = newProducts.filter((p) => !nombres.has(p.nombre.toLowerCase()));
            const duplicados = newProducts.length - nuevos.length;
            if (duplicados > 0) {
              showAlert(`${duplicados} producto(s) ya existen y no fueron agregados.`, "warning");
            }
            productos = [...existingProducts, ...nuevos];
            saveCatalog(productos);
            selectedProductIds.clear();
            updateSelectionStateUI();
            renderProductsByLetter(currentLetter);
            showAlert(`Se agregaron ${nuevos.length} producto(s) exitosamente.`, "success");
          } else {
            // Reemplazar todos
            const confirmed = await showConfirm(`¿Estás seguro de que deseas reemplazar todos los ${existingProducts.length} producto(s) existentes con los ${newProducts.length} nuevo(s)?`, "Reemplazar");
            if (confirmed) {
              productos = newProducts;
              saveCatalog(productos);
              selectedProductIds.clear();
              updateSelectionStateUI();
              renderProductsByLetter(currentLetter);
              showAlert(`Se importaron ${newProducts.length} producto(s) exitosamente.`, "success");
            }
          }
        } else {
          // No hay productos existentes, importar directamente
          productos = newProducts;
          saveCatalog(productos);
          selectedProductIds.clear();
          updateSelectionStateUI();
          renderProductsByLetter(currentLetter);
          showAlert(`Se importaron ${newProducts.length} producto(s) exitosamente.`, "success");
        }
      } catch (err) {
        showAlert("No se pudo leer el archivo Excel. Verifica el formato.", "error");
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
    if ($filterUser) $filterUser.value = "";
    if ($filterFrom) $filterFrom.value = "";
    if ($filterTo) $filterTo.value = "";
    renderHistory();
    
    // Quitar el focus del botón después de limpiar
    if ($btnClearFilters) {
      $btnClearFilters.blur();
    }
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
        <select class="form-select form-select-sm unit-selector" data-product-id="${product.id}" style="width: auto; min-width: 170px;">
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

    // Sincronizar análisis con el estado del checkbox del producto al renderizar
    const isProductSelected = selectedProductIds.has(product.id);
    if (isProductSelected) {
      let analisisSet = selectedAnalisis.get(product.id);
      if (!analisisSet || analisisSet.size === 0) {
        analisisSet = new Set();
        product.procesos.forEach((_, index) => analisisSet.add(index));
        selectedAnalisis.set(product.id, analisisSet);
      }
      // Actualizar checkboxes de análisis después de renderizar
      setTimeout(() => {
        updateAnalisisCheckboxes(product.id, product);
        updateProductTableTotals(product.id, product);
      }, 0);
    }

    const checkbox = header.querySelector(".product-select");
    checkbox.checked = selectedProductIds.has(product.id);
    checkbox.addEventListener("change", (e) => {
      if (e.target.checked) {
        selectedProductIds.add(product.id);
        // Seleccionar todos los análisis cuando se selecciona el producto
        let analisisSet = selectedAnalisis.get(product.id);
        if (!analisisSet) {
          analisisSet = new Set();
          selectedAnalisis.set(product.id, analisisSet);
        }
        // Seleccionar todos los análisis
        product.procesos.forEach((_, index) => analisisSet.add(index));
        // Actualizar los checkboxes de análisis
        updateAnalisisCheckboxes(product.id, product);
        // Recalcular totales
        updateProductTableTotals(product.id, product);
      } else {
        selectedProductIds.delete(product.id);
        // Deseleccionar todos los análisis cuando se deselecciona el producto
        let analisisSet = selectedAnalisis.get(product.id);
        if (analisisSet) {
          analisisSet.clear();
        }
        // Actualizar los checkboxes de análisis
        updateAnalisisCheckboxes(product.id, product);
        // Recalcular totales
        updateProductTableTotals(product.id, product);
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
    let sumCMtra = 0;
    
    // Obtener los análisis seleccionados
    const analisisSet = selectedAnalisis.get(productId) || new Set();
    
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
              const isSelected = analisisSet.has(index);
              const qty = parseNumStrict(row.cantidad);
              const unitValue = parseNumStrict(row[selectedUnit]);
              let calculatedTotal = 0;
              
              // Actualizar estilo de la fila
              if (isSelected) {
                tr.classList.remove("table-secondary", "opacity-75");
              } else {
                tr.classList.add("table-secondary", "opacity-75");
              }
              
              // Solo calcular si el análisis está seleccionado y el valor unitario tiene un valor válido
              if (isSelected && unitValue > 0 && qty > 0) {
                calculatedTotal = unitValue * qty;
                sumTotal += calculatedTotal;
                
                // Sumar C. Mtra. solo si está seleccionado
                const cMtra = parseNumStrict(row.cMtra ?? row.cMtra_g);
                if (isFinite(cMtra)) {
                  sumCMtra += cMtra;
                }
              }
              
              // Actualizar la celda de Vr. Total (última columna)
              const totalCell = tr.querySelector('td:last-child');
              if (totalCell) {
                totalCell.textContent = calculatedTotal > 0 ? formatMoney(calculatedTotal) : "";
              }
            }
          });
        }
        
        // Actualizar el subtotal y totales en el footer de la tabla
        const tfoot = card.querySelector('tfoot');
        if (tfoot) {
          const cells = tfoot.querySelectorAll('td');
          if (cells.length >= 11) {
            // Actualizar C. Mtra. total (columna 5, índice 4, después de la columna de checkbox)
            cells[4].textContent = formatNumber(sumCMtra);
            // Actualizar subtotal (última columna)
            cells[cells.length - 1].textContent = formatMoney(sumTotal);
          }
        }
      }
    }
    
    // Guardar totales calculados en el objeto para uso en PDF
    product._sumCMtra = sumCMtra;
    product._sumTotal = sumTotal;
  }

  function renderProductTable(product) {
    const tableWrapper = document.createElement("div");
    tableWrapper.className = "table-responsive";
    const table = document.createElement("table");
    table.className = "table table-sm table-striped table-hover mb-0";

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th style="width: 40px;"></th>
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
    
    // Obtener los análisis seleccionados para este producto
    // Solo se seleccionan cuando el checkbox del producto está marcado
    let selectedAnalisisSet = selectedAnalisis.get(product.id);
    if (!selectedAnalisisSet) {
      // Inicializar vacío - solo se seleccionarán cuando se marque el checkbox del producto
      selectedAnalisisSet = new Set();
      selectedAnalisis.set(product.id, selectedAnalisisSet);
    }
    
    // Verificar si el producto está seleccionado para sincronizar los análisis
    const isProductSelected = selectedProductIds.has(product.id);
    if (isProductSelected && selectedAnalisisSet.size === 0) {
      // Si el producto está seleccionado pero no hay análisis seleccionados, seleccionar todos
      product.procesos.forEach((_, index) => selectedAnalisisSet.add(index));
    } else if (!isProductSelected) {
      // Si el producto no está seleccionado, asegurar que no haya análisis seleccionados
      selectedAnalisisSet.clear();
    }
    
    product.procesos.forEach((row, index) => {
      const getText = (v) => (v == null ? "" : String(v));
      const valueOrFormat = (text, num, isMoney = false) => {
        if (text !== undefined && text !== null && String(text) !== "") return String(text);
        if (num === undefined || num === null || String(num) === "") return "";
        return isMoney ? formatMoney(num) : formatNumber(num);
      };

      const isSelected = selectedAnalisisSet.has(index);
      
      // Sumas usando el valor unitario seleccionado (solo si el análisis está seleccionado)
      if (isSelected) {
        const cMtraForSum = parseNumStrict(row.cMtra ?? row.cMtra_g);
        const qty = parseNumStrict(row.cantidad);
        const unitValue = parseNumStrict(row[selectedUnit]);
        
        sumCMtra += cMtraForSum;
        
        // Calcular el total usando el valor unitario seleccionado
        // Solo sumar si el valor unitario tiene un valor válido
        if (unitValue > 0 && qty > 0) {
          sumTotal += unitValue * qty;
        }
      }

      // Calcular el Vr. Total usando el valor unitario seleccionado
      const qtyForDisplay = parseNumStrict(row.cantidad);
      const unitValueForDisplay = parseNumStrict(row[selectedUnit]);
      let calculatedTotal = 0;
      
      // Solo calcular si el análisis está seleccionado y el valor unitario seleccionado tiene un valor válido
      if (isSelected && unitValueForDisplay > 0 && qtyForDisplay > 0) {
        calculatedTotal = unitValueForDisplay * qtyForDisplay;
      }
      
      const tr = document.createElement("tr");
      tr.className = isSelected ? "" : "table-secondary opacity-75";
      tr.innerHTML = `
        <td class="text-center">
          <input type="checkbox" class="form-check-input analisis-checkbox" 
                 data-product-id="${product.id}" 
                 data-analisis-index="${index}" 
                 ${isSelected ? 'checked' : ''}>
        </td>
        <td>${getText(row.codigo)}</td>
        <td>${getText(row.analisis)}</td>
        <td>${getText(row.metodo)}</td>
        <td class=\"text-center\">${valueOrFormat(row.cMtra, row.cMtra_g)}</td>
        <td class=\"text-center\">${valueOrFormat(row.cantidad, row.cantidad)}</td>
        <td class=\"text-center\" style=\"white-space: nowrap;\">${row.vrUnit1 != null && row.vrUnit1 !== "" ? formatMoney(parseNumStrict(row.vrUnit1)) : ""}</td>
        <td class=\"text-center\" style=\"white-space: nowrap;\">${row.vrUnit2 != null && row.vrUnit2 !== "" ? formatMoney(parseNumStrict(row.vrUnit2)) : ""}</td>
        <td class=\"text-center\" style=\"white-space: nowrap;\">${row.vrUnit3 != null && row.vrUnit3 !== "" ? formatMoney(parseNumStrict(row.vrUnit3)) : ""}</td>
        <td class=\"text-center\" style=\"white-space: nowrap;\">${row.vrUnitUSD != null && row.vrUnitUSD !== "" ? formatMoney(parseNumStrict(row.vrUnitUSD)) : ""}</td>
        <td class=\"text-center\" style=\"white-space: nowrap;\">${calculatedTotal > 0 ? formatMoney(calculatedTotal) : ""}</td>
      `;
      
      // Agregar event listener al checkbox
      const checkbox = tr.querySelector(".analisis-checkbox");
      checkbox.addEventListener("change", (e) => {
        const productId = e.target.getAttribute("data-product-id");
        const analisisIndex = parseInt(e.target.getAttribute("data-analisis-index"));
        const productData = productos.find(p => p.id === productId);
        
        if (!productData) return;
        
        let analisisSet = selectedAnalisis.get(productId);
        if (!analisisSet) {
          analisisSet = new Set();
          selectedAnalisis.set(productId, analisisSet);
        }
        
        if (e.target.checked) {
          analisisSet.add(analisisIndex);
        } else {
          analisisSet.delete(analisisIndex);
        }
        
        // Sincronizar checkbox del producto: si todos los análisis están seleccionados, marcar el producto
        // Si ningún análisis está seleccionado, desmarcar el producto
        const productCheckbox = document.querySelector(`.product-select[data-product-id="${productId}"]`);
        if (productCheckbox) {
          if (analisisSet.size === productData.procesos.length) {
            // Todos seleccionados: marcar el checkbox del producto
            if (!selectedProductIds.has(productId)) {
              selectedProductIds.add(productId);
              productCheckbox.checked = true;
            }
          } else if (analisisSet.size === 0) {
            // Ninguno seleccionado: desmarcar el checkbox del producto
            if (selectedProductIds.has(productId)) {
              selectedProductIds.delete(productId);
              productCheckbox.checked = false;
            }
          }
        }
        
        // Actualizar UI de selección
        updateSelectionStateUI();
        
        // Recalcular totales y actualizar la tabla
        updateProductTableTotals(productId, productData);
      });
      
      tbody.appendChild(tr);
    });

    const tfoot = document.createElement("tfoot");
    tfoot.innerHTML = `
      <tr>
        <td></td>
        <td colspan="3" class="text-end">Totales:</td>
        <td class="text-center">${formatNumber(sumCMtra)}</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="text-center fw-semibold">Subtotal</td>
        <td class="text-center fw-semibold" style="white-space: nowrap;">${formatMoney(sumTotal)}</td>
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
  
  function updateAnalisisCheckboxes(productId, product) {
    const analisisSet = selectedAnalisis.get(productId) || new Set();
    // Buscar el card del producto usando el checkbox del producto
    const productCheckbox = document.querySelector(`.product-select[data-product-id="${productId}"]`);
    if (!productCheckbox) return;
    
    const card = productCheckbox.closest('.product-card');
    if (!card) return;
    
    const tbody = card.querySelector('tbody');
    if (!tbody) return;
    
    const checkboxes = tbody.querySelectorAll(`.analisis-checkbox[data-product-id="${productId}"]`);
    checkboxes.forEach((cb) => {
      const analisisIndex = parseInt(cb.getAttribute("data-analisis-index"));
      const isSelected = analisisSet.has(analisisIndex);
      cb.checked = isSelected;
      
      // Actualizar estilo de la fila
      const row = cb.closest("tr");
      if (row) {
        if (isSelected) {
          row.classList.remove("table-secondary", "opacity-75");
        } else {
          row.classList.add("table-secondary", "opacity-75");
        }
      }
    });
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
      totalCOP: totalGeneral,
      userName: currentUser ? currentUser.name || currentUser.username : "",
      userUsername: currentUser ? currentUser.username : ""
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
    
    // Obtener los análisis seleccionados (o todos si no hay selección)
    const analisisSet = selectedAnalisis.get(product.id);
    if (!analisisSet || analisisSet.size === 0) {
      // Si no hay análisis seleccionados, no incluir nada
      return { ...product, _sumCMtra: 0, _sumTotal: 0, _selectedUnit: selectedUnit };
    }
    
    product.procesos.forEach((row, index) => {
      // Solo procesar si el análisis está seleccionado
      if (!analisisSet.has(index)) return;
      
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
    // Limpiar todas las selecciones de productos
    selectedProductIds.clear();
    
    // Limpiar todas las selecciones de análisis
    selectedAnalisis.clear();
    
    // No limpiar selectedUnitValues para mantener las selecciones del usuario
    // selectedUnitValues.clear();
    
    // Actualizar todos los checkboxes de productos
    document.querySelectorAll(".product-select").forEach((cb) => {
      cb.checked = false;
    });
    
    // Actualizar todos los checkboxes de análisis y sus filas
    document.querySelectorAll(".analisis-checkbox").forEach((cb) => {
      cb.checked = false;
      const row = cb.closest("tr");
      if (row) {
        row.classList.add("table-secondary", "opacity-75");
      }
    });
    
    // Recalcular y actualizar totales de todas las tablas
    productos.forEach((product) => {
      updateProductTableTotals(product.id, product);
    });
    
    // Actualizar UI de selección
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
    
    // Función auxiliar para formatear dinero en PDF sin espacio (evita saltos de línea)
    const formatMoneyPDF = (n) => {
      const num = Number(n) || 0;
      if (num === 0) return "";
      // Formato: $1,234,567 (sin espacio para evitar saltos de línea en PDF)
      return "$" + num.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    };
    
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

      // Obtener los análisis seleccionados para este producto
      const analisisSet = selectedAnalisis.get(p.id);
      
      // Filtrar solo los procesos seleccionados
      const procesosSeleccionados = p.procesos.filter((_, index) => {
        // Si no hay selección, incluir todos (comportamiento por defecto para compatibilidad)
        if (!analisisSet || analisisSet.size === 0) {
          return true;
        }
        return analisisSet.has(index);
      });
      
      const tableData = procesosSeleccionados.map((row) => {
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
        
        // Formatear el valor unitario con formato de moneda para PDF
        const unitValueFormatted = unitValue > 0 ? formatMoneyPDF(unitValue) : "";
        
        return [
          String(row.codigo ?? ""),
          String(row.analisis ?? ""),
          String(row.metodo ?? ""),
          String((row.cMtra ?? row.cMtra_g) ?? ""),
          String(row.cantidad ?? ""),
          unitValueFormatted,
          String(finalTotal > 0 ? formatMoneyPDF(finalTotal) : "")
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
          6: { halign: "right", cellWidth: 90 }
        },
        didDrawPage: (data) => {},
        willDrawCell: (data) => {}
      });
      y = doc.lastAutoTable.finalY + 16;

      // Subtotal por producto (alineado a la derecha bajo la tabla)
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...brandSecondary);
      doc.text(`Subtotal: ${formatMoneyPDF(p._sumTotal)}`, right, y, { align: "right" });
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
    doc.text(`Total general (COP): ${formatMoneyPDF(totalGeneral)}`, left, y);
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

    // Firma - Usar información del usuario actual
    y = checkPageBreak(y, 30);
    const user = getCurrentUser();
    const userName = user?.name || "";
    const userCargo = user?.cargo || "";
    
    if (userName) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(userName.toUpperCase(), left, y);
      y += 12;
    }
    
    if (userCargo) {
      doc.setFont("helvetica", "bold");
      doc.text(userCargo, left, y);
    }

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
      user: $filterUser ? $filterUser.value.trim().toLowerCase() : "",
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
      const userName = q.userName || q.userUsername || "N/A";
      tr.innerHTML = `
        <td>${escapeHtml(displayNumber)}</td>
        <td>${formatDateHuman(q.date)}</td>
        <td>${escapeHtml(q.clientName)}</td>
        <td>${escapeHtml(productsText)}</td>
        <td class="text-center">${formatMoney(q.totalCOP != null ? q.totalCOP : q.totalUSD)}</td>
        <td>${escapeHtml(userName)}</td>
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
      Number((q.totalCOP != null ? q.totalCOP : (q.totalUSD != null ? q.totalUSD : 0))),
      q.userName || q.userUsername || "N/A"
    ]);

    const headers = ["Consecutivo", "Fecha", "Cliente", "Productos", "Total COP", "Usuario"];
    const filename = `Historial_Cotizaciones_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // Intentar usar ExcelJS con estilos (igual que clientes, contactos y usuarios)
    const success = await applyHeaderStylesWithExcelJS(
      historyRows, 
      headers, 
      filename, 
      "Historial",
      [12, 15, 30, 50, 15, 20]
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
      { wch: 15 }, // Total COP
      { wch: 20 }  // Usuario
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

    // Establecer z-index antes de mostrar el modal para que aparezca encima de otros modales
    modalEl.style.zIndex = '1065';
    
    // Mostrar modal
    const modal = bootstrap.Modal.getOrCreateInstance(modalEl, {
      backdrop: true,
      keyboard: true
    });
    
    // Ajustar backdrop cuando el modal se muestra
    const onShown = () => {
      // Ajustar backdrop solo si hay múltiples modales abiertos
      const backdrops = document.querySelectorAll('.modal-backdrop');
      if (backdrops.length > 1) {
        // El último backdrop (del modal de alerta) debe tener z-index mayor
        const lastBackdrop = backdrops[backdrops.length - 1];
        lastBackdrop.style.zIndex = '1060';
        // Asegurar que el backdrop anterior no bloquee la interacción
        if (backdrops.length > 1) {
          backdrops[backdrops.length - 2].style.zIndex = '1050';
        }
      }
      
      modalEl.removeEventListener('shown.bs.modal', onShown);
    };
    
    const onHidden = () => {
      // Restaurar z-index por defecto
      modalEl.style.zIndex = '';
      modalEl.removeEventListener('hidden.bs.modal', onHidden);
    };
    
    modalEl.addEventListener("shown.bs.modal", onShown, { once: true });
    modalEl.addEventListener("hidden.bs.modal", onHidden, { once: true });
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

  function applyFilters(list, { consecutivo, client, product, user, from, to }) {
    return list.filter((q) => {
      const quoteNumber = (q.quoteNumber || "").toUpperCase();
      const matchesConsecutivo = consecutivo ? quoteNumber.includes(consecutivo) : true;
      const matchesClient = client ? (q.clientName || "").toLowerCase().includes(client) : true;
      const matchesProduct = product ? (q.products || []).some((p) => (p.nombre || "").toLowerCase().includes(product)) : true;
      const userName = (q.userName || q.userUsername || "").toLowerCase();
      const matchesUser = user ? userName.includes(user) : true;
      const date = q.date || "";
      const matchesFrom = from ? date >= from : true;
      const matchesTo = to ? date <= to : true;
      return matchesConsecutivo && matchesClient && matchesProduct && matchesUser && matchesFrom && matchesTo;
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
    if (num === 0) return "";
    // Formato: $ 1,234,567 (con coma como separador de miles y sin decimales)
    return "$ " + num.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
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
      // Ordenar productos alfabéticamente por nombre
      const sortedProducts = [...productos].sort((a, b) => {
        const nameA = (a.nombre || "").toLowerCase().trim();
        const nameB = (b.nombre || "").toLowerCase().trim();
        return nameA.localeCompare(nameB, 'es', { sensitivity: 'base' });
      });
      
      sortedProducts.forEach((product) => {
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
      u.email || "",
      u.cargo || "",
      u.role === "admin" ? "Administrador" : "Vendedor",
      u.active ? "Activo" : "Inactivo"
    ]);

    const headers = ["Usuario", "Nombre", "Email", "Cargo", "Rol", "Estado"];
    const filename = `Usuarios_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // Intentar usar ExcelJS con estilos
    const success = await applyHeaderStylesWithExcelJS(
      usersRows, 
      headers, 
      filename, 
      "Usuarios",
      [20, 30, 30, 35, 15, 12]
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
      { wch: 30 }, // Email
      { wch: 35 }, // Cargo
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
        <td>${escapeHtml(user.email || "")}</td>
        <td>${escapeHtml(user.cargo || "")}</td>
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
        // No ejecutar si el botón está deshabilitado
        if (btn.disabled) return;
        
        const action = btn.dataset.action;
        const id = btn.dataset.id;
        const user = users.find((u) => u.id === id);
        
        if (!user) return;
        
        if (action === "edit") {
          openUserModal(user);
        } else if (action === "delete" && user.id !== currentUser.id) {
          deleteUser(id);
        }
      });
    });
  }

  function openUserModal(user = null) {
    if (!$userModalTitle || !$userId || !$userUsername || !$userEmail || !$userName || !$userCargo || !$userPassword || !$userRole || !$userActive) return;

    if (user) {
      $userModalTitle.textContent = "Editar Usuario";
      $userId.value = user.id;
      $userUsername.value = user.username || "";
      $userEmail.value = user.email || "";
      $userName.value = user.name || "";
      $userCargo.value = user.cargo || "";
      $userPassword.value = "";
      if ($userPasswordConfirm) $userPasswordConfirm.value = "";
      $userRole.value = user.role || "";
      $userActive.checked = user.active !== false;
      
      // Ocultar requerimiento de contraseña al editar
      if ($passwordRequired) $passwordRequired.classList.add("d-none");
      if ($passwordHint) $passwordHint.classList.remove("d-none");
      if ($userPassword) $userPassword.removeAttribute("required");
      if ($passwordConfirmRequired) $passwordConfirmRequired.classList.add("d-none");
      if ($passwordConfirmHint) $passwordConfirmHint.textContent = "Dejar en blanco para mantener la actual";
      if ($userPasswordConfirm) $userPasswordConfirm.removeAttribute("required");
    } else {
      $userModalTitle.textContent = "Nuevo Usuario";
      $userId.value = "";
      $userUsername.value = "";
      $userEmail.value = "";
      $userName.value = "";
      $userCargo.value = "";
      $userPassword.value = "";
      if ($userPasswordConfirm) $userPasswordConfirm.value = "";
      $userRole.value = "";
      $userActive.checked = true;
      
      // Mostrar requerimiento de contraseña al crear
      if ($passwordRequired) $passwordRequired.classList.remove("d-none");
      if ($passwordHint) $passwordHint.classList.remove("d-none");
      if ($passwordHint) $passwordHint.textContent = "Mínimo 8 caracteres";
      if ($userPassword) $userPassword.setAttribute("required", "required");
      if ($passwordConfirmRequired) $passwordConfirmRequired.classList.remove("d-none");
      if ($passwordConfirmHint) $passwordConfirmHint.textContent = "Repite la contraseña para verificar";
      if ($userPasswordConfirm) $userPasswordConfirm.setAttribute("required", "required");
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
    const email = $userEmail.value.trim();
    const name = $userName.value.trim();
    const cargo = $userCargo.value.trim();
    const password = $userPassword.value.trim();
    const passwordConfirm = $userPasswordConfirm ? $userPasswordConfirm.value.trim() : "";
    const role = $userRole.value;
    const active = $userActive.checked;
    const id = $userId.value.trim();

    if (!username || !email || !name || !role) {
      showAlert("Por favor completa todos los campos requeridos.", "warning");
      return;
    }

    // Validar formato de email
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      showAlert("Por favor ingresa un email válido.", "warning");
      return;
    }

    // Validar contraseña
    if (id) {
      // Editar usuario: si se llena la contraseña, debe tener mínimo 8 caracteres y coincidir con la confirmación
      if (password) {
        if (password.length < 8) {
          showAlert("La contraseña debe tener mínimo 8 caracteres.", "warning");
          return;
        }
        if (password !== passwordConfirm) {
          showAlert("Las contraseñas no coinciden.", "warning");
          return;
        }
      } else if (passwordConfirm) {
        // Si se llena la confirmación pero no la contraseña
        showAlert("Debes ingresar la contraseña si deseas cambiarla.", "warning");
        return;
      }
    } else {
      // Crear usuario: la contraseña es obligatoria
      if (!password) {
        showAlert("La contraseña es requerida para nuevos usuarios.", "warning");
        return;
      }
      if (password.length < 8) {
        showAlert("La contraseña debe tener mínimo 8 caracteres.", "warning");
        return;
      }
      if (password !== passwordConfirm) {
        showAlert("Las contraseñas no coinciden.", "warning");
        return;
      }
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

      // Verificar que el email no esté en uso por otro usuario
      const emailTaken = users.some((u) => u.id !== id && u.email && u.email.toLowerCase() === email.toLowerCase());
      if (emailTaken) {
        showAlert("El email ya está en uso por otro usuario.", "warning");
        return;
      }

      updatedUsers = users.map((u) =>
        u.id === id
          ? {
              ...u,
              username,
              email,
              name,
              cargo,
              role,
              active,
              password: password ? hashPassword(password) : u.password,
              updatedAt: new Date().toISOString()
            }
          : u
      );
    } else {
      // Crear nuevo usuario
      // La validación de contraseña ya se hizo arriba

      // Verificar que el username no esté en uso
      const usernameTaken = users.some((u) => u.username.toLowerCase() === username.toLowerCase());
      if (usernameTaken) {
        showAlert("El nombre de usuario ya está en uso.", "warning");
        return;
      }

      // Verificar que el email no esté en uso
      const emailTaken = users.some((u) => u.email && u.email.toLowerCase() === email.toLowerCase());
      if (emailTaken) {
        showAlert("El email ya está en uso por otro usuario.", "warning");
        return;
      }

      const newUser = {
        id: `USER-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        username,
        email,
        password: hashPassword(password),
        name,
        cargo,
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

  // Funciones para crear/editar productos
  function generateProductId(nombre) {
    // Generar ID basado en el nombre: primera letra + nombre normalizado
    const primeraLetra = nombre.trim().charAt(0).toUpperCase();
    const nombreNormalizado = nombre.trim()
      .toUpperCase()
      .replace(/[^A-Z0-9\s]/g, '')
      .replace(/\s+/g, '-')
      .substring(0, 30);
    return `${primeraLetra}-${nombreNormalizado}`;
  }

  function openSelectProductModal() {
    if (!$selectProductModal || !$selectProductList || typeof bootstrap === "undefined") return;
    
    // Validar que el usuario sea administrador
    if (!currentUser || currentUser.role !== "admin") {
      showAlert("Solo los administradores pueden editar productos.", "error", "Acceso denegado");
      return;
    }
    
    if (productos.length === 0) {
      showAlert("No hay productos para editar.", "warning");
      return;
    }
    
    // Limpiar buscador
    if ($selectProductSearch) {
      $selectProductSearch.value = "";
    }
    
    // Renderizar lista de productos ordenados alfabéticamente
    renderSelectProductList();
    
    const modal = bootstrap.Modal.getOrCreateInstance($selectProductModal);
    modal.show();
  }

  function renderSelectProductList() {
    if (!$selectProductList) return;
    
    $selectProductList.innerHTML = "";
    
    const searchTerm = $selectProductSearch ? $selectProductSearch.value.toLowerCase().trim() : "";
    
    // Filtrar y ordenar productos
    let filteredProducts = productos;
    if (searchTerm) {
      filteredProducts = productos.filter(p => 
        (p.nombre || "").toLowerCase().includes(searchTerm)
      );
    }
    
    // Ordenar alfabéticamente
    const sortedProducts = [...filteredProducts].sort((a, b) => {
      const nameA = (a.nombre || "").toLowerCase().trim();
      const nameB = (b.nombre || "").toLowerCase().trim();
      return nameA.localeCompare(nameB, 'es', { sensitivity: 'base' });
    });
    
    if (sortedProducts.length === 0) {
      $selectProductList.innerHTML = '<div class="text-center text-muted py-3">No se encontraron productos.</div>';
      return;
    }
    
    sortedProducts.forEach((product) => {
      const item = document.createElement("a");
      item.href = "#";
      item.className = "list-group-item list-group-item-action";
      item.innerHTML = `
        <div class="d-flex justify-content-between align-items-center">
          <span>${escapeHtml(product.nombre || "")}</span>
          <i class="bi bi-chevron-right text-muted"></i>
        </div>
      `;
      item.addEventListener("click", (e) => {
        e.preventDefault();
        // Cerrar modal de selección
        if (typeof bootstrap !== "undefined") {
          const selectModal = bootstrap.Modal.getInstance($selectProductModal);
          if (selectModal) selectModal.hide();
        }
        // Abrir modal de edición con el producto seleccionado
        openProductModalForEdit(product);
      });
      $selectProductList.appendChild(item);
    });
  }

  function openProductModalForEdit(product) {
    if (!$productModal || typeof bootstrap === "undefined") return;
    
    // Validar que el usuario sea administrador
    if (!currentUser || currentUser.role !== "admin") {
      showAlert("Solo los administradores pueden editar productos.", "error", "Acceso denegado");
      return;
    }
    
    const modal = bootstrap.Modal.getOrCreateInstance($productModal);
    
    $productModalTitle.textContent = "Editar Producto";
    $productId.value = product.id;
    $productNombre.value = product.nombre;
    
    // Limpiar procesos existentes
    $procesosContainer.innerHTML = "";
    
    // Agregar procesos existentes
    if (product.procesos && product.procesos.length > 0) {
      product.procesos.forEach((proceso, index) => {
        addProcesoRow(proceso, index);
      });
    } else {
      // Si no hay procesos, agregar uno vacío
      addProcesoRow();
    }
    
    modal.show();
  }

  function openProductModal(isEdit = false) {
    if (!$productModal) {
      console.error("Modal de producto no encontrado");
      return;
    }
    if (typeof bootstrap === "undefined") {
      console.error("Bootstrap no está definido");
      return;
    }
    
    // Validar que el usuario sea administrador
    if (!currentUser || currentUser.role !== "admin") {
      showAlert("Solo los administradores pueden crear o editar productos.", "error", "Acceso denegado");
      return;
    }
    
    if (isEdit) {
      // Modo edición: abrir modal de selección
      openSelectProductModal();
      return;
    }
    
    // Modo creación
    try {
      const modal = bootstrap.Modal.getOrCreateInstance($productModal);
      if ($productModalTitle) $productModalTitle.textContent = "Nuevo Producto";
      if ($productId) $productId.value = "";
      if ($productNombre) $productNombre.value = "";
      if ($procesosContainer) $procesosContainer.innerHTML = "";
      // Agregar una fila vacía inicial
      addProcesoRow();
      
      modal.show();
    } catch (error) {
      console.error("Error al abrir el modal de producto:", error);
      showAlert("Error al abrir el formulario de producto. Por favor, recarga la página.", "error", "Error");
    }
  }

  function formatCurrencyInput(value) {
    // Remover todo excepto números
    const numStr = String(value || '').replace(/[^0-9]/g, '');
    if (!numStr) return '';
    
    // Convertir a número y formatear
    const num = parseInt(numStr, 10);
    if (isNaN(num) || num === 0) return '';
    
    // Formatear con comas
    return '$' + num.toLocaleString('es-CO');
  }

  function parseCurrencyInput(value) {
    // Remover signo de peso y comas, convertir a número
    const numStr = String(value || '').replace(/[^0-9]/g, '');
    return numStr ? parseInt(numStr, 10) : 0;
  }

  function addProcesoRow(procesoData = null, index = null) {
    if (!$procesosContainer) return;
    
    const rowIndex = index !== null ? index : $procesosContainer.children.length;
    const tr = document.createElement("tr");
    
    // Formatear valores de precio si existen
    const vrUnit1Formatted = procesoData?.vrUnit1 ? formatCurrencyInput(procesoData.vrUnit1) : '';
    const vrUnit2Formatted = procesoData?.vrUnit2 ? formatCurrencyInput(procesoData.vrUnit2) : '';
    const vrUnit3Formatted = procesoData?.vrUnit3 ? formatCurrencyInput(procesoData.vrUnit3) : '';
    const vrUnitUSDFormatted = procesoData?.vrUnitUSD ? formatCurrencyInput(procesoData.vrUnitUSD) : '';
    
    tr.innerHTML = `
      <td class="text-center align-middle" style="width: 40px;">${rowIndex + 1}</td>
      <td><input type="text" class="form-control form-control-sm proceso-codigo" value="${escapeHtml(procesoData?.codigo || '')}" placeholder="Ej: AEU-001"></td>
      <td><input type="text" class="form-control form-control-sm proceso-analisis" value="${escapeHtml(procesoData?.analisis || '')}" placeholder="Ej: Identificación" required></td>
      <td><input type="text" class="form-control form-control-sm proceso-metodo" value="${escapeHtml(procesoData?.metodo || '')}" placeholder="Ej: USP <197>"></td>
      <td><input type="number" class="form-control form-control-sm proceso-cmtra" value="${procesoData?.cMtra_g || procesoData?.cMtra || ''}" placeholder="0" min="0" step="0.01"></td>
      <td><input type="number" class="form-control form-control-sm proceso-cantidad" value="${procesoData?.cantidad !== undefined && procesoData?.cantidad !== null && procesoData?.cantidad !== '' ? procesoData.cantidad : ''}" placeholder="" min="0" step="1"></td>
      <td><input type="text" class="form-control form-control-sm proceso-vrunit1" value="${vrUnit1Formatted}" placeholder="$0"></td>
      <td><input type="text" class="form-control form-control-sm proceso-vrunit2" value="${vrUnit2Formatted}" placeholder="$0"></td>
      <td><input type="text" class="form-control form-control-sm proceso-vrunit3" value="${vrUnit3Formatted}" placeholder="$0"></td>
      <td><input type="text" class="form-control form-control-sm proceso-vrunitusd" value="${vrUnitUSDFormatted}" placeholder="$0"></td>
      <td class="text-center" style="width: 60px;">
        <button type="button" class="btn btn-sm btn-outline-danger btn-remove-proceso" title="Eliminar">
          <i class="bi bi-trash"></i>
        </button>
      </td>
    `;
    
    // Agregar event listeners para formatear campos de precio
    const priceInputs = [
      tr.querySelector(".proceso-vrunit1"),
      tr.querySelector(".proceso-vrunit2"),
      tr.querySelector(".proceso-vrunit3"),
      tr.querySelector(".proceso-vrunitusd")
    ];
    
    priceInputs.forEach(input => {
      if (input) {
        // Al hacer foco, limpiar el formato para facilitar la edición
        input.addEventListener("focus", (e) => {
          const value = e.target.value;
          // Remover formato y dejar solo números
          const numOnly = value.replace(/[^0-9]/g, '');
          if (numOnly) {
            e.target.value = numOnly;
          }
        });
        
        // Formatear al perder el foco
        input.addEventListener("blur", (e) => {
          const value = e.target.value;
          if (value.trim()) {
            const formatted = formatCurrencyInput(value);
            e.target.value = formatted;
          }
        });
        
        // Permitir solo números mientras se escribe
        input.addEventListener("input", (e) => {
          const value = e.target.value;
          const numOnly = value.replace(/[^0-9]/g, '');
          if (value !== numOnly) {
            const cursorPos = e.target.selectionStart;
            e.target.value = numOnly;
            // Ajustar posición del cursor
            const newPos = Math.max(0, cursorPos - (value.length - numOnly.length));
            e.target.setSelectionRange(newPos, newPos);
          }
        });
      }
    });
    
    // Agregar event listener al botón eliminar
    const removeBtn = tr.querySelector(".btn-remove-proceso");
    removeBtn.addEventListener("click", () => {
      tr.remove();
      updateProcesoRowNumbers();
    });
    
    $procesosContainer.appendChild(tr);
  }

  function updateProcesoRowNumbers() {
    if (!$procesosContainer) return;
    const rows = $procesosContainer.querySelectorAll("tr");
    rows.forEach((row, index) => {
      const numberCell = row.querySelector("td:first-child");
      if (numberCell) {
        numberCell.textContent = index + 1;
      }
    });
  }

  function saveProduct() {
    if (!$productFormModal || !$productFormModal.checkValidity()) {
      $productFormModal.reportValidity();
      return;
    }

    const nombre = $productNombre.value.trim();
    if (!nombre) {
      showAlert("Por favor ingresa el nombre del producto.", "warning");
      return;
    }

    // Recopilar procesos
    const procesos = [];
    const rows = $procesosContainer.querySelectorAll("tr");
    
    if (rows.length === 0) {
      showAlert("Debes agregar al menos un análisis al producto.", "warning");
      return;
    }

    rows.forEach((row) => {
      const codigo = row.querySelector(".proceso-codigo")?.value.trim() || "";
      const analisis = row.querySelector(".proceso-analisis")?.value.trim() || "";
      const metodo = row.querySelector(".proceso-metodo")?.value.trim() || "";
      const cMtra_g = parseFloat(row.querySelector(".proceso-cmtra")?.value) || 0;
      const cantidadInput = row.querySelector(".proceso-cantidad")?.value.trim();
      const cantidad = cantidadInput !== "" && !isNaN(parseInt(cantidadInput)) ? parseInt(cantidadInput) : undefined;
      const vrUnit1 = parseCurrencyInput(row.querySelector(".proceso-vrunit1")?.value) || 0;
      const vrUnit2 = parseCurrencyInput(row.querySelector(".proceso-vrunit2")?.value) || 0;
      const vrUnit3 = parseCurrencyInput(row.querySelector(".proceso-vrunit3")?.value) || 0;
      const vrUnitUSD = parseCurrencyInput(row.querySelector(".proceso-vrunitusd")?.value) || 0;

      // Validar que al menos tenga análisis
      if (analisis) {
        procesos.push({
          codigo,
          analisis,
          metodo,
          cMtra_g: cMtra_g > 0 ? cMtra_g : undefined,
          cantidad: cantidad !== undefined && cantidad > 0 ? cantidad : undefined,
          vrUnit1: vrUnit1 > 0 ? vrUnit1 : undefined,
          vrUnit2: vrUnit2 > 0 ? vrUnit2 : undefined,
          vrUnit3: vrUnit3 > 0 ? vrUnit3 : undefined,
          vrUnitUSD: vrUnitUSD > 0 ? vrUnitUSD : undefined
        });
      }
    });

    if (procesos.length === 0) {
      showAlert("Debes agregar al menos un análisis válido al producto.", "warning");
      return;
    }

    const productId = $productId.value.trim();
    let updatedProducts;

    if (productId) {
      // Editar producto existente
      const existingProduct = productos.find((p) => p.id === productId);
      if (!existingProduct) {
        showAlert("Producto no encontrado.", "error");
        return;
      }

      // Verificar que el nombre no esté en uso por otro producto
      const nameTaken = productos.some((p) => p.id !== productId && p.nombre.toLowerCase() === nombre.toLowerCase());
      if (nameTaken) {
        showAlert("Ya existe un producto con ese nombre.", "warning");
        return;
      }

      updatedProducts = productos.map((p) =>
        p.id === productId
          ? {
              ...p,
              nombre,
              procesos
            }
          : p
      );
    } else {
      // Crear nuevo producto
      // Verificar que el nombre no esté en uso
      const nameTaken = productos.some((p) => p.nombre.toLowerCase() === nombre.toLowerCase());
      if (nameTaken) {
        showAlert("Ya existe un producto con ese nombre.", "warning");
        return;
      }

      const newId = generateProductId(nombre);
      // Asegurar que el ID sea único
      let finalId = newId;
      let counter = 1;
      while (productos.some(p => p.id === finalId)) {
        finalId = `${newId}-${counter}`;
        counter++;
      }

      const newProduct = {
        id: finalId,
        nombre,
        procesos
      };
      updatedProducts = [...productos, newProduct];
    }

    // Guardar productos
    productos = updatedProducts;
    saveCatalog(productos);
    
    // Limpiar selección si se editó
    if (productId) {
      selectedProductIds.delete(productId);
      updateSelectionStateUI();
    }
    
    // Recargar productos
    renderProductsByLetter(currentLetter);
    
    // Actualizar el modal de eliminar productos si está abierto
    if ($deleteProductsModal && typeof bootstrap !== "undefined") {
      const deleteModal = bootstrap.Modal.getInstance($deleteProductsModal);
      if (deleteModal && deleteModal._isShown) {
        // El modal está abierto, actualizar la lista
        openDeleteProductsModal();
      }
    }
    
    closeProductModal();
    showAlert(productId ? "Producto actualizado correctamente." : "Producto creado correctamente.", "success");
  }

  function closeProductModal() {
    if (!$productModal || typeof bootstrap === "undefined") return;
    const modal = bootstrap.Modal.getOrCreateInstance($productModal);
    modal.hide();
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
    if ($userPasswordConfirm) $userPasswordConfirm.value = "";
    if ($passwordRequired) $passwordRequired.classList.remove("d-none");
    if ($passwordHint) $passwordHint.classList.add("d-none");
    if ($passwordHint) $passwordHint.textContent = "Mínimo 8 caracteres. Dejar en blanco para mantener la actual (solo al editar)";
    if ($passwordConfirmRequired) $passwordConfirmRequired.classList.remove("d-none");
    if ($passwordConfirmHint) $passwordConfirmHint.textContent = "Repite la contraseña para verificar";
  }

  // Manejar accesibilidad de modales - prevenir foco en modales ocultos
  function setupModalAccessibility() {
    // Escuchar cuando cualquier modal comienza a ocultarse
    document.addEventListener('hide.bs.modal', function(event) {
      const modal = event.target;
      // Remover foco de cualquier elemento dentro del modal antes de que se oculte
      const focusedElement = modal.querySelector(':focus');
      if (focusedElement) {
        focusedElement.blur();
      }
    });

    // Escuchar cuando cualquier modal se oculta completamente
    document.addEventListener('hidden.bs.modal', function(event) {
      const modal = event.target;
      // Remover foco de cualquier elemento dentro del modal
      const focusedElement = modal.querySelector(':focus');
      if (focusedElement) {
        focusedElement.blur();
      }
      // Asegurar que aria-hidden esté correctamente configurado
      if (!modal.classList.contains('show')) {
        modal.setAttribute('aria-hidden', 'true');
        // Remover tabindex de todos los elementos interactivos dentro del modal
        const interactiveElements = modal.querySelectorAll('button, a, input, select, textarea, [tabindex]');
        interactiveElements.forEach(el => {
          if (el.hasAttribute('tabindex') && el.getAttribute('tabindex') !== '-1') {
            el.setAttribute('data-original-tabindex', el.getAttribute('tabindex'));
            el.setAttribute('tabindex', '-1');
          } else if (!el.hasAttribute('tabindex')) {
            el.setAttribute('tabindex', '-1');
          }
        });
      }
    });

    // Escuchar cuando cualquier modal se muestra
    document.addEventListener('shown.bs.modal', function(event) {
      const modal = event.target;
      // Remover aria-hidden cuando el modal está visible
      modal.removeAttribute('aria-hidden');
      // Restaurar tabindex de elementos interactivos
      const interactiveElements = modal.querySelectorAll('[data-original-tabindex], [tabindex="-1"]');
      interactiveElements.forEach(el => {
        if (el.hasAttribute('data-original-tabindex')) {
          el.setAttribute('tabindex', el.getAttribute('data-original-tabindex'));
          el.removeAttribute('data-original-tabindex');
        } else if (el.tagName === 'BUTTON' || el.tagName === 'A' || el.tagName === 'INPUT' || el.tagName === 'SELECT' || el.tagName === 'TEXTAREA') {
          el.removeAttribute('tabindex');
        }
      });
    });

    // Prevenir foco en elementos dentro de modales ocultos
    document.addEventListener('focusin', function(event) {
      const target = event.target;
      const modal = target.closest('.modal');
      if (modal && !modal.classList.contains('show')) {
        // Si el elemento está dentro de un modal oculto, quitar el foco
        event.preventDefault();
        event.stopPropagation();
        target.blur();
        // Enfocar el body para evitar problemas de accesibilidad
        if (document.activeElement === target) {
          document.body.focus();
        }
      }
    }, true);
  }

  // Inicializar accesibilidad de modales
  setupModalAccessibility();

  // Funciones de depuración para verificar el estado del localStorage
  window.debugUsers = function() {
    console.log("=== DIAGNÓSTICO DE USUARIOS ===");
    console.log("1. Usuarios en localStorage:", getUsers());
    console.log("2. Datos raw en localStorage:", localStorage.getItem("olgroup_users"));
    console.log("3. Usuario actual en sessionStorage:", getCurrentUser());
    console.log("4. Estado de localStorage:", {
      disponible: typeof Storage !== "undefined",
      espacioDisponible: (function() {
        try {
          localStorage.setItem("test", "test");
          localStorage.removeItem("test");
          return "OK";
        } catch (e) {
          return "ERROR: " + e.message;
        }
      })()
    });
    return {
      users: getUsers(),
      rawData: localStorage.getItem("olgroup_users"),
      currentUser: getCurrentUser()
    };
  };

  window.backupUsers = function() {
    try {
      const users = getUsers();
      const backup = {
        timestamp: new Date().toISOString(),
        users: users
      };
      const backupStr = JSON.stringify(backup, null, 2);
      console.log("=== BACKUP DE USUARIOS ===");
      console.log(backupStr);
      // También copiar al portapapeles si es posible
      if (navigator.clipboard) {
        navigator.clipboard.writeText(backupStr).then(() => {
          console.log("Backup copiado al portapapeles");
          alert("Backup de usuarios copiado al portapapeles y mostrado en consola.");
        }).catch(() => {
          alert("Backup generado. Revisa la consola (F12) para ver los datos.");
        });
      } else {
        alert("Backup generado. Revisa la consola (F12) para ver los datos.");
      }
      return backup;
    } catch (error) {
      console.error("Error al crear backup:", error);
      alert("Error al crear backup: " + error.message);
      return null;
    }
  };

  window.restoreUsers = function(usersData) {
    if (!usersData || !Array.isArray(usersData)) {
      console.error("Error: usersData debe ser un array");
      alert("Error: Los datos deben ser un array de usuarios.");
      return false;
    }
    
    if (confirm(`¿Estás seguro de que deseas restaurar ${usersData.length} usuarios? Esto reemplazará todos los usuarios actuales.`)) {
      if (saveUsers(usersData)) {
        alert("Usuarios restaurados correctamente. Recarga la página.");
        location.reload();
        return true;
      } else {
        alert("Error al restaurar usuarios.");
        return false;
      }
    }
    return false;
  };

  // Inicio robusto: si el DOM ya está listo, inicializa de inmediato
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
