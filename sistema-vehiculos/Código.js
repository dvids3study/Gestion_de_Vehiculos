// =====================================================
// SISTEMA DE VEHÍCULOS
// CÓDIGO DEL SERVIDOR - GOOGLE APPS SCRIPT
// =====================================================

// =====================================================
// CONFIGURACIÓN GENERAL
// =====================================================

const ID_ARCHIVO = "15cfN5BSMsaY3xpNE9rJ3ObGYN6Wu_ihHLnBOcXJ9r2s";

const NOMBRE_HOJA = "VEHICULOS";

const NOMBRE_HOJA_USUARIOS = "USUARIOS";

const NOMBRE_HOJA_PARQUEADEROS = "PARQUEADEROS";

const NOMBRE_HOJA_CITAS = "CITAS";

const EMAIL_ALERTAS_CITAS = "TU_CORREO@gmail.com";

// =====================================================
// CONFIGURACIÓN DE SEGURIDAD
// =====================================================

// Nombre utilizado para guardar la clave secreta
// del sistema en las propiedades del proyecto.
const CLAVE_SECRETA_JWT = "JWT_SECRET_VEHICULOS";

// Tiempo de duración de una sesión.
// 8 horas.
const DURACION_SESION_MS = 60 * 60 * 1000;

// Tiempo de duración del token de recuperación.
// 15 minutos.
const DURACION_RECUPERACION_MS = 15 * 60 * 1000;

// =====================================================
// doGet
// =====================================================

function doGet(e) {
  const html = HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Sistema de Vehículos")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  html.addMetaTag("viewport", "width=device-width, initial-scale=1.0");

  return html;
}

// =====================================================
// INCLUIR ARCHIVOS HTML
// =====================================================

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

// =====================================================
// CONSULTA PRINCIPAL
// =====================================================

function consultarDesdeApp(pregunta) {
  return procesarConsulta(pregunta);
}

// =====================================================
// PROCESAR CONSULTA
// =====================================================

function procesarConsulta(pregunta) {
  if (!pregunta) {
    return {
      error: "No se recibió ninguna pregunta.",
    };
  }

  const texto = pregunta.toString().trim().toUpperCase();

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      error: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const rango = hoja.getDataRange();

  const datos = rango.getValues();

  const datosTexto = rango.getDisplayValues();

  let tipoConsulta = "";

  // ---------------------------------------------------
  // CONSULTA DE CIRCULACIÓN
  // ---------------------------------------------------

  if (
    texto.includes("PUEDO LLEVAR") ||
    texto.includes("PUEDO SACAR") ||
    texto.includes("PUEDE SALIR") ||
    texto.includes("PUEDE CIRCULAR") ||
    texto.includes("LLEVAR")
  ) {
    tipoConsulta = "CIRCULACION";
  }

  // ---------------------------------------------------
  // CONSULTA DE INFORMACIÓN
  // ---------------------------------------------------
  else if (
    texto.includes("INFO") ||
    texto.includes("INFORMACION") ||
    texto.includes("DATOS")
  ) {
    tipoConsulta = "INFORMACION";
  } else {
    return {
      encontrado: false,
      error: "No entendí qué necesitas.",
      ejemplos: ["PUEDO LLEVARME ONIX 2017", "INFO ONIX 2017"],
    };
  }

  // ---------------------------------------------------
  // BUSCAR VEHÍCULO
  // ---------------------------------------------------

  let vehiculo = null;

  for (let i = 1; i < datos.length; i++) {
    const nombre = datosTexto[i][0].toString().trim().toUpperCase();

    if (nombre && texto.includes(nombre)) {
      vehiculo = {
        nombre: datosTexto[i][0],

        placa: datosTexto[i][1],

        soat: convertirFecha(datosTexto[i][2]),

        tecnomecanica: convertirFecha(datosTexto[i][3]),

        picoYPlaca: datosTexto[i][4],

        precio: datos[i][5],

        kilometraje: datos[i][6],

        detalles: datosTexto[i][7],

        ubicacion: datosTexto[i][8] || "",

        puesto: datosTexto[i][9] || "",

        transito: datosTexto[i][10] || "",
      };

      break;
    }
  }

  // ---------------------------------------------------
  // VEHÍCULO NO ENCONTRADO
  // ---------------------------------------------------

  if (!vehiculo) {
    return {
      encontrado: false,

      mensaje: "No encontré ese vehículo en la lista.",
    };
  }

  // ===================================================
  // INFORMACIÓN
  // ===================================================

  if (tipoConsulta === "INFORMACION") {
    const soatVigente = fechaVigente(vehiculo.soat);

    const tecnomecanicaVigente = fechaVigente(vehiculo.tecnomecanica);

    const mensaje =
      "🚗 " +
      vehiculo.nombre +
      "\n\n" +
      "🔢 Placa: " +
      vehiculo.placa +
      "\n" +
      "💰 Precio: $" +
      formatearNumero(vehiculo.precio) +
      "\n" +
      "📏 Kilometraje: " +
      formatearNumero(vehiculo.kilometraje) +
      " km\n\n" +
      "📄 SOAT: " +
      (soatVigente ? "VIGENTE" : "VENCIDO") +
      " hasta el " +
      formatearFecha(vehiculo.soat) +
      "\n" +
      "🔧 Tecnomecánica: " +
      (tecnomecanicaVigente ? "VIGENTE" : "VENCIDA") +
      " hasta el " +
      formatearFecha(vehiculo.tecnomecanica) +
      "\n" +
      "🚦 Pico y placa: " +
      vehiculo.picoYPlaca +
      "\n\n" +
      "📍 Ubicación: " +
      (vehiculo.ubicacion || "Sin ubicación asignada") +
      "\n" +
      "🅿️ Puesto: " +
      (vehiculo.puesto || "Sin puesto asignado") +
      "\n" +
      "🏛️ Tránsito: " +
      (vehiculo.transito || "Sin tránsito registrado");

    return {
      encontrado: true,

      tipo: "INFORMACION",

      mensaje: mensaje,

      nombre: vehiculo.nombre,

      placa: vehiculo.placa,

      detalles: vehiculo.detalles,

      soat: {
        fecha: formatearFecha(vehiculo.soat),

        vigente: soatVigente,
      },

      tecnomecanica: {
        fecha: formatearFecha(vehiculo.tecnomecanica),

        vigente: tecnomecanicaVigente,
      },

      picoYPlaca: vehiculo.picoYPlaca,

      precio: vehiculo.precio,

      kilometraje: vehiculo.kilometraje,

      ubicacion: vehiculo.ubicacion,

      puesto: vehiculo.puesto,

      transito: vehiculo.transito,
    };
  }

  // ===================================================
  // CIRCULACIÓN
  // ===================================================

  if (tipoConsulta === "CIRCULACION") {
    const soatVigente = fechaVigente(vehiculo.soat);

    const tecnomecanicaVigente = fechaVigente(vehiculo.tecnomecanica);

    const hoy = new Date();

    hoy.setHours(0, 0, 0, 0);

    const diaHoy = obtenerNombreDia(hoy);

    const picoYPlaca = normalizarDia(vehiculo.picoYPlaca);

    const tienePicoHoy = picoYPlaca === normalizarDia(diaHoy);

    const proximoDia = obtenerProximoDiaDeCirculacion(hoy);

    const nombreProximoDia = obtenerNombreDia(proximoDia);

    const tienePicoProximoDia = picoYPlaca === normalizarDia(nombreProximoDia);

    let puedeSalir = true;

    let motivo = "";

    if (tienePicoHoy) {
      puedeSalir = false;

      motivo = "Tiene pico y placa hoy (" + diaHoy + ").";
    } else if (!soatVigente) {
      puedeSalir = false;

      motivo = "El SOAT está vencido.";
    } else if (!tecnomecanicaVigente) {
      puedeSalir = false;

      motivo = "La tecnomecánica está vencida.";
    } else if (tienePicoProximoDia) {
      puedeSalir = false;

      motivo =
        "Tiene pico y placa el próximo día de circulación (" +
        nombreProximoDia +
        ").";
    }

    let mensaje = "";

    if (puedeSalir) {
      mensaje = "SÍ, TE PUEDES LLEVAR EL " + vehiculo.nombre + ".";
    } else {
      mensaje =
        "NO, NO TE PUEDES LLEVAR EL " +
        vehiculo.nombre +
        ".\n\nMotivo: " +
        motivo;
    }

    return {
      encontrado: true,

      tipo: "CIRCULACION",

      mensaje: mensaje,

      puedeSalir: puedeSalir,

      motivo: motivo,

      nombre: vehiculo.nombre,
    };
  }

  return {
    error: "No se pudo procesar la consulta.",
  };
}

// =====================================================
// FESTIVOS DE COLOMBIA
// =====================================================

function obtenerFestivosColombia(año) {
  const cache = CacheService.getScriptCache();

  const clave = "festivos_" + año;

  const almacenado = cache.get(clave);

  if (almacenado) {
    return JSON.parse(almacenado);
  }

  const url = "https://api-colombia.com/api/v1/Holiday/year/" + año;

  try {
    const respuestaAPI = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
    });

    if (respuestaAPI.getResponseCode() !== 200) {
      return [];
    }

    const festivos = JSON.parse(respuestaAPI.getContentText());

    cache.put(clave, JSON.stringify(festivos), 21600);

    return festivos;
  } catch (error) {
    return [];
  }
}

// =====================================================
// VERIFICAR FESTIVO
// =====================================================

function esFestivo(fecha) {
  const año = fecha.getFullYear();

  const festivos = obtenerFestivosColombia(año);

  const fechaBuscada = Utilities.formatDate(
    fecha,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );

  for (let i = 0; i < festivos.length; i++) {
    if (festivos[i].date === fechaBuscada) {
      return true;
    }
  }

  return false;
}

// =====================================================
// PRÓXIMO DÍA DE CIRCULACIÓN
// =====================================================

function obtenerProximoDiaDeCirculacion(fechaInicial) {
  const fecha = new Date(fechaInicial);

  fecha.setHours(0, 0, 0, 0);

  while (true) {
    fecha.setDate(fecha.getDate() + 1);

    const dia = fecha.getDay();

    if (dia === 0) {
      continue;
    }

    if (dia === 6) {
      continue;
    }

    if (esFestivo(fecha)) {
      continue;
    }

    return fecha;
  }
}

// =====================================================
// CONVERTIR FECHA
// =====================================================

function convertirFecha(valor) {
  if (!valor) {
    return null;
  }

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return valor;
  }

  const texto = valor.toString().trim();

  const partes = texto.split("/");

  if (partes.length === 3) {
    const dia = parseInt(partes[0], 10);

    const mes = parseInt(partes[1], 10) - 1;

    let año = parseInt(partes[2], 10);

    if (año < 100) {
      año += 2000;
    }

    const fecha = new Date(año, mes, dia);

    if (!isNaN(fecha.getTime())) {
      return fecha;
    }
  }

  return null;
}

// =====================================================
// FECHA VIGENTE
// =====================================================

function fechaVigente(fecha) {
  if (!fecha || isNaN(fecha.getTime())) {
    return false;
  }

  const hoy = new Date();

  hoy.setHours(0, 0, 0, 0);

  const vencimiento = new Date(fecha);

  vencimiento.setHours(0, 0, 0, 0);

  return vencimiento >= hoy;
}

// =====================================================
// FORMATEAR FECHA
// =====================================================

function formatearFecha(fecha) {
  if (!fecha || isNaN(fecha.getTime())) {
    return "SIN FECHA";
  }

  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

// =====================================================
// NORMALIZAR FECHA PARA HISTORIAL
// =====================================================

function normalizarFechaHistorial(valor) {
  if (valor === null || valor === undefined || valor === "") {
    return "";
  }

  // Si ya es una fecha real
  if (Object.prototype.toString.call(valor) === "[object Date]") {
    if (isNaN(valor.getTime())) {
      return "";
    }

    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }

  const texto = valor.toString().trim();

  if (!texto) {
    return "";
  }

  let fecha = null;

  // Formato: yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
    const partes = texto.split("-");

    fecha = new Date(
      Number(partes[0]),
      Number(partes[1]) - 1,
      Number(partes[2]),
    );
  }

  // Formato: dd/MM/yyyy
  else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(texto)) {
    const partes = texto.split("/");

    fecha = new Date(
      Number(partes[2]),
      Number(partes[1]) - 1,
      Number(partes[0]),
    );
  }

  // Si encontramos una fecha válida
  if (fecha && !isNaN(fecha.getTime())) {
    return Utilities.formatDate(
      fecha,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }

  // Si no es una fecha conocida,
  // devolvemos el texto normalizado
  return texto.toLowerCase().trim();
}

// =====================================================
// FORMATEAR NÚMERO
// =====================================================

function formatearNumero(numero) {
  if (numero === null || numero === undefined || numero === "") {
    return "0";
  }

  return Number(numero).toLocaleString("es-CO");
}

// =====================================================
// OBTENER NOMBRE DEL DÍA
// =====================================================

function obtenerNombreDia(fecha) {
  const dias = [
    "DOMINGO",
    "LUNES",
    "MARTES",
    "MIERCOLES",
    "JUEVES",
    "VIERNES",
    "SABADO",
  ];

  return dias[fecha.getDay()];
}

// =====================================================
// NORMALIZAR DÍA
// =====================================================

function normalizarDia(dia) {
  if (!dia) {
    return "";
  }

  return dia
    .toString()
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// =====================================================
// HISTORIAL GLOBAL
// =====================================================

const CLAVE_HISTORIAL_GLOBAL = "historial_global_vehiculos";

const MAX_HISTORIAL_GLOBAL = 3;

// =====================================================
// OBTENER HISTORIAL
// =====================================================

function obtenerHistorialGlobal() {
  const propiedades = PropertiesService.getScriptProperties();

  const guardado = propiedades.getProperty(CLAVE_HISTORIAL_GLOBAL);

  if (!guardado) {
    return [];
  }

  try {
    const historial = JSON.parse(guardado);

    return Array.isArray(historial) ? historial : [];
  } catch (error) {
    return [];
  }
}

// =====================================================
// OBTENER VEHÍCULOS PARA AUTOCOMPLETADO
// =====================================================

function obtenerVehiculosAutocompletado() {
  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return [];
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return [];
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 1).getDisplayValues();

  return datos

    .map(function (fila) {
      return fila[0].toString().trim();
    })

    .filter(function (vehiculo) {
      return vehiculo !== "";
    });
}

// =====================================================
// GUARDAR HISTORIAL
// =====================================================

function guardarHistorialGlobal(vehiculo) {
  if (!vehiculo) {
    return obtenerHistorialGlobal();
  }

  vehiculo = vehiculo.toString().trim();

  if (!vehiculo) {
    return obtenerHistorialGlobal();
  }

  const lock = LockService.getScriptLock();

  lock.waitLock(5000);

  try {
    const propiedades = PropertiesService.getScriptProperties();

    let historial = [];

    const guardado = propiedades.getProperty(CLAVE_HISTORIAL_GLOBAL);

    if (guardado) {
      try {
        historial = JSON.parse(guardado);

        if (!Array.isArray(historial)) {
          historial = [];
        }
      } catch (error) {
        historial = [];
      }
    }

    historial = historial.filter(function (item) {
      return String(item).toLowerCase() !== vehiculo.toLowerCase();
    });

    historial.unshift(vehiculo);

    historial = historial.slice(0, MAX_HISTORIAL_GLOBAL);

    propiedades.setProperty(CLAVE_HISTORIAL_GLOBAL, JSON.stringify(historial));

    return historial;
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// =====================================================
// AUTENTICACIÓN
// =====================================================
// =====================================================

// =====================================================
// OBTENER HOJA DE USUARIOS
// =====================================================

function obtenerHojaUsuarios() {
  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_USUARIOS);

  if (!hoja) {
    throw new Error('No se encontró la hoja "USUARIOS".');
  }

  return hoja;
}

// =====================================================
// OBTENER TODOS LOS USUARIOS
// =====================================================

function obtenerUsuarios() {
  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return [];
  }

  return hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();
}

// =====================================================
// NORMALIZAR USUARIO
// =====================================================

function normalizarUsuario(usuario) {
  if (!usuario) {
    return "";
  }

  return usuario.toString().trim().toLowerCase();
}

// =====================================================
// NORMALIZAR ROL
// =====================================================

function normalizarRol(rol) {
  if (!rol) {
    return "";
  }

  return rol
    .toString()
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// =====================================================
// NORMALIZAR ESTADO
// =====================================================

function normalizarEstado(estado) {
  if (!estado) {
    return "";
  }

  return estado.toString().trim().toUpperCase();
}

// =====================================================
// HASH SHA-256
// =====================================================

function generarHash(texto) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    texto.toString(),
    Utilities.Charset.UTF_8,
  );

  return bytes
    .map(function (byte) {
      const valor = byte < 0 ? byte + 256 : byte;

      return valor.toString(16).padStart(2, "0");
    })
    .join("");
}

// =====================================================
// COMPROBAR CONTRASEÑA
// =====================================================

function comprobarPassword(contraseñaIngresada, contraseñaGuardada) {
  if (contraseñaIngresada === null || contraseñaIngresada === undefined) {
    return false;
  }

  if (contraseñaGuardada === null || contraseñaGuardada === undefined) {
    return false;
  }

  const ingresada = contraseñaIngresada.toString();

  const guardada = contraseñaGuardada.toString().trim();

  // ---------------------------------------------------
  // CONTRASEÑA CON HASH
  // ---------------------------------------------------

  if (guardada.length === 64 && /^[a-fA-F0-9]+$/.test(guardada)) {
    return generarHash(ingresada).toLowerCase() === guardada.toLowerCase();
  }

  // ---------------------------------------------------
  // COMPATIBILIDAD CON CONTRASEÑA
  // EN TEXTO PLANO
  // ---------------------------------------------------

  return ingresada === guardada;
}

// =====================================================
// INICIAR SESIÓN
// =====================================================

function iniciarSesion(usuario, contraseña) {
  usuario = normalizarUsuario(usuario);

  if (!usuario || !contraseña) {
    return {
      success: false,

      mensaje: "Debes ingresar usuario y contraseña.",
    };
  }

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,

      mensaje: "No existen usuarios registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];

    const usuarioHoja = normalizarUsuario(fila[2]);

    if (usuarioHoja !== usuario) {
      continue;
    }

    const estado = normalizarEstado(fila[6]);

    if (estado !== "ACTIVO") {
      return {
        success: false,

        mensaje: "Este usuario está inactivo.",
      };
    }

    const contraseñaCorrecta = comprobarPassword(contraseña, fila[4]);

    if (!contraseñaCorrecta) {
      return {
        success: false,

        mensaje: "Usuario o contraseña incorrectos.",
      };
    }

    // -------------------------------------------------
    // SI LA CONTRASEÑA ESTABA EN TEXTO PLANO,
    // SE CONVIERTE AUTOMÁTICAMENTE A HASH.
    // -------------------------------------------------

    const contraseñaActual = fila[4].toString().trim();

    const esHash =
      contraseñaActual.length === 64 && /^[a-fA-F0-9]+$/.test(contraseñaActual);

    if (!esHash) {
      hoja.getRange(i + 2, 5).setValue(generarHash(contraseña));
    }

    // -------------------------------------------------
    // GENERAR SESIÓN
    // -------------------------------------------------

    const tokenSesion = generarJWT(
      {
        id: String(fila[0]),

        nombre: String(fila[1]),

        usuario: String(fila[2]),

        email: String(fila[3]),

        rol: normalizarRol(fila[5]),

        tipo: "SESION",

        exp: Date.now() + DURACION_SESION_MS,
      },
      DURACION_SESION_MS,
    );

    return {
      success: true,

      mensaje: "Inicio de sesión exitoso.",

      token: tokenSesion,

      usuario: {
        id: String(fila[0]),

        nombre: String(fila[1]),

        usuario: String(fila[2]),

        email: String(fila[3]),

        rol: normalizarRol(fila[5]),
      },
    };
  }

  return {
    success: false,

    mensaje: "Usuario o contraseña incorrectos.",
  };
}

// =====================================================
// CREAR JWT
// =====================================================

function generarJWT(datos, duracion) {
  const ahora = Date.now();

  const payload = Object.assign({}, datos, {
    iat: ahora,
    exp: ahora + duracion,
  });

  const header = {
    alg: "HS256",

    typ: "JWT",
  };

  const headerEncoded = base64UrlEncode(JSON.stringify(header));

  const payloadEncoded = base64UrlEncode(JSON.stringify(payload));

  const contenido = headerEncoded + "." + payloadEncoded;

  const firma = firmarJWT(contenido);

  return contenido + "." + firma;
}

// =====================================================
// VERIFICAR JWT
// =====================================================

function verificarJWT(token) {
  if (!token) {
    return {
      valido: false,

      mensaje: "No se recibió el token.",
    };
  }

  try {
    const partes = token.toString().split(".");

    if (partes.length !== 3) {
      return {
        valido: false,

        mensaje: "Token inválido.",
      };
    }

    const contenido = partes[0] + "." + partes[1];

    const firmaEsperada = firmarJWT(contenido);

    if (firmaEsperada !== partes[2]) {
      return {
        valido: false,

        mensaje: "Firma del token inválida.",
      };
    }

    const payload = JSON.parse(base64UrlDecode(partes[1]));

    if (!payload.exp || Date.now() >= Number(payload.exp)) {
      return {
        valido: false,

        mensaje: "El token ha expirado.",
      };
    }

    return {
      valido: true,

      payload: payload,
    };
  } catch (error) {
    return {
      valido: false,

      mensaje: "Token inválido.",
    };
  }
}

// =====================================================
// FIRMAR JWT
// =====================================================

function firmarJWT(contenido) {
  const secreto = obtenerSecretoJWT();

  const firma = Utilities.computeHmacSha256Signature(
    contenido,
    secreto,
    Utilities.Charset.UTF_8,
  );

  return base64UrlEncodeBytes(firma);
}

// =====================================================
// OBTENER SECRETO JWT
// =====================================================

function obtenerSecretoJWT() {
  const propiedades = PropertiesService.getScriptProperties();

  let secreto = propiedades.getProperty(CLAVE_SECRETA_JWT);

  if (!secreto) {
    secreto =
      Utilities.getUuid() +
      "-" +
      Utilities.getUuid() +
      "-" +
      new Date().getTime();

    propiedades.setProperty(CLAVE_SECRETA_JWT, secreto);
  }

  return secreto;
}

// =====================================================
// BASE64 URL SAFE
// =====================================================

function base64UrlEncode(texto) {
  const bytes = Utilities.newBlob(texto).getBytes();

  return base64UrlEncodeBytes(bytes);
}

// =====================================================
// BASE64 URL SAFE PARA BYTES
// =====================================================

function base64UrlEncodeBytes(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, "");
}

// =====================================================
// DECODIFICAR BASE64 URL SAFE
// =====================================================

function base64UrlDecode(texto) {
  let base64 = texto.replace(/-/g, "+").replace(/_/g, "/");

  while (base64.length % 4) {
    base64 += "=";
  }

  const bytes = Utilities.base64Decode(base64);

  return Utilities.newBlob(bytes).getDataAsString();
}

// =====================================================
// VALIDAR SESIÓN
// =====================================================

function validarSesion(token) {
  const resultado = verificarJWT(token);

  if (!resultado.valido) {
    return {
      success: false,

      mensaje: resultado.mensaje,
    };
  }

  if (resultado.payload.tipo !== "SESION") {
    return {
      success: false,

      mensaje: "El token no corresponde a una sesión.",
    };
  }

  return {
    success: true,

    usuario: {
      id: resultado.payload.id,

      nombre: resultado.payload.nombre,

      usuario: resultado.payload.usuario,

      email: resultado.payload.email,

      rol: resultado.payload.rol,
    },
  };
}

// =====================================================
// RECUPERAR CONTRASEÑA
// =====================================================

function solicitarRecuperacion(email) {
  if (!email) {
    return {
      success: false,

      mensaje: "Debes ingresar tu correo.",
    };
  }

  email = email.toString().trim().toLowerCase();

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,

      mensaje: "No existen usuarios registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];

    const emailHoja = String(fila[3] || "")
      .trim()
      .toLowerCase();

    if (emailHoja !== email) {
      continue;
    }

    const estado = normalizarEstado(fila[6]);

    if (estado !== "ACTIVO") {
      return {
        success: false,

        mensaje: "Este usuario está inactivo.",
      };
    }

    // -------------------------------------------------
    // GENERAR CÓDIGO DE RECUPERACIÓN
    // -------------------------------------------------

    const codigo = Math.floor(100000 + Math.random() * 900000).toString();

    // -------------------------------------------------
    // GENERAR JWT DE RECUPERACIÓN
    // -------------------------------------------------

    const token = generarJWT(
      {
        id: String(fila[0]),

        usuario: String(fila[2]),

        email: email,

        tipo: "RECUPERACION",

        codigo: codigo,
      },

      DURACION_RECUPERACION_MS,
    );

    // -------------------------------------------------
    // GUARDAR INFORMACIÓN DEL TOKEN
    // -------------------------------------------------

    const propiedades = PropertiesService.getScriptProperties();

    const clave = "RECUP_" + generarHash(email);

    propiedades.setProperty(
      clave,

      JSON.stringify({
        token: token,

        fecha: Date.now(),
      }),
    );

    // -------------------------------------------------
    // ENVIAR CORREO
    // -------------------------------------------------

    const nombre = String(fila[1] || "Usuario");

    const asunto = "Recuperación de contraseña - Sistema de Vehículos";

    const cuerpo =
      "Hola " +
      nombre +
      ",\n\n" +
      "Hemos recibido una solicitud para cambiar " +
      "la contraseña de tu cuenta del Sistema de Vehículos.\n\n" +
      "Tu código de recuperación es:\n\n" +
      codigo +
      "\n\n" +
      "Este código tiene una vigencia de 15 minutos.\n\n" +
      "Si no solicitaste este cambio, puedes ignorar este mensaje.\n\n" +
      "Sistema de Vehículos";

    MailApp.sendEmail(email, asunto, cuerpo);

    return {
      success: true,

      mensaje: "Hemos enviado un código de recuperación a tu correo.",
    };
  }

  // -------------------------------------------------
  // POR SEGURIDAD NO REVELAMOS SI EL CORREO EXISTE
  // -------------------------------------------------

  return {
    success: true,

    mensaje:
      "Si el correo está registrado, recibirás un código de recuperación.",
  };
}

// =====================================================
// VALIDAR CÓDIGO DE RECUPERACIÓN
// =====================================================

function validarCodigoRecuperacion(email, codigo) {
  if (!email || !codigo) {
    return {
      success: false,

      mensaje: "Debes ingresar correo y código.",
    };
  }

  email = email.toString().trim().toLowerCase();

  codigo = codigo.toString().trim();

  const propiedades = PropertiesService.getScriptProperties();

  const clave = "RECUP_" + generarHash(email);

  const guardado = propiedades.getProperty(clave);

  if (!guardado) {
    return {
      success: false,

      mensaje: "El código no existe o ya fue utilizado.",
    };
  }

  try {
    const informacion = JSON.parse(guardado);

    const resultado = verificarJWT(informacion.token);

    if (!resultado.valido) {
      propiedades.deleteProperty(clave);

      return {
        success: false,

        mensaje: "El código ha expirado.",
      };
    }

    if (resultado.payload.tipo !== "RECUPERACION") {
      return {
        success: false,

        mensaje: "Token de recuperación inválido.",
      };
    }

    if (String(resultado.payload.codigo) !== codigo) {
      return {
        success: false,

        mensaje: "El código ingresado es incorrecto.",
      };
    }

    return {
      success: true,

      mensaje: "Código válido.",

      token: informacion.token,
    };
  } catch (error) {
    return {
      success: false,

      mensaje: "No fue posible validar el código.",
    };
  }
}

// =====================================================
// CAMBIAR CONTRASEÑA MEDIANTE RECUPERACIÓN
// =====================================================

function cambiarPasswordRecuperacion(token, nuevaContraseña) {
  if (!token || !nuevaContraseña) {
    return {
      success: false,

      mensaje: "Faltan datos para cambiar la contraseña.",
    };
  }

  nuevaContraseña = nuevaContraseña.toString();

  if (nuevaContraseña.length < 4) {
    return {
      success: false,

      mensaje: "La contraseña debe tener mínimo 4 caracteres.",
    };
  }

  const resultado = verificarJWT(token);

  if (!resultado.valido) {
    return {
      success: false,

      mensaje: "El código de recuperación ha expirado.",
    };
  }

  if (resultado.payload.tipo !== "RECUPERACION") {
    return {
      success: false,

      mensaje: "Token de recuperación inválido.",
    };
  }

  const email = String(resultado.payload.email).trim().toLowerCase();

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,

      mensaje: "No existen usuarios.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

  let actualizado = false;

  for (let i = 0; i < datos.length; i++) {
    const emailHoja = String(datos[i][3] || "")
      .trim()
      .toLowerCase();

    if (emailHoja !== email) {
      continue;
    }

    const estado = normalizarEstado(datos[i][6]);

    if (estado !== "ACTIVO") {
      return {
        success: false,

        mensaje: "El usuario está inactivo.",
      };
    }

    hoja.getRange(i + 2, 5).setValue(generarHash(nuevaContraseña));

    actualizado = true;

    break;
  }

  if (!actualizado) {
    return {
      success: false,

      mensaje: "No se encontró el usuario.",
    };
  }

  // ---------------------------------------------------
  // INVALIDAR TOKEN DE RECUPERACIÓN
  // ---------------------------------------------------

  const propiedades = PropertiesService.getScriptProperties();

  const clave = "RECUP_" + generarHash(email);

  propiedades.deleteProperty(clave);

  return {
    success: true,

    mensaje: "Contraseña actualizada correctamente.",
  };
}

// =====================================================
// CAMBIAR CONTRASEÑA DESDE SESIÓN
// =====================================================

function cambiarPasswordSesion(token, contraseñaActual, nuevaContraseña) {
  const sesion = validarSesion(token);

  if (!sesion.success) {
    return sesion;
  }

  if (!contraseñaActual || !nuevaContraseña) {
    return {
      success: false,

      mensaje: "Debes completar todos los campos.",
    };
  }

  if (nuevaContraseña.length < 4) {
    return {
      success: false,

      mensaje: "La nueva contraseña debe tener mínimo 4 caracteres.",
    };
  }

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

  for (let i = 0; i < datos.length; i++) {
    const usuarioHoja = normalizarUsuario(datos[i][2]);

    if (usuarioHoja !== normalizarUsuario(sesion.usuario.usuario)) {
      continue;
    }

    if (!comprobarPassword(contraseñaActual, datos[i][4])) {
      return {
        success: false,

        mensaje: "La contraseña actual es incorrecta.",
      };
    }

    hoja.getRange(i + 2, 5).setValue(generarHash(nuevaContraseña));

    return {
      success: true,

      mensaje: "Contraseña actualizada correctamente.",
    };
  }

  return {
    success: false,

    mensaje: "No se encontró el usuario.",
  };
}

// =====================================================
// OBTENER PERFIL DEL USUARIO
// =====================================================

function obtenerPerfilUsuario(token) {
  const sesion = validarSesion(token);

  if (!sesion.success) {
    return sesion;
  }

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,

      mensaje: "No existen usuarios.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getDisplayValues();

  const usuarioBuscado = normalizarUsuario(sesion.usuario.usuario);

  for (let i = 0; i < datos.length; i++) {
    if (normalizarUsuario(datos[i][2]) !== usuarioBuscado) {
      continue;
    }

    return {
      success: true,

      usuario: {
        id: datos[i][0],

        nombre: datos[i][1],

        usuario: datos[i][2],

        email: datos[i][3],

        rol: normalizarRol(datos[i][5]),

        estado: normalizarEstado(datos[i][6]),
      },
    };
  }

  return {
    success: false,

    mensaje: "No se encontró el perfil.",
  };
}

function actualizarPerfilUsuario(
  token,
  nombre,
  usuario,
  email,
  nuevaContraseña,
) {
  const sesion = validarSesion(token);

  if (!sesion.success) {
    return sesion;
  }

  nombre = nombre ? nombre.toString().trim() : "";

  usuario = normalizarUsuario(usuario);

  email = email ? email.toString().trim().toLowerCase() : "";

  nuevaContraseña = nuevaContraseña ? nuevaContraseña.toString() : "";

  if (!nombre || !usuario || !email) {
    return {
      success: false,

      mensaje: "Nombre, usuario y correo son obligatorios.",
    };
  }

  if (nuevaContraseña && nuevaContraseña.length < 4) {
    return {
      success: false,

      mensaje: "La contraseña debe tener mínimo 4 caracteres.",
    };
  }

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

  const usuarioBuscado = normalizarUsuario(sesion.usuario.usuario);

  // Verificar que el nuevo usuario
  // no pertenezca a otra cuenta.
  for (let i = 0; i < datos.length; i++) {
    const usuarioFila = normalizarUsuario(datos[i][2]);

    if (usuarioFila === usuario && usuarioFila !== usuarioBuscado) {
      return {
        success: false,

        mensaje: "Ese nombre de usuario ya está en uso.",
      };
    }
  }

  for (let i = 0; i < datos.length; i++) {
    if (normalizarUsuario(datos[i][2]) !== usuarioBuscado) {
      continue;
    }

    // Nombre
    hoja.getRange(i + 2, 2).setValue(nombre);

    // Usuario
    hoja.getRange(i + 2, 3).setValue(usuario);

    // Email
    hoja.getRange(i + 2, 4).setValue(email);

    // Contraseña
    // Solo se cambia si el usuario escribió una nueva.
    if (nuevaContraseña) {
      hoja.getRange(i + 2, 5).setValue(generarHash(nuevaContraseña));
    }

    // Generar una nueva sesión porque el
    // nombre de usuario pudo haber cambiado.
    const nuevoToken = generarJWT(
      {
        id: String(datos[i][0]),

        nombre: nombre,

        usuario: usuario,

        email: email,

        rol: sesion.usuario.rol,

        tipo: "SESION",

        exp: Date.now() + DURACION_SESION_MS,
      },
      DURACION_SESION_MS,
    );

    return {
      success: true,

      mensaje: nuevaContraseña
        ? "Perfil, usuario y contraseña actualizados correctamente."
        : "Perfil y usuario actualizados correctamente.",

      token: nuevoToken,

      usuario: {
        id: String(datos[i][0]),

        nombre: nombre,

        usuario: usuario,

        email: email,

        rol: sesion.usuario.rol,
      },
    };
  }

  return {
    success: false,

    mensaje: "No se encontró el usuario.",
  };
}

// =====================================================
// VALIDAR ROL
// =====================================================

function validarRol(token, rolPermitido) {
  const sesion = validarSesion(token);

  if (!sesion.success) {
    return sesion;
  }

  const rolUsuario = normalizarRol(sesion.usuario.rol);

  const rolNecesario = normalizarRol(rolPermitido);

  if (rolUsuario !== rolNecesario) {
    return {
      success: false,

      mensaje: "No tienes permisos para realizar esta acción.",
    };
  }

  return {
    success: true,

    usuario: sesion.usuario,
  };
}

// =====================================================
// VALIDAR ADMINISTRADOR
// =====================================================

function validarAdministrador(token) {
  return validarRol(token, "ADMINISTRADOR");
}

// =====================================================
// VALIDAR SUPERVISOR O ADMINISTRADOR
// =====================================================

function validarSupervisorOAdministrador(token) {
  const sesion = validarSesion(token);

  if (!sesion.success) {
    return sesion;
  }

  const rol = normalizarRol(sesion.usuario.rol);

  if (rol !== "SUPERVISOR" && rol !== "ADMINISTRADOR") {
    return {
      success: false,

      mensaje: "No tienes permisos para realizar esta acción.",
    };
  }

  return {
    success: true,

    usuario: sesion.usuario,
  };
}

// =====================================================
// ADMINISTRACIÓN DE UBICACIÓN Y PUESTO
// =====================================================

const COLUMNA_UBICACION = 9;
const COLUMNA_PUESTO = 10;
const COLUMNA_TRANSITO = 11;

// =====================================================
// OBTENER PARQUEADEROS
// SOLO ADMINISTRADOR
// =====================================================

function obtenerParqueaderosAdministracion(token) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_PARQUEADEROS);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "PARQUEADEROS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: true,
      parqueaderos: [],
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 2).getDisplayValues();

  const parqueaderos = datos
    .map(function (fila) {
      return {
        nombre: fila[0].toString().trim(),

        estado: fila[1].toString().trim(),
      };
    })
    .filter(function (parqueadero) {
      return parqueadero.nombre !== "";
    });

  return {
    success: true,
    parqueaderos: parqueaderos,
  };
}

// =====================================================
// CREAR PARQUEADERO
// SOLO ADMINISTRADOR
// =====================================================

function crearParqueadero(token, nombreParqueadero) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const nombre = nombreParqueadero.toString().trim();

  if (!nombre) {
    return {
      success: false,
      mensaje: "Debes ingresar el nombre del parqueadero.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_PARQUEADEROS);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "PARQUEADEROS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila >= 2) {
    const nombresExistentes = hoja
      .getRange(2, 1, ultimaFila - 1, 1)
      .getDisplayValues();

    const existe = nombresExistentes.some(function (fila) {
      return fila[0].toString().trim().toLowerCase() === nombre.toLowerCase();
    });

    if (existe) {
      return {
        success: false,
        mensaje: "Ya existe un parqueadero con ese nombre.",
      };
    }
  }

  hoja.appendRow([nombre, "ACTIVO"]);

  return {
    success: true,
    mensaje: "Parqueadero creado correctamente.",
  };
}

// =====================================================
// ACTUALIZAR PARQUEADERO
// SOLO ADMINISTRADOR
// =====================================================

function actualizarParqueadero(token, nombreAnterior, nombreNuevo) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const anterior = nombreAnterior.toString().trim();

  const nuevo = nombreNuevo.toString().trim();

  if (!anterior || !nuevo) {
    return {
      success: false,
      mensaje: "Debes indicar el parqueadero actual y el nuevo nombre.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_PARQUEADEROS);

  const hojaVehiculos = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "PARQUEADEROS".',
    };
  }

  if (!hojaVehiculos) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No hay parqueaderos registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 2).getDisplayValues();

  let filaEncontrada = -1;

  for (let i = 0; i < datos.length; i++) {
    if (
      datos[i][0].toString().trim().toLowerCase() === anterior.toLowerCase()
    ) {
      filaEncontrada = i + 2;

      break;
    }
  }

  if (filaEncontrada === -1) {
    return {
      success: false,
      mensaje: "No se encontró el parqueadero seleccionado.",
    };
  }

  const existeOtro = datos.some(function (fila, indice) {
    if (indice + 2 === filaEncontrada) {
      return false;
    }

    return fila[0].toString().trim().toLowerCase() === nuevo.toLowerCase();
  });

  if (existeOtro) {
    return {
      success: false,
      mensaje: "Ya existe otro parqueadero con ese nombre.",
    };
  }

  // =====================================================
  // ACTUALIZAR NOMBRE EN PARQUEADEROS
  // =====================================================

  hoja.getRange(filaEncontrada, 1).setValue(nuevo);

  // =====================================================
  // ACTUALIZAR UBICACIÓN DE LOS VEHÍCULOS
  // =====================================================

  const ultimaFilaVehiculos = hojaVehiculos.getLastRow();

  let vehiculosActualizados = 0;

  if (ultimaFilaVehiculos >= 2) {
    const rangoUbicaciones = hojaVehiculos.getRange(
      2,
      COLUMNA_UBICACION,
      ultimaFilaVehiculos - 1,
      1,
    );

    const ubicaciones = rangoUbicaciones.getValues();

    for (let i = 0; i < ubicaciones.length; i++) {
      const ubicacion = ubicaciones[i][0].toString().trim();

      if (ubicacion.toLowerCase() === anterior.toLowerCase()) {
        ubicaciones[i][0] = nuevo;

        vehiculosActualizados++;
      }
    }

    if (vehiculosActualizados > 0) {
      rangoUbicaciones.setValues(ubicaciones);
    }
  }

  return {
    success: true,
    mensaje:
      vehiculosActualizados > 0
        ? "Parqueadero actualizado correctamente. Se actualizaron " +
          vehiculosActualizados +
          " vehículo(s)."
        : "Parqueadero actualizado correctamente.",
  };
}

// =====================================================
// ELIMINAR PARQUEADERO
// SOLO ADMINISTRADOR
// =====================================================

function eliminarParqueadero(token, nombreParqueadero) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const nombre = nombreParqueadero.toString().trim();

  if (!nombre) {
    return {
      success: false,
      mensaje: "Debes indicar el parqueadero que deseas eliminar.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hojaParqueaderos = archivo.getSheetByName(NOMBRE_HOJA_PARQUEADEROS);

  const hojaVehiculos = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hojaParqueaderos) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "PARQUEADEROS".',
    };
  }

  if (!hojaVehiculos) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFilaParqueaderos = hojaParqueaderos.getLastRow();

  if (ultimaFilaParqueaderos < 2) {
    return {
      success: false,
      mensaje: "No hay parqueaderos registrados.",
    };
  }

  const datosParqueaderos = hojaParqueaderos
    .getRange(2, 1, ultimaFilaParqueaderos - 1, 2)
    .getDisplayValues();

  let filaEncontrada = -1;

  for (let i = 0; i < datosParqueaderos.length; i++) {
    if (
      datosParqueaderos[i][0].toString().trim().toLowerCase() ===
      nombre.toLowerCase()
    ) {
      filaEncontrada = i + 2;

      break;
    }
  }

  if (filaEncontrada === -1) {
    return {
      success: false,
      mensaje: "No se encontró el parqueadero seleccionado.",
    };
  }

  const ultimaFilaVehiculos = hojaVehiculos.getLastRow();

  if (ultimaFilaVehiculos >= 2) {
    const ubicaciones = hojaVehiculos
      .getRange(2, COLUMNA_UBICACION, ultimaFilaVehiculos - 1, 1)
      .getDisplayValues();

    let vehiculosAsignados = 0;

    ubicaciones.forEach(function (fila) {
      if (fila[0].toString().trim().toLowerCase() === nombre.toLowerCase()) {
        vehiculosAsignados++;
      }
    });

    if (vehiculosAsignados > 0) {
      return {
        success: false,
        mensaje:
          "No se puede eliminar este parqueadero porque tiene " +
          vehiculosAsignados +
          " vehículo(s) asignado(s).",
      };
    }
  }

  hojaParqueaderos.deleteRow(filaEncontrada);

  return {
    success: true,
    mensaje: "Parqueadero eliminado correctamente.",
  };
}

// =====================================================
// ACTUALIZAR INFORMACIÓN DEL VEHÍCULO
// SOLO ADMINISTRADOR
// =====================================================

function actualizarVehiculoAdministracion(
  token,
  placaOriginal,
  nombre,
  placa,
  transito,
  soat,
  tecnomecanica,
  picoYPlaca,
  precio,
  kilometraje,
  detalles,
  ubicacion,
) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  placaOriginal = placaOriginal ? placaOriginal.toString().trim() : "";

  nombre = nombre ? nombre.toString().trim().toUpperCase() : "";

  placa = placa ? placa.toString().trim().toUpperCase() : "";

  transito = transito ? transito.toString().trim().toUpperCase() : "";

  soat = soat ? soat.toString().trim() : "";

  tecnomecanica = tecnomecanica ? tecnomecanica.toString().trim() : "";

  picoYPlaca = picoYPlaca ? picoYPlaca.toString().trim() : "";

  precio =
    precio !== undefined && precio !== null ? precio.toString().trim() : "";

  kilometraje =
    kilometraje !== undefined && kilometraje !== null
      ? kilometraje.toString().trim()
      : "";

  detalles =
    detalles !== undefined && detalles !== null
      ? detalles.toString().trim().toUpperCase()
      : "";

  ubicacion =
    ubicacion !== undefined && ubicacion !== null
      ? ubicacion.toString().trim()
      : null;

  if (!placaOriginal) {
    return {
      success: false,
      mensaje: "No se recibió la placa original del vehículo.",
    };
  }

  if (!nombre) {
    return {
      success: false,
      mensaje: "El nombre del vehículo es obligatorio.",
    };
  }

  if (!placa) {
    return {
      success: false,
      mensaje: "La placa del vehículo es obligatoria.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No existen vehículos registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 11).getDisplayValues();

  const placaOriginalNormalizada = placaOriginal.trim().toUpperCase();

  // Buscar vehículo original
  for (let i = 0; i < datos.length; i++) {
    const placaHoja = datos[i][1]
      ? datos[i][1].toString().trim().toUpperCase()
      : "";

    if (placaHoja !== placaOriginalNormalizada) {
      continue;
    }

    const numeroFila = i + 2;

    const nombreAnterior = datos[i][0] || "";

    const placaAnterior = placaHoja;

    // Verificar que la nueva placa
    // no pertenezca a otro vehículo
    for (let j = 0; j < datos.length; j++) {
      if (j === i) {
        continue;
      }

      const otraPlaca = datos[j][1]
        ? datos[j][1].toString().trim().toUpperCase()
        : "";

      if (otraPlaca && otraPlaca === placa) {
        return {
          success: false,
          mensaje: "La placa indicada ya pertenece a otro vehículo.",
        };
      }
    }

    // Convertir fechas provenientes
    // del formulario HTML
    let fechaSoat = soat;
    let fechaTecnomecanica = tecnomecanica;

    if (/^\d{4}-\d{2}-\d{2}$/.test(soat)) {
      const partes = soat.split("-");

      fechaSoat = new Date(
        Number(partes[0]),
        Number(partes[1]) - 1,
        Number(partes[2]),
      );
    }

    if (/^\d{4}-\d{2}-\d{2}$/.test(tecnomecanica)) {
      const partes = tecnomecanica.split("-");

      fechaTecnomecanica = new Date(
        Number(partes[0]),
        Number(partes[1]) - 1,
        Number(partes[2]),
      );
    }

    // Precio y kilometraje
    let precioGuardar = precio;

    if (precio !== "") {
      const numeroPrecio = Number(
        precio.toString().replace(/\./g, "").replace(/,/g, ""),
      );

      if (!isNaN(numeroPrecio)) {
        precioGuardar = numeroPrecio;
      }
    }

    let kilometrajeGuardar = kilometraje;

    if (kilometraje !== "") {
      const numeroKilometraje = Number(
        kilometraje.toString().replace(/\./g, "").replace(/,/g, ""),
      );

      if (!isNaN(numeroKilometraje)) {
        kilometrajeGuardar = numeroKilometraje;
      }
    }

    // Actualizar columnas A-H
    hoja
      .getRange(numeroFila, 1, 1, 8)
      .setValues([
        [
          nombre,
          placa,
          fechaSoat,
          fechaTecnomecanica,
          picoYPlaca,
          precioGuardar,
          kilometrajeGuardar,
          detalles,
        ],
      ]);

    // Actualizar tránsito en columna K
    hoja.getRange(numeroFila, COLUMNA_TRANSITO).setValue(transito);

    if (ubicacion !== null) {
      const ubicacionAnterior = datos[i][8]
        ? datos[i][8].toString().trim()
        : "";

      if (ubicacionAnterior !== ubicacion) {
        hoja.getRange(numeroFila, COLUMNA_UBICACION).setValue(ubicacion);

        registrarMovimientoVehiculo({
          usuario: permiso.usuario.nombre,

          usuarioSistema: permiso.usuario.usuario,

          placa: placa,

          vehiculo: nombre,

          ubicacionAnterior: ubicacionAnterior,

          ubicacionNueva: ubicacion,

          tipoMovimiento: "UBICACIÓN",

          campo: "UBICACIÓN",

          valorAnterior: ubicacionAnterior,

          valorNuevo: ubicacion,

          accion: "CAMBIO DE UBICACIÓN",
        });
      }
    }

    // Registrar acción en historial
    // Registrar únicamente los campos que realmente cambiaron

    // NOMBRE
    if (nombreAnterior.toString().trim() !== nombre.toString().trim()) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "NOMBRE",

        valorAnterior: nombreAnterior,

        valorNuevo: nombre,

        tipoMovimiento: "INFORMACIÓN DEL VEHÍCULO",

        accion: "CAMBIO DE NOMBRE",
      });
    }

    // PLACA
    if (
      placaAnterior.toString().trim().toUpperCase() !==
      placa.toString().trim().toUpperCase()
    ) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "PLACA",

        valorAnterior: placaAnterior,

        valorNuevo: placa,

        tipoMovimiento: "INFORMACIÓN DEL VEHÍCULO",

        accion: "CAMBIO DE PLACA",
      });
    }

    // SOAT
    if (
      normalizarFechaHistorial(soat) !== normalizarFechaHistorial(datos[i][2])
    ) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "SOAT",

        valorAnterior: datos[i][2] || "",

        valorNuevo: soat,

        tipoMovimiento: "SOAT",

        accion: "CAMBIO DE SOAT",
      });
    }

    // TECNOMECÁNICA
    if (
      normalizarFechaHistorial(tecnomecanica) !==
      normalizarFechaHistorial(datos[i][3])
    ) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "TECNOMECÁNICA",

        valorAnterior: datos[i][3] || "",

        valorNuevo: tecnomecanica,

        tipoMovimiento: "TECNOMECÁNICA",

        accion: "CAMBIO DE TECNOMECÁNICA",
      });
    }

    // PICO Y PLACA
    if (picoYPlaca.toString().trim() !== datos[i][4].toString().trim()) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "PICO Y PLACA",

        valorAnterior: datos[i][4] || "",

        valorNuevo: picoYPlaca,

        tipoMovimiento: "PICO Y PLACA",

        accion: "CAMBIO DE PICO Y PLACA",
      });
    }

    // PRECIO
    const precioAnterior = datos[i][5]
      ? datos[i][5].toString().replace(/\./g, "").replace(/,/g, "").trim()
      : "";

    const precioNuevo =
      precioGuardar !== undefined && precioGuardar !== null
        ? precioGuardar.toString().trim()
        : "";

    if (precioAnterior !== precioNuevo) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "PRECIO",

        valorAnterior: datos[i][5] || "",

        valorNuevo: precioGuardar,

        tipoMovimiento: "PRECIO",

        accion: "CAMBIO DE PRECIO",
      });
    }

    // KILOMETRAJE
    const kilometrajeAnterior = datos[i][6]
      ? datos[i][6].toString().replace(/\./g, "").replace(/,/g, "").trim()
      : "";

    const kilometrajeNuevo =
      kilometrajeGuardar !== undefined && kilometrajeGuardar !== null
        ? kilometrajeGuardar.toString().trim()
        : "";

    if (kilometrajeAnterior !== kilometrajeNuevo) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "KILOMETRAJE",

        valorAnterior: datos[i][6] || "",

        valorNuevo: kilometrajeGuardar,

        tipoMovimiento: "KILOMETRAJE",

        accion: "CAMBIO DE KILOMETRAJE",
      });
    }

    // DETALLES
    if (datos[i][7].toString().trim() !== detalles.toString().trim()) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placa,

        vehiculo: nombre,

        campo: "DETALLES",

        valorAnterior: datos[i][7] || "",

        valorNuevo: detalles,

        tipoMovimiento: "INFORMACIÓN DEL VEHÍCULO",

        accion: "CAMBIO DE DETALLES",
      });
    }

    return {
      success: true,

      mensaje: "La información del vehículo fue actualizada correctamente.",

      vehiculo: {
        fila: numeroFila,

        nombre: nombre,

        placa: placa,

        soat: soat,

        tecnomecanica: tecnomecanica,

        picoYPlaca: picoYPlaca,

        precio: precioGuardar,

        kilometraje: kilometrajeGuardar,

        detalles: detalles,

        ubicacion: datos[i][8] || "",

        puesto: datos[i][9] || "",

        transito: datos[i][10] || "",
      },
    };
  }

  return {
    success: false,
    mensaje: "No se encontró el vehículo.",
  };
}

// =====================================================
// ELIMINAR VEHÍCULO
// SOLO ADMINISTRADOR
// =====================================================

function eliminarVehiculoAdministracion(token, fila, placa) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  fila = Number(fila);

  placa = placa ? placa.toString().trim().toUpperCase() : "";

  if (!Number.isInteger(fila) || fila < 2) {
    return {
      success: false,
      mensaje: "La fila del vehículo no es válida.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (fila > ultimaFila) {
    return {
      success: false,
      mensaje: "El vehículo ya no existe o cambió de ubicación.",
    };
  }

  // Leer A:L antes de eliminar
  const datosFila = hoja.getRange(fila, 1, 1, 13).getDisplayValues()[0];

  const filaVacia = datosFila.every(function (valor) {
    return valor.toString().trim() === "";
  });

  if (filaVacia) {
    return {
      success: false,
      mensaje: "El vehículo ya fue eliminado.",
    };
  }

  const placaActual = datosFila[1]
    ? datosFila[1].toString().trim().toUpperCase()
    : "";

  /*
    Si el vehículo tiene placa,
    comprobamos que la fila todavía
    corresponda al mismo vehículo.
  */
  if (placa && placaActual !== placa) {
    return {
      success: false,
      mensaje:
        "El vehículo seleccionado cambió. Actualiza el inventario e inténtalo nuevamente.",
    };
  }

  // =====================================================
  // REGISTRAR ELIMINACIÓN EN HISTORIAL
  // =====================================================

  registrarMovimientoVehiculo({
    usuario: permiso.usuario.nombre,

    usuarioSistema: permiso.usuario.usuario,

    vehiculo: datosFila[0] || "",

    placa: datosFila[1] || "",

    campo: "VEHÍCULO",

    valorAnterior: datosFila[0] || datosFila[1] || "Vehículo",

    valorNuevo: "",

    tipoMovimiento: "VEHÍCULO",

    accion: "ELIMINÓ VEHÍCULO",
  });

  // Limpiar A:L sin eliminar físicamente la fila
  hoja.getRange(fila, 1, 1, 13).clearContent();

  return {
    success: true,
    mensaje: "Vehículo eliminado correctamente.",
  };
}

// =====================================================
// OBTENER VEHÍCULOS PARA EL PANEL ADMINISTRATIVO
// SUPERVISOR + ADMINISTRADOR
// =====================================================

function obtenerVehiculosAdministracion(token) {
  const permiso = validarSupervisorOAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: true,
      vehiculos: [],
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 11).getDisplayValues();

  const vehiculos = [];

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];

    const nombre = fila[0] ? fila[0].toString().trim() : "";

    const placa = fila[1] ? fila[1].toString().trim() : "";

    // Ignorar filas completamente vacías
    if (!nombre && !placa) {
      continue;
    }

    vehiculos.push({
      fila: i + 2,

      nombre: nombre,

      placa: placa,

      soat: fila[2] || "",

      tecnomecanica: fila[3] || "",

      picoYPlaca: fila[4] || "",

      precio: fila[5] || "",

      kilometraje: fila[6] || "",

      detalles: fila[7] || "",

      ubicacion: fila[8] ? fila[8].toString().trim() : "",

      puesto: fila[9] ? fila[9].toString().trim() : "",

      transito: fila[10] ? fila[10].toString().trim() : "",
    });
  }

  return {
    success: true,
    vehiculos: vehiculos,
  };
}

// =====================================================
// ACTUALIZAR UBICACIÓN Y PUESTO
// SOLO ADMINISTRADOR
// =====================================================

function actualizarUbicacionVehiculo(
  token,
  placa,
  nuevaUbicacion,
  nuevoPuesto,
) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  placa = placa ? placa.toString().trim() : "";

  nuevaUbicacion = nuevaUbicacion ? nuevaUbicacion.toString().trim() : "";

  nuevoPuesto = nuevoPuesto ? nuevoPuesto.toString().trim() : "";

  if (!placa) {
    return {
      success: false,
      mensaje: "No se recibió la placa del vehículo.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No existen vehículos registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 10).getDisplayValues();

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];

    const placaHoja = fila[1] ? fila[1].toString().trim().toUpperCase() : "";

    if (placaHoja !== placa.toUpperCase()) {
      continue;
    }

    const numeroFila = i + 2;

    const ubicacionAnterior = fila[8] ? fila[8].toString().trim() : "";

    const puestoAnterior = fila[9] ? fila[9].toString().trim() : "";

    // Actualizar ubicación
    hoja.getRange(numeroFila, COLUMNA_UBICACION).setValue(nuevaUbicacion);

    // Actualizar puesto
    hoja.getRange(numeroFila, COLUMNA_PUESTO).setValue(nuevoPuesto);

    // ---------------------------------------------------
    // REGISTRAR CAMBIO DE UBICACIÓN
    // ---------------------------------------------------

    if (ubicacionAnterior !== nuevaUbicacion) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placaHoja,

        vehiculo: fila[0],

        ubicacionAnterior: ubicacionAnterior,

        ubicacionNueva: nuevaUbicacion,

        puestoAnterior: puestoAnterior,

        puestoNuevo: puestoAnterior,

        tipoMovimiento: "UBICACIÓN",

        campo: "UBICACIÓN",

        valorAnterior: ubicacionAnterior,

        valorNuevo: nuevaUbicacion,

        accion: "CAMBIO DE UBICACIÓN",
      });
    }

    // ---------------------------------------------------
    // REGISTRAR CAMBIO DE PUESTO / PARQUEADERO
    // ---------------------------------------------------

    if (puestoAnterior !== nuevoPuesto) {
      registrarMovimientoVehiculo({
        usuario: permiso.usuario.nombre,

        usuarioSistema: permiso.usuario.usuario,

        placa: placaHoja,

        vehiculo: fila[0],

        ubicacionAnterior: ubicacionAnterior,

        ubicacionNueva: nuevaUbicacion,

        puestoAnterior: puestoAnterior,

        puestoNuevo: nuevoPuesto,

        tipoMovimiento: "PUESTO / PARQUEADERO",

        campo: "PUESTO",

        valorAnterior: puestoAnterior,

        valorNuevo: nuevoPuesto,

        accion: "CAMBIO DE PUESTO",
      });
    }

    return {
      success: true,

      mensaje: "Ubicación actualizada correctamente.",

      vehiculo: {
        nombre: fila[0],

        placa: placaHoja,

        ubicacion: nuevaUbicacion,

        puesto: nuevoPuesto,
      },
    };
  }

  return {
    success: false,
    mensaje: "No se encontró el vehículo.",
  };
}

// =====================================================
// INTERCAMBIAR PARQUEADERO ENTRE DOS VEHÍCULOS
// SOLO ADMINISTRADOR
// =====================================================

function intercambiarParqueaderoVehiculos(
  token,
  placaVehiculo1,
  placaVehiculo2,
) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  placaVehiculo1 = placaVehiculo1
    ? placaVehiculo1.toString().trim().toUpperCase()
    : "";

  placaVehiculo2 = placaVehiculo2
    ? placaVehiculo2.toString().trim().toUpperCase()
    : "";

  if (!placaVehiculo1 || !placaVehiculo2) {
    return {
      success: false,
      mensaje: "Debes seleccionar dos vehículos.",
    };
  }

  if (placaVehiculo1 === placaVehiculo2) {
    return {
      success: false,
      mensaje: "No puedes intercambiar un vehículo consigo mismo.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No existen vehículos registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 10).getDisplayValues();

  let vehiculo1 = null;
  let vehiculo2 = null;

  for (let i = 0; i < datos.length; i++) {
    const placa = datos[i][1]
      ? datos[i][1].toString().trim().toUpperCase()
      : "";

    if (placa === placaVehiculo1) {
      vehiculo1 = {
        fila: i + 2,
        nombre: datos[i][0],
        placa: placa,
        ubicacion: datos[i][8] || "",
        puesto: datos[i][9] || "",
      };
    }

    if (placa === placaVehiculo2) {
      vehiculo2 = {
        fila: i + 2,
        nombre: datos[i][0],
        placa: placa,
        ubicacion: datos[i][8] || "",
        puesto: datos[i][9] || "",
      };
    }
  }

  if (!vehiculo1) {
    return {
      success: false,
      mensaje: "No se encontró el primer vehículo.",
    };
  }

  if (!vehiculo2) {
    return {
      success: false,
      mensaje: "No se encontró el segundo vehículo.",
    };
  }

  // ---------------------------------------------------
  // GUARDAR UBICACIONES ORIGINALES
  // ---------------------------------------------------

  const ubicacion1 = vehiculo1.ubicacion;

  const puesto1 = vehiculo1.puesto;

  const ubicacion2 = vehiculo2.ubicacion;

  const puesto2 = vehiculo2.puesto;

  // ---------------------------------------------------
  // REALIZAR INTERCAMBIO
  // ---------------------------------------------------

  hoja.getRange(vehiculo1.fila, COLUMNA_UBICACION).setValue(ubicacion2);

  hoja.getRange(vehiculo1.fila, COLUMNA_PUESTO).setValue(puesto2);

  hoja.getRange(vehiculo2.fila, COLUMNA_UBICACION).setValue(ubicacion1);

  hoja.getRange(vehiculo2.fila, COLUMNA_PUESTO).setValue(puesto1);

  // ---------------------------------------------------
  // REGISTRAR VEHÍCULO 1 EN HISTORIAL
  // ---------------------------------------------------

  registrarMovimientoVehiculo({
    usuario: permiso.usuario.nombre,

    usuarioSistema: permiso.usuario.usuario,

    vehiculo: vehiculo1.nombre,

    placa: vehiculo1.placa,

    ubicacionAnterior: ubicacion1,

    ubicacionNueva: ubicacion2,

    puestoAnterior: puesto1,

    puestoNuevo: puesto2,

    tipoMovimiento: "INTERCAMBIO DE PARQUEADERO",

    campo: "UBICACIÓN / PUESTO",

    valorAnterior: ubicacion1 + (puesto1 ? " / " + puesto1 : ""),

    valorNuevo: ubicacion2 + (puesto2 ? " / " + puesto2 : ""),

    accion: "INTERCAMBIO DE PARQUEADERO",
  });

  // ---------------------------------------------------
  // REGISTRAR VEHÍCULO 2 EN HISTORIAL
  // ---------------------------------------------------

  registrarMovimientoVehiculo({
    usuario: permiso.usuario.nombre,

    usuarioSistema: permiso.usuario.usuario,

    vehiculo: vehiculo2.nombre,

    placa: vehiculo2.placa,

    ubicacionAnterior: ubicacion2,

    ubicacionNueva: ubicacion1,

    puestoAnterior: puesto2,

    puestoNuevo: puesto1,

    tipoMovimiento: "INTERCAMBIO DE PARQUEADERO",

    campo: "UBICACIÓN / PUESTO",

    valorAnterior: ubicacion2 + (puesto2 ? " / " + puesto2 : ""),

    valorNuevo: ubicacion1 + (puesto1 ? " / " + puesto1 : ""),

    accion: "INTERCAMBIO DE PARQUEADERO",
  });

  return {
    success: true,

    mensaje: "Parqueaderos intercambiados correctamente.",

    vehiculo1: {
      nombre: vehiculo1.nombre,
      placa: vehiculo1.placa,
      ubicacion: ubicacion2,
      puesto: puesto2,
    },

    vehiculo2: {
      nombre: vehiculo2.nombre,
      placa: vehiculo2.placa,
      ubicacion: ubicacion1,
      puesto: puesto1,
    },
  };
}

// =====================================================
// REGISTRAR MOVIMIENTO DE VEHÍCULO
// =====================================================

function registrarMovimientoVehiculo(movimiento) {
  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  let hoja = archivo.getSheetByName("HISTORIAL_MOVIMIENTOS");

  // ---------------------------------------------------
  // CREAR HOJA SI NO EXISTE
  // ---------------------------------------------------

  if (!hoja) {
    hoja = archivo.insertSheet("HISTORIAL_MOVIMIENTOS");

    hoja.appendRow([
      "FECHA",
      "USUARIO",
      "USUARIO SISTEMA",
      "VEHÍCULO",
      "PLACA",
      "UBICACIÓN ANTERIOR",
      "UBICACIÓN NUEVA",
      "PUESTO ANTERIOR",
      "PUESTO NUEVO",
      "ACCIÓN",
      "TIPO DE MOVIMIENTO",
      "CAMPO",
      "VALOR ANTERIOR",
      "VALOR NUEVO",
    ]);
  }

  // ---------------------------------------------------
  // COMPATIBILIDAD CON LA ESTRUCTURA ACTUAL
  // ---------------------------------------------------

  hoja.appendRow([
    new Date(),

    movimiento.usuario || "",

    movimiento.usuarioSistema || "",

    movimiento.vehiculo || "",

    movimiento.placa || "",

    movimiento.ubicacionAnterior || "",

    movimiento.ubicacionNueva || "",

    movimiento.puestoAnterior || "",

    movimiento.puestoNuevo || "",

    movimiento.accion || "MOVIMIENTO",

    movimiento.tipoMovimiento || "",

    movimiento.campo || "",

    movimiento.valorAnterior !== undefined ? movimiento.valorAnterior : "",

    movimiento.valorNuevo !== undefined ? movimiento.valorNuevo : "",
  ]);
}

// =====================================================
// ADMINISTRACIÓN DE USUARIOS
// SOLO ADMINISTRADOR
// =====================================================

// =====================================================
// OBTENER USUARIOS PARA ADMINISTRACIÓN
// =====================================================

function obtenerUsuariosAdministracion(token) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: true,
      usuarios: [],
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

  const usuarios = [];

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];

    if (!fila[0] && !fila[1] && !fila[2]) {
      continue;
    }

    usuarios.push({
      id: String(fila[0] || ""),

      nombre: String(fila[1] || ""),

      usuario: String(fila[2] || ""),

      email: String(fila[3] || ""),

      rol: normalizarRol(fila[5]),

      estado: normalizarEstado(fila[6]),
    });
  }

  return {
    success: true,

    usuarios: usuarios,
  };
}

// =====================================================
// CREAR USUARIO
// SOLO ADMINISTRADOR
// =====================================================

function crearUsuarioAdministracion(
  token,
  nombre,
  usuario,
  email,
  contraseña,
  rol,
  estado,
) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  nombre = nombre ? nombre.toString().trim() : "";

  usuario = normalizarUsuario(usuario);

  email = email ? email.toString().trim().toLowerCase() : "";

  contraseña =
    contraseña !== undefined && contraseña !== null
      ? contraseña.toString()
      : "";

  rol = normalizarRol(rol);

  estado = normalizarEstado(estado);

  if (!nombre || !usuario || !email || !contraseña) {
    return {
      success: false,

      mensaje: "Completa todos los campos obligatorios.",
    };
  }

  if (contraseña.length < 4) {
    return {
      success: false,

      mensaje: "La contraseña debe tener mínimo 4 caracteres.",
    };
  }

  if (!rol) {
    rol = "USUARIO";
  }

  if (!estado) {
    estado = "ACTIVO";
  }

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  // ---------------------------------------------------
  // VERIFICAR USUARIO DUPLICADO
  // ---------------------------------------------------

  if (ultimaFila >= 2) {
    const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

    for (let i = 0; i < datos.length; i++) {
      const usuarioExistente = normalizarUsuario(datos[i][2]);

      if (usuarioExistente === usuario) {
        return {
          success: false,

          mensaje: "El nombre de usuario ya está registrado.",
        };
      }
    }
  }

  // ---------------------------------------------------
  // CREAR ID Y CONTRASEÑA
  // ---------------------------------------------------

  const idUsuario = Utilities.getUuid();

  const contraseñaHash = generarHash(contraseña);

  // ---------------------------------------------------
  // GUARDAR USUARIO
  // ---------------------------------------------------

  hoja.appendRow([
    idUsuario,

    nombre,

    usuario,

    email,

    contraseñaHash,

    rol,

    estado,
  ]);

  // ---------------------------------------------------
  // REGISTRAR EN HISTORIAL
  // ---------------------------------------------------

  registrarMovimientoVehiculo({
    usuario: permiso.usuario.nombre,

    usuarioSistema: permiso.usuario.usuario,

    vehiculo: "",

    placa: "",

    campo: "USUARIO",

    valorAnterior: "",

    valorNuevo: usuario,

    tipoMovimiento: "USUARIO",

    accion: "USUARIO CREADO",
  });

  return {
    success: true,

    mensaje: "Usuario creado correctamente.",
  };
}

// =====================================================
// BORRAR USUARIO
// SOLO ADMINISTRADOR
// =====================================================

function borrarUsuarioAdministracion(token, idUsuario) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  idUsuario =
    idUsuario !== undefined && idUsuario !== null
      ? idUsuario.toString().trim()
      : "";

  if (!idUsuario) {
    return {
      success: false,

      mensaje: "No se recibió el usuario que deseas eliminar.",
    };
  }

  const hoja = obtenerHojaUsuarios();

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,

      mensaje: "No existen usuarios registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();

  let filaEliminar = -1;

  let usuarioEliminar = null;

  for (let i = 0; i < datos.length; i++) {
    const idActual = String(datos[i][0] || "").trim();

    if (idActual !== idUsuario) {
      continue;
    }

    filaEliminar = i + 2;

    usuarioEliminar = {
      id: idActual,

      nombre: String(datos[i][1] || ""),

      usuario: normalizarUsuario(datos[i][2]),

      email: String(datos[i][3] || ""),

      rol: normalizarRol(datos[i][5]),

      estado: normalizarEstado(datos[i][6]),
    };

    break;
  }

  if (filaEliminar === -1) {
    return {
      success: false,

      mensaje: "No se encontró el usuario seleccionado.",
    };
  }

  // ---------------------------------------------------
  // IMPEDIR ELIMINAR LA PROPIA CUENTA
  // ---------------------------------------------------

  if (
    normalizarUsuario(usuarioEliminar.usuario) ===
    normalizarUsuario(permiso.usuario.usuario)
  ) {
    return {
      success: false,

      mensaje: "No puedes eliminar tu propia cuenta.",
    };
  }

  // ---------------------------------------------------
  // PROTEGER AL ÚLTIMO ADMINISTRADOR ACTIVO
  // ---------------------------------------------------

  if (
    usuarioEliminar.rol === "ADMINISTRADOR" &&
    usuarioEliminar.estado === "ACTIVO"
  ) {
    let administradoresActivos = 0;

    for (let i = 0; i < datos.length; i++) {
      if (
        normalizarRol(datos[i][5]) === "ADMINISTRADOR" &&
        normalizarEstado(datos[i][6]) === "ACTIVO"
      ) {
        administradoresActivos++;
      }
    }

    if (administradoresActivos <= 1) {
      return {
        success: false,

        mensaje: "No puedes eliminar al último administrador activo.",
      };
    }
  }

  // ---------------------------------------------------
  // REGISTRAR EN HISTORIAL
  // ANTES DE ELIMINAR
  // ---------------------------------------------------

  registrarMovimientoVehiculo({
    usuario: permiso.usuario.nombre,

    usuarioSistema: permiso.usuario.usuario,

    vehiculo: "",

    placa: "",

    campo: "USUARIO",

    valorAnterior: usuarioEliminar.usuario,

    valorNuevo: "",

    tipoMovimiento: "USUARIO",

    accion: "USUARIO ELIMINADO",
  });

  // ---------------------------------------------------
  // ELIMINAR FILA
  // ---------------------------------------------------

  hoja.deleteRow(filaEliminar);

  return {
    success: true,

    mensaje: "Usuario eliminado correctamente.",
  };
}

// =====================================================
// OBTENER HISTORIAL DE MOVIMIENTOS
// SOLO ADMINISTRADOR
// =====================================================

function obtenerHistorialMovimientos(token) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName("HISTORIAL_MOVIMIENTOS");

  if (!hoja) {
    return {
      success: true,
      movimientos: [],
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: true,
      movimientos: [],
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 14).getDisplayValues();

  const movimientos = datos
    .map(function (fila) {
      return {
        fecha: fila[0] || "",

        usuario: fila[1] || "",

        usuarioSistema: fila[2] || "",

        vehiculo: fila[3] || "",

        placa: fila[4] || "",

        ubicacionAnterior: fila[5] || "",

        ubicacionNueva: fila[6] || "",

        puestoAnterior: fila[7] || "",

        puestoNuevo: fila[8] || "",

        accion: fila[9] || "",

        tipoMovimiento: fila[10] || "",

        campo: fila[11] || "",

        valorAnterior: fila[12] || "",

        valorNuevo: fila[13] || "",
      };
    })
    .reverse();

  return {
    success: true,

    movimientos: movimientos,
  };
}

// =====================================================
// CREAR NUEVO VEHÍCULO
// SOLO ADMINISTRADOR
// =====================================================

function crearVehiculoAdministracion(
  token,
  nombre,
  placa,
  transito,
  soat,
  tecnomecanica,
  picoYPlaca,
  precio,
  kilometraje,
  detalles,
  ubicacion,
) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  /* =================================================
     NORMALIZAR DATOS
  ================================================= */

  nombre =
    nombre !== undefined && nombre !== null
      ? nombre.toString().trim().toUpperCase()
      : "";

  placa =
    placa !== undefined && placa !== null
      ? placa.toString().trim().toUpperCase()
      : "";

  transito =
    transito !== undefined && transito !== null
      ? transito.toString().trim().toUpperCase()
      : "";

  ubicacion =
    ubicacion !== undefined && ubicacion !== null
      ? ubicacion.toString().trim()
      : "";

  soat = soat !== undefined && soat !== null ? soat.toString().trim() : "";

  tecnomecanica =
    tecnomecanica !== undefined && tecnomecanica !== null
      ? tecnomecanica.toString().trim()
      : "";

  picoYPlaca =
    picoYPlaca !== undefined && picoYPlaca !== null
      ? picoYPlaca.toString().trim().toUpperCase()
      : "";

  precio =
    precio !== undefined && precio !== null ? precio.toString().trim() : "";

  kilometraje =
    kilometraje !== undefined && kilometraje !== null
      ? kilometraje.toString().trim()
      : "";

  detalles =
    detalles !== undefined && detalles !== null
      ? detalles.toString().trim().toUpperCase()
      : "";

  /* =================================================
     AL MENOS UN DATO
  ================================================= */

  if (
    !nombre &&
    !placa &&
    !transito &&
    !soat &&
    !tecnomecanica &&
    !picoYPlaca &&
    !precio &&
    !kilometraje &&
    !detalles &&
    !ubicacion
  ) {
    return {
      success: false,
      mensaje: "Debes ingresar por lo menos un dato del vehículo.",
    };
  }

  /* =================================================
     VALIDAR PLACA SOLO SI FUE INGRESADA
  ================================================= */

  if (placa && !/^[A-Z]{3}[0-9]{3}$/.test(placa)) {
    return {
      success: false,
      mensaje: "La placa debe tener el formato ABC123.",
    };
  }

  /* =================================================
     VALIDAR PRECIO SOLO SI FUE INGRESADO
  ================================================= */

  if (precio) {
    const precioNumero = Number(
      precio.replace(/\./g, "").replace(/,/g, "").replace(/\$/g, "").trim(),
    );

    if (!Number.isFinite(precioNumero) || precioNumero < 0) {
      return {
        success: false,
        mensaje: "El precio ingresado no es válido.",
      };
    }
  }

  /* =================================================
     VALIDAR KILOMETRAJE SOLO SI FUE INGRESADO
  ================================================= */

  if (kilometraje) {
    const kilometrajeNumero = Number(
      kilometraje.replace(/\./g, "").replace(/,/g, "").trim(),
    );

    if (!Number.isFinite(kilometrajeNumero) || kilometrajeNumero < 0) {
      return {
        success: false,
        mensaje: "El kilometraje ingresado no es válido.",
      };
    }
  }

  /* =================================================
     ABRIR VEHÍCULOS
  ================================================= */

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  /* =================================================
     LEER A:L
  ================================================= */

  let datos = [];

  if (ultimaFila >= 2) {
    datos = hoja.getRange(2, 1, ultimaFila - 1, 12).getDisplayValues();
  }

  /* =================================================
     EVITAR PLACA DUPLICADA
     SOLO SI LA PLACA NO ESTÁ VACÍA
  ================================================= */

  if (placa) {
    const placaDuplicada = datos.some(function (fila) {
      const placaExistente = fila[1]
        ? fila[1].toString().trim().toUpperCase()
        : "";

      return placaExistente === placa;
    });

    if (placaDuplicada) {
      return {
        success: false,
        mensaje: "Ya existe un vehículo registrado con la placa " + placa + ".",
      };
    }
  }

  /* =================================================
     BUSCAR PRIMERA FILA COMPLETAMENTE VACÍA A:L
  ================================================= */

  let filaDestino = 0;

  for (let i = 0; i < datos.length; i++) {
    const filaVacia = datos[i].every(function (valor) {
      return valor.toString().trim() === "";
    });

    if (filaVacia) {
      filaDestino = i + 2;

      break;
    }
  }

  /* =================================================
     SI NO HAY HUECO, USAR LA SIGUIENTE FILA
  ================================================= */

  if (!filaDestino) {
    filaDestino = Math.max(ultimaFila + 1, 2);
  }

  /* =================================================
     ASEGURAR QUE EXISTA LA FILA
  ================================================= */

  if (filaDestino > hoja.getMaxRows()) {
    hoja.insertRowsAfter(hoja.getMaxRows(), filaDestino - hoja.getMaxRows());
  }

  /* =================================================
     GUARDAR A:L

     A  NOMBRE
     B  PLACA
     C  SOAT
     D  TECNOMECÁNICA
     E  PICO Y PLACA
     F  PRECIO
     G  KILOMETRAJE
     H  DETALLES
     I  UBICACIÓN
     J  PUESTO
     K  TRÁNSITO
     L  IMAGEN
  ================================================= */

  hoja
    .getRange(filaDestino, 1, 1, 13)
    .setValues([
      [
        nombre,
        placa,
        soat,
        tecnomecanica,
        picoYPlaca,
        precio,
        kilometraje,
        detalles,
        ubicacion,
        "",
        transito,
        "",
        "",
      ],
    ]);

  // =====================================================
  // REGISTRAR CREACIÓN EN HISTORIAL
  // =====================================================

  registrarMovimientoVehiculo({
    usuario: permiso.usuario.nombre,

    usuarioSistema: permiso.usuario.usuario,

    vehiculo: nombre,

    placa: placa,

    campo: "VEHÍCULO",

    valorAnterior: "",

    valorNuevo: nombre || placa || "Vehículo nuevo",

    tipoMovimiento: "VEHÍCULO",

    accion: "AGREGÓ VEHÍCULO",
  });

  return {
    success: true,
    mensaje: "Vehículo agregado correctamente.",
    fila: filaDestino,
  };
}

// =====================================================
// CITAS - ADMINISTRACIÓN
// SOLO ADMINISTRADOR
// =====================================================

// =====================================================
// NORMALIZAR FECHA DE CITA
// Acepta yyyy-MM-dd o dd/MM/yyyy
// Devuelve dd/MM/yyyy
// =====================================================

function normalizarFechaCitaAdministracion(valor) {
  if (valor === undefined || valor === null) {
    return "";
  }

  const texto = valor.toString().trim();

  if (!texto) {
    return "";
  }

  let dia = 0;
  let mes = 0;
  let anio = 0;

  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
    const partes = texto.split("-");

    anio = Number(partes[0]);
    mes = Number(partes[1]);
    dia = Number(partes[2]);
  } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
    const partes = texto.split("/");

    dia = Number(partes[0]);
    mes = Number(partes[1]);
    anio = Number(partes[2]);
  } else {
    return "";
  }

  const fecha = new Date(anio, mes - 1, dia);

  if (
    isNaN(fecha.getTime()) ||
    fecha.getFullYear() !== anio ||
    fecha.getMonth() !== mes - 1 ||
    fecha.getDate() !== dia
  ) {
    return "";
  }

  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

// =====================================================
// VALIDAR TIPO DE CITA
// El valor vacío es válido porque TIPO es opcional.
// =====================================================

function validarTipoCitaAdministracion(tipo) {
  const tiposPermitidos = [
    "",
    "VER VEHÍCULO",
    "PRUEBA DE MANEJO",
    "NEGOCIACIÓN",
    "PERITAJE",
    "RECIBIR VEHÍCULO",
    "OTRO",
  ];

  return tiposPermitidos.includes(tipo);
}

// =====================================================
// VALIDAR VEHÍCULO Y PLACA PARA CITA
// VEHÍCULO es obligatorio.
// PLACA se obtiene siempre desde la hoja VEHICULOS.
// =====================================================

function validarVehiculoCitaAdministracion(vehiculo, placa) {
  vehiculo =
    vehiculo !== undefined && vehiculo !== null
      ? vehiculo.toString().trim().toUpperCase()
      : "";

  placa =
    placa !== undefined && placa !== null
      ? placa.toString().trim().toUpperCase()
      : "";

  if (!vehiculo) {
    return {
      success: false,
      mensaje: "Debes seleccionar un vehículo.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "VEHICULOS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No hay vehículos registrados.",
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 2).getDisplayValues();

  for (let i = 0; i < datos.length; i++) {
    const nombreRegistrado = String(datos[i][0] || "")
      .trim()
      .toUpperCase();

    const placaRegistrada = String(datos[i][1] || "")
      .trim()
      .toUpperCase();

    if (nombreRegistrado !== vehiculo) {
      continue;
    }

    if (placa && placa !== placaRegistrada) {
      return {
        success: false,
        mensaje: "La placa no corresponde al vehículo seleccionado.",
      };
    }

    return {
      success: true,
      vehiculo: nombreRegistrado,
      placa: placaRegistrada,
    };
  }

  return {
    success: false,
    mensaje: "El vehículo seleccionado no existe.",
  };
}

// =====================================================
// CREAR CITA
// SOLO ADMINISTRADOR
// =====================================================

function crearCitaAdministracion(
  token,
  fecha,
  hora,
  vehiculo,
  placa,
  cliente,
  telefono,
  tipo,
  observaciones,
) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const fechaOriginal =
    fecha !== undefined && fecha !== null ? fecha.toString().trim() : "";

  hora = hora !== undefined && hora !== null ? hora.toString().trim() : "";

  vehiculo =
    vehiculo !== undefined && vehiculo !== null
      ? vehiculo.toString().trim().toUpperCase()
      : "";

  placa =
    placa !== undefined && placa !== null
      ? placa.toString().trim().toUpperCase()
      : "";

  cliente =
    cliente !== undefined && cliente !== null
      ? cliente.toString().trim().toUpperCase()
      : "";

  telefono =
    telefono !== undefined && telefono !== null
      ? telefono.toString().trim()
      : "";

  tipo =
    tipo !== undefined && tipo !== null
      ? tipo.toString().trim().toUpperCase()
      : "";

  observaciones =
    observaciones !== undefined && observaciones !== null
      ? observaciones.toString().trim().toUpperCase()
      : "";

  // ===================================================
  // CAMPOS OBLIGATORIOS
  // ===================================================

  if (!fechaOriginal) {
    return {
      success: false,
      mensaje: "Debes seleccionar una fecha.",
    };
  }

  if (!hora) {
    return {
      success: false,
      mensaje: "Debes seleccionar una hora.",
    };
  }

  if (!vehiculo) {
    return {
      success: false,
      mensaje: "Debes seleccionar un vehículo.",
    };
  }

  // ===================================================
  // VALIDAR FECHA
  // ===================================================

  fecha = normalizarFechaCitaAdministracion(fechaOriginal);

  if (!fecha) {
    return {
      success: false,
      mensaje: "La fecha de la cita no es válida.",
    };
  }

  // ===================================================
  // VALIDAR HORA
  // ===================================================

  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(hora)) {
    return {
      success: false,
      mensaje: "La hora de la cita no es válida.",
    };
  }

  // ===================================================
  // VALIDAR TIPO
  // ===================================================

  if (!validarTipoCitaAdministracion(tipo)) {
    return {
      success: false,
      mensaje: "El tipo de cita no es válido.",
    };
  }

  // ===================================================
  // VALIDAR VEHÍCULO Y PLACA
  // ===================================================

  const validacionVehiculo = validarVehiculoCitaAdministracion(vehiculo, placa);

  if (!validacionVehiculo.success) {
    return validacionVehiculo;
  }

  vehiculo = validacionVehiculo.vehiculo;

  placa = validacionVehiculo.placa;

  // ===================================================
  // ABRIR HOJA CITAS
  // ===================================================

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_CITAS);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "CITAS".',
    };
  }

  // ===================================================
  // DATOS AUTOMÁTICOS
  // ===================================================

  const id = Utilities.getUuid();

  const estado = "PENDIENTE";

  const creadaPor =
    permiso.usuario.nombre || permiso.usuario.usuario || "USUARIO";

  const ahora = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss",
  );

  // ===================================================
  // GUARDAR CITA
  // ===================================================

  hoja.appendRow([
    id,
    fecha,
    hora,
    vehiculo,
    placa,
    cliente,
    telefono,
    tipo,
    estado,
    observaciones,
    creadaPor,
    ahora,
    ahora,
  ]);

  return {
    success: true,

    mensaje: "La cita fue creada correctamente.",

    cita: {
      id: id,
      fecha: fecha,
      hora: hora,
      vehiculo: vehiculo,
      placa: placa,
      cliente: cliente,
      telefono: telefono,
      tipo: tipo,
      estado: estado,
      observaciones: observaciones,
      creadaPor: creadaPor,
      fechaCreacion: ahora,
      fechaModificacion: ahora,
    },
  };
}

// =====================================================
// OBTENER CITAS
// SOLO ADMINISTRADOR
// =====================================================

function obtenerCitasAdministracion(token) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_CITAS);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "CITAS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: true,
      citas: [],
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 13).getDisplayValues();

  const citas = [];

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];

    const id = fila[0] ? fila[0].toString().trim() : "";

    if (!id) {
      continue;
    }

    citas.push({
      fila: i + 2,

      id: id,

      fecha: fila[1] || "",

      hora: fila[2] || "",

      vehiculo: fila[3] || "",

      placa: fila[4] || "",

      cliente: fila[5] || "",

      telefono: fila[6] || "",

      tipo: fila[7] || "",

      estado: fila[8] || "",

      observaciones: fila[9] || "",

      creadaPor: fila[10] || "",

      fechaCreacion: fila[11] || "",

      fechaModificacion: fila[12] || "",
    });
  }

  citas.sort(function (a, b) {
    function convertirFechaHora(cita) {
      const partesFecha = cita.fecha.split("/");

      const partesHora = cita.hora.split(":");

      if (partesFecha.length !== 3 || partesHora.length !== 2) {
        return 0;
      }

      return new Date(
        Number(partesFecha[2]),
        Number(partesFecha[1]) - 1,
        Number(partesFecha[0]),
        Number(partesHora[0]),
        Number(partesHora[1]),
        0,
        0,
      ).getTime();
    }

    return convertirFechaHora(a) - convertirFechaHora(b);
  });

  return {
    success: true,

    citas: citas,
  };
}

// =====================================================
// CAMBIAR ESTADO DE CITA
// SOLO ADMINISTRADOR
// =====================================================

function cambiarEstadoCitaAdministracion(token, idCita, nuevoEstado) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  idCita =
    idCita !== undefined && idCita !== null ? idCita.toString().trim() : "";

  nuevoEstado =
    nuevoEstado !== undefined && nuevoEstado !== null
      ? nuevoEstado.toString().trim().toUpperCase()
      : "";

  if (!idCita) {
    return {
      success: false,
      mensaje: "No se recibió la cita.",
    };
  }

  const estadosPermitidos = [
    "PENDIENTE",
    "CONFIRMADA",
    "ATENDIDA",
    "CANCELADA",
  ];

  if (!estadosPermitidos.includes(nuevoEstado)) {
    return {
      success: false,
      mensaje: "El estado de la cita no es válido.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_CITAS);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "CITAS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No existen citas registradas.",
    };
  }

  const ids = hoja.getRange(2, 1, ultimaFila - 1, 1).getDisplayValues();

  let filaCita = -1;

  for (let i = 0; i < ids.length; i++) {
    const idActual = ids[i][0] ? ids[i][0].toString().trim() : "";

    if (idActual === idCita) {
      filaCita = i + 2;

      break;
    }
  }

  if (filaCita === -1) {
    return {
      success: false,
      mensaje: "No se encontró la cita.",
    };
  }

  const ahora = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss",
  );

  hoja.getRange(filaCita, 9).setValue(nuevoEstado);

  hoja.getRange(filaCita, 13).setValue(ahora);

  return {
    success: true,

    mensaje: "El estado de la cita fue actualizado correctamente.",

    id: idCita,

    estado: nuevoEstado,

    fechaModificacion: ahora,
  };
}

// =====================================================
// OBTENER ESTADO TEMPORAL DE UNA CITA
// ATRASADA es un estado visual, no se guarda en la hoja.
// =====================================================

function obtenerEstadoTemporalCita(fecha, hora, estado) {
  estado =
    estado !== undefined && estado !== null
      ? estado.toString().trim().toUpperCase()
      : "";

  if (estado === "ATENDIDA" || estado === "CANCELADA") {
    return {
      activa: false,
      tipo: "FINALIZADA",
      minutosDiferencia: null,
    };
  }

  fecha = normalizarFechaCitaAdministracion(fecha);

  hora = hora !== undefined && hora !== null ? hora.toString().trim() : "";

  if (!fecha || !/^([01]\d|2[0-3]):[0-5]\d$/.test(hora)) {
    return {
      activa: false,
      tipo: "FECHA_INVALIDA",
      minutosDiferencia: null,
    };
  }

  const partesFecha = fecha.split("/");

  const partesHora = hora.split(":");

  const fechaCita = new Date(
    Number(partesFecha[2]),
    Number(partesFecha[1]) - 1,
    Number(partesFecha[0]),
    Number(partesHora[0]),
    Number(partesHora[1]),
    0,
    0,
  );

  if (isNaN(fechaCita.getTime())) {
    return {
      activa: false,
      tipo: "FECHA_INVALIDA",
      minutosDiferencia: null,
    };
  }

  const ahora = new Date();

  const diferenciaMs = fechaCita.getTime() - ahora.getTime();

  const minutos = Math.ceil(diferenciaMs / (1000 * 60));

  if (minutos < 0) {
    return {
      activa: true,
      tipo: "ATRASADA",
      minutosDiferencia: Math.abs(minutos),
    };
  }

  if (minutos <= 30) {
    return {
      activa: true,
      tipo: "EN_30_MINUTOS",
      minutosDiferencia: minutos,
    };
  }

  if (minutos <= 60) {
    return {
      activa: true,
      tipo: "EN_1_HORA",
      minutosDiferencia: minutos,
    };
  }

  return {
    activa: true,
    tipo: "PROXIMA",
    minutosDiferencia: minutos,
  };
}

// =====================================================
// OBTENER CITAS PARA INFORMACIÓN RÁPIDA
// SOLO ADMINISTRADOR
// =====================================================

function obtenerCitasInformacionRapida(token) {
  const respuesta = obtenerCitasAdministracion(token);

  if (!respuesta.success) {
    return respuesta;
  }

  const citas = Array.isArray(respuesta.citas) ? respuesta.citas : [];

  const atrasadas = [];
  const proximas = [];

  function obtenerFechaHoraMs(cita) {
    const partesFecha = cita.fecha.split("/");

    const partesHora = cita.hora.split(":");

    if (partesFecha.length !== 3 || partesHora.length !== 2) {
      return 0;
    }

    return new Date(
      Number(partesFecha[2]),
      Number(partesFecha[1]) - 1,
      Number(partesFecha[0]),
      Number(partesHora[0]),
      Number(partesHora[1]),
      0,
      0,
    ).getTime();
  }

  for (let i = 0; i < citas.length; i++) {
    const cita = citas[i];

    const estadoTemporal = obtenerEstadoTemporalCita(
      cita.fecha,
      cita.hora,
      cita.estado,
    );

    if (!estadoTemporal.activa) {
      continue;
    }

    const citaPreparada = {
      id: cita.id,

      fecha: cita.fecha,

      hora: cita.hora,

      vehiculo: cita.vehiculo,

      placa: cita.placa,

      cliente: cita.cliente,

      telefono: cita.telefono,

      tipo: cita.tipo,

      estado: cita.estado,

      observaciones: cita.observaciones,

      estadoTemporal: estadoTemporal.tipo,

      minutosDiferencia: estadoTemporal.minutosDiferencia,

      fechaHoraMs: obtenerFechaHoraMs(cita),
    };

    if (estadoTemporal.tipo === "ATRASADA") {
      atrasadas.push(citaPreparada);
    } else {
      proximas.push(citaPreparada);
    }
  }

  atrasadas.sort(function (a, b) {
    return a.fechaHoraMs - b.fechaHoraMs;
  });

  proximas.sort(function (a, b) {
    return a.fechaHoraMs - b.fechaHoraMs;
  });

  const proximasLimitadas = proximas.slice(0, 5);

  return {
    success: true,

    atrasadas: atrasadas,

    proximas: proximasLimitadas,

    totalAtrasadas: atrasadas.length,

    totalProximas: proximas.length,
  };
}

// =====================================================
// EDITAR CITA
// SOLO ADMINISTRADOR
// =====================================================

function actualizarCitaAdministracion(
  token,
  idCita,
  fecha,
  hora,
  vehiculo,
  placa,
  cliente,
  telefono,
  tipo,
  observaciones,
) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  idCita =
    idCita !== undefined && idCita !== null ? idCita.toString().trim() : "";

  const fechaOriginal =
    fecha !== undefined && fecha !== null ? fecha.toString().trim() : "";

  hora = hora !== undefined && hora !== null ? hora.toString().trim() : "";

  vehiculo =
    vehiculo !== undefined && vehiculo !== null
      ? vehiculo.toString().trim().toUpperCase()
      : "";

  placa =
    placa !== undefined && placa !== null
      ? placa.toString().trim().toUpperCase()
      : "";

  cliente =
    cliente !== undefined && cliente !== null
      ? cliente.toString().trim().toUpperCase()
      : "";

  telefono =
    telefono !== undefined && telefono !== null
      ? telefono.toString().trim()
      : "";

  tipo =
    tipo !== undefined && tipo !== null
      ? tipo.toString().trim().toUpperCase()
      : "";

  observaciones =
    observaciones !== undefined && observaciones !== null
      ? observaciones.toString().trim().toUpperCase()
      : "";

  // ===================================================
  // VALIDACIONES
  // ===================================================

  if (!idCita) {
    return {
      success: false,
      mensaje: "No se recibió la cita.",
    };
  }

  if (!fechaOriginal) {
    return {
      success: false,
      mensaje: "Debes seleccionar una fecha.",
    };
  }

  if (!hora) {
    return {
      success: false,
      mensaje: "Debes seleccionar una hora.",
    };
  }

  if (!vehiculo) {
    return {
      success: false,
      mensaje: "Debes seleccionar un vehículo.",
    };
  }

  fecha = normalizarFechaCitaAdministracion(fechaOriginal);

  if (!fecha) {
    return {
      success: false,
      mensaje: "La fecha de la cita no es válida.",
    };
  }

  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(hora)) {
    return {
      success: false,
      mensaje: "La hora de la cita no es válida.",
    };
  }

  if (!validarTipoCitaAdministracion(tipo)) {
    return {
      success: false,
      mensaje: "El tipo de cita no es válido.",
    };
  }

  const validacionVehiculo = validarVehiculoCitaAdministracion(vehiculo, placa);

  if (!validacionVehiculo.success) {
    return validacionVehiculo;
  }

  vehiculo = validacionVehiculo.vehiculo;

  placa = validacionVehiculo.placa;

  // ===================================================
  // ABRIR HOJA
  // ===================================================

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_CITAS);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "CITAS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No existen citas registradas.",
    };
  }

  const ids = hoja.getRange(2, 1, ultimaFila - 1, 1).getDisplayValues();

  let filaCita = -1;

  for (let i = 0; i < ids.length; i++) {
    const idActual = ids[i][0] ? ids[i][0].toString().trim() : "";

    if (idActual === idCita) {
      filaCita = i + 2;

      break;
    }
  }

  if (filaCita === -1) {
    return {
      success: false,
      mensaje: "No se encontró la cita.",
    };
  }

  const ahora = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss",
  );

  hoja
    .getRange(filaCita, 2, 1, 7)
    .setValues([[fecha, hora, vehiculo, placa, cliente, telefono, tipo]]);

  hoja.getRange(filaCita, 10).setValue(observaciones);

  hoja.getRange(filaCita, 13).setValue(ahora);

  return {
    success: true,

    mensaje: "La cita fue actualizada correctamente.",

    cita: {
      id: idCita,

      fecha: fecha,

      hora: hora,

      vehiculo: vehiculo,

      placa: placa,

      cliente: cliente,

      telefono: telefono,

      tipo: tipo,

      observaciones: observaciones,

      fechaModificacion: ahora,
    },
  };
}

// =====================================================
// ELIMINAR CITA
// SOLO ADMINISTRADOR
// Se conserva para registros creados por error.
// =====================================================

function eliminarCitaAdministracion(token, idCita) {
  const permiso = validarAdministrador(token);

  if (!permiso.success) {
    return permiso;
  }

  idCita =
    idCita !== undefined && idCita !== null ? idCita.toString().trim() : "";

  if (!idCita) {
    return {
      success: false,
      mensaje: "No se recibió la cita.",
    };
  }

  const archivo = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = archivo.getSheetByName(NOMBRE_HOJA_CITAS);

  if (!hoja) {
    return {
      success: false,
      mensaje: 'No se encontró la hoja "CITAS".',
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      success: false,
      mensaje: "No existen citas registradas.",
    };
  }

  const ids = hoja.getRange(2, 1, ultimaFila - 1, 1).getDisplayValues();

  let filaCita = -1;

  for (let i = 0; i < ids.length; i++) {
    const idActual = ids[i][0] ? ids[i][0].toString().trim() : "";

    if (idActual === idCita) {
      filaCita = i + 2;

      break;
    }
  }

  if (filaCita === -1) {
    return {
      success: false,
      mensaje: "No se encontró la cita.",
    };
  }

  hoja.deleteRow(filaCita);

  return {
    success: true,

    mensaje: "La cita fue eliminada correctamente.",

    id: idCita,
  };
}

// =====================================================
// ALERTAS AUTOMÁTICAS DE CITAS
// FUNCIONAN AUNQUE LA APP ESTÉ CERRADA
// =====================================================

function revisarCitasAutomaticamente() {
  const libro = SpreadsheetApp.openById(ID_ARCHIVO);

  const hoja = libro.getSheetByName(NOMBRE_HOJA_CITAS);

  if (!hoja) {
    return;
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return;
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 13).getValues();

  const ahora = new Date();

  datos.forEach(function (fila) {
    const id = String(fila[0] || "").trim();

    const fecha = fila[1];

    const hora = fila[2];

    const vehiculo = String(fila[3] || "").trim();

    const placa = String(fila[4] || "").trim();

    const cliente = String(fila[5] || "").trim();

    const telefono = String(fila[6] || "").trim();

    const tipo = String(fila[7] || "").trim();

    const estado = String(fila[8] || "")
      .trim()
      .toUpperCase();

    if (!id) {
      return;
    }

    if (estado === "ATENDIDA" || estado === "CANCELADA") {
      return;
    }

    const fechaHora = construirFechaHoraCitaTrigger(fecha, hora);

    if (!fechaHora) {
      return;
    }

    const diferenciaMs = fechaHora.getTime() - ahora.getTime();

    const minutos = Math.ceil(diferenciaMs / 60000);

    if (minutos < 0) {
      return;
    }

    // =============================================
    // ALERTA DE 1 HORA
    // =============================================

    if (minutos <= 60 && minutos > 30) {
      enviarAlertaCitaTrigger(
        id,
        fechaHora,
        60,
        vehiculo,
        placa,
        cliente,
        telefono,
        tipo,
      );
    }

    // =============================================
    // ALERTA DE 30 MINUTOS
    // =============================================
    else if (minutos <= 30) {
      enviarAlertaCitaTrigger(
        id,
        fechaHora,
        30,
        vehiculo,
        placa,
        cliente,
        telefono,
        tipo,
      );
    }
  });
}

function construirFechaHoraCitaTrigger(fecha, hora) {
  let fechaBase;

  if (
    Object.prototype.toString.call(fecha) === "[object Date]" &&
    !isNaN(fecha.getTime())
  ) {
    fechaBase = new Date(fecha);
  } else {
    const partes = String(fecha || "").split("/");

    if (partes.length !== 3) {
      return null;
    }

    fechaBase = new Date(
      Number(partes[2]),
      Number(partes[1]) - 1,
      Number(partes[0]),
    );
  }

  let horas = 0;
  let minutos = 0;

  if (
    Object.prototype.toString.call(hora) === "[object Date]" &&
    !isNaN(hora.getTime())
  ) {
    horas = hora.getHours();

    minutos = hora.getMinutes();
  } else {
    const partesHora = String(hora || "").split(":");

    if (partesHora.length < 2) {
      return null;
    }

    horas = Number(partesHora[0]);

    minutos = Number(partesHora[1]);
  }

  fechaBase.setHours(horas, minutos, 0, 0);

  return fechaBase;
}

function enviarAlertaCitaTrigger(
  idCita,
  fechaHora,
  minutosAlerta,
  vehiculo,
  placa,
  cliente,
  telefono,
  tipo,
) {
  const propiedades = PropertiesService.getScriptProperties();

  const clave = [
    "ALERTA_CITA",
    idCita,
    fechaHora.getTime(),
    minutosAlerta,
  ].join("_");

  if (propiedades.getProperty(clave)) {
    return;
  }

  const zonaHoraria = Session.getScriptTimeZone();

  const fechaTexto = Utilities.formatDate(fechaHora, zonaHoraria, "dd/MM/yyyy");

  const horaTexto = Utilities.formatDate(fechaHora, zonaHoraria, "HH:mm");

  const asunto =
    minutosAlerta === 60
      ? "📅 Cita en aproximadamente 1 hora"
      : "⏰ Cita en 30 minutos o menos";

  const cuerpo = `
Tienes una cita próxima.

Vehículo: ${vehiculo || "-"}
Placa: ${placa || "-"}
Fecha: ${fechaTexto}
Hora: ${horaTexto}
Cliente: ${cliente || "-"}
Teléfono: ${telefono || "-"}
Tipo: ${tipo || "-"}

Sistema de Vehículos
  `.trim();

  MailApp.sendEmail(EMAIL_ALERTAS_CITAS, asunto, cuerpo);

  propiedades.setProperty(clave, new Date().toISOString());
}

function crearTriggerAlertasCitas() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === "revisarCitasAutomaticamente") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("revisarCitasAutomaticamente")
    .timeBased()
    .everyMinutes(5)
    .create();
}
