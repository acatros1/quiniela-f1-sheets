/**
 * ====================================================================
 * SISTEMA AUTOMATIZADO DE QUINIELA / FANTASY DE FÓRMULA 1 (2025 - 2026)
 * ====================================================================
 */

const PUNTOS_POSICION = [25, 18, 15, 12, 10, 8, 6, 4, 2, 1];

const PILOTOS_RESPALDO_2025 = [
  "Alexander Albon (ALB)",
  "Andrea Kimi Antonelli (ANT)",
  "Carlos Sainz (SAI)",
  "Charles Leclerc (LEC)",
  "Esteban Ocon (OCO)",
  "Fernando Alonso (ALO)",
  "Gabriel Bortoleto (BOR)",
  "George Russell (RUS)",
  "Isack Hadjar (HAD)",
  "Jack Doohan (DOO)",
  "Lance Stroll (STR)",
  "Lando Norris (NOR)",
  "Lewis Hamilton (HAM)",
  "Liam Lawson (LAW)",
  "Max Verstappen (VER)",
  "Nico Hülkenberg (HUL)",
  "Oliver Bearman (BEA)",
  "Oscar Piastri (PIA)",
  "Pierre Gasly (GAS)",
  "Yuki Tsunoda (TSU)"
];

/**
 * Obtiene el mapeo oficial de participantes desde Script Properties
 * para no exponer correos personales en el código fuente.
 */
function obtenerMapeoJugadoresConfig() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('MAPEO_JUGADORES');
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    Logger.log('⚠️ Error al leer MAPEO_JUGADORES desde Script Properties: ' + e);
    return {};
  }
}

function notificarGlobal(mensaje) {
  Logger.log(mensaje);
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) ui.alert(mensaje);
  } catch (e) {}
}

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🏁 Quiniela F1')
      .addItem('⚙️ 1. Inicializar Tablas y Estructura', 'inicializarEstructura')
      .addSeparator()
      .addItem('👀 Actualizar Estatus de Envíos', 'actualizarEstadoPronosticos')
      .addItem('📲 Copiar Estatus para WhatsApp', 'verEstatusWhatsApp')
      .addSeparator()
      .addItem('📸 Llenar Tabla para Captura', 'generarVistaPronosticos')
      .addItem('💬 Copiar Pronósticos para WhatsApp', 'copiarPronosticosWhatsApp')
      .addSeparator()
      .addItem('🔒 Ocultar Respuestas del Formulario', 'ocultarRespuestas')
      .addItem('🔓 Mostrar Respuestas del Formulario', 'mostrarRespuestas')
      .addSeparator()
      .addItem('🔍 Verificar Estado de la API', 'verificarEstadoAPI')
      .addItem('📥 2. Descargar Resultados (API Jolpica)', 'menuObtenerResultadosAPI')
      .addItem('🧮 3. Calcular Puntos de la Carrera', 'calcularPuntuacionCarrera')
      .addItem('🏆 4. Actualizar Leaderboard y Gráficas', 'actualizarLeaderboard')
      .addSeparator()
      .addItem('📦 5. Migrar Datos Históricos (Hoja 1)', 'migrarResultadosExcel')
      .addItem('🎲 6. Simular Probabilidades (Monte Carlo)', 'calcularProbabilidadesCampeonato')
      .addToUi();
  } catch(e) {}
}

function inicializarEstructura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const pestañas = [
    {
      nombre: 'Estado_Pronosticos',
      encabezados: ['Participante', 'Estado', 'Último Envío Registrado']
    },
    {
      nombre: 'Captura_Pronosticos',
      encabezados: ['POSICIÓN']
    },
    {
      nombre: 'Respuestas_Formulario',
      encabezados: ['Marca temporal', 'P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10', 'Dirección de correo electrónico']
    },
    {
      nombre: 'Historico_Pronosticos',
      encabezados: ['Año', 'Carrera / GP', 'Fecha Envío', 'Correo Participante', 'P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10', 'Puntos Totales']
    },
    {
      nombre: 'Carreras_Anteriores',
      encabezados: ['Año', 'Carrera / GP', 'Correo Participante', 'Puntos Obtenidos']
    },
    {
      nombre: 'Resultados_API',
      encabezados: ['Posición', 'Piloto', 'Código', 'Escudería']
    },
    {
      nombre: 'Calculo_Puntos',
      encabezados: ['Correo Participante', 'P1 Pronostico', 'P1 Real', 'Pts P1', 'P2 Pronostico', 'P2 Real', 'Pts P2', 'P3 Pronostico', 'P3 Real', 'Pts P3', 'P4 Pronostico', 'P4 Real', 'Pts P4', 'P5 Pronostico', 'P5 Real', 'Pts P5', 'P6 Pronostico', 'P6 Real', 'Pts P6', 'P7 Pronostico', 'P7 Real', 'Pts P7', 'P8 Pronostico', 'P8 Real', 'Pts P8', 'P9 Pronostico', 'P9 Real', 'Pts P9', 'P10 Pronostico', 'P10 Real', 'Pts P10', 'PUNTOS TOTALES']
    },
    {
      nombre: 'Leaderboard',
      encabezados: ['Posición General', 'Participante', 'Puntos Acumulados', 'Dif. Ant.', 'Dif. Líder', 'Tendencia']
    },
    {
      nombre: 'Evolucion_Temporal',
      encabezados: ['Carrera / GP']
    }
  ];

  pestañas.forEach(p => {
    let sheet = ss.getSheetByName(p.nombre);
    if (!sheet) {
      sheet = ss.insertSheet(p.nombre);
    }
    if (sheet.getLastRow() === 0 && p.nombre !== 'Respuestas_Formulario' && p.nombre !== 'Captura_Pronosticos') {
      sheet.appendRow(p.encabezados);
      const headerRange = sheet.getRange(1, 1, 1, p.encabezados.length);
      headerRange.setBackground('#111827').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  actualizarEstadoPronosticos();
  notificarGlobal('✅ Estructura inicializada correctamente.');
}

function menuObtenerResultadosAPI() {
  try {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.prompt(
      '📥 Descargar Resultados de la API Jolpica', 
      '• Escribe "ULTIMA" (o déjalo en blanco) para descargar automáticamente el último GP disputado.\n\n' +
      '• O escribe el AÑO y RONDA separados por coma (ejemplo: 2025,15):', 
      ui.ButtonSet.OK_CANCEL
    );

    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const entrada = resp.getResponseText().trim().toUpperCase();

    if (entrada === 'ULTIMA' || entrada === 'LAST' || entrada === '') {
      consultarAPIJolpica('current', 'last');
    } else {
      const partes = entrada.split(',');
      if (partes.length === 2) {
        consultarAPIJolpica(partes[0].trim(), partes[1].trim());
      } else {
        notificarGlobal('❌ Formato incorrecto. Usa "ULTIMA" o "AÑO,RONDA" (ejemplo: 2025,15)');
      }
    }
  } catch (e) {
    notificarGlobal("Ejecuta esta función desde el menú '🏁 Quiniela F1' en Google Sheets.");
  }
}

function consultarAPIJolpica(año, carrera) {
  let url = (año === 'current' && carrera === 'last') 
    ? `https://api.jolpi.ca/ergast/f1/current/last/results.json`
    : `https://api.jolpi.ca/ergast/f1/${año}/${carrera}/results.json`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());

    const races = json.MRData.RaceTable.Races;
    if (!races || races.length === 0) {
      notificarGlobal('⚠️ No se encontraron resultados. Confirma si la carrera finalizó.');
      return;
    }

    const raceData = races[0];
    const results = raceData.Results;

    const fechaHoraGP = raceData.time ? `${raceData.date}T${raceData.time}` : raceData.date;
    PropertiesService.getScriptProperties().setProperties({
      'ULTIMO_GP_NOMBRE': raceData.raceName,
      'ULTIMO_GP_FECHA': fechaHoraGP
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAPI = ss.getSheetByName('Resultados_API');

    if (sheetAPI.getLastRow() > 1) {
      sheetAPI.getRange(2, 1, sheetAPI.getLastRow() - 1, 4).clearContent();
    }

    const top10 = [];
    for (let i = 0; i < Math.min(10, results.length); i++) {
      const res = results[i];
      const pos = res.position;
      const pilotoNombre = `${res.Driver.givenName} ${res.Driver.familyName}`;
      const codigo = res.Driver.code || res.Driver.driverId;
      const escuderia = res.Constructor.name;

      top10.push([pos, pilotoNombre, codigo, escuderia]);
    }

    sheetAPI.getRange(2, 1, top10.length, 4).setValues(top10);
    notificarGlobal(`✅ Resultados descargados exitosamente:\n\n` +
                    `🏎️ GP: ${raceData.raceName} (${raceData.season} - Ronda ${raceData.round})\n` +
                    `📍 ${raceData.Circuit.circuitName}`);

  } catch (error) {
    notificarGlobal('❌ Error al conectar con la API Jolpica: ' + error.toString());
  }
}

/**
 * Calcula puntos con sugerencia automática de Nombre/Fecha desde la API.
 * Corrige duplicados: Si un participante envió con 2 correos diferentes,
 * solo evalúa su último envío válido.
 */
function calcularPuntuacionCarrera() {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e){}

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetRespuestas = ss.getSheetByName('Respuestas_Formulario');
  const sheetAPI = ss.getSheetByName('Resultados_API');
  const sheetCalculo = ss.getSheetByName('Calculo_Puntos');
  const sheetAnteriores = ss.getSheetByName('Carreras_Anteriores');

  let sheetHistorico = ss.getSheetByName('Historico_Pronosticos');
  if (!sheetHistorico) {
    inicializarEstructura();
    sheetHistorico = ss.getSheetByName('Historico_Pronosticos');
  }

  if (!sheetRespuestas || sheetRespuestas.getLastRow() < 2) {
    if (ui) {
      const respDesierta = ui.alert(
        '🏎️ Carrera sin pronósticos registrados',
        `No hay respuestas registradas en el formulario para "${nombreGP}".\n\n` +
        `¿Deseas registrar este Gran Premio asignando 0 PUNTOS a todos los participantes oficiales?`,
        ui.ButtonSet.YES_NO
      );

      if (respDesierta === ui.Button.YES) {
        registrarCarreraEnBlanco(nombreGP);
        return;
      }
    }
    notificarGlobal('⚠️ No hay respuestas registradas en "Respuestas_Formulario".');
    return;
  }

  if (!sheetAPI || sheetAPI.getLastRow() < 11) {
    notificarGlobal('⚠️ Primero debes descargar los 10 resultados oficiales en "Resultados_API".');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const gpSugerido = props.getProperty('ULTIMO_GP_NOMBRE') || "GP Carrera";
  const fechaSugerida = props.getProperty('ULTIMO_GP_FECHA') || "";

  let nombreGP = gpSugerido;
  let fechaLimite = fechaSugerida ? new Date(fechaSugerida) : null;

  if (ui) {
    const respGP = ui.prompt('Identificador del GP', `Acepta o edita el nombre del GP para el registro:`, ui.ButtonSet.OK_CANCEL);
    if (respGP.getSelectedButton() !== ui.Button.OK) return;
    if (respGP.getResponseText().trim() !== "") {
      nombreGP = respGP.getResponseText().trim();
    }

    const respHora = ui.prompt(
      '⏱️ Hora Límite para enviar pronósticos', 
      `Fecha/Hora límite sugerida por la API: ${fechaSugerida || 'Ninguna'}\n\nIngresa una fecha (YYYY-MM-DD HH:MM) o deja igual para aplicar:`, 
      ui.ButtonSet.OK_CANCEL
    );

    if (respHora.getSelectedButton() === ui.Button.OK && respHora.getResponseText().trim() !== '') {
      const textoFecha = respHora.getResponseText().trim();
      const f = new Date(textoFecha);
      if (!isNaN(f.getTime())) {
        fechaLimite = f;
      }
    }
  }

  const datosAPI = sheetAPI.getRange(2, 1, 10, 4).getValues();
  const pilotosOficiales = datosAPI.map(f => ({
    posicion: parseInt(f[0]),
    nombreCompleto: f[1],
    codigo: f[2]
  }));

  const encabezados = sheetRespuestas.getRange(1, 1, 1, sheetRespuestas.getLastColumn()).getValues()[0];
  const respuestas = sheetRespuestas.getRange(2, 1, sheetRespuestas.getLastRow() - 1, sheetRespuestas.getLastColumn()).getValues();

  let idxCorreo = -1;
  for (let c = 0; c < encabezados.length; c++) {
    const header = String(encabezados[c]).toLowerCase().trim();
    if (header.includes("correo") || header.includes("email") || header.includes("dirección")) {
      idxCorreo = c;
      break;
    }
  }
  if (idxCorreo === -1) idxCorreo = encabezados.length - 1;

  // 1. FILTRADO INTELIGENTE: Deduplica por NOMBRE DE PARTICIPANTE (no por correo)
  const ultimasRespuestasMap = {}; // nombreJugador -> fila
  let jugadasDescartadasPorTiempo = 0;

  respuestas.forEach(fila => {
    const marcaTemporal = new Date(fila[0]);
    const correo = String(fila[idxCorreo] || '').trim();

    if (fechaLimite && marcaTemporal > fechaLimite) {
      jugadasDescartadasPorTiempo++;
      return;
    }

    if (correo) {
      // Obtenemos el nombre oficial mapeado
      const nombreJugador = obtenerNombreJugador(correo);

      // Si ya existía un envío previo del mismo jugador, conservamos el más reciente
      if (!ultimasRespuestasMap[nombreJugador] || marcaTemporal > new Date(ultimasRespuestasMap[nombreJugador][0])) {
        ultimasRespuestasMap[nombreJugador] = fila;
      }
    }
  });

  const respuestasAProcesar = Object.values(ultimasRespuestasMap);

  if (respuestasAProcesar.length === 0) {
    notificarGlobal('⚠️ No hay jugadas válidas para procesar.');
    return;
  }

  if (sheetCalculo.getLastRow() > 1) {
    sheetCalculo.getRange(2, 1, sheetCalculo.getLastRow() - 1, sheetCalculo.getLastColumn()).clearContent();
  }

  const filasCalculo = [];
  const filasHistorial = [];
  const filasRespaldoPronosticos = [];
  const anioActual = new Date().getFullYear();

  respuestasAProcesar.forEach(fila => {
    const fechaEnvio = fila[0];
    const correo = String(fila[idxCorreo] || '').trim(); 
    const pronosticos = fila.filter((val, idx) => idx !== 0 && idx !== idxCorreo).slice(0, 10);

    let puntosTotales = 0;
    let detalleFilaCalculo = [correo];

    for (let i = 0; i < 10; i++) {
      const pronostico = String(pronosticos[i] || '').trim();
      const oficial = pilotosOficiales[i];
      const oficialNombre = oficial ? oficial.nombreCompleto : '';
      const oficialCodigo = oficial ? oficial.codigo : '';

      const coincide = coincidenciaFlexible(pronostico, oficialNombre, oficialCodigo);
      const pts = coincide ? PUNTOS_POSICION[i] : 0;

      puntosTotales += pts;
      detalleFilaCalculo.push(pronostico, oficialNombre, pts);
    }

    detalleFilaCalculo.push(puntosTotales);
    filasCalculo.push(detalleFilaCalculo);

    filasHistorial.push([anioActual, nombreGP, correo, puntosTotales]);

    filasRespaldoPronosticos.push([
      anioActual,
      nombreGP,
      fechaEnvio,
      correo,
      pronosticos[0] || '',
      pronosticos[1] || '',
      pronosticos[2] || '',
      pronosticos[3] || '',
      pronosticos[4] || '',
      pronosticos[5] || '',
      pronosticos[6] || '',
      pronosticos[7] || '',
      pronosticos[8] || '',
      pronosticos[9] || '',
      puntosTotales
    ]);
  });

  // Limpiar registros viejos del mismo GP si se está recalculando
  if (sheetAnteriores.getLastRow() > 1) {
    const datosAnt = sheetAnteriores.getRange(2, 1, sheetAnteriores.getLastRow() - 1, 4).getValues();
    const datosFiltrados = datosAnt.filter(r => !(r[0] == anioActual && String(r[1]).toLowerCase().trim() == nombreGP.toLowerCase().trim()));
    sheetAnteriores.getRange(2, 1, sheetAnteriores.getLastRow() - 1, 4).clearContent();
    if (datosFiltrados.length > 0) {
      sheetAnteriores.getRange(2, 1, datosFiltrados.length, 4).setValues(datosFiltrados);
    }
  }

  if (sheetHistorico.getLastRow() > 1) {
    const datosHist = sheetHistorico.getRange(2, 1, sheetHistorico.getLastRow() - 1, 15).getValues();
    const histFiltrados = datosHist.filter(r => !(r[0] == anioActual && String(r[1]).toLowerCase().trim() == nombreGP.toLowerCase().trim()));
    sheetHistorico.getRange(2, 1, sheetHistorico.getLastRow() - 1, 15).clearContent();
    if (histFiltrados.length > 0) {
      sheetHistorico.getRange(2, 1, histFiltrados.length, 15).setValues(histFiltrados);
    }
  }

  if (filasCalculo.length > 0) {
    sheetCalculo.getRange(2, 1, filasCalculo.length, filasCalculo[0].length).setValues(filasCalculo);
  }

  if (filasHistorial.length > 0) {
    sheetAnteriores.getRange(sheetAnteriores.getLastRow() + 1, 1, filasHistorial.length, 4).setValues(filasHistorial);
  }

  if (filasRespaldoPronosticos.length > 0) {
    sheetHistorico.getRange(sheetHistorico.getLastRow() + 1, 1, filasRespaldoPronosticos.length, 15).setValues(filasRespaldoPronosticos);
  }

  let mensajeExito = `✅ Cálculo completado para ${filasCalculo.length} participante(s).\n\n` +
                     `🏎️ GP Procesado: ${nombreGP}\n` +
                     `📦 Pronósticos respaldados en "Historico_Pronosticos".`;
  if (jugadasDescartadasPorTiempo > 0) {
    mensajeExito += `\n🚫 Se descartaron ${jugadasDescartadasPorTiempo} envío(s) fuera de tiempo.`;
  }

  notificarGlobal(mensajeExito);
  actualizarLeaderboard();

  if (ui) {
    const respLimpiar = ui.alert(
      '🧹 Preparación para la siguiente carrera', 
      '¿Deseas vaciar las respuestas y preparar la quiniela para el próximo GP?\n\n' +
      '• Todos los participantes volverán a "⏳ PENDIENTE".\n' +
      '• Se vaciará la tabla de captura de pronósticos.\n' +
      '• (Tus puntos en Leaderboard e Historial quedan 100% respaldados)', 
      ui.ButtonSet.YES_NO
    );

    if (respLimpiar === ui.Button.YES) {
      if (sheetRespuestas.getLastRow() > 1) {
        sheetRespuestas.deleteRows(2, sheetRespuestas.getLastRow() - 1);
      }

      const sheetCaptura = ss.getSheetByName('Captura_Pronosticos');
      if (sheetCaptura && sheetCaptura.getLastRow() >= 3 && sheetCaptura.getLastColumn() > 1) {
        sheetCaptura.getRange(3, 2, 11, sheetCaptura.getLastColumn() - 1).clearContent();
      }

      actualizarEstadoPronosticos();
      notificarGlobal('✨ ¡Listo para el próximo GP!\n\nSe vaciaron las respuestas, la tabla de captura quedó limpia y todos los participantes están en "⏳ PENDIENTE".');
    }
  }
}

/**
 * Mapeo inteligente de correos:
 * Soporta correo único o lista de correos, ignorando tildes.
 */
function obtenerNombreJugador(correo) {
  if (!correo) return "";
  const correoClean = String(correo).toLowerCase().trim();
  const mapeo = obtenerMapeoJugadoresConfig();
  
  for (let nombre in mapeo) {
    const valor = mapeo[nombre];
    const lista = Array.isArray(valor) ? valor : [valor];
    
    for (let i = 0; i < lista.length; i++) {
      if (String(lista[i]).toLowerCase().trim() === correoClean) {
        return nombre;
      }
    }

    const nombreSinTilde = nombre.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    const usuarioCorreo = correoClean.split('@')[0].normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    if (usuarioCorreo === nombreSinTilde) {
      return nombre;
    }
  }
  return correoClean.split('@')[0].toUpperCase();
}

/**
 * ====================================================================
 * LEADERBOARD, MINIGRÁFICOS UNIVERSALES Y GRÁFICAS F1
 * ====================================================================
 */
function actualizarLeaderboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAnteriores = ss.getSheetByName('Carreras_Anteriores');
  let sheetLeaderboard = ss.getSheetByName('Leaderboard');

  if (!sheetLeaderboard) {
    sheetLeaderboard = ss.insertSheet('Leaderboard');
  }

  if (!sheetAnteriores || sheetAnteriores.getLastRow() < 2) return;

  const datosHistorial = sheetAnteriores.getRange(2, 1, sheetAnteriores.getLastRow() - 1, 4).getValues();
  const acumulado = {};

  datosHistorial.forEach(fila => {
    const correo = String(fila[2] || '').trim();
    const puntos = Number(fila[3]) || 0;
    if (correo) {
      const nombreJugador = obtenerNombreJugador(correo);
      acumulado[nombreJugador] = (acumulado[nombreJugador] || 0) + puntos;
    }
  });

  const ranking = Object.keys(acumulado).map(nombre => [nombre, acumulado[nombre]]);
  ranking.sort((a, b) => b[1] - a[1]);

  if (ranking.length === 0) return;

  const puntosLider = ranking[0][1];
  const filasTabla = [];

  for (let i = 0; i < ranking.length; i++) {
    const nombreJugador = ranking[i][0];
    const puntos = ranking[i][1];
    const pos = i + 1;

    let difAnterior = "";
    let difLider = "";

    if (i === 0) {
      difAnterior = "LÍDER";
      difLider = "LÍDER";
    } else {
      const puntosAnterior = ranking[i - 1][1];
      const gapAnt = puntosAnterior - puntos;
      const gapLid = puntosLider - puntos;

      difAnterior = gapAnt === 0 ? "0" : `-${gapAnt}`;
      difLider = gapLid === 0 ? "0" : `-${gapLid}`;
    }

    filasTabla.push([pos, nombreJugador, puntos, difAnterior, difLider, ""]);
  }

  const encabezados = ['Posición General', 'Participante', 'Puntos Acumulados', 'Dif. Ant.', 'Dif. Líder', 'Tendencia'];
  sheetLeaderboard.getRange(1, 1, 1, encabezados.length)
    .setValues([encabezados])
    .setBackground('#111827')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (sheetLeaderboard.getLastRow() > 1) {
    sheetLeaderboard.getRange(2, 1, sheetLeaderboard.getLastRow() - 1, sheetLeaderboard.getLastColumn()).clearContent();
  }

  if (filasTabla.length > 0) {
    const range = sheetLeaderboard.getRange(2, 1, filasTabla.length, 6);
    range.setValues(filasTabla);
    sheetLeaderboard.getRange(2, 1, filasTabla.length, 5).setHorizontalAlignment("center");
    sheetLeaderboard.setColumnWidth(6, 140);
    sheetLeaderboard.setRowHeights(2, filasTabla.length, 28);
  }

  generarDashboardCompletoF1(ranking);
}

function generarDashboardCompletoF1(ranking) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLeaderboard = ss.getSheetByName('Leaderboard');
  const sheetAnteriores = ss.getSheetByName('Carreras_Anteriores');
  if (!sheetLeaderboard || !sheetAnteriores || sheetAnteriores.getLastRow() < 2) return;

  let sheetEvolucion = ss.getSheetByName('Evolucion_Temporal');
  if (!sheetEvolucion) {
    sheetEvolucion = ss.insertSheet('Evolucion_Temporal');
  }
  sheetEvolucion.clear();

  const datosHist = sheetAnteriores.getRange(2, 1, sheetAnteriores.getLastRow() - 1, 4).getValues();

  const carrerasUnicas = [];
  datosHist.forEach(fila => {
    const gp = String(fila[1] || '').trim();
    if (gp && !carrerasUnicas.includes(gp)) {
      carrerasUnicas.push(gp);
    }
  });

  const jugadores = ranking.map(r => r[0]);

  const puntosPorGPyJugador = {};
  carrerasUnicas.forEach(gp => puntosPorGPyJugador[gp] = {});

  datosHist.forEach(fila => {
    const gp = String(fila[1] || '').trim();
    const correo = String(fila[2] || '').trim();
    const pts = Number(fila[3]) || 0;
    if (gp && correo) {
      const nombre = obtenerNombreJugador(correo);
      puntosPorGPyJugador[gp][nombre] = pts;
    }
  });

  const encabezadosEvolucion = ['Carrera / GP', ...jugadores];
  const filasEvolucion = [];
  const acumulador = {};
  jugadores.forEach(j => acumulador[j] = 0);

  carrerasUnicas.forEach(gp => {
    const fila = [gp];
    jugadores.forEach(j => {
      acumulador[j] += (puntosPorGPyJugador[gp][j] || 0);
      fila.push(acumulador[j]);
    });
    filasEvolucion.push(fila);
  });

  sheetEvolucion.getRange(1, 1, 1, encabezadosEvolucion.length).setValues([encabezadosEvolucion]);
  if (filasEvolucion.length > 0) {
    sheetEvolucion.getRange(2, 1, filasEvolucion.length, encabezadosEvolucion.length).setValues(filasEvolucion);
  }

  const totalGPs = carrerasUnicas.length;
  for (let i = 0; i < jugadores.length; i++) {
    const letraColumna = String.fromCharCode(66 + i);
    const formulaSparkline = '=SPARKLINE(Evolucion_Temporal!' + letraColumna + '$2:' + letraColumna + '$' + (totalGPs + 1) + ')';
    sheetLeaderboard.getRange(i + 2, 6).setFormula(formulaSparkline);
  }

  const graficasExistentes = sheetLeaderboard.getCharts();
  graficasExistentes.forEach(g => sheetLeaderboard.removeChart(g));

  const totalFilasLeaderboard = ranking.length + 1;

  const rangoBarras = sheetLeaderboard.getRange(1, 2, totalFilasLeaderboard, 2);
  const chartBarras = sheetLeaderboard.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(rangoBarras)
    .setNumHeaders(1)
    .setPosition(2, 8, 10, 10)
    .setOption('title', '🏎️ CAMPEONATO MUNDIAL DE QUINIELA F1')
    .setOption('titleTextStyle', { color: '#111827', fontSize: 13, bold: true })
    .setOption('hAxis', { title: 'Puntos Acumulados', minValue: 0 })
    .setOption('vAxis', { title: '' })
    .setOption('colors', ['#E10600'])
    .setOption('legend', { position: 'none' })
    .setOption('width', 620)
    .setOption('height', Math.max(280, totalFilasLeaderboard * 36))
    .build();

  sheetLeaderboard.insertChart(chartBarras);

  if (totalGPs > 0) {
    const filaInicioLineas = Math.max(16, totalFilasLeaderboard + 4);
    const rangoLineas = sheetEvolucion.getRange(1, 1, totalGPs + 1, jugadores.length + 1);

    const chartLineas = sheetLeaderboard.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(rangoLineas)
      .setNumHeaders(1)
      .setPosition(filaInicioLineas, 8, 10, 10)
      .setOption('title', '📈 TELEMETRÍA: EVOLUCIÓN GP TRAS GP')
      .setOption('titleTextStyle', { color: '#111827', fontSize: 13, bold: true })
      .setOption('curveType', 'function')
      .setOption('hAxis', { title: 'Gran Premio', slantedText: true })
      .setOption('vAxis', { title: 'Puntos Acumulados' })
      .setOption('legend', { position: 'top', textStyle: { fontSize: 11, bold: true } })
      .setOption('width', 620)
      .setOption('height', 380)
      .build();

    sheetLeaderboard.insertChart(chartLineas);
  }
}

/**
 * ====================================================================
 * VISTA OFICIAL DE PRONÓSTICOS (PARA CAPTURA CON FECHA Y HORA)
 * ====================================================================
 */
function generarVistaPronosticos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetResp = ss.getSheetByName('Respuestas_Formulario');

  let sheetCaptura = ss.getSheetByName('Captura_Pronosticos');
  const hojaActiva = ss.getActiveSheet();
  const tituloActiva = String(hojaActiva.getRange(1, 1).getValue() || '').toUpperCase();
  if (tituloActiva.includes('PRONÓSTICO') || tituloActiva.includes('PRONOSTICO')) {
    sheetCaptura = hojaActiva;
  }
  if (!sheetCaptura) {
    sheetCaptura = ss.insertSheet('Captura_Pronosticos');
  }

  if (!sheetResp || sheetResp.getLastRow() < 2) {
    notificarGlobal('⚠️ No hay pronósticos registrados en "Respuestas_Formulario".');
    return;
  }

  const encabezadosResp = sheetResp.getRange(1, 1, 1, sheetResp.getLastColumn()).getValues()[0];
  let idxCorreo = encabezadosResp.findIndex(h => {
    const txt = String(h).toLowerCase().trim();
    return txt.includes("correo") || txt.includes("email") || txt.includes("dirección");
  });
  if (idxCorreo === -1) idxCorreo = sheetResp.getLastColumn() - 1;

  const datosResp = sheetResp.getRange(2, 1, sheetResp.getLastRow() - 1, sheetResp.getLastColumn()).getValues();
  const pronosticosPorJugador = {};
  const formatoZona = Session.getScriptTimeZone();

  datosResp.forEach(fila => {
    const fechaRaw = fila[0];
    const correo = String(fila[idxCorreo] || '').trim();
    if (correo) {
      const nombreJugador = obtenerNombreJugador(correo);
      const picks = [];
      for (let c = 1; c < fila.length; c++) {
        if (c !== idxCorreo && picks.length < 10) {
          picks.push(fila[c]);
        }
      }

      let fechaFormateada = '—';
      if (fechaRaw instanceof Date) {
        fechaFormateada = Utilities.formatDate(fechaRaw, formatoZona, "dd/MM HH:mm:ss");
      } else if (fechaRaw) {
        fechaFormateada = String(fechaRaw);
      }

      pronosticosPorJugador[nombreJugador] = {
        fecha: fechaFormateada,
        picks: picks
      };
    }
  });

  let jugadoresColumnas = [];
  if (sheetCaptura.getLastColumn() > 1) {
    const fila2 = sheetCaptura.getRange(2, 2, 1, sheetCaptura.getLastColumn() - 1).getValues()[0];
    jugadoresColumnas = fila2.map(n => String(n).trim().toUpperCase()).filter(n => n !== "");
  }
  if (jugadoresColumnas.length === 0) {
    const mapeo = obtenerMapeoJugadoresConfig();
    const nombresRegistrados = Object.keys(mapeo);
    jugadoresColumnas = nombresRegistrados.length > 0 
      ? nombresRegistrados 
      : Object.keys(pronosticosPorJugador);
  }

  const totalCols = jugadoresColumnas.length + 1;

  sheetCaptura.getRange(1, 1, 1, totalCols).merge()
    .setValue('🏁 PRONÓSTICOS OFICIALES PARA LA CARRERA 🏁')
    .setBackground('#0B1325')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(11);

  sheetCaptura.getRange(2, 1).setValue('POSICIÓN')
    .setBackground('#0B1325')
    .setFontColor('#38BDF8')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  for (let c = 0; c < jugadoresColumnas.length; c++) {
    sheetCaptura.getRange(2, c + 2).setValue(jugadoresColumnas[c])
      .setBackground('#0B1325')
      .setFontColor('#38BDF8')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  const etiquetasPos = [['P1'], ['P2'], ['P3'], ['P4'], ['P5'], ['P6'], ['P7'], ['P8'], ['P9'], ['P10']];
  sheetCaptura.getRange(3, 1, 10, 1).setValues(etiquetasPos)
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheetCaptura.getRange(13, 1).setValue('HORA ENVÍO')
    .setBackground('#0B1325')
    .setFontColor('#FCD34D')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  for (let c = 0; c < jugadoresColumnas.length; c++) {
    const jugador = jugadoresColumnas[c];
    const colIndex = c + 2;

    let datosJugador = null;
    for (let j in pronosticosPorJugador) {
      if (j.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase() === jugador.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase()) {
        datosJugador = pronosticosPorJugador[j];
        break;
      }
    }

    if (datosJugador && datosJugador.picks && datosJugador.picks.length > 0) {
      const matrizPicks = [];
      for (let p = 0; p < 10; p++) {
        matrizPicks.push([datosJugador.picks[p] || '—']);
      }
      sheetCaptura.getRange(3, colIndex, 10, 1).setValues(matrizPicks)
        .setHorizontalAlignment('center')
        .setFontColor('#111827');

      sheetCaptura.getRange(13, colIndex).setValue(datosJugador.fecha)
        .setHorizontalAlignment('center')
        .setFontColor('#065F46')
        .setFontWeight('bold')
        .setFontSize(9);
    } else {
      const matrizVacia = Array(10).fill(['—']);
      sheetCaptura.getRange(3, colIndex, 10, 1).setValues(matrizVacia)
        .setHorizontalAlignment('center')
        .setFontColor('#9CA3AF');

      sheetCaptura.getRange(13, colIndex).setValue('⏳ PENDIENTE')
        .setHorizontalAlignment('center')
        .setFontColor('#DC2626')
        .setFontWeight('bold')
        .setFontSize(9);
    }
  }

  sheetCaptura.getRange(1, 1, 13, totalCols).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  sheetCaptura.setRowHeights(3, 11, 26);
  sheetCaptura.autoResizeColumns(1, totalCols);
  sheetCaptura.activate();

  notificarGlobal('📸 ¡Tabla lista!\n\nSe llenaron todos los pronósticos y la fecha/hora en la fila 13. Ya puedes tomar tu captura.');
}

function copiarPronosticosWhatsApp() {
  generarVistaPronosticos();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetResp = ss.getSheetByName('Respuestas_Formulario');
  if (!sheetResp || sheetResp.getLastRow() < 2) return;

  const encabezadosResp = sheetResp.getRange(1, 1, 1, sheetResp.getLastColumn()).getValues()[0];
  let idxCorreo = encabezadosResp.findIndex(h => {
    const txt = String(h).toLowerCase().trim();
    return txt.includes("correo") || txt.includes("email") || txt.includes("dirección");
  });
  if (idxCorreo === -1) idxCorreo = sheetResp.getLastColumn() - 1;

  const datosResp = sheetResp.getRange(2, 1, sheetResp.getLastRow() - 1, sheetResp.getLastColumn()).getValues();
  const ultimosPicks = {};

  datosResp.forEach(fila => {
    const correo = String(fila[idxCorreo] || '').trim();
    if (correo) {
      const nombreJugador = obtenerNombreJugador(correo);
      const picks = fila.filter((val, idx) => idx !== 0 && idx !== idxCorreo).slice(0, 10);
      ultimosPicks[nombreJugador] = picks;
    }
  });

  const mapeo = obtenerMapeoJugadoresConfig();
  const nombresRegistrados = Object.keys(mapeo);
  const listaJugadores = nombresRegistrados.length > 0 
    ? nombresRegistrados 
    : Object.keys(ultimosPicks);

  let mensaje = `🏁 *PRONÓSTICOS OFICIALES PARA LA CARRERA* 🏁\n`;
  mensaje += `🚦 _¡Luces fuera y que gane el mejor!_\n\n`;

  listaJugadores.forEach(jugador => {
    let picks = null;
    for (let j in ultimosPicks) {
      if (j.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase() === jugador.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase()) {
        picks = ultimosPicks[j];
        break;
      }
    }

    mensaje += `🏎️ *${jugador}:*\n`;
    if (picks && picks.length === 10) {
      const formatearPiloto = p => {
        const match = String(p).match(/\(([^)]+)\)/);
        return match ? match[1] : String(p).split(' ').pop();
      };

      mensaje += `🥇 ${formatearPiloto(picks[0])} | 🥈 ${formatearPiloto(picks[1])} | 🥉 ${formatearPiloto(picks[2])}\n`;
      mensaje += `P4-P10: ${picks.slice(3).map(formatearPiloto).join(' - ')}\n\n`;
    } else {
      mensaje += `❌ _No envió a tiempo_\n\n`;
    }
  });

  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert('📋 Texto para WhatsApp:\n\n' + mensaje);
  }
}

/**
 * ====================================================================
 * CONTROL PÚBLICO DE RECEPCIÓN (ESTADO DE ENVÍOS)
 * ====================================================================
 */
function actualizarEstadoPronosticos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetEstado = ss.getSheetByName('Estado_Pronosticos');
  const sheetRespuestas = ss.getSheetByName('Respuestas_Formulario');

  if (!sheetEstado) {
    sheetEstado = ss.insertSheet('Estado_Pronosticos');
  }

  if (sheetEstado.getLastRow() > 1) {
    sheetEstado.getRange(2, 1, sheetEstado.getLastRow() - 1, sheetEstado.getLastColumn()).clear();
  }

  const encabezados = ['Participante', 'Estado', 'Último Envío Registrado'];
  sheetEstado.getRange(1, 1, 1, encabezados.length)
    .setValues([encabezados])
    .setBackground('#111827')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const enviosPorNombre = {};
  const formatoZona = Session.getScriptTimeZone();

  if (sheetRespuestas && sheetRespuestas.getLastRow() > 1) {
    const encabezadosResp = sheetRespuestas.getRange(1, 1, 1, sheetRespuestas.getLastColumn()).getValues()[0];
    let idxCorreo = encabezadosResp.findIndex(h => {
      const txt = String(h).toLowerCase().trim();
      return txt.includes("correo") || txt.includes("email") || txt.includes("dirección");
    });
    if (idxCorreo === -1) idxCorreo = sheetRespuestas.getLastColumn() - 1;

    const datos = sheetRespuestas.getRange(2, 1, sheetRespuestas.getLastRow() - 1, sheetRespuestas.getLastColumn()).getValues();

    datos.forEach(fila => {
      const fecha = fila[0];
      const correo = String(fila[idxCorreo] || '').trim();
      if (correo) {
        const nombreJugador = obtenerNombreJugador(correo);
        enviosPorNombre[nombreJugador] = fecha;
      }
    });
  }

  const filas = [];
  const mapeo = obtenerMapeoJugadoresConfig();

  for (let jugador in mapeo) {
    if (enviosPorNombre[jugador]) {
      const fechaTxt = Utilities.formatDate(new Date(enviosPorNombre[jugador]), formatoZona, "dd/MM/yyyy HH:mm:ss");
      filas.push([jugador, '✅ ENVIADO', fechaTxt]);
      delete enviosPorNombre[jugador];
    } else {
      filas.push([jugador, '⏳ PENDIENTE', '—']);
    }
  }

  for (let noRegistrado in enviosPorNombre) {
    const fechaTxt = Utilities.formatDate(new Date(enviosPorNombre[noRegistrado]), formatoZona, "dd/MM/yyyy HH:mm:ss");
    filas.push([noRegistrado, '✅ ENVIADO', fechaTxt]);
  }

  filas.sort((a, b) => {
    if (a[1] === b[1]) return a[0].localeCompare(b[0]);
    return a[1] === '✅ ENVIADO' ? -1 : 1;
  });

  if (filas.length > 0) {
    const rangoDatos = sheetEstado.getRange(2, 1, filas.length, 3);
    rangoDatos.setValues(filas);
    rangoDatos.setHorizontalAlignment('center');

    for (let i = 0; i < filas.length; i++) {
      const celdaEstado = sheetEstado.getRange(i + 2, 2);
      if (filas[i][1] === '✅ ENVIADO') {
        celdaEstado.setBackground('#D1FAE5').setFontColor('#065F46').setFontWeight('bold');
      } else {
        celdaEstado.setBackground('#FEE2E2').setFontColor('#991B1B').setFontWeight('bold');
      }
    }
  }

  sheetEstado.autoResizeColumns(1, 3);
}

function verEstatusWhatsApp() {
  actualizarEstadoPronosticos();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEstado = ss.getSheetByName('Estado_Pronosticos');

  if (!sheetEstado || sheetEstado.getLastRow() < 2) {
    notificarGlobal('⚠️ No hay información de participantes.');
    return;
  }

  const datos = sheetEstado.getRange(2, 1, sheetEstado.getLastRow() - 1, 3).getValues();
  let enviados = [];
  let pendientes = [];
  const formatoZona = Session.getScriptTimeZone();

  datos.forEach(f => {
    const nombre = f[0];
    const estado = f[1];
    const fecha = f[2];

    if (estado === '✅ ENVIADO') {
      let horaCorta = '';
      if (fecha instanceof Date) {
        horaCorta = Utilities.formatDate(fecha, formatoZona, "HH:mm:ss");
      } else {
        const str = String(fecha || '');
        horaCorta = str.includes(' ') ? str.split(' ')[1] : str;
      }
      enviados.push(`✅ ${nombre} ${horaCorta && horaCorta !== '—' ? '(' + horaCorta + ')' : ''}`);
    } else {
      pendientes.push(`⏳ ${nombre}`);
    }
  });

  let mensaje = `🏁 *ESTATUS DE PRONÓSTICOS F1* 🏁\n\n`;
  mensaje += `🟢 *YA ENVIARON (${enviados.length}):*\n`;
  mensaje += enviados.length > 0 ? enviados.join('\n') : 'Ninguno aún';
  mensaje += `\n\n🔴 *PENDIENTES (${pendientes.length}):*\n`;
  mensaje += pendientes.length > 0 ? pendientes.join('\n') : '¡Todos listos! 🎉';

  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert('📋 Copia este texto para el grupo de WhatsApp:\n\n' + mensaje);
  } else {
    Logger.log(mensaje);
  }
}

/**
 * ====================================================================
 * UTILIDADES DE HISTÓRICO Y VISIBILIDAD
 * ====================================================================
 */
function coincidenciaFlexible(pronostico, nombreOficial, codigoOficial) {
  if (!pronostico) return false;
  const p = pronostico.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const n = nombreOficial.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const c = (codigoOficial || "").toLowerCase();

  return n.includes(p) || p.includes(n) || (c !== "" && p.includes(c));
}

function verificarEstadoAPI() {
  const url = 'https://api.jolpi.ca/ergast/f1/current/last/results.json';
  
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    const races = json.MRData.RaceTable.Races;

    if (!races || races.length === 0) {
      notificarGlobal('⚠️ La API aún no ha publicado los resultados de la carrera más reciente.');
      return;
    }

    const race = races[0];
    const ganador = race.Results[0].Driver.familyName;

    notificarGlobal(`🟢 API DISPONIBLE Y ACTUALIZADA\n\n` +
                    `🏎️ Carrera: ${race.raceName}\n` +
                    `📅 Temporada: ${race.season} (Ronda ${race.round})\n` +
                    `🥇 Ganador: ${race.Results[0].Driver.givenName} ${ganador}\n\n` +
                    `¡Ya puedes ejecutar la opción de descargar resultados!`);

  } catch (error) {
    notificarGlobal('❌ No se pudo conectar con la API: ' + error.toString());
  }
}

function migrarResultadosExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ss.getSheetByName("Hoja 1");
  const hojaDestino = ss.getSheetByName("Carreras_Anteriores");

  if (!hojaOrigen) {
    notificarGlobal("❌ No se encontró la pestaña 'Hoja 1'.");
    return;
  }
  if (!hojaDestino) {
    notificarGlobal("❌ No se encontró la pestaña 'Carreras_Anteriores'. Ejecuta primero '⚙️ 1. Inicializar Tablas y Estructura'.");
    return;
  }

  const datosOrigen = hojaOrigen.getDataRange().getValues();
  
  let filaEncabezado = -1;
  for (let i = 0; i < datosOrigen.length; i++) {
    const filaTexto = datosOrigen[i].map(c => String(c).toUpperCase().trim());
    if (filaTexto.includes("CIRCUITO")) {
      filaEncabezado = i;
      break;
    }
  }

  if (filaEncabezado === -1) {
    notificarGlobal("❌ No se encontró la fila con 'CIRCUITO' en 'Hoja 1'.");
    return;
  }

  const encabezados = datosOrigen[filaEncabezado];
  const indiceCircuito = encabezados.findIndex(c => String(c).toUpperCase().trim() === "CIRCUITO");
  const mapeo = obtenerMapeoJugadoresConfig();
  
  const participantes = [];
  for (let col = 0; col < encabezados.length; col++) {
    const nombre = String(encabezados[col]).trim().toUpperCase();
    if (col !== indiceCircuito && nombre !== "" && nombre !== "CARRERA" && nombre !== "TOTAL") {
      let email = `${nombre.toLowerCase()}@quiniela.com`;
      if (mapeo[nombre]) {
        const val = mapeo[nombre];
        email = Array.isArray(val) ? val[0] : val;
      }
      
      participantes.push({ nombre: nombre, columna: col, email: email });
    }
  }

  if (participantes.length === 0) {
    notificarGlobal("⚠️ No se encontraron participantes en 'Hoja 1'.");
    return;
  }

  const filasAInsertar = [];
  const ANIO_TEMPORADA = 2025;

  for (let f = filaEncabezado + 1; f < datosOrigen.length; f++) {
    const carrera = String(datosOrigen[f][indiceCircuito] || '').trim();
    if (!carrera || carrera.toLowerCase().includes("total")) break;

    participantes.forEach(p => {
      const puntos = datosOrigen[f][p.columna];
      if (puntos !== "" && puntos !== null && !isNaN(puntos)) {
        filasAInsertar.push([
          ANIO_TEMPORADA,
          carrera,
          p.email,
          Number(puntos)
        ]);
      }
    });
  }

  if (filasAInsertar.length > 0) {
    if (hojaDestino.getLastRow() > 1) {
      hojaDestino.getRange(2, 1, hojaDestino.getLastRow() - 1, 4).clearContent();
    }

    hojaDestino.getRange(2, 1, filasAInsertar.length, 4).setValues(filasAInsertar);
    actualizarLeaderboard();

    notificarGlobal(`✅ ¡MIGRACIÓN EXITOSA!\n\n` +
                    `• Registros migrados: ${filasAInsertar.length}\n` +
                    `• Participantes: ${participantes.map(p => p.nombre).join(', ')}\n` +
                    `• Tabla 'Leaderboard' actualizada con diferencias.`);
  } else {
    notificarGlobal("⚠️ No se encontraron puntos válidos.");
  }
}

function ocultarRespuestas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Respuestas_Formulario');
  if (sheet) {
    sheet.hideSheet();
    notificarGlobal('🔒 La pestaña "Respuestas_Formulario" ha sido ocultada.');
  }
}

function mostrarRespuestas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Respuestas_Formulario');
  if (sheet) {
    sheet.showSheet();
    notificarGlobal('🔓 La pestaña "Respuestas_Formulario" ahora es visible.');
  }
}
/**
 * ====================================================================
 * SIMULACIÓN MONTE CARLO: PROBABILIDAD DE CAMPEONATO
 * ====================================================================
 */
function calcularProbabilidadesCampeonato() {
  const TOTAL_GPS_TEMPORADA = 24; // Calendario oficial F1
  const SIMULACIONES = 1000;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAnteriores = ss.getSheetByName('Carreras_Anteriores');
  let sheetProb = ss.getSheetByName('Probabilidades');

  if (!sheetProb) {
    sheetProb = ss.insertSheet('Probabilidades');
  }

  if (!sheetAnteriores || sheetAnteriores.getLastRow() < 2) {
    notificarGlobal('⚠️ No hay suficientes datos en Carreras_Anteriores para simular.');
    return;
  }

  const datos = sheetAnteriores.getRange(2, 1, sheetAnteriores.getLastRow() - 1, 4).getValues();

  // 1. Extraer GPs únicos disputados
  const gpsUnicos = new Set();
  const puntosPorJugador = {}; // nombre -> [pts_gp1, pts_gp2, ...]

  datos.forEach(fila => {
    const gp = String(fila[1] || '').trim();
    const correo = String(fila[2] || '').trim();
    const pts = Number(fila[3]) || 0;

    if (gp && correo) {
      gpsUnicos.add(gp);
      const nombre = obtenerNombreJugador(correo);
      if (!puntosPorJugador[nombre]) puntosPorJugador[nombre] = [];
      puntosPorJugador[nombre].push(pts);
    }
  });

  const gpsDisputados = gpsUnicos.size;
  const gpsRestantes = Math.max(0, TOTAL_GPS_TEMPORADA - gpsDisputados);

  // 2. Calcular estadísticas (Media y Desviación estándar) de cada uno
  const statsJugadores = {};
  const jugadores = Object.keys(puntosPorJugador);

  jugadores.forEach(j => {
    const historial = puntosPorJugador[j];
    const totalPuntos = historial.reduce((a, b) => a + b, 0);
    const media = totalPuntos / historial.length;
    
    // Varianza y desviación
    const varianza = historial.reduce((sum, pts) => sum + Math.pow(pts - media, 2), 0) / historial.length;
    const desv = Math.sqrt(varianza) || 5; // Mínimo 5 para dar variabilidad

    statsJugadores[j] = {
      actuales: totalPuntos,
      media: media,
      desv: desv,
      victoriasSimuladas: 0
    };
  });

  // 3. Simulación Monte Carlo (1.000 temporadas restantes)
  for (let s = 0; s < SIMULACIONES; s++) {
    const puntosFinalesTemp = {};

    jugadores.forEach(j => {
      let acumuladoSimulado = statsJugadores[j].actuales;

      for (let g = 0; g < gpsRestantes; g++) {
        // Generador normal aleatorio (Box-Muller)
        const u1 = Math.random() || 0.0001;
        const u2 = Math.random() || 0.0001;
        const z = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);

        // Puntos simulados con piso en 0 y techo en 100
        let ptsCarrera = Math.round(statsJugadores[j].media + (z * statsJugadores[j].desv));
        ptsCarrera = Math.max(0, Math.min(100, ptsCarrera));

        acumuladoSimulado += ptsCarrera;
      }

      puntosFinalesTemp[j] = acumuladoSimulado;
    });

    // Encontrar el ganador de esta simulación
    let mejorPuntaje = -1;
    let ganador = '';
    jugadores.forEach(j => {
      if (puntosFinalesTemp[j] > mejorPuntaje) {
        mejorPuntaje = puntosFinalesTemp[j];
        ganador = j;
      }
    });

    if (ganador) {
      statsJugadores[ganador].victoriasSimuladas++;
    }
  }

  // 4. Preparar tabla de resultados
  const filasSalida = [];
  jugadores.forEach(j => {
    const prob = (statsJugadores[j].victoriasSimuladas / SIMULACIONES) * 100;
    filasSalida.push([
      j,
      statsJugadores[j].actuales,
      Math.round(statsJugadores[j].media * 10) / 10,
      prob / 100 // En decimal para formato porcentaje (0.45 = 45%)
    ]);
  });

  // Ordenar por probabilidad descendente
  filasSalida.sort((a, b) => b[3] - a[3]);

  // Escribir en la hoja
  sheetProb.clear();
  const encabezados = ['Participante', 'Puntos Actuales', 'Promedio Pts/GP', 'Probabilidad de Ganar'];
  sheetProb.getRange(1, 1, 1, encabezados.length)
    .setValues([encabezados])
    .setBackground('#111827')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (filasSalida.length > 0) {
    sheetProb.getRange(2, 1, filasSalida.length, 4).setValues(filasSalida);
    sheetProb.getRange(2, 1, filasSalida.length, 3).setHorizontalAlignment('center');
    sheetProb.getRange(2, 4, filasSalida.length, 1).setNumberFormat('0.0%').setHorizontalAlignment('center');
  }

  notificarGlobal(`🎲 Simulación Monte Carlo Completada:\n\n` +
                  `• GPs disputados: ${gpsDisputados} / ${TOTAL_GPS_TEMPORADA}\n` +
                  `• GPs restantes: ${gpsRestantes}\n` +
                  `• 1.000 temporadas simuladas con éxito.\n` +
                  `• Revisa la pestaña 'Probabilidades'.`);
}
/**
 * Registra un Gran Premio desierto (0 puntos para todos los participantes oficiales).
 * Actualiza Carreras_Anteriores, Leaderboard y Probabilidades.
 */
function registrarCarreraEnBlanco(nombreGP) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAnteriores = ss.getSheetByName('Carreras_Anteriores');
  const mapeo = obtenerMapeoJugadoresConfig();
  const anioActual = new Date().getFullYear();
  const filasCero = [];

  for (let jugador in mapeo) {
    const valor = mapeo[jugador];
    const email = Array.isArray(valor) ? valor[0] : valor;
    filasCero.push([anioActual, nombreGP, email, 0]);
  }

  if (filasCero.length > 0) {
    // Si ya existía este GP en el historial, limpiarlo para no duplicar
    if (sheetAnteriores.getLastRow() > 1) {
      const datosAnt = sheetAnteriores.getRange(2, 1, sheetAnteriores.getLastRow() - 1, 4).getValues();
      const filtrados = datosAnt.filter(r => !(r[0] == anioActual && String(r[1]).toLowerCase().trim() == nombreGP.toLowerCase().trim()));
      sheetAnteriores.getRange(2, 1, sheetAnteriores.getLastRow() - 1, 4).clearContent();
      if (filtrados.length > 0) {
        sheetAnteriores.getRange(2, 1, filtrados.length, 4).setValues(filtrados);
      }
    }

    sheetAnteriores.getRange(sheetAnteriores.getLastRow() + 1, 1, filasCero.length, 4).setValues(filasCero);
    
    // Actualizar Leaderboard y recalcular Monte Carlo
    actualizarLeaderboard();
    if (typeof calcularProbabilidadesCampeonato === 'function') {
      calcularProbabilidadesCampeonato();
    }

    notificarGlobal(`✅ Gran Premio "${nombreGP}" registrado exitosamente con 0 puntos para todos.\n\n` +
                    `• Leaderboard actualizado.\n` +
                    `• Telemetría actualizada.\n` +
                    `• Probabilidades recalculadas (se restó 1 GP al calendario).`);
  }
}