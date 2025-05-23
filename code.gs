/**
 * @OnlyCurrentDoc
 */

const PLACEHOLDER_FUNCTION_NAME = "GEMINI_MHW"; // Nombre interno de la función que se busca en las celdas
const MODEL_NAME = "gemini-2.0-flash"; // Modelo de Gemini a utilizar. ¡Asegúrate que tu API Key tenga acceso!

// --- Delimitadores para el procesamiento en lote único UNIFICADO ---
const UNIFIED_JOB_ID_PREFIX = "TASK_ID_"; // Prefijo para identificar cada sub-tarea
const UNIFIED_RESPONSE_SEPARATOR = "~~~~~GEMINI_TASK_SEPARATOR~~~~~"; // Lo que le pediremos a Gemini que use entre respuestas

/**
 * Crea un menú personalizado en la hoja de cálculo al abrirla.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('✨ GEMINI_MHW ✨') // Título del menú principal
      .addItem('1. Configurar API Key', 'showApiKeyConfigInstructions')
      .addSeparator()
      .addItem('Procesar Lote con =' + PLACEHOLDER_FUNCTION_NAME + '(...)', 'processSingleUnifiedBatchRequest')
      .addSeparator()
      .addItem('Ayuda en el uso de ' + PLACEHOLDER_FUNCTION_NAME, 'showHelpDialog') // Ajustado para que quepa mejor
      .addToUi();
}

/**
 * Muestra instrucciones para configurar la API Key.
 * También crea la hoja 'config' si no existe.
 */
function showApiKeyConfigInstructions() {
    getApiKeyFromConfig();
}

/**
 * Función auxiliar para manejar la configuración de la API Key.
 * @return {string|null} La API Key si es válida, o un string de error que comienza con "Error:".
 */
function getApiKeyFromConfig() {
  const configSheetName = "config";
  const apiKeyCell = "A1";
  const apiKeyPlaceholder = "[INGRESA_TU_API_KEY_DE_GEMINI_AQUI]";

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = spreadsheet.getSheetByName(configSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!configSheet) {
    Logger.log("Creando hoja '" + configSheetName + "'...");
    configSheet = spreadsheet.insertSheet(configSheetName);
    configSheet.getRange(apiKeyCell).setValue(apiKeyPlaceholder);
    configSheet.getRange(apiKeyCell).setFontWeight("bold").setFontColor("red").setHorizontalAlignment("left");
    configSheet.getRange("B1").setValue("<- Ingresa tu API Key de Gemini aquí y presiona Enter.");
    configSheet.getRange("A2").setValue("Puedes obtener tu API Key desde Google AI Studio.");
    configSheet.autoResizeColumn(1);
    configSheet.autoResizeColumn(2);

    ui.alert(
      'Configuración Necesaria',
      'Se ha creado la hoja "' + configSheetName + '". Por favor, ingresa tu API Key de Gemini en la celda ' + apiKeyCell + ' de esa hoja antes de usar la herramienta.',
      ui.ButtonSet.OK
    );
    Logger.log("Hoja '" + configSheetName + "' creada. Usuario notificado para ingresar API Key.");
    return "Error: Configura tu API Key en la hoja '" + configSheetName + "', celda " + apiKeyCell;
  }

  var apiKey = configSheet.getRange(apiKeyCell).getValue().toString().trim();
  Logger.log("Leyendo API Key de '" + configSheetName + "!" + apiKeyCell + "': " + (apiKey ? "Valor encontrado" : "Vacío o Placeholder"));

  if (!apiKey || apiKey === apiKeyPlaceholder || apiKey === "") {
    if (configSheet.getRange(apiKeyCell).getValue() !== apiKeyPlaceholder) {
         configSheet.getRange(apiKeyCell).setValue(apiKeyPlaceholder).setFontWeight("bold").setFontColor("red");
    }
    if (configSheet.getRange("B1").getValue() !== "<- Ingresa tu API Key de Gemini aquí y presiona Enter.") {
        configSheet.getRange("B1").setValue("<- Ingresa tu API Key de Gemini aquí y presiona Enter.");
    }

    ui.alert(
      'API Key No Configurada',
      'Por favor, ingresa tu API Key de Gemini en la hoja "' + configSheetName + '", celda ' + apiKeyCell + ', y presiona Enter.',
      ui.ButtonSet.OK
    );
    Logger.log("API Key no encontrada o es el placeholder. Usuario notificado.");
    return "Error: Ingresa tu API Key en '" + configSheetName + "!" + apiKeyCell + "'";
  }

  Logger.log("API Key encontrada y parece válida.");
  return apiKey;
}

/**
 * Realiza la llamada única a la API de Gemini con el prompt unificado.
 * @param {string} apiKey La API Key de Gemini.
 *   El prompt completo y estructurado a enviar.
 * @return {string} La respuesta de Gemini o un mensaje de error que comienza con "Error:".
 */
function _callUnifiedGeminiApi(apiKey, fullPrompt) {
  if (!apiKey || apiKey.startsWith("Error:")) {
    Logger.log("Error de API Key previo a la llamada: " + apiKey);
    return apiKey;
  }
  if (!fullPrompt) {
    Logger.log("Error: El prompt unificado está vacío.");
    return "Error: El prompt unificado no puede estar vacío.";
  }

  var apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;
  var payload = {
    "contents": [{"parts": [{"text": fullPrompt}]}],
    // Opcional: Podrías añadir aquí generationConfig si quieres controlar más la salida
    // "generationConfig": {
    //   "temperature": 0.7,
    //   "maxOutputTokens": 8192, // Ajustar según necesidad y modelo
    // }
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    Logger.log(`Enviando Prompt Unificado a API (${MODEL_NAME}). Longitud: ${fullPrompt.length}. Inicio: ${fullPrompt.substring(0,250)}...`);
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    Logger.log(`Respuesta API Unificada - Código: ${responseCode}. Longitud: ${responseBody.length}. Inicio: ${responseBody.substring(0,250)}...`);

    if (responseCode === 200) {
      var jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
          jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
          jsonResponse.candidates[0].content.parts.length > 0 &&
          typeof jsonResponse.candidates[0].content.parts[0].text === 'string') {
        return jsonResponse.candidates[0].content.parts[0].text;
      } else if (jsonResponse.promptFeedback && jsonResponse.promptFeedback.blockReason) {
        Logger.log("Error: Contenido bloqueado por Gemini. Razón: " + jsonResponse.promptFeedback.blockReason);
        return `Error: Contenido bloqueado por Gemini (${jsonResponse.promptFeedback.blockReason}). Revisa tu prompt y las políticas de Gemini.`;
      }
      Logger.log("Error: Estructura de respuesta inesperada de Gemini. Body: " + responseBody);
      return "Error: Respuesta inesperada de Gemini (formato interno). Consulta los Logs.";
    } else {
      let apiErrorMsg = `Código de respuesta: ${responseCode}. Cuerpo: ${responseBody}`;
      try {
          let errorJson = JSON.parse(responseBody);
          if (errorJson.error && errorJson.error.message) {
              apiErrorMsg = errorJson.error.message;
          }
      } catch(e) { /* no hacer nada si no es JSON, usar el mensaje completo */ }
      Logger.log(`Error de API (${responseCode}): ${apiErrorMsg}`);
      return `Error de API de Gemini (${responseCode}): ${apiErrorMsg}`;
    }
  } catch (e) {
    Logger.log(`Excepción crítica durante la llamada a la API: ${e.toString()}\nStack: ${e.stack}`);
    return `Error crítico en llamada a API: ${e.toString()}`;
  }
}

/**
 * Parsea los argumentos de la fórmula
 * @return {object} {prompt: string, info_complementaria: string, error: string|null}
 */
function _parseFormulaArguments(argsString, sheet, currentCellA1ForLog) {
  var prompt = "", info_complementaria = "", error = null;
  var argsArray = [];
  var inQuote = false;
  var currentArg = "";
  var parenthesisDepth = 0;

  for (let i = 0; i < argsString.length; i++) {
    const char = argsString[i];
    if (char === '"') {
      if (inQuote && i + 1 < argsString.length && argsString[i+1] === '"') {
        currentArg += '""';
        i++;
        continue;
      }
      inQuote = !inQuote;
    }

    if (!inQuote) {
      if (char === '(') parenthesisDepth++;
      if (char === ')') parenthesisDepth--;
      if ((char === ',' || char === ';') && parenthesisDepth === 0) {
        argsArray.push(currentArg.trim());
        currentArg = "";
        continue;
      }
    }
    currentArg += char;
  }
  argsArray.push(currentArg.trim());

  if (argsArray.length > 0 && argsArray[0]) {
    let firstArg = argsArray[0];
    if (firstArg.startsWith('"') && firstArg.endsWith('"')) {
      prompt = firstArg.substring(1, firstArg.length - 1).replace(/""/g, '"');
    } else {
      error = `#ERROR: El primer argumento (prompt) debe ser un texto entre comillas. Ej: "Analiza esto"`;
    }
  } else {
    error = `#ERROR: Falta el prompt. Uso: =${PLACEHOLDER_FUNCTION_NAME}("tu prompt", [info_adicional])`;
  }

  if (!error && argsArray.length > 1 && argsArray[1]) {
    let secondArg = argsArray[1];
    if (secondArg.startsWith('"') && secondArg.endsWith('"')) {
      info_complementaria = secondArg.substring(1, secondArg.length - 1).replace(/""/g, '"');
    } else {
      try {
        let referencedCell = sheet.getRange(secondArg);
        info_complementaria = referencedCell.getDisplayValue(); // Usar getDisplayValue para obtener el valor como se muestra
      } catch (e) {
        error = `#REF! Error al leer la celda '${secondArg}' para información complementaria.`;
        Logger.log(`Error de referencia de celda en ${currentCellA1ForLog}: ${secondArg} - ${e.message}`);
      }
    }
  }

  if (error) {
    Logger.log(`Error parseando argumentos para ${currentCellA1ForLog}: ${error}. Input: "${argsString}"`);
  }
  return { prompt, info_complementaria, error };
}


/**
 * Procesa todas las fórmulas GEMINI_MHW en la hoja activa
 * utilizando una ÚNICA llamada a la API de Gemini.
 */
function processSingleUnifiedBatchRequest() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var formulas = dataRange.getFormulas(); // Volver a A1 si el parser de argumentos espera eso

  Logger.log(`INICIO: Proceso de Lote con INFERENCIA ÚNICA en hoja: ${sheet.getName()}`);
  var apiKey = getApiKeyFromConfig();
  if (apiKey.startsWith("Error:")) {
    Logger.log("Proceso detenido por error de API Key.");
    // No es necesario un ui.alert aquí, getApiKeyFromConfig ya lo hace.
    return;
  }

  var tasks = [];
  const startRowData = dataRange.getRow();
  const startColData = dataRange.getColumn();
  const formulaRegex = new RegExp(`^=${PLACEHOLDER_FUNCTION_NAME}\\s*\\((.*)\\)\\s*$`, "i");

  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      let formulaText = formulas[r][c];
      if (formulaText && formulaRegex.test(formulaText)) {
        let currentRowInSheet = startRowData + r;
        let currentColInSheet = startColData + c;
        let cellA1 = sheet.getRange(currentRowInSheet, currentColInSheet).getA1Notation();

        let argsString = formulaText.match(formulaRegex)[1].trim();
        let parsedArgs = _parseFormulaArguments(argsString, sheet, cellA1);

        if (parsedArgs.error) {
          sheet.getRange(currentRowInSheet, currentColInSheet).setValue(parsedArgs.error);
          continue;
        }

        let taskId = `${UNIFIED_JOB_ID_PREFIX}${cellA1.replace(/[^a-zA-Z0-9_]/g, "_")}`;
        tasks.push({
          taskId: taskId,
          prompt: parsedArgs.prompt,
          info: parsedArgs.info_complementaria,
          row: currentRowInSheet,
          col: currentColInSheet,
          cellA1: cellA1
        });
        // No mostrar "Preparando lote..." individualmente aquí para una UI más limpia,
        // se mostrará un toast general.
      }
    }
  }

  if (tasks.length === 0) {
    ui.alert("Información", "No se encontraron fórmulas válidas como '=" + PLACEHOLDER_FUNCTION_NAME + "(...)' para procesar en la hoja activa.", ui.ButtonSet.OK);
    Logger.log("FIN: No hay tasks para procesar en lote único.");
    return;
  }

  // Actualizar celdas a "Procesando..." antes de la llamada larga
  tasks.forEach(task => sheet.getRange(task.row, task.col).setValue("⏳ Procesando..."));
  SpreadsheetApp.flush(); // Forzar la actualización de la UI

  Logger.log(`Recopiladas ${tasks.length} tasks para el lote único.`);
  ss.toast(`Enviando ${tasks.length} tareas en un solo lote a Gemini...`, `Procesando Lote con ${MODEL_NAME}`, -1); // -1 para que dure hasta que se quite

  let megaPromptParts = [
    "Eres un asistente avanzado de procesamiento por lotes. Te proporcionaré una lista de tareas individuales, cada una con un identificador único (taskId).",
    "Debes procesar CADA tarea según su 'Tarea Principal' y la 'Información de Apoyo' (si se proporciona).",
    `IMPORTANTE: Tu respuesta DEBE seguir ESTRICTAMENTE el siguiente formato para CADA tarea resuelta:`,
    `Comienza con el identificador de la tarea original (ej. '${UNIFIED_JOB_ID_PREFIX}Hoja1_A1').`,
    `Sigue con '::' (dos puntos).`,
    `Luego, la respuesta o resultado para esa tarea específica.`,
    `Finalmente, CADA respuesta de tarea (incluida la última) DEBE terminar con el delimitador EXACTO: "${UNIFIED_RESPONSE_SEPARATOR}"`,
    "No incluyas NINGÚN otro texto, saludo, introducción, resumen, explicación o comentario fuera de este formato estructurado de 'taskId::respuesta${UNIFIED_RESPONSE_SEPARATOR}'.",
    "Si una tarea específica no puede ser completada o resulta en un error por tu parte, para ESA tarea, devuelve 'taskId::Error: [descripción breve del error interno]' seguido del delimitador.",
    "\nA continuación, la lista de tareas a procesar:\n"
  ];

  tasks.forEach(task => {
    megaPromptParts.push(`\nTaskId: ${task.taskId}`);
    megaPromptParts.push(`Tarea Principal: ${task.prompt}`);
    if (task.info && String(task.info).trim() !== "") {
      megaPromptParts.push(`Información de Apoyo: ${task.info}`);
    }
  });
  let megaPrompt = megaPromptParts.join("\n");

  Logger.log(`Mega-Prompt construido (Longitud total: ${megaPrompt.length}). Primeros 500 caracteres: ${megaPrompt.substring(0,500)}`);

  // --- AJUSTE EN LA ADVERTENCIA DE LÍMITE DE TOKENS ---
  const FLASH_MODEL_CHAR_WARNING_THRESHOLD = 700000; // ~700k caracteres, mucho más alto
  const OTHER_MODEL_CHAR_WARNING_THRESHOLD = 28000;  // Mantenemos el umbral bajo para otros modelos

  let charWarningThreshold = OTHER_MODEL_CHAR_WARNING_THRESHOLD;
  if (MODEL_NAME.toLowerCase().includes("flash")) {
      charWarningThreshold = FLASH_MODEL_CHAR_WARNING_THRESHOLD;
  }

  if (megaPrompt.length > charWarningThreshold) {
      Logger.log(`ADVERTENCIA: El Mega-Prompt (aprox. ${megaPrompt.length} caracteres) es muy largo y podría aproximarse a los límites de tokens de ${MODEL_NAME}. Umbral de advertencia: ${charWarningThreshold} caracteres.`);
      ui.alert("Advertencia de Límite de Tokens", `El conjunto de ${tasks.length} tareas es muy grande (${megaPrompt.length} caracteres aprox.) y podría aproximarse a los límites de ${MODEL_NAME}. Si la operación falla o da resultados incompletos, intente con menos tareas o divida el trabajo.`, ui.ButtonSet.OK);
  }
  // --- FIN DEL AJUSTE ---

  let unifiedResponseText = _callUnifiedGeminiApi(apiKey, megaPrompt);

  if (unifiedResponseText.startsWith("Error:")) {
    Logger.log(`Error crítico en la llamada del Mega-Prompt: ${unifiedResponseText}`);
    ui.alert("Error en Procesamiento de Lote Único", `La llamada a Gemini para el lote falló: ${unifiedResponseText}. Las celdas no fueron actualizadas. Por favor, revise los Logs de Apps Script para más detalles.`, ui.ButtonSet.OK);
    tasks.forEach(task => sheet.getRange(task.row, task.col).setValue(`❌ Error API Lote: ${unifiedResponseText.substring(0, 100)}...`)); // Mostrar parte del error
    ss.toast("Error en API del Lote", "Fallo el procesamiento", 7);
    return;
  }

  Logger.log(`Respuesta Unificada de Gemini recibida (Longitud: ${unifiedResponseText.length}). Primeros 500 caracteres: ${unifiedResponseText.substring(0,500)}`);

  let responseSegments = unifiedResponseText.split(UNIFIED_RESPONSE_SEPARATOR);
  let updatedCount = 0;
  let errorsInResponseParsing = [];

  let tasksMap = new Map(tasks.map(task => [task.taskId, task]));

  responseSegments.forEach((segment, index) => {
    segment = segment.trim();
    if (!segment) return;

    let match = segment.match(new RegExp(`^(${UNIFIED_JOB_ID_PREFIX}[^:]+)::(.*)$`, "s")); // 's' flag para que . coincida con saltos de línea

    if (match && match[1] && typeof match[2] !== 'undefined') {
      let responseTaskId = match[1].trim();
      let taskResultText = match[2].trim();

      let task = tasksMap.get(responseTaskId);

      if (task) {
        sheet.getRange(task.row, task.col).setValue(taskResultText);
        Logger.log(`OK: Celda ${task.cellA1} (TaskId: ${responseTaskId}) actualizada.`); // No mostrar el contenido para logs más limpios
        updatedCount++;
        tasksMap.delete(responseTaskId);
      } else {
        let warnMsg = `ADVERTENCIA: Se recibió respuesta para un TaskId desconocido o ya procesado: ${responseTaskId}. Respuesta: "${String(taskResultText).substring(0,70)}..."`;
        Logger.log(warnMsg);
        errorsInResponseParsing.push(warnMsg);
      }
    } else {
      let warnMsg = `ADVERTENCIA: El segmento ${index + 1} de la respuesta de Gemini no coincide con el formato esperado (taskId::respuesta): "${segment.substring(0,100)}..."`;
      Logger.log(warnMsg);
      errorsInResponseParsing.push(warnMsg);
    }
  });

  if (tasksMap.size > 0) {
      Logger.log(`ERROR: ${tasksMap.size} tareas enviadas no tuvieron una respuesta mapeada desde Gemini:`);
      tasksMap.forEach(task => {
          Logger.log(`  - TaskId: ${task.taskId} (Celda: ${task.cellA1})`);
          sheet.getRange(task.row, task.col).setValue("#ERROR: Sin respuesta de Gemini en el lote");
          errorsInResponseParsing.push(`Tarea no resuelta: ${task.cellA1} (ID: ${task.taskId})`);
      });
  }

  if (errorsInResponseParsing.length > 0) {
      Logger.log("RESUMEN DE ERRORES DE PARSEO/MAPEADO: " + errorsInResponseParsing.join("\n - "));
  }

  let finalMessage = `Proceso de Lote Optimizado Completado.\nSe enviaron ${tasks.length} tareas a Gemini en una sola llamada.\nSe actualizaron ${updatedCount} celdas.`;
  if (tasks.length !== updatedCount || errorsInResponseParsing.length > 0) {
    let numIncidents = (tasks.length - updatedCount) + errorsInResponseParsing.filter(e => !e.startsWith("ADVERTENCIA")).length; // Contar errores reales, no solo advertencias
    if (numIncidents > 0) {
      finalMessage += `\n${numIncidents} ${numIncidents === 1 ? 'incidencia encontrada' : 'incidencias encontradas'} (respuestas no mapeadas o errores de formato). Revisa los Logs de Apps Script y las celdas marcadas con #ERROR para más detalles.`;
    }
  }
  ui.alert("Resultado del Lote Optimizado", finalMessage, ui.ButtonSet.OK);
  ss.toast(`¡Lote Optimizado Terminado! ${updatedCount}/${tasks.length} celdas actualizadas.`, "Éxito del Procesamiento", 10);
  Logger.log(`FIN: Proceso Lote Único. ${updatedCount} de ${tasks.length} tasks actualizadas.`);
}


/**
 * Muestra un diálogo de ayuda con instrucciones sobre cómo usar la herramienta.
 */
function showHelpDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('HelpDialog')
      .setWidth(550)
      .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Ayuda / Cómo Usar ✨ GEMINI_MHW ✨');
}
