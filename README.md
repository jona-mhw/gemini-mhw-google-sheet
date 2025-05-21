# GEMINI_MHW - Gemini como Fórmula en Google Sheets (MVP)

**Mira la demo en YouTube:**
[▶️ Demo de GEMINI_MHW en YouTube](https://www.youtube.com/watch?v=UdV_kD3hUWU)

**GEMINI_MHW** es un script de Google Apps Script (MVP) que te permite usar la potencia de Gemini AI como si fuera una fórmula nativa dentro de tus Google Sheets, similar a la funcionalidad `=AI()` que Google ha anunciado pero que aún podría no estar disponible para todos los usuarios (a fecha de 21 de mayo de 2025). Este script optimiza el proceso al enviar todas las solicitudes de una hoja en **una única inferencia** a la API de Gemini.

Este proyecto está orientado a usuarios que desean realizar múltiples inferencias de Gemini sobre datos en sus hojas de cálculo de manera sencilla y eficiente.

## ✨ Características Principales (MVP)

*   **Gemini como Fórmula:** Usa `=GEMINI_MHW("tu prompt", [celda_de_apoyo_o_texto])` para interactuar con Gemini.
*   **Procesamiento por Lotes Eficiente:** Todas las fórmulas `=GEMINI_MHW(...)` en la hoja activa se procesan en una sola llamada a la API.
*   **Menú Integrado:** Fácil acceso para configurar tu API Key y ejecutar el procesamiento.
*   **Modelo Configurable:** Ajusta el modelo de Gemini (ej. `gemini-2.0-flash`) en la constante `MODEL_NAME` del script.
*   **Diálogo de Ayuda:** Instrucciones de uso accesibles desde el menú.

## 🚀 Requisitos Previos

1.  Una cuenta de Google.
2.  Una **API Key de Gemini**. Puedes obtenerla desde [Google AI Studio](https://makersuite.google.com/app/apikey).

## 🛠️ Instalación y Configuración

1.  **Abre tu Google Sheet.**
2.  Ve a `Extensiones > Apps Script`.
3.  **Crea el archivo de script principal:**
    *   Renombra o vacía el archivo `Código.gs` por defecto. Copia todo el contenido del archivo [`code.gs`](./code.gs) de este repositorio y pégalo en tu archivo `code.gs`.
4.  **Crea el archivo de diálogo HTML:**
    *   En el editor de Apps Script, haz clic en `+ > HTML`. Nombra el archivo `HelpDialog.html` (exactamente así).
    *   Copia todo el contenido del archivo [`HelpDialog.html`](./HelpDialog.html) de este repositorio y pégalo en tu archivo `HelpDialog.html`.
5.  **Guarda los cambios** en el editor de Apps Script.
6.  **Vuelve a tu Google Sheet y recarga la página.** Debería aparecer un nuevo menú `✨ GEMINI_MHW ✨`.
7.  **Configura tu API Key:**
    *   Ve al menú `✨ GEMINI_MHW ✨ > 1. Configurar API Key`.
    *   Ingresa tu API Key de Gemini en la celda `A1` de la hoja `config` que se creará.

## 💡 Uso

1.  **Escribe tus prompts como fórmulas:**
    En cualquier celda, escribe:
    `=GEMINI_MHW("Tu instrucción para Gemini", A1)`
    O:
    `=GEMINI_MHW("Analiza este texto:", "Este es el texto a analizar.")`
2.  **Procesa el Lote:**
    *   Desde el menú `✨ GEMINI_MHW ✨`, selecciona `Procesar Lote con =GEMINI_MHW(...)`.
    *   Las fórmulas serán reemplazadas por las respuestas de Gemini.

## ⚠️ Limitaciones Actuales (MVP)

*   **Fórmulas Anidadas:** La función `=GEMINI_MHW(...)` **no soporta** ser anidada dentro de otras fórmulas de Google Sheets, ni tampoco anidar otras funciones dentro de sus argumentos de texto (el prompt y la información de apoyo deben ser texto directo o referencias directas a celdas que contienen texto). El parser de argumentos es básico.
*   **Manejo de Errores:** El manejo de errores es funcional pero podría ser más granular para diferentes tipos de fallos de la API.
*   **Límites de la API:** Lotes muy grandes podrían exceder los límites de tokens o tiempo de ejecución de la API de Gemini o Apps Script.

## ⚙️ Configuración del Modelo

Dentro de `code.gs`, puedes cambiar la constante `MODEL_NAME`:
```javascript
const MODEL_NAME = "gemini-2.0-flash"; // Modelo por defecto
