# GEMINI_MHW — Gemini como fórmula en Google Sheets (MVP)

Script de Google Apps Script que permite usar Gemini como si fuera una fórmula nativa
dentro de Google Sheets: `=GEMINI_MHW("tu prompt", celda)`. Todas las fórmulas de la
hoja activa se resuelven en una sola llamada a la API.

Demo: https://www.youtube.com/watch?v=UdV_kD3hUWU

## Características (MVP)

- Gemini como fórmula: `=GEMINI_MHW("tu prompt", [celda_o_texto])`.
- Procesamiento por lotes: todas las fórmulas de la hoja en una sola llamada a la API.
- Menú integrado para configurar la API Key y ejecutar el procesamiento.
- Modelo configurable en la constante `MODEL_NAME`.
- Diálogo de ayuda accesible desde el menú.

## Requisitos

1. Una cuenta de Google.
2. Una API Key de Gemini ([Google AI Studio](https://makersuite.google.com/app/apikey)).

## Instalación

1. Abre tu Google Sheet → `Extensiones > Apps Script`.
2. Vacía `Código.gs` y pega el contenido de [`code.gs`](./code.gs).
3. Crea un archivo HTML llamado `HelpDialog.html` y pega el contenido de
   [`HelpDialog.html`](./HelpDialog.html).
4. Guarda y recarga la hoja. Aparecerá el menú `GEMINI_MHW`.
5. En `GEMINI_MHW > Configurar API Key`, pega tu key en la celda A1 de la hoja `config`.

## Uso

1. Escribe el prompt como fórmula:
   `=GEMINI_MHW("Tu instrucción", A1)` o `=GEMINI_MHW("Analiza:", "texto a analizar")`.
2. En el menú `GEMINI_MHW`, elige "Procesar lote". Las fórmulas se reemplazan por las
   respuestas de Gemini.

## Limitaciones (MVP)

- No soporta anidar `=GEMINI_MHW(...)` dentro de otras fórmulas, ni funciones dentro de
  sus argumentos: el prompt y el apoyo deben ser texto directo o referencias a celdas.
  El parser de argumentos es básico.
- El manejo de errores es funcional pero poco granular.
- Lotes muy grandes pueden exceder los límites de tokens o de tiempo de ejecución de
  Gemini o Apps Script.

## Modelo

En `code.gs`, cambia la constante `MODEL_NAME`:

```javascript
const MODEL_NAME = "gemini-2.0-flash";
```
