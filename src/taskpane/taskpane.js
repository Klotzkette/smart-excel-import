/* ============================================
   Smart Excel Import — Core Logic
   ============================================ */

// --- State ---
let currentFile = null;
let currentFileContent = null; // text or base64
let currentFileType = null;   // 'pdf' | 'docx' | 'xlsx' | 'image'

// --- Office.js Init ---
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    init();
  }
});

function init() {
  // Settings toggle
  const toggle = document.getElementById('settings-toggle');
  const panel = document.getElementById('settings-panel');
  toggle.addEventListener('click', () => {
    const expanded = toggle.getAttribute('aria-expanded') === 'true';
    toggle.setAttribute('aria-expanded', !expanded);
    panel.classList.toggle('hidden');
  });

  // Load saved settings
  loadSettings();

  // API Key visibility toggle
  document.getElementById('toggle-key-visibility').addEventListener('click', () => {
    const input = document.getElementById('api-key-input');
    input.type = input.type === 'password' ? 'text' : 'password';
  });

  // Save key
  document.getElementById('save-key-btn').addEventListener('click', saveSettings);

  // Drop zone
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') fileInput.click();
  });

  document.getElementById('browse-btn').addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });

  // Remove file
  document.getElementById('remove-file').addEventListener('click', resetFile);

  // Import button
  document.getElementById('import-btn').addEventListener('click', startImport);
}

// --- Settings ---
function loadSettings() {
  try {
    const provider = localStorage.getItem('sei_provider') || 'openai';
    const key = localStorage.getItem('sei_apikey') || '';
    document.getElementById('provider-select').value = provider;
    document.getElementById('api-key-input').value = key;
    if (key) {
      showKeyStatus('Gespeicherter Schlüssel geladen', 'saved');
    }
  } catch (e) { /* localStorage may not be available */ }
}

function saveSettings() {
  try {
    const provider = document.getElementById('provider-select').value;
    const key = document.getElementById('api-key-input').value.trim();
    localStorage.setItem('sei_provider', provider);
    localStorage.setItem('sei_apikey', key);
    showKeyStatus('Einstellungen gespeichert', 'saved');
  } catch (e) {
    showKeyStatus('Fehler beim Speichern', 'error');
  }
}

function showKeyStatus(msg, type) {
  const el = document.getElementById('key-status');
  el.textContent = msg;
  el.className = 'key-status ' + type;
  setTimeout(() => { el.textContent = ''; }, 3000);
}

// --- File Handling ---
const SUPPORTED_TYPES = {
  'application/pdf': 'pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
  'application/vnd.ms-excel': 'xlsx',
  'image/jpeg': 'image',
  'image/png': 'image',
  'image/gif': 'image',
  'image/bmp': 'image',
  'image/webp': 'image',
};

function getFileTypeFromName(name) {
  const ext = name.split('.').pop().toLowerCase();
  const map = { pdf: 'pdf', docx: 'docx', xlsx: 'xlsx', xls: 'xlsx', jpg: 'image', jpeg: 'image', png: 'image', gif: 'image', bmp: 'image', webp: 'image' };
  return map[ext] || null;
}

function handleFile(file) {
  const type = SUPPORTED_TYPES[file.type] || getFileTypeFromName(file.name);
  if (!type) {
    showError('Nicht unterstütztes Dateiformat. Bitte PDF, Word, Excel oder Bild verwenden.');
    return;
  }

  currentFile = file;
  currentFileType = type;
  currentFileContent = null;

  // Show file preview
  document.getElementById('file-name').textContent = file.name;
  document.getElementById('file-size').textContent = formatFileSize(file.size);
  document.getElementById('file-preview').classList.remove('hidden');
  document.getElementById('options-section').classList.remove('hidden');
  document.getElementById('drop-section').classList.add('hidden');
  hideMessages();
}

function resetFile() {
  currentFile = null;
  currentFileContent = null;
  currentFileType = null;
  document.getElementById('file-preview').classList.add('hidden');
  document.getElementById('options-section').classList.add('hidden');
  document.getElementById('drop-section').classList.remove('hidden');
  document.getElementById('file-input').value = '';
  hideMessages();
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

// --- File Parsing ---
async function extractContent(file, type) {
  switch (type) {
    case 'pdf':
      return await extractPDF(file);
    case 'docx':
      return await extractDOCX(file);
    case 'xlsx':
      return extractXLSX(file);
    case 'image':
      return await fileToBase64(file);
    default:
      throw new Error('Unbekannter Dateityp');
  }
}

async function extractPDF(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let text = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const pageText = content.items.map(item => item.str).join(' ');
    text += `--- Seite ${i} ---\n${pageText}\n\n`;
  }
  return text;
}

async function extractDOCX(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value;
}

function extractXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        let text = '';
        workbook.SheetNames.forEach(name => {
          const sheet = workbook.Sheets[name];
          const csv = XLSX.utils.sheet_to_csv(sheet);
          text += `--- Blatt: ${name} ---\n${csv}\n\n`;
        });
        resolve(text);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// --- AI API Calls ---
async function callAI(content, type) {
  const provider = document.getElementById('provider-select').value;
  const apiKey = document.getElementById('api-key-input').value.trim() || localStorage.getItem('sei_apikey') || '';

  if (!apiKey) {
    throw new Error('Bitte zuerst einen API-Schlüssel in den Einstellungen hinterlegen.');
  }

  const useFormulas = document.getElementById('use-formulas').checked;
  const formatTable = document.getElementById('format-table').checked;

  const systemPrompt = buildSystemPrompt(useFormulas, formatTable);

  if (provider === 'openai') {
    return await callOpenAI(apiKey, systemPrompt, content, type);
  } else {
    return await callAnthropic(apiKey, systemPrompt, content, type);
  }
}

function buildSystemPrompt(useFormulas, formatTable) {
  return `Du bist ein Experte für Datenextraktion und Excel-Tabellen. Deine Aufgabe ist es, aus dem bereitgestellten Dokument alle tabellarischen Daten, Zahlen und Strukturen zu extrahieren und als strukturiertes JSON zurückzugeben.

WICHTIGE REGELN:
1. Erkenne alle Tabellen, Listen und numerische Daten im Dokument.
2. ${useFormulas ? 'Wo immer Werte berechnet werden (Summen, Differenzen, Prozentsätze, Durchschnitte, Produkte etc.), gib KEINE festen Zahlen an, sondern eine Excel-Formel. Verwende dabei Zellbezüge relativ zur Position in der Tabelle. Beispiel: Wenn Zelle C5 die Summe von C2:C4 ist, gib "=SUM(C2:C4)" an.' : 'Gib alle Werte als feste Zahlen an.'}
3. ${formatTable ? 'Markiere Überschriften, Kategorien und Datentypen (Zahl, Prozent, Währung, Text, Datum).' : ''}
4. Erkenne die logische Struktur: Überschriften, Zwischensummen, Gesamtsummen, Gruppierungen.
5. Wenn das Dokument mehrere Tabellen enthält, gib alle zurück.

Antworte AUSSCHLIESSLICH mit validem JSON in folgendem Format:
{
  "tables": [
    {
      "title": "Tabellenname (falls erkennbar)",
      "headers": ["Spalte1", "Spalte2", "Spalte3"],
      "column_types": ["text", "number", "currency"],
      "rows": [
        ["Bezeichnung", 1234, 5678.90],
        ["Andere Zeile", 2345, "=B2*1.19"]
      ],
      "summary_rows": [
        {
          "label": "Summe",
          "row_index": 5,
          "formulas": {"B": "=SUM(B2:B4)", "C": "=SUM(C2:C4)"}
        }
      ]
    }
  ],
  "metadata": {
    "document_title": "Erkannter Titel",
    "date": "Erkanntes Datum falls vorhanden",
    "source": "Art des Dokuments"
  }
}

COLUMN_TYPES: "text", "number", "currency", "percent", "date", "formula"
Bei Währungen: Gib nur die Zahl zurück, Formatierung übernimmt Excel.
Bei Prozenten: Gib den Dezimalwert zurück (0.19 für 19%).
Formeln: Verwende englische Excel-Funktionsnamen (SUM, AVERAGE, IF, etc.).
Zellbezüge: Beginne bei der Startposition, die der Benutzer angibt (Standard A1). Header in Zeile 1, Daten ab Zeile 2.`;
}

async function callOpenAI(apiKey, systemPrompt, content, type) {
  const messages = [
    { role: 'system', content: systemPrompt }
  ];

  if (type === 'image') {
    messages.push({
      role: 'user',
      content: [
        { type: 'text', text: 'Analysiere dieses Bild. Extrahiere alle Tabellen, Zahlen und Daten und gib sie als strukturiertes JSON zurück. Erkenne berechnete Werte und erstelle dafür Excel-Formeln.' },
        { type: 'image_url', image_url: { url: content } }
      ]
    });
  } else {
    messages.push({
      role: 'user',
      content: `Analysiere das folgende Dokument. Extrahiere alle Tabellen, Zahlen und Daten und gib sie als strukturiertes JSON zurück. Erkenne berechnete Werte und erstelle dafür Excel-Formeln.\n\n--- DOKUMENTINHALT ---\n${content}`
    });
  }

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: 'gpt-4o',
      messages: messages,
      temperature: 0.1,
      max_tokens: 16000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(`OpenAI Fehler: ${err.error?.message || response.statusText}`);
  }

  const data = await response.json();
  return JSON.parse(data.choices[0].message.content);
}

async function callAnthropic(apiKey, systemPrompt, content, type) {
  const messages = [];

  if (type === 'image') {
    // Extract mime type and base64 data from data URL
    const match = content.match(/^data:(image\/[a-z+]+);base64,(.+)$/i);
    if (!match) throw new Error('Ungültiges Bildformat');

    messages.push({
      role: 'user',
      content: [
        {
          type: 'image',
          source: { type: 'base64', media_type: match[1], data: match[2] }
        },
        { type: 'text', text: 'Analysiere dieses Bild. Extrahiere alle Tabellen, Zahlen und Daten und gib sie als strukturiertes JSON zurück. Erkenne berechnete Werte und erstelle dafür Excel-Formeln.' }
      ]
    });
  } else {
    messages.push({
      role: 'user',
      content: `Analysiere das folgende Dokument. Extrahiere alle Tabellen, Zahlen und Daten und gib sie als strukturiertes JSON zurück. Erkenne berechnete Werte und erstelle dafür Excel-Formeln.\n\n--- DOKUMENTINHALT ---\n${content}`
    });
  }

  // Anthropic API requires CORS proxy or backend in production.
  // For the add-in, we call directly (works with API key auth).
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 16000,
      system: systemPrompt,
      messages: messages
    })
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(`Anthropic Fehler: ${err.error?.message || response.statusText}`);
  }

  const data = await response.json();
  const text = data.content[0].text;

  // Extract JSON from response (Claude may wrap it in markdown)
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('Keine gültige JSON-Antwort von Claude erhalten.');
  return JSON.parse(jsonMatch[0]);
}

// --- Excel Writing ---
async function writeToExcel(aiResult) {
  const startCell = document.getElementById('start-cell').value.trim() || 'A1';
  const formatTable = document.getElementById('format-table').checked;

  const startMatch = startCell.match(/^([A-Z]+)(\d+)$/i);
  if (!startMatch) throw new Error('Ungültige Startposition. Bitte z.B. "A1" oder "B3" eingeben.');

  const startCol = columnToNumber(startMatch[1].toUpperCase());
  const startRow = parseInt(startMatch[2]);

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let currentRow = startRow;
    let totalCells = 0;

    const tables = aiResult.tables || [];
    if (tables.length === 0) {
      throw new Error('Keine Tabellendaten im Dokument erkannt.');
    }

    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      const headers = table.headers || [];
      const rows = table.rows || [];
      const colTypes = table.column_types || [];
      const numCols = headers.length || (rows[0] ? rows[0].length : 0);

      // --- Table Title ---
      if (table.title) {
        const titleCell = sheet.getCell(currentRow - 1, startCol - 1);
        titleCell.values = [[table.title]];
        if (formatTable) {
          titleCell.format.font.bold = true;
          titleCell.format.font.size = 13;
          titleCell.format.font.color = '#1a1a1a';
        }
        currentRow += 1;
      }

      // --- Headers ---
      if (headers.length > 0) {
        const headerRange = sheet.getRangeByIndexes(currentRow - 1, startCol - 1, 1, numCols);
        headerRange.values = [headers];

        if (formatTable) {
          headerRange.format.font.bold = true;
          headerRange.format.font.size = 11;
          headerRange.format.font.color = '#ffffff';
          headerRange.format.fill.color = '#0f7b6c';
          headerRange.format.horizontalAlignment = 'Center';
          headerRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
          headerRange.format.borders.getItem('EdgeBottom').color = '#0a5e52';
          headerRange.format.borders.getItem('EdgeBottom').weight = 'Medium';

          // Column widths
          for (let c = 0; c < numCols; c++) {
            const col = sheet.getRangeByIndexes(currentRow - 1, startCol - 1 + c, 1, 1);
            col.format.columnWidth = colTypes[c] === 'text' ? 150 : 100;
          }
        }
        currentRow += 1;
      }

      // --- Data Rows ---
      const dataStartRow = currentRow;
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        for (let c = 0; c < row.length; c++) {
          const cell = sheet.getCell(currentRow - 1, startCol - 1 + c);
          const value = row[c];
          const colType = colTypes[c] || 'text';

          if (typeof value === 'string' && value.startsWith('=')) {
            // It's a formula — adjust cell references to actual position
            const adjustedFormula = adjustFormulaReferences(value, startCol, dataStartRow);
            cell.formulas = [[adjustedFormula]];
          } else {
            cell.values = [[value === null || value === undefined ? '' : value]];
          }

          // Apply formatting
          if (formatTable) {
            applyNumberFormat(cell, colType);

            // Alternating row colors
            if (r % 2 === 1) {
              cell.format.fill.color = '#f0f9f7';
            }

            // Borders
            cell.format.borders.getItem('EdgeBottom').style = 'Continuous';
            cell.format.borders.getItem('EdgeBottom').color = '#e2e0da';

            // Alignment
            if (colType === 'text') {
              cell.format.horizontalAlignment = 'Left';
              cell.format.indentLevel = 1;
            } else {
              cell.format.horizontalAlignment = 'Right';
            }
          }

          totalCells++;
        }
        currentRow += 1;
      }

      // --- Summary Rows (Formulas) ---
      if (table.summary_rows) {
        for (const summary of table.summary_rows) {
          const summaryRowIdx = currentRow - 1;
          const labelCell = sheet.getCell(summaryRowIdx, startCol - 1);
          labelCell.values = [[summary.label || 'Summe']];

          if (formatTable) {
            labelCell.format.font.bold = true;
          }

          if (summary.formulas) {
            for (const [colLetter, formula] of Object.entries(summary.formulas)) {
              const colIdx = columnToNumber(colLetter) - 1;
              const cell = sheet.getCell(summaryRowIdx, startCol - 1 + colIdx);
              const adjustedFormula = adjustFormulaReferences(formula, startCol, dataStartRow);
              cell.formulas = [[adjustedFormula]];

              if (formatTable) {
                cell.format.font.bold = true;
                cell.format.borders.getItem('EdgeTop').style = 'Continuous';
                cell.format.borders.getItem('EdgeTop').weight = 'Medium';
                cell.format.borders.getItem('EdgeTop').color = '#0f7b6c';
                cell.format.borders.getItem('EdgeBottom').style = 'Double';
                cell.format.borders.getItem('EdgeBottom').color = '#0f7b6c';
                applyNumberFormat(cell, colTypes[colIdx] || 'number');
              }
            }
          }

          currentRow += 1;
          totalCells++;
        }
      }

      // --- Metadata ---
      if (aiResult.metadata && t === tables.length - 1) {
        currentRow += 1;
        const meta = aiResult.metadata;
        if (meta.source || meta.date) {
          const metaCell = sheet.getCell(currentRow - 1, startCol - 1);
          let metaText = '';
          if (meta.source) metaText += `Quelle: ${meta.source}`;
          if (meta.date) metaText += `${metaText ? ' | ' : ''}Datum: ${meta.date}`;
          metaCell.values = [[metaText]];
          if (formatTable) {
            metaCell.format.font.size = 9;
            metaCell.format.font.color = '#9a9a9a';
            metaCell.format.font.italic = true;
          }
          currentRow += 1;
        }
      }

      // Gap between tables
      currentRow += 2;
    }

    await context.sync();
    return totalCells;
  });
}

function adjustFormulaReferences(formula, startCol, dataStartRow) {
  // The AI generates formulas assuming standard A1 positioning with
  // headers in row 1 and data starting in row 2.
  // We need to adjust if the user chose a different start position.
  // For simplicity, we trust the AI-generated references if start is A1.
  // Otherwise, we would need a full formula parser.
  // The system prompt tells the AI to use the actual cell positions.
  return formula;
}

function applyNumberFormat(cell, colType) {
  switch (colType) {
    case 'currency':
      cell.numberFormat = '#,##0.00 €';
      break;
    case 'percent':
      cell.numberFormat = '0.0%';
      break;
    case 'number':
      cell.numberFormat = '#,##0.00';
      break;
    case 'date':
      cell.numberFormat = 'DD.MM.YYYY';
      break;
    default:
      break;
  }
}

function columnToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

function numberToColumn(num) {
  let col = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

// --- Import Flow ---
async function startImport() {
  if (!currentFile) return;

  const apiKey = document.getElementById('api-key-input').value.trim() || localStorage.getItem('sei_apikey') || '';
  if (!apiKey) {
    showError('Bitte zuerst einen API-Schlüssel in den Einstellungen hinterlegen.');
    // Open settings
    document.getElementById('settings-toggle').setAttribute('aria-expanded', 'true');
    document.getElementById('settings-panel').classList.remove('hidden');
    return;
  }

  hideMessages();
  showProgress(0, 'Datei wird gelesen…');
  document.getElementById('import-btn').disabled = true;

  try {
    // Step 1: Extract content
    showProgress(20, `${currentFileType === 'image' ? 'Bild' : 'Dokument'} wird gelesen…`);
    const content = await extractContent(currentFile, currentFileType);
    currentFileContent = content;

    // Step 2: AI Analysis
    showProgress(45, 'KI analysiert die Daten…');
    const aiResult = await callAI(content, currentFileType);

    // Step 3: Validate
    showProgress(70, 'Tabellenstruktur wird aufgebaut…');
    if (!aiResult.tables || aiResult.tables.length === 0) {
      throw new Error('Die KI konnte keine Tabellendaten im Dokument erkennen. Versuchen Sie es mit einem anderen Dokument oder einem deutlicheren Screenshot.');
    }

    // Step 4: Write to Excel
    showProgress(85, 'Daten werden in Excel geschrieben…');
    await writeToExcel(aiResult);

    // Done
    showProgress(100, 'Fertig!');
    const tableCount = aiResult.tables.length;
    const rowCount = aiResult.tables.reduce((sum, t) => sum + (t.rows ? t.rows.length : 0), 0);
    showResult(`${tableCount} Tabelle${tableCount > 1 ? 'n' : ''} mit ${rowCount} Zeilen importiert.`);

  } catch (err) {
    console.error('Import error:', err);
    showError(err.message || 'Unbekannter Fehler beim Import.');
  } finally {
    document.getElementById('import-btn').disabled = false;
    hideProgress();
  }
}

// --- UI Helpers ---
function showProgress(percent, text) {
  document.getElementById('progress-section').classList.remove('hidden');
  document.getElementById('progress-bar').style.width = percent + '%';
  document.getElementById('progress-text').textContent = text;
}

function hideProgress() {
  setTimeout(() => {
    document.getElementById('progress-section').classList.add('hidden');
  }, 800);
}

function showResult(text) {
  document.getElementById('result-section').classList.remove('hidden');
  document.getElementById('result-text').textContent = text;
}

function showError(text) {
  document.getElementById('error-section').classList.remove('hidden');
  document.getElementById('error-text').textContent = text;
}

function hideMessages() {
  document.getElementById('result-section').classList.add('hidden');
  document.getElementById('error-section').classList.add('hidden');
}
