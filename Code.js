// Constants
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const LOG_SHEET_NAME = 'Logs';

// Logging function
function logToSheet(message, level = 'INFO', details = '') {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let activeSheet = ss.getActiveSheet(); // Store the active sheet
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    
    // Create log sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      logSheet.getRange('A1:D1').setValues([['Timestamp', 'Level', 'Message', 'Details']]);
      logSheet.setFrozenRows(1);
      ss.setActiveSheet(activeSheet); // Switch back to the active sheet
    }
    
    const timestamp = new Date().toISOString();
    logSheet.appendRow([timestamp, level, message, details]);
    
    // Keep only last 1000 logs
    const maxRows = 1000;
    const currentRows = logSheet.getLastRow();
    if (currentRows > maxRows) {
      logSheet.deleteRows(2, currentRows - maxRows);
    }
    
    ss.setActiveSheet(activeSheet); // Ensure we're back on the active sheet
  } catch (error) {
    console.error('Logging failed:', error);
  }
}

// Menu creation
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Îmbogățire Date')
    .addItem('Procesează Companii', 'processCompanies')
    .addToUi();
}

// Main processing function
function processCompanies() {
  logToSheet('Starting company processing');
  
  if (!validateStructure()) {
    logToSheet('Validation failed', 'ERROR');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let processedCount = 0;

  logToSheet(`Processing rows 2 to ${lastRow}`, 'INFO');

  // Add status cell
  const statusCell = sheet.getRange("K1");
  statusCell.setValue("Status: În procesare...");

  for (let row = 2; row <= lastRow; row++) {
    if (!isRowProcessed(row)) {
      const companyName = sheet.getRange(row, 2).getValue();
      if (companyName) {
        try {
          logToSheet(`Processing company: ${companyName}`, 'INFO', `Row: ${row}`);
          statusCell.setValue(`Status: Procesare ${companyName}... (${row-1}/${lastRow-1})`);
          
          const response = callPerplexityAPI(companyName);
          logToSheet('API response received', 'DEBUG', JSON.stringify(response));
          
          if (response.error) {
            throw new Error(`API Error: ${response.error.message || 'Unknown error'}`);
          }
          
          const data = parsePerplexityResponse(response);
          logToSheet('Parsed response', 'DEBUG', JSON.stringify(data));
          
          updateSheet(row, data);
          processedCount++;
          logToSheet(`Successfully processed ${companyName}`, 'INFO', `Row: ${row}, Data: ${JSON.stringify(data)}`);
          
          // Add longer delay between requests to avoid rate limits
          Utilities.sleep(2000);
        } catch (error) {
          logToSheet(`Error processing ${companyName}`, 'ERROR', `Row: ${row}, Error: ${error.message}`);
          logError(error, row);
          
          if (error.message.includes('rate limit') || error.message.includes('RESOURCE_EXHAUSTED')) {
            const waitMinutes = 2;
            const waitMessage = `Rate limit atins. Așteptăm ${waitMinutes} minute înainte de a continua...`;
            statusCell.setValue(`Status: ${waitMessage}`);
            logToSheet('Rate limit reached, waiting...', 'WARNING');
            
            // Wait for 2 minutes before continuing
            Utilities.sleep(waitMinutes * 60 * 1000);
            
            // Retry the same row
            row--;
            continue;
          }
        }
      }
    } else {
      logToSheet(`Skipping processed row ${row}`, 'INFO');
    }
  }
  
  const finalMessage = `Status: Procesare completă. ${processedCount} companii actualizate.`;
  statusCell.setValue(finalMessage);
  logToSheet('Processing completed', 'INFO', `Processed count: ${processedCount}`);
}

// Validation function
function validateStructure() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange("B1:J1").getValues()[0];
  
  const requiredHeaders = [
    "Companie",
    "Website",
    "Cifra afaceri (2023)",
    "Profit",
    "Nr. angajati",
    "CUI"
  ];

  const missingHeaders = requiredHeaders.filter((header, index) => 
    !headers.some(h => h.toString().toLowerCase().includes(header.toLowerCase()))
  );

  if (missingHeaders.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Eroare de structură: Lipsesc următoarele coloane: ' + missingHeaders.join(', ')
    );
    return false;
  }

  if (!GEMINI_API_KEY) {
    SpreadsheetApp.getUi().alert(
      'Eroare: API Key-ul Google Gemini nu este configurat!'
    );
    return false;
  }

  return true;
}

// API interaction
function callPerplexityAPI(companyName) {
  const maxRetries = 3;
  const retryDelay = 2000; // 2 seconds
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const prompt = `Te rog caută și furnizează următoarele informații despre compania "${companyName}", in paranteza gasesti domeniul de activitate. :
      Te rog sa alegi cea mai probabil companie romaneasca si sa extragi detaliile pentru ea. 
      Important:Forteaza ca raspunsul sa fie doar ce este pus in model intre [].
      Am nevoie de ele mai departe intr-un script. Te rog mult. 

1. Numele oficial complet al companiei
2. Codul Unic de Înregistrare (CUI)
3. Cifra de afaceri pentru anul 2023 (sau cel mai recent an disponibil)
4. Profitul pentru anul 2023 (sau cel mai recent an disponibil)
5. Numărul de angajați
6. Website-ul oficial

Caută informațiile pe listafirme.ro și alte surse oficiale românești.
Răspunde strict cu informațiile găsite, în formatul:
Numele oficial: [nume]
Codul fiscal: [CUI]
Cifra de afaceri: [suma]
Profit: [suma]
Nr de angajati: [număr]
Site-ul: [URL]`;

      const options = {
        'method': 'post',
        'headers': {
          'x-goog-api-key': GEMINI_API_KEY,
          'Content-Type': 'application/json',
        },
        'payload': JSON.stringify({
          'contents': [{
            'parts': [{
              'text': prompt
            }]
          }],
          'model': 'gemini-1.5-flash',
          'generationConfig': {
            'temperature': 0,
            'topK': 1,
            'topP': 1
          },
          'tools': {
            'google_search_retrieval': {}
          }
        }),
        'muteHttpExceptions': true
      };

      const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent', options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      const responseData = JSON.parse(responseText);

      // Check for rate limit or quota exceeded
      if (responseCode === 429) {
        logToSheet(`Rate limit hit on attempt ${attempt}`, 'WARNING');
        if (attempt === maxRetries) {
          throw new Error('Rate limit exceeded after all retries');
        }
        // Wait longer with each retry
        Utilities.sleep(retryDelay * attempt);
        continue;
      }

      // Check for other errors
      if (responseCode !== 200) {
        throw new Error(`API returned code ${responseCode}: ${responseText}`);
      }

      return responseData;
    } catch (error) {
      if (attempt === maxRetries) {
        throw error;
      }
      logToSheet(`Attempt ${attempt} failed: ${error.message}`, 'WARNING');
      Utilities.sleep(retryDelay * attempt);
    }
  }
}

// Response parsing
function parsePerplexityResponse(response) {
  const content = response.candidates[0].content.parts[0].text;
  
  // Initialize default values
  const data = {
    website: 'N/A',
    revenue: 'N/A',
    profit: 'N/A',
    employees: 'N/A',
    cui: 'N/A'
  };

  // Extract information using regex patterns
  const patterns = {
    website: /Site-ul:?\s*([^\n]+)/i,
    revenue: /Cifra de afaceri:?\s*([^\n]+)/i,
    profit: /Profit:?\s*([^\n]+)/i,
    employees: /Nr de angajati:?\s*([^\n]+)/i,
    cui: /Codul fiscal:?\s*([^\n]+)/i
  };

  // Update data object with found values
  for (const [key, pattern] of Object.entries(patterns)) {
    const match = content.match(pattern);
    if (match && match[1].trim()) {
      data[key] = match[1].trim();
    }
  }

  return data;
}

// Sheet update function
function updateSheet(rowIndex, data) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Update cells with the extracted data in the correct order
  sheet.getRange(rowIndex, 6).setValue(data.cui);         // Column F - CUI
  sheet.getRange(rowIndex, 7).setValue(data.website);     // Column G - Website
  sheet.getRange(rowIndex, 8).setValue(data.revenue);     // Column H - Revenue
  sheet.getRange(rowIndex, 9).setValue(data.profit);      // Column I - Profit
  sheet.getRange(rowIndex, 10).setValue(data.employees);  // Column J - Employees
}

// Row processing check
function isRowProcessed(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getRange(rowIndex, 6, 1, 5).getValues()[0];
  return row.some(cell => cell !== '');
}

// Error logging
function logError(error, rowIndex) {
  const errorMessage = `Error processing row ${rowIndex}: ${error.message}`;
  console.error(errorMessage);
  logToSheet(errorMessage, 'ERROR');
  
  const sheet = SpreadsheetApp.getActiveSheet();
  const errorDetails = error.message.includes('code') ? error.message : 'Error: ' + error.message;
  
  // Fill all cells in the row with the error message
  const range = sheet.getRange(rowIndex, 6, 1, 5);
  const values = Array(5).fill([errorDetails]);
  range.setValues([Array(5).fill(errorDetails)]);
  
  // Make error text red
  range.setFontColor('red');
}
