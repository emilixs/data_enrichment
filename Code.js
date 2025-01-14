// Constants
const PERPLEXITY_API_KEY = PropertiesService.getScriptProperties().getProperty('PERPLEXITY_API_KEY');
const PROCESSING_LIMIT = 5; // Initial limit for testing

// Menu creation
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Îmbogățire Date')
    .addItem('Procesează Companii', 'processCompanies')
    .addToUi();
}

// Main processing function
function processCompanies() {
  if (!validateStructure()) {
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = Math.min(6, sheet.getLastRow()); // Process only rows 2-6 (5 companies)
  let processedCount = 0;

  // Add status cell
  const statusCell = sheet.getRange("K1");
  statusCell.setValue("Status: În procesare...");

  for (let row = 2; row <= lastRow; row++) {
    if (!isRowProcessed(row)) {
      const companyName = sheet.getRange(row, 2).getValue(); // Column B
      if (companyName) {
        try {
          statusCell.setValue(`Status: Procesare ${companyName}...`);
          const response = callPerplexityAPI(companyName);
          
          // Check for API errors
          if (response.error) {
            throw new Error(`API Error: ${response.error.message || 'Unknown error'}`);
          }
          
          const data = parsePerplexityResponse(response);
          updateSheet(row, data);
          processedCount++;
          // Add small delay to avoid rate limiting
          Utilities.sleep(1000);
        } catch (error) {
          logError(error, row);
          
          // Handle rate limiting
          if (error.message.includes('rate limit')) {
            statusCell.setValue("Status: Rate limit atins. Încercați mai târziu.");
            ui.alert('Rate limit atins', 'Vă rugăm să încercați din nou în câteva minute.', ui.ButtonSet.OK);
            return;
          }
        }
      }
    }
  }
  
  statusCell.setValue(`Status: Procesare completă. ${processedCount} companii actualizate.`);
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

  if (!PERPLEXITY_API_KEY) {
    SpreadsheetApp.getUi().alert(
      'Eroare: API Key-ul Perplexity nu este configurat!'
    );
    return false;
  }

  return true;
}

// API interaction
function callPerplexityAPI(companyName) {
  const prompt = `Te rog caută și furnizează următoarele informații despre compania "${companyName}" SRL:

1. Numele oficial complet al companiei
2. Codul Unic de Înregistrare (CUI)
3. Cifra de afaceri pentru anul 2023 (sau cel mai recent an disponibil)
4. Profitul pentru anul 2023 (sau cel mai recent an disponibil)
5. Numărul de angajați
6. Website-ul oficial

Te rog să răspunzi strict cu informațiile găsite, în formatul:
[nume oficial]
[Cod fiscal]
[Cifra afaceri]
[Profit]
[Numar angajati]
[URL]`;

  const options = {
    'method': 'post',
    'headers': {
      'Authorization': `Bearer ${PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json',
    },
    'payload': JSON.stringify({
      'model': 'llama-3.1-sonar-large-128k-online',
      'messages': [{
        'role': 'user',
        'content': prompt
      }]
    })
  };

  const response = UrlFetchApp.fetch('https://api.perplexity.ai/chat/completions', options);
  return JSON.parse(response.getContentText());
}

// Response parsing
function parsePerplexityResponse(response) {
  const content = response.choices[0].message.content;
  
  // Initialize default values
  const data = {
    website: 'N/A',
    revenue: 'N/A',
    profit: 'N/A',
    employees: 'N/A',
    cui: 'N/A'  // Added CUI field
  };

  // Extract information using regex patterns
  const patterns = {
    website: /Site-ul:?\s*([^\n]+)/i,
    revenue: /Cifra de afaceri:?\s*([^\n]+)/i,
    profit: /Profit:?\s*([^\n]+)/i,
    employees: /Nr de angajati:?\s*([^\n]+)/i,
    cui: /Codul fiscal:?\s*([^\n]+)/i  // Added CUI pattern
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
  console.error(`Error processing row ${rowIndex}: ${error.message}`);
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(rowIndex, 6, 1, 5);
  range.setValue('Failed');
}
