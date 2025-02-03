// Constants
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const LOG_SHEET_NAME = 'Logs';
const JOB_DESCRIPTION_SHEET = 'Job Description';
const OUTPUT_COLUMNS = {
  TECHNICAL_SCORE: 'K',
  EXPERIENCE_SCORE: 'L',
  OVERALL_SCORE: 'M',
  RECOMMENDATIONS: 'N',
  STATUS: 'O'
};

/**
 * logToSheet
 *
 * Logs a message to a dedicated "Logs" sheet in the active spreadsheet.
 * If the "Logs" sheet doesn't exist, it will be created along with headers.
 * The function appends a new row with the current timestamp, log level, message, and extra details.
 * It also keeps the number of rows within a defined limit (e.g., 1000 rows) to avoid uncontrolled growth.
 *
 * @param {string} message - The primary log message.
 * @param {string} [level="INFO"] - The severity level (e.g., INFO, DEBUG, WARNING, ERROR).
 * @param {string} [details=""] - Additional details and context for the log.
 */
function logToSheet(message, level = 'INFO', details = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    let logSheet = ss.getSheetByName('Logs');

    // Create the log sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet('Logs');
      logSheet.getRange('A1:D1').setValues([['Timestamp', 'Level', 'Message', 'Details']]);
      logSheet.setFrozenRows(1);
      ss.setActiveSheet(activeSheet); // Revert back to the originally active sheet
    }

    const timestamp = new Date().toISOString();
    logSheet.appendRow([timestamp, level, message, details]);

    // Keep only the last 1000 log entries
    const maxRows = 1000;
    const currentRows = logSheet.getLastRow();
    if (currentRows > maxRows) {
      logSheet.deleteRows(2, currentRows - maxRows);
    }

    ss.setActiveSheet(activeSheet); // Ensure we switch back to the active sheet
  } catch (error) {
    console.error('Logging failed:', error);
  }
}

// Menu creation
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Analiză Profile')
    .addItem('Configurare Job Description', 'configureJobDescription')
    .addItem('Procesează Profile', 'processProfiles')
    .addItem('Resetare Evaluări', 'resetEvaluations')
    .addToUi();
}

// Job Description Configuration
function configureJobDescription() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let jobDescSheet = ss.getSheetByName(JOB_DESCRIPTION_SHEET);
  
  if (!jobDescSheet) {
    jobDescSheet = ss.insertSheet(JOB_DESCRIPTION_SHEET);
    jobDescSheet.getRange('A1').setValue('Job Description');
  }
  
  const response = ui.prompt(
    'Configurare Job Description',
    'Introduceți descrierea job-ului pentru evaluarea profilelor:',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    jobDescSheet.getRange('A2').setValue(response.getResponseText());
    ui.alert('Job Description salvat cu succes!');
  }
}

// Reset Evaluations
function resetEvaluations() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Resetare Evaluări',
    'Această acțiune va șterge toate evaluările existente. Doriți să continuați?',
    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, OUTPUT_COLUMNS.TECHNICAL_SCORE.charCodeAt(0) - 64, lastRow - 1, 5).clearContent();
      ui.alert('Evaluările au fost resetate cu succes!');
    }
  }
}

// Main processing function
function processProfiles() {
  logToSheet('Starting profile processing');
  
  if (!validateStructure()) {
    logToSheet('Validation failed', 'ERROR');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let processedCount = 0;
  const PROFILE_LIMIT = 5;

  // Get Job Description
  const jobDesc = getJobDescription();
  if (!jobDesc) {
    ui.alert('Error', 'Job Description nu este configurat!', ui.ButtonSet.OK);
    return;
  }

  logToSheet(`Looking for first ${PROFILE_LIMIT} unprocessed profiles`, 'INFO');

  // Add status cell
  const statusCell = sheet.getRange(1, OUTPUT_COLUMNS.STATUS.charCodeAt(0) - 64);
  statusCell.setValue("Status: În procesare...");

  for (let row = 2; row <= lastRow && processedCount < PROFILE_LIMIT; row++) {
    if (!isProfileProcessed(row)) {
      const profileData = extractProfileData(row);
      if (validateProfileData(profileData)) {
        try {
          logToSheet(`Processing profile ${processedCount + 1}/${PROFILE_LIMIT}: ${profileData.firstName} ${profileData.lastName}`, 'INFO', `Row: ${row}`);
          statusCell.setValue(`Status: Procesare ${profileData.firstName} ${profileData.lastName}... (${processedCount + 1}/${PROFILE_LIMIT})`);
          
          const formattedData = formatProfileData(profileData);
          const response = callGeminiAPI(formattedData, jobDesc);
          logToSheet('API response received', 'DEBUG', JSON.stringify(response));
          
          const evaluationData = parseGeminiResponse(response);
          updateSheet(row, evaluationData);
          processedCount++;
          
          // Add delay between requests
          Utilities.sleep(2000);
        } catch (error) {
          logToSheet(`Error processing profile`, 'ERROR', `Row: ${row}, Error: ${error.message}`);
          logError(error, row);
          
          if (error.message.includes('RESOURCE_EXHAUSTED')) {
            const waitMinutes = 2;
            statusCell.setValue(`Status: Rate limit atins. Așteptăm ${waitMinutes} minute...`);
            Utilities.sleep(waitMinutes * 60 * 1000);
            row--; // Retry the same row
            continue;
          }
        }
      }
    } else {
      logToSheet(`Skipping processed profile at row ${row}`, 'INFO');
    }
  }
  
  const finalMessage = processedCount === 0 
    ? 'Status: Nu s-au găsit profile neprocesate.'
    : `Status: Procesare completă. ${processedCount} profile actualizate.`;
  
  statusCell.setValue(finalMessage);
  logToSheet('Processing completed', 'INFO', `Processed count: ${processedCount}`);
}

// Profile data extraction
function extractProfileData(row) {
  const sheet = SpreadsheetApp.getActiveSheet();
  return {
    companyIndustry: sheet.getRange(row, getColumnByName('companyIndustry')).getValue(),
    companyName: sheet.getRange(row, getColumnByName('companyName')).getValue(),
    linkedinHeadline: sheet.getRange(row, getColumnByName('linkedinHeadline')).getValue(),
    linkedinJobDateRange: sheet.getRange(row, getColumnByName('linkedinJobDateRange')).getValue(),
    linkedinJobTitle: sheet.getRange(row, getColumnByName('linkedinJobTitle')).getValue(),
    linkedinPreviousJobDateRange: sheet.getRange(row, getColumnByName('linkedinPreviousJobDateRange')).getValue(),
    linkedinPreviousJobTitle: sheet.getRange(row, getColumnByName('linkedinPreviousJobTitle')).getValue(),
    linkedinSkillsLabel: sheet.getRange(row, getColumnByName('linkedinSkillsLabel')).getValue(),
    location: sheet.getRange(row, getColumnByName('location')).getValue(),
    previousCompanyName: sheet.getRange(row, getColumnByName('previousCompanyName')).getValue(),
    linkedinSchoolDegree: sheet.getRange(row, getColumnByName('linkedinSchoolDegree')).getValue(),
    linkedinSchoolName: sheet.getRange(row, getColumnByName('linkedinSchoolName')).getValue(),
    linkedinPreviousSchoolDateRange: sheet.getRange(row, getColumnByName('linkedinPreviousSchoolDateRange')).getValue(),
    linkedinPreviousSchoolDegree: sheet.getRange(row, getColumnByName('linkedinPreviousSchoolDegree')).getValue(),
    linkedinPreviousSchoolName: sheet.getRange(row, getColumnByName('linkedinPreviousSchoolName')).getValue(),
    linkedinSchoolDateRange: sheet.getRange(row, getColumnByName('linkedinSchoolDateRange')).getValue(),
    linkedinDescription: sheet.getRange(row, getColumnByName('linkedinDescription')).getValue(),
    linkedinPreviousJobDescription: sheet.getRange(row, getColumnByName('linkedinPreviousJobDescription')).getValue(),
    linkedinSchoolDescription: sheet.getRange(row, getColumnByName('linkedinSchoolDescription')).getValue(),
    linkedinJobDescription: sheet.getRange(row, getColumnByName('linkedinJobDescription')).getValue(),
    linkedinPreviousSchoolDescription: sheet.getRange(row, getColumnByName('linkedinPreviousSchoolDescription')).getValue()
  };
}

// Format profile data for API
function formatProfileData(profileData) {
  return `
PROFIL CANDIDAT:
1. Informații Generale:
   - Industrie: ${profileData.companyIndustry || 'N/A'}
   - Companie Actuală: ${profileData.companyName || 'N/A'}
   - Titlu LinkedIn: ${profileData.linkedinHeadline || 'N/A'}
   - Locație: ${profileData.location || 'N/A'}

2. Experiență Profesională:
   - Poziție Actuală: ${profileData.linkedinJobTitle || 'N/A'} (${profileData.linkedinJobDateRange || 'N/A'})
   - Descriere: ${profileData.linkedinJobDescription || 'N/A'}
   - Poziție Anterioară: ${profileData.linkedinPreviousJobTitle || 'N/A'} la ${profileData.previousCompanyName || 'N/A'} (${profileData.linkedinPreviousJobDateRange || 'N/A'})
   - Descriere Anterioară: ${profileData.linkedinPreviousJobDescription || 'N/A'}

3. Educație:
   - Studii Actuale: ${profileData.linkedinSchoolName || 'N/A'} - ${profileData.linkedinSchoolDegree || 'N/A'} (${profileData.linkedinSchoolDateRange || 'N/A'})
   - Descriere: ${profileData.linkedinSchoolDescription || 'N/A'}
   - Studii Anterioare: ${profileData.linkedinPreviousSchoolName || 'N/A'} - ${profileData.linkedinPreviousSchoolDegree || 'N/A'} (${profileData.linkedinPreviousSchoolDateRange || 'N/A'})
   - Descriere: ${profileData.linkedinPreviousSchoolDescription || 'N/A'}

4. Competențe și Profil:
   - Competențe: ${profileData.linkedinSkillsLabel || 'N/A'}
   - Descriere Profil: ${profileData.linkedinDescription || 'N/A'}`;
}

/**
 * callGeminiAPI
 *
 * Constructs a request prompt and payload combining candidate profile data and job description,
 * then sends it to the Gemini API. The function implements a retry mechanism to handle API failures,
 * including rate limiting. The complete request prompt and payload are logged (while avoiding sensitive 
 * header details), as is the full response from the API.
 *
 * @param {string} profileData - The formatted candidate profile data.
 * @param {string} jobDescription - The job description used to evaluate the candidate.
 * @returns {Object} - The parsed JSON response from the Gemini API.
 * @throws Will throw an error if the maximum number of retries is exhausted.
 */
function callGeminiAPI(profileData, jobDescription) {
  const maxRetries = 3;
  const retryDelay = 2000; // Delay in milliseconds (increases with each retry)

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // Construct the prompt to be sent to the API
      const prompt = `${profileData}\n\nJOB DESCRIPTION:\n${jobDescription}\n\nTe rog să evaluezi și să furnizezi următoarele:
1. Evaluare Tehnică (0-100): Evaluează potrivirea competențelor tehnice cu cerințele job-ului
2. Evaluare Experiență (0-100): Evaluează relevanța experienței profesionale
3. Scor General (0-100): Calculează compatibilitatea generală
4. Recomandări: Oferă 2-3 sugestii concrete pentru îmbunătățirea profilului

Răspunde strict în următorul format:
Evaluare Tehnică: [scor]
Evaluare Experiență: [scor]
Scor General: [scor]
Recomandări:
- [recomandare 1]
- [recomandare 2]
- [recomandare 3]`;

      // Log the exact prompt that is being sent to the LLM
      logToSheet(`LLM Request Prompt (Attempt ${attempt}): ${prompt}`, 'DEBUG', '');

      // Prepare the payload for the API request
      const payload = {
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
      };

      // Log the constructed payload (excluding sensitive header info)
      logToSheet(`LLM Request Payload (Attempt ${attempt}): ${JSON.stringify(payload)}`, 'DEBUG', '');

      // Set up the HTTP request options for the Gemini API
      const options = {
        method: 'post',
        headers: {
          'x-goog-api-key': GEMINI_API_KEY,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      // Execute the API request
      const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent', options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      // Log the full API response
      logToSheet(`LLM Response (Attempt ${attempt}): ${responseText}`, 'DEBUG', `Response Code: ${responseCode}`);

      const responseData = JSON.parse(responseText);

      // Handle rate limiting by retrying on HTTP 429 response
      if (responseCode === 429) {
        logToSheet(`Rate limit hit on attempt ${attempt}`, 'WARNING');
        if (attempt === maxRetries) throw new Error('RESOURCE_EXHAUSTED');
        Utilities.sleep(retryDelay * attempt);
        continue;
      }

      // Throw an error if the response code is not 200 (OK)
      if (responseCode !== 200) {
        throw new Error(`API returned code ${responseCode}: ${responseText}`);
      }

      return responseData;
    } catch (error) {
      if (attempt === maxRetries) throw error;
      logToSheet(`Attempt ${attempt} failed: ${error.message}`, 'WARNING');
      Utilities.sleep(retryDelay * attempt);
    }
  }
}

// Response parsing
function parseGeminiResponse(response) {
  const content = response.candidates[0].content.parts[0].text;
  
  const data = {
    technicalScore: 0,
    experienceScore: 0,
    overallScore: 0,
    recommendations: []
  };

  // Extract scores
  const technicalMatch = content.match(/Evaluare Tehnică:\s*(\d+)/);
  const experienceMatch = content.match(/Evaluare Experiență:\s*(\d+)/);
  const overallMatch = content.match(/Scor General:\s*(\d+)/);
  
  if (technicalMatch) data.technicalScore = parseInt(technicalMatch[1]);
  if (experienceMatch) data.experienceScore = parseInt(experienceMatch[1]);
  if (overallMatch) data.overallScore = parseInt(overallMatch[1]);

  // Extract recommendations
  const recommendationsMatch = content.match(/Recomandări:[\s\S]*?(?=-\s.*[\s\S]*?){1,3}/g);
  if (recommendationsMatch) {
    data.recommendations = recommendationsMatch[0]
      .replace('Recomandări:', '')
      .split('-')
      .map(r => r.trim())
      .filter(r => r.length > 0);
  }

  return data;
}

// Sheet update function
function updateSheet(rowIndex, data) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Update scores
  sheet.getRange(rowIndex, OUTPUT_COLUMNS.TECHNICAL_SCORE.charCodeAt(0) - 64).setValue(data.technicalScore);
  sheet.getRange(rowIndex, OUTPUT_COLUMNS.EXPERIENCE_SCORE.charCodeAt(0) - 64).setValue(data.experienceScore);
  sheet.getRange(rowIndex, OUTPUT_COLUMNS.OVERALL_SCORE.charCodeAt(0) - 64).setValue(data.overallScore);
  
  // Update recommendations
  sheet.getRange(rowIndex, OUTPUT_COLUMNS.RECOMMENDATIONS.charCodeAt(0) - 64)
    .setValue(data.recommendations.join('\n'));
}

// Profile processing check
function isProfileProcessed(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getRange(rowIndex, OUTPUT_COLUMNS.TECHNICAL_SCORE.charCodeAt(0) - 64, 1, 4).getValues()[0];
  return row.some(cell => cell !== '');
}

// Get Job Description
function getJobDescription() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jobDescSheet = ss.getSheetByName(JOB_DESCRIPTION_SHEET);
  if (!jobDescSheet) return null;
  return jobDescSheet.getRange('A2').getValue();
}

// Validation function
function validateStructure() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const requiredColumns = [
    'companyIndustry',
    'companyName',
    'linkedinHeadline',
    'linkedinJobDateRange',
    'linkedinJobTitle',
    'linkedinPreviousJobDateRange',
    'linkedinPreviousJobTitle',
    'linkedinSkillsLabel',
    'location',
    'previousCompanyName',
    'linkedinSchoolDegree',
    'linkedinSchoolName',
    'linkedinPreviousSchoolDateRange',
    'linkedinPreviousSchoolDegree',
    'linkedinPreviousSchoolName',
    'linkedinSchoolDateRange',
    'linkedinDescription',
    'linkedinPreviousJobDescription',
    'linkedinSchoolDescription',
    'linkedinJobDescription',
    'linkedinPreviousSchoolDescription'
  ];

  const missingColumns = requiredColumns.filter(col => 
    !headers.some(h => h.toString().toLowerCase() === col.toLowerCase())
  );

  if (missingColumns.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Eroare de structură: Lipsesc următoarele coloane: ' + missingColumns.join(', ')
    );
    return false;
  }

  if (!GEMINI_API_KEY) {
    SpreadsheetApp.getUi().alert(
      'Eroare: API Key-ul Google Gemini nu este configurat!'
    );
    return false;
  }

  if (!getJobDescription()) {
    SpreadsheetApp.getUi().alert(
      'Eroare: Job Description-ul nu este configurat!'
    );
    return false;
  }

  return true;
}

// Helper function to get column index by name
function getColumnByName(columnName) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnIndex = headers.findIndex(header => header.toString().toLowerCase() === columnName.toLowerCase());
  return columnIndex + 1;
}

// Error logging
function logError(error, rowIndex) {
  const errorMessage = `Error processing row ${rowIndex}: ${error.message}`;
  console.error(errorMessage);
  logToSheet(errorMessage, 'ERROR');
  
  const sheet = SpreadsheetApp.getActiveSheet();
  const errorDetails = error.message.includes('code') ? error.message : 'Error: ' + error.message;
  
  // Update status column only
  sheet.getRange(rowIndex, OUTPUT_COLUMNS.STATUS.charCodeAt(0) - 64).setValue(errorDetails);
}
