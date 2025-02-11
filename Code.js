// Constants
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const LOG_SHEET_NAME = 'Logs';
const JOB_DESCRIPTION_SHEET = 'Job Description';
const OUTPUT_COLUMNS = {
  CRITERIA_1_SCORE: 'AV',  // Column 48
  CRITERIA_2_SCORE: 'AW', // Column 49
  CRITERIA_3_SCORE: 'AX',    // Column 50
  RECOMMENDATIONS: 'AY',  // Column 51
  STATUS: 'AZ'           // Column 52
};

const CRITERIA_SHEET_NAME = 'Criterii Evaluare CV';
const CRITERIA_PROMPT = `Analizează următorul job description și extrage cele 3 criterii cele mai importante pentru evaluarea candidaților.
Pentru fiecare criteriu, oferă un titlu și o descriere detaliată cu exemple concrete, acestea vor fi formulate ca si prompturi pentru identificarea parametrilor respectivi in CV-uri, te rog criteriile sa fie doar lucruri pe care te astepti sa le poti evalua in CV-uri.
Prompturile pentru fiecare criteriu sa fie specifice cu exemple de ce inseamna 0,30,60,90,100 ca si scor.

Job Description:
[JOB_DESCRIPTION]

Răspunde strict în următorul format:
Criteriu 1:
Titlu: [titlu]
Prompt: [descriere detaliată cu exemple]

Criteriu 2:
Titlu: [titlu]
Prompt: [descriere detaliată cu exemple]

Criteriu 3:
Titlu: [titlu]
Prompt: [descriere detaliată cu exemple]`;

// Add this at the top with other constants
const COLUMN_MAPPING = {
  'companyIndustry': 'A',
  'companyName': 'B',
  'firstName': 'C',
  'lastName': 'D',
  'linkedinCompanyUrl': 'E',
  'linkedinCompanySlug': 'F',
  'linkedinFollowersCount': 'G',
  'linkedinHeadline': 'H',
  'linkedinIsHiringBadge': 'I',
  'linkedinIsOpenToWorkBadge': 'J',
  'linkedinJobDateRange': 'K',
  'linkedinJobTitle': 'L',
  'linkedinPreviousJobDateRange': 'M',
  'linkedinPreviousJobTitle': 'N',
  'linkedinProfileSlug': 'P',
  'linkedinProfileUrl': 'Q',
  'linkedinProfileUrn': 'R',
  'linkedinSkillsLabel': 'S',
  'location': 'T',
  'previousCompanyName': 'U',
  'connectionDegree': 'V',
  'refreshedAt': 'W',
  'mutualConnectionsUrl': 'X',
  'connectionsUrl': 'Y',
  'linkedinConnectionsCount': 'Z',
  'profileUrl': 'AA',
  'linkedinSchoolUrl': 'AB',
  'linkedinSchoolCompanySlug': 'AC',
  'linkedinSchoolDegree': 'AD',
  'linkedinSchoolName': 'AE',
  'linkedinJobLocation': 'AF',
  'linkedinPreviousSchoolUrl': 'AG',
  'linkedinPreviousSchoolCompanySlug': 'AH',
  'linkedinPreviousSchoolDateRange': 'AI',
  'linkedinPreviousSchoolDegree': 'AJ',
  'linkedinPreviousSchoolName': 'AK',
  'linkedinSchoolDateRange': 'AL',
  'linkedinDescription': 'AM',
  'linkedinPreviousJobLocation': 'AN',
  'linkedinPreviousCompanySlug': 'AO',
  'linkedinPreviousJobDescription': 'AP',
  'linkedinSchoolDescription': 'AQ',
  'linkedinJobDescription': 'AR',
  'linkedinPreviousSchoolDescription': 'AS',
  'error': 'AT' // First error column
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
      logSheet.getRange('A1:G1').setValues([['Timestamp', 'Level', 'Message', 'Details', 'Prompt', 'Response', 'Conclusions']]);
      logSheet.setFrozenRows(1);
      ss.setActiveSheet(activeSheet); // Revert back to the originally active sheet
    }

    const timestamp = new Date().toISOString();
    logSheet.appendRow([timestamp, level, message, details, '', '', '']);

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

/**
 * logLLMInteraction
 *
 * Specialized logging function for LLM (Language Model) interactions.
 * Logs the prompt sent to the LLM, the response received, and any conclusions drawn.
 *
 * @param {string} prompt - The prompt sent to the LLM
 * @param {string} response - The response received from the LLM
 * @param {string} conclusions - Any conclusions or processed results from the response
 * @param {string} [level="INFO"] - Log level
 * @param {string} [message="LLM Interaction"] - Additional message for context
 */
function logLLMInteraction(prompt, response, conclusions, level = 'INFO', message = 'LLM Interaction') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    let logSheet = ss.getSheetByName('Logs');

    // Create log sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet('Logs');
      logSheet.getRange('A1:G1').setValues([['Timestamp', 'Level', 'Message', 'Details', 'Prompt', 'Response', 'Conclusions']]);
      logSheet.setFrozenRows(1);
      ss.setActiveSheet(activeSheet);
    }

    const timestamp = new Date().toISOString();
    logSheet.appendRow([
      timestamp,
      level,
      message,
      '',
      prompt,
      response,
      conclusions
    ]);

    // Format the cells for better readability
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow, 5, 1, 3).setWrap(true); // Enable text wrapping for prompt, response, and conclusions
    
    // Adjust row height to fit content
    logSheet.setRowHeight(lastRow, -1); // Auto-resize row height

    // Keep only last 1000 logs
    const maxRows = 1000;
    const currentRows = logSheet.getLastRow();
    if (currentRows > maxRows) {
      logSheet.deleteRows(2, currentRows - maxRows);
    }

    ss.setActiveSheet(activeSheet);
  } catch (error) {
    console.error('LLM logging failed:', error);
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
    
  // Set up output column headers if they don't exist
  setupOutputColumns();
}

/**
 * Extract content from a Google Docs URL
 * @param {string} url - The Google Docs URL
 * @returns {string} The document's content
 * @throws {Error} If the document cannot be accessed or URL is invalid
 */
function extractGoogleDocsContent(url) {
  try {
    // Extract the document ID from the URL
    const regex = /\/d\/([a-zA-Z0-9-_]+)/;
    const match = url.match(regex);
    
    if (!match) {
      throw new Error('URL invalid. Vă rugăm să folosiți un URL valid de Google Docs.');
    }
    
    const docId = match[1];
    
    // Try to open the document
    try {
      const doc = DocumentApp.openByUrl(url);
      return doc.getBody().getText();
    } catch (e) {
      // If that fails, try using the advanced Drive API
      const file = DriveApp.getFileById(docId);
      const content = file.getBlob().getDataAsString();
      
      // Clean up any HTML/formatting if present
      return content.replace(/<[^>]*>/g, '').trim();
    }
  } catch (error) {
    logToSheet('Failed to extract Google Docs content', 'ERROR', error.message);
    throw new Error('Nu am putut accesa documentul. Verificați că URL-ul este corect și că documentul este partajat pentru acces.');
  }
}

// Job Description Configuration with criteria extraction
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
    'Introduceți URL-ul documentului Google Docs care conține descrierea job-ului:',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const url = response.getResponseText().trim();
    
    // Basic URL validation
    if (!url.includes('docs.google.com/')) {
      ui.alert(
        'Eroare',
        'URL-ul trebuie să fie un document Google Docs valid.',
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      // Extract content from Google Docs
      const jobDescription = extractGoogleDocsContent(url);
      
      // Validate job description is not empty
      if (!jobDescription) {
        ui.alert(
          'Eroare',
          'Documentul este gol sau nu conține text valid.',
          ui.ButtonSet.OK
        );
        return;
      }

      // Store both the URL and the content
      jobDescSheet.getRange('A2').setValue(url);
      jobDescSheet.getRange('B2').setValue(jobDescription);
      
      // Log the stored job description for verification
      logToSheet(
        'Job Description stored',
        'INFO',
        `URL: ${url}\nContent length: ${jobDescription.length} characters\nFirst 100 chars: ${jobDescription.substring(0, 100)}...`
      );
      
      // Extract and update evaluation criteria
      try {
        const criteria = parseJobDescriptionForCriteria(jobDescription);
        updateEvaluationCriteria(criteria);
        ui.alert('Job Description și criteriile de evaluare au fost salvate cu succes!');
      } catch (error) {
        logToSheet('Error extracting criteria', 'ERROR', error.message);
        ui.alert('Job Description salvat, dar a apărut o eroare la extragerea criteriilor. Verificați log-urile pentru detalii.');
      }
    } catch (error) {
      logToSheet('Error configuring job description', 'ERROR', error.message);
      ui.alert('Eroare', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Parse job description to extract evaluation criteria using Gemini API
 * @param {string} jobDescription - The job description text
 * @returns {Array<Object>} Array of criteria objects with title and description
 */
function parseJobDescriptionForCriteria(jobDescription) {
  // Log the raw job description content being used for parsing
  logToSheet(
    'Parsing job description for criteria',
    'DEBUG',
    `Job Description content: ${jobDescription}`
  );

  // Clean up the job description text
  const cleanJobDescription = jobDescription.trim();
  if (!cleanJobDescription) {
    throw new Error('Job description is empty');
  }

  // Replace the placeholder in the criteria prompt with the actual job description content
  const prompt = CRITERIA_PROMPT.replace('[JOB_DESCRIPTION]', cleanJobDescription);

  // Log the constructed prompt before sending it to Gemini API
  logToSheet('Sending prompt for criteria extraction', 'DEBUG', `Prompt: ${prompt}`);
  
  try {
    // Call Gemini API with the prompt that includes the actual job description text
    const response = callGeminiAPI(prompt, '');
    
    // Verify we got a valid response
    if (!response?.candidates?.[0]?.content?.parts?.[0]?.text) {
      throw new Error('Invalid response format from Gemini API');
    }
    
    const content = response.candidates[0].content.parts[0].text;

    // Log the response content received from Gemini
    logToSheet('Received response for criteria extraction', 'DEBUG', `Response content: ${content}`);

    // Parse the response into structured criteria
    const criteria = [];
    const criteriaMatches = content.match(/Criteriu \d+:\nTitlu: (.*?)\nPrompt: (.*?)(?=\n\nCriteriu|\n*$)/gs);

    if (!criteriaMatches || criteriaMatches.length !== 3) {
      throw new Error('Invalid criteria format in API response');
    }

    criteriaMatches.forEach((match, index) => {
      const titleMatch = match.match(/Titlu: (.*?)\n/);
      const promptMatch = match.match(/Prompt: (.*?)$/s);

      if (titleMatch && promptMatch) {
        criteria.push({
          title: titleMatch[1].trim(),
          description: promptMatch[1].trim()
        });
        
        // Log each extracted criterion
        logToSheet(
          'Extracted criterion',
          'DEBUG',
          `Criterion ${index + 1}: Title="${titleMatch[1].trim()}", Prompt="${promptMatch[1].trim()}"`
        );
      }
    });

    // Log the final extracted criteria
    logToSheet('Extracted evaluation criteria', 'DEBUG', JSON.stringify(criteria, null, 2));

    return criteria;
  } catch (error) {
    logToSheet('Failed to parse job description for criteria', 'ERROR', error.message);
    throw error;
  }
}

/**
 * Update or create the evaluation criteria sheet with extracted criteria
 * @param {Array<Object>} criteria - Array of criteria objects with title and description
 */
function updateEvaluationCriteria(criteria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let criteriaSheet = ss.getSheetByName(CRITERIA_SHEET_NAME);
  
  // Create or clear the sheet
  if (!criteriaSheet) {
    criteriaSheet = ss.insertSheet(CRITERIA_SHEET_NAME);
  } else {
    criteriaSheet.clear();
  }
  
  // Set up headers
  const headers = criteria.map(c => c.title);
  criteriaSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Set up descriptions
  const descriptions = criteria.map(c => c.description);
  criteriaSheet.getRange(2, 1, 1, descriptions.length).setValues([descriptions]);
  
  // Format the sheet
  criteriaSheet.setFrozenRows(1);
  criteriaSheet.getRange(1, 1, 2, headers.length).setWrap(true);
  criteriaSheet.autoResizeColumns(1, headers.length);
  
  logToSheet('Updated evaluation criteria', 'INFO', `Added ${criteria.length} criteria`);
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
      // Clear content from AV to AZ for all rows except header
      const startCol = columnToNumber('AV');
      const numCols = 5; // AV to AZ
      sheet.getRange(2, startCol, lastRow - 1, numCols).clearContent();
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
  let skippedCount = 0;
  const PROFILE_LIMIT = 5;

  // Get Job Description
  const jobDesc = getJobDescription();
  if (!jobDesc) {
    logToSheet('Job Description not configured', 'ERROR');
    ui.alert('Error', 'Job Description nu este configurat!', ui.ButtonSet.OK);
    return;
  }

  logToSheet(`Starting to process profiles. Will process up to ${PROFILE_LIMIT} unprocessed profiles`, 'INFO');

  try {
    for (let row = 2; row <= lastRow; row++) {
      // Check if we've hit the processing limit
      if (processedCount >= PROFILE_LIMIT) {
        logToSheet(`Reached processing limit of ${PROFILE_LIMIT} profiles`, 'INFO');
        break;
      }

      if (!isProfileProcessed(row)) {
        const profileData = extractProfileData(row);
        if (validateProfileData(profileData)) {
          try {
            logToSheet(`Processing profile at row ${row} (${processedCount + 1}/${PROFILE_LIMIT})`, 'INFO');
            
            const formattedData = formatProfileData(profileData);
            const response = callGeminiAPI(formattedData, jobDesc);
            
            const evaluationData = parseGeminiResponse(response);
            updateSheet(row, evaluationData);
            processedCount++;
            
            // Add delay between requests
            if (processedCount < PROFILE_LIMIT) {
              Utilities.sleep(2000);
            }
          } catch (error) {
            logToSheet(`Error processing row ${row}`, 'ERROR', error.message);
            logError(error, row);
            
            if (error.message.includes('RESOURCE_EXHAUSTED')) {
              const waitMinutes = 2;
              const waitMessage = `Rate limit atins. Așteptăm ${waitMinutes} minute...`;
              logToSheet(waitMessage, 'WARNING');
              sheet.getRange(row, OUTPUT_COLUMNS.STATUS.charCodeAt(0) - 64)
                .setValue(`Eroare: ${waitMessage}`);
              Utilities.sleep(waitMinutes * 60 * 1000);
              row--; // Retry the same row
              continue;
            }
          }
        } else {
          // Update status for invalid profiles
          sheet.getRange(row, OUTPUT_COLUMNS.STATUS.charCodeAt(0) - 64)
            .setValue('Date profil invalide sau incomplete');
        }
      } else {
        logToSheet(`Skipping processed profile at row ${row}`, 'INFO');
        skippedCount++;
      }
    }
  } catch (error) {
    logToSheet('Processing failed', 'ERROR', error.message);
    throw error;
  }

  logToSheet('Processing completed', 'INFO', 
    `Processed: ${processedCount}, Skipped: ${skippedCount}, Total rows checked: ${lastRow - 1}`);
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
 * Gets the evaluation criteria from the Criterii Evaluare CV sheet
 * @returns {Array<Object>} Array of criteria objects with title and prompt
 */
function getEvaluationCriteria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const criteriaSheet = ss.getSheetByName(CRITERIA_SHEET_NAME);
  
  if (!criteriaSheet) {
    throw new Error('Criteriile de evaluare nu au fost configurate. Vă rugăm să configurați mai întâi job description-ul.');
  }
  
  const titles = criteriaSheet.getRange(1, 1, 1, 3).getValues()[0];
  const prompts = criteriaSheet.getRange(2, 1, 1, 3).getValues()[0];
  
  return titles.map((title, index) => ({
    title: title,
    prompt: prompts[index]
  }));
}

/**
 * Gets the stored job description text
 * @returns {string} The job description text
 */
function getJobDescription() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jobDescSheet = ss.getSheetByName(JOB_DESCRIPTION_SHEET);
  
  if (!jobDescSheet) {
    throw new Error('Job Description sheet not found');
  }
  
  const jobDesc = jobDescSheet.getRange('B2').getValue();
  if (!jobDesc) {
    throw new Error('Job Description not configured');
  }
  
  return jobDesc;
}

/**
 * callGeminiAPI
 *
 * Constructs a request prompt and payload combining candidate profile data and the job description,
 * then sends it to the Gemini API. It logs the full prompt and payload sent to the LLM and the complete response.
 * Implements a retry mechanism in case of API failures or rate limiting.
 *
 * @param {string} prompt - The prompt to send to the API
 * @param {string} jobDescription - The job description (only used for profile evaluation, not for criteria extraction)
 * @returns {Object} - The parsed JSON response from the Gemini API.
 * @throws Will throw an error if the maximum number of retries is exhausted.
 */
function callGeminiAPI(prompt, jobDescription) {
  const maxRetries = 3;
  const retryDelay = 2000;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // Only add the evaluation part if we're evaluating a profile (when jobDescription is not empty)
      let finalPrompt;
      if (jobDescription) {
        const criteria = getEvaluationCriteria();
        finalPrompt = `${prompt}\n\nJOB DESCRIPTION:\n${jobDescription}\n\nTe rog să evaluezi candidatul pe baza următoarelor criterii:\n\n`;
        
        criteria.forEach((criterion, index) => {
          finalPrompt += `${index + 1}. ${criterion.title} (0-100):\n${criterion.prompt}\n\n`;
        });
        
        finalPrompt += `Răspunde strict în următorul format:
${criteria[0].title}: [scor]
${criteria[1].title}: [scor]
${criteria[2].title}: [scor]
Recomandări:
- [recomandare 1]
- [recomandare 2]
- [recomandare 3]`;
      } else {
        finalPrompt = prompt;
      }

      const payload = {
        'contents': [{
          'parts': [{
            'text': finalPrompt
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

      const options = {
        method: 'post',
        headers: {
          'x-goog-api-key': GEMINI_API_KEY,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent', options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      const responseData = JSON.parse(responseText);

      // Log the LLM interaction
      if (responseCode === 200) {
        const conclusions = responseData.candidates[0].content.parts[0].text;
        logLLMInteraction(
          finalPrompt,
          responseText,
          conclusions,
          'INFO',
          `LLM Request (Attempt ${attempt})`
        );
      } else {
        logLLMInteraction(
          finalPrompt,
          responseText,
          `Error: Response code ${responseCode}`,
          'ERROR',
          `Failed LLM Request (Attempt ${attempt})`
        );
      }

      if (responseCode === 429) {
        logToSheet(`Rate limit hit on attempt ${attempt}`, 'WARNING');
        if (attempt === maxRetries) throw new Error('RESOURCE_EXHAUSTED');
        Utilities.sleep(retryDelay * attempt);
        continue;
      }

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

// Update parseGeminiResponse to handle dynamic criteria
function parseGeminiResponse(response) {
  const content = response.candidates[0].content.parts[0].text;
  const criteria = getEvaluationCriteria();
  
  const data = {
    criteriaScores: [],
    recommendations: []
  };

  // Extract scores for each criterion
  criteria.forEach(criterion => {
    const scoreMatch = content.match(new RegExp(`${criterion.title}:\\s*(\\d+)`));
    if (scoreMatch) {
      data.criteriaScores.push({
        title: criterion.title,
        score: parseInt(scoreMatch[1])
      });
    }
  });

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

// Helper function to convert column letters to numbers
function columnToNumber(column) {
  let result = 0;
  const length = column.length;
  
  for (let i = 0; i < length; i++) {
    result += (column.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  
  return result;
}

// Add a test function to verify column calculations
function testColumnCalculations() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const testColumns = {
    'AV': 48,  // Expected column number
    'AW': 49,
    'AX': 50,
    'AY': 51,
    'AZ': 52
  };
  
  for (const [col, expected] of Object.entries(testColumns)) {
    const calculated = columnToNumber(col);
    logToSheet(
      'Column number calculation test',
      'DEBUG',
      `Column ${col}: Expected ${expected}, Got ${calculated}`
    );
  }
}

// Update the sheet update function to handle the new format
function updateSheet(rowIndex, data) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Update scores for each criterion
  data.criteriaScores.forEach((criteriaScore, index) => {
    const column = Object.values(OUTPUT_COLUMNS)[index]; // Get the corresponding output column
    sheet.getRange(rowIndex, columnToNumber(column)).setValue(criteriaScore.score);
  });
  
  // Update recommendations
  sheet.getRange(rowIndex, columnToNumber(OUTPUT_COLUMNS.RECOMMENDATIONS))
    .setValue(data.recommendations.join('\n'));

  // Update status/conclusions
  const conclusion = `Evaluare completă - ${data.criteriaScores.map(c => `${c.title}: ${c.score}`).join(', ')}`;
  sheet.getRange(rowIndex, columnToNumber(OUTPUT_COLUMNS.STATUS)).setValue(conclusion);
}

// Profile processing check
function isProfileProcessed(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Get only the three score columns (Technical, Experience, Overall)
  const scoreColumns = [
    OUTPUT_COLUMNS.CRITERIA_1_SCORE,
    OUTPUT_COLUMNS.CRITERIA_2_SCORE,
    OUTPUT_COLUMNS.CRITERIA_3_SCORE
  ];
  
  // Log the actual column letters we're checking
  logToSheet(
    `Checking if row ${rowIndex} is processed`,
    'DEBUG',
    `Checking columns: ${scoreColumns.join(', ')}`
  );
}

/**
 * Validates the structure of the active sheet and required sheets
 * Checks for required columns and sheets
 * @returns {boolean} true if structure is valid, false otherwise
 */
function validateStructure() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Check if Job Description sheet exists and has content
    const jobDescSheet = ss.getSheetByName(JOB_DESCRIPTION_SHEET);
    if (!jobDescSheet) {
      ui.alert('Error', 'Sheet-ul "Job Description" nu există. Configurați mai întâi job description-ul.', ui.ButtonSet.OK);
      return false;
    }
    
    // Check if Criteria sheet exists and has content
    const criteriaSheet = ss.getSheetByName(CRITERIA_SHEET_NAME);
    if (!criteriaSheet) {
      ui.alert('Error', 'Sheet-ul "Criterii Evaluare CV" nu există. Configurați mai întâi job description-ul.', ui.ButtonSet.OK);
      return false;
    }
    
    // Verify criteria sheet has 3 columns with titles and prompts
    const criteriaRange = criteriaSheet.getRange(1, 1, 2, 3);
    const criteriaValues = criteriaRange.getValues();
    if (criteriaValues[0].some(title => !title) || criteriaValues[1].some(prompt => !prompt)) {
      ui.alert('Error', 'Sheet-ul "Criterii Evaluare CV" nu este configurat corect. Lipsesc titluri sau prompturi pentru criterii.', ui.ButtonSet.OK);
      return false;
    }
    
    // Check if all required input columns exist in the active sheet
    const sheet = SpreadsheetApp.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Essential columns that must exist
    const essentialColumns = ['linkedinJobTitle'];
    const missingColumns = essentialColumns.filter(col => !headers.includes(col));
    
    if (missingColumns.length > 0) {
      ui.alert('Error', `Lipsesc coloanele esențiale: ${missingColumns.join(', ')}`, ui.ButtonSet.OK);
      return false;
    }
    
    // Check if output columns exist and are properly set up
    const outputColumnLetters = Object.values(OUTPUT_COLUMNS);
    const lastColumn = sheet.getLastColumn();
    const lastColumnLetter = columnToLetter(lastColumn);
    
    const missingOutputColumns = outputColumnLetters.filter(col => 
      columnToNumber(col) > lastColumn || 
      !sheet.getRange(`${col}1`).getValue()
    );
    
    if (missingOutputColumns.length > 0) {
      setupOutputColumns(); // Attempt to set up missing output columns
    }
    
    return true;
  } catch (error) {
    logToSheet('Structure validation failed', 'ERROR', error.message);
    ui.alert('Error', 'A apărut o eroare la validarea structurii: ' + error.message, ui.ButtonSet.OK);
    return false;
  }
}

/**
 * Converts a column number to letter reference
 * @param {number} column - The column number
 * @returns {string} The column letter reference
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Sets up the output columns with proper headers based on evaluation criteria
 * Creates or updates headers for score columns, recommendations, and status
 */
function setupOutputColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Try to get criteria titles from the criteria sheet
    let criteriaTitles = [];
    const criteriaSheet = ss.getSheetByName(CRITERIA_SHEET_NAME);
    if (criteriaSheet) {
      criteriaTitles = criteriaSheet.getRange(1, 1, 1, 3).getValues()[0];
    }
    
    // Set up headers for all output columns
    const outputColumns = Object.values(OUTPUT_COLUMNS);
    outputColumns.forEach((col, index) => {
      const cell = sheet.getRange(`${col}1`);
      let headerValue;
      
      if (index < 3) { // First three columns are for criteria scores
        headerValue = criteriaTitles[index] || `Criteriu ${index + 1}`;
      } else if (index === 3) {
        headerValue = 'Recomandări';
      } else {
        headerValue = 'Status';
      }
      
      cell.setValue(headerValue);
    });
    
    // Format headers
    const headerRange = sheet.getRange('1:1');
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
    
    // Auto-resize columns to fit content
    outputColumns.forEach(col => {
      sheet.autoResizeColumn(columnToNumber(col));
    });
    
    logToSheet('Output columns setup completed', 'INFO', 
      `Set up headers for columns: ${outputColumns.join(', ')}`);
      
  } catch (error) {
    logToSheet('Failed to setup output columns', 'ERROR', error.message);
    throw new Error('Nu am putut configura coloanele de output: ' + error.message);
  }
}