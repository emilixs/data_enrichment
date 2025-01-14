# Specificație Produs - Extensie Google Sheets pentru Îmbogățire Date Companii

## 1. Prezentare Generală
Extensie Google Sheets pentru îmbogățirea automată a datelor despre companii folosind Perplexity API, procesând informațiile rând cu rând și completând automat detaliile lipsă despre companii.

## 2. Structura Datelor

### 2.1 Coloane Input/Output
- B: Companie (input)
- F: Website (output)
- G: Cifra afaceri (output)
- H: Profit (output)
- I: Nr. angajati (output)

### 2.2 Variabile Script
- `PERPLEXITY_API_KEY`: Cheie API stocată în variabilele scriptului
- `PROCESSING_LIMIT`: Setat inițial la 5 rânduri pentru testare

## 3. Funcționalități

### 3.1 Meniu și Interfață
- Buton în meniul superior al Google Sheets
- Denumire: "Îmbogățire Date"
- Submeniu cu opțiunea: "Procesează Companii"

### 3.2 Validări
- Verificare structură coloane înainte de procesare
- Verificare existență API key
- Verificare rânduri deja procesate

### 3.3 Procesare Date
1. Citire nume companie din coloana B
2. Generare prompt Perplexity:
   ```
   Gaseste urmatoarele.
   Numele oficial al companiei [nume companie]
   Codul fiscal
   Cifra de afaceri
   Profit
   Nr de angajati
   Site-ul
   ```
3. Procesare răspuns și mapare în celule

### 3.4 Gestionare Erori
- Rânduri procesate anterior: Skip automat
- Date lipsă: Inserare "N/A" în celulele respective
- Erori procesare: Marcare cu "Failed" în celulele afectate

## 4. Limitări și Constrângeri
- Procesare limitată la 5 rânduri (pentru versiunea de test)
- Funcționează doar pe structura specificată de coloane
- Limba interfață: Română
- Procesare secvențială fără posibilitate de pauză

## 5. Funcții Necesare

### 5.1 Funcții Principale
```javascript
function onOpen() {
    // Creare meniu în interfața Google Sheets
}

function validateStructure() {
    // Validare structură coloane și configurație
}

function processCompanies() {
    // Procesare principală companii
}

function callPerplexityAPI(companyName) {
    // Interogare API Perplexity
}

function parsePerplexityResponse(response) {
    // Parsare răspuns și extragere date relevante
}

function updateSheet(rowIndex, data) {
    // Actualizare celule cu datele primite
}

function isRowProcessed(rowIndex) {
    // Verificare dacă rândul a fost procesat anterior
}
```

### 5.2 Funcții Utilitare
```javascript
function getApiKey() {
    // Recuperare API key din variabile script
}

function logError(error, rowIndex) {
    // Logging erori și marcare rânduri cu probleme
}
```

## 6. Securitate
- API Key stocat în variabilele scriptului Google Apps
- Acces limitat la foaia de calcul specificată
- Validări pentru prevenirea suprascrierii accidentale

## 7. Performanță
- Procesare secvențială a rândurilor
- Pauză între request-uri API pentru evitarea rate limiting
- Limite de procesare pentru testare inițială 