# Specificație Produs - Extensie Google Sheets pentru Analiza Profilelor LinkedIn

## 1. Prezentare Generală
Extensie Google Sheets pentru analiza automată a profilelor LinkedIn folosind Gemini API, procesând informațiile despre candidați și evaluând compatibilitatea acestora cu descrierile de job-uri.

## 2. Structura Datelor

### 2.1 Coloane Input
#### 2.1.1 Coloane pentru Analiza Gemini (folosite în prompt)
1. companyIndustry: Industria companiei
2. companyName: Numele companiei
3. linkedinHeadline: Titlul profilului LinkedIn
4. linkedinJobDateRange: Perioada job actual
5. linkedinJobTitle: Poziția actuală
6. linkedinPreviousJobDateRange: Perioada job anterior
7. linkedinPreviousJobTitle: Poziția anterioară
8. linkedinSkillsLabel: Competențe listate
9. location: Locația
10. previousCompanyName: Compania anterioară
11. linkedinSchoolDegree: Diploma obținută
12. linkedinSchoolName: Numele instituției
13. linkedinPreviousSchoolDateRange: Perioada școală anterioară
14. linkedinPreviousSchoolDegree: Diploma anterioară
15. linkedinPreviousSchoolName: Numele școlii anterioare
16. linkedinSchoolDateRange: Perioada școală actuală
17. linkedinDescription: Descrierea profilului
18. linkedinPreviousJobDescription: Descrierea job anterior
19. linkedinSchoolDescription: Descrierea școlii
20. linkedinJobDescription: Descrierea job actual
21. linkedinPreviousSchoolDescription: Descrierea școlii anterioare

#### 2.1.2 Alte Coloane Input
1. firstName: Prenume
2. lastName: Nume
3. linkedinCompanyUrl: URL-ul companiei pe LinkedIn
4. linkedinCompanySlug: Slug-ul companiei
5. linkedinFollowersCount: Număr de urmăritori
6. linkedinIsHiringBadge: Badge de recrutare activă
7. linkedinIsOpenToWorkBadge: Badge disponibilitate
12. connectionDegree: Grad de conexiune
13. refreshedAt: Data actualizării
14. mutualConnectionsUrl: URL conexiuni comune
15. connectionsUrl: URL conexiuni
16. linkedinConnectionsCount: Număr conexiuni
17. profileUrl: URL profil
18. linkedinSchoolUrl: URL școală
19. linkedinSchoolCompanySlug: Slug instituție
20. linkedinJobLocation: Locația job actual
21. linkedinPreviousSchoolUrl: URL școală anterioară
22. linkedinPreviousSchoolCompanySlug: Slug școală anterioară
23. linkedinPreviousJobLocation: Locația job anterior
24. linkedinPreviousCompanySlug: Slug companie anterioară
25. linkedinPreviousJobDescription: Descrierea job anterior
26. error: Erori de procesare

### 2.2 Coloane Output
- Evaluare Tehnică: Scor bazat pe competențe tehnice (0-100)
- Evaluare Experiență: Scor bazat pe experiența relevantă (0-100)
- Potrivire Job: Scor general de compatibilitate (0-100)
- Recomandări: Sugestii de îmbunătățire
- Status: Statusul procesării

### 2.3 Structura Prompt Gemini
```
Analizează următorul profil profesional și evaluează compatibilitatea cu job description-ul dat:

PROFIL CANDIDAT:
1. Informații Generale:
   - Industrie: [companyIndustry]
   - Companie Actuală: [companyName]
   - Titlu LinkedIn: [linkedinHeadline]
   - Locație: [location]

2. Experiență Profesională:
   - Poziție Actuală: [linkedinJobTitle] ([linkedinJobDateRange])
   - Descriere: [linkedinJobDescription]
   - Poziție Anterioară: [linkedinPreviousJobTitle] la [previousCompanyName] ([linkedinPreviousJobDateRange])
   - Descriere Anterioară: [linkedinPreviousJobDescription]

3. Educație:
   - Studii Actuale: [linkedinSchoolName] - [linkedinSchoolDegree] ([linkedinSchoolDateRange])
   - Descriere: [linkedinSchoolDescription]
   - Studii Anterioare: [linkedinPreviousSchoolName] - [linkedinPreviousSchoolDegree] ([linkedinPreviousSchoolDateRange])
   - Descriere: [linkedinPreviousSchoolDescription]

4. Competențe și Profil:
   - Competențe: [linkedinSkillsLabel]
   - Descriere Profil: [linkedinDescription]

JOB DESCRIPTION:
[job_description]

Te rog să evaluezi și să furnizezi următoarele:
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
- [recomandare 3]
```

## 3. Funcționalități

### 3.1 Meniu și Interfață
- Buton în meniul superior al Google Sheets
- Denumire: "Analiză Profile"
- Submeniu cu opțiunile:
  - "Configurare Job Description"
  - "Procesează Profile"
  - "Resetare Evaluări"

### 3.2 Analiză Job Description și Criterii de Evaluare
- Detectare automată a modificărilor în Job Description
- Extragere criterii de evaluare folosind Gemini API:
  - Identificare automată a 3 criterii principale
  - Generare descrieri și exemple pentru fiecare criteriu
- Management Sheet Criterii de Evaluare:
  - Nume sheet: "Criterii Evaluare CV"
  - Structură:
    - Rând 1: Titluri criterii (ex: "Competențe Tehnice", "Experiență", "Potrivire Culturală")
    - Rând 2: Instrucțiuni evaluare și exemple
  - Actualizare automată la modificarea Job Description
  - Folosit ca referință în evaluarea candidaților

### 3.3 Validări
- Verificare structură coloane înainte de procesare
- Verificare existență API key Gemini
- Verificare existență job description configurat
- Verificare date profil:
  - Câmpuri esențiale:
    - linkedinJobTitle (poziția actuală) - Singurul câmp obligatoriu
  - Câmpuri opționale (înlocuite automat cu "N/A" dacă lipsesc):
    - companyIndustry (industria)
    - companyName (numele companiei)
    - linkedinHeadline (titlul profilului)
    - linkedinJobDateRange (perioada job actual)
    - linkedinPreviousJobDateRange (perioada job anterior)
    - linkedinPreviousJobTitle (poziția anterioară)
    - linkedinSkillsLabel (competențe)
    - location (locația)
    - previousCompanyName (compania anterioară)
    - linkedinSchoolDegree (diploma)
    - linkedinSchoolName (școala)
    - linkedinPreviousSchoolDateRange (perioada școală anterioară)
    - linkedinPreviousSchoolDegree (diploma anterioară)
    - linkedinPreviousSchoolName (școala anterioară)
    - linkedinSchoolDateRange (perioada școală)
    - linkedinDescription (descriere profil)
    - linkedinPreviousJobDescription (descriere job anterior)
    - linkedinSchoolDescription (descriere școală)
    - linkedinJobDescription (descriere job actual)
    - linkedinPreviousSchoolDescription (descriere școală anterioară)
- Validare format date (perioade, URL-uri)
- Logging detaliat pentru câmpurile lipsă:
  - WARNING pentru câmpuri esențiale lipsă
  - INFO pentru câmpuri opționale lipsă

### 3.4 Procesare Date
1. Citire job description configurat
2. Extragere date profil din coloanele specificate
3. Generare prompt Gemini conform structurii definite
4. Procesare răspuns și extragere:
   - Scoruri evaluare (tehnică, experiență, general)
   - Recomandări de îmbunătățire
5. Salvare rezultate în coloanele de output

### 3.5 Gestionare Erori
- Profile procesate anterior: Opțiune de rescriere/skip
- Date lipsă: Marcare în raport, continuare procesare
- Erori API: Retry automat cu exponential backoff
- Rate limiting: Gestionare automată cu pauze
- Logging detaliat în sheet separat

## 4. Limitări și Constrângeri
- Dependență de disponibilitatea API-ului Gemini
- Procesare secvențială a profilelor
- Necesită format specific pentru job description
- Evaluare bazată pe date publice LinkedIn
- Confidențialitate și conformitate GDPR
- Limba: Interfață în Română, procesare în Engleză

## 5. Funcții Necesare

### 5.1 Funcții Principale
```javascript
function onOpen() {
    // Creare meniu în interfața Google Sheets
}

function validateStructure() {
    // Validare structură coloane și configurație
}

function configureJobDescription() {
    // Configurare și salvare job description
}

function processProfiles() {
    // Procesare principală profile
}

function callGeminiAPI(profileData, jobDescription) {
    // Interogare API Gemini cu retry logic
}

function parseGeminiResponse(response) {
    // Parsare răspuns și extragere evaluări
}

function updateSheet(rowIndex, evaluationData) {
    // Actualizare celule cu rezultatele evaluării
}

function isProfileProcessed(rowIndex) {
    // Verificare dacă profilul a fost procesat
}

function resetEvaluations() {
    // Resetare rezultate evaluări anterioare
}

function parseJobDescriptionForCriteria(jobDescription) {
    // Extragere criterii evaluare din job description folosind Gemini API
    // Return: Array de obiecte cu title și description pentru fiecare criteriu
}

function updateEvaluationCriteria(criteria) {
    // Actualizare/creare sheet criterii evaluare
    // Populare cu titluri și descrieri
}

function onJobDescriptionChange() {
    // Handler pentru modificări job description
    // Trigger pentru actualizare criterii
}
```

### 5.2 Funcții Utilitare
```