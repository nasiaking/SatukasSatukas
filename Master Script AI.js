// ------------------- File: AI.gs (Logika AI) -------------------

// #region AI - MAIN ENTRY POINT
// =================================================================
//                    AI - MAIN ENTRY POINT
// =================================================================

/**
 * ENTRY: Dipanggil frontend. Sekarang mendukung konteks owner & intent eksplisit.
 * @param {string} userQuestion
 * @param {string} walletOwner (wajib dari UI – environment AI sendiri)
 * @param {string} intent One of: DirectQuery|ComparisonQuery|AnalyticalQuery|GeneralAdvisory
 */
function getAiResponse(userQuestion, walletOwner, intent) {
  userQuestion = (userQuestion || '').trim();
  walletOwner = (walletOwner || '').trim();
  intent = (intent || '').trim();
  if (!userQuestion) return 'Pertanyaan kosong.';
  if (!walletOwner) return 'Pilih Owner terlebih dahulu.';
  if (!intent) return 'Pilih Question Type / Intent.';

  const properties = PropertiesService.getScriptProperties();
  const MAIN_SHEET_ID = properties.getProperty('MAIN_SHEET_ID');
  const QUERIES_SHEET_NAME = 'Queries';
  const DAILY_LIMIT = 10; // simple per-sheet total limit

  const queryId = 'ai-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
  const now = new Date();

  // Hitung limit harian (total semua owner) – sederhana dulu
  let todayUsageCount = 0;
  try {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheet = ss.getSheetByName(QUERIES_SHEET_NAME);
    if (sheet) {
      const values = sheet.getDataRange().getValues();
      if (values.length > 1) {
        const headers = values.shift();
        const tsIdx = headers.indexOf('Timestamp');
        const today = new Date(); today.setHours(0,0,0,0);
        for (let i = values.length - 1; i >= 0; i--) {
          const row = values[i];
            const d = new Date(row[tsIdx]);
            if (d < today) break;
            todayUsageCount++;
        }
      }
    }
  } catch(e) {
    console.warn('Daily limit check failed:', e && e.message);
  }
  if (todayUsageCount >= DAILY_LIMIT) {
    logToQueriesSheet(queryId, now, walletOwner, intent, userQuestion, 'Limit reached');
    return 'Maaf, Anda telah mencapai batas penggunaan AI harian (10 kali).';
  }

  try {
    const answer = routeUserQuestion_(walletOwner, userQuestion, intent); // intent override
    logToQueriesSheet(queryId, now, walletOwner, intent, userQuestion, answer);
    return answer;
  } catch(err) {
    const msg = 'Maaf, terjadi kesalahan saat memproses permintaan AI.';
    logToQueriesSheet(queryId, now, walletOwner, intent, userQuestion, 'Error: '+ (err && err.message));
    console.error('Error di getAiResponse:', err && err.stack || err);
    return msg;
  }
}

// Logging ke sheet Queries (kompatibel versi lama tanpa kolom Intent)
function logToQueriesSheet(queryId, timestamp, walletOwner, intent, userQuestion, aiAnswer) {
  try {
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.openById(props.getProperty('MAIN_SHEET_ID'));
    const sheet = ss.getSheetByName('Queries');
    if (!sheet) throw new Error("Sheet 'Queries' not found");
    const row = [queryId, timestamp, walletOwner, intent, userQuestion, aiAnswer];
    // Jika header belum punya Intent (5 kolom), tetap append 6 kolom (baru).
    sheet.appendRow(row);
  } catch(e) { console.error('Failed to log to Queries sheet:', e && e.message); }
}

/**
 * Ambil history Q/A (default 30 terbaru) filter by walletOwner + optional intent.
 * @param {string} walletOwner
 * @param {string} intentFilter 'All' atau salah satu intent
 * @param {number} limit jumlah maksimum
 * @returns {Array<{queryId:string,timestamp:Date,walletOwner:string,intent:string,question:string,answer:string}>}
 */
function getAIQueryHistory(walletOwner, intentFilter, limit) {
  walletOwner = (walletOwner||'').trim();
  intentFilter = intentFilter || 'All';
  limit = limit || 30;
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('MAIN_SHEET_ID'));
  const sheet = ss.getSheetByName('Queries');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift();
  const idxQuery = headers.indexOf('QueryID');
  const idxTs = headers.indexOf('Timestamp');
  const idxOwner = headers.indexOf('WalletOwner');
  const idxIntent = headers.indexOf('Intent'); // may be -1
  // Backward compatibility headers w/o Intent
  // Pattern lama: [QueryID, Timestamp, WalletOwner, UserQuestion, AI_Answer]
  const idxQuestion = headers.indexOf('UserQuestion');
  const idxAnswer = headers.indexOf('AI_Answer');

  const out = [];
  for (let i = values.length - 1; i >= 0; i--) { // iterate reverse (newest last assumed)
    const r = values[i];
    const ownerVal = r[idxOwner];
    if (walletOwner && ownerVal !== walletOwner) continue;
    const intentVal = idxIntent >= 0 ? (r[idxIntent] || 'Unknown') : 'Unknown';
    if (intentFilter && intentFilter !== 'All' && intentVal !== intentFilter) continue;
    out.push({
      queryId: idxQuery >=0 ? r[idxQuery] : '',
      timestamp: idxTs >=0 ? new Date(r[idxTs]) : null,
      walletOwner: ownerVal || '',
      intent: intentVal,
      question: idxQuestion >=0 ? r[idxQuestion] : '',
      answer: idxAnswer >=0 ? r[idxAnswer] : ''
    });
    if (out.length >= limit) break;
  }
  // sort newest first by timestamp
  out.sort((a,b)=> (b.timestamp?.getTime()||0) - (a.timestamp?.getTime()||0));
  return out;
}

// #endregion

// #region AI - USAGE LIMIT CHECK
// =================================================================
//                 AI - USAGE LIMIT CHECK
// =================================================================

/**
 * Memeriksa apakah pengguna telah mencapai batas penggunaan API harian.
 * @param {string} customerEmail Email pengguna yang akan diperiksa.
 * @returns {boolean} True jika batas belum tercapai, false jika sudah.
 */
function checkApiUsageLimit_(customerEmail) {
  const properties = PropertiesService.getScriptProperties();
  const MAIN_SHEET_ID = properties.getProperty('MAIN_SHEET_ID');
  const QUERIES_SHEET_NAME = "Queries";
  const DAILY_LIMIT = 10;

  try {
    const logSpreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const queriesSheet = logSpreadsheet.getSheetByName(QUERIES_SHEET_NAME);
    if (!queriesSheet) return true; // Failsafe

    const data = queriesSheet.getDataRange().getValues();
    const headers = data.shift();
    const timestampCol = headers.indexOf("Timestamp");

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let todayUsageCount = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const logDate = new Date(row[timestampCol]);
      if (logDate < today) break;
      todayUsageCount++;
    }
    return todayUsageCount < DAILY_LIMIT;
  } catch (e) {
    console.error("Error checking usage limit in Queries sheet: " + e.message);
    return true; // Failsafe
  }
}

// #endregion

// #region AI - ROUTING & HANDLERS
// =================================================================
//                  AI - ROUTING & HANDLERS
// =================================================================

/**
 * Mengklasifikasikan niat pengguna dan mengarahkannya ke handler yang tepat.
 * @param {string} walletOwner Pengguna yang bertanya (email).
 * @param {string} userQuestion Pertanyaan pengguna.
 * @returns {string} Jawaban yang dihasilkan oleh handler yang sesuai.
 */
function routeUserQuestion_(walletOwner, userQuestion, intentOverride) {
  const globalContext = getGlobalContext_();
  let intent = (intentOverride || '').trim();
  if (!intent) {
    const classificationPrompt = `Classify the user's financial question into "DirectQuery", "ComparisonQuery", "AnalyticalQuery", or "GeneralAdvisory". Output ONLY the type. Question: "${userQuestion}"`;
    try { intent = callGeminiApi_(classificationPrompt, false).trim(); } catch(e){ intent = 'AnalyticalQuery'; }
  }
  const entities = { originalQuestion: userQuestion, walletOwner, context: globalContext };
  switch (intent) {
    case 'DirectQuery': return handleDirectQuery_(entities);
    case 'ComparisonQuery': return handleComparisonQuery_(entities);
    case 'AnalyticalQuery': return handleAnalyticalQuery_(entities);
    case 'GeneralAdvisory': return handleGeneralAdvisory_(entities);
    default: return handleAnalyticalQuery_(entities);
  }
}

/**
 * Menjawab pertanyaan analisis, prediksi, atau insight.
 */
function handleAnalyticalQuery_(entities) {
  const extractedFilters = extractEntities_(entities.originalQuestion, entities.walletOwner, entities.context);
  const today = new Date();
  const pastDate = new Date(); pastDate.setDate(today.getDate() - 90);
  const futureDate = new Date(); futureDate.setDate(today.getDate() + 90);
  const historicalData = getFilteredData_({ ...extractedFilters, startDate: pastDate, endDate: today });
  const scheduledData = getScheduledData_({ ...extractedFilters, startDate: today, endDate: futureDate });
  const historicalDataAsText = historicalData.length ? formatDataForAI_(historicalData, false) : 'Tidak ada data historis yang relevan.';
  const scheduledDataAsText = scheduledData.length ? formatDataForAI_(scheduledData, true) : 'Tidak ada data terjadwal yang relevan.';
  const analysisPrompt = `You are a proactive and smart financial advisor. Provide analysis, predictions, or insights to answer the user's question based on the historical and scheduled data.
Rules:
- Use historical data to identify patterns.
- Use scheduled data for future commitments.
- Combine both for a comprehensive cash flow view.
- Answer must be practical & actionable.
- Format dates 'DD MMM YYYY' and currency as 'IDR 1,234,500'.
- Use bullet points (-) and bold important terms.
---
FINANCIAL CONTEXT:
${entities.context}
---
HISTORICAL DATA (Last 90 days):
${historicalDataAsText}
---
SCHEDULED DATA (Next 90 days):
${scheduledDataAsText}
---
USER QUESTION: "${entities.originalQuestion}"
---
PROACTIVE ANALYSIS:`;
  if (historicalData.length < 3 && scheduledData.length === 0) {
    return 'Maaf, data transaksi tidak cukup untuk saya analisis secara mendalam.';
  }
  return callGeminiApi_(analysisPrompt, false);
}

/**
 * Menjawab pertanyaan langsung (total, rincian, daftar).
 */
function handleDirectQuery_(entities) {
  const extractedFilters = extractEntities_(entities.originalQuestion, entities.walletOwner, entities.context);
  let transactionData = getFilteredData_(extractedFilters);
  let isScheduled = false;

  // Fallback: Jika tidak ada di data historis, cek data terjadwal
  if (transactionData.length === 0) {
    transactionData = getScheduledData_(extractedFilters);
    if (transactionData.length > 0) isScheduled = true;
  }

  if (transactionData.length === 0) {
    return `Maaf, saya tidak menemukan data transaksi dengan filter yang Anda berikan.`;
  }

  const dataAsText = formatDataForAI_(transactionData, isScheduled);
  const dataTypeMessage = isScheduled ? "DATA TRANSAKSI TERJADWAL (AKAN DATANG)" : "DATA HISTORIS";

  const analysisPrompt = `
    Anda adalah analis data. Jawab pertanyaan pengguna HANYA berdasarkan data berikut.
    - Jika diminta total, hitung jumlahnya.
    - Jika diminta rincian, buat daftarnya.
    - Jawaban harus singkat dan langsung ke intinya.
    - Format tanggal sebagai 'DD MMM YYYY' dan mata uang sebagai 'IDR 1.234.500'.
    - JANGAN gunakan markdown.

    ---
    ${dataTypeMessage}:
    ${dataAsText}
    ---
    PERTANYAAN: "${entities.originalQuestion}"
    ---
    JAWABAN:
  `;
  return callGeminiApi_(analysisPrompt, false);
}

/**
 * Menjawab pertanyaan perbandingan.
 */
function handleComparisonQuery_(entities) {
  const extractedFilters = extractEntities_(entities.originalQuestion, entities.walletOwner, entities.context, true); // true for multi-entity extraction
  const people = extractedFilters.person_beneficiary ? extractedFilters.person_beneficiary.split(',').map(p => p.trim()) : [entities.walletOwner];
  let comparisonDataText = "";
  let dataFound = false;

  people.forEach(person => {
    const filters = { ...extractedFilters, person_beneficiary: person };
    const personData = getFilteredData_(filters);
    if (personData.length > 0) {
      dataFound = true;
      comparisonDataText += `--- DATA UNTUK ${person.toUpperCase()} ---\n${formatDataForAI_(personData, false)}\n\n`;
    }
  });

  if (!dataFound) {
    return "Maaf, saya tidak dapat menemukan data yang cukup untuk dibuat perbandingan.";
  }

  const analysisPrompt = `
    Anda adalah analis keuangan. Bandingkan data dari beberapa entitas berikut untuk menjawab pertanyaan pengguna.
    Fokus pada perbedaan total, rata-rata, dan pola yang menonjol.
    - Format tanggal sebagai 'DD MMM YYYY' dan mata uang sebagai 'IDR 1.234.500'.
    - JANGAN gunakan markdown.

    ---
    DATA UNTUK DIBANDINGKAN:
    ${comparisonDataText}
    ---
    PERTANYAAN: "${entities.originalQuestion}"
    ---
    HASIL PERBANDINGAN:
  `;
  return callGeminiApi_(analysisPrompt, false);
}

/**
 * Menjawab pertanyaan umum dan terkait goals.
 */
function handleGeneralAdvisory_(entities) {
  const advisoryPrompt = `Anda adalah penasihat keuangan pribadi. Jawab secara edukatif & memotivasi.
Konteks & aturan:
- Jangan mengarang detail di luar data.
- Jika terkait goal, kaitkan dengan goals di konteks.
- Gunakan bullet points (-) dan bold istilah penting.
---
KONTEKS:
${entities.context}
---
PERTANYAAN: "${entities.originalQuestion}"
---
JAWABAN:`;
  return callGeminiApi_(advisoryPrompt, false);
}

// #endregion

// #region AI - CONTEXT & ENTITY EXTRACTION
// =================================================================
//              AI - CONTEXT & ENTITY EXTRACTION
// =================================================================

/**
 * Mengumpulkan semua konteks dari sheet-sheet setup di spreadsheet aktif.
 * @returns {string} Sebuah blok teks berisi rangkuman konteks.
 */
function getGlobalContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let context = "=== USER'S FINANCIAL CONTEXT ===\n";
  try {
    const wallets = ss.getSheetByName(WALLET_SETUP_SHEET).getDataRange().getValues();
    const walletHeaders = wallets.shift();
    const walletOwnerCol = walletHeaders.indexOf('Wallet Owner');
    const walletOwners = [...new Set(wallets.map(row => row[walletOwnerCol]))].join(', ');
    context += `- Wallet Owners (Payers): ${walletOwners}\n`;

    const goalsSheet = ss.getSheetByName(GOALS_SHEET);
    if (goalsSheet) {
      const goals = goalsSheet.getDataRange().getValues();
      const goalsHeaders = goals.shift();
      const goalNameCol = goalsHeaders.indexOf('Goals');
      const goalOwnerCol = goalsHeaders.indexOf('Goal Owner');
      const goalTargetCol = goalsHeaders.indexOf('Nominal Needed');
      const goalDeadlineCol = goalsHeaders.indexOf('Deadline');
      
      context += "- Active Goals:\n";
      goals.forEach(goal => {
        if (goal[goalNameCol]) {
          context += `  - Name: ${goal[goalNameCol]}, Owner: ${goal[goalOwnerCol]}, Target: ${formatCurrency_(goal[goalTargetCol])}, Deadline: ${formatDateForDisplay_(goal[goalDeadlineCol])}\n`;
        }
      });
    }
  } catch (e) {
    console.error("Failed to load global context: " + e.message);
    return "Context could not be fully loaded.";
  }
  return context;
}


/**
 * Mengekstrak entitas dari pertanyaan pengguna dengan bantuan konteks.
 */
function extractEntities_(question, defaultOwner, context, allowMultiplePeople = false) {
  const today = new Date().toISOString().split('T')[0];
  const extractionPrompt = `
    You are an expert entity extractor. Based on the user's financial context and their question, extract entities into a clean JSON.
    CONTEXT:
    ${context}
    - "saya", "aku", "kita" refers to "${defaultOwner}".
    - Today's date is ${today}.
    JSON STRUCTURE:
    {
      "person_beneficiary": "string or null",
      "category": "string or null",
      "startDate": "YYYY-MM-DD or null",
      "endDate": "YYYY-MM-DD or null"
    }
    RULES:
    - If an entity is not mentioned, its value MUST be null.
    - ${allowMultiplePeople ? 'If multiple beneficiaries are mentioned (e.g., "Bapaknya dan Family"), combine them with a comma: "Bapaknya, Family".' : 'If multiple people are mentioned, just pick the first one.'}
    - ONLY output the JSON.
    ANALYZE: "${question}"
  `;
  const geminiResponse = callGeminiApi_(extractionPrompt, true);
  try {
    const parsed = JSON.parse(geminiResponse.replace(/```json\n|```/g, '').trim());
    if (!parsed.person_beneficiary) {
        parsed.person_beneficiary = defaultOwner;
    }
    return parsed;
  } catch (e) {
    console.error("Failed to parse JSON entities: ", geminiResponse);
    return { person_beneficiary: defaultOwner, category: null, startDate: null, endDate: null };
  }
}

// #endregion

// #region AI - DATA RETRIEVAL & FORMATTING
// =================================================================
//              AI - DATA RETRIEVAL & FORMATTING
// =================================================================

/**
 * Mengambil data transaksi historis dari sheet 'Input' di spreadsheet aktif.
 */
function getFilteredData_(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) throw new Error(`Sheet '${DATA_SHEET}' not found.`);

  const data = sheet.getDataRange().getValues();
  const headers = data.shift().map(h => String(h || '').trim());

  const colMap = {
    date: headers.indexOf("Date"),
    payer: headers.indexOf("Wallet Owner"),
    beneficiary: headers.indexOf("Expense Purpose"),
    category: headers.indexOf("Category"),
    amount: headers.indexOf("Amount"),
    description: headers.indexOf("Description")
  };

  const results = [];
  const beneficiaries = filters.person_beneficiary ? filters.person_beneficiary.toLowerCase().split(',').map(n => n.trim()) : null;
  const startDate = filters.startDate ? new Date(filters.startDate) : null;
  const endDate = filters.endDate ? new Date(filters.endDate) : null;
  if (startDate) startDate.setHours(0, 0, 0, 0);
  if (endDate) endDate.setHours(23, 59, 59, 999);

  data.forEach(row => {
    let isMatch = true;
    if (beneficiaries && !beneficiaries.includes(String(row[colMap.beneficiary] || "").toLowerCase())) isMatch = false;
    if (isMatch && filters.category && String(row[colMap.category] || "").toLowerCase() !== filters.category.toLowerCase()) isMatch = false;
    
    if (isMatch && startDate && endDate) {
      const rowDate = new Date(row[colMap.date]);
      if (isNaN(rowDate.getTime()) || rowDate < startDate || rowDate > endDate) isMatch = false;
    }
    
    if (isMatch) {
      results.push({
        Date: new Date(row[colMap.date]),
        Payer: row[colMap.payer],
        Beneficiary: row[colMap.beneficiary],
        Category: row[colMap.category] || 'N/A',
        Amount: parseFloat(row[colMap.amount]) || 0,
        Description: row[colMap.description] || 'No Desc'
      });
    }
  });
  return results;
}

/**
 * Mengambil data transaksi terjadwal dari sheet 'ScheduledTransactions' di spreadsheet aktif.
 */
function getScheduledData_(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULED_SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift().map(h => String(h || '').trim());
  
  const colMap = {
    date: headers.indexOf("NextDueDate"),
    payer: headers.indexOf("Wallet Owner"),
    beneficiary: headers.indexOf("Expense Purpose"),
    category: headers.indexOf("Category"),
    amount: headers.indexOf("Amount"),
    status: headers.indexOf("Status"),
    description: headers.indexOf("Description")
  };

  const results = [];
  const beneficiaries = filters.person_beneficiary ? filters.person_beneficiary.toLowerCase().split(',').map(n => n.trim()) : null;
  const startDate = filters.startDate ? new Date(filters.startDate) : null;
  const endDate = filters.endDate ? new Date(filters.endDate) : null;
  if (startDate) startDate.setHours(0, 0, 0, 0);
  if (endDate) endDate.setHours(23, 59, 59, 999);

  data.forEach(row => {
    const status = String(row[colMap.status] || "").toLowerCase();
    if (status !== 'active' && status !== 'upcoming') return;

    let isMatch = true;
    if (beneficiaries && !beneficiaries.includes(String(row[colMap.beneficiary] || "").toLowerCase())) isMatch = false;
    if (isMatch && filters.category && String(row[colMap.category] || "").toLowerCase() !== filters.category.toLowerCase()) isMatch = false;
    
    if (isMatch && startDate && endDate) {
      const rowDate = new Date(row[colMap.date]);
      if (isNaN(rowDate.getTime()) || rowDate < startDate || rowDate > endDate) isMatch = false;
    }
    
    if (isMatch) {
      results.push({
        Date: new Date(row[colMap.date]),
        Payer: row[colMap.payer],
        Beneficiary: row[colMap.beneficiary] || 'N/A',
        Category: row[colMap.category] || 'N/A',
        Amount: parseFloat(row[colMap.amount]) || 0,
        Description: row[colMap.description] || 'No Desc'
      });
    }
  });
  return results;
}


/**
 * Utility untuk mengubah data array of objects menjadi teks CSV untuk AI.
 */
function formatDataForAI_(data, isScheduled) {
  const header = isScheduled ? "NextDueDate,Payer,Beneficiary,Category,Description,Amount\n" : "Date,Payer,Beneficiary,Category,Description,Amount\n";
  return header + data.map(d => [ d.Date.toISOString().split('T')[0], d.Payer, d.Beneficiary, d.Category, (d.Description || '').replace(/,/g, ';'), d.Amount ].join(',')).join('\n');
}

// #endregion

// #region AI - API CALL & UTILITIES
// =================================================================
//                 AI - API CALL & UTILITIES
// =================================================================

/**
 * Memanggil Gemini API dengan prompt yang diberikan.
 */
function callGeminiApi_(prompt, isJsonOutput) {
  // Mengambil API Key dari Script Properties untuk keamanan.
  // Cara set: Buka editor script > Project Settings (ikon gerigi) > Script Properties.
  const properties = PropertiesService.getScriptProperties();
  const API_KEY_GEMINI = properties.getProperty('API_KEY_GEMINI');

  if (!API_KEY_GEMINI) {
    return "ERROR: Kunci API Gemini belum diatur di Script Properties.";
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY_GEMINI}`;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { "temperature": 0.2, "maxOutputTokens": 2048 }
  };
  
  if (isJsonOutput) {
    payload.generationConfig.responseMimeType = "application/json";
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let res = UrlFetchApp.fetch(url, options);
  let code = res.getResponseCode();
  let text = res.getContentText();
  if (code === 429) {
    // Single backoff retry
    Utilities.sleep(1200);
    try {
      res = UrlFetchApp.fetch(url, options);
      code = res.getResponseCode();
      text = res.getContentText();
    } catch (e) { /* ignore, will fall through */ }
  }
  if (code === 200) {
    try {
      const data = JSON.parse(text);
      return data.candidates?.[0]?.content?.parts?.[0]?.text || "(No content returned from API)";
    } catch(parseErr) {
      console.error('Parse AI response error', parseErr);
      return '(AI response parse error)';
    }
  }
  if (code === 429) {
    return 'Layanan AI sedang dibatasi (429). Coba ulang dalam 30–60 detik.';
  }
  console.error(`API Error (${code}): ${text}`);
  throw new Error(`Failed to contact AI (Error ${code}).`);
}

/**
 * Helper untuk format mata uang.
 */
function formatCurrency_(number) {
    if (number == null || isNaN(number)) return 'IDR 0';
    // Format as currency, then replace "Rp" with "IDR" for consistency
    return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 })
        .format(number)
        .replace(/^Rp/, 'IDR')
        .replace(/\u00A0/g, ' ');
}

/**
 * Helper untuk format tanggal.
 */
function formatDateForDisplay_(date) {
  if (!date) return 'N/A';
  try {
    return Utilities.formatDate(new Date(date), "GMT+7", "d MMM yyyy");
  } catch (e) {
    return String(date);
  }
}

// #endregion


// ------------------- File: Logger.gs (Fungsi Logging) -------------------

// #region LOGGER
// =================================================================
//                            LOGGER
// =================================================================

/**
 * Mencatat sebuah event ke sheet "Queries" di MAIN_SHEET_ID.
 * @param {string} queryId ID unik untuk query/permintaan.
 * @param {Date} timestamp Waktu log.
 * @param {string} walletOwner Nama pemilik wallet.
 * @param {string} userQuestion Pertanyaan user.
 * @param {string} aiAnswer Jawaban AI.
 */
function logApiUsage(queryId, timestamp, walletOwner, userQuestion, aiAnswer) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const MAIN_SHEET_ID = properties.getProperty('MAIN_SHEET_ID');
    const QUERIES_SHEET_NAME = "Queries";
    const logSpreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const queriesSheet = logSpreadsheet.getSheetByName(QUERIES_SHEET_NAME);
    if (!queriesSheet) throw new Error(`Sheet '${QUERIES_SHEET_NAME}' not found.`);
    queriesSheet.appendRow([queryId, timestamp, walletOwner, userQuestion, aiAnswer]);
  } catch (error) {
    console.error("Failed to log to Queries sheet: " + error.message);
  }
}

// #endregion
