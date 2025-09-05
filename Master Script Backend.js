// ------------------- File: Code.gs (Backend Utama) -------------------
/**
 * @OnlyCurrentDoc
 * File utama yang berfungsi sebagai backend untuk dashboard dan AI.
 * Script ini terikat (bound) ke Google Sheet pelanggan.
 */

// #region CONFIGURATION
// =================================================================
//                        CONFIGURATION
// =================================================================

// Ambang batas penggunaan budget untuk perhitungan status
const BUDGET_OVER_THRESHOLD = 100;
const BUDGET_WARNING_THRESHOLD = 80;

/**
 * Nama-nama sheet yang digunakan oleh script.
 */
const WALLET_SETUP_SHEET = 'Wallet Setup';
const CATEGORY_SHEET = "Category Setup";
const GOALS_SHEET = "Goals Setup";
const SCHEDULED_SHEET_NAME = "ScheduledTransactions";
const DATA_SHEET = 'Input';

// #endregion

// #region WEB APP & MAIN DATA FETCHER
// =================================================================
//                 WEB APP & MAIN DATA FETCHER
// =================================================================

/**
 * Fungsi untuk menyajikan file Dashboard.html sebagai web app.
 * @returns {HtmlOutput} Output HTML yang akan dirender.
 */
function doGet(e) {
  // Handle permintaan untuk Service Worker
  if (e.parameter.sw === '1') {
    // Panggil fungsi yang berisi kode Service Worker
    const swContent = getServiceWorkerJs_(); 
    return ContentService.createTextOutput(swContent)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Handle permintaan untuk manifest
  if (e.parameter.manifest === '1') {
    // Panggil fungsi yang menghasilkan JSON manifest
    const manifestContent = getManifestJson_();
    return ContentService.createTextOutput(manifestContent)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Tampilkan halaman HTML utama
  return HtmlService.createTemplateFromFile('Master Script Frontend').evaluate()
    .setTitle('SatukasSatukas — Financial Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Fungsi ini adalah jembatan utama antara frontend (HTML) dan backend (Apps Script).
 * Frontend akan memanggil fungsi ini untuk mendapatkan semua data yang diperlukan dashboard.
 * @param {string} period Periode yang diminta oleh pengguna (misal: 'current_month').
 * @param {object} filters Objek filter yang diterapkan oleh pengguna.
 * @param {boolean} forceRefresh Jika true, akan mengabaikan cache dan mengambil data baru.
 * @returns {object} Objek berisi semua data yang sudah diproses untuk setiap komponen dashboard.
 * @property {Array<Object>} goalsStatus Status progres setiap tujuan finansial.
 * @property {Array<Object>} netFlow Data arus kas bersih per periode.
 * @property {Array<Object>} budgetStatus Status penggunaan budget.
 * @property {Array<Object>} liabilitiesUpcoming Gabungan data utang dan transaksi mendatang.
 * @property {Array<Object>} ratios Rasio pengeluaran (Living, Playing, Saving).
 * @property {Array<Array>} sankeyData Data untuk Google Sankey Chart, berupa array berisi [owner, purpose, amount].
 * @property {number} totalSaving Total nilai tabungan dari dompet tipe 'Other Asset' dan 'Savings'.
 */
function getDashboardData(period, filters, forceRefresh) {
  try {
    const cache = CacheService.getUserCache();
    const cacheKey = `dashboardData_${period}_${JSON.stringify(filters)}`;
    let cachedData = cache.get(cacheKey);

    if (cachedData != null && !forceRefresh) {
      console.log("Mengambil data dashboard dari cache.");
      return JSON.parse(cachedData);
    }

    // Jika tidak ada periode atau filter, gunakan default
    const safePeriod = period || 'current_month';
    const safeFilters = filters || {};

    // --- Mengambil semua data mentah yang diperlukan sekali saja ---
    // Untuk DATA_SHEET dan SCHEDULED_SHEET_NAME, kita prioritaskan kesegaran data (forceRefresh = true)
    // Untuk sheet setup, kita bisa gunakan cache yang lebih lama
    const allTransactionsData = getRawSheetData_(DATA_SHEET, forceRefresh);
    const scheduledTransactionsData = getRawSheetData_(SCHEDULED_SHEET_NAME, forceRefresh);
    const walletSetupData = getRawSheetData_(WALLET_SETUP_SHEET, false); // Cache untuk Wallet Setup
    const categorySetupData = getRawSheetData_(CATEGORY_SHEET, false); // Cache untuk Category Setup
    const goalsSetupData = getRawSheetData_(GOALS_SHEET, false); // Cache untuk Goals Setup

    const { startDate, endDate } = getPeriodDates_(period, filters.startDate, filters.endDate);
    const { startDate: prevStartDate, endDate: prevEndDate } = getPreviousPeriodDates_(period, startDate);

    const transactions = getFilteredTransactions_(allTransactionsData, filters, startDate, endDate);
    const prevTransactions = getFilteredTransactions_(allTransactionsData, filters, prevStartDate, prevEndDate);

    // Hitung KPI utama (income/expense/net flow)
    const kpiSummary = calculateKpiSummary_(transactions, prevTransactions);

    // Saving (sudah ada)
    const totalSaving = calculateTotalSaving_(transactions);
    const prevTotalSaving = calculateTotalSaving_(prevTransactions);
    kpiSummary.saving = totalSaving;
    kpiSummary.prev_saving = prevTotalSaving;

    // Wallet status (dipakai juga untuk aset terkini tampilan)
  let walletStatus = calculateWalletStatus_(allTransactionsData, walletSetupData);

    // compute liquid assets by summing up from the walletStatus result
    const liquidAssets = (walletStatus || []).reduce((total, wallet) => {
      const type = (wallet.Type || '').toLowerCase();
      // Cek apakah tipe wallet termasuk dalam kategori likuid
      if (type.includes('cash') || type.includes('bank') || type.includes('e-wallet')) {
        return total + (wallet.Balance || 0);
      }
      return total;
    }, 0);


    // assign liquidAssets into KPI summary for frontend convenience
    kpiSummary.liquidAssets = liquidAssets;

    // --- Net Worth Snapshot (BARU) ---
    const currentNetWorthSnapshot = calculateNetWorthSnapshot_(allTransactionsData, endDate); // global snapshot
    let previousNetWorthSnapshot = { assets: 0, liabilities: 0, netWorth: 0 };
    if (prevEndDate && !isNaN(prevEndDate.getTime()) && prevEndDate.getFullYear() > 1970) {
      previousNetWorthSnapshot = calculateNetWorthSnapshot_(allTransactionsData, prevEndDate); // global prev snapshot
    }
    kpiSummary.netWorth = currentNetWorthSnapshot.netWorth;
    kpiSummary.prev_netWorth = previousNetWorthSnapshot.netWorth;

  // Flag filtered state (currently only walletOwner drives Net Worth override)
  kpiSummary.isFiltered = !!safeFilters.walletOwner;

  // === Override Net Worth & Wallet Status when walletOwner filter applied ===
  if (safeFilters.walletOwner) {
      try {
        const ownerFilteredCurrent = calculateNetWorthSnapshot_(allTransactionsData, endDate, safeFilters.walletOwner);
        let ownerFilteredPrev = { assets: 0, liabilities: 0, netWorth: 0 };
        if (prevEndDate && !isNaN(prevEndDate.getTime()) && prevEndDate.getFullYear() > 1970) {
          ownerFilteredPrev = calculateNetWorthSnapshot_(allTransactionsData, prevEndDate, safeFilters.walletOwner);
        }
        kpiSummary.netWorth = ownerFilteredCurrent.netWorth;
        kpiSummary.prev_netWorth = ownerFilteredPrev.netWorth;
        // Filter walletStatus list itself to reflect owner scope
        walletStatus = (walletStatus || []).filter(w => (w.Owner || '') === safeFilters.walletOwner);
      } catch(e) {
        console.warn('Owner-filtered net worth calculation failed:', e && e.message);
      }
    }

    // Komponen lain
    const goalsStatus = calculateGoalsStatus_(goalsSetupData, transactions, filters);
    const netFlow = calculateNetFlow_(allTransactionsData, safePeriod, safeFilters);
    const budgetStatus = calculateBudgetStatus_(categorySetupData, allTransactionsData, safePeriod, safeFilters);
    const liabilitiesUpcoming = calculateLiabilitiesUpcoming_(scheduledTransactionsData, allTransactionsData, safePeriod, safeFilters);
    const ratios = calculateRatios_(categorySetupData, allTransactionsData, safePeriod, safeFilters);
    const sankeyData = calculateSankeyData_(allTransactionsData, safePeriod, safeFilters);
    const expenseTreeMap = calculateExpenseTreeMapData_(transactions, prevTransactions);

    const dashboardData = {
      kpiSummary,
      goalsStatus,
      netFlow,
      budgetStatus,
      liabilitiesUpcoming,
      ratios,
      sankeyData,
      totalSaving,
      expenseTreeMap,
      walletStatus,
      liquidAssets
    };

    // --- NEW: build snapshot cards (compact placeholders) ---
    try {
      dashboardData.snapshotCards = computeFinancialSnapshot_(dashboardData, allTransactionsData, walletSetupData, goalsSetupData);
    } catch (e) {
      console.warn('computeFinancialSnapshot_ failed:', e && e.message);
      dashboardData.snapshotCards = {};
    }

    // --- Financial Insights (Major Spent & Big Change + Moving Average) ---
    try {
      const cats = (expenseTreeMap && expenseTreeMap.byCategory) || [];
      let majorSpent = null, bigChange = null;
      let majorSpentPct = 0, bigChangePct = 0, bigChangeType = '';
      if (cats.length) {
        // Major Spent
        const total = expenseTreeMap.total || cats.reduce((s, c) => s + (c.value || 0), 0);
        majorSpent = cats.reduce((a, b) => (b.value || 0) > (a.value || 0) ? b : a, cats[0]);
        majorSpentPct = total > 0 ? ((majorSpent.value || 0) / total * 100) : 0;

        // Big Change (deviasi terbesar dari rata2 prev_value)
        const avgPrev = cats.reduce((s, c) => s + (c.prev_value || 0), 0) / Math.max(1, cats.filter(c => c.prev_value > 0).length);
        let biggest = null, biggestDelta = 0;
        cats.forEach(c => {
          const delta = c.value - (c.prev_value || avgPrev);
          if (Math.abs(delta) > Math.abs(biggestDelta)) {
            biggest = c;
            biggestDelta = delta;
          }
        });
        if (biggest && Math.abs(biggestDelta) > 0.3 * (avgPrev || 1)) {
          bigChange = biggest;
          bigChangeType = biggestDelta > 0 ? 'increase' : 'decrease';
          bigChangePct = (biggest.prev_value && biggest.prev_value > 0)
            ? ((biggest.value - biggest.prev_value) / biggest.prev_value * 100)
            : null;
        }
      }
      dashboardData.financialInsights = {
        majorSpent: majorSpent ? { name: majorSpent.name, value: majorSpent.value, pct: majorSpentPct } : null,
        bigChange: bigChange ? { name: bigChange.name, type: bigChangeType, pct: bigChangePct } : null
      };

      // --- NEW: Moving Average untuk Spending Behavior ---
      const currExp = kpiSummary.expense;
      const prevExp = kpiSummary.prev_expense;
      const changePct = prevExp > 0
        ? parseFloat(((currExp - prevExp) / prevExp * 100).toFixed(1))
        : null;
      dashboardData.financialInsights.movingAverage = {
        current: currExp,
        previous: prevExp,
        changePct: changePct
      };

    } catch(e) {
      dashboardData.financialInsights = {};
    }

    // Perbarui cache dengan data terbaru (hanya jika ukuran kecil)
    putCacheIfSmall_(cache, cacheKey, dashboardData, 300); // Cache selama 5 menit (300 detik)
    console.log(forceRefresh ? "Force refresh: Data dashboard diperbarui dan disimpan ke cache." : "Data dashboard disimpan ke cache.");
    return dashboardData;
  } catch (error) {
    console.error("Error di getDashboardData: " + error.stack);
    // Mengirim pesan error yang lebih informatif ke frontend
    throw new Error("Gagal mengambil data dashboard. Penyebab: " + error.message);
  }
}

/**
 * Mengambil opsi filter unik dari data untuk mengisi dropdown di frontend.
 * @param {boolean} forceRefresh Jika true, akan mengabaikan cache dan mengambil opsi baru.
 * @returns {object} Objek yang berisi array unik untuk setiap filter.
 */
function getFilterOptions(forceRefresh) {
    const cache = CacheService.getUserCache();
    const cacheKey = "filterOptions";
    let cachedOptions = cache.get(cacheKey);

    if (cachedOptions != null && !forceRefresh) {
      console.log("Mengambil opsi filter dari cache.");
      return JSON.parse(cachedOptions);
    }

    // Mengambil data mentah dari sheet yang relevan
    const walletData = getRawSheetData_(WALLET_SETUP_SHEET, forceRefresh);
    const inputData = getRawSheetData_(DATA_SHEET, forceRefresh);
    const categoryData = getRawSheetData_(CATEGORY_SHEET, forceRefresh);

    if (!walletData || !inputData) {
        console.error("Sheet 'Wallet Setup' atau 'Input' tidak ditemukan.");
        return { wallets: [], walletOwners: [], expensePurposes: [] };
    }

    // Memproses data 'Wallet Setup'
    const walletHeaders = walletData.shift();
    const walletColIndex = walletHeaders.indexOf('Wallet');
    const ownerColIndex = walletHeaders.indexOf('Wallet Owner');
    
    const wallets = [...new Set(walletData.map(row => row[walletColIndex]).filter(Boolean))];
    const walletOwners = [...new Set(walletData.map(row => row[ownerColIndex]).filter(Boolean))];

    // Memproses data 'Input'
    const inputHeaders = inputData.shift();
    const purposeColIndex = inputHeaders.indexOf('Expense Purpose');
    const noteColIndex = inputHeaders.indexOf('Note');
    const expensePurposes = [...new Set(inputData.map(row => row[purposeColIndex]).filter(Boolean))];
    const notes = noteColIndex >= 0 ? [...new Set(inputData.map(row => (row[noteColIndex]!==null && row[noteColIndex]!==undefined ? String(row[noteColIndex]).trim() : '')).filter(v=>v))] : [];

    // Memproses data 'Category Setup' untuk daftar Category/Subcategory
    let categories = [];
    let subcategories = [];
    try {
      const catHeaders = categoryData.shift();
      const catIdx = catHeaders.indexOf('Category');
      const subIdx = catHeaders.indexOf('Subcategory');
      categories = [...new Set(categoryData.map(r => r[catIdx]).filter(Boolean))];
      subcategories = [...new Set(categoryData.map(r => r[subIdx]).filter(Boolean))];
    } catch (e) {
      // fallback dari inputData bila category setup tidak ada
      const catIdx = inputHeaders.indexOf('Category');
      const subIdx = inputHeaders.indexOf('Subcategory');
      if (catIdx >= 0) categories = [...new Set(inputData.map(r => r[catIdx]).filter(Boolean))];
      if (subIdx >= 0) subcategories = [...new Set(inputData.map(r => r[subIdx]).filter(Boolean))];
    }
    
    const options = {
        wallets,
        walletOwners,
        expensePurposes,
        categories,
        subcategories,
        notes
    };

    // Perbarui cache dengan opsi terbaru (hanya jika ukuran kecil)
    putCacheIfSmall_(cache, cacheKey, options, 3600); // Cache selama 1 jam (3600 detik)
    console.log(forceRefresh ? "Force refresh: Opsi filter diperbarui dan disimpan ke cache." : "Opsi filter disimpan ke cache.");
    return options;
}


// #endregion

// #region DASHBOARD CALCULATION FUNCTIONS
// =================================================================
//               DASHBOARD CALCULATION FUNCTIONS
// =================================================================

/**
 * Menghitung ringkasan KPI (Income, Expense, Net) dari data transaksi yang sudah difilter.
 * @param {Array<Object>} currentTransactions Transaksi untuk periode saat ini.
 * @param {Array<Object>} previousTransactions Transaksi untuk periode sebelumnya.
 * @returns {Object} Objek berisi ringkasan KPI.
 */
function calculateKpiSummary_(currentTransactions, previousTransactions) {
  // --- TAMBAHAN: Definisikan ulang regex untuk deteksi tabungan tersamarkan ---
  // Ini perlu agar kita bisa memisahkan expense murni dari 'disguised saving'.
  const disguisedSavingRegex = new RegExp('(^|[^a-z])(tabungan|menabung|saving|savings?|autosave|investment|investasi|deposito?|reksadana|mutualfund|saham|stock|equity|obligasi|bond|pensiun|retirement|emergencyfund|aset|asset|capital)([^a-z]|$)', 'i');

  let currentIncome = 0, currentExpense = 0;
  (currentTransactions || []).forEach(t => {
    // Menggunakan Amount yang sudah dinormalisasi (+/-) dari getFilteredTransactions_
    if (t.Amount > 0) {
      currentIncome += t.Amount;
    } else { // Amount < 0 (expense atau disguised saving)
      const category = (t.Category || '').toString().trim();
      const subcategory = (t.Subcategory || '').toString().trim();
      const type = (t.Type || '').toString().trim().toLowerCase();

      // Cek apakah ini 'disguised saving'
      const isDisguisedSaving = type === 'expense' && (disguisedSavingRegex.test(category) || disguisedSavingRegex.test(subcategory));

      // Hanya tambahkan ke expense jika BUKAN disguised saving
      if (!isDisguisedSaving) {
        currentExpense += -t.Amount; // Expense disimpan sebagai nilai positif
      }
    }
  });

  let prevIncome = 0, prevExpense = 0;
  (previousTransactions || []).forEach(t => {
    if (t.Amount > 0) {
      prevIncome += t.Amount;
    } else { // Amount < 0
      const category = (t.Category || '').toString().trim();
      const subcategory = (t.Subcategory || '').toString().trim();
      const type = (t.Type || '').toString().trim().toLowerCase();
      const isDisguisedSaving = type === 'expense' && (disguisedSavingRegex.test(category) || disguisedSavingRegex.test(subcategory));
      if (!isDisguisedSaving) {
        prevExpense += -t.Amount;
      }
    }
  });

  return {
    income: currentIncome,
    expense: currentExpense,
    net: currentIncome - currentExpense,        // Net Flow (tetap)
    prev_income: prevIncome,
    prev_expense: prevExpense,
    prev_net: prevIncome - prevExpense,
    saving: 0,
    prev_saving: 0,
    netWorth: 0,          // BARU
    prev_netWorth: 0      // BARU
  };
}


/**
 * Menghitung progres setiap tujuan finansial.
 * @param {Array<Array>} goalsData Data mentah dari sheet Goals Setup.
 * @param {Array<Object>} filteredTransactions Data transaksi yang sudah difilter.
 * @param {object} filters Filter tambahan.
 * @returns {Array<Object>} Array objek yang berisi data status untuk setiap goal.
 */
function calculateGoalsStatus_(goalsData, filteredTransactions, filters) {
  if (!goalsData || !filteredTransactions) throw new Error('Data sheet Goals Setup atau transaksi tidak tersedia.');

  const currentGoalsData = [...goalsData];
  const goalsHeaders = currentGoalsData.shift();
  const goalNameCol = goalsHeaders.indexOf('Goals');
  const goalOwnerCol = goalsHeaders.indexOf('Goal Owner');
  const nominalNeededCol = goalsHeaders.indexOf('Nominal Needed');
  const deadlineCol = goalsHeaders.indexOf('Deadline');

  // Ambil semua transaksi mentah (ALL) untuk mencari earliest contribution (T0)
  let allTxRaw = [];
  try { allTxRaw = getRawSheetData_(DATA_SHEET, false); } catch(e) { /* ignore */ }
  const allTxHeaders = allTxRaw.length ? allTxRaw[0] : [];
  const idxDate = allTxHeaders.indexOf('Date');
  const idxType = allTxHeaders.indexOf('Transaction Type');
  const idxAmount = allTxHeaders.indexOf('Amount');
  const idxCategory = allTxHeaders.indexOf('Category');
  const idxSubcat = allTxHeaders.indexOf('Subcategory');
  const idxPurpose = allTxHeaders.indexOf('Expense Purpose');

  const today = new Date(); today.setHours(0,0,0,0);

  const materializedData = [];

  currentGoalsData.forEach(goal => {
    const goalName = goal[goalNameCol];
    if (!goalName) return;
    const goalOwner = goal[goalOwnerCol];
    const totalNeeded = parseFloat(goal[nominalNeededCol]) || 0;
    const deadline = goal[deadlineCol] ? new Date(goal[deadlineCol]) : null;
    if (deadline && !isNaN(deadline.getTime())) deadline.setHours(0,0,0,0);

    // 1. Hitung collected (dari filteredTransactions = periode aktif) sesuai kriteria existing
    const collected = filteredTransactions
      .filter(t => t.Purpose === goalOwner && t.Category === 'Saving/Investment' && t.Subcategory === goalName && t.Amount > 0)
      .reduce((s,t)=> s + t.Amount, 0);

    // 2. Earliest contribution (T0) scanning seluruh transaksi historis
    let earliest = null;
    if (allTxRaw.length > 1) {
      for (let r = 1; r < allTxRaw.length; r++) {
        const row = allTxRaw[r];
        const cat = row[idxCategory];
        const sub = row[idxSubcat];
        if (cat !== 'Saving/Investment' || sub !== goalName) continue;
        const purpose = idxPurpose >=0 ? row[idxPurpose] : '';
        if (purpose !== goalOwner) continue; // konsisten dengan logika existing
        let amt = normalizeNumber_(row[idxAmount]);
        const tType = (row[idxType] || '').toString().toLowerCase();
        const subLower = (row[idxSubcat] || '').toString().toLowerCase();
        // Positive contribution criteria
        if (tType === 'income' || (tType === 'transfer' && subLower === 'transfer-in') || amt > 0) {
          if (amt <= 0) continue; // pastikan positif setelah normalisasi
          const dRaw = row[idxDate];
            const d = new Date(dRaw);
            if (isNaN(d.getTime())) continue;
            d.setHours(0,0,0,0);
            if (!earliest || d < earliest) earliest = d;
        }
      }
    }
    // Fallback jika belum ada kontribusi
    if (!earliest) { earliest = new Date(); earliest.setHours(0,0,0,0); }

    // 3. Pacing metrics
    const periodSpanMs = (deadline && !isNaN(deadline.getTime())) ? (deadline.getTime() - earliest.getTime()) : 0;
    const elapsedMs = Math.max(0, today.getTime() - earliest.getTime());
    let elapsedRatio = 0;
    if (periodSpanMs > 0) elapsedRatio = Math.min(1, Math.max(0, elapsedMs / periodSpanMs));

    const targetCumulative = totalNeeded * elapsedRatio;
    const gapAmount = collected - targetCumulative;
    const gapPct = totalNeeded > 0 ? (gapAmount / totalNeeded) : 0; // bisa negatif
    const remainingAmount = Math.max(0, totalNeeded - collected);
    const daysLeft = (deadline && !isNaN(deadline.getTime())) ? Math.round((deadline.getTime() - today.getTime())/86400000) : null;
    const elapsedDays = Math.max(1, Math.round(elapsedMs/86400000));
    const paceNeededPerDay = (daysLeft !== null && daysLeft > 0) ? (remainingAmount / daysLeft) : (remainingAmount > 0 ? remainingAmount : 0);
    const actualPacePerDay = collected > 0 ? (collected / elapsedDays) : 0;
    const projectedFinish = (actualPacePerDay > 0 && remainingAmount > 0) ? new Date(today.getTime() + (remainingAmount/actualPacePerDay)*86400000) : (collected >= totalNeeded ? today : null);

    // 4. Status determination
    const pctAchieved = totalNeeded > 0 ? (collected / totalNeeded) : (collected > 0 ? 1 : 0);
    let status = 'On Track';
    if (totalNeeded === 0) {
      status = collected > 0 ? 'Completed' : 'On Track';
    } else if (pctAchieved >= 1.1) {
      status = 'Overfunded';
    } else if (pctAchieved >= 1.0) {
      status = 'Completed';
    } else if (deadline && today > deadline) {
      // Deadline terlewati & belum 100%
      if (pctAchieved >= 0.95) status = 'Completed';
      else if (pctAchieved >= 0.80) status = 'Overdue';
      else status = 'Failed';
    } else if (collected === 0 && elapsedRatio > 0.25) {
      status = 'No Activity';
    } else {
      // GapPct thresholds
      if (gapPct >= 0.05) status = 'Ahead';
      else if (gapPct > -0.05) status = 'On Track';
      else if (gapPct > -0.15) status = 'Slightly Behind';
      else if (gapPct > -0.30) status = 'At Risk';
      else status = 'Off Track';
    }

    // 5. Risk score (0–100) – lebih tinggi = lebih berisiko
    let riskScore = 0;
    if (totalNeeded > 0) {
      const deficitRatio = 1 - pctAchieved; // 0 (aman) .. 1 (belum mulai)
      const timeBuffer = 1 - elapsedRatio;  // 1 (baru mulai) .. 0 (hampir deadline)
      const paceRatio = (paceNeededPerDay > 0) ? (paceNeededPerDay / (actualPacePerDay || paceNeededPerDay)) : 0; // >=1 berarti butuh pace >= actual
      riskScore = (deficitRatio * 60) + (timeBuffer * 20) + (paceRatio * 20);
      riskScore = Math.max(0, Math.min(100, Math.round(riskScore)));
    }

    const progressPercentage = totalNeeded > 0 ? (collected / totalNeeded) * 100 : (collected>0?100:0);

    materializedData.push({
      UniqueID: Utilities.getUuid(),
      GoalName: goalName,
      Deadline: deadline ? formatDateForDisplay_(deadline) : 'N/A',
      StartDate: earliest ? formatDateForDisplay_(earliest) : 'N/A',
      ProgressPercentage: parseFloat(progressPercentage.toFixed(1)),
      RemainingAmount: remainingAmount,
      Collected: collected,
      TotalNeeded: totalNeeded,
      TargetCumulative: parseFloat(targetCumulative.toFixed(2)),
      GapAmount: parseFloat(gapAmount.toFixed(2)),
      GapPct: parseFloat((gapPct*100).toFixed(2)),
      ElapsedRatio: parseFloat((elapsedRatio*100).toFixed(1)),
      PaceNeededPerDay: parseFloat(paceNeededPerDay.toFixed(2)),
      ActualPacePerDay: parseFloat(actualPacePerDay.toFixed(2)),
      DaysLeft: daysLeft,
      ProjectedFinish: projectedFinish ? formatDateForDisplay_(projectedFinish) : '',
      RiskScore: riskScore,
      Status: status
    });
  });

  return materializedData;
}

/**
 * Menghitung arus kas bersih (Pemasukan - Pengeluaran) per periode.
 * @param {Array<Array>} allTransactionsData Data mentah dari sheet Input.
 * @param {string} period Periode waktu yang dipilih.
 * @param {object} filters Filter tambahan.
 * @returns {Array<Object>}
 */
function calculateNetFlow_(allTransactionsData, period, filters) {
  const { startDate, endDate } = getPeriodDates_(period);
  
  if (!allTransactionsData) throw new Error(`Data sheet Input tidak tersedia.`);
  
  let transactionsData = getFilteredTransactions_(allTransactionsData, filters, startDate, endDate);

  const netFlowByPeriod = {};

  transactionsData.forEach(t => {
      const date = new Date(t.Date);
      const yearMonth = Utilities.formatDate(date, "GMT+7", "yyyy-MM");
      
      if (!netFlowByPeriod[yearMonth]) {
        netFlowByPeriod[yearMonth] = { Income: 0, Expense: 0 };
      }
      
      // Menggunakan Amount yang sudah signed (+/-)
      if (t.Amount > 0) {
        netFlowByPeriod[yearMonth].Income += t.Amount;
      } else {
        netFlowByPeriod[yearMonth].Expense += -t.Amount; // Expense disimpan sebagai nilai positif
      }
  });

  const materializedData = Object.keys(netFlowByPeriod).map(key => {
    const data = netFlowByPeriod[key];
    return {
      UniqueID: Utilities.getUuid(),
      PeriodLabel: key,
      Income: data.Income,
      Expense: data.Expense,
      NetFlowAmount: data.Income - data.Expense,
    };
  });

  return materializedData.sort((a, b) => b.PeriodLabel.localeCompare(a.PeriodLabel));
}

/**
 * Menghitung status budget.
 * @param {Array<Array>} categoryData Data mentah dari sheet Category Setup.
 * @param {Array<Array>} allTransactionsData Data mentah dari sheet Input.
 * @param {string} period Periode waktu.
 * @param {object} filters Filter tambahan.
 * @returns {Array<Object>} Data Budget Status.
 */
function calculateBudgetStatus_(categoryData, allTransactionsData, period, filters) {
  const { startDate, endDate } = getPeriodDates_(period);
  
  if (!categoryData || !allTransactionsData) throw new Error(`Data sheet Category Setup atau Input tidak tersedia.`);

  // Buat salinan data karena shift() akan memodifikasi array asli
  const currentCategoriesData = [...categoryData];

  const transactions = getFilteredTransactions_(allTransactionsData, filters, startDate, endDate);
  
  const catHeaders = currentCategoriesData.shift();
  
  const budgetTree = {};

  // 1. Bangun struktur budget dari 'Category Setup'
  currentCategoriesData.forEach(cat => {
    const categoryName = cat[catHeaders.indexOf('Category')];
    const subcategoryName = cat[catHeaders.indexOf('Subcategory')];
    const budgetAmount = parseFloat(cat[catHeaders.indexOf('Budget Subcategory')] || 0);

    if (budgetAmount > 0) {
      if (!budgetTree[categoryName]) {
        budgetTree[categoryName] = { BudgetAmount: 0, ActualExpense: 0, subcategories: {} };
      }
      budgetTree[categoryName].BudgetAmount += budgetAmount;
      budgetTree[categoryName].subcategories[subcategoryName] = { BudgetAmount: budgetAmount, ActualExpense: 0 };
    }
  });

  // 2. Akumulasi pengeluaran dari transaksi yang sudah difilter
  transactions.forEach(t => {
    // Pengeluaran adalah transaksi dengan Amount negatif
    if (t.Amount < 0) {
      if (budgetTree[t.Category]) {
        const expenseAmount = -t.Amount; // Gunakan nilai absolut
        budgetTree[t.Category].ActualExpense += expenseAmount;
        if (budgetTree[t.Category].subcategories[t.Subcategory]) {
          budgetTree[t.Category].subcategories[t.Subcategory].ActualExpense += expenseAmount;
        }
      }
    }
  });

  // 3. Format output final
  const finalBudgetStatus = [];
  Object.keys(budgetTree).forEach(categoryName => {
    const categoryData = budgetTree[categoryName];
    // Hanya tampilkan jika ada budget
    if (categoryData.BudgetAmount > 0) {
        finalBudgetStatus.push(formatBudgetRow_(categoryName, 'All', categoryData.BudgetAmount, categoryData.ActualExpense));
        Object.keys(categoryData.subcategories).forEach(subcategoryName => {
            const subcatData = categoryData.subcategories[subcategoryName];
            finalBudgetStatus.push(formatBudgetRow_(categoryName, subcategoryName, subcatData.BudgetAmount, subcatData.ActualExpense));
        });
    }
  });

  return finalBudgetStatus;
}


/**
 * Menghitung utang dan transaksi mendatang.
 * @param {Array<Array>} scheduledData Data mentah dari sheet ScheduledTransactions.
 * @param {Array<Array>} walletData Data mentah dari sheet Wallet Setup.
 * @param {string} period Periode waktu.
 * @param {object} filters Filter tambahan.
 * @returns {Array<Object>} Gabungan data Liabilities & Upcoming.
 */
function calculateLiabilitiesUpcoming_(scheduledData, allTransactionsData, period, filters) {
  const { startDate, endDate } = getPeriodDates_(period);
  if (!scheduledData || !allTransactionsData) throw new Error('Data ScheduledTransactions atau Input tidak tersedia.');

  const out = [];

  // --- Upcoming (ScheduledTransactions) ---
  const schedCopy = [...scheduledData];
  const schedHeaders = schedCopy.shift() || [];
  const idxSchedStatus   = ciIndex_(schedHeaders,'Status');
  const idxSchedNextDue  = ciIndex_(schedHeaders,'NextDueDate');
  const idxSchedDesc     = ciIndex_(schedHeaders,'Description');
  const idxSchedCat      = ciIndex_(schedHeaders,'Category');
  const idxSchedAmount   = ciIndex_(schedHeaders,'Amount');
  const idxSchedWallet   = ciIndex_(schedHeaders,'Wallet');
  const idxSchedOwner    = ciIndex_(schedHeaders,'Wallet Owner');

  schedCopy.forEach(row => {
    const statusRaw = normStr_(row[idxSchedStatus]);
    if (statusRaw !== 'active') return;
    const rawDateStr = row[idxSchedNextDue];
    if (!rawDateStr) return;
    const due = new Date(rawDateStr);
    if (isNaN(due.getTime())) return;
    if (due < startDate || due > endDate) return;

    const iso = Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const display = formatDateForDisplay_(due);
    const today = new Date(); today.setHours(0,0,0,0);

    const ownerVal = idxSchedOwner !== -1 ? row[idxSchedOwner] : '';
    if (filters.walletOwner && ownerVal !== filters.walletOwner) return; // skip if owner filter not match
    out.push({
      UniqueID: Utilities.getUuid(),
      Type: 'Upcoming',
      Name: row[idxSchedDesc] || row[idxSchedCat],
      Amount: normalizeNumber_(row[idxSchedAmount]),
      Wallet: row[idxSchedWallet],
      Owner: ownerVal,
      RawDueDate: iso,
      DisplayDate: display,
      DueDate: display,
      isOverdue: due < today
    });
  });

  // --- Liabilities (Input) ---
  const inputCopy = [...allTransactionsData];
  const inputHeaders = inputCopy.shift() || [];
  const idxSource  = ciIndex_(inputHeaders,'Source');
  const idxDate    = ciIndex_(inputHeaders,'Date');
  const idxAmount  = ciIndex_(inputHeaders,'Amount');
  const idxDesc    = ciIndex_(inputHeaders,'Description');
  const idxCat     = ciIndex_(inputHeaders,'Category');
  const idxSubcat  = ciIndex_(inputHeaders,'Subcategory');
  const idxWallet  = ciIndex_(inputHeaders,'Wallet');
  const idxOwner   = ciIndex_(inputHeaders,'Wallet Owner');

  // Keyword set (EN + ID) to classify liabilities (case-insensitive)
  const LIAB_KEYWORDS = ['liability','liabilities','debt','loan','credit','installment','repayment','mortgage','hutang','utang','pinjaman','cicilan','kredit','angsuran'];
  const hasLiabKeyword = v => {
    if (!v) return false; const s = String(v).toLowerCase();
    return LIAB_KEYWORDS.some(k => s.includes(k));
  };

  inputCopy.forEach(tx => {
    const sourceVal = idxSource !== -1 ? tx[idxSource] : '';
    const catVal = idxCat !== -1 ? tx[idxCat] : '';
    const subcatVal = idxSubcat !== -1 ? tx[idxSubcat] : '';
    if (!(hasLiabKeyword(sourceVal) || hasLiabKeyword(catVal) || hasLiabKeyword(subcatVal))) return;

    // Transaction date (may be used to assume due date end-of-month)
    let displayDate = ''; let txDate = null;
    if (idxDate !== -1) {
      const d = new Date(tx[idxDate]);
      if (!isNaN(d.getTime())) { txDate = d; displayDate = formatDateForDisplay_(d); }
    }
    // Jika ada tanggal transaksi, hormati filter periode; jika tidak ada tanggal, abaikan (tetap masuk)
    if (txDate) {
      if (txDate < startDate || txDate > endDate) return; // luar periode
    }
    const rawAmt = normalizeNumber_(tx[idxAmount]);
    const name = tx[idxDesc] || tx[idxCat] || 'Liability';
    const ownerVal = idxOwner !== -1 ? tx[idxOwner] : '';
    if (filters.walletOwner && ownerVal !== filters.walletOwner) return;

    // Assumed due date = last day of same month (if txDate exists)
    let dueIso = ''; let dueDisplay = ''; let isOverdue = false;
    if (txDate) {
      const due = new Date(txDate.getFullYear(), txDate.getMonth()+1, 0);
      due.setHours(0,0,0,0);
      dueIso = Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      dueDisplay = formatDateForDisplay_(due);
      const todayMid = new Date(); todayMid.setHours(0,0,0,0);
      isOverdue = due < todayMid;
    }

    // Fallback: jika tidak ada txDate (atau gagal parse) maka set due ke akhir periode (endDate)
    if (!dueIso) {
      try {
        const fallbackDue = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
        fallbackDue.setHours(0,0,0,0);
        dueIso = Utilities.formatDate(fallbackDue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        dueDisplay = formatDateForDisplay_(fallbackDue);
        const todayMid = new Date(); todayMid.setHours(0,0,0,0);
        isOverdue = fallbackDue < todayMid;
      } catch(e) {
        // diamkan; biarkan kosong jika benar-benar gagal
      }
    }

    out.push({
      UniqueID: Utilities.getUuid(),
      Type: 'Liabilities',
      Name: name,
      Amount: Math.abs(rawAmt),
      Wallet: idxWallet !== -1 ? tx[idxWallet] : '',
      Owner: ownerVal,
      DisplayDate: displayDate,
      DueDate: dueDisplay,
      RawDueDate: dueIso,
      isOverdue
    });
  });

  // Debug ringan untuk memastikan liabilities punya RawDueDate
  try {
    const liabCount = out.filter(o => o.Type === 'Liabilities').length;
    const liabNoDue = out.filter(o => o.Type === 'Liabilities' && !o.RawDueDate).length;
    console.log(`[LiabilitiesUpcoming] Liabilities=${liabCount} tanpaDue=${liabNoDue}`);
  } catch(e) {}

  return out;
}

/**
 * Menghitung rasio pengeluaran (Living, Playing, Saving) dengan breakdown Source.
 * @param {Array<Array>} categoryData Data mentah dari sheet Category Setup.
 * @param {Array<Array>} allTransactionsData Data mentah dari sheet Input.
 * @param {string} period Periode waktu.
 * @param {object} filters Filter tambahan.
 * @returns {Array<Object>} Array objek yang berisi data rasio pengeluaran + breakdown source.
 */
function calculateRatios_(categoryData, allTransactionsData, period, filters) {
  const { startDate, endDate } = getPeriodDates_(period);

  if (!categoryData || !allTransactionsData) throw new Error(`Data sheet Category Setup atau Input tidak tersedia.`);

  // Buat salinan data karena shift() akan memodifikasi array asli
  const currentCategoriesData = [...categoryData];

  const transactions = getFilteredTransactions_(allTransactionsData, filters, startDate, endDate);

  const catHeaders = currentCategoriesData.shift();

  const ratioMapping = {};
  currentCategoriesData.forEach(cat => {
    const ratio = cat[catHeaders.indexOf('Ratios')];
    if (ratio) {
      const key = `${cat[catHeaders.indexOf('Category')]}-${cat[catHeaders.indexOf('Subcategory')]}`;
      ratioMapping[key] = ratio;
    }
  });

  // PATCH: breakdown by actual Source value (not just Liabilities/CashBank)
  const expenseByRatio = {};
  transactions.forEach(t => {
    if (t.Amount < 0) {
      const key = `${t.Category}-${t.Subcategory}`;
      const ratioType = ratioMapping[key] || 'Uncategorized';
      const source = (t.Source || 'Unknown').toString().trim() || 'Unknown';

      if (!expenseByRatio[ratioType]) {
        expenseByRatio[ratioType] = { total: 0, bySource: {} };
      }
      expenseByRatio[ratioType].total += -t.Amount;
      if (!expenseByRatio[ratioType].bySource[source]) expenseByRatio[ratioType].bySource[source] = 0;
      expenseByRatio[ratioType].bySource[source] += -t.Amount;
    }
  });

  return Object.keys(expenseByRatio).map(ratioType => ({
    RatioType: ratioType,
    TotalExpense: expenseByRatio[ratioType].total,
    BySource: expenseByRatio[ratioType].bySource
  }));
}

/**
 * Menghitung data untuk Sankey Chart.
 * @param {Array<Array>} allTransactionsData Data mentah dari sheet Input.
 * @param {string} period Periode waktu.
 * @param {object} filters Filter tambahan.
 * @returns {Array<Array>} Data yang diformat untuk Google Sankey Chart.
 */
function calculateSankeyData_(allTransactionsData, period, filters) {
  try {
    if (!allTransactionsData || allTransactionsData.length < 2) return [];

    const { startDate, endDate } = getPeriodDates_(period);
    const tx = getFilteredTransactions_(allTransactionsData, filters || {}, startDate, endDate);

    const PAYER_PREFIX = 'Payer: ';
    const BENEFICIARY_PREFIX = 'Beneficiary: ';

    // Aggregate by raw names (keep raw to detect collisions)
    const aggregate = {}; // payerRaw -> beneficiaryRaw -> total
    const ownerSet = new Set();
    const purposeSet = new Set();
    let inspected = 0, expenseCount = 0;

    tx.forEach(t => {
      inspected++;
      if (t.Amount >= 0) return; // hanya expense
      expenseCount++;

      const payerRaw = (t.Owner || t.Wallet || 'Unknown').toString().trim() || 'Unknown';
      const beneRaw = (t.Purpose || 'Unspecified').toString().trim() || 'Unspecified';

      ownerSet.add(payerRaw);
      purposeSet.add(beneRaw);

      if (!aggregate[payerRaw]) aggregate[payerRaw] = {};
      if (!aggregate[payerRaw][beneRaw]) aggregate[payerRaw][beneRaw] = 0;
      aggregate[payerRaw][beneRaw] += Math.abs(t.Amount);
    });

    // detect names that appear on both sides (for logging only)
    const collisions = [...purposeSet].filter(p => ownerSet.has(p));
    if (collisions.length) {
      console.log('[Sankey] Name collisions (same raw name on both sides):', collisions.join(', '));
    }

    const sankeyRows = [['From','To','Amount']];
    Object.keys(aggregate).forEach(payerRaw => {
      Object.keys(aggregate[payerRaw]).forEach(beneRaw => {
        const amount = aggregate[payerRaw][beneRaw];
        if (!amount) return;
        // Force bipartite by prefixing — this prevents any cycle caused by identical raw strings
        const fromLabel = payerRaw;
        const toLabel = beneRaw;
        // double-check self-loop (shouldn't happen because of prefix) and skip if amount <= 0
        if (fromLabel === toLabel) {
          console.warn('[Sankey] Skipping self-loop for', payerRaw);
          return;
        }
        sankeyRows.push([fromLabel, toLabel, amount]);
      });
    });

    if (sankeyRows.length === 1) {
      console.log(`[Sankey] Empty inspected=${inspected} expenses=${expenseCount}`);
      return [];
    }
    console.log(`[Sankey] Rows=${sankeyRows.length-1} inspected=${inspected} expenses=${expenseCount}`);
    return sankeyRows;
  } catch (err) {
    console.error('calculateSankeyData_ error', err.stack || err);
    return [];
  }
}

/**
 * Menghitung total tabungan berdasarkan transaksi Input:
 * - Mendeteksi wallet yang di-tag melalui kolom Source = 'Saving/Investment' atau 'Other Asset'
 * - Mengambil saldo per wallet (pakai calculateWalletStatus_ jika ada), lalu jumlahkan saldo wallet yang terdeteksi
 * - Jika tidak ada deteksi Source, fallback pakai Wallet Type di Wallet Setup ('Other Asset' / 'Savings')
 */
function calculateTotalSaving_(filteredTransactions) {
  // Jika tidak ada transaksi yang sudah difilter, hasilnya 0
  if (!filteredTransactions || filteredTransactions.length === 0) {
    return 0;
  }

  // Skenario 1: Kumpulan kata kunci untuk 'Source' pada tabungan normal (pemasukan)
  const normalSavingSourceKeys = new Set(['saving/investment', 'other asset', 'investment', 'otherasset']);

  // Skenario 2: Aturan untuk tabungan tersamarkan (dicatat sebagai expense)
  const disguisedSavingRegex = new RegExp('(^|[^a-z])(tabungan|menabung|saving|savings?|autosave|investment|investasi|deposito?|reksadana|mutualfund|saham|stock|equity|obligasi|bond|pensiun|retirement|emergencyfund|aset|asset|capital)([^a-z]|$)', 'i');
  const disguisedSavingSourceKeys = new Set(['saving/investment', 'other asset', 'cash and bank', 'cash & bank']);

  let totalSavedInPeriod = 0;

  filteredTransactions.forEach(t => {
    // Normalisasi data. Untuk category/subcategory, kita tidak ubah ke lowercase
    // karena regex sudah case-insensitive ('i') dan butuh case asli untuk boundary check.
    const source = (t.Source || '').toString().trim().toLowerCase();
    const category = (t.Category || '').toString().trim();
    const subcategory = (t.Subcategory || '').toString().trim();
    const type = (t.Type || '').toString().trim().toLowerCase();

    // Skenario 1: Tabungan Normal (Pemasukan ke 'Source' tabungan)
    // Ini adalah kasus utama, biasanya dari 'Transfer-In'.
    if (t.Amount > 0 && normalSavingSourceKeys.has(source)) {
      totalSavedInPeriod += t.Amount;
      return; // Lanjut ke transaksi berikutnya agar tidak dihitung ganda.
    }

    // Skenario 2: Tabungan Tersamarkan (dicatat sebagai 'Expense' dengan kriteria ketat)
    if (type === 'expense' &&
        (disguisedSavingRegex.test(category) || disguisedSavingRegex.test(subcategory)))
    {
      // Karena type 'expense', t.Amount akan negatif. Kita ambil nilai absolutnya.
      totalSavedInPeriod += Math.abs(t.Amount);
    }
  });

  return totalSavedInPeriod;
}

/**
 * Menghitung data expense untuk TreeMap berdasarkan kategori dan subkategori untuk periode saat ini dan sebelumnya.
 * @param {Array<Object>} currentTransactions Transaksi untuk periode saat ini.
 * @param {Array<Object>} previousTransactions Transaksi untuk periode sebelumnya.
 * @returns {Object} Objek dengan data untuk TreeMap, termasuk nilai periode sebelumnya.
 */
function calculateExpenseTreeMapData_(currentTransactions, previousTransactions) {
  
  // Helper function untuk mengakumulasi pengeluaran dari daftar transaksi
  const _accumulateExpenses = (transactions) => {
    const result = {
      total: 0,
      categories: {},
      allSubcategories: {}
    };
    
    const expenseTransactions = (transactions || []).filter(t => t.Amount < 0);

    expenseTransactions.forEach(t => {
      const category = t.Category || 'Uncategorized';
      const subcategory = t.Subcategory || 'General';
      const amount = Math.abs(t.Amount);
      
      result.total += amount;
      
      if (!result.categories[category]) {
        result.categories[category] = { name: category, value: 0, subcategories: {} };
      }
      result.categories[category].value += amount;
      
      if (!result.categories[category].subcategories[subcategory]) {
        result.categories[category].subcategories[subcategory] = { name: subcategory, value: 0 };
      }
      result.categories[category].subcategories[subcategory].value += amount;
      
      if (!result.allSubcategories[subcategory]) {
        result.allSubcategories[subcategory] = { name: subcategory, value: 0, category: category };
      }
      result.allSubcategories[subcategory].value += amount;
    });
    return result;
  };

  const currentData = _accumulateExpenses(currentTransactions);
  const prevData = _accumulateExpenses(previousTransactions);

  // Helper untuk menggabungkan data saat ini dengan data sebelumnya
  const mergeWithPrev = (currentItems, prevItems, isSubcategory = false) => {
    return Object.values(currentItems).map(currentItem => {
      // Kunci pencarian adalah nama item (kategori atau subkategori)
      const prevItem = prevItems[currentItem.name];
      const mergedItem = {
        name: currentItem.name,
        value: currentItem.value,
        prev_value: prevItem ? prevItem.value : 0
      };
      // Jika ini subkategori, tambahkan properti 'category'
      if (isSubcategory && currentItem.category) {
        mergedItem.category = currentItem.category;
      }
      return mergedItem;
    });
  };

  const byCategory = mergeWithPrev(currentData.categories, prevData.categories);
  const bySubcategory = mergeWithPrev(currentData.allSubcategories, prevData.allSubcategories, true);

  // Buat data hierarkis dari data yang sudah digabung
  const hierarchical = byCategory.map(cat => {
    const children = bySubcategory
      .filter(sub => sub.category === cat.name)
      .map(sub => ({
        name: sub.name,
        value: sub.value,
        prev_value: sub.prev_value
      }));
      
    return {
      name: cat.name,
      value: cat.value,
      prev_value: cat.prev_value,
      children: children
    };
  });

  return {
    total: currentData.total,
    hierarchical: hierarchical,
    byCategory: byCategory,
    bySubcategory: bySubcategory
  };
}

/**
 * Menghitung status wallet berdasarkan transaksi di sheet Input dan metadata dari Wallet Setup.
 * @param {Array<Array>} allTransactionsData Data mentah dari sheet Input.
 * @param {Array<Array>} walletData Data mentah dari sheet Wallet Setup (untuk Type dan Owner).
 * @returns {Array<Object>} Data wallet dengan UniqueID, Wallet, Type, Owner, Balance.
 */
function calculateWalletStatus_(allTransactionsData, walletData) {
  if (!allTransactionsData) throw new Error('Data Input tidak tersedia.');

  // 1. Hitung balance per wallet dari transaksi di Input
  const balanceMap = {}; // wallet -> total balance
  const currentAllTransactions = [...allTransactionsData];
  const transHeaders = currentAllTransactions.shift();
  const transWalletCol = transHeaders.indexOf('Wallet');
  const transAmountCol = transHeaders.indexOf('Amount');
  const transTypeCol = transHeaders.indexOf('Transaction Type');
  const transSubcatCol = transHeaders.indexOf('Subcategory');
  const transSourceCol = transHeaders.indexOf('Source');

  // Map wallet -> set of observed sources (to infer Type)
  const walletSources = {};

  currentAllTransactions.forEach(row => {
    const wallet = row[transWalletCol];
    if (!wallet) return;

    let amount = normalizeNumber_(row[transAmountCol]);
    const transactionType = (row[transTypeCol] || '').toString().toLowerCase();
    const subcategoryRaw = (row[transSubcatCol] || '').toString().toLowerCase();

    if (transactionType === 'income') amount = Math.abs(amount);
    else if (transactionType === 'expense') amount = -Math.abs(amount);
    else if (transactionType === 'transfer') {
      if (subcategoryRaw === 'transfer-out') amount = -Math.abs(amount);
      else if (subcategoryRaw === 'transfer-in') amount = Math.abs(amount);
      else amount = 0;
    }

    if (!balanceMap[wallet]) balanceMap[wallet] = 0;
    balanceMap[wallet] += amount;

    // collect source samples for this wallet
    if (transSourceCol >= 0) {
      const src = row[transSourceCol];
      if (src) {
        walletSources[wallet] = walletSources[wallet] || new Set();
        walletSources[wallet].add(String(src).trim());
      }
    }
  });

  // 2. Ambil metadata (Type, Owner) dari Wallet Setup (jika ada)
  const metadataMap = {};
  if (walletData) {
    const currentWalletData = [...walletData];
    const walletHeaders = currentWalletData.shift();
    const walletCol = walletHeaders.indexOf('Wallet');
    const typeCol = walletHeaders.indexOf('Wallet Type');
    const ownerCol = walletHeaders.indexOf('Wallet Owner');

    currentWalletData.forEach(row => {
      const wallet = row[walletCol];
      if (wallet) {
        metadataMap[wallet] = {
          Type: row[typeCol] || '',
          Owner: row[ownerCol] || ''
        };
      }
    });
  }

  // 3. Gabungkan balance dan metadata; jika Type kosong, infer dari observed Source
  const result = [];
  Object.keys(balanceMap).forEach(wallet => {
    const balance = balanceMap[wallet];
    const meta = metadataMap[wallet] || { Type: '', Owner: '' };
    let inferredType = meta.Type || '';
    const walletNameLower = (wallet || '').toLowerCase();

    // --- LOGIKA INFERENSI BARU ---
    // Prioritas 1: Coba tebak dari nama wallet itu sendiri
    if (!inferredType) {
        if (/bca|mandiri|bni|bri|cimb|dbs|uob|ocbc|bank|rekening/.test(walletNameLower)) {
            inferredType = 'Cash & Bank';
        } else if (/gopay|ovo|dana|shopeepay|linkaja|ewallet|e-wallet/.test(walletNameLower)) {
            inferredType = 'E-Wallet';
        }
    }

    // Prioritas 2: Jika masih kosong, tebak dari 'Source' transaksi (logika lama)
    if (!inferredType && walletSources[wallet] && walletSources[wallet].size) {
      for (const s of walletSources[wallet]) {
        if (isLiquidSource_(s)) { inferredType = 'Cash & Bank'; break; }
      }
      if (!inferredType) {
        for (const s of walletSources[wallet]) {
          const sn = normStr_(s);
          if (sn.includes('e-wallet') || sn.includes('ewallet')) { inferredType = 'E-Wallet'; break; }
        }
      }
    }
    // --- AKHIR LOGIKA BARU ---

    result.push({
      UniqueID: Utilities.getUuid(),
      Wallet: wallet,
      Type: inferredType,
      Owner: meta.Owner || '',
      Balance: balance,
      Sources: Array.from(walletSources[wallet] || []) // <-- expose observed Source samples
    });
  });

  return result;
}

/**
 * Menghitung snapshot Net Worth (Assets - Liabilities) sampai cutoffDate (inklusif).
 * Assets: saldo semua wallet dihitung dari transaksi <= cutoffDate.
 * Liabilities: jumlah absolut transaksi dengan Source 'Liabilities' / 'Liability' (<= cutoffDate).
 * @param {Array<Array>} allTransactionsData - Data mentah sheet Input (termasuk header).
 * @param {Date} cutoffDate - Tanggal akhir snapshot.
 * @returns {{assets:number, liabilities:number, netWorth:number}}
 */
function calculateNetWorthSnapshot_(allTransactionsData, cutoffDate, ownerFilter) {
  if (!allTransactionsData || allTransactionsData.length < 2 || !cutoffDate) {
    return { assets: 0, liabilities: 0, netWorth: 0 };
  }
  const copy = [...allTransactionsData];
  const headers = copy.shift();
  const idxDate = headers.indexOf('Date');
  const idxWallet = headers.indexOf('Wallet');
  const idxAmount = headers.indexOf('Amount');
  const idxType = headers.indexOf('Transaction Type');
  const idxSubcat = headers.indexOf('Subcategory');
  const idxSource = headers.indexOf('Source');
  const idxOwner = headers.indexOf('Wallet Owner');

  const balanceMap = {}; // wallet -> balance
  let totalLiabilities = 0;

  copy.forEach(row => {
    const rawDate = row[idxDate];
    const d = new Date(rawDate);
    if (isNaN(d.getTime()) || d > cutoffDate) return;

  const wallet = idxWallet >= 0 ? row[idxWallet] : '';
  const ownerVal = idxOwner >= 0 ? row[idxOwner] : '';
  if (ownerFilter && ownerVal !== ownerFilter) return; // skip if not matching owner when filtering
    let amount = normalizeNumber_(row[idxAmount]);
    const type = (idxType >= 0 ? row[idxType] : '').toString().toLowerCase();
    const subcat = (idxSubcat >= 0 ? row[idxSubcat] : '').toString().toLowerCase();
    const sourceRaw = idxSource >= 0 ? row[idxSource] : '';
    const sourceNorm = normStr_(sourceRaw);

    // Normalisasi arah transaksi (sama logika getFilteredTransactions_)
    if (type === 'income') {
      amount = Math.abs(amount);
    } else if (type === 'expense') {
      amount = -Math.abs(amount);
    } else if (type === 'transfer') {
      if (subcat === 'transfer-out') amount = -Math.abs(amount);
      else if (subcat === 'transfer-in') amount = Math.abs(amount);
      else amount = 0;
    }

    if (wallet) {
      if (!balanceMap[wallet]) balanceMap[wallet] = 0;
      balanceMap[wallet] += amount;
    }

    // Liabilities: gunakan helper isLiabilitiesSource_
    if (isLiabilitiesSource_(sourceNorm)) {
      // Only count liability if owner matches (or no owner filter)
      if (!ownerFilter || ownerVal === ownerFilter) {
        totalLiabilities += Math.abs(amount);
      }
    }
  });

  const totalAssets = Object.values(balanceMap).reduce((s, v) => s + v, 0);
  return {
    assets: totalAssets,
    liabilities: totalLiabilities,
    netWorth: totalAssets - totalLiabilities
  };
}

// #endregion

// #region HELPER FUNCTIONS
// =================================================================
//                        HELPER FUNCTIONS
// =================================================================

/**
 * Helper untuk mengambil data mentah dari sheet tertentu, dengan dukungan cache.
 * @param {string} sheetName Nama sheet yang akan diambil datanya.
 * @param {boolean} forceRefresh Jika true, akan mengabaikan cache.
 * @returns {Array<Array>} Data mentah dari sheet (termasuk header).
 */
function getRawSheetData_(sheetName, forceRefresh) {
  const cache = CacheService.getUserCache();
  const cacheKey = `rawSheetData_${sheetName}`;
  let cachedData = cache.get(cacheKey);

  if (cachedData != null && !forceRefresh) {
    console.log(`Mengambil data sheet '${sheetName}' dari cache.`);
    return JSON.parse(cachedData);
  }

  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('MAIN_SHEET_ID'));
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.error(`Sheet '${sheetName}' tidak ditemukan.`);
    throw new Error(`Sheet '${sheetName}' tidak ditemukan.`);
  }

  const data = sheet.getDataRange().getValues();
  
  // Cache untuk sheet yang lebih statis (Wallet, Category, Goals) lebih lama
  // Untuk Input dan ScheduledTransactions, cache sangat singkat atau tidak sama sekali (jika forceRefresh true)
  let cacheExpiration = 300; // Default 5 menit
  if (sheetName === WALLET_SETUP_SHEET || sheetName === CATEGORY_SHEET || sheetName === GOALS_SHEET) {
    cacheExpiration = 3600; // 1 jam untuk sheet setup
  } 
  // Jika forceRefresh true, data tidak akan diambil dari cache, tapi akan diperbarui ke cache.
  // Jika forceRefresh false, tapi cache kadaluarsa, maka akan diambil dari sheet dan disimpan ke cache.
  // Untuk DATA_SHEET dan SCHEDULED_SHEET_NAME, kita biarkan default 5 menit atau langsung refresh jika tombol refresh ditekan.

  // Simpan ke cache hanya jika ukuran data kecil (hindari limit 100KB CacheService)
  putCacheIfSmall_(cache, cacheKey, data, cacheExpiration);
  console.log(`Data sheet '${sheetName}' diambil dari sheet dan disimpan ke cache.`);
  return data;
}

/**
 * Menyimpan ke CacheService hanya jika payload cukup kecil (< ~95KB).
 * Menghindari error: "Data larger than maximum size allowed" saat data besar.
 * @param {Cache} cache Instance cache (User/Script cache)
 * @param {string} key Key cache
 * @param {any} value Objek/string yang akan disimpan
 * @param {number} seconds Durasi TTL
 */
function putCacheIfSmall_(cache, key, value, seconds) {
  try {
    const str = (typeof value === 'string') ? value : JSON.stringify(value);
    // Batas resmi per entry CacheService ~100KB. Pakai ambang konservatif 95KB.
    if (str && str.length <= 95000) {
      cache.put(key, str, seconds);
    } else {
      console.log(`Skip cache for key '${key}' (size=${str ? str.length : 0} bytes)`);
    }
  } catch (e) {
    // Jangan gagal hanya karena caching bermasalah
    console.warn(`Gagal menyimpan cache untuk '${key}': ${e && e.message}`);
  }
}


 /**
 * Helper untuk memfilter transaksi berdasarkan kriteria.
 * @param {Array<Array>} allData Data mentah dari sheet Input (termasuk header).
 * @param {object} filters Objek filter.
 * @param {Date} startDate Tanggal mulai.
 * @param {Date} endDate Tanggal akhir.
 * @returns {Array<Object>} Array objek transaksi yang sudah difilter.
 */
function getFilteredTransactions_(allData, filters, startDate, endDate) {
  const currentAllData = [...allData];
  const headers = currentAllData.shift();
  const colMap = {
    date: headers.indexOf('Date'),
    type: headers.indexOf('Transaction Type'),
    amount: headers.indexOf('Amount'),
    wallet: headers.indexOf('Wallet'),
    owner: headers.indexOf('Wallet Owner'),
    purpose: headers.indexOf('Expense Purpose'),
    category: headers.indexOf('Category'),
    subcategory: headers.indexOf('Subcategory'),
    note: headers.indexOf('Note'),
    description: headers.indexOf('Description'),
    source: headers.indexOf('Source') // <-- TAMBAHKAN INI
  };

  return currentAllData.map(row => {
    const date = new Date(row[colMap.date]);
    if (date < startDate || date > endDate) return null;

    // Normalisasi angka mentah
    let amount = normalizeNumber_(row[colMap.amount]);

    const transactionType = (row[colMap.type] || '').toString().toLowerCase();
    const subcategoryRaw = (colMap.subcategory >= 0 ? (row[colMap.subcategory] || '') : '').toString().toLowerCase();

    if (transactionType === 'income') {
      amount = Math.abs(amount);
    } else if (transactionType === 'expense') {
      amount = -Math.abs(amount);
    } else if (transactionType === 'transfer') {
      if (subcategoryRaw === 'transfer-out') amount = -Math.abs(amount);
      else if (subcategoryRaw === 'transfer-in') amount = Math.abs(amount);
      else amount = 0;
    }

    const transaction = {
      Date: date,
      Type: row[colMap.type],
      Amount: amount,
      Wallet: row[colMap.wallet],
      Owner: row[colMap.owner],
      Purpose: row[colMap.purpose],
      Category: row[colMap.category],
      Subcategory: row[colMap.subcategory],
      Note: colMap.note >= 0 ? row[colMap.note] : '',
      Description: colMap.description >= 0 ? (row[colMap.description] || '') : '',
      Source: colMap.source >= 0 ? row[colMap.source] : ''
    };

    if (filters.wallet && transaction.Wallet !== filters.wallet) return null;
    if (filters.walletOwner && transaction.Owner !== filters.walletOwner) return null;
    if (filters.expensePurpose && transaction.Purpose !== filters.expensePurpose) return null;
    if (filters.category && transaction.Category !== filters.category) return null;
    if (filters.subcategory && transaction.Subcategory !== filters.subcategory) return null;
    if (filters.note && transaction.Note !== filters.note) return null;
    if (filters.description) {
      const q = String(filters.description || '').toLowerCase();
      if (!String(transaction.Description || '').toLowerCase().includes(q)) return null;
    }

    return transaction;
  }).filter(Boolean);
}


/**
 * Helper untuk memformat baris data budget.
 * @param {string} category Nama kategori.
 * @param {string} subcategory Nama subkategori.
 * @param {number} budget Jumlah budget.
 * @param {number} expense Jumlah pengeluaran.
 * @returns {object} Objek data budget yang diformat.
 */
function formatBudgetRow_(category, subcategory, budget, expense) {
    const usagePercentage = budget > 0 ? (expense / budget) * 100 : 0;
    let status;
    if (usagePercentage > BUDGET_OVER_THRESHOLD) status = 'Over';
    else if (usagePercentage > BUDGET_WARNING_THRESHOLD) status = 'Warning';
    else status = 'On Track';
    
    return {
      UniqueID: Utilities.getUuid(),
      Category: category,
      Subcategory: subcategory,
      BudgetAmount: budget,
      ActualExpense: expense,
      RemainingBudget: budget - expense,
      UsagePercentage: parseFloat(usagePercentage.toFixed(1)),
      Status: status,
    };
}

/**
 * Menghitung tanggal awal dan akhir berdasarkan periode yang diberikan.
 * @param {string} period Periode waktu.
 * @returns {{startDate: Date, endDate: Date}} Objek berisi tanggal awal dan akhir.
 */
function getPeriodDates_(period, customStart, customEnd) {
  const today = new Date();
  let startDate, endDate;

  // Prioritaskan custom range dari filter
  if (period === 'custom' && customStart && customEnd) {
    startDate = new Date(customStart);
    endDate = new Date(customEnd);
    startDate.setHours(0,0,0,0);
    endDate.setHours(23,59,59,999);
    return { startDate, endDate };
  }

  switch (period) {
    case 'all':
      // rentang sangat luas untuk mencakup seluruh history sheet
      startDate = new Date(1900, 0, 1);
      endDate = new Date(9999, 11, 31);
      break;
    case 'today':
      startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      break;
    case 'yesterday':
      startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      break;
    case 'this_week':
      const firstDayOfWeek = today.getDate() - today.getDay();
      startDate = new Date(today.setDate(firstDayOfWeek));
      endDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 6);
      break;
    case 'last_7_days':
      endDate = new Date(today);
      startDate = new Date();
      startDate.setDate(endDate.getDate() - 6);
      break;
    case 'last_month':
      startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      endDate = new Date(today.getFullYear(), today.getMonth(), 0);
      break;
    case 'current_year':
      startDate = new Date(today.getFullYear(), 0, 1);
      endDate = new Date(today.getFullYear(), 11, 31);
      break;
    case 'last_year':
      startDate = new Date(today.getFullYear() - 1, 0, 1);
      endDate = new Date(today.getFullYear() - 1, 11, 31);
      break;
    case 'current_month':
    default:
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      break;
  }

  // pastikan inklusif waktu
  startDate.setHours(0,0,0,0);
  endDate.setHours(23,59,59,999);
  return { startDate, endDate };
}

/**
 * Menghitung tanggal periode sebelumnya berdasarkan periode saat ini.
 * @param {string} period Periode waktu saat ini.
 * @param {Date} currentStartDate Tanggal mulai periode saat ini.
 * @returns {{startDate: Date, endDate: Date}} Objek berisi tanggal awal dan akhir periode sebelumnya.
 */
function getPreviousPeriodDates_(period, currentStartDate) {
  let startDate, endDate;
  const prevEndDate = new Date(currentStartDate.getTime() - 1); // Sehari sebelum periode saat ini mulai

  switch (period) {
    case 'today':
    case 'yesterday':
      startDate = new Date(prevEndDate.getFullYear(), prevEndDate.getMonth(), prevEndDate.getDate());
      endDate = new Date(prevEndDate.getFullYear(), prevEndDate.getMonth(), prevEndDate.getDate());
      break;
    case 'last_7_days':
    case 'this_week':
      startDate = new Date(prevEndDate.getTime() - 6 * 24 * 60 * 60 * 1000);
      endDate = prevEndDate;
      break;
    case 'current_month':
    case 'last_month':
      startDate = new Date(prevEndDate.getFullYear(), prevEndDate.getMonth(), 1);
      endDate = prevEndDate;
      break;
    case 'current_year':
    case 'last_year':
      startDate = new Date(prevEndDate.getFullYear(), 0, 1);
      endDate = prevEndDate;
      break;
    default: // 'all' atau 'custom'
      // Untuk 'all' atau 'custom', perbandingan tidak didefinisikan, jadi kembalikan rentang kosong
      return { startDate: new Date(0), endDate: new Date(0) };
  }

  startDate.setHours(0,0,0,0);
  endDate.setHours(23,59,59,999);
  return { startDate, endDate };
}


 /**
 * Memformat objek tanggal untuk ditampilkan di UI.
 * @param {Date} date Objek tanggal.
 * @returns {string} Tanggal berformat 'd MMM yyyy'.
 */
function formatDateForDisplay_(date) {
  if (!date || !(date instanceof Date)) return 'N/A';
  try {
    return Utilities.formatDate(date, "GMT+7", "d MMM yyyy");
  } catch (e) {
    return 'Invalid Date';
  }
}

// ---------- Case-insensitive helpers ----------
function ciIndex_(headers, target) {
  if (!headers) return -1;
  const want = String(target).trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === want) return i;
  }
  return -1;
}
function normStr_(v) {
  return String(v === null || v === undefined ? '' : v).trim().toLowerCase();
}
function isLiabilitiesSource_(v) {
  const s = normStr_(v);
  return s === 'liabilities' || s === 'liability';
}

// tambah helper baru untuk deteksi jenis Source likuid
function isLiquidSource_(src) {
  const s = normStr_(src || '');
  return ['cash & bank','cash and bank','cash','bank','e-wallet','ewallet','digital wallet','gopay','ovo'].some(k => s.includes(k));
}

/**
 * Normalisasi angka string berbagai format (ID / EN).
 * Contoh:
 *  "1.250.000"    -> 1250000
 *  "1,250,000"    -> 1250000
 *  "1.234,56"     -> 1234.56
 *  "1,234.56"     -> 1234.56
 *  "Rp 2.500,75"  -> 2500.75
 *  "- 3.000"      -> -3000
 */
function normalizeNumber_(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  let s = String(val).trim();
  if (!s) return 0;

  // Allow leading sign
  let sign = 1;
  if (/^-/.test(s.replace(/\s+/g,''))) sign = -1;
  s = s.replace(/[()+\s]|Rp|IDR/gi,'').trim();

  // Keep only digits , .'
  s = s.replace(/[^0-9,.-]/g,'');

  // Positions
  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  let decimalSep = null;
  if (lastComma !== -1 && lastDot !== -1) {
    decimalSep = lastComma > lastDot ? ',' : '.';
  } else if (lastComma !== -1) {
    if (s.length - lastComma - 1 <= 2) decimalSep = ',';
  } else if (lastDot !== -1) {
    if (s.length - lastDot - 1 <= 2) decimalSep = '.';
  }

  if (decimalSep) {
    const thousandSep = decimalSep === ',' ? '.' : ',';
    s = s.split(thousandSep).join('');
    if (decimalSep === ',') s = s.replace(',', '.');
  } else {
    // remove all separators
    s = s.replace(/[.,]/g,'');

  }

  const num = parseFloat(s);
  return isNaN(num) ? 0 : num * sign;
}

/**
 * Hitung jumlah aset likuid dari transaksi sampai cutoffDate.
 * @param {Array<Array>} allTransactionsData Data mentah dari sheet Input (termasuk header)
 * @param {Date} cutoffDate Tanggal akhir snapshot
 * @returns {Number} Total aset likuid per cutoffDate
 */
function computeLiquidAssetsSnapshot_(allTransactionsData, cutoffDate) {
  if (!allTransactionsData || allTransactionsData.length < 2 || !cutoffDate) return 0;
  const copy = [...allTransactionsData];
  const headers = copy.shift();
  const idxDate = headers.indexOf('Date');
  const idxWallet = headers.indexOf('Wallet');
  const idxAmount = headers.indexOf('Amount');
  const idxType = headers.indexOf('Transaction Type');
  const idxSubcat = headers.indexOf('Subcategory');
  const idxSource = headers.indexOf('Source');

  // Hitung saldo per wallet sampai cutoffDate
  const balanceMap = {};
  const walletSources = {};

  copy.forEach(row => {
    const rawDate = row[idxDate];
    const d = new Date(rawDate);
    if (isNaN(d.getTime()) || d > cutoffDate) return;

    const wallet = row[idxWallet];
    if (!wallet) return;

    let amount = normalizeNumber_(row[idxAmount]);
    const transactionType = (row[idxType] || '').toString().toLowerCase();
    const subcategoryRaw = (row[idxSubcat] || '').toString().toLowerCase();

    if (transactionType === 'income') amount = Math.abs(amount);
    else if (transactionType === 'expense') amount = -Math.abs(amount);
    else if (transactionType === 'transfer') {
      if (subcategoryRaw === 'transfer-out') amount = -Math.abs(amount);
      else if (subcategoryRaw === 'transfer-in') amount = Math.abs(amount);
      else amount = 0;
    }

    if (!balanceMap[wallet]) balanceMap[wallet] = 0;
    balanceMap[wallet] += amount;

    // collect source samples for this wallet
    if (idxSource >= 0) {
      const src = row[idxSource];
      if (src) {
        walletSources[wallet] = walletSources[wallet] || new Set();
        walletSources[wallet].add(String(src).trim());
      }
    }
  });

  // Deteksi wallet likuid
  let totalLiquid = 0;
  Object.keys(balanceMap).forEach(wallet => {
    let isLiquid = false;
    const sources = Array.from(walletSources[wallet] || []);
    if (sources.length > 0) {
      isLiquid = sources.some(source => isLiquidSource_(source));
    }
    if (!isLiquid) {
      const walletName = (wallet || '').toString().toLowerCase();
      if (/bca|bank|gopay|ovo/.test(walletName)) isLiquid = true;
    }
    if (isLiquid) totalLiquid += balanceMap[wallet];
  });

  return totalLiquid;
}

// #endregion

// #region PWA HELPERS (Manifest & Service Worker)
// =================================================================
//      PWA HELPERS (Manifest JSON & Service Worker JS)
// =================================================================

/**
 * Menghasilkan isi manifest PWA sebagai string JSON.
 */
function getManifestJson_() {
  var manifest = {
    name: 'SatukasSatukas — Financial Dashboard',
    short_name: 'SkSk',
    start_url: './',
    display: 'standalone',
    background_color: '#f7fafc',
    theme_color: '#f7fafc',
    description: 'Personal financial dashboard for SatukasSatukas',
    icons: [
      { src: 'data:image/svg+xml;base64,' + Utilities.base64Encode(getIconSvg_()), sizes: 'any', type: 'image/svg+xml', purpose: 'any' },
      { src: 'data:image/svg+xml;base64,' + Utilities.base64Encode(getMaskableSvg_()), sizes: 'any', type: 'image/svg+xml', purpose: 'maskable' }
    ]
  };
  return JSON.stringify(manifest);
}

/**
 * Menghasilkan isi Service Worker sebagai string JS.
 */
function getServiceWorkerJs_() {
  return (
    "const CACHE_NAME='sksk-cache-v1';" +
    "const CORE_ASSETS=['./','?offline=1'];" +
    "self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE_NAME).then(c=>c.addAll(CORE_ASSETS)).then(()=>self.skipWaiting()))});" +
    "self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(keys=>Promise.all(keys.map(k=>k!==CACHE_NAME?caches.delete(k):Promise.resolve()))).then(()=>self.clients.claim()))});" +
    "async function networkFirst(r,f){try{const fresh=await fetch(r);const c=await caches.open(CACHE_NAME);c.put(r,fresh.clone());return fresh}catch(err){const cached=await caches.match(r);if(cached)return cached;if(f)return caches.match(f);throw err}}" +
    "async function cacheFirst(r){const cached=await caches.match(r);if(cached)return cached;const fresh=await fetch(r);const c=await caches.open(CACHE_NAME);c.put(r,fresh.clone());return fresh}" +
    "self.addEventListener('fetch',e=>{const r=e.request;if(r.mode==='navigate'){e.respondWith(networkFirst(r,'?offline=1'));return}const d=r.destination;if(['style','script','image','font'].includes(d)){e.respondWith(cacheFirst(r));return}e.respondWith(fetch(r).then(res=>{const copy=res.clone();caches.open(CACHE_NAME).then(c=>c.put(r,copy)).catch(()=>{});return res}).catch(()=>caches.match(r)))})"
  );
}

// Inline SVG icon sources (agar tidak perlu file statik terpisah)
function getIconSvg_(){
  return '<svg xmlns="http://www.w3.org/2000/svg" width="512" height="512" viewBox="0 0 512 512">'+
         '<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">'+
         '<stop offset="0%" stop-color="#06b6d4"/><stop offset="100%" stop-color="#3b82f6"/></linearGradient></defs>'+
         '<rect width="512" height="512" rx="96" fill="url(#g)"/>'+
         '<g fill="#fff" font-family="Inter, system-ui, Arial" font-weight="700" font-size="200" text-anchor="middle">'+
         '<text x="256" y="300">SkSk</text></g></svg>';
}

function getMaskableSvg_(){
  return '<svg xmlns="http://www.w3.org/2000/svg" width="512" height="512" viewBox="0 0 512 512">'+
         '<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">'+
         '<stop offset="0%" stop-color="#22d3ee"/><stop offset="100%" stop-color="#60a5fa"/></linearGradient>'+
         '<clipPath id="maskable"><path d="M68 0h376c37.6 0 68 30.4 68 68v376c0 37.6-30.4 68-68 68H68c-37.6 0-68-30.4-68-68V68C0 30.4 30.4 0 68 0z"/></clipPath></defs>'+
         '<rect width="512" height="512" fill="url(#g)" clip-path="url(#maskable)"/>'+
         '<g fill="#ffffff" font-family="Inter, system-ui, Arial" font-weight="700" font-size="200" text-anchor="middle" clip-path="url(#maskable)">'+
         '<text x="256" y="300">SkSk</text></g></svg>';
}

// #endregion

const SankeyChart = {
        props: ['data'],
        data() {
            return {
                topN: 5, // ensure default Top 5 here too
                linkOptions: [5, 7, 10]
            };
        },
        computed: {
            // Process input rows into aggregated flows where "Other (Payer)" groups PAYERS outside topN
            processedRows() {
                if (!this.data || this.data.length === 0) return [];
                // If backend sent header row, skip non-number rows in col 2
                const rawRows = this.data.filter(r => Array.isArray(r) && !isNaN(Number(r[2]))).map(r => [String(r[0]), String(r[1]), Number(r[2])]);
                if (rawRows.length === 0) return [];

                // Compute total outgoing per payer
                const totalsByPayer = {};
                rawRows.forEach(r => {
                    const payer = r[0];
                    totalsByPayer[payer] = (totalsByPayer[payer] || 0) + (Number(r[2]) || 0);
                });

                // Determine topN payers by total outgoing
                const payersSorted = Object.keys(totalsByPayer).sort((a,b) => totalsByPayer[b] - totalsByPayer[a]);
                const topPayers = new Set(payersSorted.slice(0, this.topN));

                // Aggregate flows
                const keptFlows = []; // flows where payer in topPayers
                const otherToBenef = {}; // beneficiary -> total from all other payers

                rawRows.forEach(r => {
                    const payer = r[0];
                    const benef = r[1];
                    const value = r[2];

                    if (topPayers.has(payer)) {
                        // Payer is in topN, keep this flow

                        keptFlows.push({ from: payer, to: benef, value });
                    } else {
                        // Aggregate to "Other (Payer)"
                        if (!otherToBenef[benef]) otherToBenef[benef] = 0;
                        otherToBenef[benef] += value;
                    }
                });

                const otherFlows = Object.keys(otherToBenef).map(b => ({ from: 'Other (Payer)', to: b, value: otherToBenef[b] }));

                // Merge and filter zero
                const merged = keptFlows.concat(otherFlows).filter(f => f.value > 0);
                // Sort by value desc for stable rendering
                merged.sort((a,b) => b.value - a.value);
                return merged;
            }
        },
        template: `
        <section class="card p-5 flex flex-col h-full">
            <div class="flex items-start justify-between mb-3">
                <div>
                    <h3 class="text-slate-800 font-semibold tracking-tight mb-1">Owner to Purpose Flow</h3>
                    <p class="text-xs text-slate-500">Top flows (links) grouped; smaller payers aggregated as "Other (Payer)"</p>
                </div>
                <div class="flex items-center gap-3">
                    <label class="text-xs text-slate-500">Top</label>
                    <select v-model.number="topN" class="header-control text-sm">
                        <option v-for="n in linkOptions" :key="n" :value="n">{{ n }}</option>
                    </select>
                    <!-- refresh button is provided in the dashboard header; no in-card icon -->
                </div>
            </div>

            <div ref="chartContainer" class="w-full" style="height:300px; min-height:300px;">
                <div ref="chartdiv" style="width:100%; height:100%;"></div>
            </div>

            <div v-if="processedRows.length===0" class="text-center text-xs text-slate-400 mt-4">No data available for selected filters.</div>
        </section>
    `,
    mounted() {
        // defer initial draw slightly to avoid layout overflow on first render
        this.$nextTick(()=> setTimeout(this.drawChart, 80));
        window.addEventListener('resize', this.onResize);
    },
    watch: {
        data() { this.$nextTick(()=> setTimeout(this.drawChart, 80)); },
        topN() { this.$nextTick(this.drawChart); }
    },
    methods: {
        onResize() {
            clearTimeout(this._rz);
            this._rz = setTimeout(() => this.drawChart(), 150);
        },
        drawChart() {
            if (!this.processedRows.length) {
                if (this.$refs.chartdiv) this.$refs.chartdiv.innerHTML = '';
                return;
            }
            google.charts.setOnLoadCallback(() => {
                try {
                    const dt = new google.visualization.DataTable();
                    dt.addColumn('string', 'From');
                    dt.addColumn('string', 'To');
                    dt.addColumn('number', 'Weight');
                    dt.addColumn({ type: 'string', role: 'tooltip', p: { html: true } });

                    const LEFT_PREFIX = 'Payer: ';
                    const RIGHT_PREFIX = 'Beneficiary: ';

                    // Prepare totals per 'to' for percentage
                    const totalTo = {};
                    this.processedRows.forEach(f => {
                        const toLabel = RIGHT_PREFIX + String(f.to);
                        totalTo[toLabel] = (totalTo[toLabel] || 0) + f.value;
                    });

                    const rows = this.processedRows.map(f => {
                        // Node label tetap pakai prefix
                        const fromLabel = LEFT_PREFIX + String(f.from);
                        const toLabel = RIGHT_PREFIX + String(f.to);
                        const percent = totalTo[toLabel] ? (f.value / totalTo[toLabel]) * 100 : 0;
                        // Tooltip tanpa prefix double
                        const tip = `<div style="padding:6px 8px;white-space:nowrap">
  <div><strong>${f.from}</strong> → <strong>${f.to}</strong></div>
  <div>Amount: ${formatCurrency(f.value)}</div>
  <div>Share of '${f.to}': ${percent.toFixed(1)}%</div>
</div>`;
                        return [fromLabel, toLabel, f.value, tip];
                    });

                    dt.addRows(rows);

                    const containerWidth = this.$refs.chartContainer ? this.$refs.chartContainer.clientWidth : undefined;
                    const options = {
                        width: containerWidth,
                        height: this.$refs.chartContainer ? this.$refs.chartContainer.clientHeight : 300,
                        backgroundColor: 'transparent',
                        tooltip: { isHtml: true },
                        sankey: {
                            node: {
                                width: 12,
                                nodePadding: 18,
                                label: { fontName: 'Inter', fontSize: 11, color: '#1e293b' }
                            },
                            link: { colorMode: 'gradient' }
                        }
                    };

                    const el = this.$refs.chartdiv;
                    el.innerHTML = '';
                    const chart = new google.visualization.Sankey(el);
                    chart.draw(dt, options);
                } catch (e) {
                    console.error('Sankey draw error', e);
                }
            });
        }
    }
};

// ================== EXPORT CSV API (NEW) ==================
/**
 * Ekspor gabungan data dashboard + transaksi terfilter (dan opsional raw all time) dalam satu CSV ber-seksi.
 * @param {string} period
 * @param {Object} filters
 * @param {boolean} includeRawAll Jika true sertakan semua transaksi all time tanpa filter.
 * @returns {{filename:string, mime:string, contentBase64:string}}
 */
function exportDashboardCsv(period, filters, includeRawAll) {
  try {
    const safePeriod = period || 'current_month';
    const safeFilters = filters || {};
    // Ambil data dashboard (tidak force refresh agar pakai cache jika ada)
    const dash = getDashboardData(safePeriod, safeFilters, false);

    // Ambil transaksi terfilter (re-run agar kita dapat objek transaksi)
    const allTransactionsData = getRawSheetData_(DATA_SHEET, false);
    const { startDate, endDate } = getPeriodDates_(safePeriod, safeFilters.startDate, safeFilters.endDate);
    const filteredTx = getFilteredTransactions_(allTransactionsData, safeFilters, startDate, endDate) || [];

    // Raw all (ignore filters) bila diminta
    let rawAllTx = [];
    if (includeRawAll) {
      const allRange = getPeriodDates_('all');
      rawAllTx = getFilteredTransactions_(allTransactionsData, {}, allRange.startDate, allRange.endDate) || [];
    }

    const lines = [];
    const pushBlank = () => { if (lines.length && lines[lines.length-1] !== '') lines.push(''); };
    const esc = v => {
      if (v === null || v === undefined) return '';
      const s = String(v);
      return /[",\n]/.test(s) ? '"'+ s.replace(/"/g,'""') +'"' : s;
    };
    const writeSection = (title, rows, headersOrder) => {
      pushBlank();
      lines.push(`# ${title}`);
      if (!rows || !rows.length) { lines.push('(no rows)'); return; }
      const headers = headersOrder || Object.keys(rows[0]);
      lines.push(headers.map(esc).join(','));
      rows.forEach(r => {
        lines.push(headers.map(h => esc(r[h])).join(','));
      });
    };

    // Meta
    const now = new Date();
    lines.push(`# EXPORT SatukasSatukas`);
    lines.push(`Generated,${now.toISOString()}`);
    lines.push(`Period,${safePeriod}`);
    Object.keys(safeFilters).forEach(k => { if (safeFilters[k]) lines.push(`Filter:${k},${esc(safeFilters[k])}`); });

    // KPI Summary
    const kpi = dash.kpiSummary || {};
    // Hanya 4 KPI sesuai kartu di dashboard (Income, Expense, NetWorth, Saving)
    const kpiRows = [
      { Metric:'Income', Current:kpi.income||0, Previous:kpi.prev_income||0 },
      { Metric:'Expense', Current:kpi.expense||0, Previous:kpi.prev_expense||0 },
      { Metric:'NetWorth', Current:kpi.netWorth||0, Previous:kpi.prev_netWorth||0 },
      { Metric:'Saving', Current:kpi.saving||0, Previous:kpi.prev_saving||0 }
    ].map(r => { const diff = r.Current - r.Previous; const pct = r.Previous ? (diff/Math.abs(r.Previous))*100 : (r.Current?100:0); return { ...r, Diff: diff, DiffPct: pct.toFixed(2) }; });
    writeSection('KPI Summary', kpiRows, ['Metric','Current','Previous','Diff','DiffPct']);

    // Wallet Status
    writeSection('Wallet Status', (dash.walletStatus||[]).map(w => ({ Wallet:w.Wallet, Type:w.Type, Owner:w.Owner, Balance:w.Balance })), ['Wallet','Type','Owner','Balance']);

  // Goals Status (extended pacing fields)
  writeSection('Goals Status', (dash.goalsStatus||[]), ['GoalName','StartDate','Deadline','TotalNeeded','Collected','ProgressPercentage','TargetCumulative','GapAmount','GapPct','RemainingAmount','ElapsedRatio','PaceNeededPerDay','ActualPacePerDay','DaysLeft','ProjectedFinish','RiskScore','Status']);

    // Budget Status
    writeSection('Budget Status', (dash.budgetStatus||[]), ['Category','Subcategory','BudgetAmount','ActualExpense','RemainingBudget','UsagePercentage','Status']);

    // Liabilities & Upcoming
    writeSection('Liabilities Upcoming', (dash.liabilitiesUpcoming||[]).map(l=>({Type:l.Type, Name:l.Name, Amount:l.Amount, Wallet:l.Wallet, Owner:l.Owner, DisplayDate:l.DisplayDate, DueDate:l.DueDate, isOverdue:l.isOverdue})), ['Type','Name','Amount','Wallet','Owner','DisplayDate','DueDate','isOverdue']);

    // Ratios
    writeSection('Expense Ratios', (dash.ratios||[]).map(r=>({RatioType:r.RatioType, TotalExpense:r.TotalExpense, Sources: JSON.stringify(r.BySource||{}) })), ['RatioType','TotalExpense','Sources']);

    // Net Flow
    writeSection('Net Flow', (dash.netFlow||[]), ['PeriodLabel','Income','Expense','NetFlowAmount']);

    // Sankey
    if (dash.sankeyData && dash.sankeyData.length>1) {
      const sankeyRows = dash.sankeyData.slice(1).map(r => ({ From:r[0], To:r[1], Amount:r[2] }));
      writeSection('Sankey Flows', sankeyRows, ['From','To','Amount']);
    }

    // Expense TreeMap Category
    if (dash.expenseTreeMap) {
      writeSection('Expense Category', (dash.expenseTreeMap.byCategory||[]).map(c=>({Category:c.name, Value:c.value, PrevValue:c.prev_value})), ['Category','Value','PrevValue']);
      writeSection('Expense Subcategory', (dash.expenseTreeMap.bySubcategory||[]).map(c=>({Subcategory:c.name, Category:c.category, Value:c.value, PrevValue:c.prev_value})), ['Subcategory','Category','Value','PrevValue']);
    }

    // Helper aman untuk format tanggal (skip invalid)
    const safeDate = d => {
      if (!d) return '';
      try { if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime())) return d.toISOString().split('T')[0]; } catch(e){}
      // Jika bukan Date valid, coba parse string
      try { const nd = new Date(d); if (!isNaN(nd.getTime())) return nd.toISOString().split('T')[0]; } catch(e){}
      return '';
    };

    // Filtered Transactions
    writeSection('Filtered Transactions', filteredTx.map(t=>({
      Date: safeDate(t.Date),
      Type: t.Type,
      Amount: t.Amount,
      Wallet: t.Wallet,
      Owner: t.Owner,
      Purpose: t.Purpose,
      Category: t.Category,
      Subcategory: t.Subcategory,
      Note: t.Note,
      Description: t.Description,
      Source: t.Source
    })), ['Date','Type','Amount','Wallet','Owner','Purpose','Category','Subcategory','Note','Description','Source']);

    if (includeRawAll) {
      writeSection('Raw All Transactions', rawAllTx.map(t=>({
        Date: safeDate(t.Date),
        Type: t.Type,
        Amount: t.Amount,
        Wallet: t.Wallet,
        Owner: t.Owner,
        Purpose: t.Purpose,
        Category: t.Category,
        Subcategory: t.Subcategory,
        Note: t.Note,
        Description: t.Description,
        Source: t.Source
      })), ['Date','Type','Amount','Wallet','Owner','Purpose','Category','Subcategory','Note','Description','Source']);
    }

    const csv = lines.join('\n');
    const encoded = Utilities.base64Encode(csv, Utilities.Charset.UTF_8);
    return {
      filename: `sksk-transactions-${safePeriod}.csv`,
      mime: 'text/csv',
      contentBase64: encoded,
      base64: encoded // alias untuk kompatibilitas frontend awal
    };
  } catch(e) {
    console.error('exportDashboardCsv error', e.stack||e);
    throw new Error('Gagal menghasilkan export: '+ e.message);
  }
}
