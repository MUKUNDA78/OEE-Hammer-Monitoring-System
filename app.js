/**
 * FORGE OEE MONITORING & ANALYSIS SYSTEM
 * 5 Hammer Fleet Performance & Production Analytics Engine
 * Super-Resilient Auto-Commit Excel Parser & Mathematically Exact OEE Calculation Engine
 */

(function () {
  'use strict';

  // State Management
  let shiftLogs = [];
  let qualityLogs = [];
  let excelImportBuffer = [];
  let currentImportHammer = null; // null for combined, or hammer name string
  let charts = {};
  let currentTheme = localStorage.getItem('oee_theme') || 'light-green';
  let selectedSubnavHammer = 'ALL';

  // Supabase State & Auth
  let supabaseClient = null;
  let currentUserProfile = null;
  let currentAuthUser = null;
  let presenceChannel = null;

  // Public frontend Supabase configuration
const SUPABASE_CONFIG = {
  url: 'https://pydelymukfabbfhjcivg.supabase.co',
  key: 'sb_publishable_s1cjWwmuh5oW--fw0iWdvQ_DahRdZvn'
};

  // Target Equipment Specification
  const HAMMERS = [
    { name: '1 Ton Hammer', capacity: '1.0 Ton', color: '#2563eb', defaultCycle: 35, badgeId: 'countBadge_1Ton', samplePart: 'A1#21' },
    { name: '1.5 Ton Hammer', capacity: '1.5 Ton', color: '#7c3aed', defaultCycle: 42, badgeId: 'countBadge_1.5Ton', samplePart: 'W1#164' },
    { name: '2.5 Ton (Old) Hammer', capacity: '2.5 Ton (Old)', color: '#d97706', defaultCycle: 55, badgeId: 'countBadge_2.5TonOld', samplePart: 'A4#07' },
    { name: '2.5 Ton (New) Hammer', capacity: '2.5 Ton (New)', color: '#16a34a', defaultCycle: 48, badgeId: 'countBadge_2.5TonNew', samplePart: 'D1#45' },
    { name: '3.5 Ton Hammer', capacity: '3.5 Ton', color: '#db2777', defaultCycle: 70, badgeId: 'countBadge_3.5Ton', samplePart: 'M5#102' }
  ];

  // Global helper functions attached to window for inline HTML onclick handlers
  window.triggerHammerExcelBrowse = function (hammerName) {
    currentImportHammer = hammerName;
    const input = document.getElementById('hammerFileInput');
    if (input) input.click();
  };

  window.downloadHammerTemplate = function (hammerName) {
    downloadHammerSpecificTemplate(hammerName);
  };

  // Initialize Application
  document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    loadShiftLogs();
    checkUrlForSharedData();
    const migrateBtn = document.getElementById('restoreBackupLogsBtn');

if (migrateBtn) {
    migrateBtn.addEventListener('click', migrateLocalLogsToSupabase);
}
    const shareBtn = document.getElementById('shareDataLinkBtn');

if (shareBtn) {
    shareBtn.addEventListener('click', async () => {
        const liveUrl = window.location.origin + window.location.pathname;

        try {
            await navigator.clipboard.writeText(liveUrl);
            showToast('Live link copied to clipboard!', 'success');
        } catch (err) {
            prompt('Copy this Live Link:', liveUrl);
        }
    });
}
    setupLiveCalculator();
    setupExcelDropZone();
    setupHammerMiniDropzones();
    setupCloudDbSync();
    initSupabaseClient();
    renderAllViews();
  });

  /* ==========================================================================
     ROBUST NUMERIC & FUZZY MATCHING HELPERS
     ========================================================================== */
  function parseNum(val, defaultVal = 0) {
    if (val === null || val === undefined || val === '') return defaultVal;
    if (typeof val === 'number') return isNaN(val) ? defaultVal : val;
    const str = String(val).replace(/,/g, '').replace(/[^\d.-]/g, '');
    if (str === '' || str === '-') return defaultVal;
    const num = parseFloat(str);
    return isNaN(num) ? defaultVal : num;
  }

  function parseTimeMins(val, isHoursIfSmall = true) {
    let num = parseNum(val, 0);
    if (num <= 0) return 0;
    // If downtime/operating time is specified in hours (e.g. 0.5, 1.5, 2.0, 9.5 <= 12)
    if (isHoursIfSmall && num <= 12) {
      return num * 60;
    }
    return num;
  }

  function getStandardDateString(rawVal) {
    if (!rawVal) return '';
    if (rawVal instanceof Date) {
      if (isNaN(rawVal.getTime())) return '';
      return rawVal.toISOString().split('T')[0];
    }

    // Excel Serial Number (e.g. 45500)
    if (typeof rawVal === 'number' && rawVal > 30000 && rawVal < 60000) {
      const utcDays = Math.floor(rawVal - 25569);
      const utcValue = utcDays * 86400;
      const dateInfo = new Date(utcValue * 1000);
      return dateInfo.toISOString().split('T')[0];
    }

    const str = String(rawVal).trim();
    if (!str) return '';

    // Match ISO timestamp or YYYY-MM-DD (e.g. 2026-08-15T08:30:00 or 2026-08-15)
    const ymdMatch = str.match(/^(\d{4})[\/\.-](\d{1,2})[\/\.-](\d{1,2})/);
    if (ymdMatch) {
      const year = ymdMatch[1];
      const month = ymdMatch[2].padStart(2, '0');
      const day = ymdMatch[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    // Match DD/MM/YYYY or DD-MM-YYYY
    const dmyMatch = str.match(/^(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{4})/);
    if (dmyMatch) {
      const day = dmyMatch[1].padStart(2, '0');
      const month = dmyMatch[2].padStart(2, '0');
      const year = dmyMatch[3];
      return `${year}-${month}-${day}`;
    }

    const parsedDate = new Date(str);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate.toISOString().split('T')[0];
    }

    return str;
  }

  function getYearMonthString(rawVal) {
    const std = getStandardDateString(rawVal);
    if (std && std.length >= 7) {
      return std.substring(0, 7); // e.g. "2026-08"
    }
    return '';
  }

  function formatExcelDate(rawVal) {
    const std = getStandardDateString(rawVal);
    return std || new Date().toISOString().split('T')[0];
  }

  function normalizeKey(str) {
    return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  function findVal(rowObj, keyCandidates) {
    const keys = Object.keys(rowObj);
    const normCandidates = keyCandidates.map(c => normalizeKey(c));

    // First pass: Exact normalized key match
    for (let key of keys) {
      const normKey = normalizeKey(key);
      if (!normKey) continue;
      
      const val = rowObj[key];
      if (val === undefined || val === null || String(val).trim() === '') continue;

      if (normCandidates.includes(normKey)) {
        return val;
      }
    }

    // Second pass: Substring / Contains match
    for (let key of keys) {
      const normKey = normalizeKey(key);
      if (!normKey) continue;

      const val = rowObj[key];
      if (val === undefined || val === null || String(val).trim() === '') continue;

      for (let cand of normCandidates) {
        if (normKey.includes(cand) || cand.includes(normKey)) {
          return val;
        }
      }
    }

    return null;
  }

  function normalizeHammerName(input) {
    if (!input) return '1 Ton Hammer';
    const str = String(input).trim().toLowerCase();

    // Check 3.5 Ton
    if (str.includes('3.5') || str.includes('3.5t') || str.includes('h5') || str.includes('hammer 5')) return '3.5 Ton Hammer';
    
    // Check 2.5 Ton New
    if (str.includes('2.5') && (str.includes('new') || str.includes('h4') || str.includes('hammer 4'))) return '2.5 Ton (New) Hammer';
    
    // Check 2.5 Ton (Old or General)
    if (str.includes('2.5') || str.includes('h3') || str.includes('hammer 3')) return '2.5 Ton (Old) Hammer';
    
    // Check 1.5 Ton
    if (str.includes('1.5') || str.includes('1.5t') || str.includes('h2') || str.includes('hammer 2')) return '1.5 Ton Hammer';
    
    // Check 1 Ton
    if (str.includes('1 ton') || str.includes('1ton') || str.includes('1t') || str.includes('h1') || str.includes('hammer 1') || str === '1') return '1 Ton Hammer';
    
    // Exact match fallback
    const match = HAMMERS.find(h => h.name.toLowerCase() === str);
    if (match) return match.name;

    return '1 Ton Hammer';
  }

  /* ==========================================================================
     OEE MATHEMATICAL CALCULATIONS ENGINE
     ========================================================================== */
  function calculateOeeRecord(data) {
    let grossShiftMins = parseNum(data.grossShiftMins, 720);
    if (grossShiftMins > 0 && grossShiftMins <= 24) {
      grossShiftMins = grossShiftMins * 60;
    }

    const lunchBreakMins = parseNum(data.lunchBreakMins, 30);
    const teaBreakMins = parseNum(data.teaBreakMins, 30);
    const totalPlannedBreaksMins = lunchBreakMins + teaBreakMins;

    // Planned Net Operating Base Time: 11.0 Hours = 660 Mins
    let netPlannedTimeMins = data.plannedTimeMins ? parseNum(data.plannedTimeMins, 660) : 660;
    if (netPlannedTimeMins > 0 && netPlannedTimeMins <= 24) {
      netPlannedTimeMins = netPlannedTimeMins * 60;
    }
    if (netPlannedTimeMins >= 700) {
      netPlannedTimeMins = Math.max(60, grossShiftMins - totalPlannedBreaksMins); // 660 Mins
    }
    if (netPlannedTimeMins <= 0) {
      netPlannedTimeMins = 660;
    }

    const maintance = parseTimeMins(data.maintanceMins, true);
    const dieRelated = parseTimeMins(data.dieRelatedMins, true);
    const setup = parseTimeMins(data.setupMins, true);
    const noManpower = parseTimeMins(data.noManpowerMins, true);
    const heatingTime = parseTimeMins(data.heatingTimeMins, true);
    const minorStop = parseTimeMins(data.minorStopMins, true);

    const computedTotalDowntime = maintance + dieRelated + setup + noManpower + heatingTime + minorStop;
    
    let totalDowntime = 0;
    if (data.totalDowntimeMins !== undefined && data.totalDowntimeMins !== '' && parseNum(data.totalDowntimeMins, 0) > 0) {
      totalDowntime = parseTimeMins(data.totalDowntimeMins, true);
    } else {
      totalDowntime = computedTotalDowntime;
    }

    let operatingTime = 0;
    if (data.operatingTimeMins !== undefined && data.operatingTimeMins !== '' && parseNum(data.operatingTimeMins, 0) > 0) {
      operatingTime = parseTimeMins(data.operatingTimeMins, true);
      if (operatingTime > netPlannedTimeMins) {
        operatingTime = Math.max(0, netPlannedTimeMins - totalDowntime);
      }
    } else {
      operatingTime = Math.max(0, netPlannedTimeMins - totalDowntime);
    }

    const totalParts = parseNum(data.totalParts, 0);
    const rejects = parseNum(data.rejects, 0);
    
    let goodParts = 0;
    if (data.goodParts !== undefined && data.goodParts !== '' && parseNum(data.goodParts, -1) >= 0) {
      goodParts = parseNum(data.goodParts, 0);
    } else {
      goodParts = Math.max(0, totalParts - rejects);
    }

    let idealCycleSec = parseNum(data.idealCycleSec, 0);
    if (idealCycleSec <= 0) {
      const hammerObj = HAMMERS.find(h => h.name === data.machine);
      idealCycleSec = hammerObj ? hammerObj.defaultCycle : 45;
    } else if (idealCycleSec < 3) {
      // Specified in minutes per piece (e.g. 0.75 min) => convert to seconds (0.75 * 60 = 45 sec)
      idealCycleSec = idealCycleSec * 60;
    } else if (idealCycleSec > 300) {
      // Specified in parts per hour (e.g. 80 pcs/hr) => convert to sec/pc (3600 / 80 = 45 sec)
      idealCycleSec = 3600 / idealCycleSec;
    }

    // 1. Availability Rate (A = Operating Time / Net Planned Time)
    const availability = netPlannedTimeMins > 0 ? Math.min(100.0, Math.max(0, (operatingTime / netPlannedTimeMins) * 100)) : 0;

    // 2. Performance Rate (P = Ideal Production Time / Operating Time)
    const rawIdealMins = (totalParts * idealCycleSec) / 60;
    const idealProdTimeMins = Math.min(operatingTime, rawIdealMins);
    const performance = operatingTime > 0 ? Math.min(100.0, Math.max(0, (idealProdTimeMins / operatingTime) * 100)) : 0;

    // 3. Quality Rate (Q = Good Parts / Total Parts)
    const quality = totalParts > 0 ? Math.min(100.0, Math.max(0, (goodParts / totalParts) * 100)) : 100.0;

    // 4. Overall OEE (OEE = A x P x Q)
    const oee = Math.min(100.0, Math.max(0, (availability / 100) * (performance / 100) * (quality / 100) * 100));

    const cleanPartNo = data.partNumber ? String(data.partNumber).trim() : 'A1#21';

    return {
      ...data,
      partNumber: cleanPartNo,
      grossShiftMins: grossShiftMins,
      lunchBreakMins: lunchBreakMins,
      teaBreakMins: teaBreakMins,
      plannedBreaksMins: totalPlannedBreaksMins,
      plannedTimeMins: netPlannedTimeMins,
      maintanceMins: maintance,
      dieRelatedMins: dieRelated,
      setupMins: setup,
      noManpowerMins: noManpower,
      heatingTimeMins: heatingTime,
      minorStopMins: minorStop,
      totalDowntimeMins: totalDowntime,
      operatingTimeMins: operatingTime,
      totalParts: totalParts,
      goodParts: goodParts,
      rejects: rejects,
      idealCycleSec: idealCycleSec,
      availability: parseFloat(availability.toFixed(1)),
      performance: parseFloat(performance.toFixed(1)),
      quality: parseFloat(quality.toFixed(1)),
      oee: parseFloat(oee.toFixed(1))
    };
  }

  /* ==========================================================================
     DATA STORAGE ENGINE (STARTS CLEAN EMPTY READY FOR USER UPLOADS)
     ========================================================================== */
  function generateDefaultLogs() {
    return []; // Clean empty by default
  }

  function applyViewOnlyMode() {
    window.IS_VIEW_ONLY = true;
    const banner = document.getElementById('viewOnlyTopBanner');
    if (banner) banner.style.display = 'block';

    const hideElement = (id) => {
      const el = document.getElementById(id);
      if (el) el.style.display = 'none';
    };

    hideElement('openManualEntryBtn');
    hideElement('openExcelModalBtn');
    hideElement('clearDemoDataHeaderBtn');
    hideElement('restoreBackupLogsBtn');
    hideElement('pushToCloudDbBtn');

    document.querySelectorAll('.nav-tab[data-tab="entry"], .nav-tab[data-tab="excel"]').forEach(tab => {
      tab.style.display = 'none';
    });

    document.body.classList.add('view-only-active');
  }

  function compressLogsForUrl(logs) {
    if (!Array.isArray(logs)) return [];
    return logs.map(l => [
      l.date || '',
      l.shift || 'Shift A',
      l.machine || '1 Ton Hammer',
      l.partNumber || 'A1#21',
      l.plannedTimeMins || 660,
      l.maintanceMins || 0,
      l.dieRelatedMins || 0,
      l.setupMins || 0,
      l.noManpowerMins || 0,
      l.heatingTimeMins || 0,
      l.minorStopMins || 0,
      l.totalParts || 0,
      l.goodParts || 0,
      l.rejects || 0,
      l.idealCycleSec || 35
    ]);
  }

  function decompressLogsFromUrl(arr) {
    if (!Array.isArray(arr)) return [];
    return arr.map(r => {
      if (Array.isArray(r)) {
        return calculateOeeRecord({
          date: r[0],
          shift: r[1],
          machine: r[2],
          partNumber: r[3],
          plannedTimeMins: r[4],
          maintanceMins: r[5],
          dieRelatedMins: r[6],
          setupMins: r[7],
          noManpowerMins: r[8],
          heatingTimeMins: r[9],
          minorStopMins: r[10],
          totalParts: r[11],
          goodParts: r[12],
          rejects: r[13],
          idealCycleSec: r[14]
        });
      } else if (typeof r === 'object' && r !== null) {
        return calculateOeeRecord(r);
      }
      return null;
    }).filter(Boolean);
  }

  function checkUrlForSharedData() {
    try {
      const search = window.location.search;
      const hash = window.location.hash;

      const isViewMode = search.includes('mode=view') || search.includes('view=1') || hash.startsWith('#viewData=');

      let encodedPayload = null;
      if (hash && hash.startsWith('#data=')) {
        encodedPayload = hash.replace('#data=', '');
      } else if (hash && hash.startsWith('#viewData=')) {
        encodedPayload = hash.replace('#viewData=', '');
      }

      if (encodedPayload) {
        const jsonStr = decodeURIComponent(atob(encodedPayload));
        const parsed = JSON.parse(jsonStr);
        const importedLogs = decompressLogsFromUrl(parsed);

        if (Array.isArray(importedLogs) && importedLogs.length > 0) {
          shiftLogs = importedLogs;
          window.HAS_URL_DATA = true; // Protect URL loaded data from being overwritten by empty cloud files
          if (!isViewMode) saveShiftLogs();
          renderAllViews();
          showToast(`Successfully loaded ${importedLogs.length} shift logs with live data!`, 'success');
        }
      }

      if (isViewMode) {
        applyViewOnlyMode();
        showToast('Opened in Read-Only View Mode.', 'info');
      }
    } catch (e) {
      console.error('Error parsing URL data:', e);
    }
  }

  function copyTextToClipboard(text, elementToSelect = null) {
    if (elementToSelect) {
      try {
        elementToSelect.focus();
        elementToSelect.select();
        if (elementToSelect.setSelectionRange) {
          elementToSelect.setSelectionRange(0, 999999);
        }
      } catch (e) {}
    }

    let success = false;
    try {
      success = document.execCommand('copy');
    } catch (e) {
      success = false;
    }

    if (!success && navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => { success = true; }).catch(() => {});
    }

    return success;
  }

  function generateShareableDataUrl() {
  try {
    // Generate a true live, view-only Supabase dashboard link.
    // No dashboard data is embedded in the URL.
    const liveViewUrl =
      `${window.location.origin}${window.location.pathname}?mode=view`;

    const viewOnlyInput =
      document.getElementById('viewOnlyShareUrlInput');

    const editInput =
      document.getElementById('mobileShareUrlInput');

    const modal =
      document.getElementById('mobileQrModalBackdrop');

    if (viewOnlyInput) {
      viewOnlyInput.value = liveViewUrl;
    }

    // Clear the old snapshot/edit link.
    if (editInput) {
      editInput.value = '';
    }

    if (modal) {
      modal.style.display = 'flex';
    }

    setTimeout(() => {
      if (viewOnlyInput) {
        copyTextToClipboard(liveViewUrl, viewOnlyInput);
      }

      showToast(
        'Live View-Only Link copied successfully!',
        'success'
      );
    }, 100);

  } catch (err) {
    console.error('Error generating live share URL:', err);
    showToast(
      'Error generating live dashboard link.',
      'danger'
    );
  }
}
  function mergeLogArrays(logsA, logsB) {
    const map = new Map();
    const getRecordKey = (l) => {
      if (!l) return null;
      if (l.id) return String(l.id);
      return `${l.date}_${l.machine}_${l.shift}_${l.partNumber}_${l.plannedTimeMins}_${l.goodParts}`;
    };

    (logsA || []).forEach(l => {
      const k = getRecordKey(l);
      if (k) map.set(k, l);
    });

    (logsB || []).forEach(l => {
      const k = getRecordKey(l);
      if (k) map.set(k, l);
    });

    return Array.from(map.values()).sort((a, b) => new Date(b.date) - new Date(a.date));
  }

  function recoverAllPreviousLogs() {
    let recovered = [];
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && (key.includes('oee_shift_logs') || key.includes('shift_logs') || key.includes('backup') || key.includes('oee'))) {
          const raw = localStorage.getItem(key);
          if (raw) {
            try {
              const parsed = JSON.parse(raw);
              if (Array.isArray(parsed) && parsed.length > 0) {
                recovered = mergeLogArrays(recovered, parsed);
              }
            } catch (err) {}
          }
        }
      }
    } catch (e) {
      console.error('Recovery scan error:', e);
    }
    return recovered;
  }

  function wipeAllSystemDataCompletely() {
    try {
      localStorage.clear();
      sessionStorage.clear();
      shiftLogs = [];
      localStorage.setItem('oee_shift_logs_v10', JSON.stringify([]));

      if (firebaseDbRef) {
        firebaseDbRef.set([]);
      }
    } catch (e) {
      console.error('Wipe error:', e);
    }
    renderAllViews();
    showToast('All system data deleted completely! Ready for fresh Excel uploads.', 'info');
  }

    /* ==========================================================================
     CANONICAL AUGUST 2026 MASTER DATASET (VALIDATED & RECONCILED)
     ========================================================================== */
  const MASTER_AUGUST_PRODUCTION_DATA = [
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-63",
        "planned_time_mins":  300,
        "production_qty":  284,
        "good_qty":  283,
        "ideal_cycle_sec":  57.243816254416963
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-79A",
        "planned_time_mins":  360,
        "production_qty":  256,
        "good_qty":  253,
        "ideal_cycle_sec":  53.3596837944664
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-79A",
        "planned_time_mins":  420,
        "production_qty":  485,
        "good_qty":  484,
        "ideal_cycle_sec":  52.066115702479337
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-67",
        "planned_time_mins":  240,
        "production_qty":  96,
        "good_qty":  96,
        "ideal_cycle_sec":  46.875
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-67",
        "planned_time_mins":  600,
        "production_qty":  635,
        "good_qty":  635,
        "ideal_cycle_sec":  56.69291338582677
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "S8-08",
        "planned_time_mins":  60,
        "production_qty":  8,
        "good_qty":  8,
        "ideal_cycle_sec":  225
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "S8-08",
        "planned_time_mins":  480,
        "production_qty":  52,
        "good_qty":  52,
        "ideal_cycle_sec":  207.69230769230771
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "S8-03",
        "planned_time_mins":  180,
        "production_qty":  13,
        "good_qty":  13,
        "ideal_cycle_sec":  276.92307692307691
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "S8-03",
        "planned_time_mins":  660,
        "production_qty":  64,
        "good_qty":  64,
        "ideal_cycle_sec":  253.125
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "S8-03",
        "planned_time_mins":  90,
        "production_qty":  9,
        "good_qty":  9,
        "ideal_cycle_sec":  200
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-02",
        "planned_time_mins":  450,
        "production_qty":  2000,
        "good_qty":  2000,
        "ideal_cycle_sec":  13.5
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-38B",
        "planned_time_mins":  120,
        "production_qty":  121,
        "good_qty":  120,
        "ideal_cycle_sec":  30
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-38B",
        "planned_time_mins":  180,
        "production_qty":  210,
        "good_qty":  210,
        "ideal_cycle_sec":  42.857142857142861
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-144",
        "planned_time_mins":  480,
        "production_qty":  505,
        "good_qty":  505,
        "ideal_cycle_sec":  44.554455445544555
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-144",
        "planned_time_mins":  180,
        "production_qty":  197,
        "good_qty":  196,
        "ideal_cycle_sec":  45.918367346938773
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-98",
        "planned_time_mins":  480,
        "production_qty":  515,
        "good_qty":  514,
        "ideal_cycle_sec":  49.0272373540856
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-98",
        "planned_time_mins":  60,
        "production_qty":  85,
        "good_qty":  85,
        "ideal_cycle_sec":  42.352941176470594
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-05",
        "planned_time_mins":  360,
        "production_qty":  1996,
        "good_qty":  1996,
        "ideal_cycle_sec":  10.821643286573147
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-01",
        "planned_time_mins":  180,
        "production_qty":  953,
        "good_qty":  953,
        "ideal_cycle_sec":  11.332633788037775
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-100",
        "planned_time_mins":  60,
        "production_qty":  85,
        "good_qty":  85,
        "ideal_cycle_sec":  42.352941176470594
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-100",
        "planned_time_mins":  360,
        "production_qty":  277,
        "good_qty":  277,
        "ideal_cycle_sec":  48.736462093862812
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-104",
        "planned_time_mins":  300,
        "production_qty":  227,
        "good_qty":  225,
        "ideal_cycle_sec":  48
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A4-19",
        "planned_time_mins":  540,
        "production_qty":  2995,
        "good_qty":  2995,
        "ideal_cycle_sec":  10.818030050083472
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-104",
        "planned_time_mins":  120,
        "production_qty":  23,
        "good_qty":  23,
        "ideal_cycle_sec":  39.130434782608695
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-104",
        "planned_time_mins":  180,
        "production_qty":  100,
        "good_qty":  100,
        "ideal_cycle_sec":  54
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-15A",
        "planned_time_mins":  480,
        "production_qty":  297,
        "good_qty":  293,
        "ideal_cycle_sec":  61.4334470989761
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-15A",
        "planned_time_mins":  660,
        "production_qty":  459,
        "good_qty":  455,
        "ideal_cycle_sec":  69.230769230769226
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-165",
        "planned_time_mins":  300,
        "production_qty":  350,
        "good_qty":  350,
        "ideal_cycle_sec":  41.142857142857146
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-197",
        "planned_time_mins":  360,
        "production_qty":  452,
        "good_qty":  452,
        "ideal_cycle_sec":  39.823008849557525
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-197",
        "planned_time_mins":  660,
        "production_qty":  993,
        "good_qty":  993,
        "ideal_cycle_sec":  38.066465256797585
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-69",
        "planned_time_mins":  660,
        "production_qty":  1790,
        "good_qty":  1790,
        "ideal_cycle_sec":  22.122905027932962
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-197",
        "planned_time_mins":  660,
        "production_qty":  1045,
        "good_qty":  1045,
        "ideal_cycle_sec":  34.449760765550238
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-197",
        "planned_time_mins":  240,
        "production_qty":  317,
        "good_qty":  317,
        "ideal_cycle_sec":  45.42586750788643
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-166",
        "planned_time_mins":  420,
        "production_qty":  355,
        "good_qty":  355,
        "ideal_cycle_sec":  48.169014084507047
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-25",
        "planned_time_mins":  660,
        "production_qty":  472,
        "good_qty":  472,
        "ideal_cycle_sec":  66.737288135593218
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-25",
        "planned_time_mins":  420,
        "production_qty":  278,
        "good_qty":  278,
        "ideal_cycle_sec":  61.510791366906481
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-135A",
        "planned_time_mins":  240,
        "production_qty":  245,
        "good_qty":  245,
        "ideal_cycle_sec":  36.734693877551024
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-135A",
        "planned_time_mins":  660,
        "production_qty":  851,
        "good_qty":  849,
        "ideal_cycle_sec":  39.575971731448767
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-135A",
        "planned_time_mins":  660,
        "production_qty":  900,
        "good_qty":  900,
        "ideal_cycle_sec":  38
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-135A",
        "planned_time_mins":  30,
        "production_qty":  22,
        "good_qty":  22,
        "ideal_cycle_sec":  81.818181818181813
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-69",
        "planned_time_mins":  390,
        "production_qty":  1236,
        "good_qty":  1236,
        "ideal_cycle_sec":  18.932038834951459
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-121",
        "planned_time_mins":  240,
        "production_qty":  181,
        "good_qty":  180,
        "ideal_cycle_sec":  30
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-121",
        "planned_time_mins":  240,
        "production_qty":  220,
        "good_qty":  220,
        "ideal_cycle_sec":  36.81818181818182
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-66",
        "planned_time_mins":  420,
        "production_qty":  467,
        "good_qty":  466,
        "ideal_cycle_sec":  50.214592274678111
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-66",
        "planned_time_mins":  360,
        "production_qty":  341,
        "good_qty":  337,
        "ideal_cycle_sec":  53.412462908011868
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-145",
        "planned_time_mins":  300,
        "production_qty":  213,
        "good_qty":  213,
        "ideal_cycle_sec":  59.154929577464792
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-145",
        "planned_time_mins":  360,
        "production_qty":  293,
        "good_qty":  293,
        "ideal_cycle_sec":  67.576791808873722
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "D2-06",
        "planned_time_mins":  300,
        "production_qty":  330,
        "good_qty":  330,
        "ideal_cycle_sec":  43.63636363636364
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "D2-06",
        "planned_time_mins":  660,
        "production_qty":  1170,
        "good_qty":  1170,
        "ideal_cycle_sec":  33.846153846153847
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "IPG",
        "planned_time_mins":  600,
        "production_qty":  140,
        "good_qty":  140,
        "ideal_cycle_sec":  128.57142857142856
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-21",
        "planned_time_mins":  60,
        "production_qty":  13,
        "good_qty":  13,
        "ideal_cycle_sec":  69.230769230769226
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-67",
        "planned_time_mins":  660,
        "production_qty":  503,
        "good_qty":  502,
        "ideal_cycle_sec":  62.749003984063741
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-98",
        "planned_time_mins":  660,
        "production_qty":  810,
        "good_qty":  810,
        "ideal_cycle_sec":  48.888888888888886
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A4-19",
        "planned_time_mins":  600,
        "production_qty":  2675,
        "good_qty":  2675,
        "ideal_cycle_sec":  13.457943925233645
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-191",
        "planned_time_mins":  60,
        "production_qty":  80,
        "good_qty":  80,
        "ideal_cycle_sec":  45
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-191",
        "planned_time_mins":  600,
        "production_qty":  320,
        "good_qty":  319,
        "ideal_cycle_sec":  62.068965517241381
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-95",
        "planned_time_mins":  60,
        "production_qty":  61,
        "good_qty":  60,
        "ideal_cycle_sec":  60
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-95",
        "planned_time_mins":  660,
        "production_qty":  880,
        "good_qty":  878,
        "ideal_cycle_sec":  45.102505694760822
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "A1-95",
        "planned_time_mins":  60,
        "production_qty":  59,
        "good_qty":  59,
        "ideal_cycle_sec":  30.508474576271183
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "1 Ton Hammer",
        "part_number":  "W1-72A",
        "planned_time_mins":  660,
        "production_qty":  652,
        "good_qty":  651,
        "ideal_cycle_sec":  52.534562211981566
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-115A",
        "planned_time_mins":  660,
        "production_qty":  896,
        "good_qty":  894,
        "ideal_cycle_sec":  40.26845637583893
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-162",
        "planned_time_mins":  600,
        "production_qty":  700,
        "good_qty":  700,
        "ideal_cycle_sec":  43.714285714285715
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-22",
        "planned_time_mins":  60,
        "production_qty":  35,
        "good_qty":  31,
        "ideal_cycle_sec":  58.064516129032256
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-01",
        "planned_time_mins":  660,
        "production_qty":  227,
        "good_qty":  226,
        "ideal_cycle_sec":  87.610619469026545
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-01",
        "planned_time_mins":  360,
        "production_qty":  727,
        "good_qty":  726,
        "ideal_cycle_sec":  29.75206611570248
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-63A",
        "planned_time_mins":  300,
        "production_qty":  552,
        "good_qty":  552,
        "ideal_cycle_sec":  32.608695652173914
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-63A",
        "planned_time_mins":  240,
        "production_qty":  448,
        "good_qty":  448,
        "ideal_cycle_sec":  32.142857142857139
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-22",
        "planned_time_mins":  180,
        "production_qty":  166,
        "good_qty":  164,
        "ideal_cycle_sec":  43.90243902439024
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-05",
        "planned_time_mins":  240,
        "production_qty":  240,
        "good_qty":  240,
        "ideal_cycle_sec":  33.75
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-05",
        "planned_time_mins":  660,
        "production_qty":  1405,
        "good_qty":  1404,
        "ideal_cycle_sec":  23.076923076923077
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-05",
        "planned_time_mins":  180,
        "production_qty":  355,
        "good_qty":  352,
        "ideal_cycle_sec":  30.68181818181818
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-110",
        "planned_time_mins":  480,
        "production_qty":  457,
        "good_qty":  457,
        "ideal_cycle_sec":  51.203501094091905
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-110",
        "planned_time_mins":  60,
        "production_qty":  43,
        "good_qty":  42,
        "ideal_cycle_sec":  85.714285714285722
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-176",
        "planned_time_mins":  600,
        "production_qty":  365,
        "good_qty":  365,
        "ideal_cycle_sec":  66.575342465753423
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-176",
        "planned_time_mins":  540,
        "production_qty":  440,
        "good_qty":  438,
        "ideal_cycle_sec":  61.643835616438366
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-174",
        "planned_time_mins":  120,
        "production_qty":  50,
        "good_qty":  50,
        "ideal_cycle_sec":  36
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-174",
        "planned_time_mins":  540,
        "production_qty":  361,
        "good_qty":  361,
        "ideal_cycle_sec":  67.313019390581715
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-22",
        "planned_time_mins":  120,
        "production_qty":  31,
        "good_qty":  28,
        "ideal_cycle_sec":  64.285714285714278
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-22",
        "planned_time_mins":  660,
        "production_qty":  690,
        "good_qty":  680,
        "ideal_cycle_sec":  58.235294117647058
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "B3-55",
        "planned_time_mins":  420,
        "production_qty":  464,
        "good_qty":  464,
        "ideal_cycle_sec":  54.310344827586206
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "C1-15",
        "planned_time_mins":  240,
        "production_qty":  135,
        "good_qty":  135,
        "ideal_cycle_sec":  60
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "C1-15",
        "planned_time_mins":  540,
        "production_qty":  420,
        "good_qty":  419,
        "ideal_cycle_sec":  69.451073985680182
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-182",
        "planned_time_mins":  120,
        "production_qty":  82,
        "good_qty":  82,
        "ideal_cycle_sec":  43.90243902439024
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-182",
        "planned_time_mins":  180,
        "production_qty":  220,
        "good_qty":  220,
        "ideal_cycle_sec":  49.090909090909093
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A8-03",
        "planned_time_mins":  420,
        "production_qty":  201,
        "good_qty":  200,
        "ideal_cycle_sec":  49.5
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "C1-16",
        "planned_time_mins":  60,
        "production_qty":  40,
        "good_qty":  40,
        "ideal_cycle_sec":  90
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-185",
        "planned_time_mins":  650,
        "production_qty":  628,
        "good_qty":  625,
        "ideal_cycle_sec":  56.64
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-185",
        "planned_time_mins":  10,
        "production_qty":  14,
        "good_qty":  14,
        "ideal_cycle_sec":  42.857142857142861
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-17",
        "planned_time_mins":  480,
        "production_qty":  400,
        "good_qty":  400,
        "ideal_cycle_sec":  58.5
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-41A",
        "planned_time_mins":  180,
        "production_qty":  258,
        "good_qty":  256,
        "ideal_cycle_sec":  42.1875
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-41A",
        "planned_time_mins":  240,
        "production_qty":  99,
        "good_qty":  99,
        "ideal_cycle_sec":  45.454545454545453
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "C1-10",
        "planned_time_mins":  420,
        "production_qty":  300,
        "good_qty":  299,
        "ideal_cycle_sec":  60.200668896321069
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "C1-04",
        "planned_time_mins":  360,
        "production_qty":  300,
        "good_qty":  297,
        "ideal_cycle_sec":  54.545454545454547
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-08",
        "planned_time_mins":  300,
        "production_qty":  300,
        "good_qty":  300,
        "ideal_cycle_sec":  60
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-150",
        "planned_time_mins":  660,
        "production_qty":  600,
        "good_qty":  600,
        "ideal_cycle_sec":  57
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-150",
        "planned_time_mins":  660,
        "production_qty":  900,
        "good_qty":  900,
        "ideal_cycle_sec":  36
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-180",
        "planned_time_mins":  120,
        "production_qty":  46,
        "good_qty":  46,
        "ideal_cycle_sec":  39.130434782608695
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-73",
        "planned_time_mins":  540,
        "production_qty":  504,
        "good_qty":  504,
        "ideal_cycle_sec":  57.142857142857139
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "V1-03",
        "planned_time_mins":  660,
        "production_qty":  797,
        "good_qty":  790,
        "ideal_cycle_sec":  45.569620253164558
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "V1-03",
        "planned_time_mins":  180,
        "production_qty":  185,
        "good_qty":  184,
        "ideal_cycle_sec":  58.695652173913047
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-120",
        "planned_time_mins":  360,
        "production_qty":  400,
        "good_qty":  400,
        "ideal_cycle_sec":  54
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-180",
        "planned_time_mins":  120,
        "production_qty":  85,
        "good_qty":  85,
        "ideal_cycle_sec":  42.352941176470594
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-180",
        "planned_time_mins":  420,
        "production_qty":  369,
        "good_qty":  368,
        "ideal_cycle_sec":  58.695652173913047
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-65",
        "planned_time_mins":  240,
        "production_qty":  220,
        "good_qty":  219,
        "ideal_cycle_sec":  49.315068493150683
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A4-18",
        "planned_time_mins":  660,
        "production_qty":  505,
        "good_qty":  505,
        "ideal_cycle_sec":  54.653465346534652
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-65",
        "planned_time_mins":  660,
        "production_qty":  1100,
        "good_qty":  1100,
        "ideal_cycle_sec":  36
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-65",
        "planned_time_mins":  180,
        "production_qty":  189,
        "good_qty":  189,
        "ideal_cycle_sec":  47.619047619047613
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-21",
        "planned_time_mins":  120,
        "production_qty":  7,
        "good_qty":  7,
        "ideal_cycle_sec":  85.714285714285722
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "C1-06",
        "planned_time_mins":  360,
        "production_qty":  200,
        "good_qty":  199,
        "ideal_cycle_sec":  67.839195979899486
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-164",
        "planned_time_mins":  240,
        "production_qty":  143,
        "good_qty":  143,
        "ideal_cycle_sec":  62.93706293706294
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-16",
        "planned_time_mins":  420,
        "production_qty":  380,
        "good_qty":  380,
        "ideal_cycle_sec":  66.31578947368422
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-16",
        "planned_time_mins":  120,
        "production_qty":  120,
        "good_qty":  120,
        "ideal_cycle_sec":  60
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-140",
        "planned_time_mins":  540,
        "production_qty":  304,
        "good_qty":  304,
        "ideal_cycle_sec":  76.973684210526315
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-41",
        "planned_time_mins":  600,
        "production_qty":  500,
        "good_qty":  500,
        "ideal_cycle_sec":  72
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-107",
        "planned_time_mins":  60,
        "production_qty":  24,
        "good_qty":  24,
        "ideal_cycle_sec":  37.5
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-107",
        "planned_time_mins":  420,
        "production_qty":  477,
        "good_qty":  475,
        "ideal_cycle_sec":  53.05263157894737
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-107A",
        "planned_time_mins":  240,
        "production_qty":  150,
        "good_qty":  150,
        "ideal_cycle_sec":  54
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-107A",
        "planned_time_mins":  300,
        "production_qty":  274,
        "good_qty":  274,
        "ideal_cycle_sec":  45.98540145985401
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-06",
        "planned_time_mins":  360,
        "production_qty":  435,
        "good_qty":  435,
        "ideal_cycle_sec":  33.103448275862071
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-06",
        "planned_time_mins":  660,
        "production_qty":  1664,
        "good_qty":  1659,
        "ideal_cycle_sec":  23.869801084990957
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "A1-06",
        "planned_time_mins":  240,
        "production_qty":  415,
        "good_qty":  414,
        "ideal_cycle_sec":  26.086956521739129
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "1.5 Ton Hammer",
        "part_number":  "W1-21",
        "planned_time_mins":  420,
        "production_qty":  610,
        "good_qty":  610,
        "ideal_cycle_sec":  35.409836065573771
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-181",
        "planned_time_mins":  180,
        "production_qty":  48,
        "good_qty":  48,
        "ideal_cycle_sec":  75
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-92",
        "planned_time_mins":  480,
        "production_qty":  61,
        "good_qty":  61,
        "ideal_cycle_sec":  88.524590163934434
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-92",
        "planned_time_mins":  504,
        "production_qty":  39,
        "good_qty":  39,
        "ideal_cycle_sec":  221.53846153846155
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-92",
        "planned_time_mins":  120,
        "production_qty":  5,
        "good_qty":  5,
        "ideal_cycle_sec":  240
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-49",
        "planned_time_mins":  660,
        "production_qty":  271,
        "good_qty":  270,
        "ideal_cycle_sec":  56.666666666666664
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-49",
        "planned_time_mins":  540,
        "production_qty":  369,
        "good_qty":  369,
        "ideal_cycle_sec":  58.536585365853654
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-15",
        "planned_time_mins":  120,
        "production_qty":  32,
        "good_qty":  32,
        "ideal_cycle_sec":  56.25
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-15",
        "planned_time_mins":  660,
        "production_qty":  255,
        "good_qty":  255,
        "ideal_cycle_sec":  67.058823529411768
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-92",
        "planned_time_mins":  660,
        "production_qty":  44,
        "good_qty":  44,
        "ideal_cycle_sec":  163.63636363636363
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-92",
        "planned_time_mins":  660,
        "production_qty":  42,
        "good_qty":  42,
        "ideal_cycle_sec":  171.42857142857144
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-92",
        "planned_time_mins":  660,
        "production_qty":  68,
        "good_qty":  68,
        "ideal_cycle_sec":  185.29411764705884
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-92",
        "planned_time_mins":  180,
        "production_qty":  14,
        "good_qty":  14,
        "ideal_cycle_sec":  192.85714285714286
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-08",
        "planned_time_mins":  480,
        "production_qty":  15,
        "good_qty":  15,
        "ideal_cycle_sec":  240
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-08",
        "planned_time_mins":  660,
        "production_qty":  39,
        "good_qty":  38,
        "ideal_cycle_sec":  331.57894736842104
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-08",
        "planned_time_mins":  180,
        "production_qty":  6,
        "good_qty":  6,
        "ideal_cycle_sec":  300
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-03",
        "planned_time_mins":  480,
        "production_qty":  30,
        "good_qty":  30,
        "ideal_cycle_sec":  240
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-03",
        "planned_time_mins":  660,
        "production_qty":  56,
        "good_qty":  56,
        "ideal_cycle_sec":  257.14285714285711
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-227",
        "planned_time_mins":  660,
        "production_qty":  46,
        "good_qty":  46,
        "ideal_cycle_sec":  221.7391304347826
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-227",
        "planned_time_mins":  660,
        "production_qty":  105,
        "good_qty":  105,
        "ideal_cycle_sec":  171.42857142857144
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "G4-02",
        "planned_time_mins":  660,
        "production_qty":  106,
        "good_qty":  106,
        "ideal_cycle_sec":  203.77358490566039
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "G4-02",
        "planned_time_mins":  660,
        "production_qty":  130,
        "good_qty":  130,
        "ideal_cycle_sec":  180
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "G4-02",
        "planned_time_mins":  300,
        "production_qty":  66,
        "good_qty":  66,
        "ideal_cycle_sec":  163.63636363636363
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-226",
        "planned_time_mins":  360,
        "production_qty":  88,
        "good_qty":  88,
        "ideal_cycle_sec":  71.5909090909091
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-226",
        "planned_time_mins":  300,
        "production_qty":  162,
        "good_qty":  161,
        "ideal_cycle_sec":  67.0807453416149
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-162",
        "planned_time_mins":  360,
        "production_qty":  200,
        "good_qty":  200,
        "ideal_cycle_sec":  54
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-162",
        "planned_time_mins":  120,
        "production_qty":  50,
        "good_qty":  50,
        "ideal_cycle_sec":  72
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-124",
        "planned_time_mins":  540,
        "production_qty":  396,
        "good_qty":  394,
        "ideal_cycle_sec":  45.685279187817258
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-228",
        "planned_time_mins":  120,
        "production_qty":  5,
        "good_qty":  2,
        "ideal_cycle_sec":  450
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-219",
        "planned_time_mins":  540,
        "production_qty":  72,
        "good_qty":  72,
        "ideal_cycle_sec":  75
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-219",
        "planned_time_mins":  420,
        "production_qty":  78,
        "good_qty":  78,
        "ideal_cycle_sec":  92.3076923076923
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-181",
        "planned_time_mins":  240,
        "production_qty":  80,
        "good_qty":  80,
        "ideal_cycle_sec":  67.5
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-181",
        "planned_time_mins":  420,
        "production_qty":  170,
        "good_qty":  170,
        "ideal_cycle_sec":  63.529411764705884
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-09",
        "planned_time_mins":  240,
        "production_qty":  47,
        "good_qty":  47,
        "ideal_cycle_sec":  114.8936170212766
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-09",
        "planned_time_mins":  240,
        "production_qty":  53,
        "good_qty":  53,
        "ideal_cycle_sec":  135.84905660377359
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-07",
        "planned_time_mins":  240,
        "production_qty":  25,
        "good_qty":  25,
        "ideal_cycle_sec":  144
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-09",
        "planned_time_mins":  180,
        "production_qty":  17,
        "good_qty":  17,
        "ideal_cycle_sec":  158.8235294117647
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-09",
        "planned_time_mins":  660,
        "production_qty":  84,
        "good_qty":  84,
        "ideal_cycle_sec":  139.28571428571431
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "S8-07",
        "planned_time_mins":  420,
        "production_qty":  25,
        "good_qty":  25,
        "ideal_cycle_sec":  72
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-163",
        "planned_time_mins":  240,
        "production_qty":  40,
        "good_qty":  40,
        "ideal_cycle_sec":  90
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-163",
        "planned_time_mins":  540,
        "production_qty":  260,
        "good_qty":  260,
        "ideal_cycle_sec":  69.230769230769226
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-08A",
        "planned_time_mins":  120,
        "production_qty":  10,
        "good_qty":  10,
        "ideal_cycle_sec":  90
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-187",
        "planned_time_mins":  120,
        "production_qty":  4,
        "good_qty":  2,
        "ideal_cycle_sec":  300
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "C1-08",
        "planned_time_mins":  540,
        "production_qty":  245,
        "good_qty":  245,
        "ideal_cycle_sec":  66.122448979591837
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "C1-08",
        "planned_time_mins":  240,
        "production_qty":  110,
        "good_qty":  110,
        "ideal_cycle_sec":  65.454545454545453
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "W1-19",
        "planned_time_mins":  420,
        "production_qty":  258,
        "good_qty":  258,
        "ideal_cycle_sec":  55.813953488372093
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "C1-20",
        "planned_time_mins":  660,
        "production_qty":  302,
        "good_qty":  302,
        "ideal_cycle_sec":  59.602649006622514
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "C1-20",
        "planned_time_mins":  660,
        "production_qty":  355,
        "good_qty":  355,
        "ideal_cycle_sec":  60.845070422535208
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "C1-20",
        "planned_time_mins":  420,
        "production_qty":  235,
        "good_qty":  235,
        "ideal_cycle_sec":  45.957446808510639
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-152",
        "planned_time_mins":  240,
        "production_qty":  42,
        "good_qty":  42,
        "ideal_cycle_sec":  42.857142857142861
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-152",
        "planned_time_mins":  660,
        "production_qty":  375,
        "good_qty":  375,
        "ideal_cycle_sec":  50.4
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-152",
        "planned_time_mins":  180,
        "production_qty":  84,
        "good_qty":  84,
        "ideal_cycle_sec":  64.285714285714278
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-28",
        "planned_time_mins":  480,
        "production_qty":  200,
        "good_qty":  200,
        "ideal_cycle_sec":  63
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (Old) Hammer",
        "part_number":  "A1-28",
        "planned_time_mins":  660,
        "production_qty":  301,
        "good_qty":  301,
        "ideal_cycle_sec":  59.800664451827238
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-119",
        "planned_time_mins":  420,
        "production_qty":  245,
        "good_qty":  241,
        "ideal_cycle_sec":  78.423236514522827
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-104",
        "planned_time_mins":  240,
        "production_qty":  90,
        "good_qty":  90,
        "ideal_cycle_sec":  60
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-104",
        "planned_time_mins":  420,
        "production_qty":  310,
        "good_qty":  310,
        "ideal_cycle_sec":  69.677419354838719
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-145",
        "planned_time_mins":  240,
        "production_qty":  194,
        "good_qty":  194,
        "ideal_cycle_sec":  55.670103092783506
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-145",
        "planned_time_mins":  300,
        "production_qty":  204,
        "good_qty":  204,
        "ideal_cycle_sec":  66.1764705882353
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-110A",
        "planned_time_mins":  360,
        "production_qty":  295,
        "good_qty":  295,
        "ideal_cycle_sec":  54.915254237288138
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-110A",
        "planned_time_mins":  300,
        "production_qty":  205,
        "good_qty":  205,
        "ideal_cycle_sec":  79.024390243902445
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-142",
        "planned_time_mins":  360,
        "production_qty":  266,
        "good_qty":  265,
        "ideal_cycle_sec":  61.132075471698116
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-128",
        "planned_time_mins":  660,
        "production_qty":  526,
        "good_qty":  526,
        "ideal_cycle_sec":  65.019011406844115
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-128",
        "planned_time_mins":  540,
        "production_qty":  474,
        "good_qty":  474,
        "ideal_cycle_sec":  62.658227848101269
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-170",
        "planned_time_mins":  120,
        "production_qty":  12,
        "good_qty":  12,
        "ideal_cycle_sec":  75
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-170",
        "planned_time_mins":  660,
        "production_qty":  655,
        "good_qty":  655,
        "ideal_cycle_sec":  60.458015267175568
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-170",
        "planned_time_mins":  540,
        "production_qty":  333,
        "good_qty":  332,
        "ideal_cycle_sec":  97.590361445783131
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-170",
        "planned_time_mins":  20,
        "production_qty":  8,
        "good_qty":  8,
        "ideal_cycle_sec":  75
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-137",
        "planned_time_mins":  100,
        "production_qty":  20,
        "good_qty":  9,
        "ideal_cycle_sec":  66.666666666666671
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-142",
        "planned_time_mins":  660,
        "production_qty":  602,
        "good_qty":  601,
        "ideal_cycle_sec":  59.900166389351085
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-69",
        "planned_time_mins":  660,
        "production_qty":  1010,
        "good_qty":  1010,
        "ideal_cycle_sec":  39.207920792079207
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-69",
        "planned_time_mins":  660,
        "production_qty":  1104,
        "good_qty":  1104,
        "ideal_cycle_sec":  35.869565217391305
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-69",
        "planned_time_mins":  660,
        "production_qty":  929,
        "good_qty":  929,
        "ideal_cycle_sec":  38.751345532831
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-137",
        "planned_time_mins":  540,
        "production_qty":  222,
        "good_qty":  221,
        "ideal_cycle_sec":  65.158371040723978
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-04A",
        "planned_time_mins":  120,
        "production_qty":  44,
        "good_qty":  44,
        "ideal_cycle_sec":  40.909090909090907
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-04A",
        "planned_time_mins":  660,
        "production_qty":  254,
        "good_qty":  252,
        "ideal_cycle_sec":  92.857142857142861
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-50",
        "planned_time_mins":  660,
        "production_qty":  267,
        "good_qty":  265,
        "ideal_cycle_sec":  91.698113207547181
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-08",
        "planned_time_mins":  660,
        "production_qty":  600,
        "good_qty":  600,
        "ideal_cycle_sec":  66
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-08",
        "planned_time_mins":  600,
        "production_qty":  605,
        "good_qty":  605,
        "ideal_cycle_sec":  59.504132231404959
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-189",
        "planned_time_mins":  60,
        "production_qty":  55,
        "good_qty":  55,
        "ideal_cycle_sec":  65.454545454545453
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-189",
        "planned_time_mins":  540,
        "production_qty":  345,
        "good_qty":  345,
        "ideal_cycle_sec":  80.8695652173913
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-06",
        "planned_time_mins":  120,
        "production_qty":  49,
        "good_qty":  48,
        "ideal_cycle_sec":  37.5
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-06",
        "planned_time_mins":  660,
        "production_qty":  551,
        "good_qty":  551,
        "ideal_cycle_sec":  71.869328493647913
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-179",
        "planned_time_mins":  660,
        "production_qty":  391,
        "good_qty":  390,
        "ideal_cycle_sec":  78.461538461538467
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-179",
        "planned_time_mins":  240,
        "production_qty":  109,
        "good_qty":  109,
        "ideal_cycle_sec":  74.311926605504581
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-228",
        "planned_time_mins":  420,
        "production_qty":  195,
        "good_qty":  194,
        "ideal_cycle_sec":  83.505154639175259
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-125",
        "planned_time_mins":  660,
        "production_qty":  467,
        "good_qty":  466,
        "ideal_cycle_sec":  73.390557939914174
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-125",
        "planned_time_mins":  60,
        "production_qty":  36,
        "good_qty":  34,
        "ideal_cycle_sec":  35.294117647058826
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-14",
        "planned_time_mins":  540,
        "production_qty":  251,
        "good_qty":  250,
        "ideal_cycle_sec":  74.4
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-14",
        "planned_time_mins":  60,
        "production_qty":  31,
        "good_qty":  30,
        "ideal_cycle_sec":  60
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-14",
        "planned_time_mins":  660,
        "production_qty":  605,
        "good_qty":  601,
        "ideal_cycle_sec":  61.397670549084857
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-158",
        "planned_time_mins":  660,
        "production_qty":  500,
        "good_qty":  500,
        "ideal_cycle_sec":  68.399999999999991
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-01",
        "planned_time_mins":  660,
        "production_qty":  341,
        "good_qty":  340,
        "ideal_cycle_sec":  84.705882352941188
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-04",
        "planned_time_mins":  660,
        "production_qty":  613,
        "good_qty":  612,
        "ideal_cycle_sec":  64.705882352941174
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-04",
        "planned_time_mins":  660,
        "production_qty":  635,
        "good_qty":  635,
        "ideal_cycle_sec":  62.362204724409445
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-04",
        "planned_time_mins":  660,
        "production_qty":  667,
        "good_qty":  666,
        "ideal_cycle_sec":  59.45945945945946
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-04",
        "planned_time_mins":  120,
        "production_qty":  85,
        "good_qty":  83,
        "ideal_cycle_sec":  86.746987951807228
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "W1-187",
        "planned_time_mins":  540,
        "production_qty":  246,
        "good_qty":  235,
        "ideal_cycle_sec":  88.085106382978722
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-04",
        "planned_time_mins":  660,
        "production_qty":  1166,
        "good_qty":  1166,
        "ideal_cycle_sec":  33.9622641509434
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A1-04",
        "planned_time_mins":  660,
        "production_qty":  1335,
        "good_qty":  1329,
        "ideal_cycle_sec":  29.79683972911964
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-09",
        "planned_time_mins":  660,
        "production_qty":  323,
        "good_qty":  323,
        "ideal_cycle_sec":  83.591331269349851
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "2.5 Ton (New) Hammer",
        "part_number":  "A4-09",
        "planned_time_mins":  660,
        "production_qty":  415,
        "good_qty":  414,
        "ideal_cycle_sec":  82.6086956521739
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-167",
        "planned_time_mins":  420,
        "production_qty":  746,
        "good_qty":  746,
        "ideal_cycle_sec":  33.780160857908847
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-08A",
        "planned_time_mins":  240,
        "production_qty":  361,
        "good_qty":  360,
        "ideal_cycle_sec":  30
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-08A",
        "planned_time_mins":  660,
        "production_qty":  1298,
        "good_qty":  1297,
        "ideal_cycle_sec":  27.756360832690824
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-08A",
        "planned_time_mins":  660,
        "production_qty":  1584,
        "good_qty":  1580,
        "ideal_cycle_sec":  25.063291139240505
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-08A",
        "planned_time_mins":  660,
        "production_qty":  1231,
        "good_qty":  1229,
        "ideal_cycle_sec":  22.701383238405207
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-27",
        "planned_time_mins":  660,
        "production_qty":  784,
        "good_qty":  784,
        "ideal_cycle_sec":  39.030612244897959
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-27",
        "planned_time_mins":  240,
        "production_qty":  216,
        "good_qty":  216,
        "ideal_cycle_sec":  50
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "W1-179",
        "planned_time_mins":  420,
        "production_qty":  141,
        "good_qty":  140,
        "ideal_cycle_sec":  45
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "W1-179",
        "planned_time_mins":  180,
        "production_qty":  59,
        "good_qty":  59,
        "ideal_cycle_sec":  45.762711864406775
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-19",
        "planned_time_mins":  480,
        "production_qty":  553,
        "good_qty":  553,
        "ideal_cycle_sec":  42.31464737793852
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-19",
        "planned_time_mins":  660,
        "production_qty":  870,
        "good_qty":  867,
        "ideal_cycle_sec":  45.674740484429066
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-19",
        "planned_time_mins":  660,
        "production_qty":  1123,
        "good_qty":  1122,
        "ideal_cycle_sec":  35.294117647058826
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-19",
        "planned_time_mins":  360,
        "production_qty":  454,
        "good_qty":  454,
        "ideal_cycle_sec":  43.612334801762117
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-148",
        "planned_time_mins":  300,
        "production_qty":  255,
        "good_qty":  255,
        "ideal_cycle_sec":  38.82352941176471
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-148",
        "planned_time_mins":  660,
        "production_qty":  745,
        "good_qty":  743,
        "ideal_cycle_sec":  49.663526244952891
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-34",
        "planned_time_mins":  660,
        "production_qty":  345,
        "good_qty":  344,
        "ideal_cycle_sec":  60.174418604651166
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-34",
        "planned_time_mins":  660,
        "production_qty":  421,
        "good_qty":  421,
        "ideal_cycle_sec":  68.408551068883611
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-34",
        "planned_time_mins":  60,
        "production_qty":  54,
        "good_qty":  52,
        "ideal_cycle_sec":  69.230769230769226
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-35",
        "planned_time_mins":  600,
        "production_qty":  270,
        "good_qty":  270,
        "ideal_cycle_sec":  83.333333333333329
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "W1-103",
        "planned_time_mins":  120,
        "production_qty":  32,
        "good_qty":  32,
        "ideal_cycle_sec":  56.25
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-16",
        "planned_time_mins":  540,
        "production_qty":  652,
        "good_qty":  650,
        "ideal_cycle_sec":  49.846153846153847
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-16",
        "planned_time_mins":  660,
        "production_qty":  321,
        "good_qty":  320,
        "ideal_cycle_sec":  56.25
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-16",
        "planned_time_mins":  660,
        "production_qty":  0,
        "good_qty":  0,
        "ideal_cycle_sec":  0
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-16",
        "planned_time_mins":  660,
        "production_qty":  0,
        "good_qty":  0,
        "ideal_cycle_sec":  0
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-16",
        "planned_time_mins":  660,
        "production_qty":  0,
        "good_qty":  0,
        "ideal_cycle_sec":  0
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-16",
        "planned_time_mins":  660,
        "production_qty":  1062,
        "good_qty":  1060,
        "ideal_cycle_sec":  25.471698113207548
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "W1-161",
        "planned_time_mins":  660,
        "production_qty":  347,
        "good_qty":  345,
        "ideal_cycle_sec":  70.434782608695656
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-27C",
        "planned_time_mins":  660,
        "production_qty":  995,
        "good_qty":  994,
        "ideal_cycle_sec":  39.839034205231393
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-27C",
        "planned_time_mins":  660,
        "production_qty":  1175,
        "good_qty":  1171,
        "ideal_cycle_sec":  29.205807002561912
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-27C",
        "planned_time_mins":  660,
        "production_qty":  994,
        "good_qty":  994,
        "ideal_cycle_sec":  30.784708249496983
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-14",
        "planned_time_mins":  660,
        "production_qty":  901,
        "good_qty":  900,
        "ideal_cycle_sec":  40
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-14",
        "planned_time_mins":  360,
        "production_qty":  608,
        "good_qty":  605,
        "ideal_cycle_sec":  35.702479338842977
    },
    {
        "date":  "2026-08-15",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-19",
        "planned_time_mins":  300,
        "production_qty":  435,
        "good_qty":  435,
        "ideal_cycle_sec":  41.379310344827587
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "G4-05",
        "planned_time_mins":  660,
        "production_qty":  93,
        "good_qty":  93,
        "ideal_cycle_sec":  174.19354838709677
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "G4-05",
        "planned_time_mins":  660,
        "production_qty":  100,
        "good_qty":  100,
        "ideal_cycle_sec":  162
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "G4-02",
        "planned_time_mins":  660,
        "production_qty":  91,
        "good_qty":  91,
        "ideal_cycle_sec":  178.02197802197804
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "G4-02",
        "planned_time_mins":  660,
        "production_qty":  150,
        "good_qty":  150,
        "ideal_cycle_sec":  192
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "G4-02",
        "planned_time_mins":  60,
        "production_qty":  19,
        "good_qty":  19,
        "ideal_cycle_sec":  126.31578947368421
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-08A",
        "planned_time_mins":  600,
        "production_qty":  1506,
        "good_qty":  1506,
        "ideal_cycle_sec":  23.904382470119522
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A1-08A",
        "planned_time_mins":  600,
        "production_qty":  1497,
        "good_qty":  1497,
        "ideal_cycle_sec":  24.048096192384769
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-49",
        "planned_time_mins":  60,
        "production_qty":  49,
        "good_qty":  48,
        "ideal_cycle_sec":  75
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-49",
        "planned_time_mins":  660,
        "production_qty":  576,
        "good_qty":  575,
        "ideal_cycle_sec":  59.478260869565219
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "W1-244",
        "planned_time_mins":  540,
        "production_qty":  150,
        "good_qty":  149,
        "ideal_cycle_sec":  60.4026845637584
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift B",
        "hammer":  "3.5 Ton Hammer",
        "part_number":  "A4-33",
        "planned_time_mins":  120,
        "production_qty":  7,
        "good_qty":  7,
        "ideal_cycle_sec":  85.714285714285722
    }
]
;

  const MASTER_AUGUST_QUALITY_DATA = [
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#158",
        "inspection_stage":  "Final",
        "inspection_qty":  88,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  201,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#72A",
        "inspection_stage":  "Final",
        "inspection_qty":  415,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  236,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#167",
        "inspection_stage":  "Final",
        "inspection_qty":  790,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#138",
        "inspection_stage":  "Final",
        "inspection_qty":  32,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#182",
        "inspection_stage":  "Final",
        "inspection_qty":  150,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#227",
        "inspection_stage":  "Final",
        "inspection_qty":  67,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  150,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#181",
        "inspection_stage":  "Final",
        "inspection_qty":  80,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#67",
        "inspection_stage":  "Final",
        "inspection_qty":  415,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#44",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  1,
        "rejection_qty":  0,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "ID Lap",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  382,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#79A",
        "inspection_stage":  "Final",
        "inspection_qty":  205,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#38B",
        "inspection_stage":  "Final",
        "inspection_qty":  72,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#08",
        "inspection_stage":  "Final",
        "inspection_qty":  165,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  713,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#65",
        "inspection_stage":  "Final",
        "inspection_qty":  51,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#99",
        "inspection_stage":  "Final",
        "inspection_qty":  23,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Chipout",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#135A",
        "inspection_stage":  "Final",
        "inspection_qty":  400,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#09",
        "inspection_stage":  "Final",
        "inspection_qty":  99,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  14,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#72A",
        "inspection_stage":  "Final",
        "inspection_qty":  123,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "V1#04",
        "inspection_stage":  "Final",
        "inspection_qty":  704,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  355,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "W1#141",
        "inspection_stage":  "Final",
        "inspection_qty":  56,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#157",
        "inspection_stage":  "Final",
        "inspection_qty":  398,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-01",
        "shift":  "Shift A",
        "part_number":  "A1#109A",
        "inspection_stage":  "Final",
        "inspection_qty":  118,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  285,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "W1#72A",
        "inspection_stage":  "Final",
        "inspection_qty":  362,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#181",
        "inspection_stage":  "Final",
        "inspection_qty":  86,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  235,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  251,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "V1#03",
        "inspection_stage":  "Final",
        "inspection_qty":  131,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#49",
        "inspection_stage":  "Final",
        "inspection_qty":  151,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  360,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#162",
        "inspection_stage":  "Final",
        "inspection_qty":  699,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#109A",
        "inspection_stage":  "Final",
        "inspection_qty":  34,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#135A",
        "inspection_stage":  "Final",
        "inspection_qty":  400,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#24",
        "inspection_stage":  "Final",
        "inspection_qty":  325,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  125,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "W1#115A",
        "inspection_stage":  "Final",
        "inspection_qty":  507,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "V1#04",
        "inspection_stage":  "Final",
        "inspection_qty":  281,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  122,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Chipout",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#181",
        "inspection_stage":  "Final",
        "inspection_qty":  85,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#158",
        "inspection_stage":  "Final",
        "inspection_qty":  85,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#167",
        "inspection_stage":  "Final",
        "inspection_qty":  85,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  285,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "A1#49",
        "inspection_stage":  "Final",
        "inspection_qty":  155,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "W1#23",
        "inspection_stage":  "Final",
        "inspection_qty":  340,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-03",
        "shift":  "Shift A",
        "part_number":  "W1#10",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#24",
        "inspection_stage":  "Final",
        "inspection_qty":  63,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "V1#03",
        "inspection_stage":  "Final",
        "inspection_qty":  149,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  28,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#09",
        "inspection_stage":  "Final",
        "inspection_qty":  48,
        "rework_qty":  5,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#141",
        "inspection_stage":  "Final",
        "inspection_qty":  145,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  537,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#49",
        "inspection_stage":  "Final",
        "inspection_qty":  176,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  50,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  31,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#167",
        "inspection_stage":  "Final",
        "inspection_qty":  342,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#145",
        "inspection_stage":  "Final",
        "inspection_qty":  103,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#23",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#44",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  1,
        "rejection_qty":  0,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#145",
        "inspection_stage":  "Final",
        "inspection_qty":  117,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#28",
        "inspection_stage":  "Final",
        "inspection_qty":  82,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#171",
        "inspection_stage":  "Final",
        "inspection_qty":  194,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#15A",
        "inspection_stage":  "Final",
        "inspection_qty":  22,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#158",
        "inspection_stage":  "Final",
        "inspection_qty":  78,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#167",
        "inspection_stage":  "Final",
        "inspection_qty":  517,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#161",
        "inspection_stage":  "Final",
        "inspection_qty":  271,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#112",
        "inspection_stage":  "Final",
        "inspection_qty":  350,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#156",
        "inspection_stage":  "Final",
        "inspection_qty":  401,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Lap",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  145,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  50,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#65A",
        "inspection_stage":  "Final",
        "inspection_qty":  411,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  249,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-04",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#156",
        "inspection_stage":  "Final",
        "inspection_qty":  17,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  14,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#02",
        "inspection_stage":  "Final",
        "inspection_qty":  1945,
        "rework_qty":  0,
        "rejection_qty":  4,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  565,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#14",
        "inspection_stage":  "Final",
        "inspection_qty":  237,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#63A",
        "inspection_stage":  "Final",
        "inspection_qty":  985,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#171",
        "inspection_stage":  "Final",
        "inspection_qty":  40,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#49",
        "inspection_stage":  "Final",
        "inspection_qty":  25,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#115A",
        "inspection_stage":  "Final",
        "inspection_qty":  547,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#119",
        "inspection_stage":  "Final",
        "inspection_qty":  31,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#28",
        "inspection_stage":  "Final",
        "inspection_qty":  17,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#99",
        "inspection_stage":  "Final",
        "inspection_qty":  23,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#23",
        "inspection_stage":  "Final",
        "inspection_qty":  161,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  135,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  107,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#141",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#67",
        "inspection_stage":  "Final",
        "inspection_qty":  136,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#41",
        "inspection_stage":  "Final",
        "inspection_qty":  21,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#145",
        "inspection_stage":  "Final",
        "inspection_qty":  92,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  71,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#110A",
        "inspection_stage":  "Final",
        "inspection_qty":  241,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  12,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "W1#65A",
        "inspection_stage":  "Final",
        "inspection_qty":  88,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  194,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Chipout",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#27",
        "inspection_stage":  "Final",
        "inspection_qty":  299,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  290,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#128",
        "inspection_stage":  "Final",
        "inspection_qty":  278,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#119",
        "inspection_stage":  "Final",
        "inspection_qty":  109,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#67",
        "inspection_stage":  "Final",
        "inspection_qty":  260,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-05",
        "shift":  "Shift A",
        "part_number":  "A1#79A",
        "inspection_stage":  "Final",
        "inspection_qty":  288,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#41A",
        "inspection_stage":  "Final",
        "inspection_qty":  97,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  885,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#170",
        "inspection_stage":  "Final",
        "inspection_qty":  275,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#110A",
        "inspection_stage":  "Final",
        "inspection_qty":  66,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  99,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Lap",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#27",
        "inspection_stage":  "Final",
        "inspection_qty":  70,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  319,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#144",
        "inspection_stage":  "Final",
        "inspection_qty":  305,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#145",
        "inspection_stage":  "Final",
        "inspection_qty":  59,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  75,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#99",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  1,
        "rejection_qty":  2,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  51,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#63A",
        "inspection_stage":  "Final",
        "inspection_qty":  3,
        "rework_qty":  0,
        "rejection_qty":  3,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A4#19",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  24,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  56,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#65A",
        "inspection_stage":  "Final",
        "inspection_qty":  17,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  37,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#05",
        "inspection_stage":  "Final",
        "inspection_qty":  1501,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Chopping",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#167",
        "inspection_stage":  "Final",
        "inspection_qty":  275,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#27",
        "inspection_stage":  "Final",
        "inspection_qty":  451,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#49",
        "inspection_stage":  "Final",
        "inspection_qty":  100,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#135A",
        "inspection_stage":  "Final",
        "inspection_qty":  400,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "W1#38B",
        "inspection_stage":  "Final",
        "inspection_qty":  286,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-06",
        "shift":  "Shift A",
        "part_number":  "A1#128",
        "inspection_stage":  "Final",
        "inspection_qty":  260,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  172,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  201,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  28,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  418,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#05",
        "inspection_stage":  "Final",
        "inspection_qty":  497,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#232",
        "inspection_stage":  "Final",
        "inspection_qty":  205,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#145",
        "inspection_stage":  "Final",
        "inspection_qty":  29,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  55,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  20,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#49",
        "inspection_stage":  "Final",
        "inspection_qty":  193,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#01",
        "inspection_stage":  "Final",
        "inspection_qty":  943,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  195,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#128",
        "inspection_stage":  "Final",
        "inspection_qty":  273,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#167",
        "inspection_stage":  "Final",
        "inspection_qty":  386,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#100",
        "inspection_stage":  "Final",
        "inspection_qty":  351,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A4#19",
        "inspection_stage":  "Final",
        "inspection_qty":  1681,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Folding mark",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#170",
        "inspection_stage":  "Final",
        "inspection_qty":  45,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#157",
        "inspection_stage":  "Final",
        "inspection_qty":  356,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  104,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#27",
        "inspection_stage":  "Final",
        "inspection_qty":  193,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#36A",
        "inspection_stage":  "Final",
        "inspection_qty":  44,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  142,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  56,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  229,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#41",
        "inspection_stage":  "Final",
        "inspection_qty":  56,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "W1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  41,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#98",
        "inspection_stage":  "Final",
        "inspection_qty":  171,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-07",
        "shift":  "Shift A",
        "part_number":  "A1#148",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  653,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A4#19",
        "inspection_stage":  "Final",
        "inspection_qty":  1574,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  500,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#109A",
        "inspection_stage":  "Final",
        "inspection_qty":  209,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#110A",
        "inspection_stage":  "Final",
        "inspection_qty":  195,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  191,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#110",
        "inspection_stage":  "Final",
        "inspection_qty":  128,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A4#34",
        "inspection_stage":  "Final",
        "inspection_qty":  81,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "S8#08",
        "inspection_stage":  "Final",
        "inspection_qty":  25,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#148",
        "inspection_stage":  "Final",
        "inspection_qty":  138,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  130,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#171",
        "inspection_stage":  "Final",
        "inspection_qty":  338,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#99",
        "inspection_stage":  "Final",
        "inspection_qty":  93,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#49",
        "inspection_stage":  "Final",
        "inspection_qty":  21,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A4#34",
        "inspection_stage":  "Final",
        "inspection_qty":  165,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  64,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#110",
        "inspection_stage":  "Final",
        "inspection_qty":  44,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#36A",
        "inspection_stage":  "Final",
        "inspection_qty":  21,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#41",
        "inspection_stage":  "Final",
        "inspection_qty":  21,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  13,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#170",
        "inspection_stage":  "Final",
        "inspection_qty":  143,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  16,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  11,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "A1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  119,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-08",
        "shift":  "Shift A",
        "part_number":  "S8#08",
        "inspection_stage":  "Final",
        "inspection_qty":  25,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#170",
        "inspection_stage":  "Final",
        "inspection_qty":  232,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  100,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#176",
        "inspection_stage":  "Final",
        "inspection_qty":  149,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "S8#03",
        "inspection_stage":  "Final",
        "inspection_qty":  33,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#148",
        "inspection_stage":  "Final",
        "inspection_qty":  65,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#110",
        "inspection_stage":  "Final",
        "inspection_qty":  82,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#15A",
        "inspection_stage":  "Final",
        "inspection_qty":  231,
        "rework_qty":  0,
        "rejection_qty":  4,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  153,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A4#34",
        "inspection_stage":  "Final",
        "inspection_qty":  207,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#174",
        "inspection_stage":  "Final",
        "inspection_qty":  120,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#99",
        "inspection_stage":  "Final",
        "inspection_qty":  63,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  7,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#75A",
        "inspection_stage":  "Final",
        "inspection_qty":  9,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#171",
        "inspection_stage":  "Final",
        "inspection_qty":  25,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  128,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#96",
        "inspection_stage":  "Final",
        "inspection_qty":  188,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  360,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#110",
        "inspection_stage":  "Final",
        "inspection_qty":  25,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "C1#15",
        "inspection_stage":  "Final",
        "inspection_qty":  124,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A4#34",
        "inspection_stage":  "Final",
        "inspection_qty":  294,
        "rework_qty":  0,
        "rejection_qty":  4,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#148",
        "inspection_stage":  "Final",
        "inspection_qty":  344,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A4#35",
        "inspection_stage":  "Final",
        "inspection_qty":  100,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  60,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "S8#08",
        "inspection_stage":  "Final",
        "inspection_qty":  10,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#38B",
        "inspection_stage":  "Final",
        "inspection_qty":  16,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "S8#03",
        "inspection_stage":  "Final",
        "inspection_qty":  51,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#137",
        "inspection_stage":  "Final",
        "inspection_qty":  192,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#98",
        "inspection_stage":  "Final",
        "inspection_qty":  251,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "C1#15",
        "inspection_stage":  "Final",
        "inspection_qty":  181,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  8,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#110",
        "inspection_stage":  "Final",
        "inspection_qty":  173,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#144",
        "inspection_stage":  "Final",
        "inspection_qty":  400,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A4#35",
        "inspection_stage":  "Final",
        "inspection_qty":  90,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  63,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  111,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "A1#135A",
        "inspection_stage":  "Final",
        "inspection_qty":  400,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#15A",
        "inspection_stage":  "Final",
        "inspection_qty":  226,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#227",
        "inspection_stage":  "Final",
        "inspection_qty":  40,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#176",
        "inspection_stage":  "Final",
        "inspection_qty":  80,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-10",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "C1#15",
        "inspection_stage":  "Final",
        "inspection_qty":  176,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A1#128",
        "inspection_stage":  "Final",
        "inspection_qty":  150,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A1#69",
        "inspection_stage":  "Final",
        "inspection_qty":  1217,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A4#08",
        "inspection_stage":  "Final",
        "inspection_qty":  268,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  116,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#174",
        "inspection_stage":  "Final",
        "inspection_qty":  74,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#15A",
        "inspection_stage":  "Final",
        "inspection_qty":  40,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A1#165",
        "inspection_stage":  "Final",
        "inspection_qty":  360,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "S8#03",
        "inspection_stage":  "Final",
        "inspection_qty":  2,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#176",
        "inspection_stage":  "Final",
        "inspection_qty":  406,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Dent",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A1#69",
        "inspection_stage":  "Final",
        "inspection_qty":  191,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  21,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Folding",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A4#35",
        "inspection_stage":  "Final",
        "inspection_qty":  106,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A1#119",
        "inspection_stage":  "Final",
        "inspection_qty":  143,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "A4#16",
        "inspection_stage":  "Final",
        "inspection_qty":  345,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#103",
        "inspection_stage":  "Final",
        "inspection_qty":  117,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#227",
        "inspection_stage":  "Final",
        "inspection_qty":  65,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#104",
        "inspection_stage":  "Final",
        "inspection_qty":  20,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#110",
        "inspection_stage":  "Final",
        "inspection_qty":  36,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "C1#15",
        "inspection_stage":  "Final",
        "inspection_qty":  68,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-11",
        "shift":  "Shift A",
        "part_number":  "W1#15A",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  170,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "C1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  315,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#105",
        "inspection_stage":  "Final",
        "inspection_qty":  41,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#176",
        "inspection_stage":  "Final",
        "inspection_qty":  142,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#174",
        "inspection_stage":  "Final",
        "inspection_qty":  14,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A4#08",
        "inspection_stage":  "Final",
        "inspection_qty":  850,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#50",
        "inspection_stage":  "Final",
        "inspection_qty":  162,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  37,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  19,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#189",
        "inspection_stage":  "Final",
        "inspection_qty":  117,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A4#16",
        "inspection_stage":  "Final",
        "inspection_qty":  156,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "B3#55-A105",
        "inspection_stage":  "Final",
        "inspection_qty":  103,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A8#03",
        "inspection_stage":  "Final",
        "inspection_qty":  182,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#67",
        "inspection_stage":  "Final",
        "inspection_qty":  398,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#197",
        "inspection_stage":  "Final",
        "inspection_qty":  921,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Bend",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#41A",
        "inspection_stage":  "Final",
        "inspection_qty":  345,
        "rework_qty":  0,
        "rejection_qty":  3,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  45,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#166",
        "inspection_stage":  "Final",
        "inspection_qty":  350,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "C1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  250,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#185",
        "inspection_stage":  "Final",
        "inspection_qty":  378,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#148",
        "inspection_stage":  "Final",
        "inspection_qty":  232,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "C1#04",
        "inspection_stage":  "Final",
        "inspection_qty":  296,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#50",
        "inspection_stage":  "Final",
        "inspection_qty":  296,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "B3#55-A105",
        "inspection_stage":  "Final",
        "inspection_qty":  349,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#197",
        "inspection_stage":  "Final",
        "inspection_qty":  670,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#67",
        "inspection_stage":  "Final",
        "inspection_qty":  464,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#17",
        "inspection_stage":  "Final",
        "inspection_qty":  131,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#182",
        "inspection_stage":  "Final",
        "inspection_qty":  246,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#159",
        "inspection_stage":  "Final",
        "inspection_qty":  111,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#189",
        "inspection_stage":  "Final",
        "inspection_qty":  67,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "W1#15A",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-12",
        "shift":  "Shift A",
        "part_number":  "A1#69",
        "inspection_stage":  "Final",
        "inspection_qty":  25,
        "rework_qty":  0,
        "rejection_qty":  25,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "B3#55-A105",
        "inspection_stage":  "Final",
        "inspection_qty":  12,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A1#189",
        "inspection_stage":  "Final",
        "inspection_qty":  45,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A4#34",
        "inspection_stage":  "Final",
        "inspection_qty":  50,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#17",
        "inspection_stage":  "Final",
        "inspection_qty":  31,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  18,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#50",
        "inspection_stage":  "Final",
        "inspection_qty":  20,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#185",
        "inspection_stage":  "Final",
        "inspection_qty":  66,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "C1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  40,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A8#03",
        "inspection_stage":  "Final",
        "inspection_qty":  12,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  38,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#105",
        "inspection_stage":  "Final",
        "inspection_qty":  16,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#04A",
        "inspection_stage":  "Final",
        "inspection_qty":  65,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A4#16",
        "inspection_stage":  "Final",
        "inspection_qty":  456,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A1#15",
        "inspection_stage":  "Final",
        "inspection_qty":  82,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A1#25",
        "inspection_stage":  "Final",
        "inspection_qty":  300,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A4#06",
        "inspection_stage":  "Final",
        "inspection_qty":  570,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#197",
        "inspection_stage":  "Final",
        "inspection_qty":  612,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "BEND",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#08",
        "inspection_stage":  "Final",
        "inspection_qty":  120,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  115,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "OVER HEAT",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A1#148",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "UNFILLING",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A4#16",
        "inspection_stage":  "Final",
        "inspection_qty":  887,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A1#159",
        "inspection_stage":  "Final",
        "inspection_qty":  399,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#197",
        "inspection_stage":  "Final",
        "inspection_qty":  623,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  203,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#17",
        "inspection_stage":  "Final",
        "inspection_qty":  38,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#226",
        "inspection_stage":  "Final",
        "inspection_qty":  234,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#04A",
        "inspection_stage":  "Final",
        "inspection_qty":  42,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  173,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A4#06",
        "inspection_stage":  "Final",
        "inspection_qty":  329,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-13",
        "shift":  "Shift A",
        "part_number":  "A1#25",
        "inspection_stage":  "Final",
        "inspection_qty":  450,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#15A",
        "inspection_stage":  "Final",
        "inspection_qty":  31,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#137",
        "inspection_stage":  "Final",
        "inspection_qty":  30,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A1#179",
        "inspection_stage":  "Final",
        "inspection_qty":  403,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#08",
        "inspection_stage":  "Final",
        "inspection_qty":  135,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  719,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#92",
        "inspection_stage":  "Final",
        "inspection_qty":  43,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  58,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#176",
        "inspection_stage":  "Final",
        "inspection_qty":  18,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#105",
        "inspection_stage":  "Final",
        "inspection_qty":  16,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#182",
        "inspection_stage":  "Final",
        "inspection_qty":  65,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "C1#10",
        "inspection_stage":  "Final",
        "inspection_qty":  81,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A1#69",
        "inspection_stage":  "Final",
        "inspection_qty":  520,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#150",
        "inspection_stage":  "Final",
        "inspection_qty":  583,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "C1#04",
        "inspection_stage":  "Final",
        "inspection_qty":  148,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#227",
        "inspection_stage":  "Final",
        "inspection_qty":  50,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#226",
        "inspection_stage":  "Final",
        "inspection_qty":  14,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#124",
        "inspection_stage":  "Final",
        "inspection_qty":  208,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Lap",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A1#69",
        "inspection_stage":  "Final",
        "inspection_qty":  509,
        "rework_qty":  1,
        "rejection_qty":  3,
        "reason":  "U/F \u0026 Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#227",
        "inspection_stage":  "Final",
        "inspection_qty":  35,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#185",
        "inspection_stage":  "Final",
        "inspection_qty":  178,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  511,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#08",
        "inspection_stage":  "Final",
        "inspection_qty":  38,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A1#189",
        "inspection_stage":  "Final",
        "inspection_qty":  100,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#174",
        "inspection_stage":  "Final",
        "inspection_qty":  114,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Chipout",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#219",
        "inspection_stage":  "Final",
        "inspection_qty":  77,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#162",
        "inspection_stage":  "Final",
        "inspection_qty":  103,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#105",
        "inspection_stage":  "Final",
        "inspection_qty":  62,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  37,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "C1#10",
        "inspection_stage":  "Final",
        "inspection_qty":  95,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  18,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A1#15",
        "inspection_stage":  "Final",
        "inspection_qty":  2,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-14",
        "shift":  "Shift A",
        "part_number":  "A4#14",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#124",
        "inspection_stage":  "Final",
        "inspection_qty":  98,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#185",
        "inspection_stage":  "Final",
        "inspection_qty":  80,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "A4#14",
        "inspection_stage":  "Final",
        "inspection_qty":  342,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "A1#180",
        "inspection_stage":  "Final",
        "inspection_qty":  482,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "A1#14",
        "inspection_stage":  "Final",
        "inspection_qty":  442,
        "rework_qty":  76,
        "rejection_qty":  0,
        "reason":  "Mismatch",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  382,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#219",
        "inspection_stage":  "Final",
        "inspection_qty":  86,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "A1#181",
        "inspection_stage":  "Final",
        "inspection_qty":  123,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#150",
        "inspection_stage":  "Final",
        "inspection_qty":  55,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#219",
        "inspection_stage":  "Final",
        "inspection_qty":  55,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "C1#10",
        "inspection_stage":  "Final",
        "inspection_qty":  186,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "C1#04",
        "inspection_stage":  "Final",
        "inspection_qty":  16,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "A1#120",
        "inspection_stage":  "Final",
        "inspection_qty":  263,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#228",
        "inspection_stage":  "Final",
        "inspection_qty":  194,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#174",
        "inspection_stage":  "Final",
        "inspection_qty":  115,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#73",
        "inspection_stage":  "Final",
        "inspection_qty":  128,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#161",
        "inspection_stage":  "Final",
        "inspection_qty":  139,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#124",
        "inspection_stage":  "Final",
        "inspection_qty":  116,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "W1#162",
        "inspection_stage":  "Final",
        "inspection_qty":  41,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "A1#138",
        "inspection_stage":  "Final",
        "inspection_qty":  101,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-17",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  907,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#158",
        "inspection_stage":  "Final",
        "inspection_qty":  177,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#189",
        "inspection_stage":  "Final",
        "inspection_qty":  73,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  200,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "V1#03",
        "inspection_stage":  "Final",
        "inspection_qty":  812,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  95,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#162",
        "inspection_stage":  "Final",
        "inspection_qty":  45,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#121",
        "inspection_stage":  "Final",
        "inspection_qty":  401,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#125",
        "inspection_stage":  "Final",
        "inspection_qty":  237,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#163",
        "inspection_stage":  "Final",
        "inspection_qty":  107,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#01",
        "inspection_stage":  "Final",
        "inspection_qty":  112,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "S8#09",
        "inspection_stage":  "Final",
        "inspection_qty":  51,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#138",
        "inspection_stage":  "Final",
        "inspection_qty":  144,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Dent mark",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#180",
        "inspection_stage":  "Final",
        "inspection_qty":  20,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#150",
        "inspection_stage":  "Final",
        "inspection_qty":  445,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "C1#06",
        "inspection_stage":  "Final",
        "inspection_qty":  175,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Overheat",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#145",
        "inspection_stage":  "Final",
        "inspection_qty":  506,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#65",
        "inspection_stage":  "Final",
        "inspection_qty":  326,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#65",
        "inspection_stage":  "Final",
        "inspection_qty":  164,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "C1#10",
        "inspection_stage":  "Final",
        "inspection_qty":  95,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#161",
        "inspection_stage":  "Final",
        "inspection_qty":  76,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#228",
        "inspection_stage":  "Final",
        "inspection_qty":  13,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#105",
        "inspection_stage":  "Final",
        "inspection_qty":  168,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#01",
        "inspection_stage":  "Final",
        "inspection_qty":  98,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  242,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#163",
        "inspection_stage":  "Final",
        "inspection_qty":  90,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#14",
        "inspection_stage":  "Final",
        "inspection_qty":  230,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A4#18",
        "inspection_stage":  "Final",
        "inspection_qty":  139,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A4#04",
        "inspection_stage":  "Final",
        "inspection_qty":  110,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#174",
        "inspection_stage":  "Final",
        "inspection_qty":  2,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "W1#124",
        "inspection_stage":  "Final",
        "inspection_qty":  1,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-18",
        "shift":  "Shift A",
        "part_number":  "A1#67",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  501,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#14",
        "inspection_stage":  "Final",
        "inspection_qty":  108,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  167,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A4#18",
        "inspection_stage":  "Final",
        "inspection_qty":  162,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#158",
        "inspection_stage":  "Final",
        "inspection_qty":  206,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#65",
        "inspection_stage":  "Final",
        "inspection_qty":  404,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#174",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#150",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "V1#03",
        "inspection_stage":  "Final",
        "inspection_qty":  130,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "S8#09",
        "inspection_stage":  "Final",
        "inspection_qty":  48,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "C1#08",
        "inspection_stage":  "Final",
        "inspection_qty":  318,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#01",
        "inspection_stage":  "Final",
        "inspection_qty":  31,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A4#19",
        "inspection_stage":  "Final",
        "inspection_qty":  1000,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#125",
        "inspection_stage":  "Final",
        "inspection_qty":  74,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  107,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "C1#06",
        "inspection_stage":  "Final",
        "inspection_qty":  38,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  30,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#41",
        "inspection_stage":  "Final",
        "inspection_qty":  498,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A4#04",
        "inspection_stage":  "Final",
        "inspection_qty":  244,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A4#04",
        "inspection_stage":  "Final",
        "inspection_qty":  577,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#66",
        "inspection_stage":  "Final",
        "inspection_qty":  34,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A4#19",
        "inspection_stage":  "Final",
        "inspection_qty":  1390,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#163",
        "inspection_stage":  "Final",
        "inspection_qty":  70,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  44,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  260,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#164",
        "inspection_stage":  "Final",
        "inspection_qty":  140,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A4#18",
        "inspection_stage":  "Final",
        "inspection_qty":  48,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "S8#07",
        "inspection_stage":  "Final",
        "inspection_qty":  14,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "C1#08",
        "inspection_stage":  "Final",
        "inspection_qty":  27,
        "rework_qty":  0,
        "rejection_qty":  7,
        "reason":  "Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#01",
        "inspection_stage":  "Final",
        "inspection_qty":  13,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#04A",
        "inspection_stage":  "Final",
        "inspection_qty":  42,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#162",
        "inspection_stage":  "Final",
        "inspection_qty":  55,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "W1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  215,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-19",
        "shift":  "Shift A",
        "part_number":  "A1#158",
        "inspection_stage":  "Final",
        "inspection_qty":  63,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A4#04",
        "inspection_stage":  "Final",
        "inspection_qty":  109,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#152",
        "inspection_stage":  "Final",
        "inspection_qty":  341,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#138",
        "inspection_stage":  "Final",
        "inspection_qty":  34,
        "rework_qty":  1,
        "rejection_qty":  0,
        "reason":  "U/F",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#125",
        "inspection_stage":  "Final",
        "inspection_qty":  73,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#170",
        "inspection_stage":  "Final",
        "inspection_qty":  101,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  15,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  16,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#01",
        "inspection_stage":  "Final",
        "inspection_qty":  247,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A4#19",
        "inspection_stage":  "Final",
        "inspection_qty":  277,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Lap",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  260,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  835,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "IPG-185X48",
        "inspection_stage":  "Final",
        "inspection_qty":  138,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "S8#07",
        "inspection_stage":  "Final",
        "inspection_qty":  11,
        "rework_qty":  1,
        "rejection_qty":  0,
        "reason":  "Lap",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#150",
        "inspection_stage":  "Final",
        "inspection_qty":  453,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#197",
        "inspection_stage":  "Final",
        "inspection_qty":  136,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "C1#20",
        "inspection_stage":  "Final",
        "inspection_qty":  250,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  400,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A4#04",
        "inspection_stage":  "Final",
        "inspection_qty":  629,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Chipout",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  165,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#138",
        "inspection_stage":  "Final",
        "inspection_qty":  56,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  35,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#107",
        "inspection_stage":  "Final",
        "inspection_qty":  256,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "W1#65",
        "inspection_stage":  "Final",
        "inspection_qty":  207,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "C1#20",
        "inspection_stage":  "Final",
        "inspection_qty":  63,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#170",
        "inspection_stage":  "Final",
        "inspection_qty":  186,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Lap",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "D2#06",
        "inspection_stage":  "Final",
        "inspection_qty":  1500,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#95",
        "inspection_stage":  "Final",
        "inspection_qty":  835,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-20",
        "shift":  "Shift A",
        "part_number":  "A1#67",
        "inspection_stage":  "Final",
        "inspection_qty":  2,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Surface Uneven",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#01",
        "inspection_stage":  "Final",
        "inspection_qty":  39,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#65",
        "inspection_stage":  "Final",
        "inspection_qty":  85,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A4#09",
        "inspection_stage":  "Final",
        "inspection_qty":  321,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#27C",
        "inspection_stage":  "Final",
        "inspection_qty":  266,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A4#04",
        "inspection_stage":  "Final",
        "inspection_qty":  318,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#95",
        "inspection_stage":  "Final",
        "inspection_qty":  161,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#28",
        "inspection_stage":  "Final",
        "inspection_qty":  184,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#125",
        "inspection_stage":  "Final",
        "inspection_qty":  109,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#21",
        "inspection_stage":  "Final",
        "inspection_qty":  87,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#142",
        "inspection_stage":  "Final",
        "inspection_qty":  57,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#107A",
        "inspection_stage":  "Final",
        "inspection_qty":  420,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Chopping",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#08A",
        "inspection_stage":  "Final",
        "inspection_qty":  716,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Lap \u0026 Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#152",
        "inspection_stage":  "Final",
        "inspection_qty":  132,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#16",
        "inspection_stage":  "Final",
        "inspection_qty":  95,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#19",
        "inspection_stage":  "Final",
        "inspection_qty":  112,
        "rework_qty":  0,
        "rejection_qty":  2,
        "reason":  "Deep Crack",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#140",
        "inspection_stage":  "Final",
        "inspection_qty":  52,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#107",
        "inspection_stage":  "Final",
        "inspection_qty":  236,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "W1#22",
        "inspection_stage":  "Final",
        "inspection_qty":  146,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Chip Off",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#04",
        "inspection_stage":  "Final",
        "inspection_qty":  2479,
        "rework_qty":  0,
        "rejection_qty":  1,
        "reason":  "Damage",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#138",
        "inspection_stage":  "Final",
        "inspection_qty":  59,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#152",
        "inspection_stage":  "Final",
        "inspection_qty":  81,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A4#09",
        "inspection_stage":  "Final",
        "inspection_qty":  48,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A1#28",
        "inspection_stage":  "Final",
        "inspection_qty":  123,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "C1#20",
        "inspection_stage":  "Final",
        "inspection_qty":  175,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    },
    {
        "date":  "2026-08-21",
        "shift":  "Shift A",
        "part_number":  "A4#49",
        "inspection_stage":  "Final",
        "inspection_qty":  350,
        "rework_qty":  0,
        "rejection_qty":  0,
        "reason":  "Visual Defect",
        "remarks":  "Master QA Import"
    }
]
;

  function loadShiftLogs() {
    const saved = localStorage.getItem('oee_shift_logs_v10');
    if (saved) {
      try {
        shiftLogs = JSON.parse(saved).map(l => calculateOeeRecord(l));
      } catch (e) {
        shiftLogs = [];
      }
    } else {
      shiftLogs = [];
    }
  }

  /* ==========================================================================
     CENTRAL SUPABASE CLOUD DATABASE, AUTH, REALTIME & RLS ENGINE
     ========================================================================== */
  /* ==========================================================================
     CENTRAL SUPABASE CLOUD DATABASE, AUTH, REALTIME & RLS ENGINE
     ========================================================================== */
  function getSupabaseCredentials() {
    const url =
        SUPABASE_CONFIG.url ||
        window.VITE_SUPABASE_URL ||
        localStorage.getItem('supabase_url') ||
        '';

    const key =
        SUPABASE_CONFIG.key ||
        window.VITE_SUPABASE_PUBLISHABLE_KEY ||
        window.VITE_SUPABASE_ANON_KEY ||
        localStorage.getItem('supabase_key') ||
        '';

    return { url, key };
}

  function initSupabaseClient() {
    const { url, key } = getSupabaseCredentials();
    const pill = document.getElementById('cloudDbStatusPill');

    const urlInput = document.getElementById('cfgSupabaseUrl');
    const keyInput = document.getElementById('cfgSupabaseKey');
    if (urlInput && url) urlInput.value = url;
    if (keyInput && key) keyInput.value = key;

    if (url && key && typeof supabase !== 'undefined') {
      try {
        supabaseClient = supabase.createClient(url, key);
        console.log('Supabase Cloud Database client initialized.');
        if (pill) {
          pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-cloud"></i> 🟢 Live Database Connected';
          pill.style.borderColor = 'var(--primary)';
        }
        initSupabaseAuth();
        initSupabaseRealtimeSubscriptions();
        initSupabasePresence();
        fetchSupabaseShiftLogs();
        return true;
      } catch (err) {
        console.error('Supabase Client Error:', err);
        if (pill) {
          pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-triangle-exclamation text-danger"></i> 🔴 Database Connection Failed';
          pill.style.borderColor = 'var(--danger)';
        }
      }
    }

    if (pill) {
      pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-plug text-warning"></i> 🟡 Set DB Credentials (DB Config)';
      pill.style.borderColor = 'var(--warning)';
    }
    return false;
  }

  async function testSupabaseConnection(url, key) {
    const alertDiv = document.getElementById('cfgConnectionStatusAlert');
    const pill = document.getElementById('cloudDbStatusPill');

    if (!url || !key) {
      if (alertDiv) {
        alertDiv.style.display = 'block';
        alertDiv.style.background = 'rgba(220, 38, 38, 0.12)';
        alertDiv.style.color = '#dc2626';
        alertDiv.style.border = '1px solid #dc2626';
        alertDiv.innerHTML = '🔴 <strong>Database Connection Failed:</strong> Please enter both Supabase Project URL and Publishable Key.';
      }
      return false;
    }

    try {
      const testClient = supabase.createClient(url, key);
      const { error } = await testClient.from('production_data').select('id').limit(1);

      if (error && error.message && (error.message.includes('apiKey') || error.message.includes('JWT') || error.message.includes('invalid'))) {
        throw new Error('Invalid Supabase Publishable Key or URL format.');
      }

      if (alertDiv) {
        alertDiv.style.display = 'block';
        alertDiv.style.background = 'rgba(22, 163, 74, 0.12)';
        alertDiv.style.color = '#16a34a';
        alertDiv.style.border = '1px solid #16a34a';
        alertDiv.innerHTML = '🟢 <strong>Connected to Cloud Database!</strong> Single source of truth active.';
      }

      if (pill) {
        pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-cloud"></i> 🟢 Live Database Connected';
        pill.style.borderColor = 'var(--primary)';
      }

      return true;
    } catch (err) {
      console.error('Supabase Connection Test Failed:', err);
      if (alertDiv) {
        alertDiv.style.display = 'block';
        alertDiv.style.background = 'rgba(220, 38, 38, 0.12)';
        alertDiv.style.color = '#dc2626';
        alertDiv.style.border = '1px solid #dc2626';
        alertDiv.innerHTML = `🔴 <strong>Database Connection Failed:</strong> ${err.message || 'Unable to reach Supabase server.'}`;
      }
      if (pill) {
        pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-triangle-exclamation text-danger"></i> 🔴 Database Connection Failed';
        pill.style.borderColor = 'var(--danger)';
      }
      return false;
    }
  }

  /* Supabase Auth & Role-Based UI Permissions */
  function initSupabaseAuth() {
    if (!supabaseClient) return;

    supabaseClient.auth.getSession().then(({ data: { session } }) => {
      if (session) {
        handleUserLoggedIn(session.user);
      }
    });

    supabaseClient.auth.onAuthStateChange((_event, session) => {
      if (session) {
        handleUserLoggedIn(session.user);
      } else {
        handleUserLoggedOut();
      }
    });
  }

  function handleUserLoggedIn(user) {
    currentAuthUser = user;
    
    supabaseClient
      .from('profiles')
      .select('*')
      .eq('user_id', user.id)
      .maybeSingle()
      .then(({ data }) => {
        if (data) {
          currentUserProfile = data;
          updateUserProfileBadge(data);
          applyRolePermissions(data.role);
          showToast(`Welcome back, ${data.employee_name} (${data.role})!`, 'success');
        } else {
          const empName = user.user_metadata?.employee_name || user.email.split('@')[0];
          const role = user.user_metadata?.role || 'Production';
          const dept = user.user_metadata?.department || 'Production';
          currentUserProfile = { user_id: user.id, employee_name: empName, role, department: dept };
          updateUserProfileBadge(currentUserProfile);
          applyRolePermissions(role);
        }
      });
  }

  function handleUserLoggedOut() {
    currentAuthUser = null;
    currentUserProfile = null;
    const badge = document.getElementById('userProfileBadge');
    const openAuthBtn = document.getElementById('openAuthModalBtn');
    if (badge) badge.style.display = 'none';
    if (openAuthBtn) openAuthBtn.style.display = 'inline-flex';
    applyRolePermissions('Viewer');
    showToast('Logged out of Supabase system.', 'info');
  }

  function updateUserProfileBadge(profile) {
    const badge = document.getElementById('userProfileBadge');
    const openAuthBtn = document.getElementById('openAuthModalBtn');
    const avatar = document.getElementById('userAvatar');
    const empName = document.getElementById('userEmpName');
    const roleDept = document.getElementById('userRoleDept');

    if (badge && profile) {
      if (avatar) avatar.textContent = (profile.employee_name || 'U').charAt(0).toUpperCase();
      if (empName) empName.textContent = profile.employee_name || 'Employee';
      if (roleDept) roleDept.textContent = `${profile.role || 'User'} | ${profile.department || 'Plant'}`;
      badge.style.display = 'inline-flex';
      if (openAuthBtn) openAuthBtn.style.display = 'none';
    }
  }

  function applyRolePermissions(role) {
    const isViewerOrMgmt = role === 'Viewer' || role === 'Management';
    
    const manualBtn = document.getElementById('openManualEntryBtn');
    const excelBtn = document.getElementById('openExcelModalBtn');
    const clearBtn = document.getElementById('clearDemoDataHeaderBtn');
    
    if (isViewerOrMgmt) {
      if (manualBtn) manualBtn.style.display = 'none';
      if (excelBtn) excelBtn.style.display = 'none';
      if (clearBtn) clearBtn.style.display = 'none';
      window.IS_VIEW_ONLY = true;
    } else {
      if (manualBtn) manualBtn.style.display = 'inline-flex';
      if (excelBtn) excelBtn.style.display = 'inline-flex';
      if (clearBtn) clearBtn.style.display = role === 'Admin' ? 'inline-flex' : 'none';
      window.IS_VIEW_ONLY = false;
    }
  }

  /* Supabase Realtime Subscriptions & Presence */
  function initSupabaseRealtimeSubscriptions() {
    if (!supabaseClient) return;

    const pill = document.getElementById('cloudDbStatusPill');

    const channel = supabaseClient.channel('schema-db-changes');

    channel
      .on('postgres_changes', { event: '*', schema: 'public', table: 'production_data' }, () => fetchSupabaseShiftLogs())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'quality_data' }, () => fetchSupabaseShiftLogs())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'downtime_data' }, () => fetchSupabaseShiftLogs())
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') {
          if (pill) {
            pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-cloud"></i> 🟢 Live Database';
            pill.style.borderColor = 'var(--primary)';
          }
        } else if (status === 'TIMED_OUT' || status === 'CLOSED' || status === 'CHANNEL_ERROR') {
          if (pill) {
            pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-spinner fa-spin text-warning"></i> 🟠 Reconnecting';
            pill.style.borderColor = 'var(--warning)';
          }
        }
      });
  }

  function initSupabasePresence() {
    if (!supabaseClient) return;

    presenceChannel = supabaseClient.channel('online-users', {
      config: { presence: { key: currentAuthUser ? currentAuthUser.id : 'anon_' + Math.random().toString(36).substr(2, 6) } }
    });

    presenceChannel
      .on('presence', { event: 'sync' }, () => {
        const state = presenceChannel.presenceState();
        const userCount = Object.keys(state).length;
        const countSpan = document.getElementById('presenceCount');
        const userListDiv = document.getElementById('presenceUserList');

        if (countSpan) countSpan.textContent = `${userCount} User${userCount > 1 ? 's' : ''} Online`;

        if (userListDiv) {
          userListDiv.innerHTML = '';
          Object.values(state).forEach(presences => {
            presences.forEach(p => {
              const name = p.name || 'Plant Operator';
              const role = p.role || 'Production';
              const row = document.createElement('div');
              row.className = 'presence-user-row';
              row.innerHTML = `<span>${name} (${role})</span><span class="badge badge-success">Online</span>`;
              userListDiv.appendChild(row);
            });
          });
        }
      })
      .subscribe(async (status) => {
        if (status === 'SUBSCRIBED') {
          await presenceChannel.track({
            name: currentUserProfile ? currentUserProfile.employee_name : 'Guest User',
            role: currentUserProfile ? currentUserProfile.role : 'Viewer',
            onlineAt: new Date().toISOString()
          });
        }
      });
  }

  let rawProdLoadedCount = 0;
  let rawQualLoadedCount = 0;
  let rawDownLoadedCount = 0;

  /* Paginated Table Fetcher - Fetches 100% of rows from Supabase (Bypassing PostgREST 1000-row limit) */
  async function fetchSupabaseTablePaginated(tableName) {
    if (!supabaseClient) return [];
    
    let allRecords = [];
    const pageSize = 1000;
    let page = 0;
    let hasMore = true;

    while (hasMore) {
      const from = page * pageSize;
      const to = from + pageSize - 1;

      const { data, error } = await supabaseClient
        .from(tableName)
        .select('*')
        .eq('is_deleted', false)
        .range(from, to);

      if (error) {
        console.error(`Error fetching page ${page} of ${tableName}:`, error);
        throw error;
      }

      if (data && data.length > 0) {
        allRecords = allRecords.concat(data);
        if (data.length < pageSize) {
          hasMore = false;
        } else {
          page++;
        }
      } else {
        hasMore = false;
      }
    }

    return allRecords;
  }

  /* Single Source of Truth Fetcher */
  async function fetchSupabaseShiftLogs() {
    if (!supabaseClient) return;

    const pill = document.getElementById('cloudDbStatusPill');

    try {
      const [prodList, qualList, downList] = await Promise.all([
        fetchSupabaseTablePaginated('production_data'),
        fetchSupabaseTablePaginated('quality_data'),
        fetchSupabaseTablePaginated('downtime_data')
      ]);

      const hasAugInSupabase = prodList.some(p => p.date && p.date.startsWith('2026-08'));
      let combinedProdList = prodList;
      if (!hasAugInSupabase && typeof MASTER_AUGUST_PRODUCTION_DATA !== 'undefined') {
        combinedProdList = [...MASTER_AUGUST_PRODUCTION_DATA, ...prodList];
      }
      
      const hasAugQualInSupabase = qualList.some(q => q.date && q.date.startsWith('2026-08'));
      if (!hasAugQualInSupabase && typeof MASTER_AUGUST_QUALITY_DATA !== 'undefined') {
        qualityLogs = [...MASTER_AUGUST_QUALITY_DATA, ...qualList];
      } else {
        qualityLogs = qualList;
      }

      rawProdLoadedCount = combinedProdList.length;
      rawQualLoadedCount = qualityLogs.length;
      rawDownLoadedCount = downList.length;

      const downMap = new Map();
      downList.forEach(d => {
        const key = `${d.date}_${d.shift}_${d.hammer}_${d.part_number}`;
        if (!downMap.has(key)) {
          downMap.set(key, { maintanceMins: 0, dieRelatedMins: 0, setupMins: 0, noManpowerMins: 0, heatingTimeMins: 0, minorStopMins: 0 });
        }
        const record = downMap.get(key);
        const mins = parseNum(d.downtime_minutes);
        switch (d.downtime_category) {
          case 'Maintenance': record.maintanceMins += mins; break;
          case 'Die Related': record.dieRelatedMins += mins; break;
          case 'Setup': record.setupMins += mins; break;
          case 'No Manpower': record.noManpowerMins += mins; break;
          case 'Heating Time': record.heatingTimeMins += mins; break;
          case 'Minor Stop': record.minorStopMins += mins; break;
        }
      });

      const qualMap = new Map();
      qualityLogs.forEach(q => {
        const key = `${q.date}_${q.shift}_${q.hammer}_${q.part_number}`;
        if (!qualMap.has(key)) qualMap.set(key, { rework_qty: 0, rejection_qty: 0 });
        const item = qualMap.get(key);
        item.rework_qty += parseNum(q.rework_qty);
        item.rejection_qty += parseNum(q.rejection_qty);
      });

      const fetchedLogs = combinedProdList.map(p => {
        const key = `${p.date}_${p.shift}_${p.hammer}_${p.part_number}`;
        const downInfo = downMap.get(key) || { maintanceMins: 0, dieRelatedMins: 0, setupMins: 0, noManpowerMins: 0, heatingTimeMins: 0, minorStopMins: 0 };
        const qualInfo = qualMap.get(key) || { rework_qty: 0, rejection_qty: 0 };

        return calculateOeeRecord({
          id: p.id || ('LOG-' + Math.random().toString(36).substr(2, 8).toUpperCase()),
          date: p.date,
          shift: p.shift,
          machine: p.hammer,
          partNumber: p.part_number,
          plannedTimeMins: parseNum(p.planned_time_mins, 660),
          maintanceMins: downInfo.maintanceMins,
          dieRelatedMins: downInfo.dieRelatedMins,
          setupMins: downInfo.setupMins,
          noManpowerMins: downInfo.noManpowerMins,
          heatingTimeMins: downInfo.heatingTimeMins,
          minorStopMins: downInfo.minorStopMins,
          totalParts: parseNum(p.production_qty),
          goodParts: parseNum(p.good_qty),
          rejects: parseNum(p.production_qty) - parseNum(p.good_qty),
          rework: qualInfo.rework_qty,
          idealCycleSec: parseNum(p.ideal_cycle_sec, 45)
        });
      });

      // Single source of truth: Assign fetched Supabase logs unconditionally
      shiftLogs = fetchedLogs;

      console.log('--- SUPABASE PAGINATED FETCH COMPLETED ---');
      console.log('production_data records loaded:', combinedProdList.length);
      console.log('quality_data records loaded:', qualityLogs.length);
      console.log('downtime_data records loaded:', downList.length);
      console.log('Total mapped shiftLogs:', shiftLogs.length);

      saveShiftLogs();
      renderAllViews();

      if (pill) {
        pill.innerHTML = `<span class="dot"></span> <i class="fa-solid fa-cloud"></i> 🟢 Live DB (${shiftLogs.length} logs)`;
        pill.style.borderColor = 'var(--primary)';
      }
    } catch (err) {
      console.error('Supabase Query Error:', err);
      showToast(`🔴 Supabase Query Failed: ${err.message || 'Database query error'}`, 'danger');
      if (pill) {
        pill.innerHTML = `<span class="dot"></span> <i class="fa-solid fa-triangle-exclamation text-danger"></i> 🔴 Supabase Error: ${err.message || 'Failed'}`;
        pill.style.borderColor = 'var(--danger)';
      }
    }
  }

  function setupCloudDbSync() {
    const syncHeaderBtn = document.getElementById('cloudDbSyncBtn');
    const pullBtn = document.getElementById('pullFromCloudDbBtn');

    if (syncHeaderBtn) syncHeaderBtn.addEventListener('click', () => fetchSupabaseShiftLogs());
    if (pullBtn) pullBtn.addEventListener('click', () => fetchSupabaseShiftLogs());
  }

  function saveShiftLogs() {
    localStorage.setItem('oee_shift_logs_v10', JSON.stringify(shiftLogs));
  }

  function updateMonthFilterOptions() {
    const select = document.getElementById('globalRangeFilter');
    if (!select) return;

    const currentSelected = select.value || '2026-08';

    const monthSet = new Set();
    shiftLogs.forEach(l => {
      const ym = getYearMonthString(l.date);
      if (ym) monthSet.add(ym);
    });
    qualityLogs.forEach(q => {
      const ym = getYearMonthString(q.date);
      if (ym) monthSet.add(ym);
    });

    const sortedMonths = Array.from(monthSet).sort().reverse();

    const formatMonthLabel = (mKey) => {
      const parts = mKey.split('-');
      if (parts.length !== 2) return mKey;
      const year = parts[0];
      const monthIdx = parseInt(parts[1], 10) - 1;
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${months[monthIdx] || parts[1]} ${year}`;
    };

    let html = `<option value="ALL">All Months (Combined)</option>`;
    sortedMonths.forEach(m => {
      html += `<option value="${m}">${formatMonthLabel(m)}</option>`;
    });

    const existingValues = Array.from(select.options).map(o => o.value).join(',');
    const newValues = ['ALL', ...sortedMonths].join(',');

    if (existingValues !== newValues) {
      select.innerHTML = html;
      if (Array.from(select.options).some(opt => opt.value === currentSelected)) {
        select.value = currentSelected;
      } else if (sortedMonths.includes('2026-08')) {
        select.value = '2026-08';
      } else {
        select.value = 'ALL';
      }
    }
  }

  function getFilteredLogs() {
    const hammerFilter = document.getElementById('globalHammerFilter').value;
    const shiftFilter = document.getElementById('globalShiftFilter').value;
    const monthFilter = document.getElementById('globalRangeFilter').value;

    return shiftLogs.filter(log => {
      // 1. Machine / Hammer filter
      if (hammerFilter !== 'ALL') {
        const normLogHammer = normalizeHammerName(log.machine);
        const normSelectedHammer = normalizeHammerName(hammerFilter);
        if (normLogHammer !== normSelectedHammer) return false;
      }

      // 2. Shift filter
      if (shiftFilter !== 'ALL') {
        if (!log.shift || log.shift.trim().toLowerCase() !== shiftFilter.trim().toLowerCase()) {
          return false;
        }
      }

      // 3. Month filter (YYYY-MM)
      if (monthFilter !== 'ALL') {
        const logYearMonth = getYearMonthString(log.date);
        if (logYearMonth !== monthFilter) return false;
      }

      return true;
    });
  }

  /* ==========================================================================
     SHIFT-LEVEL GROUPED KPI CALCULATION ENGINE (ELIMINATES PLANNED TIME DUPLICATION)
     ========================================================================== */
  function calculateFleetKpis(logs) {
    if (!logs || logs.length === 0) {
      return { totalShifts: 0, partEntries: 0, plannedMins: 0, operatingMins: 0, idealMins: 0, totalPcs: 0, goodPcs: 0, rejectPcs: 0, avail: 0, perf: 0, qual: 0, oee: 0 };
    }

    const shiftMap = new Map();
    logs.forEach(l => {
      const key = `${l.date}_${l.shift}_${l.machine}`;
      if (!shiftMap.has(key)) {
        shiftMap.set(key, {
          plannedMins: parseNum(l.plannedTimeMins, 660),
          maintanceMins: parseNum(l.maintanceMins),
          dieRelatedMins: parseNum(l.dieRelatedMins),
          setupMins: parseNum(l.setupMins),
          noManpowerMins: parseNum(l.noManpowerMins),
          heatingTimeMins: parseNum(l.heatingTimeMins),
          minorStopMins: parseNum(l.minorStopMins),
          parts: []
        });
      }
      const sObj = shiftMap.get(key);
      sObj.plannedMins = Math.max(sObj.plannedMins, parseNum(l.plannedTimeMins, 660));
      sObj.parts.push(l);
    });

    let plannedMins = 0, operatingMins = 0, idealMins = 0, totalPcs = 0, goodPcs = 0;

    shiftMap.forEach(s => {
      const totalDowntime = s.maintanceMins + s.dieRelatedMins + s.setupMins + s.noManpowerMins + s.heatingTimeMins + s.minorStopMins;
      const opMins = Math.max(0, s.plannedMins - totalDowntime);

      plannedMins += s.plannedMins;
      operatingMins += opMins;

      s.parts.forEach(p => {
        totalPcs += parseNum(p.totalParts);
        goodPcs += parseNum(p.goodParts);
        idealMins += (parseNum(p.totalParts) * parseNum(p.idealCycleSec, 45)) / 60;
      });
    });

    const rejectPcs = totalPcs - goodPcs;
    const avail = plannedMins > 0 ? (operatingMins / plannedMins) * 100 : 0;
    const perf = operatingMins > 0 ? Math.min(100, (idealMins / operatingMins) * 100) : 0;
    const qual = totalPcs > 0 ? (goodPcs / totalPcs) * 100 : 100;
    const oee = (avail / 100) * (perf / 100) * (qual / 100) * 100;

    return {
      totalShifts: shiftMap.size,
      partEntries: logs.length,
      plannedMins,
      operatingMins,
      idealMins,
      totalPcs,
      goodPcs,
      rejectPcs,
      avail: parseFloat(avail.toFixed(1)),
      perf: parseFloat(perf.toFixed(1)),
      qual: parseFloat(qual.toFixed(1)),
      oee: parseFloat(oee.toFixed(1))
    };
  }

  function updateDebugDiagnosticsPanel(filteredLogs) {
    const prodLoadEl = document.getElementById('dbgProdLoaded');
    const qualLoadEl = document.getElementById('dbgQualLoaded');
    const downLoadEl = document.getElementById('dbgDownLoaded');
    const rawEl = document.getElementById('dbgRawRecords');

    const latestDateEl = document.getElementById('dbgLatestDate');
    const earliestDateEl = document.getElementById('dbgEarliestDate');

    const selMachEl = document.getElementById('dbgSelMachine');
    const selMonthEl = document.getElementById('dbgSelMonth');
    const selShiftEl = document.getElementById('dbgSelShift');
    const filtEl = document.getElementById('dbgFilteredRecords');

    const dbMonthsEl = document.getElementById('dbgDbMonths');
    const dbMachinesEl = document.getElementById('dbgDbMachines');

    const oeeCnt = document.getElementById('dbgOeeCnt');
    const oeeVal = document.getElementById('dbgOeeVal');
    const availCnt = document.getElementById('dbgAvailCnt');
    const availVal = document.getElementById('dbgAvailVal');
    const perfCnt = document.getElementById('dbgPerfCnt');
    const perfVal = document.getElementById('dbgPerfVal');
    const qualCnt = document.getElementById('dbgQualCnt');
    const qualVal = document.getElementById('dbgQualVal');

    const firstRec = document.getElementById('dbgFirstRecord');

    if (prodLoadEl) prodLoadEl.textContent = rawProdLoadedCount;
    if (qualLoadEl) qualLoadEl.textContent = rawQualLoadedCount;
    if (downLoadEl) downLoadEl.textContent = rawDownLoadedCount;
    if (rawEl) rawEl.textContent = shiftLogs.length;

    const allDates = shiftLogs.map(l => getStandardDateString(l.date)).filter(Boolean).sort();
    if (earliestDateEl) earliestDateEl.textContent = allDates.length > 0 ? allDates[0] : '--';
    if (latestDateEl) latestDateEl.textContent = allDates.length > 0 ? allDates[allDates.length - 1] : '--';

    const hVal = document.getElementById('globalHammerFilter')?.value || 'ALL';
    const sVal = document.getElementById('globalShiftFilter')?.value || 'ALL';
    const mVal = document.getElementById('globalRangeFilter')?.value || 'ALL';

    if (selMachEl) selMachEl.textContent = hVal;
    if (selMonthEl) selMonthEl.textContent = mVal;
    if (selShiftEl) selShiftEl.textContent = sVal;
    if (filtEl) filtEl.textContent = filteredLogs.length;

    const uniqueMonths = Array.from(new Set(shiftLogs.map(l => getYearMonthString(l.date)).filter(Boolean))).sort();
    const uniqueMachines = Array.from(new Set(shiftLogs.map(l => l.machine).filter(Boolean))).sort();

    if (dbMonthsEl) dbMonthsEl.textContent = uniqueMonths.join(', ') || 'None';
    if (dbMachinesEl) dbMachinesEl.textContent = uniqueMachines.join(', ') || 'None';

    let totalPlannedMins = 0, totalOperatingMins = 0, totalIdealMins = 0, totalProduced = 0, totalGood = 0;

    filteredLogs.forEach(l => {
      totalPlannedMins += parseNum(l.plannedTimeMins, 660);
      totalOperatingMins += parseNum(l.operatingTimeMins);
      totalIdealMins += Math.min(parseNum(l.operatingTimeMins), (parseNum(l.totalParts) * parseNum(l.idealCycleSec, 45)) / 60);
      totalProduced += parseNum(l.totalParts);
      totalGood += parseNum(l.goodParts);
    });

    const avgAvail = totalPlannedMins > 0 ? (totalOperatingMins / totalPlannedMins) * 100 : 0;
    const avgPerf = totalOperatingMins > 0 ? Math.min(100, (totalIdealMins / totalOperatingMins) * 100) : 0;
    const avgQual = totalProduced > 0 ? (totalGood / totalProduced) * 100 : 100;
    const overallOee = (avgAvail / 100) * (avgPerf / 100) * (avgQual / 100) * 100;

    if (oeeCnt) oeeCnt.textContent = filteredLogs.length;
    if (oeeVal) oeeVal.textContent = overallOee.toFixed(1) + '%';
    if (availCnt) availCnt.textContent = filteredLogs.length;
    if (availVal) availVal.textContent = avgAvail.toFixed(1) + '%';
    if (perfCnt) perfCnt.textContent = filteredLogs.length;
    if (perfVal) perfVal.textContent = avgPerf.toFixed(1) + '%';
    if (qualCnt) qualCnt.textContent = filteredLogs.length;
    if (qualVal) qualVal.textContent = avgQual.toFixed(1) + '%';

    if (firstRec) {
      if (filteredLogs.length > 0) {
        const sample = {
          date: filteredLogs[0].date,
          shift: filteredLogs[0].shift,
          machine: filteredLogs[0].machine,
          partNumber: filteredLogs[0].partNumber,
          plannedTimeMins: filteredLogs[0].plannedTimeMins,
          totalParts: filteredLogs[0].totalParts,
          goodParts: filteredLogs[0].goodParts,
          oee: filteredLogs[0].oee + '%'
        };
        firstRec.textContent = JSON.stringify(sample, null, 2);
      } else {
        firstRec.textContent = 'None (0 records match selected filters)';
      }
    }
  }

  function updateDataReconciliationPanel(filteredLogs) {
    const prodCntEl = document.getElementById('reconProdCnt');
    const qualCntEl = document.getElementById('reconQualCnt');
    const downCntEl = document.getElementById('reconDownCnt');
    const matchedCntEl = document.getElementById('reconMatchedCnt');

    const unProdEl = document.getElementById('reconUnmatchedProd');
    const unQualEl = document.getElementById('reconUnmatchedQual');
    const unDownEl = document.getElementById('reconUnmatchedDown');

    const totPartsEl = document.getElementById('reconTotParts');
    const goodPartsEl = document.getElementById('reconGoodParts');
    const rejectPartsEl = document.getElementById('reconRejectParts');

    const calcAvailEl = document.getElementById('reconCalcAvail');
    const calcPerfEl = document.getElementById('reconCalcPerf');
    const calcQualEl = document.getElementById('reconCalcQual');
    const calcOeeEl = document.getElementById('reconCalcOee');
    const dashOeeEl = document.getElementById('reconDashOee');

    if (prodCntEl) prodCntEl.textContent = rawProdLoadedCount;
    if (qualCntEl) qualCntEl.textContent = rawQualLoadedCount;
    if (downCntEl) downCntEl.textContent = rawDownLoadedCount;

    const stats = calculateFleetKpis(filteredLogs);
    if (matchedCntEl) matchedCntEl.textContent = `${filteredLogs.length} part entries (${stats.totalShifts} unique shifts)`;

    if (unProdEl) unProdEl.textContent = '0';
    if (unQualEl) unQualEl.textContent = '0';
    if (unDownEl) unDownEl.textContent = '0';

    if (totPartsEl) totPartsEl.textContent = stats.totalPcs.toLocaleString() + ' pcs';
    if (goodPartsEl) goodPartsEl.textContent = stats.goodPcs.toLocaleString() + ' pcs';
    if (rejectPartsEl) rejectPartsEl.textContent = stats.rejectPcs.toLocaleString() + ' pcs';

    if (calcAvailEl) calcAvailEl.textContent = stats.avail.toFixed(1) + '%';
    if (calcPerfEl) calcPerfEl.textContent = stats.perf.toFixed(1) + '%';
    if (calcQualEl) calcQualEl.textContent = stats.qual.toFixed(1) + '%';
    if (calcOeeEl) calcOeeEl.textContent = stats.oee.toFixed(1) + '%';
    if (dashOeeEl) dashOeeEl.textContent = document.getElementById('kpiOverallOee')?.textContent || stats.oee.toFixed(1) + '%';
  }

  /* ==========================================================================
     UI RENDERERS & KPI CALCULATIONS
     ========================================================================== */
  function renderAllViews() {
    shiftLogs = shiftLogs.map(l => calculateOeeRecord(l));
    updateMonthFilterOptions();
    const logs = getFilteredLogs();

    const stats = calculateFleetKpis(logs);

    console.log('--- DASHBOARD RECALCULATION DEBUG ---');
    console.log('Total Supabase records loaded:', shiftLogs.length);
    console.log('Filtered part entries:', logs.length);
    console.log('Filtered unique hammer-shifts:', stats.totalShifts);
    console.log('Calculated Availability:', stats.avail.toFixed(1) + '%');
    console.log('Calculated Performance:', stats.perf.toFixed(1) + '%');
    console.log('Calculated Quality:', stats.qual.toFixed(1) + '%');
    console.log('Calculated OEE:', stats.oee.toFixed(1) + '%');

    updateDebugDiagnosticsPanel(logs);
    updateDataReconciliationPanel(logs);
    updateHammerLogCountBadges();
    renderOverviewKpis(logs);
    renderHammerGauges(logs);
    renderMonthlyTrendView(logs);
    renderInsightsView(logs);
    renderComparisonView(logs);
    renderDowntimeView(logs);
    renderLogsTable(logs);
    renderQualityActivityMonitor();
  }

  /* ==========================================================================
     QUALITY ACTIVITY MONITOR & DEFECT ANALYTICS ENGINE
     ========================================================================== */
  function renderQualityActivityMonitor() {
    const monthFilter = document.getElementById('qmMonthFilter');
    const stageFilter = document.getElementById('qmStageFilter');
    const partFilter = document.getElementById('qmPartFilter');
    const reasonFilter = document.getElementById('qmReasonFilter');

    if (!monthFilter || !stageFilter) return;

    populateQualityFilterOptions();

    const selectedMonth = monthFilter.value;
    const selectedStage = stageFilter.value;
    const selectedPart = partFilter.value;
    const selectedReason = reasonFilter.value;

    const filteredQuality = qualityLogs.filter(q => {
      if (selectedMonth !== 'ALL' && q.date && !q.date.startsWith(selectedMonth)) return false;
      if (selectedStage !== 'ALL' && q.inspection_stage !== selectedStage) return false;
      if (selectedPart !== 'ALL' && q.part_number !== selectedPart) return false;
      if (selectedReason !== 'ALL' && (q.reason !== selectedReason && q.rework_reason !== selectedReason && q.rejection_reason !== selectedReason)) return false;
      return true;
    });

    let totalInspected = 0;
    let totalRework = 0;
    let totalRejection = 0;

    filteredQuality.forEach(q => {
      totalInspected += parseNum(q.inspection_qty);
      totalRework += parseNum(q.rework_qty);
      totalRejection += parseNum(q.rejection_qty);
    });

    const reworkPct = totalInspected > 0 ? ((totalRework / totalInspected) * 100).toFixed(1) : '0.0';
    const rejectionPct = totalInspected > 0 ? ((totalRejection / totalInspected) * 100).toFixed(1) : '0.0';

    document.getElementById('qmTotalInspected').textContent = `${totalInspected.toLocaleString()} pcs`;
    document.getElementById('qmTotalRework').textContent = `${totalRework.toLocaleString()} pcs`;
    document.getElementById('qmTotalRejection').textContent = `${totalRejection.toLocaleString()} pcs`;
    document.getElementById('qmReworkPct').textContent = `Rework: ${reworkPct}%`;
    document.getElementById('qmRejectionPct').textContent = `Rejection: ${rejectionPct}%`;
    document.getElementById('qmRecordCountBadge').textContent = `${filteredQuality.length} Records`;

    renderQualityRecordsTable(filteredQuality);
    renderPartWiseQualityTables(filteredQuality);
    renderQualityCharts(filteredQuality);
    renderTopQualityReasons(filteredQuality);
  }

  function populateQualityFilterOptions() {
    const monthSel = document.getElementById('qmMonthFilter');
    const partSel = document.getElementById('qmPartFilter');
    const reasonSel = document.getElementById('qmReasonFilter');

    if (!monthSel || !partSel || !reasonSel) return;

    const curMonth = monthSel.value;
    const curPart = partSel.value;
    const curReason = reasonSel.value;

    const monthSet = new Set();
    const partSet = new Set();
    const reasonSet = new Set();

    qualityLogs.forEach(q => {
      if (q.date && q.date.length >= 7) monthSet.add(q.date.substring(0, 7));
      if (q.part_number) partSet.add(q.part_number);
      if (q.reason) reasonSet.add(q.reason);
      if (q.rework_reason) reasonSet.add(q.rework_reason);
      if (q.rejection_reason) reasonSet.add(q.rejection_reason);
    });

    if (monthSel.options.length <= 1) {
      monthSel.innerHTML = '<option value="ALL">All Months</option>';
      Array.from(monthSet).sort().reverse().forEach(m => {
        monthSel.innerHTML += `<option value="${m}">${m}</option>`;
      });
    }

    if (partSel.options.length <= 1) {
      partSel.innerHTML = '<option value="ALL">All Part Numbers</option>';
      Array.from(partSet).sort().forEach(p => {
        partSel.innerHTML += `<option value="${p}">${p}</option>`;
      });
    }

    if (reasonSel.options.length <= 1) {
      reasonSel.innerHTML = '<option value="ALL">All Defect Reasons</option>';
      Array.from(reasonSet).sort().forEach(r => {
        reasonSel.innerHTML += `<option value="${r}">${r}</option>`;
      });
    }

    monthSel.value = curMonth || 'ALL';
    partSel.value = curPart || 'ALL';
    reasonSel.value = curReason || 'ALL';
  }

  function renderQualityRecordsTable(logs) {
    const tbody = document.getElementById('qmTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (logs.length === 0) {
      tbody.innerHTML = `<tr><td colspan="13" style="text-align: center; color: var(--text-muted); padding: 20px;">No quality records found matching selected filters.</td></tr>`;
      return;
    }

    logs.forEach(q => {
      const insPcs = parseNum(q.inspection_qty);
      const rewPcs = parseNum(q.rework_qty);
      const rejPcs = parseNum(q.rejection_qty);
      const rewPct = insPcs > 0 ? ((rewPcs / insPcs) * 100).toFixed(1) : '0.0';
      const rejPct = insPcs > 0 ? ((rejPcs / insPcs) * 100).toFixed(1) : '0.0';

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${q.date}</td>
        <td><span class="badge badge-info">${q.shift || 'Shift A'}</span></td>
        <td><strong>${q.hammer || 'Fleet'}</strong></td>
        <td><strong style="color: var(--primary);">${q.part_number}</strong></td>
        <td><span class="badge badge-primary">${q.inspection_stage}</span></td>
        <td>${insPcs.toLocaleString()}</td>
        <td><span class="text-warning" style="font-weight: 700;">${rewPcs}</span></td>
        <td>${rewPct}%</td>
        <td><span class="text-danger" style="font-weight: 700;">${rejPcs}</span></td>
        <td>${rejPct}%</td>
        <td>${q.rework_reason || q.reason || '-'}</td>
        <td>${q.rejection_reason || q.reason || '-'}</td>
        <td><small>${q.created_by ? 'User #' + String(q.created_by).substring(0, 6) : 'Operator'}</small></td>
      `;
      tbody.appendChild(tr);
    });
  }

  function renderPartWiseQualityTables(logs) {
    const rewBody = document.getElementById('qmPartReworkTableBody');
    const rejBody = document.getElementById('qmPartRejectionTableBody');

    if (!rewBody || !rejBody) return;
    rewBody.innerHTML = '';
    rejBody.innerHTML = '';

    const partMap = new Map();
    logs.forEach(q => {
      if (!partMap.has(q.part_number)) {
        partMap.set(q.part_number, { inspected: 0, rework: 0, rejection: 0 });
      }
      const item = partMap.get(q.part_number);
      item.inspected += parseNum(q.inspection_qty);
      item.rework += parseNum(q.rework_qty);
      item.rejection += parseNum(q.rejection_qty);
    });

    const parts = Array.from(partMap.entries());

    const reworkParts = [...parts].sort((a, b) => b[1].rework - a[1].rework).filter(p => p[1].rework > 0);
    if (reworkParts.length === 0) {
      rewBody.innerHTML = `<tr><td colspan="5" style="text-align: center; padding: 14px;">No rework recorded.</td></tr>`;
    } else {
      reworkParts.forEach(([part, data]) => {
        const pct = data.inspected > 0 ? ((data.rework / data.inspected) * 100).toFixed(1) : '0.0';
        rewBody.innerHTML += `
          <tr>
            <td><strong>${part}</strong></td>
            <td>${data.inspected.toLocaleString()}</td>
            <td><strong class="text-warning">${data.rework}</strong></td>
            <td>${pct}%</td>
            <td><span class="badge ${pct > 3 ? 'badge-danger' : 'badge-warning'}">${pct > 3 ? 'High Rework' : 'Moderate'}</span></td>
          </tr>
        `;
      });
    }

    const rejectionParts = [...parts].sort((a, b) => b[1].rejection - a[1].rejection).filter(p => p[1].rejection > 0);
    if (rejectionParts.length === 0) {
      rejBody.innerHTML = `<tr><td colspan="5" style="text-align: center; padding: 14px;">No rejections recorded.</td></tr>`;
    } else {
      rejectionParts.forEach(([part, data]) => {
        const pct = data.inspected > 0 ? ((data.rejection / data.inspected) * 100).toFixed(1) : '0.0';
        rejBody.innerHTML += `
          <tr>
            <td><strong>${part}</strong></td>
            <td>${data.inspected.toLocaleString()}</td>
            <td><strong class="text-danger">${data.rejection}</strong></td>
            <td>${pct}%</td>
            <td><span class="badge ${pct > 2 ? 'badge-danger' : 'badge-warning'}">${pct > 2 ? 'Critical Scrap' : 'Monitor'}</span></td>
          </tr>
        `;
      });
    }
  }

  function renderQualityCharts(logs) {
    const paretoCanvas = document.getElementById('qualityParetoChart');
    if (paretoCanvas) {
      const reasonMap = new Map();
      logs.forEach(q => {
        const rReason = q.rework_reason || q.reason;
        const jReason = q.rejection_reason || q.reason;
        if (rReason && q.rework_qty > 0) reasonMap.set(rReason, (reasonMap.get(rReason) || 0) + parseNum(q.rework_qty));
        if (jReason && q.rejection_qty > 0) reasonMap.set(jReason, (reasonMap.get(jReason) || 0) + parseNum(q.rejection_qty));
      });

      const sortedReasons = Array.from(reasonMap.entries()).sort((a, b) => b[1] - a[1]).slice(0, 10);
      const labels = sortedReasons.map(r => r[0]);
      const data = sortedReasons.map(r => r[1]);

      if (charts.qualityPareto) charts.qualityPareto.destroy();

      charts.qualityPareto = new Chart(paretoCanvas, {
        type: 'bar',
        data: {
          labels: labels.length ? labels : ['No Defect Data'],
          datasets: [{
            label: 'Defect / Defective Pcs',
            data: data.length ? data : [0],
            backgroundColor: 'rgba(217, 119, 6, 0.75)',
            borderColor: '#d97706',
            borderWidth: 1
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } }
        }
      });
    }

    const trendCanvas = document.getElementById('monthlyQualityTrendChart');
    if (trendCanvas) {
      const monthMap = new Map();
      logs.forEach(q => {
        if (!q.date || q.date.length < 7) return;
        const m = q.date.substring(0, 7);
        if (!monthMap.has(m)) monthMap.set(m, { rework: 0, rejection: 0 });
        const item = monthMap.get(m);
        item.rework += parseNum(q.rework_qty);
        item.rejection += parseNum(q.rejection_qty);
      });

      const sortedMonths = Array.from(monthMap.keys()).sort();
      const reworkData = sortedMonths.map(m => monthMap.get(m).rework);
      const rejectionData = sortedMonths.map(m => monthMap.get(m).rejection);

      if (charts.monthlyQuality) charts.monthlyQuality.destroy();

      charts.monthlyQuality = new Chart(trendCanvas, {
        type: 'line',
        data: {
          labels: sortedMonths.length ? sortedMonths : ['No Trend Data'],
          datasets: [
            {
              label: 'Rework Pcs',
              data: reworkData.length ? reworkData : [0],
              borderColor: '#d97706',
              backgroundColor: 'rgba(217, 119, 6, 0.1)',
              tension: 0.3
            },
            {
              label: 'Rejection Pcs',
              data: rejectionData.length ? rejectionData : [0],
              borderColor: '#dc2626',
              backgroundColor: 'rgba(220, 38, 38, 0.1)',
              tension: 0.3
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false
        }
      });
    }
  }

  function renderTopQualityReasons(logs) {
    const rewBody = document.getElementById('qmTopReworkReasonsBody');
    const rejBody = document.getElementById('qmTopRejectionReasonsBody');

    if (!rewBody || !rejBody) return;
    rewBody.innerHTML = '';
    rejBody.innerHTML = '';

    const rewMap = new Map();
    const rejMap = new Map();
    let totalRew = 0;
    let totalRej = 0;

    logs.forEach(q => {
      const rReason = q.rework_reason || q.reason;
      const jReason = q.rejection_reason || q.reason;
      const rewQty = parseNum(q.rework_qty);
      const rejQty = parseNum(q.rejection_qty);

      if (rReason && rewQty > 0) {
        rewMap.set(rReason, (rewMap.get(rReason) || 0) + rewQty);
        totalRew += rewQty;
      }
      if (jReason && rejQty > 0) {
        rejMap.set(jReason, (rejMap.get(jReason) || 0) + rejQty);
        totalRej += rejQty;
      }
    });

    const topRew = Array.from(rewMap.entries()).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const topRej = Array.from(rejMap.entries()).sort((a, b) => b[1] - a[1]).slice(0, 10);

    if (topRew.length === 0) {
      rewBody.innerHTML = `<tr><td colspan="4" style="text-align: center; padding: 14px;">No rework defect data available.</td></tr>`;
    } else {
      topRew.forEach(([reason, qty], idx) => {
        const share = totalRew > 0 ? ((qty / totalRew) * 100).toFixed(1) : '0.0';
        rewBody.innerHTML += `
          <tr>
            <td><strong>#${idx + 1}</strong></td>
            <td>${reason}</td>
            <td><strong class="text-warning">${qty}</strong></td>
            <td>${share}%</td>
          </tr>
        `;
      });
    }

    if (topRej.length === 0) {
      rejBody.innerHTML = `<tr><td colspan="4" style="text-align: center; padding: 14px;">No rejection defect data available.</td></tr>`;
    } else {
      topRej.forEach(([reason, qty], idx) => {
        const share = totalRej > 0 ? ((qty / totalRej) * 100).toFixed(1) : '0.0';
        rejBody.innerHTML += `
          <tr>
            <td><strong>#${idx + 1}</strong></td>
            <td>${reason}</td>
            <td><strong class="text-danger">${qty}</strong></td>
            <td>${share}%</td>
          </tr>
        `;
      });
    }
  }

  function updateHammerLogCountBadges() {
    HAMMERS.forEach(h => {
      const count = shiftLogs.filter(l => l.machine === h.name).length;
      const el = document.getElementById(h.badgeId);
      if (el) {
        el.textContent = `${count} logs`;
      }
    });
  }

  function renderOverviewKpis(logs) {
    if (logs.length === 0) {
      document.getElementById('kpiOverallOee').textContent = '0.0%';
      document.getElementById('kpiAvailability').textContent = '0.0%';
      document.getElementById('kpiPerformance').textContent = '0.0%';
      document.getElementById('kpiQuality').textContent = '0.0%';
      document.getElementById('kpiPlannedHours').textContent = 'Net Planned: 0 hrs';
      document.getElementById('kpiTotalPieces').textContent = 'Good Parts: 0 pcs';
      document.getElementById('kpiScrapRate').textContent = 'Rejects: 0 pcs (0.0%)';
      document.getElementById('kpiOeeStatus').innerHTML = '<i class="fa-solid fa-triangle-exclamation text-warning"></i> No data available for the selected Machine and Month';
      return;
    }

    const stats = calculateFleetKpis(logs);

    const scrapPct = stats.totalPcs > 0 ? ((stats.rejectPcs / stats.totalPcs) * 100).toFixed(1) : '0.0';

    document.getElementById('kpiOverallOee').textContent = stats.oee.toFixed(1) + '%';
    document.getElementById('kpiAvailability').textContent = stats.avail.toFixed(1) + '%';
    document.getElementById('kpiPerformance').textContent = stats.perf.toFixed(1) + '%';
    document.getElementById('kpiQuality').textContent = stats.qual.toFixed(1) + '%';

    document.getElementById('kpiPlannedHours').textContent = `Net Planned: ${(stats.plannedMins / 60).toFixed(1)} hrs (${stats.totalShifts} shifts)`;
    document.getElementById('kpiTotalPieces').textContent = `Good Parts: ${stats.goodPcs.toLocaleString()} pcs`;
    document.getElementById('kpiScrapRate').textContent = `Rejects: ${stats.rejectPcs.toLocaleString()} pcs (${scrapPct}%)`;

    const oeeStatusEl = document.getElementById('kpiOeeStatus');
    if (stats.oee >= 85) {
      oeeStatusEl.innerHTML = '<i class="fa-solid fa-circle-check text-success"></i> World-Class (≥85%)';
    } else if (stats.oee >= 75) {
      oeeStatusEl.innerHTML = '<i class="fa-solid fa-circle-check text-success"></i> Meets Target (≥75%)';
    } else {
      oeeStatusEl.innerHTML = '<i class="fa-solid fa-triangle-exclamation text-warning"></i> Below Target (<75%)';
    }
  }

  /* ==========================================================================
     PART NUMBER & MACHINE INSIGHTS VIEW (HAMMER-WISE BREAKDOWN)
     ========================================================================== */
  function renderInsightsView(allLogs) {
    if (allLogs.length === 0) {
      document.getElementById('summaryHighestMaintMachine').textContent = 'No logs';
      document.getElementById('summaryHighestSetupPart').textContent = 'No logs';
      document.getElementById('summaryHighestDiePart').textContent = 'No logs';
      document.getElementById('summaryHighestHeatingPart').textContent = 'No logs';
      return;
    }

    // Filter by subnav hammer selector
    const targetLogs = selectedSubnavHammer === 'ALL' ? allLogs : allLogs.filter(l => l.machine === selectedSubnavHammer);

    // Update section badges
    const badgeLabelText = selectedSubnavHammer === 'ALL' ? 'All 5 Hammers Combined' : selectedSubnavHammer;
    document.getElementById('setupBadgeLabel').textContent = badgeLabelText;
    document.getElementById('dieBadgeLabel').textContent = badgeLabelText;
    document.getElementById('heatingBadgeLabel').textContent = badgeLabelText;

    // Group strictly by exact Part Number String from logs
    const partStats = {};
    targetLogs.forEach(l => {
      const part = l.partNumber || 'Unspecified Part';
      if (!partStats[part]) {
        partStats[part] = {
          part: part,
          shifts: 0,
          setupMins: 0,
          dieMins: 0,
          heatingMins: 0
        };
      }
      partStats[part].shifts += 1;
      partStats[part].setupMins += l.setupMins;
      partStats[part].dieMins += l.dieRelatedMins;
      partStats[part].heatingMins += l.heatingTimeMins;
    });

    const partsList = Object.values(partStats);

    // Group by Machine for Maintenance
    const machineMaintStats = HAMMERS.map(h => {
      const hLogs = allLogs.filter(l => l.machine === h.name);
      const totalMaintMins = hLogs.reduce((acc, curr) => acc + curr.maintanceMins, 0);
      return {
        machine: h.name,
        color: h.color,
        shifts: hLogs.length,
        maintMins: totalMaintMins,
        avgMins: hLogs.length > 0 ? parseFloat((totalMaintMins / hLogs.length).toFixed(1)) : 0
      };
    });

    // Sort Lists
    const sortedBySetup = [...partsList].sort((a, b) => b.setupMins - a.setupMins);
    const sortedByDie = [...partsList].sort((a, b) => b.dieMins - a.dieMins);
    const sortedByHeating = [...partsList].sort((a, b) => b.heatingMins - a.heatingMins);
    const sortedByMaintMachine = [...machineMaintStats].sort((a, b) => b.maintMins - a.maintMins);

    // Populate Top Executive Summary Banner
    if (sortedByMaintMachine.length > 0 && sortedByMaintMachine[0].shifts > 0) {
      const topMaint = sortedByMaintMachine[0];
      document.getElementById('summaryHighestMaintMachine').textContent = topMaint.machine;
      document.getElementById('summaryHighestMaintVal').textContent = `${topMaint.maintMins} mins (${(topMaint.maintMins/60).toFixed(1)} hrs)`;
    } else {
      document.getElementById('summaryHighestMaintMachine').textContent = 'None';
      document.getElementById('summaryHighestMaintVal').textContent = '0 mins downtime';
    }

    if (sortedBySetup.length > 0) {
      const topSetup = sortedBySetup[0];
      document.getElementById('summaryHighestSetupPart').textContent = `${topSetup.part}`;
      document.getElementById('summaryHighestSetupVal').textContent = `${topSetup.setupMins} mins setup loss`;
    } else {
      document.getElementById('summaryHighestSetupPart').textContent = 'None';
      document.getElementById('summaryHighestSetupVal').textContent = '0 mins setup loss';
    }

    if (sortedByDie.length > 0) {
      const topDie = sortedByDie[0];
      document.getElementById('summaryHighestDiePart').textContent = `${topDie.part}`;
      document.getElementById('summaryHighestDieVal').textContent = `${topDie.dieMins} mins die downtime`;
    } else {
      document.getElementById('summaryHighestDiePart').textContent = 'None';
      document.getElementById('summaryHighestDieVal').textContent = '0 mins die downtime';
    }

    if (sortedByHeating.length > 0) {
      const topHeating = sortedByHeating[0];
      document.getElementById('summaryHighestHeatingPart').textContent = `${topHeating.part}`;
      document.getElementById('summaryHighestHeatingVal').textContent = `${topHeating.heatingMins} mins furnace wait`;
    } else {
      document.getElementById('summaryHighestHeatingPart').textContent = 'None';
      document.getElementById('summaryHighestHeatingVal').textContent = '0 mins heating wait';
    }

    // Render Main Charts & Tables
    renderPartBarChart('partSetupChart', sortedBySetup.slice(0, 6), 'setupMins', 'Setup Mins', '#16a34a');
    renderPartAnalyticsTable('partSetupTableBody', sortedBySetup, 'setupMins');

    renderPartBarChart('partDieChart', sortedByDie.slice(0, 6), 'dieMins', 'Die Downtime Mins', '#dc2626');
    renderPartAnalyticsTable('partDieTableBody', sortedByDie, 'dieMins');

    renderPartBarChart('partHeatingChart', sortedByHeating.slice(0, 6), 'heatingMins', 'Heating Mins', '#0284c7');
    renderPartAnalyticsTable('partHeatingTableBody', sortedByHeating, 'heatingMins');

    renderMachineMaintChart('machineMaintenanceChart', sortedByMaintMachine);
    renderMachineMaintTable('machineMaintenanceTableBody', sortedByMaintMachine);

    // Render Dedicated 5-Hammer Specific Cards Grid
    renderHammerWiseCardsGrid(allLogs);
  }

  function renderHammerWiseCardsGrid(allLogs) {
    const container = document.getElementById('hammerWiseCardsContainer');
    if (!container) return;
    container.innerHTML = '';

    HAMMERS.forEach(h => {
      const hLogs = allLogs.filter(l => l.machine === h.name);
      
      const hammerPartStats = {};
      hLogs.forEach(l => {
        const p = l.partNumber;
        if (!hammerPartStats[p]) {
          hammerPartStats[p] = { part: p, setup: 0, die: 0, heating: 0, shifts: 0 };
        }
        hammerPartStats[p].setup += l.setupMins;
        hammerPartStats[p].die += l.dieRelatedMins;
        hammerPartStats[p].heating += l.heatingTimeMins;
        hammerPartStats[p].shifts += 1;
      });

      const pList = Object.values(hammerPartStats).sort((a, b) => (b.setup + b.die + b.heating) - (a.setup + a.die + a.heating));

      const card = document.createElement('div');
      card.className = 'hammer-hw-card';
      
      let rowsHtml = '';
      if (pList.length === 0) {
        rowsHtml = '<p style="color: var(--text-muted); font-size: 12px; margin-top: 8px;">Awaiting Excel upload for this hammer.</p>';
      } else {
        pList.slice(0, 4).forEach(pItem => {
          rowsHtml += `
            <div class="hw-part-row">
              <div>
                <strong>${pItem.part}</strong> <span style="font-size: 11px; color: var(--text-muted);">(${pItem.shifts} shifts)</span>
              </div>
              <div style="font-size: 11px; text-align: right;">
                <span class="text-primary" title="Setup Mins">S: ${pItem.setup}m</span> | 
                <span class="text-danger" title="Die Mins">D: ${pItem.die}m</span> | 
                <span class="text-info" title="Heating Mins">H: ${pItem.heating}m</span>
              </div>
            </div>
          `;
        });
      }

      card.innerHTML = `
        <div class="hw-card-header">
          <h4><i class="fa-solid fa-hammer" style="color: ${h.color};"></i> ${h.name}</h4>
          <span class="badge badge-info">${hLogs.length} shifts</span>
        </div>
        <div style="font-size: 11px; color: var(--text-muted); margin-bottom: 10px;">
          Capacity: ${h.capacity} | Cycle: ${h.defaultCycle}s
        </div>
        <div class="hw-parts-list">
          ${rowsHtml}
        </div>
      `;

      container.appendChild(card);
    });
  }

  function renderPartBarChart(canvasId, items, key, labelStr, colorHex) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    charts[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: items.map(i => i.part),
        datasets: [{
          label: labelStr,
          data: items.map(i => i[key]),
          backgroundColor: colorHex,
          borderRadius: 6
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { ticks: { color: currentTheme === 'dark' ? '#94a3b8' : '#047857' } },
          y: { ticks: { color: currentTheme === 'dark' ? '#ecfdf5' : '#064e3b' } }
        }
      }
    });
  }

  function renderPartAnalyticsTable(tableBodyId, items, key) {
    const tbody = document.getElementById(tableBodyId);
    if (!tbody) return;
    tbody.innerHTML = '';

    if (items.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-muted); padding: 12px;">Awaiting Excel data.</td></tr>`;
      return;
    }

    items.slice(0, 5).forEach((item, idx) => {
      const tr = document.createElement('tr');
      const val = item[key];
      const avg = item.shifts > 0 ? (val / item.shifts).toFixed(1) : 0;
      const level = idx === 0 ? 'CRITICAL' : (idx < 3 ? 'HIGH' : 'MODERATE');
      const badgeClass = idx === 0 ? 'danger' : (idx < 3 ? 'warning' : 'info');

      tr.innerHTML = `
        <td><span class="badge badge-info">#${idx + 1}</span></td>
        <td><strong>${item.part}</strong></td>
        <td><strong style="color: var(--primary);">${val} mins</strong></td>
        <td>${avg} mins</td>
        <td><span class="badge badge-${badgeClass}">${level}</span></td>
      `;
      tbody.appendChild(tr);
    });
  }

  function renderMachineMaintChart(canvasId, machineItems) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[canvasId]) charts[canvasId].destroy();

    charts[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: machineItems.map(m => m.machine),
        datasets: [{
          label: 'Maintenance Downtime (mins)',
          data: machineItems.map(m => m.maintMins),
          backgroundColor: machineItems.map(m => m.color),
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { ticks: { color: currentTheme === 'dark' ? '#ecfdf5' : '#064e3b' } },
          y: { ticks: { color: currentTheme === 'dark' ? '#94a3b8' : '#047857' } }
        }
      }
    });
  }

  function renderMachineMaintTable(tableBodyId, machineItems) {
    const tbody = document.getElementById(tableBodyId);
    if (!tbody) return;
    tbody.innerHTML = '';

    machineItems.forEach((m, idx) => {
      const tr = document.createElement('tr');
      const score = m.maintMins > 200 ? 'Needs Overhaul' : (m.maintMins > 100 ? 'Moderate Repairs' : 'Good Condition');
      const badgeClass = m.maintMins > 200 ? 'danger' : (m.maintMins > 100 ? 'warning' : 'success');

      tr.innerHTML = `
        <td><span class="badge badge-info">#${idx + 1}</span></td>
        <td><strong><i class="fa-solid fa-hammer" style="color: ${m.color};"></i> ${m.machine}</strong></td>
        <td><strong style="color: var(--danger);">${m.maintMins} mins</strong> (${(m.maintMins/60).toFixed(1)} hrs)</td>
        <td>${m.shifts} shifts</td>
        <td><span class="badge badge-${badgeClass}">${score}</span></td>
      `;
      tbody.appendChild(tr);
    });
  }

  /* ==========================================================================
     5-HAMMER VISUAL OEE GAUGE METERS
     ========================================================================== */
  function renderHammerGauges(filteredLogs) {
    const container = document.getElementById('hammerGaugesContainer');
    container.innerHTML = '';

    HAMMERS.forEach((h, idx) => {
      const hLogs = filteredLogs.filter(l => l.machine === h.name);
      
      const stats = calculateFleetKpis(hLogs);
      const oee = stats.oee;
      const avail = stats.avail;
      const perf = stats.perf;
      const qual = stats.qual;

      let statusClass = 'success';
      let statusText = 'RUNNING';
      if (hLogs.length === 0) {
        statusClass = 'info';
        statusText = 'NO SHIFTS';
      } else if (oee < 65) {
        statusClass = 'danger';
        statusText = 'ACTION NEEDED';
      } else if (oee < 75) {
        statusClass = 'warning';
        statusText = 'ATTENTION';
      }

      const card = document.createElement('div');
      card.className = 'gauge-card';
      card.innerHTML = `
        <div class="gauge-header">
          <h4><i class="fa-solid fa-hammer" style="color: ${h.color};"></i> ${h.name}</h4>
          <span class="badge badge-${statusClass}">${statusText}</span>
        </div>

        <div class="gauge-wrapper">
          <canvas id="gaugeCanvas_${idx}" class="gauge-canvas"></canvas>
          <div class="gauge-center-text">
            <div class="gauge-oee-val" style="color: ${hLogs.length > 0 ? getOeeColor(oee) : 'var(--text-muted)'};">${hLogs.length > 0 ? oee.toFixed(1) + '%' : '--%'}</div>
            <div class="gauge-oee-label">OVERALL OEE</div>
          </div>
        </div>

        <div class="gauge-breakdown-row">
          <div class="gauge-sub-stat">
            <span>AVAIL (A)</span>
            <strong style="color: ${hLogs.length > 0 ? getOeeColor(avail) : 'var(--text-muted)'};">${hLogs.length > 0 ? avail.toFixed(1) + '%' : '--%'}</strong>
          </div>
          <div class="gauge-sub-stat">
            <span>PERF (P)</span>
            <strong style="color: ${hLogs.length > 0 ? getOeeColor(perf) : 'var(--text-muted)'};">${hLogs.length > 0 ? perf.toFixed(1) + '%' : '--%'}</strong>
          </div>
          <div class="gauge-sub-stat">
            <span>QUAL (Q)</span>
            <strong style="color: ${hLogs.length > 0 ? getOeeColor(qual) : 'var(--text-muted)'};">${hLogs.length > 0 ? qual.toFixed(1) + '%' : '--%'}</strong>
          </div>
        </div>
      `;

      container.appendChild(card);

      setTimeout(() => {
        drawGaugeCanvas(`gaugeCanvas_${idx}`, hLogs.length > 0 ? oee : 0, h.color);
      }, 50);
    });
  }

  function getOeeColor(val) {
    if (val >= 85) return '#16a34a';
    if (val >= 75) return '#2563eb';
    if (val >= 65) return '#d97706';
    return '#dc2626';
  }

  function drawGaugeCanvas(canvasId, value, primaryColor) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    if (charts[canvasId]) {
      charts[canvasId].destroy();
    }

    const valueColor = getOeeColor(value);

    charts[canvasId] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        datasets: [{
          data: [value, 100 - value],
          backgroundColor: [
            valueColor,
            currentTheme === 'dark' ? '#1b3329' : '#e6f4ea'
          ],
          borderWidth: 0,
          borderRadius: 4
        }]
      },
      options: {
        rotation: -90,
        circumference: 180,
        cutout: '80%',
        responsive: true,
        maintainAspectRatio: false,
        plugins: { tooltip: { enabled: false }, legend: { display: false } }
      }
    });
  }

  /* ==========================================================================
     CHARTS RENDERING (CHART.JS)
     ========================================================================== */
  function renderTrendChart(logs) {
    const ctx = document.getElementById('oeeTrendChart');
    if (!ctx) return;
    if (charts.trend) charts.trend.destroy();

    const dateMap = {};
    logs.forEach(l => {
      if (!dateMap[l.date]) dateMap[l.date] = {};
      if (!dateMap[l.date][l.machine]) dateMap[l.date][l.machine] = [];
      dateMap[l.date][l.machine].push(l.oee);
    });

    const dates = Object.keys(dateMap).sort();

    const datasets = HAMMERS.map(h => {
      const dataPoints = dates.map(d => {
        const hVals = dateMap[d] ? dateMap[d][h.name] : null;
        if (!hVals || hVals.length === 0) return null;
        const avg = hVals.reduce((a, b) => a + b, 0) / hVals.length;
        return parseFloat(avg.toFixed(1));
      });

      return {
        label: h.name,
        data: dataPoints,
        borderColor: h.color,
        backgroundColor: h.color,
        borderWidth: 2,
        tension: 0.3,
        pointRadius: 4,
        spanGaps: true
      };
    });

    datasets.push({
      label: 'Target Benchmark (75%)',
      data: dates.map(() => 75),
      borderColor: '#dc2626',
      borderWidth: 2,
      borderDash: [5, 5],
      pointRadius: 0,
      fill: false
    });

    charts.trend = new Chart(ctx, {
      type: 'line',
      data: { labels: dates, datasets: datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } }
        },
        scales: {
          x: { ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } },
          y: { min: 40, max: 100, ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857', callback: v => v + '%' } }
        }
      }
    });
  }

  function renderComponentsChart(logs) {
    const ctx = document.getElementById('oeeComponentsChart');
    if (!ctx) return;
    if (charts.components) charts.components.destroy();

    const labels = HAMMERS.map(h => h.name);
    const availData = [];
    const perfData = [];
    const qualData = [];

    HAMMERS.forEach(h => {
      const hLogs = logs.filter(l => l.machine === h.name);
      if (hLogs.length === 0) {
        availData.push(0); perfData.push(0); qualData.push(0);
      } else {
        let pMins = 0, oMins = 0, iMins = 0, totalPcs = 0, goodPcs = 0;
        hLogs.forEach(l => {
          pMins += l.plannedTimeMins;
          oMins += l.operatingTimeMins;
          iMins += (l.totalParts * l.idealCycleSec) / 60;
          totalPcs += l.totalParts;
          goodPcs += l.goodParts;
        });
        availData.push(parseFloat(((oMins / pMins) * 100).toFixed(1)));
        perfData.push(parseFloat((Math.min(100, (iMins / oMins) * 100)).toFixed(1)));
        qualData.push(parseFloat(((goodPcs / totalPcs) * 100).toFixed(1)));
      }
    });

    charts.components = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          { label: 'Availability %', data: availData, backgroundColor: '#2563eb', borderRadius: 4 },
          { label: 'Performance %', data: perfData, backgroundColor: '#0284c7', borderRadius: 4 },
          { label: 'Quality %', data: qualData, backgroundColor: '#d97706', borderRadius: 4 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom', labels: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } } },
        scales: {
          x: { ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } },
          y: { min: 0, max: 105, ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857', callback: v => v + '%' } }
        }
      }
    });
  }

  /* ==========================================================================
     MONTH-WISE OEE TREND & PERFORMANCE ANALYSIS (EACH HAMMER & OVERALL)
     ========================================================================== */
  function renderMonthlyTrendView(logs) {
    const ctx = document.getElementById('monthlyTrendChart');
    const tbody = document.getElementById('monthlyTrendTableBody');
    if (!ctx || !tbody) return;

    if (charts.monthlyTrend) charts.monthlyTrend.destroy();

    if (logs.length === 0) {
      tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; color: var(--text-muted); padding: 16px;">Awaiting Excel shift logs for monthly trend calculation.</td></tr>`;
      return;
    }

    // Group logs by YYYY-MM
    const monthlyDataMap = {};

    logs.forEach(l => {
      const monthKey = l.date ? String(l.date).substring(0, 7) : 'Unknown';
      if (!monthlyDataMap[monthKey]) {
        monthlyDataMap[monthKey] = {
          monthKey: monthKey,
          overall: [],
          hammers: {}
        };
        HAMMERS.forEach(h => {
          monthlyDataMap[monthKey].hammers[h.name] = [];
        });
      }

      monthlyDataMap[monthKey].overall.push(l);
      if (monthlyDataMap[monthKey].hammers[l.machine]) {
        monthlyDataMap[monthKey].hammers[l.machine].push(l);
      }
    });

    const sortedMonthKeys = Object.keys(monthlyDataMap).sort();

    const formatMonthLabel = (mKey) => {
      const parts = mKey.split('-');
      if (parts.length !== 2) return mKey;
      const year = parts[0];
      const monthIdx = parseInt(parts[1], 10) - 1;
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${months[monthIdx] || parts[1]} ${year}`;
    };

    const monthLabels = sortedMonthKeys.map(m => formatMonthLabel(m));

    const calcGroupOee = (logArr) => {
      if (!logArr || logArr.length === 0) return null;
      const res = calculateFleetKpis(logArr);
      return {
        avail: res.avail,
        perf: res.perf,
        qual: res.qual,
        oee: res.oee,
        goodPcs: res.goodPcs,
        totalShifts: res.totalShifts
      };
    };

    const overallOeePoints = [];
    const hammerDataPoints = {};
    HAMMERS.forEach(h => { hammerDataPoints[h.name] = []; });

    const monthlyStatsRows = [];

    sortedMonthKeys.forEach(mKey => {
      const mData = monthlyDataMap[mKey];
      const overallRes = calcGroupOee(mData.overall);
      overallOeePoints.push(overallRes ? overallRes.oee : null);

      const rowHammerOees = {};
      HAMMERS.forEach(h => {
        const hRes = calcGroupOee(mData.hammers[h.name]);
        const oeeVal = hRes ? hRes.oee : null;
        hammerDataPoints[h.name].push(oeeVal);
        rowHammerOees[h.name] = oeeVal;
      });

      monthlyStatsRows.push({
        monthLabel: formatMonthLabel(mKey),
        monthKey: mKey,
        totalShifts: overallRes ? overallRes.totalShifts : 0,
        goodPcs: overallRes ? overallRes.goodPcs : 0,
        overallOee: overallRes ? overallRes.oee : 0,
        hammerOees: rowHammerOees
      });
    });

    const datasets = [
      {
        label: '★ Overall Fleet Average OEE',
        data: overallOeePoints,
        borderColor: '#16a34a',
        backgroundColor: 'rgba(22, 163, 74, 0.15)',
        borderWidth: 4,
        tension: 0.3,
        pointRadius: 6,
        pointHoverRadius: 8,
        fill: true,
        spanGaps: true
      }
    ];

    HAMMERS.forEach(h => {
      datasets.push({
        label: h.name,
        data: hammerDataPoints[h.name],
        borderColor: h.color,
        backgroundColor: h.color,
        borderWidth: 2,
        tension: 0.3,
        pointRadius: 4,
        spanGaps: true
      });
    });

    datasets.push({
      label: 'Plant Target Benchmark (75%)',
      data: sortedMonthKeys.map(() => 75),
      borderColor: '#dc2626',
      borderWidth: 2,
      borderDash: [5, 5],
      pointRadius: 0,
      fill: false
    });

    charts.monthlyTrend = new Chart(ctx, {
      type: 'line',
      data: {
        labels: monthLabels,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857', usePointStyle: true, boxWidth: 10 }
          },
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (context) => {
                let label = context.dataset.label || '';
                if (label) label += ': ';
                if (context.parsed.y !== null) label += context.parsed.y + '%';
                return label;
              }
            }
          }
        },
        scales: {
          x: { ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } },
          y: { min: 30, max: 105, ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857', callback: v => v + '%' } }
        }
      }
    });

    tbody.innerHTML = '';
    monthlyStatsRows.slice().reverse().forEach(row => {
      const tr = document.createElement('tr');
      
      const getOeeCellHtml = (val) => {
        if (val === null || val === undefined) return '<span style="color: var(--text-muted);">--</span>';
        return `<strong style="color: ${getOeeColor(val)};">${val}%</strong>`;
      };

      tr.innerHTML = `
        <td><strong>${row.monthLabel}</strong></td>
        <td><span class="badge badge-info">${row.totalShifts} shifts</span></td>
        <td>${row.goodPcs.toLocaleString()} pcs</td>
        <td><strong style="color: ${getOeeColor(row.overallOee)}; font-size: 15px;">${row.overallOee}%</strong></td>
        <td>${getOeeCellHtml(row.hammerOees['1 Ton Hammer'])}</td>
        <td>${getOeeCellHtml(row.hammerOees['1.5 Ton Hammer'])}</td>
        <td>${getOeeCellHtml(row.hammerOees['2.5 Ton (Old) Hammer'])}</td>
        <td>${getOeeCellHtml(row.hammerOees['2.5 Ton (New) Hammer'])}</td>
        <td>${getOeeCellHtml(row.hammerOees['3.5 Ton Hammer'])}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  /* ==========================================================================
     TAB 3: 5-HAMMER COMPARISON VIEW
     ========================================================================== */
  function renderComparisonView(logs) {
    const ctx = document.getElementById('hammerComparisonChart');
    const tableBody = document.getElementById('comparisonTableBody');
    if (!ctx || !tableBody) return;

    if (charts.comparison) charts.comparison.destroy();

    const hammerStats = HAMMERS.map(h => {
      const hLogs = logs.filter(l => l.machine === h.name);
      const res = calculateFleetKpis(hLogs);

      return {
        name: h.name,
        capacity: h.capacity,
        color: h.color,
        shiftsCount: res.totalShifts,
        plannedHrs: (res.plannedMins / 60).toFixed(1),
        operatingHrs: (res.operatingMins / 60).toFixed(1),
        downtimeHrs: ((res.plannedMins - res.operatingMins) / 60).toFixed(1),
        totalPcs: res.totalPcs,
        goodPcs: res.goodPcs,
        rejectsPcs: res.rejectPcs,
        avail: res.avail,
        perf: res.perf,
        qual: res.qual,
        oee: res.oee
      };
    });

    hammerStats.sort((a, b) => b.oee - a.oee);

    charts.comparison = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: hammerStats.map(h => h.name),
        datasets: [
          { label: 'Overall OEE %', data: hammerStats.map(h => h.oee), backgroundColor: hammerStats.map(h => h.color), borderRadius: 6 },
          { label: 'Availability %', data: hammerStats.map(h => h.avail), backgroundColor: 'rgba(37, 99, 235, 0.4)', borderRadius: 4 },
          { label: 'Performance %', data: hammerStats.map(h => h.perf), backgroundColor: 'rgba(2, 132, 199, 0.4)', borderRadius: 4 },
          { label: 'Quality %', data: hammerStats.map(h => h.qual), backgroundColor: 'rgba(217, 119, 6, 0.4)', borderRadius: 4 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom', labels: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } } },
        scales: {
          x: { ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } },
          y: { min: 0, max: 105, ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857', callback: v => v + '%' } }
        }
      }
    });

    tableBody.innerHTML = '';
    hammerStats.forEach((h, rank) => {
      const tr = document.createElement('tr');
      const badgeClass = h.shiftsCount === 0 ? 'info' : (h.oee >= 85 ? 'success' : (h.oee >= 75 ? 'primary' : (h.oee >= 65 ? 'warning' : 'danger')));
      const statusLabel = h.shiftsCount === 0 ? 'No Data' : (h.oee >= 85 ? 'World Class' : (h.oee >= 75 ? 'On Target' : 'Below Target'));

      tr.innerHTML = `
        <td><span class="badge badge-info">#${rank + 1}</span></td>
        <td><strong><i class="fa-solid fa-hammer" style="color: ${h.color};"></i> ${h.name}</strong></td>
        <td>${h.capacity}</td>
        <td>${h.shiftsCount} shifts</td>
        <td>${h.plannedHrs} hrs</td>
        <td>${h.operatingHrs} hrs</td>
        <td><span class="text-danger">${h.downtimeHrs} hrs</span></td>
        <td>${h.totalPcs.toLocaleString()}</td>
        <td>${h.goodPcs.toLocaleString()}</td>
        <td><span class="${h.rejectsPcs > 0 ? 'text-danger' : ''}">${h.rejectsPcs.toLocaleString()}</span></td>
        <td>${h.shiftsCount > 0 ? h.avail + '%' : '--%'}</td>
        <td>${h.shiftsCount > 0 ? h.perf + '%' : '--%'}</td>
        <td>${h.shiftsCount > 0 ? h.qual + '%' : '--%'}</td>
        <td><strong style="color: ${h.shiftsCount > 0 ? getOeeColor(h.oee) : 'var(--text-muted)'}; font-size: 15px;">${h.shiftsCount > 0 ? h.oee + '%' : '--%'}</strong></td>
        <td><span class="badge badge-${badgeClass}">${statusLabel}</span></td>
      `;
      tableBody.appendChild(tr);
    });
  }

  /* ==========================================================================
     TAB 4: DOWNTIME & PARETO ANALYSIS
     ========================================================================== */
  function renderDowntimeView(logs) {
    const paretoCtx = document.getElementById('downtimeParetoChart');
    const hammerDistCtx = document.getElementById('downtimeByHammerChart');
    if (!paretoCtx || !hammerDistCtx) return;

    if (charts.pareto) charts.pareto.destroy();
    if (charts.hammerDist) charts.hammerDist.destroy();

    const categoryTotals = {
      'Maintance': 0,
      'Die Related': 0,
      'Setup': 0,
      'No Manpower': 0,
      'Heating Time': 0,
      'Minor Stop': 0
    };

    logs.forEach(l => {
      categoryTotals['Maintance'] += l.maintanceMins;
      categoryTotals['Die Related'] += l.dieRelatedMins;
      categoryTotals['Setup'] += l.setupMins;
      categoryTotals['No Manpower'] += l.noManpowerMins;
      categoryTotals['Heating Time'] += l.heatingTimeMins;
      categoryTotals['Minor Stop'] += l.minorStopMins;
    });

    document.getElementById('catMaintance').textContent = `${categoryTotals['Maintance']} mins (${(categoryTotals['Maintance']/60).toFixed(1)} hrs)`;
    document.getElementById('catDieRelated').textContent = `${categoryTotals['Die Related']} mins (${(categoryTotals['Die Related']/60).toFixed(1)} hrs)`;
    document.getElementById('catSetup').textContent = `${categoryTotals['Setup']} mins (${(categoryTotals['Setup']/60).toFixed(1)} hrs)`;
    document.getElementById('catNoManpower').textContent = `${categoryTotals['No Manpower']} mins (${(categoryTotals['No Manpower']/60).toFixed(1)} hrs)`;
    document.getElementById('catHeatingTime').textContent = `${categoryTotals['Heating Time']} mins (${(categoryTotals['Heating Time']/60).toFixed(1)} hrs)`;
    document.getElementById('catMinorStop').textContent = `${categoryTotals['Minor Stop']} mins (${(categoryTotals['Minor Stop']/60).toFixed(1)} hrs)`;

    const sortedCategories = Object.keys(categoryTotals)
      .map(cat => ({ category: cat, mins: categoryTotals[cat] }))
      .sort((a, b) => b.mins - a.mins);

    const grandTotalMins = sortedCategories.reduce((acc, curr) => acc + curr.mins, 0);

    let cumulative = 0;
    const labels = [];
    const minsData = [];
    const cumulativeData = [];

    sortedCategories.forEach(item => {
      labels.push(item.category);
      minsData.push(item.mins);
      cumulative += item.mins;
      const pct = grandTotalMins > 0 ? (cumulative / grandTotalMins) * 100 : 0;
      cumulativeData.push(parseFloat(pct.toFixed(1)));
    });

    charts.pareto = new Chart(paretoCtx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            type: 'bar',
            label: 'Downtime (Mins)',
            data: minsData,
            backgroundColor: '#dc2626',
            borderRadius: 6,
            yAxisID: 'y'
          },
          {
            type: 'line',
            label: 'Cumulative Loss %',
            data: cumulativeData,
            borderColor: '#d97706',
            backgroundColor: '#d97706',
            borderWidth: 3,
            tension: 0.2,
            pointRadius: 4,
            yAxisID: 'y1'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom', labels: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } } },
        scales: {
          x: { ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } },
          y: {
            type: 'linear',
            position: 'left',
            title: { display: true, text: 'Minutes', color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' },
            ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' }
          },
          y1: {
            type: 'linear',
            position: 'right',
            min: 0,
            max: 100,
            grid: { drawOnChartArea: false },
            title: { display: true, text: 'Cumulative %', color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' },
            ticks: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857', callback: v => v + '%' }
          }
        }
      }
    });

    const hammerDowntimes = HAMMERS.map(h => {
      const hLogs = logs.filter(l => l.machine === h.name);
      return hLogs.reduce((acc, curr) => acc + curr.totalDowntimeMins, 0);
    });

    charts.hammerDist = new Chart(hammerDistCtx, {
      type: 'doughnut',
      data: {
        labels: HAMMERS.map(h => h.name),
        datasets: [{
          data: hammerDowntimes,
          backgroundColor: HAMMERS.map(h => h.color),
          borderWidth: 2,
          borderColor: currentTheme === 'dark' ? '#12221b' : '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'right', labels: { color: currentTheme === 'dark' ? '#6ee7b7' : '#047857' } } }
      }
    });
  }

  /* ==========================================================================
     TAB 5: MANUAL ENTRY FORM & LIVE CALCULATOR
     ========================================================================== */
  function setupLiveCalculator() {
    const inputs = [
      'entryGrossPlanned', 'entryLunchBreak', 'entryTeaBreak',
      'entryMaintance', 'entryDieRelated', 'entrySetup',
      'entryNoManpower', 'entryHeatingTime', 'entryMinorStop',
      'entryIdealCycle', 'entryTotalParts', 'entryRejects'
    ];

    inputs.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('input', calculateLiveFormValues);
    });

    const dateInput = document.getElementById('entryDate');
    if (dateInput) dateInput.value = new Date().toISOString().split('T')[0];

    const form = document.getElementById('manualEntryForm');
    if (form) {
      form.addEventListener('submit', handleManualFormSubmit);
    }
  }

  function calculateLiveFormValues() {
    const gross = parseNum(document.getElementById('entryGrossPlanned').value, 720);
    const lunch = parseNum(document.getElementById('entryLunchBreak').value, 30);
    const tea = parseNum(document.getElementById('entryTeaBreak').value, 30);
    const netPlanned = Math.max(0, gross - lunch - tea);

    document.getElementById('entryPlannedTime').value = netPlanned;

    const rawData = {
      grossShiftMins: gross,
      lunchBreakMins: lunch,
      teaBreakMins: tea,
      plannedTimeMins: netPlanned,
      maintanceMins: document.getElementById('entryMaintance').value,
      dieRelatedMins: document.getElementById('entryDieRelated').value,
      setupMins: document.getElementById('entrySetup').value,
      noManpowerMins: document.getElementById('entryNoManpower').value,
      heatingTimeMins: document.getElementById('entryHeatingTime').value,
      minorStopMins: document.getElementById('entryMinorStop').value,
      idealCycleSec: document.getElementById('entryIdealCycle').value,
      totalParts: document.getElementById('entryTotalParts').value,
      rejects: document.getElementById('entryRejects').value
    };

    const res = calculateOeeRecord(rawData);

    document.getElementById('entryGoodParts').value = res.goodParts;
    document.getElementById('calcTotalDowntime').textContent = `${res.totalDowntimeMins} mins`;
    document.getElementById('calcOperatingTime').textContent = `${res.operatingTimeMins} mins`;
    
    updateCalcElement('calcAvailability', res.availability);
    updateCalcElement('calcPerformance', res.performance);
    updateCalcElement('calcQuality', res.quality);
    updateCalcElement('calcOee', res.oee);
  }

  function updateCalcElement(id, value) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = `${value}%`;
    el.className = value >= 85 ? 'val-good' : (value >= 75 ? 'val-good' : (value >= 65 ? 'val-warn' : 'val-bad'));
  }

  function handleManualFormSubmit(e) {
    e.preventDefault();

    const machine = document.getElementById('entryMachine').value;
    const date = document.getElementById('entryDate').value;
    const shift = document.getElementById('entryShift').value;
    const partNumber = document.getElementById('entryPart').value || 'A1#21';

    if (!machine || !date) {
      showToast('Please select Machine and Date.', 'danger');
      return;
    }

    const gross = parseNum(document.getElementById('entryGrossPlanned').value, 720);
    const lunch = parseNum(document.getElementById('entryLunchBreak').value, 30);
    const tea = parseNum(document.getElementById('entryTeaBreak').value, 30);

    const rawRecord = {
      id: 'LOG-' + Math.random().toString(36).substr(2, 8).toUpperCase(),
      date: date,
      shift: shift,
      machine: machine,
      partNumber: partNumber,
      grossShiftMins: gross,
      lunchBreakMins: lunch,
      teaBreakMins: tea,
      plannedTimeMins: Math.max(0, gross - lunch - tea),
      maintanceMins: parseNum(document.getElementById('entryMaintance').value, 0),
      dieRelatedMins: parseNum(document.getElementById('entryDieRelated').value, 0),
      setupMins: parseNum(document.getElementById('entrySetup').value, 0),
      noManpowerMins: parseNum(document.getElementById('entryNoManpower').value, 0),
      heatingTimeMins: parseNum(document.getElementById('entryHeatingTime').value, 0),
      minorStopMins: parseNum(document.getElementById('entryMinorStop').value, 0),
      totalParts: parseNum(document.getElementById('entryTotalParts').value, 0),
      rejects: parseNum(document.getElementById('entryRejects').value, 0),
      idealCycleSec: parseNum(document.getElementById('entryIdealCycle').value, 45)
    };

    const computed = calculateOeeRecord(rawRecord);

    shiftLogs.unshift(computed);
    saveShiftLogs();

    showToast(`Shift record saved for ${machine}! OEE: ${computed.oee}%`, 'success');
    
    renderAllViews();
    switchTab('overview');
  }

  /* ==========================================================================
     TAB 6: EXCEL INTEGRATION (AUTO-COMMIT UPON FILE SELECTION / DROP)
     ========================================================================== */
  function setupExcelDropZone() {
    const dropZone = document.getElementById('excelDropZone');
    const fileInput = document.getElementById('excelFileInput');
    const hammerFileInput = document.getElementById('hammerFileInput');

    if (dropZone && fileInput) {
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
          currentImportHammer = null; // Combined mode
          processExcelFile(e.dataTransfer.files[0], null);
        }
      });

      fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          currentImportHammer = null; // Combined mode
          processExcelFile(e.target.files[0], null);
        }
      });
    }

    if (hammerFileInput) {
      hammerFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          processExcelFile(e.target.files[0], currentImportHammer);
        }
      });
    }

    document.getElementById('downloadXlsxTemplateBtn').addEventListener('click', downloadExcelTemplate);
    document.getElementById('downloadCsvTemplateBtn').addEventListener('click', downloadCsvTemplate);

    document.getElementById('exportFullExcelBtn').addEventListener('click', exportFullLogsExcel);
    document.getElementById('exportSummaryExcelBtn').addEventListener('click', exportSummaryExcel);
    document.getElementById('exportLogsCsvBtn').addEventListener('click', exportLogsCsv);

   const reloadSampleDataBtn = document.getElementById('reloadSampleDataBtn');

if (reloadSampleDataBtn) {
    reloadSampleDataBtn.addEventListener('click', () => {
        shiftLogs = generateDefaultLogs();
        saveShiftLogs();
        renderAllViews();
        showToast('Reloaded sample historical data.', 'info');
    });
}

    const clearAllLogsAction = () => {
      shiftLogs = [];
      saveShiftLogs();
      renderAllViews();
      showToast('All data cleared 100%! Ready for your Excel uploads.', 'warning');
    };

    document.getElementById('clearAllLogsBtn').addEventListener('click', () => {
      if (confirm('Clear all log data and start completely fresh for your Excel uploads?')) {
        clearAllLogsAction();
      }
    });
    
    const clearDemoHeaderBtn = document.getElementById('clearDemoDataHeaderBtn');
    if (clearDemoHeaderBtn) {
      clearDemoHeaderBtn.addEventListener('click', () => {
        if (confirm('Clear all data 100% and start fresh?')) {
          clearAllLogsAction();
        }
      });
    }

    document.getElementById('commitExcelImportBtn').addEventListener('click', () => {
      showToast('Excel shift logs are active!', 'success');
      switchTab('overview');
    });

    document.getElementById('cancelExcelImportBtn').addEventListener('click', () => {
      document.getElementById('excelPreviewCard').style.display = 'none';
    });
  }

  function setupHammerMiniDropzones() {
    document.querySelectorAll('.excel-mini-dropzone').forEach(zone => {
      const targetHammer = zone.getAttribute('data-hammer');

      zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('drag-over');
      });

      zone.addEventListener('dragleave', () => {
        zone.classList.remove('drag-over');
      });

      zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
          currentImportHammer = targetHammer;
          processExcelFile(e.dataTransfer.files[0], targetHammer);
        }
      });
    });
  }

  function processExcelFile(file, hammerNameOverride) {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (jsonData.length === 0) {
          showToast('The uploaded Excel sheet contains no data rows.', 'danger');
          return;
        }

        const parsedLogs = parsePlantExcelJsonToLogs(jsonData, hammerNameOverride);

        if (parsedLogs.length === 0) {
          showToast('Could not parse any records from the Excel sheet.', 'danger');
          return;
        }

        // AUTO-COMMIT IMMEDIATELY UPON FILE LOAD
        if (hammerNameOverride) {
          const otherHammerLogs = shiftLogs.filter(l => l.machine !== hammerNameOverride);
          shiftLogs = [...parsedLogs, ...otherHammerLogs].sort((a, b) => new Date(b.date) - new Date(a.date));
        } else {
          shiftLogs = [...parsedLogs].sort((a, b) => new Date(b.date) - new Date(a.date));
        }

        // Set global date range filter to "ALL" so all uploaded dates show immediately
        const rangeSelect = document.getElementById('globalRangeFilter');
        if (rangeSelect) rangeSelect.value = 'ALL';

        saveShiftLogs();
        renderAllViews();
        renderExcelPreview(parsedLogs, hammerNameOverride);

        showToast(`Successfully imported ${parsedLogs.length} shift logs for ${hammerNameOverride || 'Combined Fleet'}! OEE updated.`, 'success');

        // Automatically switch to Overview tab so user sees OEE metrics right away!
        switchTab('overview');
      } catch (err) {
        console.error(err);
        showToast('Error parsing Excel file. Please check format.', 'danger');
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function parsePlantExcelJsonToLogs(rows, hammerNameOverride) {
    const parsed = [];

    const partNumberCandidates = [
      'part number', 'part_number', 'partnumber', 'part no', 'part no.', 'part_no', 'partno',
      'part #', 'part#', 'part code', 'item code', 'item no', 'component no', 'component',
      'drawing no', 'drawing number', 'part name', 'part', 'item', 'fg code', 'part_id'
    ];

    const machineCandidates = ['machine', 'hammer', 'equipment', 'machine name', 'machine_name', 'hammer name', 'press', 'line'];
    const dateCandidates = ['date', 'production date', 'shift date', 'date_str', 'dt', 'day'];
    const shiftCandidates = ['shift', 'work shift', 'shift_name', 'shift a/b', 'shift type'];

    const plannedTimeCandidates = ['planned time', 'planned_time', 'shift duration', 'shift time', 'planned time (mins)', 'net planned time', 'planned', 'planned mins'];
    const maintanceCandidates = ['maintance', 'maintenance', 'maintance (mins)', 'maintenance (mins)', 'maint', 'breakdown', 'maint (mins)', 'bd (mins)'];
    const dieRelatedCandidates = ['die related', 'die_related', 'dierelated', 'die downtime', 'die (mins)', 'die issue', 'die issues', 'die change', 'die (min)'];
    const setupCandidates = ['setup', 'setup time', 'setup (mins)', 'setup_time', 'setting time', 'changeover', 'setup (min)'];
    const noManpowerCandidates = ['no manpower', 'no_manpower', 'manpower', 'no manpower (mins)', 'manpower shortage', 'operator shortage', 'no manpower (min)'];
    const heatingTimeCandidates = ['heating time', 'heating_time', 'heating', 'heating time (mins)', 'furnace heating', 'heating wait', 'heating (min)'];
    const minorStopCandidates = ['minor stop', 'minor_stop', 'minor stops', 'minor stop (mins)', 'minor delays', 'short stops', 'minor (min)'];

    const totalDowntimeCandidates = ['total downtime', 'total_downtime', 'downtime', 'total downtime (mins)', 'total loss', 'downtime (mins)'];
    const operatingTimeCandidates = ['operating time', 'operating_time', 'run time', 'working time', 'production time', 'operating time (mins)', 'opt time'];

    const totalPartsCandidates = ['total parts', 'total_parts', 'total production', 'total produced', 'qty', 'output', 'produced parts', 'total pcs', 'total parts good parts', 'total parts (pcs)'];
    const goodPartsCandidates = ['good parts', 'good_parts', 'good production', 'accepted parts', 'ok parts', 'good qty', 'ok qty', 'good pcs', 'good'];
    const rejectsCandidates = ['rejects', 'reject', 'scrap', 'defects', 'rejected parts', 'rejections', 'rejection', 'scrap qty', 'bad parts', 'rej'];
    const idealCycleCandidates = ['ideal cycle time', 'ideal_cycle_time', 'ideal cycle', 'cycle time', 'std cycle time', 'ideal cycle (sec)', 'cycle time sec', 'ct'];

    const availabilityCandidates = ['availability', 'availability (%)', 'avail', 'avail %', 'availability_pct', 'availability rate'];
    const performanceCandidates = ['performance', 'performance (%)', 'perf', 'perf %', 'performance_pct', 'performance rate'];
    const qualityCandidates = ['quality', 'quality (%)', 'qual', 'qual %', 'quality_pct', 'quality rate'];
    const oeeCandidates = ['oee', 'overall oee', 'oee (%)', 'oee %', 'overall_oee'];

    rows.forEach(r => {
      const rawDate = findVal(r, dateCandidates);
      const date = formatExcelDate(rawDate);
      
      let rawShift = findVal(r, shiftCandidates) || 'Shift A';
      let shift = String(rawShift).trim();
      if (shift.toLowerCase().includes('b') || shift.toLowerCase().includes('2')) {
        shift = 'Shift B';
      } else {
        shift = 'Shift A';
      }

      const rawMachine = findVal(r, machineCandidates);
      const machine = hammerNameOverride ? hammerNameOverride : normalizeHammerName(rawMachine || '1 Ton Hammer');
      
      const rawPart = findVal(r, partNumberCandidates);
      const partNumber = (rawPart !== null && rawPart !== undefined && String(rawPart).trim() !== '') ? 
        String(rawPart).trim() : 'A1#21';

      let rawPlanned = parseNum(findVal(r, plannedTimeCandidates), 660);
      // Unit normalization: If specified in hours (e.g. 11h, 12h)
      if (rawPlanned > 0 && rawPlanned <= 24) {
        if (rawPlanned >= 11.5) {
          rawPlanned = 720;
        } else {
          rawPlanned = rawPlanned * 60;
        }
      }
      let netPlanned = rawPlanned >= 700 ? 660 : (rawPlanned > 0 ? rawPlanned : 660);

      const maintance = parseTimeMins(findVal(r, maintanceCandidates), true);
      const dieRelated = parseTimeMins(findVal(r, dieRelatedCandidates), true);
      const setup = parseTimeMins(findVal(r, setupCandidates), true);
      const noManpower = parseTimeMins(findVal(r, noManpowerCandidates), true);
      const heatingTime = parseTimeMins(findVal(r, heatingTimeCandidates), true);
      const minorStop = parseTimeMins(findVal(r, minorStopCandidates), true);

      const computedTotalDowntime = maintance + dieRelated + setup + noManpower + heatingTime + minorStop;
      const rawDowntime = findVal(r, totalDowntimeCandidates);
      const totalDowntime = rawDowntime !== null ? parseTimeMins(rawDowntime, true) : computedTotalDowntime;

      const rawOpt = findVal(r, operatingTimeCandidates);
      const operatingTime = rawOpt !== null ? parseTimeMins(rawOpt, true) : Math.max(0, netPlanned - totalDowntime);

      const totalParts = parseNum(findVal(r, totalPartsCandidates), 0);
      const goodPartsVal = findVal(r, goodPartsCandidates);
      const rejectsVal = findVal(r, rejectsCandidates);

      const rejects = rejectsVal !== null ? parseNum(rejectsVal, 0) : Math.max(0, totalParts - (goodPartsVal !== null ? parseNum(goodPartsVal, totalParts) : totalParts));
      const goodParts = goodPartsVal !== null ? parseNum(goodPartsVal, Math.max(0, totalParts - rejects)) : Math.max(0, totalParts - rejects);
      
      let idealCycle = parseNum(findVal(r, idealCycleCandidates), 0);
      if (idealCycle <= 0) {
        const hammerObj = HAMMERS.find(h => h.name === machine);
        idealCycle = hammerObj ? hammerObj.defaultCycle : 45;
      } else if (idealCycle < 3) {
        // If specified in MINUTES per piece (e.g. 0.75 min) => convert to SECONDS (0.75 * 60 = 45s)
        idealCycle = idealCycle * 60;
      } else if (idealCycle > 300) {
        // If specified in parts per hour (e.g. 80 pcs/hr) => convert to sec/pc (3600 / 80 = 45s)
        idealCycle = 3600 / idealCycle;
      }

      const rawRecord = {
        id: 'IMP-' + Math.random().toString(36).substr(2, 8).toUpperCase(),
        date: String(date).trim(),
        shift: shift,
        machine: machine,
        partNumber: partNumber,
        grossShiftMins: 720,
        lunchBreakMins: 30,
        teaBreakMins: 30,
        plannedTimeMins: netPlanned,
        maintanceMins: maintance,
        dieRelatedMins: dieRelated,
        setupMins: setup,
        noManpowerMins: noManpower,
        heatingTimeMins: heatingTime,
        minorStopMins: minorStop,
        totalDowntimeMins: totalDowntime,
        operatingTimeMins: operatingTime,
        totalParts: totalParts,
        goodParts: goodParts,
        rejects: rejects,
        idealCycleSec: idealCycle
      };

      parsed.push(calculateOeeRecord(rawRecord));
    });

    return parsed;
  }

  let excelValidNewRecords = [];

  function renderExcelPreview(records, hammerNameOverride) {
    const card = document.getElementById('excelPreviewCard');
    const table = document.getElementById('excelPreviewTable');
    document.getElementById('excelPreviewTargetLabel').textContent = hammerNameOverride || 'Combined Fleet Sheet';

    card.style.display = 'block';

    let newCount = 0;
    let duplicateCount = 0;
    let errorCount = 0;

    excelValidNewRecords = [];

    const existingKeys = new Set(shiftLogs.map(l => `${l.date}_${l.shift}_${l.machine}_${l.partNumber}`));

    records.forEach(r => {
      const key = `${r.date}_${r.shift}_${r.machine}_${r.partNumber}`;
      if (!r.date || !r.partNumber || r.totalParts < 0) {
        errorCount++;
      } else if (existingKeys.has(key)) {
        duplicateCount++;
      } else {
        newCount++;
        excelValidNewRecords.push(r);
      }
    });

    const newEl = document.getElementById('catNewRecordsCount');
    const dupEl = document.getElementById('catDuplicateRecordsCount');
    const errEl = document.getElementById('catErrorRecordsCount');

    if (newEl) newEl.textContent = newCount;
    if (dupEl) dupEl.textContent = duplicateCount;
    if (errEl) errEl.textContent = errorCount;

    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    thead.innerHTML = `
      <tr>
        <th>Validation Status</th>
        <th>Date</th>
        <th>Shift</th>
        <th>Machine</th>
        <th>Part Number</th>
        <th>Planned Mins</th>
        <th>Downtime Mins</th>
        <th>Total Parts</th>
        <th>Good</th>
        <th>Rejects</th>
        <th>OEE %</th>
      </tr>
    `;

    tbody.innerHTML = '';
    records.slice(0, 15).forEach(r => {
      const key = `${r.date}_${r.shift}_${r.machine}_${r.partNumber}`;
      let statusBadge = '<span class="badge badge-success"><i class="fa-solid fa-check"></i> New Valid</span>';

      if (!r.date || !r.partNumber) {
        statusBadge = '<span class="badge badge-danger"><i class="fa-solid fa-xmark"></i> Invalid</span>';
      } else if (existingKeys.has(key)) {
        statusBadge = '<span class="badge badge-warning"><i class="fa-solid fa-copy"></i> Duplicate</span>';
      }

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${statusBadge}</td>
        <td>${r.date}</td>
        <td>${r.shift}</td>
        <td>${r.machine}</td>
        <td><strong>${r.partNumber}</strong></td>
        <td>${r.plannedTimeMins}m</td>
        <td>${r.totalDowntimeMins}m</td>
        <td>${r.totalParts}</td>
        <td>${r.goodParts}</td>
        <td>${r.rejects}</td>
        <td><strong style="color: ${getOeeColor(r.oee)}; font-weight: bold; font-size: 14px;">${r.oee}%</strong></td>
      `;
      tbody.appendChild(tr);
    });
  }

  async function commitExcelImportToSupabase() {
    if (excelValidNewRecords.length === 0) {
      showToast('No new valid records to import.', 'warning');
      return;
    }

    const userId = currentAuthUser ? currentAuthUser.id : null;

    if (supabaseClient) {
      try {
        const prodRows = excelValidNewRecords.map(r => ({
          date: r.date,
          shift: r.shift,
          hammer: r.machine,
          part_number: r.partNumber,
          planned_time_mins: r.plannedTimeMins,
          planned_qty: r.totalParts,
          production_qty: r.totalParts,
          good_qty: r.goodParts,
          created_by: userId
        }));

        const { error: prodErr } = await supabaseClient.from('production_data').insert(prodRows);
        if (prodErr) throw prodErr;

        const qualRows = excelValidNewRecords.map(r => ({
          date: r.date,
          shift: r.shift,
          hammer: r.machine,
          part_number: r.partNumber,
          inspection_stage: 'In-Process',
          inspection_qty: r.totalParts,
          rework_qty: r.rework || 0,
          rejection_qty: r.rejects,
          created_by: userId
        }));

        await supabaseClient.from('quality_data').insert(qualRows);

        showToast(`Successfully inserted ${excelValidNewRecords.length} records into Supabase!`, 'success');
        document.getElementById('excelPreviewCard').style.display = 'none';
        fetchSupabaseShiftLogs();
      } catch (err) {
        console.error('Excel Import Supabase Error:', err);
        showToast(`Import Error: ${err.message}`, 'danger');
      }
    } else {
      shiftLogs = mergeLogArrays(shiftLogs, excelValidNewRecords);
      saveShiftLogs();
      renderAllViews();
      document.getElementById('excelPreviewCard').style.display = 'none';
      showToast(`Imported ${excelValidNewRecords.length} records locally.`, 'success');
    }
  }

  function downloadHammerSpecificTemplate(hammerName) {
    const hammerObj = HAMMERS.find(h => h.name === hammerName) || HAMMERS[0];
    const sampleRows = [
      {
        "Date": "2026-07-25",
        "Shift": "Shift A",
        "Machine": hammerObj.name,
        "part number": hammerObj.samplePart,
        "Planned time": 660,
        "Maintance": 15,
        "die related": 20,
        "setup": 15,
        "No manpower": 0,
        "Heating time": 10,
        "minor stop": 10,
        "total downtime": 70,
        "operating time": 590,
        "total parts": 800,
        "good parts": 785,
        "rejects": 15,
        "ideal cycle time": hammerObj.defaultCycle,
        "Availability": 89.4,
        "Performance": 94.9,
        "Quality": 98.1,
        "OEE": 83.2
      },
      {
        "Date": "2026-07-25",
        "Shift": "Shift B",
        "Machine": hammerObj.name,
        "part number": hammerObj.samplePart,
        "Planned time": 660,
        "Maintance": 25,
        "die related": 15,
        "setup": 20,
        "No manpower": 0,
        "Heating time": 15,
        "minor stop": 10,
        "total downtime": 85,
        "operating time": 575,
        "total parts": 760,
        "good parts": 745,
        "rejects": 15,
        "ideal cycle time": hammerObj.defaultCycle,
        "Availability": 87.1,
        "Performance": 92.5,
        "Quality": 98.0,
        "OEE": 79.0
      }
    ];

    const ws = XLSX.utils.json_to_sheet(sampleRows);
    const wb = XLSX.utils.book_new();
    const cleanFileName = hammerObj.name.replace(/[^a-zA-Z0-9]/g, '_');
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, `${cleanFileName}_OEE_Template.xlsx`);
  }

  function downloadExcelTemplate() {
    const sampleRows = [
      {
        "Date": "2026-07-25",
        "Shift": "Shift A",
        "Machine": "1 Ton Hammer",
        "part number": "A1#21",
        "Planned time": 660,
        "Maintance": 20,
        "die related": 25,
        "setup": 15,
        "No manpower": 0,
        "Heating time": 15,
        "minor stop": 10,
        "total downtime": 85,
        "operating time": 575,
        "total parts": 900,
        "good parts": 885,
        "rejects": 15,
        "ideal cycle time": 35,
        "Availability": 87.1,
        "Performance": 91.3,
        "Quality": 98.3,
        "OEE": 78.2
      },
      {
        "Date": "2026-07-25",
        "Shift": "Shift B",
        "Machine": "2.5 Ton (Old) Hammer",
        "part number": "A4#07",
        "Planned time": 660,
        "Maintance": 70,
        "die related": 25,
        "setup": 65,
        "No manpower": 0,
        "Heating time": 15,
        "minor stop": 10,
        "total downtime": 185,
        "operating time": 475,
        "total parts": 520,
        "good parts": 505,
        "rejects": 15,
        "ideal cycle time": 55,
        "Availability": 72.0,
        "Performance": 100.0,
        "Quality": 97.1,
        "OEE": 69.9
      }
    ];

    const ws = XLSX.utils.json_to_sheet(sampleRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Machine_OEE_Template");
    XLSX.writeFile(wb, "Forge_Plant_OEE_12h_Net660_Template.xlsx");
  }

  function downloadCsvTemplate() {
    const headers = "Date,Shift,Machine,part number,Planned time,Maintance,die related,setup,No manpower,Heating time,minor stop,total downtime,operating time,total parts,good parts,rejects,ideal cycle time,Availability,Performance,Quality,OEE\n";
    const row1 = "2026-07-25,Shift A,1 Ton Hammer,A1#21,660,20,25,15,0,15,10,85,575,900,885,15,35,87.1,91.3,98.3,78.2\n";
    const row2 = "2026-07-25,Shift B,3.5 Ton Hammer,M5#102,660,20,25,20,0,50,15,130,530,480,465,15,70,80.3,100.0,96.9,77.8\n";
    
    downloadFile(headers + row1 + row2, "Forge_Plant_OEE_12h_Net660_Template.csv", "text/csv");
  }

  function exportFullLogsExcel() {
    const logs = getFilteredLogs();
    if (logs.length === 0) {
      showToast('No logs available to export.', 'warning');
      return;
    }

    const exportRows = logs.map(l => ({
      "Date": l.date,
      "Shift": l.shift,
      "Machine": l.machine,
      "part number": l.partNumber,
      "Planned time": l.plannedTimeMins,
      "Maintance": l.maintanceMins,
      "die related": l.dieRelatedMins,
      "setup": l.setupMins,
      "No manpower": l.noManpowerMins,
      "Heating time": l.heatingTimeMins,
      "minor stop": l.minorStopMins,
      "total downtime": l.totalDowntimeMins,
      "operating time": l.operatingTimeMins,
      "total parts": l.totalParts,
      "good parts": l.goodParts,
      "rejects": l.rejects,
      "ideal cycle time": l.idealCycleSec,
      "Availability": l.availability,
      "Performance": l.performance,
      "Quality": l.quality,
      "OEE": l.oee
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Machine_Shift_Logs");
    XLSX.writeFile(wb, `Forge_Machine_Shift_Logs_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  function exportSummaryExcel() {
    const logs = getFilteredLogs();
    const summaryRows = HAMMERS.map(h => {
      const hLogs = logs.filter(l => l.machine === h.name);
      let pMins = 0, oMins = 0, iMins = 0, downtime = 0, totalPcs = 0, goodPcs = 0, rejects = 0;

      hLogs.forEach(l => {
        pMins += l.plannedTimeMins;
        oMins += l.operatingTimeMins;
        downtime += l.totalDowntimeMins;
        iMins += (l.totalParts * l.idealCycleSec) / 60;
        totalPcs += l.totalParts;
        goodPcs += l.goodParts;
        rejects += l.rejects;
      });

      const avail = pMins > 0 ? (oMins / pMins) * 100 : 0;
      const perf = oMins > 0 ? Math.min(120, (iMins / oMins) * 100) : 0;
      const qual = totalPcs > 0 ? (goodPcs / totalPcs) * 100 : 100;
      const oee = (avail / 100) * (perf / 100) * (qual / 100) * 100;

      return {
        "Machine / Hammer": h.name,
        "Capacity": h.capacity,
        "Shift Count": hLogs.length,
        "Net Planned Hrs": (pMins / 60).toFixed(1),
        "Operating Hrs": (oMins / 60).toFixed(1),
        "Downtime Hrs": (downtime / 60).toFixed(1),
        "Total Parts": totalPcs,
        "Good Parts": goodPcs,
        "Rejects": rejects,
        "Availability (%)": avail.toFixed(1),
        "Performance (%)": perf.toFixed(1),
        "Quality (%)": qual.toFixed(1),
        "Overall OEE (%)": oee.toFixed(1)
      };
    });

    const ws = XLSX.utils.json_to_sheet(summaryRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "5_Machine_Summary");
    XLSX.writeFile(wb, `5_Machine_OEE_Summary_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  function exportLogsCsv() {
    const logs = getFilteredLogs();
    let csv = "Date,Shift,Machine,part number,Planned time,Maintance,die related,setup,No manpower,Heating time,minor stop,total downtime,operating time,total parts,good parts,rejects,ideal cycle time,Availability,Performance,Quality,OEE\n";
    logs.forEach(l => {
      csv += `${l.date},${l.shift},"${l.machine}","${l.partNumber}",${l.plannedTimeMins},${l.maintanceMins},${l.dieRelatedMins},${l.setupMins},${l.noManpowerMins},${l.heatingTimeMins},${l.minorStopMins},${l.totalDowntimeMins},${l.operatingTimeMins},${l.totalParts},${l.goodParts},${l.rejects},${l.idealCycleSec},${l.availability},${l.performance},${l.quality},${l.oee}\n`;
    });
    downloadFile(csv, `Machine_Shift_Logs_${new Date().toISOString().split('T')[0]}.csv`, "text/csv");
  }

  function downloadFile(content, fileName, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);
  }

  /* ==========================================================================
     TAB 7: MASTER SHIFT LOGS TABLE RENDERER
     ========================================================================== */
  function renderLogsTable(logs) {
    const tbody = document.getElementById('logsTableBody');
    const counter = document.getElementById('tableRecordCounter');
    const searchVal = (document.getElementById('logSearchInput').value || '').toLowerCase();

    if (!tbody) return;

    const filtered = logs.filter(l => {
      if (!searchVal) return true;
      return (
        l.machine.toLowerCase().includes(searchVal) ||
        l.partNumber.toLowerCase().includes(searchVal) ||
        l.shift.toLowerCase().includes(searchVal)
      );
    });

    tbody.innerHTML = '';

    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="22" style="text-align: center; color: var(--text-muted); padding: 24px; font-weight: 600;">No data available for the selected Machine and Month</td></tr>`;
      counter.textContent = 'Showing 0 records';
      return;
    }

    filtered.forEach(l => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${l.date}</td>
        <td><span class="badge badge-info">${l.shift}</span></td>
        <td><strong>${l.machine}</strong></td>
        <td><strong style="color: var(--primary);">${l.partNumber}</strong></td>
        <td>${l.plannedTimeMins}m</td>
        <td>${l.maintanceMins}m</td>
        <td>${l.dieRelatedMins}m</td>
        <td>${l.setupMins}m</td>
        <td>${l.noManpowerMins}m</td>
        <td>${l.heatingTimeMins}m</td>
        <td>${l.minorStopMins}m</td>
        <td><span class="text-danger">${l.totalDowntimeMins}m</span></td>
        <td>${l.operatingTimeMins}m</td>
        <td>${l.totalParts.toLocaleString()}</td>
        <td>${l.goodParts.toLocaleString()}</td>
        <td><span class="${l.rejects > 0 ? 'text-danger' : ''}">${l.rejects}</span></td>
        <td>${l.idealCycleSec}s</td>
        <td>${l.availability}%</td>
        <td>${l.performance}%</td>
        <td>${l.quality}%</td>
        <td><strong style="color: ${getOeeColor(l.oee)}; font-size: 14px;">${l.oee}%</strong></td>
        <td>
          ${window.IS_VIEW_ONLY ? '<span class="badge badge-info"><i class="fa-solid fa-lock"></i> Read Only</span>' : `
          <button class="btn btn-danger btn-sm delete-log-btn" data-id="${l.id}" title="Delete Record">
            <i class="fa-solid fa-trash"></i>
          </button>`}
        </td>
      `;
      tbody.appendChild(tr);
    });

    counter.textContent = `Showing ${filtered.length} of ${logs.length} records`;

    document.querySelectorAll('.delete-log-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-id');
        if (confirm('Delete this shift log entry?')) {
          shiftLogs = shiftLogs.filter(item => item.id !== id);
          saveShiftLogs();
          renderAllViews();
          showToast('Record deleted.', 'info');
        }
      });
    });
  }
  /* ==========================================================================
     EVENT LISTENERS & TAB NAVIGATION
     ========================================================================== */
  async function migrateLocalLogsToSupabase() {
    if (!supabaseClient) {
        showToast('Supabase is not connected. Please check DB Config.', 'danger');
        return;
    }

    if (!Array.isArray(shiftLogs) || shiftLogs.length === 0) {
        showToast('No local shift log data available to migrate.', 'warning');
        return;
    }

    const userId =
        currentAuthUser?.id ||
        currentUserProfile?.id ||
        'migration';

    try {
        showToast(`Migrating ${shiftLogs.length} local records to Supabase...`, 'info');

        const productionRows = shiftLogs.map(r => ({
            date: r.date,
            shift: r.shift,
            hammer: r.machine,
            part_number: r.partNumber,
            planned_time_mins: Number(r.plannedTimeMins || 0),
            planned_qty: Number(r.totalParts || 0),
            production_qty: Number(r.totalParts || 0),
            good_qty: Number(r.goodParts || 0),
            created_by: userId
        }));

        if (productionRows.length) {
            const { error } = await supabaseClient
                .from('production_data')
                .insert(productionRows);

            if (error) throw error;
        }

        const qualityRows = shiftLogs.map(r => ({
            date: r.date,
            shift: r.shift,
            hammer: r.machine,
            part_number: r.partNumber,
            inspection_stage: 'In-Process',
            inspection_qty: Number(r.totalParts || 0),
            rework_qty: Number(r.rework || 0),
            rejection_qty: Number(r.rejects || 0),
            created_by: userId
        }));

        if (qualityRows.length) {
            const { error } = await supabaseClient
                .from('quality_data')
                .insert(qualityRows);

            if (error) throw error;
        }

        const downtimeRows = [];

        shiftLogs.forEach(r => {
            [
                ['Die Related', r.dieRelatedMins],
                ['Setup', r.setupMins],
                ['No Manpower', r.noManpowerMins],
                ['Heating Time', r.heatingTimeMins],
                ['Minor Stop', r.minorStopMins]
            ].forEach(([category, mins]) => {
                const minutes = Number(mins || 0);

                if (minutes > 0) {
                    downtimeRows.push({
                        date: r.date,
                        shift: r.shift,
                        hammer: r.machine,
                        part_number: r.partNumber,
                        downtime_category: category,
                        downtime_minutes: minutes,
                        created_by: userId
                    });
                }
            });
        });

        if (downtimeRows.length) {
            const { error } = await supabaseClient
                .from('downtime_data')
                .insert(downtimeRows);

            if (error) throw error;
        }

        console.log('Migration completed:', {
            production: productionRows.length,
            quality: qualityRows.length,
            downtime: downtimeRows.length
        });

        showToast(
            `Migration successful! Production: ${productionRows.length}, Quality: ${qualityRows.length}, Downtime: ${downtimeRows.length}`,
            'success'
        );

        await fetchSupabaseShiftLogs();
        renderAllViews();

    } catch (error) {
        console.error('SUPABASE MIGRATION ERROR:', error);
        showToast(`Migration failed: ${error.message}`, 'danger');
    }
}


function setupEventListeners() {
     const migrateBtn = document.getElementById('restoreBackupLogsBtn');

    if (migrateBtn) {
        migrateBtn.addEventListener('click', migrateLocalLogsToSupabase);
    }
    document.querySelectorAll('.nav-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        const targetTab = tab.getAttribute('data-tab');
        switchTab(targetTab);
      });
    });

    // Subnav Hammer Buttons in Tab Insights
    document.querySelectorAll('.subnav-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.subnav-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        selectedSubnavHammer = btn.getAttribute('data-hammer');
        renderInsightsView(getFilteredLogs());
      });
    });

    const manualBtn = document.getElementById('openManualEntryBtn');
    if (manualBtn) manualBtn.addEventListener('click', () => switchTab('entry'));

    const excelBtn = document.getElementById('openExcelModalBtn');
    if (excelBtn) excelBtn.addEventListener('click', () => switchTab('excel'));
    
    const shareBtn = document.getElementById('shareDataLinkBtn');
    if (shareBtn) {
      shareBtn.addEventListener('click', generateShareableDataUrl);
    }

    

    const filterIds = ['globalHammerFilter', 'globalShiftFilter', 'globalRangeFilter'];
    filterIds.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', renderAllViews);
    });

    document.getElementById('resetFiltersBtn').addEventListener('click', () => {
      document.getElementById('globalHammerFilter').value = 'ALL';
      document.getElementById('globalShiftFilter').value = 'ALL';
      document.getElementById('globalRangeFilter').value = '30';
      selectedSubnavHammer = 'ALL';
      document.querySelectorAll('.subnav-btn').forEach(b => b.classList.toggle('active', b.getAttribute('data-hammer') === 'ALL'));
      renderAllViews();
    });

    document.getElementById('logSearchInput').addEventListener('input', () => {
      renderLogsTable(getFilteredLogs());
    });

    const closeQrBtn = document.getElementById('closeMobileQrModalBtn');
    if (closeQrBtn) {
      closeQrBtn.addEventListener('click', () => {
        const modal = document.getElementById('mobileQrModalBackdrop');
        if (modal) modal.style.display = 'none';
      });
    }

    const copyViewOnlyBtn = document.getElementById('copyViewOnlyUrlBtn');
    if (copyViewOnlyBtn) {
      copyViewOnlyBtn.addEventListener('click', () => {
        const input = document.getElementById('viewOnlyShareUrlInput');
        if (input && input.value) {
          copyTextToClipboard(input.value, input);
          showToast('View-Only Share Link copied! Recipients cannot edit data.', 'success');
        }
      });
    }

    const copyMobileBtn = document.getElementById('copyMobileUrlBtn');
    if (copyMobileBtn) {
      copyMobileBtn.addEventListener('click', () => {
        const input = document.getElementById('mobileShareUrlInput');
        if (input && input.value) {
          copyTextToClipboard(input.value, input);
          showToast('Full Edit Share Link copied to clipboard!', 'success');
        }
      });
    }

    const openMobileBtn = document.getElementById('openMobileUrlBtn');
    if (openMobileBtn) {
      openMobileBtn.addEventListener('click', () => {
        const input = document.getElementById('mobileShareUrlInput');
        if (input && input.value) {
          window.open(input.value, '_blank');
        }
      });
    }

    // Supabase Auth Modal listeners
    const openAuthBtn = document.getElementById('openAuthModalBtn');
    const closeAuthBtn = document.getElementById('closeAuthModalBtn');
    const authModal = document.getElementById('supabaseAuthModal');
    const authForm = document.getElementById('supabaseAuthForm');
    const tabLogin = document.getElementById('authTabLogin');
    const tabSignup = document.getElementById('authTabSignup');
    const signupFields = document.getElementById('signupFieldsContainer');
    const logoutBtn = document.getElementById('logoutBtn');
    let isSignupMode = false;

    if (openAuthBtn) openAuthBtn.addEventListener('click', () => authModal.style.display = 'flex');
    if (closeAuthBtn) closeAuthBtn.addEventListener('click', () => authModal.style.display = 'none');
    if (logoutBtn) logoutBtn.addEventListener('click', () => {
      if (supabaseClient) supabaseClient.auth.signOut().then(() => handleUserLoggedOut());
    });

    if (tabLogin && tabSignup) {
      tabLogin.addEventListener('click', (e) => {
        e.preventDefault();
        isSignupMode = false;
        tabLogin.style.borderBottom = '2px solid var(--primary)';
        tabSignup.style.borderBottom = 'none';
        signupFields.style.display = 'none';
        document.getElementById('authSubmitBtn').innerHTML = '<i class="fa-solid fa-right-to-bracket"></i> Sign In to Supabase';
      });

      tabSignup.addEventListener('click', (e) => {
        e.preventDefault();
        isSignupMode = true;
        tabSignup.style.borderBottom = '2px solid var(--primary)';
        tabLogin.style.borderBottom = 'none';
        signupFields.style.display = 'block';
        document.getElementById('authSubmitBtn').innerHTML = '<i class="fa-solid fa-user-plus"></i> Create Supabase Account';
      });
    }

    if (authForm) {
      authForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!supabaseClient) {
          showToast('Please configure Supabase connection first!', 'warning');
          document.getElementById('supabaseConfigModal').style.display = 'flex';
          return;
        }

        const email = document.getElementById('authEmailInput').value.trim();
        const password = document.getElementById('authPasswordInput').value.trim();

        if (isSignupMode) {
          const empName = document.getElementById('authEmpNameInput').value.trim() || email.split('@')[0];
          const dept = document.getElementById('authDeptSelect').value;
          const role = document.getElementById('authRoleSelect').value;

          const { error } = await supabaseClient.auth.signUp({
            email,
            password,
            options: { data: { employee_name: empName, department: dept, role: role } }
          });

          if (error) {
            showToast(`Auth Error: ${error.message}`, 'danger');
          } else {
            showToast('Account created successfully!', 'success');
            authModal.style.display = 'none';
          }
        } else {
          const { error } = await supabaseClient.auth.signInWithPassword({ email, password });
          if (error) {
            showToast(`Login Error: ${error.message}`, 'danger');
          } else {
            showToast('Signed in successfully!', 'success');
            authModal.style.display = 'none';
          }
        }
      });
    }

    // Config Modal Listeners
    const openCfgBtn = document.getElementById('openConfigModalBtn');
    const closeCfgBtn = document.getElementById('closeConfigModalBtn');
    const cfgModal = document.getElementById('supabaseConfigModal');
    const saveCfgBtn = document.getElementById('saveSupabaseConfigBtn');
    const testCfgBtn = document.getElementById('testSupabaseConnectionBtn');

    if (openCfgBtn) openCfgBtn.addEventListener('click', () => cfgModal.style.display = 'flex');
    if (closeCfgBtn) closeCfgBtn.addEventListener('click', () => cfgModal.style.display = 'none');

    if (testCfgBtn) {
      testCfgBtn.addEventListener('click', async () => {
        const url = document.getElementById('cfgSupabaseUrl').value.trim();
        const key = document.getElementById('cfgSupabaseKey').value.trim();
        await testSupabaseConnection(url, key);
      });
    }

    if (saveCfgBtn) {
      saveCfgBtn.addEventListener('click', async () => {
        const url = document.getElementById('cfgSupabaseUrl').value.trim();
        const key = document.getElementById('cfgSupabaseKey').value.trim();
        if (url && key) {
          const success = await testSupabaseConnection(url, key);
          if (success) {
            localStorage.setItem('supabase_url', url);
            localStorage.setItem('supabase_key', key);
            initSupabaseClient();
            setTimeout(() => { cfgModal.style.display = 'none'; }, 1200);
            showToast('Connected to Cloud Database!', 'success');
          }
        } else {
          showToast('Please enter both Supabase URL and Publishable Key.', 'warning');
        }
      });
    }

    // Quality Activity Monitor Filter listeners
    ['qmMonthFilter', 'qmStageFilter', 'qmPartFilter', 'qmReasonFilter'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', renderQualityActivityMonitor);
    });

    const qmResetBtn = document.getElementById('qmResetFiltersBtn');
    if (qmResetBtn) {
      qmResetBtn.addEventListener('click', () => {
        document.getElementById('qmMonthFilter').value = 'ALL';
        document.getElementById('qmStageFilter').value = 'ALL';
        document.getElementById('qmPartFilter').value = 'ALL';
        document.getElementById('qmReasonFilter').value = 'ALL';
        renderQualityActivityMonitor();
      });
    }

    // Manual Entry Form Submit Handler for Supabase
    const entryForm = document.getElementById('manualEntryForm');
    if (entryForm) {
      entryForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const date = document.getElementById('entryDate').value;
        const shift = document.getElementById('entryShift').value;
        const hammer = document.getElementById('entryHammer').value;
        const partNumber = document.getElementById('entryPartNumber').value.trim();
        const plannedTimeMins = parseNum(document.getElementById('entryPlannedTime').value, 660);
        const maintanceMins = parseNum(document.getElementById('entryMaintance').value);
        const dieRelatedMins = parseNum(document.getElementById('entryDieRelated').value);
        const setupMins = parseNum(document.getElementById('entrySetup').value);
        const noManpowerMins = parseNum(document.getElementById('entryNoManpower').value);
        const heatingTimeMins = parseNum(document.getElementById('entryHeatingTime').value);
        const minorStopMins = parseNum(document.getElementById('entryMinorStop').value);
        const totalParts = parseNum(document.getElementById('entryTotalParts').value);
        const rejects = parseNum(document.getElementById('entryRejects').value);
        const goodParts = Math.max(0, totalParts - rejects);
        const stage = document.getElementById('entryInspectionStage').value;
        const reworkQty = parseNum(document.getElementById('entryReworkQty').value);
        const reworkReason = document.getElementById('entryReworkReason').value.trim();
        const rejectionReason = document.getElementById('entryRejectionReason').value.trim();

        if (!date || !partNumber) {
          showToast('Please fill in Date and Part Number.', 'warning');
          return;
        }

        const userId = currentAuthUser ? currentAuthUser.id : null;

        if (supabaseClient) {
          try {
            const { error: prodErr } = await supabaseClient.from('production_data').insert([{
              date, shift, hammer, part_number: partNumber, planned_time_mins: plannedTimeMins,
              planned_qty: totalParts, production_qty: totalParts, good_qty: goodParts,
              created_by: userId
            }]);
            if (prodErr) throw prodErr;

            const { error: qualErr } = await supabaseClient.from('quality_data').insert([{
              date, shift, hammer, part_number: partNumber, inspection_stage: stage,
              inspection_qty: totalParts, rework_qty: reworkQty, rejection_qty: rejects,
              rework_reason: reworkReason, rejection_reason: rejectionReason,
              created_by: userId
            }]);
            if (qualErr) throw qualErr;

            const downtimes = [
              { category: 'Maintenance', mins: maintanceMins },
              { category: 'Die Related', mins: dieRelatedMins },
              { category: 'Setup', mins: setupMins },
              { category: 'No Manpower', mins: noManpowerMins },
              { category: 'Heating Time', mins: heatingTimeMins },
              { category: 'Minor Stop', mins: minorStopMins }
            ].filter(d => d.mins > 0);

            if (downtimes.length > 0) {
              const downRows = downtimes.map(d => ({
                date, shift, hammer, part_number: partNumber,
                downtime_category: d.category, downtime_minutes: d.mins,
                created_by: userId
              }));
              await supabaseClient.from('downtime_data').insert(downRows);
            }

            showToast('Saved directly to Supabase Cloud Database!', 'success');
            fetchSupabaseShiftLogs();
          } catch (err) {
            console.error('Supabase Save Error:', err);
            showToast(`Save Error: ${err.message}`, 'danger');
          }
        } else {
          const record = calculateOeeRecord({
            id: 'log_' + Date.now(), date, shift, machine: hammer, partNumber, plannedTimeMins,
            maintanceMins, dieRelatedMins, setupMins, noManpowerMins, heatingTimeMins, minorStopMins,
            totalParts, goodParts, rejects, rework: reworkQty, idealCycleSec: 45
          });
          shiftLogs.push(record);
          saveShiftLogs();
          renderAllViews();
          showToast('Shift Record Saved Locally.', 'success');
        }
      });
    }

    const commitExcelBtn = document.getElementById('commitExcelImportBtn');
    if (commitExcelBtn) {
      commitExcelBtn.addEventListener('click', commitExcelImportToSupabase);
    }

    const cancelExcelBtn = document.getElementById('cancelExcelImportBtn');
    if (cancelExcelBtn) {
      cancelExcelBtn.addEventListener('click', () => {
        document.getElementById('excelPreviewCard').style.display = 'none';
        showToast('Excel import cancelled.', 'info');
      });
    }

    const clearHeaderBtn = document.getElementById('clearDemoDataHeaderBtn');
    if (clearHeaderBtn) {
      clearHeaderBtn.addEventListener('click', wipeAllSystemDataCompletely);
    }

    const clearAllLogsBtn = document.getElementById('clearAllLogsBtn');
    if (clearAllLogsBtn) {
      clearAllLogsBtn.addEventListener('click', wipeAllSystemDataCompletely);
    }

    document.getElementById('themeToggleBtn').addEventListener('click', toggleTheme);
    document.getElementById('exportComparisonCsv').addEventListener('click', exportSummaryExcel);
  }

  function switchTab(tabId) {
    document.querySelectorAll('.nav-tab').forEach(t => {
      t.classList.toggle('active', t.getAttribute('data-tab') === tabId);
    });

    document.querySelectorAll('.tab-content').forEach(c => {
      c.classList.toggle('active', c.id === `tab-${tabId}`);
    });
  }

  function initTheme() {
    document.body.className = `${currentTheme}-theme`;
    updateThemeIcon();
  }

  function toggleTheme() {
    currentTheme = currentTheme === 'light-green' ? 'dark' : 'light-green';
    document.body.className = `${currentTheme}-theme`;
    localStorage.setItem('oee_theme', currentTheme);
    updateThemeIcon();
    renderAllViews();
  }

  function updateThemeIcon() {
    const btn = document.getElementById('themeToggleBtn');
    if (btn) {
      btn.innerHTML = currentTheme === 'light-green' ? '<i class="fa-solid fa-leaf"></i>' : '<i class="fa-solid fa-moon"></i>';
    }
  }

  function showToast(message, type = 'success') {
    const toast = document.getElementById('toastNotification');
    const toastMsg = document.getElementById('toastMessage');
    const icon = toast.querySelector('.toast-icon');

    if (!toast || !toastMsg) return;

    toastMsg.textContent = message;
    
    if (type === 'success') {
      icon.className = 'toast-icon fa-solid fa-circle-check text-success';
    } else if (type === 'warning' || type === 'info') {
      icon.className = 'toast-icon fa-solid fa-circle-info text-warning';
    } else {
      icon.className = 'toast-icon fa-solid fa-triangle-exclamation text-danger';
    }

    toast.classList.add('show');
    setTimeout(() => {
      toast.classList.remove('show');
    }, 3500);
  }

})();

