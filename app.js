/**
 * FORGE OEE MONITORING & ANALYSIS SYSTEM
 * 5 Hammer Fleet Performance & Production Analytics Engine
 * Super-Resilient Auto-Commit Excel Parser & Mathematically Exact OEE Calculation Engine
 */

(function () {
  'use strict';

  // State Management
  let shiftLogs = [];
  let excelImportBuffer = [];
  let currentImportHammer = null; // null for combined, or hammer name string
  let charts = {};
  let currentTheme = localStorage.getItem('oee_theme') || 'light-green';
  let selectedSubnavHammer = 'ALL';

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
    setupEventListeners();
    setupLiveCalculator();
    setupExcelDropZone();
    setupHammerMiniDropzones();
    setupCloudDbSync();
    renderAllViews();
    pullFromCloudDb(true);
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

  function formatExcelDate(rawVal) {
    if (!rawVal) return new Date().toISOString().split('T')[0];

    // Excel Serial Number (e.g. 45500)
    if (typeof rawVal === 'number' && rawVal > 30000 && rawVal < 60000) {
      const utcDays = Math.floor(rawVal - 25569);
      const utcValue = utcDays * 86400;
      const dateInfo = new Date(utcValue * 1000);
      return dateInfo.toISOString().split('T')[0];
    }

    const str = String(rawVal).trim();
    if (!str) return new Date().toISOString().split('T')[0];

    // Match DD/MM/YYYY or DD-MM-YYYY
    const dmyMatch = str.match(/^(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{4})$/);
    if (dmyMatch) {
      const day = dmyMatch[1].padStart(2, '0');
      const month = dmyMatch[2].padStart(2, '0');
      const year = dmyMatch[3];
      return `${year}-${month}-${day}`;
    }

    // Match YYYY-MM-DD
    const ymdMatch = str.match(/^(\d{4})[\/\.-](\d{1,2})[\/\.-](\d{1,2})$/);
    if (ymdMatch) {
      const year = ymdMatch[1];
      const month = ymdMatch[2].padStart(2, '0');
      const day = ymdMatch[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    const parsedDate = new Date(str);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate.toISOString().split('T')[0];
    }

    return new Date().toISOString().split('T')[0];
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
    if (shiftLogs.length === 0) {
      showToast('No shift logs to share yet. Upload Excel files first!', 'warning');
      return;
    }

    try {
      const compressedArr = compressLogsForUrl(shiftLogs);
      const jsonStr = JSON.stringify(compressedArr);
      const encoded = btoa(encodeURIComponent(jsonStr));

      const viewOnlyUrl = `${window.location.origin}${window.location.pathname}?mode=view#data=${encoded}`;
      const fullEditUrl = `${window.location.origin}${window.location.pathname}#data=${encoded}`;
      
      const viewOnlyInput = document.getElementById('viewOnlyShareUrlInput');
      const editInput = document.getElementById('mobileShareUrlInput');
      const modal = document.getElementById('mobileQrModalBackdrop');

      if (viewOnlyInput) viewOnlyInput.value = viewOnlyUrl;
      if (editInput) editInput.value = fullEditUrl;
      if (modal) modal.style.display = 'flex';

      setTimeout(() => {
        if (viewOnlyInput) copyTextToClipboard(viewOnlyUrl, viewOnlyInput);
        showToast(`View-Only Share Link selected & copied! (${shiftLogs.length} logs included)`, 'success');
      }, 100);

    } catch (err) {
      console.error(err);
      showToast('Error generating share URL.', 'danger');
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

  let cloudSyncTimer = null;
  const defaultCloudEndpoint = 'https://raw.githubusercontent.com/MUKUNDA78/OEE-Hammer-Monitoring-System/master/shift_logs_db.json';

  function getCloudEndpoint() {
    const input = document.getElementById('cloudDbEndpointInput');
    return (input && input.value.trim()) ? input.value.trim() : defaultCloudEndpoint;
  }

  function pullFromCloudDb(silent = false) {
    const endpoint = getCloudEndpoint();
    const pill = document.getElementById('cloudDbStatusPill');

    if (pill) {
      pill.innerHTML = '<span class="dot"></span> <i class="fa-solid fa-spinner fa-spin"></i> Syncing...';
    }

    fetch(endpoint, { cache: 'no-cache' })
      .then(res => {
        if (!res.ok) throw new Error(`HTTP error ${res.status}`);
        return res.json();
      })
      .then(data => {
        let fetchedLogs = [];
        if (Array.isArray(data)) {
          fetchedLogs = data;
        } else if (data && Array.isArray(data.shiftLogs)) {
          fetchedLogs = data.shiftLogs;
        }

        if (fetchedLogs.length > 0 && !window.HAS_URL_DATA) {
          shiftLogs = mergeLogArrays(shiftLogs, fetchedLogs);
          saveShiftLogs();
          renderAllViews();
          if (!silent) showToast(`Online DB Synced! ${shiftLogs.length} logs active.`, 'success');
        } else if (!silent) {
          showToast('Online DB connected. Ready for multi-user shift uploads.', 'info');
        }

        if (pill) {
          pill.innerHTML = `<span class="dot"></span> <i class="fa-solid fa-cloud"></i> Online DB (${shiftLogs.length} logs)`;
          pill.style.borderColor = 'var(--primary)';
        }
      })
      .catch(err => {
        console.warn('Cloud DB pull failed:', err);
        if (!silent) showToast('Local storage active. (Click Share Live Link to sync devices)', 'info');
        if (pill) {
          pill.innerHTML = `<span class="dot"></span> <i class="fa-solid fa-cloud-arrow-up"></i> Multi-User Share Active`;
        }
      });
  }

  let firebaseDbRef = null;
  let isRemoteUpdating = false;

  function initRealtimeCloudDatabase() {
    const endpoint = getCloudEndpoint();
    const pill = document.getElementById('cloudDbStatusPill');

    if (typeof firebase !== 'undefined' && endpoint.includes('firebaseio.com')) {
      try {
        if (!firebase.apps.length) {
          firebase.initializeApp({ databaseURL: endpoint });
        }
        firebaseDbRef = firebase.database().ref('shiftLogs');

        firebaseDbRef.on('value', (snapshot) => {
          const val = snapshot.val();
          if (val) {
            let fetchedLogs = Array.isArray(val) ? val : Object.values(val);
            if (fetchedLogs.length > 0) {
              isRemoteUpdating = true;
              shiftLogs = mergeLogArrays(shiftLogs, fetchedLogs);
              localStorage.setItem('oee_shift_logs_v10', JSON.stringify(shiftLogs));
              localStorage.setItem('oee_shift_logs_backup', JSON.stringify(shiftLogs));
              renderAllViews();
              isRemoteUpdating = false;
              if (pill) {
                pill.innerHTML = `<span class="dot"></span> <i class="fa-solid fa-bolt"></i> Realtime Sync (${shiftLogs.length} logs)`;
                pill.style.borderColor = 'var(--primary)';
              }
            }
          } else if (shiftLogs.length > 0) {
            // Remote DB empty, broadcast local logs to populate Cloud DB!
            broadcastDataToCloud();
          }
        });
        console.log('Firebase Realtime Multi-User Live Sync initialized.');
      } catch (err) {
        console.warn('Firebase init warning:', err);
      }
    }
  }

  function broadcastDataToCloud() {
    if (isRemoteUpdating) return;

    if (firebaseDbRef && shiftLogs.length > 0) {
      firebaseDbRef.set(shiftLogs)
        .then(() => {
          showToast('⚡ Live data broadcasted to all connected users!', 'success');
        })
        .catch(err => {
          console.warn('Firebase broadcast failed:', err);
        });
    }
  }

  function setupCloudDbSync() {
    initRealtimeCloudDatabase();

    const syncHeaderBtn = document.getElementById('cloudDbSyncBtn');
    const pullBtn = document.getElementById('pullFromCloudDbBtn');
    const pushBtn = document.getElementById('pushToCloudDbBtn');
    const autoSelect = document.getElementById('autoCloudSyncSelect');

    if (syncHeaderBtn) syncHeaderBtn.addEventListener('click', () => pullFromCloudDb(false));
    if (pullBtn) pullBtn.addEventListener('click', () => pullFromCloudDb(false));
    if (pushBtn) pushBtn.addEventListener('click', () => broadcastDataToCloud());

    if (autoSelect) {
      autoSelect.addEventListener('change', () => {
        if (cloudSyncTimer) clearInterval(cloudSyncTimer);
        const val = autoSelect.value;
        if (val !== 'OFF') {
          const sec = parseInt(val, 10);
          cloudSyncTimer = setInterval(() => pullFromCloudDb(true), sec * 1000);
          showToast(`Auto Cloud Sync enabled every ${sec}s!`, 'info');
        } else {
          showToast('Auto Cloud Sync disabled. Use Manual Sync.', 'info');
        }
      });

      cloudSyncTimer = setInterval(() => pullFromCloudDb(true), 10000);
    }
  }

  function saveShiftLogs() {
    localStorage.setItem('oee_shift_logs_v10', JSON.stringify(shiftLogs));
    broadcastDataToCloud();
  }

  function updateMonthFilterOptions() {
    const select = document.getElementById('globalRangeFilter');
    if (!select) return;

    const currentSelected = select.value || 'ALL';

    const monthSet = new Set();
    shiftLogs.forEach(l => {
      if (l.date && l.date.length >= 7) {
        monthSet.add(l.date.substring(0, 7));
      }
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
    if (sortedMonths.length === 0) {
      html += `
        <option value="2026-07">Jul 2026</option>
        <option value="2026-06">Jun 2026</option>
        <option value="2026-05">May 2026</option>
        <option value="2026-04">Apr 2026</option>
        <option value="2026-03">Mar 2026</option>
        <option value="2026-02">Feb 2026</option>
        <option value="2026-01">Jan 2026</option>
      `;
    } else {
      sortedMonths.forEach(m => {
        html += `<option value="${m}">${formatMonthLabel(m)}</option>`;
      });
    }

    select.innerHTML = html;
    if (Array.from(select.options).some(opt => opt.value === currentSelected)) {
      select.value = currentSelected;
    } else {
      select.value = 'ALL';
    }
  }

  function getFilteredLogs() {
    const hammerFilter = document.getElementById('globalHammerFilter').value;
    const shiftFilter = document.getElementById('globalShiftFilter').value;
    const monthFilter = document.getElementById('globalRangeFilter').value;

    return shiftLogs.filter(log => {
      if (hammerFilter !== 'ALL' && log.machine !== hammerFilter) return false;
      if (shiftFilter !== 'ALL' && log.shift !== shiftFilter) return false;

      if (monthFilter !== 'ALL') {
        if (!log.date || !log.date.startsWith(monthFilter)) return false;
      }

      return true;
    });
  }

  /* ==========================================================================
     UI RENDERERS & KPI CALCULATIONS
     ========================================================================== */
  function renderAllViews() {
    shiftLogs = shiftLogs.map(l => calculateOeeRecord(l));
    updateMonthFilterOptions();
    const logs = getFilteredLogs();

    updateHammerLogCountBadges();
    renderOverviewKpis(logs);
    renderHammerGauges(logs);
    renderTrendChart(logs);
    renderComponentsChart(logs);
    renderMonthlyTrendView(shiftLogs); // Always show all months history in the Month-Wise Trend section
    renderInsightsView(logs);
    renderComparisonView(logs);
    renderDowntimeView(logs);
    renderLogsTable(logs);
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
      document.getElementById('kpiOverallOee').textContent = '--%';
      document.getElementById('kpiAvailability').textContent = '--%';
      document.getElementById('kpiPerformance').textContent = '--%';
      document.getElementById('kpiQuality').textContent = '--%';
      document.getElementById('kpiPlannedHours').textContent = 'Net Planned: 0 hrs';
      document.getElementById('kpiTotalPieces').textContent = 'Good Parts: 0 pcs';
      document.getElementById('kpiScrapRate').textContent = 'Rejects: 0 pcs (0.0%)';
      document.getElementById('kpiOeeStatus').innerHTML = '<i class="fa-solid fa-cloud-arrow-up text-primary"></i> Ready for Excel Upload';
      return;
    }

    let totalPlannedMins = 0;
    let totalOperatingMins = 0;
    let totalIdealMins = 0;
    let totalProduced = 0;
    let totalGood = 0;
    let totalRejects = 0;

    logs.forEach(l => {
      totalPlannedMins += l.plannedTimeMins;
      totalOperatingMins += l.operatingTimeMins;
      totalIdealMins += Math.min(l.operatingTimeMins, (l.totalParts * l.idealCycleSec) / 60);
      totalProduced += l.totalParts;
      totalGood += l.goodParts;
      totalRejects += l.rejects;
    });

    const avgAvail = totalPlannedMins > 0 ? Math.min(100, (totalOperatingMins / totalPlannedMins) * 100) : 0;
    const avgPerf = totalOperatingMins > 0 ? Math.min(100, (totalIdealMins / totalOperatingMins) * 100) : 0;
    const avgQual = totalProduced > 0 ? Math.min(100, (totalGood / totalProduced) * 100) : 100;
    const overallOee = (avgAvail / 100) * (avgPerf / 100) * (avgQual / 100) * 100;

    const scrapPct = totalProduced > 0 ? ((totalRejects / totalProduced) * 100).toFixed(1) : '0.0';

    document.getElementById('kpiOverallOee').textContent = overallOee.toFixed(1) + '%';
    document.getElementById('kpiAvailability').textContent = avgAvail.toFixed(1) + '%';
    document.getElementById('kpiPerformance').textContent = avgPerf.toFixed(1) + '%';
    document.getElementById('kpiQuality').textContent = avgQual.toFixed(1) + '%';

    document.getElementById('kpiPlannedHours').textContent = `Net Planned: ${(totalPlannedMins / 60).toFixed(1)} hrs`;
    document.getElementById('kpiTotalPieces').textContent = `Good Parts: ${totalGood.toLocaleString()} pcs`;
    document.getElementById('kpiScrapRate').textContent = `Rejects: ${totalRejects.toLocaleString()} pcs (${scrapPct}%)`;

    const oeeStatusEl = document.getElementById('kpiOeeStatus');
    if (overallOee >= 85) {
      oeeStatusEl.innerHTML = '<i class="fa-solid fa-circle-check text-success"></i> World-Class (≥85%)';
    } else if (overallOee >= 75) {
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
      
      let oee = 0, avail = 0, perf = 0, qual = 0;

      if (hLogs.length > 0) {
        let pMins = 0, oMins = 0, iMins = 0, totalParts = 0, goodParts = 0;
        hLogs.forEach(l => {
          pMins += l.plannedTimeMins;
          oMins += l.operatingTimeMins;
          iMins += (l.totalParts * l.idealCycleSec) / 60;
          totalParts += l.totalParts;
          goodParts += l.goodParts;
        });

        avail = pMins > 0 ? (oMins / pMins) * 100 : 0;
        perf = oMins > 0 ? Math.min(100, (iMins / oMins) * 100) : 0;
        qual = totalParts > 0 ? (goodParts / totalParts) * 100 : 100;
        oee = (avail / 100) * (perf / 100) * (qual / 100) * 100;
      }

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
      let pMins = 0, oMins = 0, iMins = 0, totalPcs = 0, goodPcs = 0;
      logArr.forEach(l => {
        pMins += l.plannedTimeMins;
        oMins += l.operatingTimeMins;
        iMins += (l.totalParts * l.idealCycleSec) / 60;
        totalPcs += l.totalParts;
        goodPcs += l.goodParts;
      });

      const avail = pMins > 0 ? (oMins / pMins) * 100 : 0;
      const perf = oMins > 0 ? Math.min(100, (iMins / oMins) * 100) : 0;
      const qual = totalPcs > 0 ? (goodPcs / totalPcs) * 100 : 100;
      const oee = (avail / 100) * (perf / 100) * (qual / 100) * 100;

      return {
        avail: parseFloat(avail.toFixed(1)),
        perf: parseFloat(perf.toFixed(1)),
        qual: parseFloat(qual.toFixed(1)),
        oee: parseFloat(oee.toFixed(1)),
        goodPcs: goodPcs,
        totalShifts: logArr.length
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
        totalShifts: mData.overall.length,
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
      let pMins = 0, oMins = 0, iMins = 0, downtimeMins = 0;
      let totalPcs = 0, goodPcs = 0, rejectsPcs = 0;

      hLogs.forEach(l => {
        pMins += l.plannedTimeMins;
        oMins += l.operatingTimeMins;
        downtimeMins += l.totalDowntimeMins;
        iMins += Math.min(l.operatingTimeMins, (l.totalParts * l.idealCycleSec) / 60);
        totalPcs += l.totalParts;
        goodPcs += l.goodParts;
        rejectsPcs += l.rejects;
      });

      const avail = pMins > 0 ? Math.min(100, (oMins / pMins) * 100) : 0;
      const perf = oMins > 0 ? Math.min(100, (iMins / oMins) * 100) : 0;
      const qual = totalPcs > 0 ? Math.min(100, (goodPcs / totalPcs) * 100) : 100;
      const oee = (avail / 100) * (perf / 100) * (qual / 100) * 100;

      return {
        name: h.name,
        capacity: h.capacity,
        color: h.color,
        shiftsCount: hLogs.length,
        plannedHrs: (pMins / 60).toFixed(1),
        operatingHrs: (oMins / 60).toFixed(1),
        downtimeHrs: (downtimeMins / 60).toFixed(1),
        totalPcs: totalPcs,
        goodPcs: goodPcs,
        rejectsPcs: rejectsPcs,
        avail: parseFloat(avail.toFixed(1)),
        perf: parseFloat(perf.toFixed(1)),
        qual: parseFloat(qual.toFixed(1)),
        oee: parseFloat(oee.toFixed(1))
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

    document.getElementById('reloadSampleDataBtn').addEventListener('click', () => {
      shiftLogs = generateDefaultLogs();
      saveShiftLogs();
      renderAllViews();
      showToast('Reloaded sample historical data.', 'info');
    });

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

  function renderExcelPreview(records, hammerNameOverride) {
    const card = document.getElementById('excelPreviewCard');
    const table = document.getElementById('excelPreviewTable');
    document.getElementById('excelRecordCount').textContent = records.length;
    document.getElementById('excelPreviewTargetLabel').textContent = hammerNameOverride || 'Combined Fleet Sheet';

    card.style.display = 'block';

    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    thead.innerHTML = `
      <tr>
        <th>Date</th>
        <th>Shift</th>
        <th>Machine</th>
        <th>Part Number</th>
        <th>Net Planned</th>
        <th>Downtime</th>
        <th>Total Parts</th>
        <th>Good</th>
        <th>Rejects</th>
        <th>Avail %</th>
        <th>Perf %</th>
        <th>Qual %</th>
        <th>OEE %</th>
      </tr>
    `;

    tbody.innerHTML = '';
    records.slice(0, 10).forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${r.date}</td>
        <td>${r.shift}</td>
        <td>${r.machine}</td>
        <td><strong>${r.partNumber}</strong></td>
        <td>${r.plannedTimeMins}m</td>
        <td>${r.totalDowntimeMins}m</td>
        <td>${r.totalParts}</td>
        <td>${r.goodParts}</td>
        <td>${r.rejects}</td>
        <td>${r.availability}%</td>
        <td>${r.performance}%</td>
        <td>${r.quality}%</td>
        <td><strong style="color: ${getOeeColor(r.oee)}; font-weight: bold; font-size: 14px;">${r.oee}%</strong></td>
      `;
      tbody.appendChild(tr);
    });
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
      tbody.innerHTML = `<tr><td colspan="22" style="text-align: center; color: var(--text-muted); padding: 24px;">No shift records found matching current filters.</td></tr>`;
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

  function recoverAndRestoreData() {
    const recovered = recoverAllPreviousLogs();
    if (recovered.length > 0) {
      shiftLogs = mergeLogArrays(shiftLogs, recovered);
      saveShiftLogs();
      renderAllViews();
      showToast(`Recovered ${shiftLogs.length} shift logs!`, 'success');
      alert(`Data Recovery Completed!\n\nSuccessfully scanned browser storage and recovered ${shiftLogs.length} total shift records. All data has been saved and broadcasted.`);
    } else {
      showToast('No backup logs found in local browser storage.', 'info');
    }
  }

  /* ==========================================================================
     EVENT LISTENERS & TAB NAVIGATION
     ========================================================================== */
  function setupEventListeners() {
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

    const recoverBtn = document.getElementById('restoreBackupLogsBtn');
    if (recoverBtn) {
      recoverBtn.addEventListener('click', recoverAndRestoreData);
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
