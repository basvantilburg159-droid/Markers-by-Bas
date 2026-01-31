(() => {
  const table = document.getElementById('table');
  const statusEl = document.getElementById('status');
  const projectEl = document.getElementById('project');
  const clientEl = document.getElementById('client');
  const pipeIdEl = document.getElementById('pipeId');
  const themeToggle = document.getElementById('themeToggle');
  const syncIdEl = document.getElementById('syncId');
  const syncEnabledEl = document.getElementById('syncEnabled');
  const cloudFileNameEl = document.getElementById('cloudFileName');
  const cloudFilesListEl = document.getElementById('cloudFilesList');
  const cloudSaveBtn = document.getElementById('cloudSaveBtn');
  const cloudLoadBtn = document.getElementById('cloudLoadBtn');
  const cloudDeleteBtn = document.getElementById('cloudDeleteBtn');
  const cloudRefreshBtn = document.getElementById('cloudRefreshBtn');
  const cloudStatusEl = document.getElementById('cloudStatus');
  const userModeToggle = document.getElementById('userModeToggle');
  const speedEl = document.getElementById('speed');
  const flowEl = document.getElementById('flow');
  const flowOffsetEl = document.getElementById('flowOffset');
  const volEl = document.getElementById('vol');
  const pipeSizeEl = document.getElementById('pipeSize');
  const pipeWtEl = document.getElementById('pipeWt');
  const speedRow = document.getElementById('speedRow');
  const flowRow = document.getElementById('flowRow');
  const flowOffsetRow = document.getElementById('flowOffsetRow');
  const volRow = document.getElementById('volRow');
  const pipeSizeRow = document.getElementById('pipeSizeRow');
  const pipeWtRow = document.getElementById('pipeWtRow');
  const modeSelectEl = document.getElementById('modeSelect');
  const launchDateEl = document.getElementById('launchDate');
  const launchEl = document.getElementById('launch');
  const fileInput = document.getElementById('fileInput');
  const importXlsxBtn = document.getElementById('importXlsxBtn');
  const importXlsxInput = document.getElementById('importXlsxInput');
  const markerImportModal = document.getElementById('markerImportModal');
  const markerImportClose = document.getElementById('markerImportClose');
  const markerImportChoose = document.getElementById('markerImportChoose');
  const markerDownloadSample = document.getElementById('markerDownloadSample');
  const markerPaste = document.getElementById('markerPaste');
  const markerApplyPaste = document.getElementById('markerApplyPaste');
  const graphsBtn = document.getElementById('graphsBtn');
  const exportBtn = document.getElementById('exportPdf');
  const exportExcelBtn = document.getElementById('exportXlsx');
  const exportKmzBtn = document.getElementById('exportKmz');
  const newProjectBtn = document.getElementById('newProject');
  const addMarkerBtn = document.getElementById('addMarker');
  const loadJsonBtn = document.getElementById('loadJson');
  const saveJsonBtn = document.getElementById('saveJson');
  const exportExcelModal = document.getElementById('exportExcelModal');
  const exportExcelClose = document.getElementById('exportExcelClose');
  const exportExcelConfirm = document.getElementById('exportExcelConfirm');
  const exportExcelPreviewTable = document.getElementById('exportExcelPreviewTable');
  const exportExcelModalContent = exportExcelModal ? exportExcelModal.querySelector('.modal-content') : null;
  const exportExcelName = document.getElementById('exportExcelName');
  const exportExcelIncludeCoords = document.getElementById('exportExcelIncludeCoords');
  const exportExcelIncludeDates = document.getElementById('exportExcelIncludeDates');
  const exportExcelIncludeExpTimes = document.getElementById('exportExcelIncludeExpTimes');
  const exportExcelIncludeExpVol = document.getElementById('exportExcelIncludeExpVol');
  const exportExcelIncludeActualTimes = document.getElementById('exportExcelIncludeActualTimes');
  const exportExcelIncludeMissed = document.getElementById('exportExcelIncludeMissed');
  const exportExcelEpsg4326 = document.getElementById('exportExcelEpsg4326');
  const exportExcelEpsg3857 = document.getElementById('exportExcelEpsg3857');
  const exportExcelEpsg28992 = document.getElementById('exportExcelEpsg28992');
  const exportExcelEpsgUtm = document.getElementById('exportExcelEpsgUtm');
  const exportKmzModal = document.getElementById('exportKmzModal');
  const exportKmzClose = document.getElementById('exportKmzClose');
  const exportKmzConfirm = document.getElementById('exportKmzConfirm');
  const exportKmzTable = document.getElementById('exportKmzTable');
  const exportKmzIncludeLine = document.getElementById('exportKmzIncludeLine');
  const exportKmzSymbolAll = document.getElementById('exportKmzSymbolAll');
  const exportKmzApplySymbol = document.getElementById('exportKmzApplySymbol');
  const exportKmzToggleDesc = document.getElementById('exportKmzToggleDesc');
  const exportPreviewModal = document.getElementById('exportPreviewModal');
  const exportPreviewClose = document.getElementById('exportPreviewClose');
  const exportPreviewConfirm = document.getElementById('exportPreviewConfirm');
  const exportPreviewTable = document.getElementById('exportPreviewTable');
  const exportPdfName = document.getElementById('exportPdfName');
  const exportIncludeCoords = document.getElementById('exportIncludeCoords');
  const exportIncludeDates = document.getElementById('exportIncludeDates');
  const exportIncludeExpTimes = document.getElementById('exportIncludeExpTimes');
  const exportIncludeActualTimes = document.getElementById('exportIncludeActualTimes');
  const exportEpsgSelect = document.getElementById('exportEpsg');
  const exportOrientationSelect = document.getElementById('exportOrientation');
  const calcSpeedDisplay = document.getElementById('calcSpeedDisplay');
  const gpsBtn = document.getElementById('gpsBtn');
  const gpsModal = document.getElementById('gpsModal');
  const gpsClose = document.getElementById('gpsClose');
  const gpsLoadFile = document.getElementById('gpsLoadFile');
  const gpsFileInput = document.getElementById('gpsFileInput');
  const gpsCsvBtn = document.getElementById('gpsCsvBtn');
  const gpsCsvInput = document.getElementById('gpsCsvInput');
  const gpsCsvModal = document.getElementById('gpsCsvModal');
  const gpsCsvClose = document.getElementById('gpsCsvClose');
  const gpsCsvApply = document.getElementById('gpsCsvApply');
  const gpsCsvName = document.getElementById('gpsCsvName');
  const gpsCsvEpsg = document.getElementById('gpsCsvEpsg');
  const gpsCsvUtmZone = document.getElementById('gpsCsvUtmZone');
  const gpsCsvLat = document.getElementById('gpsCsvLat');
  const gpsCsvLon = document.getElementById('gpsCsvLon');
  const gpsCsvAlt = document.getElementById('gpsCsvAlt');
  const gpsCsvPreview = document.getElementById('gpsCsvPreview');
  const gpsUtmZonePick = document.getElementById('gpsUtmZonePick');
  const gpsCsvUtmZonePick = document.getElementById('gpsCsvUtmZonePick');
  const toggleNavBtn = document.getElementById('toggleNavBtn');
  const utmPickerModal = document.getElementById('utmPickerModal');
  const utmPickerClose = document.getElementById('utmPickerClose');
  const utmPickerApply = document.getElementById('utmPickerApply');
  const utmCountrySelect = document.getElementById('utmCountrySelect');
  const utmCountryReadout = document.getElementById('utmCountryReadout');
  const utmMap = document.getElementById('utmMap');
  const utmMapReadout = document.getElementById('utmMapReadout');
  const gpsCoordsInput = document.getElementById('gpsCoordsInput');
  const gpsParseText = document.getElementById('gpsParseText');
  const gpsStatus = document.getElementById('gpsStatus');
  const gpsPreview = document.getElementById('gpsPreview');
  const gpsApply = document.getElementById('gpsApply');
  const gpsSampleXlsx = document.getElementById('gpsSampleXlsx');
  const gpsSelectTable = document.getElementById('gpsSelectTable');
  const gpsSelectAll = document.getElementById('gpsSelectAll');
  const gpsSelectNone = document.getElementById('gpsSelectNone');
  const gpsMoveUp = document.getElementById('gpsMoveUp');
  const gpsMoveDown = document.getElementById('gpsMoveDown');
  const gpsEpsgSelect = document.getElementById('gpsEpsg');
  const gpsUtmZoneInput = document.getElementById('gpsUtmZone');
  const displayCoordsBtn = document.getElementById('displayCoordsBtn');
  const displayEpsgSelect = document.getElementById('displayEpsg');

  const defaultState = () => ({
    project: 'New Run',
    client: '',
    pipeId: '',
    locked: false,
    userMode: 'expert',
    mode: 'manual',
    speed: 1000,
    flow: '',
    offset: 0,
    vol: '',
    pipeSize: '',
    pipeWt: '',
    pipeWtAuto: true,
    launchDate: '',
    launch: '08:30:00',
    showCoords: false,
    displayEpsg: 'EPSG:4326',
    showNavigate: true,
    exportName: '',
    theme: 'dark',
    syncId: '',
    syncEnabled: false,
    pts: [
      { n: 'Launcher', d: 0, t: '', missed: false, lat: null, lon: null, alt: null },
      { n: 'M1', d: 0, t: '', missed: false, lat: null, lon: null, alt: null },
      { n: 'M2', d: 0, t: '', missed: false, lat: null, lon: null, alt: null },
      { n: 'Receiver', d: 0, t: '', missed: false, lat: null, lon: null, alt: null }
    ]
  });

  let state = defaultState();

  const saveLocal = () => {
    localStorage.setItem('pigging-state', JSON.stringify(state));
    scheduleSyncSave();
  };

  const normalizeState = () => {
    if (!state) return;
    state.showCoords = !!state.showCoords;
    state.launchDate = state.launchDate || '';
    state.displayEpsg = state.displayEpsg || 'EPSG:4326';
    state.client = state.client || '';
    state.pipeId = state.pipeId || '';
    state.userMode = state.userMode === 'operator' ? 'operator' : 'expert';
    state.showNavigate = state.showNavigate !== false;
    state.exportName = state.exportName || '';
    state.theme = state.theme === 'light' ? 'light' : 'dark';
    state.syncId = state.syncId || '';
    state.syncEnabled = !!state.syncEnabled;
    state.pipeSize = state.pipeSize ?? '';
    state.pipeWt = state.pipeWt ?? '';
    state.pipeWtAuto = state.pipeWtAuto !== false;
    state.pts = (state.pts || []).map((p, idx) => ({
      n: p.n || `M${idx}`,
      d: Number(p.d) || 0,
      t: p.t || '',
      missed: !!p.missed,
      lat: p.lat ?? null,
      lon: p.lon ?? null,
      alt: p.alt ?? null
    }));
  };

  const toSec = (t) => {
    if (!t) return null;
    const p = t.split(':').map(Number);
    return p[0] * 3600 + (p[1] || 0) * 60 + (p[2] || 0);
  };

  const toTime = (s) => {
    if (s == null) return '';
    const h = Math.floor(s / 3600) % 24;
    const m = Math.floor((s % 3600) / 60);
    const sec = Math.floor(s % 60);
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(sec).padStart(2, '0')}`;
  };

  const formatDateFromSeconds = (seconds, baseDateStr) => {
    if (seconds == null || !baseDateStr) return '';
    const base = new Date(`${baseDateStr}T00:00:00`);
    if (Number.isNaN(base.getTime())) return '';
    const day = Math.floor(seconds / 86400);
    base.setDate(base.getDate() + day);
    const yyyy = base.getFullYear();
    const mm = String(base.getMonth() + 1).padStart(2, '0');
    const dd = String(base.getDate()).padStart(2, '0');
    return `${dd}/${mm}/${yyyy}`;
  };

  const clientId = (() => {
    const key = 'pigging-client-id';
    const existing = localStorage.getItem(key);
    if (existing) return existing;
    const id = (crypto && crypto.randomUUID) ? crypto.randomUUID() : String(Date.now() + Math.random());
    localStorage.setItem(key, id);
    return id;
  })();

  let firestoreDb = null;
  let syncUnsub = null;
  let isApplyingRemote = false;
  let lastRemoteTs = 0;
  let syncTimer = null;

  const cloudNameKey = 'pigging-cloud-file-name';

  const setCloudStatus = (msg) => {
    if (cloudStatusEl) cloudStatusEl.textContent = msg || '';
  };

  const sanitizeDocId = (raw) => {
    if (!raw) return '';
    return String(raw).trim().replace(/[^a-z0-9-_ ]/gi, '').replace(/\s+/g, '_').slice(0, 64);
  };

  const getSyncDocId = () => sanitizeDocId(state.syncId || state.project || '');

  const initFirebase = () => {
    if (!window.firebase || !window.FIREBASE_CONFIG || !window.FIREBASE_CONFIG.projectId) return false;
    if (!firestoreDb) {
      firebase.initializeApp(window.FIREBASE_CONFIG);
      firestoreDb = firebase.firestore();
    }
    return true;
  };

  const ensureFirebase = () => {
    const ready = initFirebase();
    if (!ready) setCloudStatus('Firebase config missing. Add firebase-config.js values.');
    return ready;
  };

  const getCloudDocId = (name) => sanitizeDocId(name || '');

  const loadCloudList = () => {
    if (!cloudFilesListEl) return;
    if (!ensureFirebase()) return;
    setCloudStatus('Loading cloud list...');
    firestoreDb.collection('markerfiles').orderBy('updatedAt', 'desc').limit(200).get()
      .then((snap) => {
        const options = [];
        snap.forEach((doc) => {
          const data = doc.data() || {};
          const updatedAt = data.updatedAt ? new Date(data.updatedAt) : null;
          const labelDate = updatedAt ? updatedAt.toLocaleString() : 'unknown';
          const name = data.name || doc.id;
          options.push({ id: doc.id, name, label: `${name} — ${labelDate}` });
        });
        cloudFilesListEl.innerHTML = '';
        const placeholder = document.createElement('option');
        placeholder.value = '';
        placeholder.textContent = options.length ? 'Select a cloud file' : 'No cloud files yet';
        cloudFilesListEl.appendChild(placeholder);
        options.forEach((opt) => {
          const el = document.createElement('option');
          el.value = opt.id;
          el.textContent = opt.label;
          el.dataset.name = opt.name;
          cloudFilesListEl.appendChild(el);
        });
        setCloudStatus(`Loaded ${options.length} cloud file${options.length === 1 ? '' : 's'}.`);
      })
      .catch(() => setCloudStatus('Failed to load cloud list.'));
  };

  const getCloudNameInput = () => (cloudFileNameEl ? cloudFileNameEl.value.trim() : '');

  const saveCloudFile = () => {
    if (!ensureFirebase()) return;
    const name = getCloudNameInput() || state.project || '';
    const docId = getCloudDocId(name);
    if (!docId) {
      setCloudStatus('Enter a cloud file name.');
      return;
    }
    if (cloudFileNameEl) cloudFileNameEl.value = name;
    localStorage.setItem(cloudNameKey, name);
    setCloudStatus('Saving to cloud...');
    firestoreDb.collection('markerfiles').doc(docId).set({
      name,
      state,
      updatedAt: Date.now(),
      clientId
    }, { merge: true })
      .then(() => {
        setCloudStatus('Saved to cloud.');
        loadCloudList();
        if (cloudFilesListEl) cloudFilesListEl.value = docId;
      })
      .catch(() => setCloudStatus('Failed to save cloud file.'));
  };

  const loadCloudFile = () => {
    if (!ensureFirebase()) return;
    const docId = (cloudFilesListEl && cloudFilesListEl.value) || getCloudDocId(getCloudNameInput());
    if (!docId) {
      setCloudStatus('Select or enter a cloud file name.');
      return;
    }
    setCloudStatus('Loading from cloud...');
    firestoreDb.collection('markerfiles').doc(docId).get()
      .then((doc) => {
        if (!doc.exists) {
          setCloudStatus('Cloud file not found.');
          return;
        }
        const data = doc.data() || {};
        if (!data.state) {
          setCloudStatus('Cloud file is empty.');
          return;
        }
        state = data.state;
        normalizeState();
        render();
        saveLocal();
        if (cloudFileNameEl) cloudFileNameEl.value = data.name || docId;
        localStorage.setItem(cloudNameKey, data.name || docId);
        setCloudStatus('Loaded from cloud.');
      })
      .catch(() => setCloudStatus('Failed to load cloud file.'));
  };

  const deleteCloudFile = () => {
    if (!ensureFirebase()) return;
    const docId = cloudFilesListEl ? cloudFilesListEl.value : '';
    if (!docId) {
      setCloudStatus('Select a cloud file to delete.');
      return;
    }
    if (!window.confirm('Delete this cloud file?')) return;
    setCloudStatus('Deleting cloud file...');
    firestoreDb.collection('markerfiles').doc(docId).delete()
      .then(() => {
        setCloudStatus('Cloud file deleted.');
        if (cloudFilesListEl) cloudFilesListEl.value = '';
        loadCloudList();
      })
      .catch(() => setCloudStatus('Failed to delete cloud file.'));
  };

  const stopSync = () => {
    if (syncUnsub) syncUnsub();
    syncUnsub = null;
  };

  const startSync = () => {
    if (!state.syncEnabled) return;
    if (!initFirebase()) {
      if (statusEl) statusEl.textContent = 'Firebase config missing. Add firebase-config.js values.';
      return;
    }
    const docId = getSyncDocId();
    if (!docId) {
      if (statusEl) statusEl.textContent = 'Enter a Sync ID to enable realtime updates.';
      return;
    }
    stopSync();
    const ref = firestoreDb.collection('markerlists').doc(docId);
    syncUnsub = ref.onSnapshot((snap) => {
      if (!snap.exists) return;
      const data = snap.data() || {};
      if (data.clientId === clientId) return;
      const ts = Number(data.updatedAt || 0);
      if (ts && ts <= lastRemoteTs) return;
      if (!data.state) return;
      isApplyingRemote = true;
      state = data.state;
      normalizeState();
      render();
      lastRemoteTs = ts || Date.now();
      isApplyingRemote = false;
    });
  };

  const scheduleSyncSave = () => {
    if (!state.syncEnabled || !firestoreDb || isApplyingRemote) return;
    const docId = getSyncDocId();
    if (!docId) return;
    if (syncTimer) clearTimeout(syncTimer);
    syncTimer = setTimeout(() => {
      const ref = firestoreDb.collection('markerlists').doc(docId);
      ref.set({
        state,
        updatedAt: Date.now(),
        clientId
      }, { merge: true }).catch(() => {});
    }, 600);
  };

  const haversineMeters = (lat1, lon1, lat2, lon2) => {
    const toRad = (v) => (v * Math.PI) / 180;
    const R = 6371000;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
    return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  };

  const parseKmlCoordString = (text, name = '') => {
    const list = [];
    if (!text) return list;
    const trimmed = text.trim();
    // If commas are present, treat tokens as lon,lat[,alt]
    if (trimmed.includes(',')) {
      const tokens = trimmed.split(/\s+/);
      tokens.forEach((tok) => {
        if (!tok) return;
        const parts = tok.split(',');
        if (parts.length < 2) return;
        const lon = parseFloat(parts[0]);
        const lat = parseFloat(parts[1]);
        const alt = parts.length > 2 ? parseFloat(parts[2]) : null;
        if (Number.isFinite(lat) && Number.isFinite(lon)) list.push({ lat, lon, alt: Number.isFinite(alt) ? alt : null, name });
      });
      return list;
    }
    // Fallback: whitespace-separated lon lat [alt] (gx:coord style)
    const nums = trimmed.split(/\s+/).map((n) => parseFloat(n)).filter((n) => Number.isFinite(n));
    const step = nums.length % 3 === 0 ? 3 : 2;
    for (let i = 0; i + 1 < nums.length; i += step) {
      const lon = nums[i];
      const lat = nums[i + 1];
      const alt = step === 3 && (i + 2 < nums.length) ? nums[i + 2] : null;
      if (Number.isFinite(lat) && Number.isFinite(lon)) list.push({ lat, lon, alt: Number.isFinite(alt) ? alt : null, name });
    }
    return list;
  };

  const toUtm = (lat, lon) => {
    if (lat == null || lon == null || !Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    const a = 6378137;
    const f = 1 / 298.257223563;
    const k0 = 0.9996;
    const e2 = f * (2 - f);
    const ePrime2 = e2 / (1 - e2);
    const latRad = (lat * Math.PI) / 180;
    const lonRad = (lon * Math.PI) / 180;
    const zone = Math.floor((lon + 180) / 6) + 1;
    const lon0 = ((zone - 1) * 6 - 180 + 3) * (Math.PI / 180); // central meridian

    const N = a / Math.sqrt(1 - e2 * Math.sin(latRad) ** 2);
    const T = Math.tan(latRad) ** 2;
    const C = ePrime2 * Math.cos(latRad) ** 2;
    const A = Math.cos(latRad) * (lonRad - lon0);

    const e4 = e2 ** 2;
    const e6 = e4 * e2;
    const M =
      a *
      ((1 - e2 / 4 - 3 * e4 / 64 - 5 * e6 / 256) * latRad -
        (3 * e2 / 8 + 3 * e4 / 32 + 45 * e6 / 1024) * Math.sin(2 * latRad) +
        (15 * e4 / 256 + 45 * e6 / 1024) * Math.sin(4 * latRad) -
        (35 * e6 / 3072) * Math.sin(6 * latRad));

    const east =
      k0 *
        N *
        (A + ((1 - T + C) * A ** 3) / 6 + ((5 - 18 * T + T ** 2 + 72 * C - 58 * ePrime2) * A ** 5) / 120) +
      500000;

    let north =
      k0 *
      (M +
        N * Math.tan(latRad) *
          (A ** 2 / 2 + ((5 - T + 9 * C + 4 * C ** 2) * A ** 4) / 24 + ((61 - 58 * T + T ** 2 + 600 * C - 330 * ePrime2) * A ** 6) / 720));

    if (lat < 0) north += 10000000;
    const hemi = lat >= 0 ? 'N' : 'S';
    return { zone, hemi, east, north };
  };

  const utmToLatLon = (east, north, zone, hemi = 'N') => {
    if (![east, north, zone].every((v) => Number.isFinite(v))) return null;
    const a = 6378137;
    const f = 1 / 298.257223563;
    const k0 = 0.9996;
    const e2 = f * (2 - f);
    const ePrime2 = e2 / (1 - e2);

    const x = east - 500000;
    let y = north;
    const northern = hemi.toUpperCase() !== 'S';
    if (!northern) y -= 10000000;

    const lon0 = ((zone - 1) * 6 - 180 + 3) * (Math.PI / 180);

    const M = y / k0;
    const mu = M / (a * (1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256));

    const e1 = (1 - Math.sqrt(1 - e2)) / (1 + Math.sqrt(1 - e2));
    const J1 = (3 * e1) / 2 - (27 * e1 ** 3) / 32;
    const J2 = (21 * e1 ** 2) / 16 - (55 * e1 ** 4) / 32;
    const J3 = (151 * e1 ** 3) / 96;
    const J4 = (1097 * e1 ** 4) / 512;

    const fp = mu + J1 * Math.sin(2 * mu) + J2 * Math.sin(4 * mu) + J3 * Math.sin(6 * mu) + J4 * Math.sin(8 * mu);

    const sinFp = Math.sin(fp);
    const cosFp = Math.cos(fp);
    const tanFp = Math.tan(fp);

    const C1 = ePrime2 * cosFp ** 2;
    const T1 = tanFp ** 2;
    const N1 = a / Math.sqrt(1 - e2 * sinFp ** 2);
    const R1 = (a * (1 - e2)) / (1 - e2 * sinFp ** 2) ** 1.5;
    const D = x / (N1 * k0);

    const Q1 = tanFp / (2 * R1 * N1 * k0 ** 2);
    const Q2 = (5 + 3 * T1 + 10 * C1 - 4 * C1 ** 2 - 9 * ePrime2) * tanFp / (24 * R1 * N1 ** 3 * k0 ** 4);
    const Q3 = (61 + 90 * T1 + 298 * C1 + 45 * T1 ** 2 - 252 * ePrime2 - 3 * C1 ** 2) * tanFp / (720 * R1 * N1 ** 5 * k0 ** 6);

    const lat = fp - Q1 * D ** 2 + Q2 * D ** 4 - Q3 * D ** 6;

    const lon = lon0 +
      (D - (1 + 2 * T1 + C1) * D ** 3 / 6 + (5 - 2 * C1 + 28 * T1 - 3 * C1 ** 2 + 8 * ePrime2 + 24 * T1 ** 2) * D ** 5 / 120) /
      cosFp;

    return { lat: (lat * 180) / Math.PI, lon: (lon * 180) / Math.PI };
  };

  const wgsToRd = (lat, lon) => {
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    const lat0 = 52.15517440;
    const lon0 = 5.38720621;
    const dLat = 0.36 * (lat - lat0);
    const dLon = 0.36 * (lon - lon0);

    const dLat2 = dLat * dLat;
    const dLat3 = dLat2 * dLat;
    const dLat4 = dLat3 * dLat;
    const dLon2 = dLon * dLon;
    const dLon3 = dLon2 * dLon;

    const x = 155000
      + 190094.945 * dLon
      + -11832.228 * dLat
      + -114.221 * dLon2
      + -32.391 * dLat2
      + -0.705 * dLon3
      + -2.340 * dLon * dLat2
      + -0.608 * dLat3
      + -0.008 * dLon2 * dLat
      + 0.148 * dLon * dLat3;

    const y = 463000
      + 309056.544 * dLat
      + 3638.893 * dLon2
      + 73.077 * dLat2
      + -157.984 * dLon2 * dLat
      + 59.788 * dLat3
      + 0.433 * dLon3
      + -6.439 * dLon2 * dLat2
      + -0.032 * dLon * dLat3
      + 0.092 * dLon3 * dLat
      + -0.054 * dLat4;

    return { x, y };
  };

  const rdToWgs = (x, y) => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
    const x0 = 155000;
    const y0 = 463000;
    const lat0 = 52.15517440;
    const lon0 = 5.38720621;

    const dX = (x - x0) / 100000;
    const dY = (y - y0) / 100000;

    const dX2 = dX * dX;
    const dX3 = dX2 * dX;
    const dX4 = dX3 * dX;
    const dX5 = dX4 * dX;
    const dY2 = dY * dY;
    const dY3 = dY2 * dY;
    const dY4 = dY3 * dY;

    const lat = lat0
      + (3235.65389 * dY)
      + (-32.58297 * dX2)
      + (-0.24750 * dY2)
      + (-0.84978 * dX2 * dY)
      + (-0.06550 * dY3)
      + (-0.01709 * dX2 * dY2)
      + (-0.00738 * dX)
      + (0.00530 * dX4)
      + (-0.00039 * dX2 * dY3)
      + (0.00033 * dX4 * dY)
      + (-0.00012 * dX * dY);

    const lon = lon0
      + (5260.52916 * dX)
      + (105.94684 * dX * dY)
      + (2.45656 * dX * dY2)
      + (-0.81885 * dX3)
      + (0.05594 * dX * dY3)
      + (-0.05607 * dX3 * dY)
      + (0.01199 * dY)
      + (-0.00256 * dX3 * dY2)
      + (0.00128 * dX * dY4)
      + (0.00022 * dY2)
      + (-0.00022 * dX2)
      + (0.00026 * dX5);

    return { lat, lon };
  };

  const formatCoordForEpsg = (lat, lon, epsgCode = 'EPSG:4326') => {
    if (lat == null || lon == null) return '';
    const code = (epsgCode || 'EPSG:4326').toUpperCase();
    if (code === 'EPSG:4326') return `${lat.toFixed(6)}, ${lon.toFixed(6)}`;
    // Simple Web Mercator projection (EPSG:3857)
    if (code === 'EPSG:3857') {
      const R = 6378137;
      const x = R * (lon * Math.PI / 180);
      const y = R * Math.log(Math.tan(Math.PI / 4 + (lat * Math.PI / 180) / 2));
      return `${x.toFixed(2)}, ${y.toFixed(2)}`;
    }
    if (code === 'EPSG:28992') {
      const rd = wgsToRd(lat, lon);
      if (!rd) return '';
      return `${rd.x.toFixed(2)}, ${rd.y.toFixed(2)}`;
    }
    if (code === 'UTM') {
      const utm = toUtm(lat, lon);
      if (!utm) return '';
      return `${utm.zone}${utm.hemi} ${utm.east.toFixed(2)}, ${utm.north.toFixed(2)}`;
    }
    // Fallback: return WGS84 if unknown EPSG
    return `${lat.toFixed(6)}, ${lon.toFixed(6)}`;
  };

  const projectForEpsg = (lat, lon, epsgCode = 'EPSG:4326') => {
    if (lat == null || lon == null) return { x: '', y: '', zone: '' };
    const code = (epsgCode || 'EPSG:4326').toUpperCase();
    if (code === 'UTM') {
      const utm = toUtm(lat, lon);
      if (!utm) return { x: '', y: '', zone: '' };
      return { x: utm.east.toFixed(2), y: utm.north.toFixed(2), zone: `${utm.zone}${utm.hemi}` };
    }
    if (code === 'EPSG:3857') {
      const R = 6378137;
      const x = R * (lon * Math.PI / 180);
      const y = R * Math.log(Math.tan(Math.PI / 4 + (lat * Math.PI / 180) / 2));
      return { x: x.toFixed(2), y: y.toFixed(2), zone: '' };
    }
    if (code === 'EPSG:28992') {
      const rd = wgsToRd(lat, lon);
      if (!rd) return { x: '', y: '', zone: '' };
      return { x: rd.x.toFixed(2), y: rd.y.toFixed(2), zone: '' };
    }
    // WGS84: use lon/lat as X/Y
    return { x: lon.toFixed(6), y: lat.toFixed(6), zone: '' };
  };

  const parseUtmZoneString = (value) => {
    if (!value) return null;
    const cleaned = String(value).trim().toUpperCase().replace(/\s+/g, '');
    const m = cleaned.match(/^(\d{1,2})([A-Z])?$/);
    if (!m) return null;
    const zone = Number(m[1]);
    if (!Number.isFinite(zone)) return null;
    let hemi = 'N';
    if (m[2]) {
      const letter = m[2];
      if (/[NS]/.test(letter)) hemi = letter;
      else hemi = letter >= 'N' ? 'N' : 'S';
    }
    return { zone, hemi };
  };

  const zoneFromLonLat = (lon, lat) => {
    if (!Number.isFinite(lon) || !Number.isFinite(lat)) return null;
    let zone = Math.floor((lon + 180) / 6) + 1;
    zone = Math.max(1, Math.min(60, zone));
    const hemi = lat >= 0 ? 'N' : 'S';
    return { zone, hemi };
  };

  const formatZone = (zoneObj) => zoneObj ? `${zoneObj.zone}${zoneObj.hemi}` : '';

  const UTM_COUNTRIES = [
    { name: 'Netherlands', lon: 5.29, lat: 52.13 },
    { name: 'Belgium', lon: 4.35, lat: 50.85 },
    { name: 'Luxembourg', lon: 6.13, lat: 49.61 },
    { name: 'Germany', lon: 10.45, lat: 51.16 },
    { name: 'France', lon: 2.21, lat: 46.23 },
    { name: 'United Kingdom', lon: -1.50, lat: 52.35 },
    { name: 'Ireland', lon: -8.0, lat: 53.3 },
    { name: 'Spain', lon: -3.7, lat: 40.4 },
    { name: 'Portugal', lon: -8.0, lat: 39.4 },
    { name: 'Italy', lon: 12.5, lat: 42.8 },
    { name: 'Switzerland', lon: 8.2, lat: 46.8 },
    { name: 'Austria', lon: 14.3, lat: 47.5 },
    { name: 'Denmark', lon: 10.0, lat: 56.0 },
    { name: 'Norway', lon: 8.5, lat: 61.0 },
    { name: 'Sweden', lon: 15.0, lat: 62.0 },
    { name: 'Finland', lon: 26.0, lat: 64.0 },
    { name: 'Poland', lon: 19.1, lat: 52.1 },
    { name: 'Czechia', lon: 15.5, lat: 49.8 },
    { name: 'United States (contiguous)', lon: -98.5, lat: 39.8 },
    { name: 'Canada', lon: -106.3, lat: 56.1 },
    { name: 'Mexico', lon: -102.5, lat: 23.6 },
    { name: 'Brazil', lon: -51.9, lat: -10.8 },
    { name: 'Argentina', lon: -64.0, lat: -34.6 },
    { name: 'Chile', lon: -71.0, lat: -35.7 },
    { name: 'Australia', lon: 134.5, lat: -25.5 },
    { name: 'New Zealand', lon: 172.0, lat: -41.0 },
    { name: 'South Africa', lon: 24.0, lat: -29.0 },
    { name: 'Egypt', lon: 30.0, lat: 26.8 },
    { name: 'Saudi Arabia', lon: 45.0, lat: 24.0 },
    { name: 'United Arab Emirates', lon: 54.3, lat: 24.4 },
    { name: 'India', lon: 78.9, lat: 22.6 },
    { name: 'China', lon: 104.2, lat: 35.9 },
    { name: 'Japan', lon: 138.0, lat: 36.2 },
    { name: 'Indonesia', lon: 113.9, lat: -0.8 }
  ];

  let utmPickerTarget = null;
  let utmPickerZone = null;

  const openUtmPicker = (targetInput) => {
    utmPickerTarget = targetInput || null;
    utmPickerZone = null;
    if (utmCountrySelect) {
      const opts = UTM_COUNTRIES.map((c, i) => `<option value="${i}">${c.name}</option>`).join('');
      utmCountrySelect.innerHTML = `<option value="">Select country…</option>${opts}`;
    }
    if (utmCountryReadout) utmCountryReadout.textContent = 'Select a country to set zone.';
    if (utmMapReadout) utmMapReadout.textContent = 'Lon: —, Lat: —, Zone: —';
    if (utmPickerModal) utmPickerModal.classList.add('show');
  };

  const closeUtmPicker = () => {
    if (utmPickerModal) utmPickerModal.classList.remove('show');
  };

  const convertImportCoords = (coords, epsgCode = 'EPSG:4326', opts = {}) => {
    const list = [];
    const code = (epsgCode || 'EPSG:4326').toUpperCase();
    const utmHint = opts.utmZoneHint ? parseUtmZoneString(opts.utmZoneHint) : null;
    const zoneFromText = (txt) => {
      const raw = String(txt || '').trim();
      if (!/[A-Za-z]/.test(raw)) return null;
      return parseUtmZoneString(raw);
    };

    const invMercator = (x, y) => {
      const R = 6378137;
      const lon = (x / R) * (180 / Math.PI);
      const lat = (2 * Math.atan(Math.exp(y / R)) - Math.PI / 2) * (180 / Math.PI);
      return { lat, lon };
    };

    const invRd = (x, y) => rdToWgs(x, y);

    (coords || []).forEach((p) => {
      if (!p || p.lat == null || p.lon == null || !Number.isFinite(p.lat) || !Number.isFinite(p.lon)) return;
      const alt = Number.isFinite(p.alt) ? p.alt : null;
      const name = p.name || '';
      let out = null;

      if (code === 'EPSG:4326') {
        out = { lat: p.lat, lon: p.lon, alt, name };
      } else if (code === 'EPSG:3857') {
        const geo = invMercator(p.lon, p.lat);
        out = { lat: geo.lat, lon: geo.lon, alt, name };
      } else if (code === 'EPSG:28992') {
        const geo = invRd(p.lon, p.lat);
        if (!geo) throw new Error('Could not convert RD coordinate to lat/lon.');
        out = { lat: geo.lat, lon: geo.lon, alt, name };
      } else if (code === 'UTM') {
        const zoneInfo = p.zoneHint ? parseUtmZoneString(p.zoneHint) : null;
        const zone = zoneInfo || utmHint || zoneFromText(name);
        if (!zone) throw new Error('Provide UTM zone (e.g., 31N) to convert UTM coordinates.');
        const geo = utmToLatLon(p.lon, p.lat, zone.zone, zone.hemi || 'N');
        if (!geo) throw new Error('Could not convert UTM coordinate to lat/lon.');
        out = { lat: geo.lat, lon: geo.lon, alt, name };
      } else {
        out = { lat: p.lat, lon: p.lon, alt, name };
      }

      list.push(out);
    });

    return list;
  };

  const parseKmlCoordinates = (kmlText) => {
    const coords = [];
    if (!kmlText) return coords;
    const doc = new DOMParser().parseFromString(kmlText, 'application/xml');

    // Iterate placemarks to capture names
    doc.querySelectorAll('Placemark').forEach((pm) => {
      const pmName = (pm.querySelector('name')?.textContent || '').trim();
      pm.querySelectorAll('coordinates').forEach((el) => {
        coords.push(...parseKmlCoordString(el.textContent, pmName));
      });
      pm.querySelectorAll('gx\\:coord, coord').forEach((el) => {
        coords.push(...parseKmlCoordString(el.textContent, pmName));
      });
    });

    // Fallback: any coordinates outside placemarks
    if (!coords.length) {
      doc.querySelectorAll('coordinates').forEach((el) => {
        coords.push(...parseKmlCoordString(el.textContent, ''));
      });
      doc.querySelectorAll('gx\\:coord, coord').forEach((el) => {
        coords.push(...parseKmlCoordString(el.textContent, ''));
      });
    }

    return coords;
  };

  const parseManualCoordinates = (text, opts = {}) => {
    const coords = [];
    if (!text) return coords;

    const forcedCode = opts.epsgCode ? opts.epsgCode.toUpperCase() : null;
    const utmZoneHint = opts.utmZone || opts.utmZoneHint || '';

    const parseCoordVal = (raw) => {
      if (!raw) return null;
      let s = raw.trim().replace(/,/g, '.');
      const hemi = s.match(/[NSEW]/i);
      const sign = hemi ? ((/[SW]/i.test(hemi[0])) ? -1 : 1) : 1;
      s = s.replace(/[NSEW]/gi, '').trim();
      const val = parseFloat(s);
      return Number.isFinite(val) ? sign * val : null;
    };

    const parseUtmParts = (parts) => {
      if (parts.length < 3) return null;
      let zoneStr = parts[0];
      let hemi = null;
      let eastIdx = 1;
      let northIdx = 2;
      // Allow formats: "31U 500000 5640000 [h]" or "31 U 500000 5640000 [h]"
      const zoneMatch = zoneStr.match(/^(\d{1,2})([C-HJ-NP-X])?$/i);
      if (!zoneMatch && parts.length >= 4) {
        const altZone = parts[0].match(/^\d{1,2}$/) && parts[1].match(/^[NS]$/i);
        if (altZone) {
          zoneStr = parts[0];
          hemi = parts[1];
          eastIdx = 2;
          northIdx = 3;
        }
      } else if (zoneMatch && parts.length >= 4 && parts[1].match(/^[NS]$/i)) {
        hemi = parts[1];
        eastIdx = 2;
        northIdx = 3;
      }

      const zoneMatch2 = zoneStr.match(/^(\d{1,2})([C-HJ-NP-X])?$/i);
      if (!zoneMatch2) return null;
      const zone = Number(zoneMatch2[1]);
      hemi = hemi || (zoneMatch2[2] ? zoneMatch2[2] : 'N');
      const east = parseFloat(parts[eastIdx]);
      const north = parseFloat(parts[northIdx]);
      const alt = parts.length > northIdx + 1 ? parseFloat(parts[northIdx + 1]) : null;
      if (!Number.isFinite(east) || !Number.isFinite(north)) return null;
      const geo = utmToLatLon(east, north, zone, hemi);
      if (!geo) return null;
      return { lat: geo.lat, lon: geo.lon, alt: Number.isFinite(alt) ? alt : null };
    };

    const lines = text.split(/\n+/).map((l) => l.trim()).filter(Boolean);
    lines.forEach((line, idx) => {
      const parts = line.split(/\t|;/).length > 1 ? line.split(/\t|;/) : line.split(/\s+/);
      if (!parts.length) return;

      const headerish = idx === 0 && parts[0].toLowerCase().includes('name') && parts.some((p) => /lat/i.test(p)) && parts.some((p) => /lon/i.test(p));
      if (headerish && !forcedCode) return;

      if (forcedCode && forcedCode !== 'EPSG:4326') {
        const zoneToken = forcedCode === 'UTM' ? (parts.find((p) => parseUtmZoneString(p))) : null;
        const zoneHint = zoneToken || utmZoneHint;
        const toNum = (val) => {
          const cleaned = val.replace(/,/g, '.');
          if (!cleaned.match(/^-?\d+(?:\.\d+)?$/)) return null;
          const num = parseFloat(cleaned);
          return Number.isFinite(num) ? num : null;
        };
        const nums = parts.map((p) => toNum(p)).filter((v) => Number.isFinite(v));
        let name = '';
        const firstNum = toNum(parts[0]);
        if (!Number.isFinite(firstNum)) name = parts[0];
        const east = nums[0];
        const north = nums[1];
        const alt = nums.length > 2 ? nums[2] : null;
        if (Number.isFinite(east) && Number.isFinite(north)) {
          coords.push({ lat: north, lon: east, alt: Number.isFinite(alt) ? alt : null, name, zoneHint });
        }
        return;
      }

      // AUTO / EPSG:4326 parsing with hemisphere letters and optional UTM auto-detect
      const maybeUtm = parseUtmParts(parts);
      if (maybeUtm) {
        coords.push(maybeUtm);
        return;
      }

      if (parts.length >= 3) {
        const name = parts[0].trim();
        const lat = parseCoordVal(parts[1]);
        const lon = parseCoordVal(parts[2]);
        const alt = parts.length > 3 ? parseCoordVal(parts[3]) : null;
        if (Number.isFinite(lat) && Number.isFinite(lon)) coords.push({ name, lat, lon, alt: Number.isFinite(alt) ? alt : null });
        return;
      }

      if (parts.length >= 2) {
        const lat = parseCoordVal(parts[0]);
        const lon = parseCoordVal(parts[1]);
        const alt = parts.length > 2 ? parseCoordVal(parts[2]) : null;
        if (Number.isFinite(lat) && Number.isFinite(lon)) coords.push({ lat, lon, alt: Number.isFinite(alt) ? alt : null });
      }
    });
    return coords;
  };

  const parseDecimal = (value) => {
    if (value == null) return null;
    const raw = String(value).trim().replace(',', '.');
    if (!raw || raw === '-') return null;
    const num = Number(raw);
    return Number.isFinite(num) ? num : null;
  };

  const parseCoordText = (raw) => {
    if (raw == null) return null;
    let s = String(raw).trim();
    if (!s) return null;
    const hemi = s.match(/[NSEW]/i);
    const sign = hemi ? ((/[SW]/i.test(hemi[0])) ? -1 : 1) : 1;
    s = s.replace(/[NSEW]/gi, '').trim();
    // Normalize thousands/decimal separators (e.g., 5.728.055.692 -> 5728055.692)
    const hasDot = s.includes('.');
    const hasComma = s.includes(',');
    if (hasDot || hasComma) {
      const lastDot = s.lastIndexOf('.');
      const lastComma = s.lastIndexOf(',');
      const decimalPos = Math.max(lastDot, lastComma);
      const decimalChar = decimalPos >= 0 ? s[decimalPos] : null;
      const intPart = decimalPos >= 0 ? s.slice(0, decimalPos) : s;
      const fracPart = decimalPos >= 0 ? s.slice(decimalPos + 1) : '';
      const cleanedInt = intPart.replace(/[.,\s]/g, '');
      const cleanedFrac = fracPart.replace(/[.,\s]/g, '');
      s = cleanedFrac ? `${cleanedInt}.${cleanedFrac}` : cleanedInt;
    }
    const val = parseFloat(s);
    return Number.isFinite(val) ? sign * val : null;
  };

  const parseCoordWithHint = (raw, range) => {
    const primary = parseCoordText(raw);
    const cleaned = String(raw ?? '').replace(/[NSEW]/gi, '').trim();
    const noSep = cleaned.replace(/[.,\s]/g, '');
    const noSepVal = noSep ? Number(noSep) : null;
    const candidates = [primary, Number.isFinite(noSepVal) ? noSepVal : null].filter((v) => Number.isFinite(v));
    if (!range) return primary;
    const match = candidates.find((v) => v >= range.min && v <= range.max);
    return Number.isFinite(match) ? match : primary;
  };

  const parseUtmCsvValue = (raw) => {
    if (raw == null) return null;
    const s = String(raw).trim();
    if (!s) return null;
    if (/[.,]/.test(s)) {
      const num = parseCoordText(s);
      return Number.isFinite(num) ? num : null;
    }
    const cleaned = s.replace(/[^0-9-]/g, '');
    if (!cleaned || cleaned === '-') return null;
    const num = Number(cleaned);
    return Number.isFinite(num) ? num : null;
  };

  const ASME_NPS_OD_MM = [
    { nps: 0.125, od: 10.29 },
    { nps: 0.25, od: 13.72 },
    { nps: 0.375, od: 17.15 },
    { nps: 0.5, od: 21.34 },
    { nps: 0.75, od: 26.67 },
    { nps: 1, od: 33.40 },
    { nps: 1.25, od: 42.16 },
    { nps: 1.5, od: 48.26 },
    { nps: 2, od: 60.33 },
    { nps: 2.5, od: 73.03 },
    { nps: 3, od: 88.90 },
    { nps: 3.5, od: 101.60 },
    { nps: 4, od: 114.30 },
    { nps: 5, od: 141.30 },
    { nps: 6, od: 168.28 },
    { nps: 8, od: 219.08 },
    { nps: 10, od: 273.05 },
    { nps: 12, od: 323.85 },
    { nps: 14, od: 355.60 },
    { nps: 16, od: 406.40 },
    { nps: 18, od: 457.20 },
    { nps: 20, od: 508.00 },
    { nps: 22, od: 558.80 },
    { nps: 24, od: 609.60 },
    { nps: 26, od: 660.40 },
    { nps: 28, od: 711.20 },
    { nps: 30, od: 762.00 },
    { nps: 32, od: 812.80 },
    { nps: 34, od: 863.60 },
    { nps: 36, od: 914.40 }
  ];

  const ASME_STD_WT_IN = [
    { nps: 0.125, wt: 0.068 },
    { nps: 0.25, wt: 0.088 },
    { nps: 0.375, wt: 0.091 },
    { nps: 0.5, wt: 0.109 },
    { nps: 0.75, wt: 0.113 },
    { nps: 1, wt: 0.133 },
    { nps: 1.25, wt: 0.140 },
    { nps: 1.5, wt: 0.145 },
    { nps: 2, wt: 0.154 },
    { nps: 2.5, wt: 0.203 },
    { nps: 3, wt: 0.216 },
    { nps: 3.5, wt: 0.226 },
    { nps: 4, wt: 0.237 },
    { nps: 5, wt: 0.258 },
    { nps: 6, wt: 0.280 },
    { nps: 8, wt: 0.322 },
    { nps: 10, wt: 0.365 },
    { nps: 12, wt: 0.375 },
    { nps: 14, wt: 0.375 },
    { nps: 16, wt: 0.375 },
    { nps: 18, wt: 0.375 },
    { nps: 20, wt: 0.375 },
    { nps: 22, wt: 0.375 },
    { nps: 24, wt: 0.375 },
    { nps: 26, wt: 0.375 },
    { nps: 28, wt: 0.375 },
    { nps: 30, wt: 0.375 },
    { nps: 32, wt: 0.375 },
    { nps: 34, wt: 0.375 },
    { nps: 36, wt: 0.375 }
  ];

  const getAsmeOdMm = (nps) => {
    if (!Number.isFinite(nps)) return null;
    const tol = 0.001;
    const match = ASME_NPS_OD_MM.find((r) => Math.abs(r.nps - nps) <= tol);
    return match ? match.od : null;
  };

  const getAsmeStdWtMm = (nps) => {
    if (!Number.isFinite(nps)) return null;
    const tol = 0.001;
    const match = ASME_STD_WT_IN.find((r) => Math.abs(r.nps - nps) <= tol);
    return match ? match.wt * 25.4 : null;
  };

  const computePipeVolumeM3PerKm = () => {
    const nps = parseDecimal(state.pipeSize);
    const wt = parseDecimal(state.pipeWt);
    if (!Number.isFinite(nps) || !Number.isFinite(wt)) return null;
    const odMm = getAsmeOdMm(nps);
    if (!Number.isFinite(odMm)) return null;
    const idMm = odMm - 2 * wt;
    if (!Number.isFinite(idMm) || idMm <= 0) return null;
    const idM = idMm / 1000;
    const area = Math.PI * (idM / 2) ** 2;
    return area * 1000; // m³ per km
  };

  const getEffectiveVolume = () => {
    if (state.mode === 'pipe') return computePipeVolumeM3PerKm();
    return parseDecimal(state.vol);
  };

  const isOperatorMode = () => state.userMode === 'operator';
  const isEditLocked = () => state.locked || isOperatorMode();

  const calcSpeed = () => {
    if (state.mode === 'calc' || state.mode === 'pipe') {
      const effectiveVol = getEffectiveVolume();
      if (!state.flow || !Number.isFinite(effectiveVol)) return null;
      const flowAdj = Number(state.flow || 0) + Number(state.offset || 0);
      return (flowAdj / effectiveVol) * 1000; // m/h
    }
    return Number(state.speed) || null;
  };

  const updateExpectations = () => {
    const epsgCode = exportEpsgSelect ? exportEpsgSelect.value : 'EPSG:4326';
    const vRaw = calcSpeed();
    const v = vRaw ? vRaw / 3600 : null;
    const effectiveVol = getEffectiveVolume();
    let total = 0;
    let ref = toSec(state.launch);
    state.pts.forEach((p, i) => {
      const expVol = Number.isFinite(effectiveVol) ? ((total * Number(effectiveVol)) / 1000).toFixed(1) : '';
      const timeEl = table.querySelector(`[data-exp-time="${i}"]`);
      const volEl = table.querySelector(`[data-exp-vol="${i}"]`);
      if (displayCoordsBtn) displayCoordsBtn.textContent = state.showCoords ? 'Hide coordinates' : 'Show coordinates';
      if (displayEpsgSelect) displayEpsgSelect.value = state.displayEpsg || 'EPSG:4326';
      if (timeEl) {
        const hasExp = ref != null && !!v;
        timeEl.textContent = hasExp ? toTime(ref) : '';
        const dateEl = table.querySelector(`[data-exp-date="${i}"]`);
        if (dateEl) dateEl.textContent = hasExp ? formatDateFromSeconds(ref, state.launchDate) : '';
      }
      if (volEl) volEl.textContent = expVol;
      if (p.t && !p.missed) ref = toSec(p.t);
      if (i < state.pts.length - 1 && v) ref += Number(p.d || 0) / v;
      total += Number(p.d || 0);
    });
  };

  const render = () => {
    const operatorMode = isOperatorMode();
    const editLocked = isEditLocked();

    projectEl.value = state.project;
    if (clientEl) clientEl.value = state.client || '';
    if (pipeIdEl) pipeIdEl.value = state.pipeId || '';
    if (syncIdEl) syncIdEl.value = state.syncId || '';
    if (syncEnabledEl) syncEnabledEl.checked = !!state.syncEnabled;
    if (userModeToggle) userModeToggle.checked = !operatorMode;
    if (modeSelectEl) modeSelectEl.value = state.mode || 'manual';
    speedEl.value = state.speed;
    flowEl.value = state.flow;
    flowOffsetEl.value = state.offset;
    if (pipeSizeEl) pipeSizeEl.value = state.pipeSize;
    if (pipeWtEl) pipeWtEl.value = state.pipeWt;
    const pipeVol = state.mode === 'pipe' ? computePipeVolumeM3PerKm() : null;
    const effectiveVol = state.mode === 'pipe' ? pipeVol : getEffectiveVolume();
    volEl.value = state.mode === 'pipe'
      ? (Number.isFinite(pipeVol) ? pipeVol.toFixed(2) : '')
      : (state.vol ?? '');
    if (launchDateEl) launchDateEl.value = state.launchDate || '';
    launchEl.value = state.launch;

    const vDerived = (state.mode === 'calc' || state.mode === 'pipe') ? calcSpeed() : null;
    if (calcSpeedDisplay) {
      if ((state.mode === 'calc' || state.mode === 'pipe') && vDerived) {
        calcSpeedDisplay.textContent = `Calculated speed: ${vDerived.toFixed(1)} m/h`;
      } else if (state.mode === 'pipe') {
        calcSpeedDisplay.textContent = 'Enter flow, pipe size, and WT to calculate speed';
      } else if (state.mode === 'calc') {
        calcSpeedDisplay.textContent = 'Enter flow and volume to calculate speed';
      } else {
        calcSpeedDisplay.textContent = '';
      }
    }

    const vRaw = calcSpeed();
    const v = vRaw ? vRaw / 3600 : null; // m/s


    // Manual mode: allow editing speed, keep flow disabled. Calc mode: disable speed, allow flow unless locked.
    if (state.mode === 'calc' || state.mode === 'pipe') {
      if (speedRow) speedRow.style.display = 'none';
      if (flowRow) flowRow.style.display = '';
      if (flowOffsetRow) flowOffsetRow.style.display = '';
      if (volRow) volRow.style.display = '';
      if (pipeSizeRow) pipeSizeRow.style.display = state.mode === 'pipe' ? '' : 'none';
      if (pipeWtRow) pipeWtRow.style.display = state.mode === 'pipe' ? '' : 'none';
      speedEl.disabled = true;
      flowEl.disabled = editLocked;
      flowOffsetEl.disabled = editLocked;
    } else {
      if (speedRow) speedRow.style.display = '';
      if (flowRow) flowRow.style.display = 'none';
      if (flowOffsetRow) flowOffsetRow.style.display = 'none';
      if (volRow) volRow.style.display = '';
      if (pipeSizeRow) pipeSizeRow.style.display = 'none';
      if (pipeWtRow) pipeWtRow.style.display = 'none';
      speedEl.disabled = editLocked;
      flowEl.disabled = true;
      flowOffsetEl.disabled = true;
    }
    volEl.disabled = editLocked || state.mode === 'pipe';
    if (pipeSizeEl) pipeSizeEl.disabled = editLocked || state.mode !== 'pipe';
    if (pipeWtEl) pipeWtEl.disabled = editLocked || state.mode !== 'pipe';
    if (launchDateEl) launchDateEl.disabled = editLocked;
    launchEl.disabled = editLocked;
    if (addMarkerBtn) addMarkerBtn.disabled = editLocked;
    if (newProjectBtn) newProjectBtn.disabled = editLocked;
    if (modeSelectEl) modeSelectEl.disabled = editLocked;
    if (projectEl) projectEl.disabled = editLocked;
    if (clientEl) clientEl.disabled = editLocked;
    if (pipeIdEl) pipeIdEl.disabled = editLocked;
    if (syncIdEl) syncIdEl.disabled = editLocked;
    if (syncEnabledEl) syncEnabledEl.disabled = editLocked;
    if (cloudFileNameEl) cloudFileNameEl.disabled = editLocked;
    if (cloudFilesListEl) cloudFilesListEl.disabled = editLocked;
    if (cloudSaveBtn) cloudSaveBtn.disabled = editLocked;
    if (cloudLoadBtn) cloudLoadBtn.disabled = editLocked;
    if (cloudDeleteBtn) cloudDeleteBtn.disabled = editLocked;
    if (cloudRefreshBtn) cloudRefreshBtn.disabled = editLocked;
    if (themeToggle) themeToggle.checked = state.theme === 'dark';
    document.body.dataset.theme = state.theme;
    if (importXlsxBtn) importXlsxBtn.disabled = operatorMode;
    if (gpsBtn) gpsBtn.disabled = operatorMode;
    if (gpsCsvBtn) gpsCsvBtn.disabled = operatorMode;
    if (exportBtn) exportBtn.disabled = operatorMode;
    if (exportExcelBtn) exportExcelBtn.disabled = operatorMode;
    if (exportKmzBtn) exportKmzBtn.disabled = operatorMode;
    if (graphsBtn) graphsBtn.disabled = operatorMode;
    if (displayEpsgSelect) displayEpsgSelect.disabled = operatorMode;
    if (loadJsonBtn) loadJsonBtn.disabled = false;
    if (saveJsonBtn) saveJsonBtn.disabled = false;
    // Marker inputs honor lock except actual passing times

    if (modeSelectEl) modeSelectEl.value = state.mode || 'manual';

    const displayEpsg = state.displayEpsg || 'EPSG:4326';
    if (displayEpsgSelect) displayEpsgSelect.value = displayEpsg;
    if (displayCoordsBtn) displayCoordsBtn.textContent = state.showCoords ? 'Hide coordinates' : 'Show coordinates';
    if (toggleNavBtn) toggleNavBtn.textContent = state.showNavigate ? 'Google Maps on' : 'Google Maps off';

    const coordMeta = (() => {
      const code = displayEpsg.toUpperCase();
      if (code === 'EPSG:3857') {
        return {
          headers: ['X (m)', 'Y (m)', 'Height (m)'],
          build: (p) => {
            const proj = projectForEpsg(p.lat, p.lon, code);
            return [proj.x, proj.y, p.alt != null ? p.alt : ''];
          }
        };
      }
      if (code === 'EPSG:28992') {
        return {
          headers: ['RD X (m)', 'RD Y (m)', 'Height (m)'],
          build: (p) => {
            const proj = projectForEpsg(p.lat, p.lon, code);
            return [proj.x, proj.y, p.alt != null ? p.alt : ''];
          }
        };
      }
      if (code === 'UTM') {
        return {
          headers: ['East (m)', 'North (m)', 'Zone', 'Height (m)'],
          build: (p) => {
            const proj = projectForEpsg(p.lat, p.lon, code);
            return [proj.x, proj.y, proj.zone, p.alt != null ? p.alt : ''];
          }
        };
      }
      return {
        headers: ['Lat', 'Lon', 'Height (m)'],
        build: (p) => [p.lat != null ? p.lat.toFixed(6) : '', p.lon != null ? p.lon.toFixed(6) : '', p.alt != null ? p.alt : '']
      };
    })();

    const coordHeader = state.showCoords ? coordMeta.headers.map((h) => `<th>${h}</th>`).join('') : '';

    table.innerHTML = `
      <tr>
        <th>Marker name</th>
        <th>Distance between markers (m)</th>
        <th>Total (m)</th>
        ${coordHeader}
        <th>Expected volume (m³)</th>
        <th>Expected time</th>
        <th>Actual passing (hh:mm:ss)</th>
        ${state.showNavigate ? '<th>Google Maps</th>' : ''}
        <th></th>
      </tr>`;

    let total = 0;
    let ref = toSec(state.launch);

    state.pts.forEach((p, i) => {
      const expVol = Number.isFinite(effectiveVol) ? ((total * Number(effectiveVol)) / 1000).toFixed(1) : '';

      const coordCells = state.showCoords ? coordMeta.build(p, displayEpsg) : [];

      const row = table.insertRow();
      const hasCoords = Number.isFinite(p.lat) && Number.isFinite(p.lon);
      const navUrl = hasCoords ? `https://www.google.com/maps/dir/?api=1&destination=${p.lat},${p.lon}` : '';
      row.innerHTML = `
        <td><input value="${p.n}" ${editLocked ? 'disabled' : ''} data-name="${i}"></td>
        <td></td>
        <td>${total}</td>
        ${coordCells.map((c) => `<td>${c ?? ''}</td>`).join('')}
        <td><span data-exp-vol="${i}">${expVol}</span></td>
        <td><span data-exp-time="${i}">${ref && v ? toTime(ref) : ''}</span> <span class="hint" data-exp-date="${i}">${ref && v ? formatDateFromSeconds(ref, state.launchDate) : ''}</span></td>
        <td>
          ${p.missed && i > 0 && i < state.pts.length - 1 ? `<span class="missed">Not detected</span>` : `<input type="time" step="1" value="${p.t}" data-time="${i}">`}
          ${i > 0 && i < state.pts.length - 1 ? `<button class="small secondary" data-missed="${i}">Not detected</button>` : ''}
        </td>
        ${state.showNavigate ? `<td>${hasCoords ? `<a class="small secondary" href="${navUrl}" target="_blank" rel="noopener">Open</a>` : ''}</td>` : ''}
          <td>${i > 0 && i < state.pts.length - 1 ? `<button class="small" data-del="${i}" aria-label="Remove" ${editLocked ? 'disabled' : ''}>🗑</button>` : ''}</td>`;

      if (p.t && !p.missed) ref = toSec(p.t);

      if (i < state.pts.length - 1) {
        const seg = table.insertRow();
        const coordCount = state.showCoords ? coordMeta.headers.length : 0;
        const remainingCols = coordCount + (state.showNavigate ? 6 : 5);
        seg.innerHTML = `
          <td>to ${state.pts[i + 1].n}</td>
          <td><input value="${p.d}" ${editLocked ? 'disabled' : ''} data-dist="${i}"></td>
          <td colspan="${remainingCols}"></td>`;
        if (ref && v) ref += Number(p.d || 0) / v;
        total += Number(p.d || 0);
      }
    });

    saveLocal();
  };

  const showGraphs = () => {
    const win = window.open('', '_blank');
    if (!win) {
      statusEl.textContent = 'Allow pop-ups to view graphs.';
      return;
    }

    try {
      const pts = state.pts;
      const vRaw = calcSpeed();
      const v = vRaw ? vRaw / 3600 : null;

    // Build expected cumulative seconds from launch + cumulative distance/volume
    const expCum = [];
    const distCum = [];
    const expVolCum = [];
    const effectiveVol = getEffectiveVolume();
    let total = 0;
    let ref = toSec(state.launch);
    pts.forEach((p, i) => {
      expCum[i] = ref;
      distCum[i] = total;
      expVolCum[i] = Number.isFinite(effectiveVol) ? ((total * Number(effectiveVol)) / 1000) : null;
      if (i < pts.length - 1 && v) {
        ref += Number(p.d || 0) / v;
      }
      total += Number(p.d || 0);
    });

    // Actual times per marker with day rollovers handled sequentially (use last actual, not just previous index)
    const actualSec = [];
    let dayOffset = 0;
    let lastActual = null;
    pts.forEach((p, i) => {
      if (p.t && !p.missed) {
        let t = toSec(p.t) + dayOffset;
        if (lastActual != null && t <= lastActual) {
          // push to next day until strictly later than last actual
          while (t <= lastActual) {
            t += 86400;
            dayOffset += 86400;
          }
        }
        actualSec[i] = t;
        lastActual = t;
      } else {
        actualSec[i] = null;
      }
    });

    // Segment speeds (between markers) using actual times — include all segments, null where missing
    const segSpeeds = [];
    const segLabels = [];
    for (let i = 1; i < pts.length; i++) {
      const t1 = actualSec[i - 1];
      const t2 = actualSec[i];
      const d = Number(pts[i - 1].d || 0);
      const speed = (t1 != null && t2 != null && t2 > t1) ? (d / (t2 - t1)) * 3600 : null;
      segLabels.push(`${pts[i - 1].n}→${pts[i].n}`);
      segSpeeds.push(speed);
    }

    // Segment durations (actual vs expected)
    const segDurationLabels = [];
    const segDurationActual = [];
    const segDurationExpected = [];
    for (let i = 1; i < pts.length; i++) {
      const d = Number(pts[i - 1].d || 0);
      const expected = v ? d / v : null;
      const t1 = actualSec[i - 1];
      const t2 = actualSec[i];
      const actual = (t1 != null && t2 != null && t2 > t1) ? (t2 - t1) : null;
      segDurationLabels.push(`${pts[i - 1].n}→${pts[i].n}`);
      segDurationActual.push(actual != null ? actual / 60 : null);
      segDurationExpected.push(expected != null ? expected / 60 : null);
    }

    // Rolling average speed (window=3 over segment speeds)
    const rollingLabels = [...segLabels];
    const rollingSpeeds = [];
    const window = 3;
    for (let i = 0; i < segSpeeds.length; i++) {
      const start = Math.max(0, i - window + 1);
      const slice = segSpeeds.slice(start, i + 1).filter((x) => x != null);
      if (slice.length) {
        rollingSpeeds.push(slice.reduce((a, b) => a + b, 0) / slice.length);
      } else {
        rollingSpeeds.push(null);
      }
    }

    // Deviation: actual - expected at markers (minutes)
    const devLabels = [];
    const devValues = [];
    pts.forEach((p, i) => {
      if (actualSec[i] != null && expCum[i] != null) {
        devLabels.push(p.n);
        devValues.push((actualSec[i] - expCum[i]) / 60); // minutes
      }
    });

    // Volume vs distance (expected only)
    const volDistanceLabels = [...distCum];
    const volDistanceValues = expVolCum;

    // Speed vs flow chart removed

    // Prepare window content
    const html = `<!DOCTYPE html>
      <html><head>
        <meta charset="utf-8">
        <title>Graphs - ${state.project}</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
        <style>
          body{font-family:Segoe UI,Arial,sans-serif;background:#0d0d0d;color:#e6e6e6;margin:12px;}
          h2{margin:6px 0;color:#f6b000;}
          button.action{background:#f6b000;color:#111;border:none;border-radius:4px;padding:8px 12px;font-weight:600;cursor:pointer;}
          .actions{display:flex;gap:8px;margin:6px 0 12px 0;}
          .grid{display:grid;gap:16px;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));}
          .card{background:#151515;border:1px solid #2a2a2a;border-radius:8px;padding:12px;}
          canvas{width:100%;height:280px;cursor:pointer;}
          .muted{color:#aaa;font-size:12px;}
          .overlay{position:fixed;inset:0;background:rgba(0,0,0,0.85);display:none;align-items:center;justify-content:center;z-index:9999;}
          .overlay.show{display:flex;}
          .overlay img{max-width:95vw;max-height:95vh;object-fit:contain;border:1px solid #444;border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,0.4);}
        </style>
      </head><body>
        <h2>Graphs - ${state.project}</h2>
        <div class="actions">
          <button id="closeTab" class="action">Back</button>
          <button id="exportGraphsPdf" class="action">Export to PDF</button>
        </div>
        <div id="exportSelect" class="muted" style="margin:0 0 12px 0; font-size:12px;"></div>
        <div id="overlay" class="overlay"><img id="overlayImg" alt="chart"></div>
        <div class="grid">
          <div class="card"><div class="muted">Average speed between markers (all segments)</div><canvas id="chartSegments"></canvas></div>
          <div class="card"><div class="muted">Time deviation (actual - expected, min)</div><canvas id="chartDeviation"></canvas></div>
          <div class="card"><div class="muted">Cumulative distance vs time</div><canvas id="chartDistanceTime"></canvas></div>
          <div class="card"><div class="muted">Segment duration (min) actual vs expected</div><canvas id="chartSegDuration"></canvas></div>
          <div class="card"><div class="muted">Rolling average speed (segments)</div><canvas id="chartRolling"></canvas></div>
          <div class="card"><div class="muted">Expected volume vs distance</div><canvas id="chartVolume"></canvas></div>
        </div>
      <script>
        const segLabels = ${JSON.stringify(segLabels)};
        const segSpeeds = ${JSON.stringify(segSpeeds)};
        const segDurationLabels = ${JSON.stringify(segDurationLabels)};
        const segDurationActual = ${JSON.stringify(segDurationActual)};
        const segDurationExpected = ${JSON.stringify(segDurationExpected)};
        const distCum = ${JSON.stringify(distCum)};
        const expCum = ${JSON.stringify(expCum.map((t) => (t != null ? t / 60 : null)))};
        const rollingLabels = ${JSON.stringify(rollingLabels)};
        const rollingSpeeds = ${JSON.stringify(rollingSpeeds)};
        const devLabels = ${JSON.stringify(devLabels)};
        const devValues = ${JSON.stringify(devValues)};
        const volDistanceLabels = ${JSON.stringify(volDistanceLabels)};
        const volDistanceValues = ${JSON.stringify(volDistanceValues)};
        const actualSec = ${JSON.stringify(actualSec)};
        const projectName = ${JSON.stringify(state.project || '')};

        const palette = ['#f6b000','#5ad8a6','#4e8cff','#f55d7a','#9b7bff'];

        const exportOptions = [
          { id:'chartSegments', label:'Average speed between markers' },
          { id:'chartDeviation', label:'Time deviation' },
          { id:'chartDistanceTime', label:'Cumulative distance vs time' },
          { id:'chartSegDuration', label:'Segment duration' },
          { id:'chartRolling', label:'Rolling average speed' },
          { id:'chartVolume', label:'Expected volume vs distance' }
        ];

        // Render at higher backing resolution for sharper exports/zoom
        Chart.defaults.devicePixelRatio = 3;

        const overlay = document.getElementById('overlay');
        const overlayImg = document.getElementById('overlayImg');
        const exportSelect = document.getElementById('exportSelect');
        const renderExportSelect = () => {
          if (!exportSelect) return;
          exportSelect.innerHTML = 'Select graphs to include: ' + exportOptions.map((o) => {
            return '<label style="margin-right:10px; cursor:pointer;"><input type="checkbox" data-export="' + o.id + '" checked> ' + o.label + '</label>';
          }).join('');
        };
        const showOverlay = (src, w, h) => {
          overlayImg.src = src;
          const maxW = window.innerWidth * 0.95;
          const maxH = window.innerHeight * 0.95;
          let targetW = w ? w * 5 : maxW;
          let targetH = h ? h * 5 : maxH;
          if (w && h) {
            const aspect = w / h;
            if (targetW > maxW) { targetW = maxW; targetH = targetW / aspect; }
            if (targetH > maxH) { targetH = maxH; targetW = targetH * aspect; }
          } else {
            targetW = maxW;
            targetH = maxH;
          }
          overlayImg.style.width = targetW + 'px';
          overlayImg.style.height = targetH + 'px';
          overlay.classList.add('show');
        };
        overlay.onclick = () => {
          overlay.classList.remove('show');
        };

        renderExportSelect();

        document.getElementById('closeTab').onclick = () => window.close();
        document.getElementById('exportGraphsPdf').onclick = async () => {
          try {
            if (!(window.jspdf && window.jspdf.jsPDF)) {
              alert('PDF library not loaded yet.');
              return;
            }
            const doc = new window.jspdf.jsPDF('p','mm','a4');
            const selected = Array.from(document.querySelectorAll('input[data-export]:checked')).map((el) => el.getAttribute('data-export'));
            const ids = selected.length ? selected : exportOptions.map((o) => o.id);
            const existingIds = ids.filter((id) => document.getElementById(id));
            if (!existingIds.length) {
              alert('No charts available to export.');
              return;
            }
            let y = 10;
            const maxW = 190;
            const pageH = 297;
            existingIds.forEach((id, idx) => {
              const c = document.getElementById(id);
              if (!c) return;
              const ratio = c.height && c.width ? c.height / c.width : 0.6;
              const h = maxW * (ratio || 0.6);
              if (y + h > pageH - 10 && idx !== 0) {
                doc.addPage();
                y = 10;
              }
              const labelEl = c.previousElementSibling;
              const label = labelEl ? labelEl.textContent : '';
              doc.text(label, 10, y);
              y += 6;
              doc.addImage(c.toDataURL('image/png'), 'PNG', 10, y, maxW, h);
              y += h + 10;
            });
            const baseName = (projectName || 'graphs').replace(/[^a-z0-9-_ ]/gi,'') || 'graphs';
            const name = baseName + ' - Graphs.pdf';
            if (window.showSaveFilePicker) {
              try {
                const handle = await showSaveFilePicker({
                  suggestedName: name,
                  types: [{ description: 'PDF', accept: { 'application/pdf': ['.pdf'] } }]
                });
                const writable = await handle.createWritable();
                await writable.write(doc.output('blob'));
                await writable.close();
                return;
              } catch (e) {
                if (e && e.name === 'AbortError') return; // user canceled
              }
            }
            doc.save(name);
          } catch (err) {
            console.error('Export failed', err);
            alert('Export failed: ' + (err && err.message ? err.message : 'unknown error'));
          }
        };

        const attachZoom = () => {
          document.querySelectorAll('canvas').forEach((c) => {
            c.onclick = () => {
              try {
                showOverlay(c.toDataURL('image/png'), c.width, c.height);
              } catch (e) {}
            };
          });
        };

        new Chart(document.getElementById('chartSegments'), {
          type:'line',
          data:{
            labels:segLabels,
            datasets:[
              { label:'m/h', data:segSpeeds, borderColor:palette[1], backgroundColor:'rgba(90,216,166,0.15)', tension:0.2, fill:true, spanGaps:true, pointRadius:4, pointHoverRadius:6 }
            ]
          },
          options:{
            plugins:{legend:{display:false}},
            scales:{y:{title:{text:'m/h',display:true}, beginAtZero:true}}
          }
        });

        new Chart(document.getElementById('chartDeviation'), {
          type:'line',
          data:{ labels:devLabels, datasets:[{ label:'minutes', data:devValues, borderColor:palette[2], backgroundColor:'rgba(78,140,255,0.15)', tension:0.2, fill:true }]},
          options:{plugins:{legend:{display:false}}, scales:{y:{title:{text:'Minutes',display:true}}}}
        });

        new Chart(document.getElementById('chartDistanceTime'), {
          type:'line',
          data:{
            labels:expCum,
            datasets:[
              { label:'Expected', data:expCum.map((t, i) => ({ x:t, y:distCum[i] })), borderColor:palette[0], fill:false, tension:0.1 },
              { label:'Actual', data:distCum.map((d,i)=> actualSec[i]!=null ? ({ x:actualSec[i]/60, y:d }) : null).filter(Boolean), borderColor:palette[2], fill:false, tension:0.1 }
            ]
          },
          options:{plugins:{legend:{display:true}}, scales:{x:{title:{text:'Minutes from launch',display:true}}, y:{title:{text:'Distance (m)',display:true}}}}
        });

        new Chart(document.getElementById('chartSegDuration'), {
          type:'bar',
          data:{
            labels:segDurationLabels,
            datasets:[
              { label:'Actual (min)', data:segDurationActual, backgroundColor:palette[2] },
              { label:'Expected (min)', data:segDurationExpected, backgroundColor:palette[0] }
            ]
          },
          options:{responsive:true, plugins:{legend:{display:true}}, scales:{y:{title:{text:'Minutes',display:true}}}}
        });

        new Chart(document.getElementById('chartRolling'), {
          type:'line',
          data:{ labels:rollingLabels, datasets:[{ label:'Rolling speed (m/h)', data:rollingSpeeds, borderColor:palette[3], backgroundColor:'rgba(245,93,122,0.15)', tension:0.2, fill:true }]},
          options:{plugins:{legend:{display:false}}, scales:{y:{title:{text:'m/h',display:true}}}}
        });

        new Chart(document.getElementById('chartVolume'), {
          type:'line',
          data:{ labels:volDistanceLabels, datasets:[{ label:'Expected volume (m³)', data:volDistanceValues, borderColor:palette[0], backgroundColor:'rgba(246,176,0,0.12)', fill:true }]},
          options:{plugins:{legend:{display:false}}, scales:{x:{title:{text:'Distance (m)',display:true}}, y:{title:{text:'Volume (m³)',display:true}}}}
        });

        attachZoom();
      <\/script>
      </body></html>`;
      win.document.write(html);
      win.document.close();
      statusEl.textContent = 'Graphs opened in new tab.';
    } catch (err) {
      statusEl.textContent = 'Could not render graphs.';
      try { win.close(); } catch (e) {}
    }
  };

  let jsPdfLoader = null;
  const ensureJsPdf = () => {
    if (jsPdfLoader) return jsPdfLoader;
    jsPdfLoader = new Promise((resolve, reject) => {
      if (window.jspdf && window.jspdf.jsPDF) return resolve(window.jspdf.jsPDF);
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js';
      s.onload = () => resolve(window.jspdf.jsPDF);
      s.onerror = () => reject(new Error('Failed to load jsPDF'));
      document.head.appendChild(s);
    });
    return jsPdfLoader;
  };

  let jsZipLoader = null;
  const ensureJsZip = () => {
    if (jsZipLoader) return jsZipLoader;
    jsZipLoader = new Promise((resolve, reject) => {
      if (window.JSZip) return resolve(window.JSZip);
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
      s.onload = () => resolve(window.JSZip);
      s.onerror = () => reject(new Error('Failed to load JSZip'));
      document.head.appendChild(s);
    });
    return jsZipLoader;
  };

  const buildExportRows = (opts) => {
    const includeCoords = opts?.includeCoords;
    const includeDates = opts?.includeDates;
    const epsgCode = opts?.epsgCode || 'EPSG:4326';
    const rows = [];

    const vRaw = calcSpeed();
    const v = vRaw ? vRaw / 3600 : null;
    const effectiveVol = getEffectiveVolume();
    let total = 0;
    let ref = toSec(state.launch);

    state.pts.forEach((p, i) => {
      const expTime = ref && v ? toTime(ref) : '';
      const expDate = includeDates ? formatDateFromSeconds(ref, state.launchDate) : '';
      const expVol = Number.isFinite(effectiveVol) ? ((total * Number(effectiveVol)) / 1000).toFixed(1) : '';
      const actual = p.missed ? 'Missed' : (p.t || '');
      const hasCoord = includeCoords && p.lat != null && p.lon != null;
      const coord = hasCoord ? formatCoordForEpsg(p.lat, p.lon, epsgCode) : '';
      const proj = hasCoord ? projectForEpsg(p.lat, p.lon, epsgCode) : { x: '', y: '', zone: '' };
      const latVal = hasCoord ? p.lat.toFixed(6) : '';
      const lonVal = hasCoord ? p.lon.toFixed(6) : '';
      const altVal = includeCoords && p.alt != null ? Number(p.alt).toFixed(2) : '';
      const distToNext = Number(p.d || 0);

      // Keep marker rows focused on totals; show distance only on the segment row that follows
      rows.push({
        name: p.n || '',
        dist: '',
        total,
        coord,
        lat: latVal,
        lon: lonVal,
        alt: altVal,
        x: proj.x || '',
        y: proj.y || '',
        zone: proj.zone || '',
        expTime,
        expDate,
        expVol,
        actual
      });

      // Insert a segment row between markers to highlight the distance to the next marker
      if (i < state.pts.length - 1) {
        const next = state.pts[i + 1];
        rows.push({
          name: `to ${next?.n || ''}`,
          dist: distToNext,
          total: '',
          coord: '',
          lat: '',
          lon: '',
          alt: '',
          x: '',
          y: '',
          zone: '',
          expTime: '',
          expDate: '',
          actual: ''
        });
      }

      if (p.t && !p.missed) ref = toSec(p.t);
      if (i < state.pts.length - 1 && v) ref += distToNext / v;
      total += distToNext;
    });

    return rows;
  };

  const sanitizeExportName = (raw, fallback = 'Markerlist') => {
    return (raw || fallback)
      .replace(/[^a-z0-9-_ ]/gi, '')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/ /g, '_') || fallback;
  };

  const buildExportBaseName = () => {
    const parts = [state.project, state.client, state.pipeId]
      .map((v) => (v || '').trim())
      .filter((v) => v.length > 0);
    const raw = parts.length ? parts.join('_') : 'pigging-run';
    return sanitizeExportName(raw, 'pigging-run');
  };

  const buildMarkerlistName = () => {
    const parts = [state.project, state.client, state.pipeId]
      .map((v) => (v || '').trim())
      .filter((v) => v.length > 0);
    const raw = parts.length ? `Markerlist_${parts.join('_')}` : 'Markerlist';
    return sanitizeExportName(raw, 'Markerlist');
  };

  const getExportName = () => {
    const raw = state.exportName && state.exportName.trim() ? state.exportName.trim() : buildMarkerlistName();
    return sanitizeExportName(raw, 'Markerlist');
  };

  const kmzSymbols = [
    { id: 'ylw-pushpin', label: 'Yellow pushpin', href: 'https://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png' },
    { id: 'red-pushpin', label: 'Red pushpin', href: 'https://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png' },
    { id: 'blu-pushpin', label: 'Blue pushpin', href: 'https://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png' },
    { id: 'grn-pushpin', label: 'Green pushpin', href: 'https://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png' },
    { id: 'wht-blank', label: 'White circle', href: 'https://maps.google.com/mapfiles/kml/paddle/wht-blank.png' },
    { id: 'red-circle', label: 'Red circle', href: 'https://maps.google.com/mapfiles/kml/paddle/red-circle.png' },
    { id: 'blu-circle', label: 'Blue circle', href: 'https://maps.google.com/mapfiles/kml/paddle/blu-circle.png' }
  ];

  let kmzWorking = [];
  let kmzShowDescriptions = true;

  const buildKmzWorking = () => {
    kmzWorking = state.pts.map((p, idx) => {
      const hasCoord = Number.isFinite(p.lat) && Number.isFinite(p.lon);
      return {
        idx,
        name: p.n || `M${idx + 1}`,
        description: p.n ? `Marker: ${p.n}` : '',
        include: hasCoord,
        hasCoord,
        symbol: kmzSymbols[0]?.id || 'ylw-pushpin',
        lat: p.lat,
        lon: p.lon,
        alt: p.alt
      };
    });
  };

  const renderKmzTable = () => {
    if (!exportKmzTable) return;
    if (exportKmzSymbolAll) {
      exportKmzSymbolAll.innerHTML = kmzSymbols.map((s) => `<option value="${s.id}">${s.label}</option>`).join('');
    }
    if (exportKmzToggleDesc) exportKmzToggleDesc.textContent = kmzShowDescriptions ? 'Hide descriptions' : 'Show descriptions';
    exportKmzTable.innerHTML = `
      <tr>
        <th>Include</th>
        <th>Marker</th>
        <th>Symbol</th>
        ${kmzShowDescriptions ? '<th>Description</th>' : ''}
        <th>Status</th>
      </tr>`;

    kmzWorking.forEach((m) => {
      const validLatLon = Number.isFinite(m.lat) && Number.isFinite(m.lon) && Math.abs(m.lat) <= 90 && Math.abs(m.lon) <= 180;
      const row = exportKmzTable.insertRow();
      const symbolOptions = kmzSymbols.map((s) => `<option value="${s.id}" ${s.id === m.symbol ? 'selected' : ''}>${s.label}</option>`).join('');
      const symbol = kmzSymbols.find((s) => s.id === m.symbol);
      const iconCell = symbol ? `<img class="kmz-icon-preview" src="${symbol.href}" alt="">` : '';
      row.innerHTML = `
        <td><input type="checkbox" data-kmz-include="${m.idx}" ${m.include ? 'checked' : ''} ${m.hasCoord ? '' : 'disabled'}></td>
        <td>${m.name}</td>
        <td>
          ${iconCell}
          <select data-kmz-symbol="${m.idx}" ${m.hasCoord ? '' : 'disabled'}>
            ${symbolOptions}
          </select>
        </td>
        ${kmzShowDescriptions ? `<td><input class="kmz-desc-input" type="text" data-kmz-desc="${m.idx}" value="${(m.description || '').replace(/"/g, '&quot;')}" ${m.hasCoord ? '' : 'disabled'}></td>` : ''}
        <td>${validLatLon ? '<span style="color:#3ad56b; font-weight:600;">OK</span>' : '<span style="color:#ff5c5c; font-weight:600;">Error</span>'}</td>`;
    });
  };

  const openExportKmzModal = () => {
    if (!exportKmzModal) return;
    buildKmzWorking();
    renderKmzTable();
    exportKmzModal.classList.add('show');
  };

  const closeExportKmzModal = () => {
    if (!exportKmzModal) return;
    exportKmzModal.classList.remove('show');
  };

  const buildKmzKml = (items, includeLine) => {
    const styles = new Map();
    items.forEach((m) => {
      if (m.include && m.hasCoord && !styles.has(m.symbol)) {
        const sym = kmzSymbols.find((s) => s.id === m.symbol);
        if (sym) styles.set(m.symbol, sym.href);
      }
    });

    const styleXml = Array.from(styles.entries()).map(([id, href]) => {
      return `
        <Style id="${id}">
          <IconStyle>
            <scale>1.0</scale>
            <Icon><href>${href}</href></Icon>
          </IconStyle>
        </Style>`;
    }).join('');

    const lineStyleXml = includeLine ? `
      <Style id="lineStyle">
        <LineStyle><color>ff0000ff</color><width>3</width></LineStyle>
      </Style>` : '';

    const placemarks = items.filter((m) => m.include && m.hasCoord).map((m) => {
      const alt = Number.isFinite(m.alt) ? m.alt : 0;
      const descRaw = String(m.description || '');
      const desc = descRaw.replace(/\]\]>/g, ']]]]><![CDATA[>');
      return `
        <Placemark>
          <name>${String(m.name || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</name>
          ${descRaw ? `<description><![CDATA[${desc}]]></description>` : ''}
          <styleUrl>#${m.symbol}</styleUrl>
          <Point><coordinates>${m.lon},${m.lat},${alt}</coordinates></Point>
        </Placemark>`;
    }).join('');

    let linePlacemark = '';
    if (includeLine) {
      const coords = items
        .filter((m) => m.include && m.hasCoord)
        .map((m) => `${m.lon},${m.lat},${Number.isFinite(m.alt) ? m.alt : 0}`)
        .join(' ');
      if (coords) {
        linePlacemark = `
        <Placemark>
          <name>Marker line</name>
          <styleUrl>#lineStyle</styleUrl>
          <LineString>
            <tessellate>1</tessellate>
            <coordinates>${coords}</coordinates>
          </LineString>
        </Placemark>`;
      }
    }

    return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>${buildExportBaseName()}</name>
    ${styleXml}
    ${lineStyleXml}
    ${placemarks}
    ${linePlacemark}
  </Document>
</kml>`;
  };

  const exportKmz = async () => {
    try {
      const JSZip = await ensureJsZip();
      const includeLine = exportKmzIncludeLine ? exportKmzIncludeLine.checked : true;
      const selected = kmzWorking.filter((m) => m.include && m.hasCoord);
      if (!selected.length) {
        statusEl.textContent = 'No markers with coordinates selected.';
        return;
      }
      const kml = buildKmzKml(kmzWorking, includeLine);
      const zip = new JSZip();
      zip.file('doc.kml', kml);
      const blob = await zip.generateAsync({ type: 'blob' });
      const name = buildExportBaseName() + '.kmz';

      if (window.showSaveFilePicker) {
        try {
          const h = await showSaveFilePicker({
            suggestedName: name,
            types: [{ description: 'KMZ', accept: { 'application/vnd.google-earth.kmz': ['.kmz'] } }]
          });
          const w = await h.createWritable();
          await w.write(blob);
          await w.close();
          statusEl.textContent = 'KMZ saved.';
          return;
        } catch (e) {
          if (e && e.name === 'AbortError') return;
        }
      }

      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.click();
      statusEl.textContent = 'KMZ downloaded.';
    } catch (err) {
      statusEl.textContent = 'Could not export KMZ.';
    }
  };

  const computeColumnWidths = (doc, columns, rows, padding = 3, margin = 20) => {
    const avail = doc.internal.pageSize.getWidth() - margin;
    const widths = columns.map((col) => {
      let max = doc.getTextWidth(col.label || '') + padding;
      rows.forEach((r) => {
        const txt = String(r[col.key] ?? '');
        const w = doc.getTextWidth(txt) + padding;
        if (w > max) max = w;
      });
      return Math.max(max, 16);
    });
    const total = widths.reduce((a, b) => a + b, 0);
    if (total > avail && total > 0) {
      const scale = avail / total;
      for (let i = 0; i < widths.length; i++) widths[i] = Math.max(widths[i] * scale, 16);
    }
    return widths;
  };

  const renderExportPreview = () => {
    if (!exportPreviewTable) return;
    const includeCoords = exportIncludeCoords ? exportIncludeCoords.checked : false;
    const includeDates = exportIncludeDates ? exportIncludeDates.checked : false;
    const includeExpTimes = exportIncludeExpTimes ? exportIncludeExpTimes.checked : true;
    const includeActualTimes = exportIncludeActualTimes ? exportIncludeActualTimes.checked : true;
    const epsgCode = exportEpsgSelect ? exportEpsgSelect.value : 'EPSG:4326';
    const rows = buildExportRows({ includeCoords, includeDates, epsgCode });

    let coordHeader = '';
    if (includeCoords) {
      if (epsgCode === 'EPSG:4326') {
        coordHeader = '<th>Lat</th><th>Lon</th><th>Height (m)</th>';
      } else if (epsgCode === 'EPSG:3857') {
        coordHeader = '<th>X (m)</th><th>Y (m)</th><th>Height (m)</th>';
      } else if (epsgCode === 'EPSG:28992') {
        coordHeader = '<th>RD X (m)</th><th>RD Y (m)</th><th>Height (m)</th>';
      } else if (epsgCode === 'UTM') {
        coordHeader = '<th>East (m)</th><th>North (m)</th><th>Zone</th><th>Height (m)</th>';
      }
    }
    const dateHeader = includeDates ? '<th>Expected date</th>' : '';
    const expTimeHeader = includeExpTimes ? '<th>Expected time</th>' : '';
    const actualHeader = includeActualTimes ? '<th>Actual</th>' : '';

    exportPreviewTable.innerHTML = `
      <tr>
        <th>Marker</th>
        <th>Distance (m)</th>
        <th>Total (m)</th>
        ${coordHeader}
        ${expTimeHeader}
        ${dateHeader}
        ${actualHeader}
      </tr>`;

    rows.forEach((r) => {
      let coordCell = '';
      if (includeCoords) {
        if (epsgCode === 'EPSG:4326') {
          coordCell = `<td>${r.lat || ''}</td><td>${r.lon || ''}</td><td>${r.alt || ''}</td>`;
        } else if (epsgCode === 'EPSG:3857') {
          coordCell = `<td>${r.x || ''}</td><td>${r.y || ''}</td><td>${r.alt || ''}</td>`;
        } else if (epsgCode === 'EPSG:28992') {
          coordCell = `<td>${r.x || ''}</td><td>${r.y || ''}</td><td>${r.alt || ''}</td>`;
        } else if (epsgCode === 'UTM') {
          coordCell = `<td>${r.x || ''}</td><td>${r.y || ''}</td><td>${r.zone || ''}</td><td>${r.alt || ''}</td>`;
        }
      }
      const dateCell = includeDates ? `<td>${r.expDate}</td>` : '';
      const expTimeCell = includeExpTimes ? `<td>${r.expTime}</td>` : '';
      const actualCell = includeActualTimes ? `<td>${r.actual}</td>` : '';
      const row = exportPreviewTable.insertRow();
      row.innerHTML = `
        <td>${r.name}</td>
        <td>${r.dist}</td>
        <td>${r.total}</td>
        ${coordCell}
        ${expTimeCell}
        ${dateCell}
        ${actualCell}`;
    });
  };

  const openExportPreview = () => {
    if (exportPreviewModal) exportPreviewModal.classList.add('show');
    if (!state.exportName) state.exportName = buildMarkerlistName();
    if (exportPdfName) exportPdfName.value = getExportName();
    renderExportPreview();
  };

  const closeExportPreview = () => {
    if (exportPreviewModal) exportPreviewModal.classList.remove('show');
  };

  const openExportExcelModal = () => {
    if (exportExcelModal) exportExcelModal.classList.add('show');
    if (!state.exportName) state.exportName = buildMarkerlistName();
    if (exportExcelName) exportExcelName.value = getExportName();
    renderExcelPreview();
  };

  const closeExportExcelModal = () => {
    if (exportExcelModal) exportExcelModal.classList.remove('show');
  };

  const exportPdf = async (opts) => {
    try {
      const includeCoords = opts?.includeCoords ?? (exportIncludeCoords ? exportIncludeCoords.checked : false);
      const includeDates = opts?.includeDates ?? (exportIncludeDates ? exportIncludeDates.checked : false);
      const includeExpTimes = opts?.includeExpTimes ?? (exportIncludeExpTimes ? exportIncludeExpTimes.checked : true);
      const includeActualTimes = opts?.includeActualTimes ?? (exportIncludeActualTimes ? exportIncludeActualTimes.checked : true);
      const epsgCode = opts?.epsgCode ?? (exportEpsgSelect ? exportEpsgSelect.value : 'EPSG:4326');
      const orientationChoice = opts?.orientation ?? (exportOrientationSelect ? exportOrientationSelect.value : 'auto');
      const jsPDF = await ensureJsPdf();

      const columns = [
        { key: 'name', label: 'Marker' },
        { key: 'dist', label: 'Dist (m)' },
        { key: 'total', label: 'Total (m)' },
      ];
      if (includeCoords) {
        if (epsgCode === 'EPSG:4326') {
          columns.push(
            { key: 'lat', label: 'Lat' },
            { key: 'lon', label: 'Lon' },
            { key: 'alt', label: 'Height (m)' }
          );
        } else if (epsgCode === 'EPSG:3857') {
          columns.push(
            { key: 'x', label: 'X (m)' },
            { key: 'y', label: 'Y (m)' },
            { key: 'alt', label: 'Height (m)' }
          );
        } else if (epsgCode === 'EPSG:28992') {
          columns.push(
            { key: 'x', label: 'RD X (m)' },
            { key: 'y', label: 'RD Y (m)' },
            { key: 'alt', label: 'Height (m)' }
          );
        } else if (epsgCode === 'UTM') {
          columns.push(
            { key: 'x', label: 'East (m)' },
            { key: 'y', label: 'North (m)' },
            { key: 'zone', label: 'Zone' },
            { key: 'alt', label: 'Height (m)' }
          );
        }
      }
      if (includeExpTimes) columns.push({ key: 'expTime', label: 'Expected passing time' });
      if (includeDates) columns.push({ key: 'expDate', label: 'Date' });
      if (includeActualTimes) columns.push({ key: 'actual', label: 'Actual passing time' });

      const rows = buildExportRows({ includeCoords, includeDates, epsgCode });
      const tableX = 10;
      const bodyRowH = 7;
      const minFontSize = 6;
      const maxFontSize = 12;
      const minColWidth = 12;

      const computePdfColumnWidths = (docRef, cols, data, padding = 4) => {
        return cols.map((col) => {
          let max = 0;
          const label = col.label || '';
          const words = String(label).split(' ');
          words.forEach((w) => {
            const wWidth = docRef.getTextWidth(w) + padding;
            if (wWidth > max) max = wWidth;
          });
          data.forEach((r) => {
            const txt = String(r[col.key] ?? '');
            const wWidth = docRef.getTextWidth(txt) + padding;
            if (wWidth > max) max = wWidth;
          });
          return Math.max(max, minColWidth);
        });
      };

      const computeLayout = (orientation, fontSize) => {
        const measureDoc = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
        measureDoc.setFontSize(fontSize);
        const widths = computePdfColumnWidths(measureDoc, columns, rows);
        const total = widths.reduce((a, b) => a + b, 0);
        const pageW = measureDoc.internal.pageSize.getWidth();
        return { orientation, fontSize, colWidths: widths, tableW: total, pageW };
      };

      const pickLayoutForOrientation = (orientation) => {
        for (let fs = maxFontSize; fs >= minFontSize; fs--) {
          const layout = computeLayout(orientation, fs);
          if (layout.tableW <= layout.pageW - 20) return layout;
        }
        return null;
      };

      let chosen = null;
      if (orientationChoice === 'portrait' || orientationChoice === 'landscape') {
        const o = orientationChoice === 'portrait' ? 'p' : 'l';
        chosen = pickLayoutForOrientation(o);
        if (!chosen) {
          const fallback = computeLayout(o, minFontSize);
          const scale = (fallback.pageW - 20) / fallback.tableW;
          fallback.colWidths = fallback.colWidths.map((w) => Math.max(Math.floor(w * scale), minColWidth));
          fallback.tableW = fallback.colWidths.reduce((a, b) => a + b, 0);
          chosen = fallback;
        }
      } else {
        for (let fs = maxFontSize; fs >= minFontSize; fs--) {
          const portrait = computeLayout('p', fs);
          if (portrait.tableW <= portrait.pageW - 20) { chosen = portrait; break; }
          const landscape = computeLayout('l', fs);
          if (landscape.tableW <= landscape.pageW - 20) { chosen = landscape; break; }
        }
        if (!chosen) {
          const fallback = computeLayout('l', minFontSize);
          const scale = (fallback.pageW - 20) / fallback.tableW;
          fallback.colWidths = fallback.colWidths.map((w) => Math.max(Math.floor(w * scale), minColWidth));
          fallback.tableW = fallback.colWidths.reduce((a, b) => a + b, 0);
          chosen = fallback;
        }
      }

      if (chosen.tableW < chosen.pageW - 20) {
        const expand = (chosen.pageW - 20) / chosen.tableW;
        chosen.colWidths = chosen.colWidths.map((w) => Math.max(Math.floor(w * expand), minColWidth));
        chosen.tableW = chosen.colWidths.reduce((a, b) => a + b, 0);
      }

      const doc = new jsPDF({ orientation: chosen.orientation, unit: 'mm', format: 'a4' });
      let y = 14;

      const header = getExportName();
      doc.setFontSize(14);
      doc.text(header, 10, y);
      y += 8;

      doc.setFontSize(10);
      doc.text(`Launch: ${state.launch || '-'}`, 10, y);
      y += 6;
      const derived = calcSpeed();
      const effectiveVol = getEffectiveVolume();
      if (state.mode === 'calc') {
        const calcLine = `Mode: From flow/volume   Flow: ${state.flow || '-'} m³/h   Volume: ${Number.isFinite(effectiveVol) ? effectiveVol.toFixed(2) : '-'} m³/km   Offset: ${state.offset || 0}`;
        const speedLine = derived ? `Calculated speed: ${derived.toFixed(1)} m/h` : '';
        doc.text(calcLine, 10, y);
        y += 6;
        if (speedLine) {
          doc.text(speedLine, 10, y);
          y += 6;
        }
      } else if (state.mode === 'pipe') {
        const calcLine = `Mode: From pipe specs   Flow: ${state.flow || '-'} m³/h   Pipe: ${state.pipeSize || '-'} in   WT: ${state.pipeWt || '-'} mm   Volume: ${Number.isFinite(effectiveVol) ? effectiveVol.toFixed(2) : '-'} m³/km   Offset: ${state.offset || 0}`;
        const speedLine = derived ? `Calculated speed: ${derived.toFixed(1)} m/h` : '';
        doc.text(calcLine, 10, y);
        y += 6;
        if (speedLine) {
          doc.text(speedLine, 10, y);
          y += 6;
        }
      } else {
        const manualLine = `Mode: Manual speed   Speed: ${state.speed || '-'} m/h`;
        doc.text(manualLine, 10, y);
        y += 6;
      }
      y += 4;

      const colWidths = chosen.colWidths;
      const tableW = chosen.tableW;
      doc.setFontSize(chosen.fontSize);
      const colPositions = [];
      {
        let x = tableX;
        columns.forEach((col, idx) => {
          x += colWidths[idx];
          colPositions.push(x);
        });
      }

      const headerLineHeight = doc.getTextDimensions('Mg').h || 4;
      const headerLines = columns.map((col, idx) => doc.splitTextToSize(col.label, Math.max(colWidths[idx] - 4, minColWidth)));
      const headerRowH = Math.max(bodyRowH, Math.ceil(Math.max(...headerLines.map((l) => l.length)) * headerLineHeight + 4));

      const addHeader = () => {
        doc.setFontSize(chosen.fontSize);
        doc.setDrawColor(40, 96, 144);
        doc.setFillColor(36, 114, 169);
        doc.setTextColor(255, 255, 255);
        const top = y;
        doc.rect(tableX, top, tableW, headerRowH, 'FD');
        let x = tableX + 2;
        columns.forEach((col, idx) => {
          const lines = headerLines[idx];
          doc.text(lines, x, top + 2, { baseline: 'top' });
          x += colWidths[idx];
        });
        // Vertical dividers for header
        colPositions.forEach((px) => doc.line(px, top, px, top + headerRowH));
        y = top + headerRowH;
        doc.setTextColor(34, 34, 34);
      };

      const pageH = doc.internal.pageSize.getHeight();
      const maybeBreak = () => {
        if (y + bodyRowH > pageH - 10) {
          doc.addPage();
          y = 14;
          addHeader();
        }
      };

      addHeader();

      doc.setFontSize(chosen.fontSize);
      doc.setDrawColor(205, 205, 205);
      rows.forEach((r, idx) => {
        const stripe = idx % 2 === 1;
        const top = y;
        if (stripe) {
          doc.setFillColor(245, 245, 245);
          doc.rect(tableX, top, tableW, bodyRowH, 'FD');
        } else {
          doc.setFillColor(255, 255, 255);
          doc.rect(tableX, top, tableW, bodyRowH, 'FD');
        }

        let x = tableX + 2;
        columns.forEach((col, cIdx) => {
          const txt = String(r[col.key] ?? '');
          doc.text(txt, x, top + bodyRowH - 2);
          x += colWidths[cIdx];
        });
        // Vertical dividers for body rows
        colPositions.forEach((px) => doc.line(px, top, px, top + bodyRowH));
        y += bodyRowH;
        maybeBreak();
      });

      const blob = doc.output('blob');
      const name = getExportName() + '.pdf';

      if (window.showSaveFilePicker) {
        try {
          const h = await showSaveFilePicker({ suggestedName: name, types: [{ description: 'PDF', accept: { 'application/pdf': ['.pdf'] } }] });
          const w = await h.createWritable();
          await w.write(blob);
          await w.close();
          statusEl.textContent = 'PDF saved.';
          return;
        } catch (e) {
          if (e && e.name === 'AbortError') return; // user canceled
        }
      }

      // Fallback to download
      doc.save(name);
      statusEl.textContent = 'PDF downloaded.';
    } catch (err) {
      statusEl.textContent = 'Could not export PDF.';
    }
  };

  let sheetJsLoader = null;
  const ensureSheetJs = () => {
    if (sheetJsLoader) return sheetJsLoader;
    sheetJsLoader = new Promise((resolve, reject) => {
      if (window.XLSX) return resolve(window.XLSX);
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
      s.onload = () => resolve(window.XLSX);
      s.onerror = () => reject(new Error('Failed to load XLSX'));
      document.head.appendChild(s);
    });
    return sheetJsLoader;
  };

  const getExcelEpsgSelections = () => {
    const selections = [];
    if (exportExcelEpsg4326 && exportExcelEpsg4326.checked) selections.push('EPSG:4326');
    if (exportExcelEpsg3857 && exportExcelEpsg3857.checked) selections.push('EPSG:3857');
    if (exportExcelEpsg28992 && exportExcelEpsg28992.checked) selections.push('EPSG:28992');
    if (exportExcelEpsgUtm && exportExcelEpsgUtm.checked) selections.push('UTM');
    return selections;
  };

  const renderExcelPreview = () => {
    if (!exportExcelPreviewTable) return;
    const includeCoords = exportExcelIncludeCoords ? exportExcelIncludeCoords.checked : true;
    const includeExpTimes = exportExcelIncludeExpTimes ? exportExcelIncludeExpTimes.checked : true;
    const includeDates = exportExcelIncludeDates ? exportExcelIncludeDates.checked : true;
    const includeExpVol = exportExcelIncludeExpVol ? exportExcelIncludeExpVol.checked : true;
    const includeActualTimes = exportExcelIncludeActualTimes ? exportExcelIncludeActualTimes.checked : true;
    const includeMissed = exportExcelIncludeMissed ? exportExcelIncludeMissed.checked : true;
    let epsgCodes = getExcelEpsgSelections();
    if (includeCoords && epsgCodes.length === 0) epsgCodes = ['EPSG:4326'];
    if (!includeCoords) epsgCodes = [];

    const rows = buildExportRows({ includeCoords, includeDates, epsgCode: epsgCodes[0] || 'EPSG:4326' });

    const headerCells = ['<th>Marker</th>', '<th>Dist (m)</th>', '<th>Total (m)</th>'];
    if (includeCoords) {
      epsgCodes.forEach((code) => {
        if (code === 'EPSG:4326') headerCells.push('<th>Lat (WGS84)</th>', '<th>Lon (WGS84)</th>');
        else if (code === 'EPSG:3857') headerCells.push('<th>X (m) (EPSG:3857)</th>', '<th>Y (m) (EPSG:3857)</th>');
        else if (code === 'EPSG:28992') headerCells.push('<th>RD X (m) (EPSG:28992)</th>', '<th>RD Y (m) (EPSG:28992)</th>');
        else if (code === 'UTM') headerCells.push('<th>East (m) (UTM)</th>', '<th>North (m) (UTM)</th>', '<th>Zone (UTM)</th>');
      });
      headerCells.push('<th>Height (m)</th>');
    }
    if (includeExpTimes) headerCells.push('<th>Expected time</th>');
    if (includeDates) headerCells.push('<th>Expected date</th>');
    if (includeExpVol) headerCells.push('<th>Expected vol (m³)</th>');
    if (includeActualTimes) headerCells.push('<th>Actual passing time</th>');
    if (includeMissed) headerCells.push('<th>Missed</th>');

    exportExcelPreviewTable.innerHTML = `<tr>${headerCells.join('')}</tr>`;

    rows.forEach((r) => {
      const row = exportExcelPreviewTable.insertRow();
      const cells = [
        `<td>${r.name ?? ''}</td>`,
        `<td>${r.dist ?? ''}</td>`,
        `<td>${r.total ?? ''}</td>`
      ];

      if (includeCoords) {
        epsgCodes.forEach((code) => {
          if (code === 'EPSG:4326') {
            cells.push(`<td>${r.lat || ''}</td>`, `<td>${r.lon || ''}</td>`);
          } else if (code === 'EPSG:3857') {
            const proj = (r.lat && r.lon) ? projectForEpsg(Number(r.lat), Number(r.lon), 'EPSG:3857') : { x: '', y: '' };
            cells.push(`<td>${r.x || proj.x || ''}</td>`, `<td>${r.y || proj.y || ''}</td>`);
          } else if (code === 'EPSG:28992') {
            const proj = (r.lat && r.lon) ? projectForEpsg(Number(r.lat), Number(r.lon), 'EPSG:28992') : { x: '', y: '' };
            cells.push(`<td>${r.x || proj.x || ''}</td>`, `<td>${r.y || proj.y || ''}</td>`);
          } else if (code === 'UTM') {
            cells.push(`<td>${r.x || ''}</td>`, `<td>${r.y || ''}</td>`, `<td>${r.zone || ''}</td>`);
          }
        });
        cells.push(`<td>${r.alt || ''}</td>`);
      }

      if (includeExpTimes) cells.push(`<td>${r.expTime || ''}</td>`);
      if (includeDates) cells.push(`<td>${r.expDate || ''}</td>`);
      if (includeExpVol) cells.push(`<td>${r.expVol || ''}</td>`);
      if (includeActualTimes) cells.push(`<td>${r.actual || ''}</td>`);
      if (includeMissed) cells.push(`<td>${r.actual === 'Missed' ? 'Yes' : ''}</td>`);

      row.innerHTML = cells.join('');
    });

    if (exportExcelModalContent) {
      const padding = 48;
      const tableWidth = exportExcelPreviewTable.scrollWidth || exportExcelPreviewTable.offsetWidth || 0;
      const maxWidth = Math.floor(window.innerWidth * 0.98);
      const targetWidth = Math.min(tableWidth + padding, maxWidth);
      exportExcelModalContent.style.width = `${targetWidth}px`;
      exportExcelModalContent.style.maxWidth = '98vw';
    }
  };

  const exportExcel = async (opts = {}) => {
    try {
      const XLSX = await ensureSheetJs();
      const includeCoords = opts.includeCoords ?? (exportExcelIncludeCoords ? exportExcelIncludeCoords.checked : true);
      const includeExpTimes = opts.includeExpTimes ?? (exportExcelIncludeExpTimes ? exportExcelIncludeExpTimes.checked : true);
      const includeDates = opts.includeDates ?? (exportExcelIncludeDates ? exportExcelIncludeDates.checked : true);
      const includeExpVol = opts.includeExpVol ?? (exportExcelIncludeExpVol ? exportExcelIncludeExpVol.checked : true);
      const includeActualTimes = opts.includeActualTimes ?? (exportExcelIncludeActualTimes ? exportExcelIncludeActualTimes.checked : true);
      const includeMissed = opts.includeMissed ?? (exportExcelIncludeMissed ? exportExcelIncludeMissed.checked : true);
      let epsgCodes = opts.epsgCodes ?? getExcelEpsgSelections();
      if (includeCoords && (!epsgCodes || epsgCodes.length === 0)) epsgCodes = ['EPSG:4326'];
      if (!includeCoords) epsgCodes = [];
      const rows = [];

      const effectiveVol = getEffectiveVolume();

      rows.push([getExportName()]);
      rows.push(['Launch', state.launch || '-']);
      if (state.mode === 'calc') {
        rows.push(['Mode', 'From flow/volume']);
        rows.push(['Flow (m³/h)', state.flow || '-']);
        rows.push(['Volume (m³/km)', Number.isFinite(effectiveVol) ? effectiveVol.toFixed(2) : '-']);
        rows.push(['Offset (m³/h)', state.offset || 0]);
        const derived = calcSpeed();
        if (derived) rows.push(['Calculated speed (m/h)', derived.toFixed(1)]);
      } else if (state.mode === 'pipe') {
        rows.push(['Mode', 'From pipe specs']);
        rows.push(['Flow (m³/h)', state.flow || '-']);
        rows.push(['Pipe size (in)', state.pipeSize || '-']);
        rows.push(['WT (mm)', state.pipeWt || '-']);
        rows.push(['Volume (m³/km)', Number.isFinite(effectiveVol) ? effectiveVol.toFixed(2) : '-']);
        rows.push(['Offset (m³/h)', state.offset || 0]);
        const derived = calcSpeed();
        if (derived) rows.push(['Calculated speed (m/h)', derived.toFixed(1)]);
      } else {
        rows.push(['Mode', 'Manual speed']);
        rows.push(['Speed (m/h)', state.speed || '-']);
      }
      rows.push([]);

      const header = ['Marker', 'Dist (m)', 'Total (m)'];
      if (includeCoords) {
        epsgCodes.forEach((code) => {
          if (code === 'EPSG:4326') {
            header.push('Lat (WGS84)', 'Lon (WGS84)');
          } else if (code === 'EPSG:3857') {
            header.push('X (m) (EPSG:3857)', 'Y (m) (EPSG:3857)');
          } else if (code === 'EPSG:28992') {
            header.push('RD X (m) (EPSG:28992)', 'RD Y (m) (EPSG:28992)');
          } else if (code === 'UTM') {
            header.push('East (m) (UTM)', 'North (m) (UTM)', 'Zone (UTM)');
          }
        });
        header.push('Height (m)');
      }
      if (includeExpTimes) header.push('Expected time');
      if (includeDates) header.push('Expected date');
      if (includeExpVol) header.push('Expected vol (m³)');
      if (includeActualTimes) header.push('Actual passing time');
      if (includeMissed) header.push('Missed');
      rows.push(header);

      const vRaw = calcSpeed();
      const v = vRaw ? vRaw / 3600 : null;
      let total = 0;
      let ref = toSec(state.launch);

      state.pts.forEach((p, i) => {
        const expTime = ref && v ? toTime(ref) : '';
        const expDate = ref && v ? formatDateFromSeconds(ref, state.launchDate) : '';
        const expVol = Number.isFinite(effectiveVol) ? ((total * Number(effectiveVol)) / 1000).toFixed(1) : '';
        const actual = p.missed ? '' : (p.t || '');
        const missed = p.missed ? 'Yes' : '';
        const hasCoord = p.lat != null && p.lon != null;
        const latVal = hasCoord ? p.lat.toFixed(6) : '';
        const lonVal = hasCoord ? p.lon.toFixed(6) : '';
        const altVal = p.alt != null ? Number(p.alt).toFixed(2) : '';
        const projVal = (v) => (Number.isFinite(v) ? Number(v).toFixed(2) : '');

        const row = [
          p.n || '',
          Number(p.d || 0),
          total
        ];

        if (includeCoords) {
          epsgCodes.forEach((code) => {
            if (!hasCoord) {
              if (code === 'EPSG:4326') row.push('', '');
              else if (code === 'EPSG:3857') row.push('', '');
              else if (code === 'EPSG:28992') row.push('', '');
              else if (code === 'UTM') row.push('', '', '');
              return;
            }
            if (code === 'EPSG:4326') {
              row.push(latVal, lonVal);
            } else if (code === 'EPSG:3857') {
              const proj = projectForEpsg(p.lat, p.lon, 'EPSG:3857');
              row.push(projVal(proj.x), projVal(proj.y));
            } else if (code === 'EPSG:28992') {
              const proj = projectForEpsg(p.lat, p.lon, 'EPSG:28992');
              row.push(projVal(proj.x), projVal(proj.y));
            } else if (code === 'UTM') {
              const proj = projectForEpsg(p.lat, p.lon, 'UTM');
              row.push(projVal(proj.x), projVal(proj.y), proj.zone || '');
            }
          });
          row.push(altVal);
        }

        if (includeExpTimes) row.push(expTime);
        if (includeDates) row.push(expDate);
        if (includeExpVol) row.push(expVol);
        if (includeActualTimes) row.push(actual);
        if (includeMissed) row.push(missed);

        rows.push(row);

        if (p.t && !p.missed) ref = toSec(p.t);
        if (i < state.pts.length - 1 && v) ref += Number(p.d || 0) / v;
        total += Number(p.d || 0);
      });

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, 'Run');

      const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const name = getExportName() + '.xlsx';

      if (window.showSaveFilePicker) {
        try {
          const h = await showSaveFilePicker({
            suggestedName: name,
            types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
          });
          const w = await h.createWritable();
          await w.write(blob);
          await w.close();
          statusEl.textContent = 'Excel saved.';
          return;
        } catch (e) {
          if (e && e.name === 'AbortError') return; // user canceled
        }
      }

      // Fallback to download
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.click();
      statusEl.textContent = 'Excel downloaded.';
    } catch (err) {
      statusEl.textContent = 'Could not export Excel.';
    }
  };

  const markPassNow = (idx) => {
    if (state.locked) return;
    const now = new Date();
    state.pts[idx].t = toTime(now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds());
    render();
  };

  const addMarker = () => {
    if (isEditLocked()) return;
    const idx = state.pts.length - 1;
    state.pts.splice(idx, 0, { n: `M${idx}`, d: 0, t: '', missed: false, lat: null, lon: null });
    render();
  };

  const newProject = () => {
    state = defaultState();
    render();
    if (statusEl) statusEl.textContent = 'Started new project.';
  };

  const removeMarker = (i) => {
    if (isEditLocked() || i === 0 || i === state.pts.length - 1) return;
    state.pts.splice(i, 1);
    render();
  };

  const applyPaste = () => {
    if (isEditLocked() || !markerPaste) return;
    const lines = markerPaste.value.trim().split(/\r?\n/).filter(Boolean);
    if (!lines.length) return;
    state.pts = lines.map((l, i) => {
      const c = l.split(/\t| {2,}/);
      return { n: c[0] ? c[0].trim() : `M${i + 1}`, d: Number(c[1]) || 0, t: '', missed: false, lat: null, lon: null };
    });
    if (state.pts.length < 2) state.pts.push({ n: 'Receiver', d: 0, t: '', missed: false, lat: null, lon: null });
    state.pts[state.pts.length - 1].d = 0;
    render();
    closeMarkerImport();
  };

  const importMarkersXlsx = async (file) => {
    if (!file) return;
    try {
      const XLSX = await ensureSheetJs();
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!rows.length) throw new Error('No rows found');
      let startIdx = 0;
      const header = rows[0].map((h) => String(h).toLowerCase().trim());
      const hasHeader = header.includes('name') && (header.includes('distance') || header.includes('dist') || header.includes('meter'));
      if (hasHeader) startIdx = 1;
      const pts = [];
      for (let i = startIdx; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length < 2) continue;
        const name = String(r[0]).trim();
        const dist = Number(r[1]);
        if (!name || !Number.isFinite(dist)) continue;
        pts.push({ n: name, d: dist, t: '', missed: false, lat: null, lon: null });
      }
      if (pts.length < 2) throw new Error('Need at least two rows with name and distance');
      pts[pts.length - 1].d = 0;
      state.pts = pts;
      render();
      statusEl.textContent = 'Imported markers from Excel.';
      closeMarkerImport();
    } catch (err) {
      statusEl.textContent = 'Could not import markers from Excel.';
      alert(err && err.message ? err.message : 'Import failed');
    } finally {
      if (importXlsxInput) importXlsxInput.value = '';
    }
  };

  const openMarkerImport = () => {
    if (markerImportModal) markerImportModal.classList.add('show');
  };

  const closeMarkerImport = () => {
    if (markerImportModal) markerImportModal.classList.remove('show');
  };

  const downloadMarkerSample = async () => {
    try {
      const XLSX = await ensureSheetJs();
      const rows = [
        ['Name', 'Distance'],
        ['Launcher', 500],
        ['M1', 600],
        ['M2', 800],
        ['Receiver', 0],
        [],
        ['in this example Launcher 500 is the distance from Launcher to M1. therefore receiver is always 0']
      ];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, 'Markers');
      const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const name = 'markers_example.xlsx';
      if (window.showSaveFilePicker) {
        try {
          const h = await showSaveFilePicker({
            suggestedName: name,
            types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
          });
          const w = await h.createWritable();
          await w.write(blob);
          await w.close();
          statusEl.textContent = 'Example file saved.';
          return;
        } catch (e) {
          if (e && e.name === 'AbortError') return;
        }
      }
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.click();
      statusEl.textContent = 'Example file downloaded.';
    } catch (err) {
      statusEl.textContent = 'Could not generate sample.';
    }
  };

  const saveJSON = async () => {
    const name = buildExportBaseName() + '.json';
    const data = JSON.stringify(state, null, 2);
    if (window.showSaveFilePicker) {
      try {
        const h = await showSaveFilePicker({ suggestedName: name, types: [{ accept: { 'application/json': ['.json'] } }] });
        const w = await h.createWritable();
        await w.write(data);
        await w.close();
        statusEl.textContent = 'Saved to file.';
        return;
      } catch (e) {
        if (e && e.name === 'AbortError') return; // user canceled
      }
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([data], { type: 'application/json' }));
    a.download = name;
    a.click();
    statusEl.textContent = 'Downloaded JSON.';
  };

  const loadJSON = () => fileInput.click();

  let gpsCoords = null;
  let gpsSegments = null;
  let gpsWorking = [];
  let gpsActiveIndex = null;
  let gpsCsvHeaders = [];
  let gpsCsvRows = [];
  let gpsSourceEpsg = 'EPSG:4326';
  let gpsSourceUtmZone = '';

  const setGpsStatus = (msg) => { if (gpsStatus) gpsStatus.textContent = msg; };

  const setGpsData = (coords, opts = {}) => {
    gpsSourceEpsg = (opts.epsg || 'EPSG:4326').toUpperCase();
    gpsSourceUtmZone = opts.utmZone || '';
    gpsWorking = (coords || []).map((p) => ({
      lat: p.lat,
      lon: p.lon,
      alt: p.alt ?? null,
      name: p.name || '',
      include: true,
      utm: p.utm ? { east: p.utm.east, north: p.utm.north } : null
    }));
    gpsActiveIndex = null;
    renderGpsList();
    renderGpsPreview();
  };

  const parseCsvLine = (line, sep) => {
    const out = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }
      if (ch === sep && !inQuotes) {
        out.push(cur);
        cur = '';
        continue;
      }
      cur += ch;
    }
    out.push(cur);
    return out.map((v) => v.trim());
  };

  const parseCsvText = (text) => {
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (!lines.length) return { headers: [], rows: [] };
    const sample = lines.slice(0, 5).join('\n');
    const sep = sample.includes(';') ? ';' : (sample.includes('\t') ? '\t' : ',');
    const headerRow = parseCsvLine(lines[0], sep);
    const rows = lines.slice(1).map((l) => parseCsvLine(l, sep));
    const headers = headerRow.map((h, i) => h || `Column ${i + 1}`);
    return { headers, rows };
  };

  const renderCsvPreview = () => {
    if (!gpsCsvPreview) return;
    const headers = gpsCsvHeaders;
    const epsg = gpsCsvEpsg ? gpsCsvEpsg.value : 'EPSG:4326';
    const latIdx = gpsCsvLat ? Number(gpsCsvLat.value) : NaN;
    const lonIdx = gpsCsvLon ? Number(gpsCsvLon.value) : NaN;
    gpsCsvPreview.innerHTML = `
      <tr>${headers.map((h) => `<th>${h}</th>`).join('')}</tr>`;
    gpsCsvRows.slice(0, 20).forEach((row) => {
      const tr = gpsCsvPreview.insertRow();
      tr.innerHTML = headers.map((_, i) => {
        let val = row[i] ?? '';
        if (epsg === 'UTM' && (i === latIdx || i === lonIdx)) {
          const parsed = parseUtmCsvValue(val);
          if (Number.isFinite(parsed)) val = parsed.toFixed(2);
        }
        return `<td>${val}</td>`;
      }).join('');
    });
  };

  const renderCsvColumns = () => {
    const opts = gpsCsvHeaders.map((h, i) => `<option value="${i}">${h}</option>`).join('');
    if (gpsCsvName) gpsCsvName.innerHTML = `<option value="">(none)</option>${opts}`;
    if (gpsCsvLat) gpsCsvLat.innerHTML = opts;
    if (gpsCsvLon) gpsCsvLon.innerHTML = opts;
    if (gpsCsvAlt) gpsCsvAlt.innerHTML = `<option value="">(none)</option>${opts}`;
    const pickIndex = (needle) => gpsCsvHeaders.findIndex((h) => h.toLowerCase().includes(needle));
    const nameIdx = pickIndex('name');
    const latIdx = pickIndex('lat');
    const lonIdx = pickIndex('lon');
    const altIdx = pickIndex('alt') >= 0 ? pickIndex('alt') : pickIndex('height');
    if (gpsCsvName && nameIdx >= 0) gpsCsvName.value = String(nameIdx);
    if (gpsCsvLat && latIdx >= 0) gpsCsvLat.value = String(latIdx);
    if (gpsCsvLon && lonIdx >= 0) gpsCsvLon.value = String(lonIdx);
    if (gpsCsvAlt && altIdx >= 0) gpsCsvAlt.value = String(altIdx);
  };

  const openCsvModal = () => {
    if (!gpsCsvModal) return;
    if (gpsCsvUtmZone) gpsCsvUtmZone.value = '';
    renderCsvColumns();
    renderCsvPreview();
    gpsCsvModal.classList.add('show');
  };

  const closeCsvModal = () => {
    if (!gpsCsvModal) return;
    gpsCsvModal.classList.remove('show');
  };

  const updateCsvEpsgUi = () => {
    const isUtm = gpsCsvEpsg && gpsCsvEpsg.value === 'UTM';
    const label = document.getElementById('gpsCsvUtmZoneLabel');
    if (label) label.style.display = isUtm ? 'block' : 'none';
    if (!isUtm && gpsCsvUtmZone) gpsCsvUtmZone.value = '';
    const latLabel = document.getElementById('gpsCsvLatLabel');
    const lonLabel = document.getElementById('gpsCsvLonLabel');
    if (latLabel) latLabel.textContent = isUtm ? 'Easting (X) column' : 'Latitude column';
    if (lonLabel) lonLabel.textContent = isUtm ? 'Northing (Y) column' : 'Longitude column';
  };

  if (utmCountrySelect) {
    utmCountrySelect.onchange = () => {
      const idx = utmCountrySelect.value !== '' ? Number(utmCountrySelect.value) : null;
      if (idx == null || !UTM_COUNTRIES[idx]) return;
      const c = UTM_COUNTRIES[idx];
      const zone = zoneFromLonLat(c.lon, c.lat);
      utmPickerZone = zone;
      const zoneStr = formatZone(zone);
      if (utmCountryReadout) utmCountryReadout.textContent = `${c.name} → ${zoneStr}`;
    };
  }

  if (utmMap) {
    utmMap.addEventListener('click', (e) => {
      const rect = utmMap.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      const lon = (x / rect.width) * 360 - 180;
      const lat = 90 - (y / rect.height) * 180;
      const zone = zoneFromLonLat(lon, lat);
      utmPickerZone = zone;
      if (utmMapReadout) utmMapReadout.textContent = `Lon: ${lon.toFixed(2)}, Lat: ${lat.toFixed(2)}, Zone: ${formatZone(zone)}`;
      if (utmPickerTarget && utmPickerZone) {
        utmPickerTarget.value = formatZone(utmPickerZone);
      }
    });
  }

  if (utmPickerApply) {
    utmPickerApply.onclick = () => {
      if (!utmPickerTarget) return;
      if (!utmPickerZone) {
        if (utmCountryReadout) utmCountryReadout.textContent = 'Pick a country or click on the map first.';
        return;
      }
      utmPickerTarget.value = formatZone(utmPickerZone);
      closeUtmPicker();
    };
  }

  if (utmPickerClose) utmPickerClose.onclick = closeUtmPicker;
  if (utmPickerModal) bindModalClose(utmPickerModal, closeUtmPicker);
  if (gpsUtmZonePick && gpsUtmZoneInput) gpsUtmZonePick.onclick = () => openUtmPicker(gpsUtmZoneInput);
  if (gpsCsvUtmZonePick && gpsCsvUtmZone) gpsCsvUtmZonePick.onclick = () => openUtmPicker(gpsCsvUtmZone);

  const renderGpsPreview = () => {
    if (!gpsPreview) return;
    const coords = gpsWorking.filter((p) => p.include);
    gpsPreview.innerHTML = '';
    gpsSegments = null;
    gpsCoords = coords;
    if (!coords || coords.length < 2) {
      setGpsStatus('Need at least two selected points.');
      return;
    }
    gpsSegments = [];
    let total = 0;
    const isUtmDisplay = gpsSourceEpsg === 'UTM' && gpsSourceUtmZone;
    const colA = isUtmDisplay ? 'Easting (m)' : 'Lat';
    const colB = isUtmDisplay ? 'Northing (m)' : 'Lon';
    gpsPreview.innerHTML = `<tr><th>#</th><th>Name</th><th>${colA}</th><th>${colB}</th><th>Alt</th><th>Segment (m)</th><th>Cumulative (m)</th></tr>`;
    coords.forEach((p, i) => {
      let seg = '';
      if (i > 0) {
        const d = haversineMeters(coords[i - 1].lat, coords[i - 1].lon, p.lat, p.lon);
        gpsSegments.push(d);
        total += d;
        seg = Math.round(d).toLocaleString();
      }
      const aVal = isUtmDisplay && p.utm ? p.utm.east.toFixed(2) : p.lat.toFixed(6);
      const bVal = isUtmDisplay && p.utm ? p.utm.north.toFixed(2) : p.lon.toFixed(6);
      const row = gpsPreview.insertRow();
      row.innerHTML = `<td>${i + 1}</td><td>${p.name || ''}</td><td>${aVal}</td><td>${bVal}</td><td>${p.alt != null ? p.alt : ''}</td><td>${seg}</td><td>${i === 0 ? '' : Math.round(total).toLocaleString()}</td>`;
    });
    setGpsStatus(`Selected ${coords.length} points · total ${Math.round(total).toLocaleString()} m`);
  };

  const renderGpsList = () => {
    if (!gpsSelectTable) return;
    const isUtmDisplay = gpsSourceEpsg === 'UTM' && gpsSourceUtmZone;
    const colA = isUtmDisplay ? 'Easting (m)' : 'Lat';
    const colB = isUtmDisplay ? 'Northing (m)' : 'Lon';
    gpsSelectTable.innerHTML = `<tr><th>#</th><th>Name</th><th>${colA}</th><th>${colB}</th><th>Alt</th><th>Include</th></tr>`;
    gpsWorking.forEach((p, i) => {
      const row = gpsSelectTable.insertRow();
      row.dataset.idx = String(i);
      if (gpsActiveIndex === i) row.classList.add('active');
      const aVal = isUtmDisplay && p.utm ? p.utm.east.toFixed(2) : p.lat.toFixed(6);
      const bVal = isUtmDisplay && p.utm ? p.utm.north.toFixed(2) : p.lon.toFixed(6);
      row.innerHTML = `<td>${i + 1}</td><td>${p.name || ''}</td><td>${aVal}</td><td>${bVal}</td><td>${p.alt != null ? p.alt : ''}</td><td><input type="checkbox" data-include="${i}" ${p.include ? 'checked' : ''}></td>`;
    });
  };

  const openGpsModal = () => {
    if (gpsModal) gpsModal.classList.add('show');
  };

  const closeGpsModal = () => {
    if (gpsModal) gpsModal.classList.remove('show');
  };

  const applyGpsToMarkers = () => {
    if (state.locked) {
      alert('Unlock to apply GPS distances.');
      return;
    }
    if (!gpsCoords || !gpsSegments || gpsCoords.length < 2) {
      alert('Load GPS data first.');
      return;
    }
    const pts = gpsCoords.map((p, idx) => {
      const isFirst = idx === 0;
      const isLast = idx === gpsCoords.length - 1;
      const fallback = isFirst ? 'Launcher' : (isLast ? 'Receiver' : `M${idx}`);
      const name = p.name && p.name.trim() ? p.name.trim() : fallback;
      const dist = isLast ? 0 : Math.round(gpsSegments[idx] || 0);
      return { n: name, d: dist, t: '', missed: false, lat: p.lat ?? null, lon: p.lon ?? null, alt: p.alt ?? null };
    });
    state.pts = pts;
    render();
    setGpsStatus('Applied to markers.');
    closeGpsModal();
  };

  const readKmz = async (file) => {
    const JSZip = await ensureJsZip();
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const kmlEntry = Object.values(zip.files).find((f) => f.name.toLowerCase().endsWith('.kml'));
    if (!kmlEntry) throw new Error('No KML found inside KMZ');
    return kmlEntry.async('text');
  };

  const getGpsImportEpsg = () => (gpsEpsgSelect ? gpsEpsgSelect.value : 'EPSG:4326');
  const getGpsImportUtmZone = () => (gpsUtmZoneInput ? gpsUtmZoneInput.value : '');

  const updateGpsEpsgUi = () => {
    const isUtm = gpsEpsgSelect && gpsEpsgSelect.value === 'UTM';
    const label = document.getElementById('gpsUtmZoneLabel');
    if (label) label.style.display = isUtm ? 'block' : 'none';
    if (!isUtm && gpsUtmZoneInput) gpsUtmZoneInput.value = '';
  };

  const handleGpsFile = async (file) => {
    if (!file) return;
    try {
      setGpsStatus('Loading file...');
      const name = file.name.toLowerCase();
      let kmlText = '';
      if (name.endsWith('.kmz')) {
        kmlText = await readKmz(file);
      } else if (name.endsWith('.kml')) {
        kmlText = await file.text();
      } else {
        throw new Error('Use a KMZ or KML file');
      }
      const coords = parseKmlCoordinates(kmlText);
      if (!coords.length) throw new Error('No coordinates found in file');
      const epsg = getGpsImportEpsg();
      const utmZoneHint = getGpsImportUtmZone();
      const converted = convertImportCoords(coords, epsg, { utmZoneHint });
      if (!converted.length) throw new Error('No usable coordinates after conversion');
      setGpsData(converted, { epsg, utmZone: utmZoneHint });
    } catch (err) {
      setGpsStatus('Could not load file.');
      alert(err && err.message ? err.message : 'Failed to load file');
    } finally {
      if (gpsFileInput) gpsFileInput.value = '';
    }
  };

  const handleGpsText = () => {
    try {
      const epsg = getGpsImportEpsg();
      const utmZoneHint = getGpsImportUtmZone();
      const parsed = parseManualCoordinates(gpsCoordsInput ? gpsCoordsInput.value : '', { epsgCode: epsg, utmZoneHint });
      const coords = convertImportCoords(parsed, epsg, { utmZoneHint });
      if (!coords.length) {
        setGpsStatus('Could not parse pasted coordinates.');
        return;
      }
      setGpsData(coords, { epsg, utmZone: utmZoneHint });
    } catch (err) {
      setGpsStatus(err && err.message ? err.message : 'Could not parse pasted coordinates.');
    }
  };

  const downloadGpsSample = async () => {
    try {
      const XLSX = await ensureSheetJs();
      const rows = [
        ['Example 1'],
        ['Name', 'Latitude', 'Longitude'],
        ['Zero point', '51.690109', '-5.023715'],
        ['MP2 flange gasoline', '51.692755', '-5.028808'],
        ['MP3 offtake gasoline', '51.697523', '-5.029099'],
        ['End point', '51.697425', '-5.019123'],
        [],
        ['Example 2'],
        ['Name', 'Latitude', 'Longitude'],
        ['Zero point', 'N 51.6860106823', 'W 5.0236768364'],
        ['MP2 flange gasoline', 'N 51.6901092629', 'W 5.0237152286'],
        ['MP3 offtake gasoline', 'N 51.6927552979', 'W 5.028808381'],
        ['End point', 'N 51.697522647', 'W 5.0290986279']
      ];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, 'Sample');
      const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const name = 'gps_import_example.xlsx';
      if (window.showSaveFilePicker) {
        try {
          const h = await showSaveFilePicker({
            suggestedName: name,
            types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
          });
          const w = await h.createWritable();
          await w.write(blob);
          await w.close();
          setGpsStatus('Example Excel saved.');
          return;
        } catch (e) {
          if (e && e.name === 'AbortError') return;
        }
      }
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.click();
      setGpsStatus('Example Excel downloaded.');
    } catch (err) {
      setGpsStatus('Could not generate sample.');
    }
  };

  fileInput.onchange = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = () => {
      try {
        state = JSON.parse(r.result);
        normalizeState();
        render();
        statusEl.textContent = 'Loaded JSON.';
      } catch (err) {
        statusEl.textContent = 'Could not parse file.';
      }
    };
    r.readAsText(f);
  };

  function bindModalClose(modalEl, closeFn) {
    if (!modalEl || typeof closeFn !== 'function') return;
    modalEl.addEventListener('click', (e) => {
      if (e.target === modalEl || (e.target && e.target.classList && e.target.classList.contains('modal-backdrop'))) closeFn();
    });
  }

  if (gpsBtn) gpsBtn.onclick = openGpsModal;
  if (gpsClose) gpsClose.onclick = closeGpsModal;
  bindModalClose(gpsModal, closeGpsModal);
  if (gpsLoadFile) gpsLoadFile.onclick = () => { if (gpsFileInput) gpsFileInput.click(); };
  if (gpsFileInput) gpsFileInput.onchange = (e) => { const f = e.target.files[0]; handleGpsFile(f); };
  if (gpsParseText) gpsParseText.onclick = handleGpsText;
  if (gpsSampleXlsx) gpsSampleXlsx.onclick = downloadGpsSample;
  if (gpsApply) gpsApply.onclick = applyGpsToMarkers;
  if (gpsEpsgSelect) {
    gpsEpsgSelect.onchange = updateGpsEpsgUi;
    updateGpsEpsgUi();
  }
  if (gpsSelectAll) gpsSelectAll.onclick = () => { gpsWorking.forEach((p) => (p.include = true)); renderGpsList(); renderGpsPreview(); };
  if (gpsSelectNone) gpsSelectNone.onclick = () => { gpsWorking.forEach((p) => (p.include = false)); renderGpsList(); renderGpsPreview(); };
  const swapGps = (dir) => {
    if (gpsActiveIndex == null) return;
    const target = gpsActiveIndex + dir;
    if (target < 0 || target >= gpsWorking.length) return;
    const tmp = gpsWorking[gpsActiveIndex];
    gpsWorking[gpsActiveIndex] = gpsWorking[target];
    gpsWorking[target] = tmp;
    gpsActiveIndex = target;
    renderGpsList();
    renderGpsPreview();
  };
  if (gpsMoveUp) gpsMoveUp.onclick = () => swapGps(-1);
  if (gpsMoveDown) gpsMoveDown.onclick = () => swapGps(1);
  if (gpsSelectTable) {
    gpsSelectTable.addEventListener('click', (e) => {
      const t = e.target;
      if (t.dataset && t.dataset.include) {
        const idx = Number(t.dataset.include);
        if (Number.isInteger(idx) && gpsWorking[idx]) {
          gpsWorking[idx].include = t.checked;
          renderGpsPreview();
        }
        return;
      }
      let row = t.closest('tr');
      if (!row || !row.dataset.idx) return;
      gpsActiveIndex = Number(row.dataset.idx);
      renderGpsList();
    });
  }

  if (gpsCsvBtn) gpsCsvBtn.onclick = () => { if (gpsCsvInput) gpsCsvInput.click(); };
  if (gpsCsvInput) {
    gpsCsvInput.onchange = async (e) => {
      const f = e.target.files[0];
      if (!f) return;
      try {
        const text = await f.text();
        const parsed = parseCsvText(text);
        gpsCsvHeaders = parsed.headers;
        gpsCsvRows = parsed.rows;
        openCsvModal();
        updateCsvEpsgUi();
      } catch (err) {
        if (gpsStatus) gpsStatus.textContent = 'Could not read CSV file.';
      } finally {
        if (gpsCsvInput) gpsCsvInput.value = '';
      }
    };
  }
  if (gpsCsvClose) gpsCsvClose.onclick = closeCsvModal;
  bindModalClose(gpsCsvModal, closeCsvModal);
  if (gpsCsvEpsg) gpsCsvEpsg.onchange = () => { updateCsvEpsgUi(); renderCsvPreview(); };
  if (gpsCsvLat) gpsCsvLat.onchange = renderCsvPreview;
  if (gpsCsvLon) gpsCsvLon.onchange = renderCsvPreview;
  if (gpsCsvApply) {
    gpsCsvApply.onclick = () => {
      const nameIdx = gpsCsvName && gpsCsvName.value !== '' ? Number(gpsCsvName.value) : null;
      const latIdx = gpsCsvLat ? Number(gpsCsvLat.value) : NaN;
      const lonIdx = gpsCsvLon ? Number(gpsCsvLon.value) : NaN;
      const altIdx = gpsCsvAlt && gpsCsvAlt.value !== '' ? Number(gpsCsvAlt.value) : null;
      if (!Number.isInteger(latIdx) || !Number.isInteger(lonIdx)) {
        if (gpsStatus) gpsStatus.textContent = 'Select valid X/Y columns first.';
        return;
      }
      const coords = [];
      const epsg = gpsCsvEpsg ? gpsCsvEpsg.value : 'EPSG:4326';
      const useLatLonCols = false;
      const latRefIdx = -1;
      const lonRefIdx = -1;
      if (epsg === 'UTM') {
        const zoneStr = gpsCsvUtmZone ? gpsCsvUtmZone.value : '';
        if (!parseUtmZoneString(zoneStr)) {
          if (gpsStatus) gpsStatus.textContent = 'Please select a UTM zone before applying.';
          alert('Please select a UTM zone (e.g., 30N) before applying.');
          if (gpsCsvUtmZone) gpsCsvUtmZone.focus();
          return;
        }
      }
      const utmZoneHint = epsg === 'UTM' && gpsCsvUtmZone ? gpsCsvUtmZone.value : '';
      gpsCsvRows.forEach((row) => {
        const rawA = row[latIdx];
        const rawB = row[lonIdx];
        const a = epsg === 'UTM' ? parseUtmCsvValue(rawA) : parseCoordText(rawA);
        const b = epsg === 'UTM' ? parseUtmCsvValue(rawB) : parseCoordText(rawB);
        let lat = epsg === 'UTM' ? b : a;
        let lon = epsg === 'UTM' ? a : b;
        let latLonOverride = null;
        if (epsg === 'UTM' && useLatLonCols && latRefIdx >= 0 && lonRefIdx >= 0) {
          const latRef = parseCoordText(row[latRefIdx]);
          const lonRef = parseCoordText(row[lonRefIdx]);
          if (Number.isFinite(latRef) && Number.isFinite(lonRef)) {
            lat = latRef;
            lon = lonRef;
            latLonOverride = { lat: latRef, lon: lonRef };
          }
        }
        const alt = altIdx != null ? parseCoordText(row[altIdx]) : null;
        const name = nameIdx != null ? String(row[nameIdx] ?? '').trim() : '';
        if (Number.isFinite(lat) && Number.isFinite(lon)) {
          const item = { lat, lon, alt: Number.isFinite(alt) ? alt : null, name, latLonOverride };
          if (epsg === 'UTM') {
            item.utm = { east: a, north: b };
            if (utmZoneHint) item.zoneHint = utmZoneHint;
          }
          coords.push(item);
        }
      });
      if (!coords.length) {
        if (gpsStatus) gpsStatus.textContent = 'No valid coordinates found in selected columns.';
        return;
      }
      try {
        let converted = [];
        if (epsg === 'EPSG:4326') {
          converted = coords;
        } else if (epsg === 'UTM' && useLatLonCols) {
          coords.forEach((item) => {
            if (item.latLonOverride) {
              converted.push({ lat: item.latLonOverride.lat, lon: item.latLonOverride.lon, alt: item.alt ?? null, name: item.name || '' });
            } else {
              const conv = convertImportCoords([item], epsg, { utmZoneHint });
              if (conv && conv[0]) converted.push(conv[0]);
            }
          });
        } else {
          converted = convertImportCoords(coords, epsg, { utmZoneHint });
        }
        if (epsg === 'UTM') {
          converted = converted.map((p, i) => ({ ...p, utm: coords[i] && coords[i].utm ? coords[i].utm : null }));
        }
        setGpsData(converted, { epsg, utmZone: utmZoneHint });
        if (gpsStatus) gpsStatus.textContent = `Loaded ${coords.length} points from CSV.`;
        openGpsModal();
        closeCsvModal();
      } catch (err) {
        if (gpsStatus) gpsStatus.textContent = err && err.message ? err.message : 'Could not convert CSV coordinates.';
      }
    };
  }

  if (addMarkerBtn) addMarkerBtn.onclick = () => { if (!isEditLocked()) addMarker(); };
  if (saveJsonBtn) saveJsonBtn.onclick = saveJSON;
  if (loadJsonBtn) loadJsonBtn.onclick = loadJSON;
  if (importXlsxBtn) importXlsxBtn.onclick = openMarkerImport;
  if (importXlsxInput) importXlsxInput.onchange = (e) => { const f = e.target.files[0]; importMarkersXlsx(f); };
  if (markerImportClose) markerImportClose.onclick = closeMarkerImport;
  bindModalClose(markerImportModal, closeMarkerImport);
  if (markerImportChoose) markerImportChoose.onclick = () => { if (importXlsxInput) importXlsxInput.click(); };
  if (markerDownloadSample) markerDownloadSample.onclick = downloadMarkerSample;
  if (markerApplyPaste) markerApplyPaste.onclick = () => { if (!isEditLocked()) applyPaste(); };
  if (newProjectBtn) newProjectBtn.onclick = () => {
    if (isEditLocked()) return;
    const ok = window.confirm('Start a new project? This clears the current marker list.');
    if (!ok) return;
    newProject();
  };
  exportBtn.onclick = openExportPreview;
  if (exportExcelBtn) exportExcelBtn.onclick = openExportExcelModal;
  if (displayCoordsBtn) displayCoordsBtn.onclick = () => { state.showCoords = !state.showCoords; render(); };
  if (toggleNavBtn) toggleNavBtn.onclick = () => { state.showNavigate = !state.showNavigate; render(); };
  if (displayEpsgSelect) displayEpsgSelect.onchange = () => { state.displayEpsg = displayEpsgSelect.value || 'EPSG:4326'; render(); };

  if (exportPreviewClose) exportPreviewClose.onclick = closeExportPreview;
  bindModalClose(exportPreviewModal, closeExportPreview);
  if (exportPreviewConfirm) exportPreviewConfirm.onclick = () => exportPdf({});
  if (exportIncludeCoords) exportIncludeCoords.onchange = renderExportPreview;
  if (exportIncludeDates) exportIncludeDates.onchange = renderExportPreview;
  if (exportIncludeExpTimes) exportIncludeExpTimes.onchange = renderExportPreview;
  if (exportIncludeActualTimes) exportIncludeActualTimes.onchange = renderExportPreview;
  if (exportEpsgSelect) exportEpsgSelect.onchange = renderExportPreview;
  if (exportPdfName) exportPdfName.oninput = () => {
    state.exportName = exportPdfName.value;
    if (exportExcelName) exportExcelName.value = exportPdfName.value;
  };

  const updateExcelCoordUi = () => {
    const enabled = exportExcelIncludeCoords ? exportExcelIncludeCoords.checked : true;
    if (exportExcelEpsg4326) exportExcelEpsg4326.disabled = !enabled;
    if (exportExcelEpsg3857) exportExcelEpsg3857.disabled = !enabled;
    if (exportExcelEpsg28992) exportExcelEpsg28992.disabled = !enabled;
    if (exportExcelEpsgUtm) exportExcelEpsgUtm.disabled = !enabled;
  };

  if (exportExcelClose) exportExcelClose.onclick = closeExportExcelModal;
  bindModalClose(exportExcelModal, closeExportExcelModal);
  if (exportExcelConfirm) exportExcelConfirm.onclick = () => {
    exportExcel({});
    closeExportExcelModal();
  };
  if (exportExcelName) exportExcelName.oninput = () => {
    state.exportName = exportExcelName.value;
    if (exportPdfName) exportPdfName.value = exportExcelName.value;
  };
  if (exportKmzBtn) exportKmzBtn.onclick = openExportKmzModal;
  if (graphsBtn) graphsBtn.onclick = showGraphs;
  if (exportKmzClose) exportKmzClose.onclick = closeExportKmzModal;
  bindModalClose(exportKmzModal, closeExportKmzModal);
  if (exportKmzConfirm) exportKmzConfirm.onclick = () => {
    exportKmz();
    closeExportKmzModal();
  };
  if (exportKmzApplySymbol) {
    exportKmzApplySymbol.onclick = () => {
      if (!exportKmzSymbolAll) return;
      const val = exportKmzSymbolAll.value;
      kmzWorking.forEach((m) => {
        if (m.hasCoord) m.symbol = val;
      });
      renderKmzTable();
    };
  }
  if (exportKmzToggleDesc) {
    exportKmzToggleDesc.onclick = () => {
      kmzShowDescriptions = !kmzShowDescriptions;
      renderKmzTable();
    };
  }
  if (exportKmzTable) {
    exportKmzTable.addEventListener('change', (e) => {
      const t = e.target;
      if (t.dataset && t.dataset.kmzInclude != null) {
        const idx = Number(t.dataset.kmzInclude);
        const item = kmzWorking.find((m) => m.idx === idx);
        if (item) item.include = t.checked;
        return;
      }
      if (t.dataset && t.dataset.kmzSymbol != null) {
        const idx = Number(t.dataset.kmzSymbol);
        const item = kmzWorking.find((m) => m.idx === idx);
        if (item) item.symbol = t.value;
        renderKmzTable();
        return;
      }
      if (t.dataset && t.dataset.kmzDesc != null) {
        const idx = Number(t.dataset.kmzDesc);
        const item = kmzWorking.find((m) => m.idx === idx);
        if (item) item.description = t.value;
      }
    });
  }
  if (exportExcelIncludeCoords) exportExcelIncludeCoords.onchange = () => { updateExcelCoordUi(); renderExcelPreview(); };
  if (exportExcelIncludeDates) exportExcelIncludeDates.onchange = renderExcelPreview;
  if (exportExcelIncludeExpTimes) exportExcelIncludeExpTimes.onchange = renderExcelPreview;
  if (exportExcelIncludeExpVol) exportExcelIncludeExpVol.onchange = renderExcelPreview;
  if (exportExcelIncludeActualTimes) exportExcelIncludeActualTimes.onchange = renderExcelPreview;
  if (exportExcelIncludeMissed) exportExcelIncludeMissed.onchange = renderExcelPreview;
  if (exportExcelEpsg4326) exportExcelEpsg4326.onchange = renderExcelPreview;
  if (exportExcelEpsg3857) exportExcelEpsg3857.onchange = renderExcelPreview;
  if (exportExcelEpsg28992) exportExcelEpsg28992.onchange = renderExcelPreview;
  if (exportExcelEpsgUtm) exportExcelEpsgUtm.onchange = renderExcelPreview;
  updateExcelCoordUi();
  renderExcelPreview();

  projectEl.oninput = () => { state.project = projectEl.value; saveLocal(); };
  if (clientEl) clientEl.oninput = () => { state.client = clientEl.value; saveLocal(); };
  if (pipeIdEl) pipeIdEl.oninput = () => { state.pipeId = pipeIdEl.value; saveLocal(); };
  if (syncIdEl) syncIdEl.oninput = () => {
    state.syncId = syncIdEl.value;
    saveLocal();
    if (state.syncEnabled) startSync();
  };
  if (syncEnabledEl) syncEnabledEl.onchange = () => {
    state.syncEnabled = !!syncEnabledEl.checked;
    saveLocal();
    if (state.syncEnabled) startSync();
    else stopSync();
  };
  if (cloudFileNameEl) cloudFileNameEl.oninput = () => {
    localStorage.setItem(cloudNameKey, cloudFileNameEl.value.trim());
  };
  if (cloudFilesListEl) cloudFilesListEl.onchange = () => {
    const selectedOption = cloudFilesListEl.options[cloudFilesListEl.selectedIndex];
    if (selectedOption && cloudFileNameEl) {
      cloudFileNameEl.value = selectedOption.dataset.name || selectedOption.value || '';
      localStorage.setItem(cloudNameKey, cloudFileNameEl.value.trim());
    }
  };
  if (cloudSaveBtn) cloudSaveBtn.onclick = saveCloudFile;
  if (cloudLoadBtn) cloudLoadBtn.onclick = loadCloudFile;
  if (cloudDeleteBtn) cloudDeleteBtn.onclick = deleteCloudFile;
  if (cloudRefreshBtn) cloudRefreshBtn.onclick = loadCloudList;
  speedEl.oninput = () => { state.speed = speedEl.value; render(); };
  flowEl.oninput = () => { state.flow = flowEl.value; render(); };
  flowOffsetEl.oninput = () => {
    const raw = flowOffsetEl.value;
    if (raw === '-' || raw === '') return; // allow typing minus before number
    state.offset = Number(raw) || 0;
    render();
  };
  volEl.oninput = () => { state.vol = volEl.value; render(); };
  if (pipeSizeEl) {
    pipeSizeEl.oninput = () => {
      state.pipeSize = pipeSizeEl.value;
      const nps = parseDecimal(state.pipeSize);
      const stdWt = getAsmeStdWtMm(nps);
      if (Number.isFinite(stdWt) && (state.pipeWtAuto || !state.pipeWt)) {
        state.pipeWt = stdWt.toFixed(2);
        state.pipeWtAuto = true;
      } else if (!Number.isFinite(stdWt) && state.pipeWtAuto) {
        state.pipeWt = '';
      }
      render();
    };
  }
  if (pipeWtEl) pipeWtEl.oninput = () => {
    state.pipeWt = pipeWtEl.value;
    state.pipeWtAuto = false;
    render();
  };
  if (launchDateEl) launchDateEl.oninput = () => { state.launchDate = launchDateEl.value; render(); };
  launchEl.oninput = () => { state.launch = launchEl.value; render(); };

  if (modeSelectEl) {
    modeSelectEl.onchange = () => {
      state.mode = modeSelectEl.value || 'manual';
      render();
    };
  }

  if (userModeToggle) {
    userModeToggle.onchange = () => {
      state.userMode = userModeToggle.checked ? 'expert' : 'operator';
      render();
    };
  }

  if (themeToggle) {
    themeToggle.onchange = () => {
      state.theme = themeToggle.checked ? 'dark' : 'light';
      render();
    };
  }

  // Update state live, but render only on commit (change or Enter) to avoid disrupting typing.
  table.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.name) {
      if (isEditLocked()) return;
      state.pts[Number(t.dataset.name)].n = t.value;
      return; // defer render until commit
    }
    if (t.dataset.dist) {
      if (isEditLocked()) return;
      state.pts[Number(t.dataset.dist)].d = Number(t.value) || 0;
      updateExpectations(); // keep expected times in sync while typing distances
      return; // render on commit
    }
    if (t.dataset.time) {
      const idx = Number(t.dataset.time);
      state.pts[idx].t = t.value;
      state.pts[idx].missed = false;
      updateExpectations();
      return; // render on commit
    }
  });

  table.addEventListener('change', (e) => {
    const t = e.target;
    if (isEditLocked() && (t.dataset.name || t.dataset.dist)) return;
    if (t.dataset.name || t.dataset.dist) {
      render();
    }
    // Time changes are handled live without full render to preserve focus while typing
  });

  table.addEventListener('keydown', (e) => {
    const t = e.target;
    if (e.key === 'Enter' && (t.dataset.name || t.dataset.dist)) {
      if (isEditLocked()) {
        e.preventDefault();
        return;
      }
      e.preventDefault();
      render();
    }
    if (e.key === 'Enter' && t.dataset.time) {
      e.preventDefault();
      // Avoid full render on Enter to keep focus stable
    }
  });

  table.addEventListener('click', (e) => {
    const t = e.target;
    if (t.dataset.del) removeMarker(Number(t.dataset.del));
    if (t.dataset.missed) {
      const idx = Number(t.dataset.missed);
      const p = state.pts[idx];
      if (p) {
        const nowMissed = !p.missed;
        p.missed = nowMissed;
        p.t = '';
        render();
      }
    }
  });

  const registerSW = async () => {
    if ('serviceWorker' in navigator) {
      try {
        await navigator.serviceWorker.register('service-worker.js');
        statusEl.textContent = 'Offline cache ready.';
      } catch (e) {}
    }
  };

  normalizeState();
  const cachedCloudName = localStorage.getItem(cloudNameKey);
  if (cloudFileNameEl && cachedCloudName) cloudFileNameEl.value = cachedCloudName;
  render();
  if (state.syncEnabled) startSync();
  loadCloudList();
  registerSW();
})();
