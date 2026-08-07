// ==UserScript==
// @name         Carnet de bord — Alternance (kDrive)
// @namespace    harry.carnet
// @version      4.0
// @description  Suivi des candidatures, lu et réécrit directement dans Carnet de Bord.xlsx sur kDrive
// @author       Harry
// @updateURL    https://raw.githubusercontent.com/Elvalio/carnet-de-bord/main/carnet-de-bord.user.js
// @downloadURL  https://raw.githubusercontent.com/Elvalio/carnet-de-bord/main/carnet-de-bord.user.js
// @match        https://internsearch.com/*
// @match        https://example.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @grant        GM_getResourceText
// @connect      api.infomaniak.com
// @connect      nominatim.openstreetmap.org
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @require      https://unpkg.com/leaflet@1.9.4/dist/leaflet.js
// @resource     leafletcss https://unpkg.com/leaflet@1.9.4/dist/leaflet.css
// @run-at       document-idle
// ==/UserScript==

/*
  Ouvrir  https://internsearch.com  (ou https://example.com/#carnet en secours).
  Token API Infomaniak : scopes drive + user_info. Il reste dans le stockage local de Violentmonkey.
  Le fichier kDrive ne doit pas être ouvert dans OnlyOffice au moment d'enregistrer.
*/

(function () {
  'use strict';
  // Sur le domaine dédié, l'app prend toute la place : taper internsearch.com suffit.
  // Ailleurs (example.com de secours), il faut le fragment #carnet.
  const HOTE_DEDIE = 'internsearch.com';
  const surHoteDedie = location.hostname === HOTE_DEDIE || location.hostname.endsWith('.' + HOTE_DEDIE);
  if (!surHoteDedie && !location.hash.startsWith('#carnet')) return;

  const XHR = typeof GM_xmlhttpRequest === 'function' ? GM_xmlhttpRequest : GM.xmlHttpRequest;

  /* ================= configuration ================= */
  const CFG_KEY = 'carnet.config';
  const CACHE_KEY = 'carnet.cache';
  const THEME_KEY = 'carnet.theme';
  const DEFAULTS = {
    token: '', driveId: '2606717', fileId: '291',
    parentId: '289', fileName: 'Carnet de Bord.xlsx', uploadMode: 'fileid'
  };
  let cfg = Object.assign({}, DEFAULTS, JSON.parse(GM_getValue(CFG_KEY, '{}')));
  const saveCfg = () => GM_setValue(CFG_KEY, JSON.stringify(cfg));
  let theme = GM_getValue(THEME_KEY, 'dark');

  /* ================= modèle ================= */
  const COLS = ['Sujet', 'Premier envoi', 'Etat', 'Nom entreprise', 'Localisation', 'Adresse', 'Poste', 'Source', 'Notes externes', 'Relances'];
  const WIDTHS = [26, 15, 14, 22, 20, 40, 46, 44, 50, 24];
  const GEO = 'Geo';
  const PROSPECTS = 'A postuler';
  const COLS_P = ['Entreprise', 'Poste visé', 'Localisation', 'Adresse', 'Source', 'Repérée le', 'Notes'];
  const WIDTHS_P = [24, 40, 20, 40, 44, 14, 50];
  const TEINTES = {
    'En attente': '#f9e2af', 'Relancé': '#fab387', 'Entretien': '#89dceb',
    'Offre reçue': '#a6e3a1', 'Refus': '#f38ba8', 'Offre retirée': '#6c7086'
  };
  const JOURNAL = 'Journal';
  const VILLES = 'Villes';
  const COLS_J = ['Horodatage', 'Entreprise', 'Champ', 'Avant', 'Après'];
  const WIDTHS_J = [19, 24, 18, 44, 44];
  const MAX_JOURNAL = 600;

  const ETATS = {
    'En attente':    { tone: 'var(--yellow)', actif: true },
    'Relancé':       { tone: 'var(--peach)', actif: true },
    'Entretien':     { tone: 'var(--sky)', actif: true },
    'Offre reçue':   { tone: 'var(--green)', actif: false },
    'Refus':         { tone: 'var(--red)', actif: false },
    'Offre retirée': { tone: 'var(--overlay0)', actif: false }
  };
  const SUJETS = ['Offre publiée', 'Candidature spontanée', 'Candidature spontanée (mail)', 'Poste classique'];
  const CHAMPS = {
    type: 'Sujet', date: 'Premier envoi', etat: 'Etat', entreprise: 'Nom entreprise',
    ville: 'Localisation', adresse: 'Adresse', poste: 'Poste', source: 'Source', notes: 'Notes externes', relances: 'Relances'
  };
  const SEUIL = 12;

  let items = [];
  let reference = [];        // état du fichier au dernier chargement / enregistrement
  let journal = [];
  let villes = [];
  let prospects = [], referenceProspects = [];
  let onglet = 'candidatures';
  let geo = {};              // "adresse ou ville" -> { lat, lon }
  let carte = null, coucheMarqueurs = null, geocodageEnCours = false, repositionner = null;
  let marqueurs = new Map();
  const trierVilles = () => villes.sort((a, b) => a.localeCompare(b, 'fr'));
  let sheetName = 'Candidatures';
  let charge = false, modifie = false;
  let versionDistante = null, conflit = false, controleEnCours = false, dernierControle = 0;
  const INTERVALLE = 5 * 60 * 1000;
  let filtre = 'toutes', recherche = '', editId = null, brouillonRelances = [];

  /* ================= utilitaires ================= */
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const today = () => new Date().toISOString().slice(0, 10);
  const uid = () => 'c' + Math.random().toString(36).slice(2, 9);
  const sansAccents = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();

  function horodatage() {
    const d = new Date();
    const p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }
  function versIso(v) {
    if (!v) return '';
    if (v instanceof Date) return new Date(v.getTime() - v.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    return '';
  }
  const dateFr = iso => iso ? iso.split('-').reverse().join('/') : '—';
  function jours(iso) {
    if (!iso) return null;
    const d = Math.floor((new Date(today()) - new Date(iso)) / 86400000);
    return isNaN(d) ? null : Math.max(0, d);
  }
  const dernierContact = it => (it.relances && it.relances.length) ? it.relances[it.relances.length - 1] : it.date;
  const attente = it => jours(dernierContact(it));
  const estActif = it => !!(ETATS[it.etat] && ETATS[it.etat].actif);
  const aRelancer = it => estActif(it) && it.etat !== 'Entretien' && (attente(it) ?? 0) >= SEUIL;
  function tonAge(j) {
    if (j == null) return 'var(--overlay0)';
    if (j >= 21) return 'var(--red)';
    if (j >= SEUIL) return 'var(--peach)';
    if (j >= 7) return 'var(--yellow)';
    return 'var(--green)';
  }

  /* ================= API kDrive ================= */
  function api(opts) {
    return new Promise((resolve, reject) => {
      XHR({
        method: opts.method || 'GET',
        url: opts.url,
        responseType: opts.responseType,
        headers: Object.assign({ Authorization: 'Bearer ' + cfg.token }, opts.headers || {}),
        data: opts.data,
        timeout: 90000,
        onload: r => resolve(r),
        onerror: () => reject(new Error('Requête réseau échouée')),
        ontimeout: () => reject(new Error('Délai dépassé'))
      });
    });
  }

  async function lireFichier() {
    const r = await api({
      url: `https://api.infomaniak.com/2/drive/${cfg.driveId}/files/${cfg.fileId}/download`,
      responseType: 'arraybuffer'
    });
    if (r.status !== 200) throw new Error(`Téléchargement refusé (HTTP ${r.status})`);
    return r.response;
  }

  async function metaFichier() {
    const r = await api({ url: `https://api.infomaniak.com/2/drive/${cfg.driveId}/files/${cfg.fileId}` });
    if (r.status !== 200) throw new Error('HTTP ' + r.status);
    return JSON.parse(r.responseText)?.data?.last_modified_at ?? null;
  }

  const BASE_UPLOAD = () => `https://api.infomaniak.com/3/drive/${cfg.driveId}/upload`;

  function urlUpload(taille, mode, cible) {
    const id = cible?.fileId ?? cfg.fileId;
    const nom = cible?.fileName ?? cfg.fileName;
    switch (mode || cfg.uploadMode) {
      case 'replace':
        return `${BASE_UPLOAD()}?directory_id=${encodeURIComponent(cfg.parentId)}`
          + `&file_name=${encodeURIComponent(nom)}&total_size=${taille}&conflict=version`;
      case 'nouveau':
        return `${BASE_UPLOAD()}?directory_id=${encodeURIComponent(cfg.parentId)}`
          + `&file_name=${encodeURIComponent(nom)}&total_size=${taille}`;
      default:
        return `${BASE_UPLOAD()}?file_id=${encodeURIComponent(id)}&total_size=${taille}`;
    }
  }

  async function envoyer(buf, mode, cible) {
    const r = await api({
      method: 'POST',
      url: urlUpload(buf.byteLength, mode, cible),
      headers: { 'Content-Type': 'application/octet-stream' },
      data: new Blob([buf], { type: 'application/octet-stream' })
    });
    let corps = null;
    try { corps = JSON.parse(r.responseText); } catch (e) { /* réponse non JSON */ }
    return { status: r.status, corps, brut: r.responseText || '' };
  }

  async function ecrireFichier(buf) {
    const r = await envoyer(buf);
    if (r.status >= 200 && r.status < 300 && (!r.corps || r.corps.result !== 'error')) return r.corps;
    if (r.corps?.error?.code === 'lock_error') {
      throw new Error('Fichier verrouillé : il est ouvert dans l\'éditeur OnlyOffice de kDrive. Fermez l\'onglet, attendez une minute, puis réessayez. Vos modifications sont conservées ici.');
    }
    throw new Error(`Envoi refusé (HTTP ${r.status}) — ${r.brut.slice(0, 900)}`);
  }

  /* ================= xlsx ================= */
  function parser(buf) {
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });

    let nomFeuille = wb.SheetNames.find(n => {
      const ws = wb.Sheets[n];
      const entetes = XLSX.utils.sheet_to_json(ws, { header: 1, range: 0 })[0] || [];
      return entetes.includes('Sujet') || entetes.includes('Nom entreprise');
    }) || wb.SheetNames.find(n => n !== JOURNAL) || wb.SheetNames[0];
    sheetName = nomFeuille;

    const lignes = XLSX.utils.sheet_to_json(wb.Sheets[nomFeuille], { defval: '', raw: false, cellDates: true });
    const liste = lignes
      .filter(l => COLS.some(c => String(l[c] ?? '').trim()))
      .map(l => ({
        id: uid(),
        type: String(l['Sujet'] ?? '').trim(),
        date: versIso(l['Premier envoi']),
        etat: String(l['Etat'] ?? '').trim() || 'En attente',
        entreprise: String(l['Nom entreprise'] ?? '').trim(),
        ville: String(l['Localisation'] ?? l['Ville'] ?? '').trim(),
        adresse: String(l['Adresse'] ?? '').trim(),
        poste: String(l['Poste'] ?? '').trim(),
        source: String(l['Source'] ?? '').trim(),
        notes: String(l['Notes externes'] ?? '').trim(),
        relances: String(l['Relances'] ?? '').split(/[\s,;]+/).map(versIso).filter(Boolean)
      }));

    let hist = [];
    if (wb.SheetNames.includes(JOURNAL)) {
      hist = XLSX.utils.sheet_to_json(wb.Sheets[JOURNAL], { defval: '', raw: false, cellDates: true })
        .map(l => ({
          at: String(l['Horodatage'] ?? '').trim(),
          entreprise: String(l['Entreprise'] ?? '').trim(),
          champ: String(l['Champ'] ?? '').trim(),
          avant: String(l['Avant'] ?? '').trim(),
          apres: String(l['Après'] ?? '').trim()
        }))
        .filter(e => e.at || e.entreprise);
    }
    let listeVilles = [];
    if (wb.SheetNames.includes(VILLES)) {
      listeVilles = XLSX.utils.sheet_to_json(wb.Sheets[VILLES], { defval: '', raw: false })
        .map(l => String(l['Ville'] ?? '').trim()).filter(Boolean);
    }
    listeVilles = Array.from(new Set(listeVilles.concat(liste.map(x => x.ville).filter(Boolean))));

    const listeProspects = wb.SheetNames.includes(PROSPECTS)
      ? XLSX.utils.sheet_to_json(wb.Sheets[PROSPECTS], { defval: '', raw: false, cellDates: true })
          .filter(l => COLS_P.some(c => String(l[c] ?? '').trim()))
          .map(l => ({
            id: uid(),
            entreprise: String(l['Entreprise'] ?? '').trim(),
            poste: String(l['Poste visé'] ?? '').trim(),
            ville: String(l['Localisation'] ?? '').trim(),
            adresse: String(l['Adresse'] ?? '').trim(),
            source: String(l['Source'] ?? '').trim(),
            reperee: versIso(l['Repérée le']),
            notes: String(l['Notes'] ?? '').trim()
          }))
      : [];

    const cacheGeo = {};
    if (wb.SheetNames.includes(GEO)) {
      XLSX.utils.sheet_to_json(wb.Sheets[GEO], { defval: '', raw: false }).forEach(l => {
        const q = String(l['Requête'] ?? l['Adresse'] ?? '').trim();
        const lat = parseFloat(String(l['Latitude'] ?? '').replace(',', '.'));
        const lon = parseFloat(String(l['Longitude'] ?? '').replace(',', '.'));
        if (q && !isNaN(lat) && !isNaN(lon)) cacheGeo[q] = { lat, lon };
      });
    }

    return { liste, hist, listeVilles, cacheGeo, listeProspects };
  }

  function feuilleCandidatures() {
    const aoa = [COLS.slice()];
    items.forEach(it => aoa.push([
      it.type, it.date ? new Date(it.date + 'T00:00:00') : '', it.etat,
      it.entreprise, it.ville, it.adresse, it.poste, it.source, it.notes, (it.relances || []).join(' ')
    ]));
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    ws['!cols'] = WIDTHS.map(w => ({ wch: w }));
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    for (let i = 0; i < items.length; i++) {
      const ref = XLSX.utils.encode_cell({ r: i + 1, c: 1 });
      if (ws[ref] && ws[ref].t === 'd') ws[ref].z = 'dd/mm/yyyy;@';
    }
    return ws;
  }

  function feuilleJournal() {
    const aoa = [COLS_J.slice()];
    journal.slice().reverse().forEach(e => aoa.push([e.at, e.entreprise, e.champ, e.avant, e.apres]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = WIDTHS_J.map(w => ({ wch: w }));
    return ws;
  }

  function feuilleVilles() {
    const ws = XLSX.utils.aoa_to_sheet([['Ville'], ...villes.map(v => [v])]);
    ws['!cols'] = [{ wch: 28 }];
    return ws;
  }

  function feuilleProspects() {
    const aoa = [COLS_P.slice()];
    prospects.forEach(p => aoa.push([
      p.entreprise, p.poste, p.ville, p.adresse, p.source,
      p.reperee ? new Date(p.reperee + 'T00:00:00') : '', p.notes
    ]));
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    ws['!cols'] = WIDTHS_P.map(w => ({ wch: w }));
    for (let i = 0; i < prospects.length; i++) {
      const ref = XLSX.utils.encode_cell({ r: i + 1, c: 5 });
      if (ws[ref] && ws[ref].t === 'd') ws[ref].z = 'dd/mm/yyyy;@';
    }
    return ws;
  }

  function feuilleGeo() {
    const aoa = [['Requête', 'Latitude', 'Longitude']];
    Object.keys(geo).sort().forEach(q => aoa.push([q, geo[q].lat, geo[q].lon]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 46 }, { wch: 12 }, { wch: 12 }];
    return ws;
  }

  function construire() {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, feuilleCandidatures(), sheetName);
    XLSX.utils.book_append_sheet(wb, feuilleProspects(), PROSPECTS);
    XLSX.utils.book_append_sheet(wb, feuilleJournal(), JOURNAL);
    XLSX.utils.book_append_sheet(wb, feuilleVilles(), VILLES);
    XLSX.utils.book_append_sheet(wb, feuilleGeo(), GEO);
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellDates: true });
  }

  /* ================= journal des modifications ================= */
  const copie = liste => liste.map(it => Object.assign({}, it, { relances: (it.relances || []).slice() }));
  const lisible = (champ, v) => {
    if (champ === 'date') return dateFr(v);
    if (champ === 'relances') return (v || []).map(dateFr).join(', ') || '(aucune)';
    return String(v ?? '').trim();
  };

  function diff() {
    const at = horodatage();
    const avant = new Map(reference.map(it => [it.id, it]));
    const apres = new Map(items.map(it => [it.id, it]));
    const entrees = [];

    items.forEach(it => {
      const a = avant.get(it.id);
      if (!a) {
        entrees.push({ at, entreprise: it.entreprise, champ: 'Candidature', avant: '', apres: `ajoutée — ${it.poste || 'poste non précisé'}` });
        return;
      }
      Object.keys(CHAMPS).forEach(k => {
        const x = lisible(k, a[k]), y = lisible(k, it[k]);
        if (x !== y) entrees.push({ at, entreprise: it.entreprise, champ: CHAMPS[k], avant: x, apres: y });
      });
    });
    const avantP = new Map(referenceProspects.map(p => [p.id, p]));
    const apresP = new Map(prospects.map(p => [p.id, p]));
    prospects.forEach(p => {
      if (!avantP.has(p.id)) entrees.push({ at, entreprise: p.entreprise, champ: 'À postuler', avant: '', apres: `ajoutée — ${p.poste || 'poste non précisé'}` });
    });
    referenceProspects.forEach(p => {
      if (!apresP.has(p.id)) entrees.push({ at, entreprise: p.entreprise, champ: 'À postuler', avant: p.poste || 'poste non précisé', apres: 'retirée de la liste' });
    });

    reference.forEach(a => {
      if (!apres.has(a.id)) entrees.push({ at, entreprise: a.entreprise, champ: 'Candidature', avant: `${a.poste || 'poste non précisé'} — ${a.etat}`, apres: 'supprimée' });
    });
    return entrees;
  }

  /* ================= nettoyage des textes ================= */
  function majuscule(s) {
    s = String(s ?? '').replace(/\s+/g, ' ').trim();
    return s ? s[0].toUpperCase() + s.slice(1) : s;
  }
  function normSujet(v) {
    const n = sansAccents(v);
    if (!n) return 'Candidature spontanée';
    if (n.includes('classique')) return 'Poste classique';
    if (n.includes('spontan')) return n.includes('mail') ? 'Candidature spontanée (mail)' : 'Candidature spontanée';
    if (n.includes('offre')) return 'Offre publiée';
    return majuscule(v);
  }
  function normEtat(v) {
    const n = sansAccents(v);
    if (n.includes('retir')) return 'Offre retirée';
    if (n.includes('refus')) return 'Refus';
    if (n.includes('relanc')) return 'Relancé';
    if (n.includes('entretien')) return 'Entretien';
    if (n.includes('offre')) return 'Offre reçue';
    if (n.includes('attente')) return 'En attente';
    return majuscule(v) || 'En attente';
  }
  function normNote(v) {
    let s = majuscule(v);
    if (!s) return '';
    if (!/[.!?…]$/.test(s)) s += '.';
    return s;
  }
  function propositions() {
    return items.map(it => {
      const p = {
        type: normSujet(it.type),
        etat: normEtat(it.etat),
        entreprise: String(it.entreprise ?? '').replace(/\s+/g, ' ').trim(),
        ville: String(it.ville ?? '').replace(/\s+/g, ' ').trim(),
        adresse: String(it.adresse ?? '').replace(/\s+/g, ' ').trim(),
        poste: majuscule(it.poste),
        source: String(it.source ?? '').trim(),
        notes: normNote(it.notes)
      };
      const changes = Object.keys(p).filter(k => p[k] !== (it[k] ?? '')).map(k => ({ champ: CHAMPS[k], k, avant: it[k], apres: p[k] }));
      return { id: it.id, entreprise: it.entreprise, p, changes };
    }).filter(x => x.changes.length);
  }

  /* ================= actions ================= */
  async function charger(silencieux) {
    if (!cfg.token) { ouvrirReglages(); return; }
    majBandeau('Lecture du fichier sur kDrive…');
    try {
      const r = parser(await lireFichier());
      items = r.liste; journal = r.hist;
      villes = r.listeVilles || []; trierVilles();
      geo = r.cacheGeo || {};
      prospects = r.listeProspects || [];
      referenceProspects = copie(prospects);
      reference = copie(items);
      charge = true; modifie = false; conflit = false;
      try { versionDistante = await metaFichier(); dernierControle = Date.now(); } catch (e) { /* sans gravité */ }
      GM_setValue(CACHE_KEY, JSON.stringify({ items, journal, villes, geo, prospects, at: Date.now() }));
      rendre();
      if (!silencieux) toast(`${items.length} candidatures chargées.`);
    } catch (e) {
      charge = false; rendre();
      toast('Lecture impossible : ' + e.message);
      const cache = JSON.parse(GM_getValue(CACHE_KEY, 'null'));
      if (cache && !items.length) {
        items = cache.items; journal = cache.journal || []; villes = cache.villes || []; geo = cache.geo || {}; prospects = cache.prospects || []; reference = []; referenceProspects = [];
        modifie = true; rendre();
        toast('Copie locale restaurée — pensez à renvoyer le fichier.');
      }
    }
  }

  async function enregistrer(force) {
    if (!charge) { toast('Chargez d\'abord le fichier depuis kDrive.'); return; }
    if (force !== true) {
      try {
        const d = await metaFichier();
        dernierControle = Date.now();
        if (d && versionDistante && d > versionDistante) {
          conflit = true; rendre();
          if (!confirm('Le fichier a été modifié sur kDrive depuis votre chargement, sans doute depuis un autre appareil.\n\nEnregistrer maintenant écrasera ces changements. Continuer ?')) return;
        }
      } catch (e) { /* si la vérification échoue, on tente l'envoi quand même */ }
    }
    const entrees = diff();
    journal = journal.concat(entrees).slice(-MAX_JOURNAL);
    majBandeau('Envoi du fichier vers kDrive…');
    try {
      await ecrireFichier(construire());
      reference = copie(items);
      referenceProspects = copie(prospects);
      modifie = false; conflit = false;
      try { versionDistante = await metaFichier(); dernierControle = Date.now(); } catch (e) { /* sans gravité */ }
      rendre();
      toast(entrees.length ? `Enregistré — ${entrees.length} modification${entrees.length > 1 ? 's' : ''} consignée${entrees.length > 1 ? 's' : ''}.` : 'Fichier enregistré sur kDrive.');
    } catch (e) {
      journal = journal.slice(0, journal.length - entrees.length);
      rendre();
      toast(e.message);
    }
  }

  // Contrôle périodique : on ne télécharge rien tant que la date de modification n'a pas bougé.
  async function verifier() {
    if (!charge || !cfg.token || controleEnCours) return;
    controleEnCours = true;
    try {
      const d = await metaFichier();
      dernierControle = Date.now();
      if (d && versionDistante && d > versionDistante) {
        if (modifie) {
          conflit = true; rendre();
          toast('Le fichier a changé sur kDrive depuis un autre appareil.');
        } else {
          await charger(true);
          toast('Carnet actualisé depuis kDrive.');
        }
      } else rendre();
    } catch (e) { /* hors ligne ou API indisponible : on réessaiera */ }
    finally { controleEnCours = false; }
  }

  function toucher() {
    modifie = true;
    GM_setValue(CACHE_KEY, JSON.stringify({ items, journal, villes, geo, prospects, at: Date.now() }));
  }

  /* ================= diagnostic ================= */
  async function diagnostic() {
    const sortie = [];
    const essai = async (nom, opts) => {
      try {
        const r = await api(opts);
        const t = (r.responseText || '').slice(0, 400);
        sortie.push(`${nom}\n  HTTP ${r.status}\n  ${t || '(binaire, ' + (r.response?.byteLength ?? 0) + ' octets)'}`);
      } catch (e) { sortie.push(`${nom}\n  ÉCHEC : ${e.message}`); }
    };
    await essai('1. Profil (scope user_info)', { url: 'https://api.infomaniak.com/2/profile' });
    await essai('2. Métadonnées du fichier', { url: `https://api.infomaniak.com/2/drive/${cfg.driveId}/files/${cfg.fileId}` });
    await essai('3. Téléchargement', { url: `https://api.infomaniak.com/2/drive/${cfg.driveId}/files/${cfg.fileId}/download`, responseType: 'arraybuffer' });
    document.getElementById('diag').textContent = sortie.join('\n\n');
  }

  async function sonder() {
    const diag = document.getElementById('diag');
    const log = [];
    const dire = t => { log.push(t); diag.textContent = log.join('\n'); };
    const charge0 = new TextEncoder().encode('sonde carnet de bord').buffer;
    const nom = `carnet-sonde-${Date.now()}.txt`;
    let idJetable = null;

    dire('Création d\'un fichier jetable…');
    try {
      const r = await envoyer(charge0, 'nouveau', { fileName: nom });
      dire(`  création → HTTP ${r.status}`);
      if (r.status >= 300 || r.corps?.result === 'error') { dire('  ' + r.brut.slice(0, 600)); return; }
      idJetable = r.corps?.data?.id;
      dire(`  fichier jetable id ${idJetable}`);
    } catch (e) { dire('  ÉCHEC : ' + e.message); return; }

    const modes = [['fileid', 'file_id seul'], ['replace', 'directory_id + file_name + conflict=version']];
    const gagnants = [];
    for (const [mode, libelle] of modes) {
      try {
        const r = await envoyer(charge0, mode, { fileId: idJetable, fileName: nom });
        const ok = r.status < 300 && r.corps?.result !== 'error';
        dire(`\n${libelle} → HTTP ${r.status} ${ok ? 'OK' : ''}`);
        if (!ok) dire('  ' + r.brut.slice(0, 500)); else gagnants.push(mode);
      } catch (e) { dire(`\n${libelle} → ÉCHEC : ${e.message}`); }
    }
    if (idJetable) {
      try {
        await api({ method: 'DELETE', url: `https://api.infomaniak.com/2/drive/${cfg.driveId}/files/${idJetable}` });
        dire('\nFichier jetable supprimé.');
      } catch (e) { dire(`\nFichier jetable ${nom} à supprimer à la main.`); }
    }
    if (gagnants.length) {
      cfg.uploadMode = gagnants[0];
      document.getElementById('c-mode').value = gagnants[0];
      saveCfg();
      dire(`\n→ Mode retenu : ${gagnants[0]}.`);
    } else dire('\n→ Aucune variante ne passe.');
  }

  /* ================= styles ================= */
  const CSS = `
  :root[data-theme="dark"]{--base:#1e1e2e;--mantle:#181825;--crust:#11111b;--surface0:#313244;--surface1:#45475a;
  --surface2:#585b70;--overlay0:#6c7086;--overlay1:#7f849c;--text:#cdd6f4;--subtext0:#a6adc8;--blue:#89b4fa;
  --lavender:#b4befe;--mauve:#cba6f7;--red:#f38ba8;--peach:#fab387;--yellow:#f9e2af;--green:#a6e3a1;--sky:#89dceb;
  --ombre:rgba(0,0,0,.45);--voile:rgba(17,17,27,.75);}
  :root[data-theme="light"]{--base:#eff1f5;--mantle:#e6e9ef;--crust:#dce0e8;--surface0:#ccd0da;--surface1:#bcc0cc;
  --surface2:#acb0be;--overlay0:#9ca0b0;--overlay1:#8c8fa1;--text:#4c4f69;--subtext0:#6c6f85;--blue:#1e66f5;
  --lavender:#7287fd;--mauve:#8839ef;--red:#d20f39;--peach:#fe640b;--yellow:#df8e1d;--green:#40a02b;--sky:#04a5e5;
  --ombre:rgba(76,79,105,.16);--voile:rgba(220,224,232,.8);}
  :root{--mono:ui-monospace,"JetBrains Mono",Menlo,Consolas,monospace;--sans:ui-sans-serif,system-ui,"Segoe UI",Roboto,sans-serif;}
  html,body{margin:0!important;padding:0!important;background:var(--base)!important;}
  body{font-family:var(--sans);color:var(--text);font-size:15px;line-height:1.5;-webkit-font-smoothing:antialiased;}
  *{box-sizing:border-box;}
  .cb-app{max-width:1080px;margin:0 auto;padding:22px 16px 110px;}
  .cb-head{display:flex;flex-wrap:wrap;gap:16px;align-items:flex-end;justify-content:space-between;margin-bottom:16px;}
  .cb-head h1{font-size:25px;margin:0;font-weight:650;letter-spacing:-.02em;}
  .cb-head h1 span{color:var(--mauve);}
  .cb-file{margin:4px 0 0;color:var(--overlay1);font-family:var(--mono);font-size:11.5px;}
  .cb-file b{color:var(--peach);font-weight:500;}
  .cb-stats{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
  .cb-stat{background:var(--mantle);border:1px solid var(--surface0);border-radius:10px;padding:7px 12px;min-width:70px;}
  .cb-stat b{display:block;font-family:var(--mono);font-size:18px;line-height:1.15;font-weight:600;}
  .cb-stat small{color:var(--overlay0);font-size:10px;text-transform:uppercase;letter-spacing:.08em;font-family:var(--mono);}
  .cb-outils-head{display:flex;gap:6px;flex:0 0 auto;}
  .cb-theme{background:var(--mantle);border:1px solid var(--surface0);color:var(--subtext0);border-radius:10px;
  padding:9px 11px;cursor:pointer;font-size:15px;line-height:1;}
  .cb-theme:hover{color:var(--text);border-color:var(--surface2);}
  .cb-sync{display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:var(--mantle);border:1px solid var(--surface0);
  border-radius:10px;padding:10px 12px;margin-bottom:14px;}
  .cb-sync span{font-family:var(--mono);font-size:11.5px;color:var(--overlay1);margin-right:auto;}
  .cb-radar{border:1px solid var(--surface0);border-left:3px solid var(--peach);background:var(--mantle);border-radius:10px;padding:11px 15px;margin-bottom:16px;}
  .cb-radar.quiet{border-left-color:var(--surface2);}
  .cb-radar-t{font-family:var(--mono);font-size:10.5px;text-transform:uppercase;letter-spacing:.09em;color:var(--overlay1);margin-bottom:5px;}
  .cb-radar-b{font-size:13.5px;color:var(--subtext0);}
  .cb-radar-b button{background:none;border:none;color:var(--peach);font:inherit;padding:0;cursor:pointer;text-decoration:underline;text-underline-offset:3px;}
  .cb-onglets{display:flex;gap:4px;background:var(--mantle);border:1px solid var(--surface0);
  border-radius:11px;padding:4px;margin-bottom:14px;}
  .cb-onglet{flex:1;background:none;border:none;border-radius:8px;padding:9px 12px;cursor:pointer;
  color:var(--overlay1);font:inherit;font-size:14px;font-weight:560;transition:background .12s,color .12s;}
  .cb-onglet:hover{color:var(--text);}
  .cb-onglet[aria-selected="true"]{background:var(--surface0);color:var(--text);}
  .cb-onglet i{font-style:normal;font-family:var(--mono);font-size:11px;color:var(--overlay0);margin-left:7px;}
  .cb-onglet[aria-selected="true"] i{color:var(--mauve);}
  .cb-prospect{border-left-color:var(--lavender)!important;}
  .cb-tools{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:14px;}
  .cb-search{flex:1 1 200px;}
  .cb-search input{width:100%;background:var(--mantle);border:1px solid var(--surface0);color:var(--text);border-radius:10px;padding:9px 12px;font:inherit;font-size:14px;}
  .cb-chips{display:flex;gap:6px;flex-wrap:wrap;}
  .cb-chip{background:var(--mantle);border:1px solid var(--surface0);color:var(--subtext0);border-radius:999px;padding:6px 12px;font-family:var(--mono);font-size:12px;cursor:pointer;}
  .cb-chip:hover{border-color:var(--surface2);color:var(--text);}
  .cb-chip[aria-pressed="true"]{background:var(--surface0);color:var(--text);border-color:var(--surface2);}
  .cb-chip i{font-style:normal;color:var(--overlay0);margin-left:5px;}
  .cb-btn{background:var(--surface0);border:1px solid var(--surface1);color:var(--text);border-radius:10px;padding:8px 14px;font:inherit;font-size:14px;cursor:pointer;}
  .cb-btn:hover{background:var(--surface1);}
  .cb-btn.p{background:var(--mauve);border-color:var(--mauve);color:var(--base);font-weight:600;}
  .cb-btn.p:hover{background:var(--lavender);border-color:var(--lavender);}
  .cb-btn.g{background:none;border-color:var(--surface0);color:var(--subtext0);}
  .cb-btn.g:hover{background:var(--mantle);color:var(--text);}
  .cb-btn.d{color:var(--red);border-color:var(--surface0);background:none;}
  .cb-btn[disabled]{opacity:.45;cursor:not-allowed;}
  .cb-list{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:8px;}
  .cb-card{background:var(--mantle);border:1px solid var(--surface0);border-radius:10px;padding:13px 15px;display:grid;
  grid-template-columns:1fr auto;gap:4px 16px;position:relative;overflow:hidden;}
  .cb-card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--tone,var(--surface2));}
  .cb-t{display:flex;align-items:baseline;gap:9px;flex-wrap:wrap;}
  .cb-t b{font-size:15.5px;font-weight:620;}
  .cb-t a{color:var(--blue);text-decoration:none;font-family:var(--mono);font-size:11px;}
  .cb-lieu{font-family:var(--mono);font-size:11.5px;color:var(--sky);border:1px solid var(--sky);
  border-radius:5px;padding:2px 8px;white-space:nowrap;}
  .cb-lieu::before{content:"◎ ";opacity:.7;}
  .cb-poste{color:var(--subtext0);font-size:14px;}
  .cb-notes{color:var(--overlay1);font-size:13px;margin-top:6px;border-left:2px solid var(--surface0);padding-left:9px;}
  .cb-meta{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:8px;font-family:var(--mono);font-size:11px;color:var(--overlay0);}
  .cb-tag{border:1px solid var(--surface0);border-radius:5px;padding:2px 7px;}
  .cb-etat{color:var(--tone);border-color:var(--tone);background:transparent;}
  .cb-side{display:flex;flex-direction:column;align-items:flex-end;gap:7px;}
  .cb-age b{display:block;font-family:var(--mono);font-size:19px;font-weight:600;line-height:1;color:var(--age,var(--subtext0));text-align:right;}
  .cb-age small{font-family:var(--mono);font-size:9.5px;text-transform:uppercase;letter-spacing:.07em;color:var(--overlay0);}
  .cb-gauge{width:90px;height:4px;border-radius:3px;background:var(--surface0);overflow:hidden;}
  .cb-gauge span{display:block;height:100%;background:var(--age,var(--subtext0));}
  .cb-ico{background:none;border:1px solid transparent;color:var(--overlay0);border-radius:7px;padding:4px 7px;cursor:pointer;font-family:var(--mono);font-size:11px;}
  .cb-ico:hover{color:var(--text);border-color:var(--surface1);background:var(--surface0);}
  .cb-empty{text-align:center;padding:52px 20px;border:1px dashed var(--surface1);border-radius:10px;color:var(--overlay1);}
  .cb-empty b{display:block;color:var(--text);font-size:16px;margin-bottom:6px;}
  .cb-bar{position:fixed;left:50%;bottom:16px;transform:translate(-50%,140%);display:flex;align-items:center;gap:12px;
  background:var(--mantle);border:1px solid var(--peach);border-radius:14px;padding:11px 14px;box-shadow:0 8px 28px var(--ombre);
  z-index:9998;max-width:calc(100vw - 24px);transition:transform .22s ease;}
  .cb-bar.on{transform:translate(-50%,0);}
  .cb-bar.conflit{border-color:var(--red);}
  .cb-bar p{margin:0;font-size:13.5px;color:var(--subtext0);}
  .cb-bar p b{display:block;color:var(--text);font-size:14px;font-weight:620;}
  .cb-ov{position:fixed;inset:0;background:var(--voile);display:none;align-items:flex-start;justify-content:center;padding:20px 14px;overflow-y:auto;z-index:99999;}
  .cb-ov.open{display:flex;}
  .cb-sheet{background:var(--base);border:1px solid var(--surface1);border-radius:14px;width:100%;max-width:560px;padding:20px;}
  .cb-sheet.large{max-width:760px;}
  .cb-sheet h2{margin:0 0 14px;font-size:17px;font-weight:620;}
  .cb-grid{display:grid;grid-template-columns:1fr 1fr;gap:11px;}
  .cb-f{display:flex;flex-direction:column;gap:5px;}
  .cb-f.w{grid-column:1/-1;}
  .cb-f label{font-family:var(--mono);font-size:10.5px;text-transform:uppercase;letter-spacing:.07em;color:var(--overlay1);}
  .cb-f input,.cb-f select,.cb-f textarea{background:var(--mantle);border:1px solid var(--surface0);color:var(--text);border-radius:8px;padding:9px 11px;font:inherit;font-size:14px;}
  .cb-f textarea{resize:vertical;min-height:60px;}
  .cb-sf{display:flex;gap:8px;margin-top:16px;justify-content:flex-end;flex-wrap:wrap;}
  .cb-sf .cb-btn.d{margin-right:auto;}
  .cb-rel{font-family:var(--mono);font-size:12px;color:var(--overlay1);}
  #map{height:58vh;min-height:300px;border-radius:10px;border:1px solid var(--surface0);background:var(--mantle);}
  #map .leaflet-popup-content-wrapper,#map .leaflet-popup-tip{background:var(--mantle);color:var(--text);}
  #map .leaflet-popup-content{font-family:var(--sans);font-size:13px;line-height:1.45;}
  #map .leaflet-popup-content a{color:var(--blue);}
  .cb-mini{background:var(--surface0);border:1px solid var(--surface1);color:var(--text);border-radius:6px;
  padding:3px 8px;margin:6px 6px 0 0;font-family:var(--mono);font-size:10.5px;cursor:pointer;}
  .cb-mini:hover{background:var(--surface1);}
  #map .leaflet-control-attribution{background:var(--mantle);color:var(--overlay0);}
  #map .leaflet-control-attribution a{color:var(--overlay1);}
  #map .cb-tip{background:var(--crust);color:var(--text);border:1px solid var(--surface1);border-radius:6px;
  font-family:var(--mono);font-size:11px;padding:2px 7px;box-shadow:0 2px 8px var(--ombre);white-space:nowrap;}
  #map .cb-tip::before{border-top-color:var(--surface1);}
  .cb-sheet.xlarge{max-width:1020px;}
  .cb-zone{display:grid;grid-template-columns:1fr 290px;gap:12px;align-items:stretch;}
  .cb-panneau{height:58vh;min-height:300px;overflow-y:auto;border:1px solid var(--surface0);border-radius:10px;
  background:var(--mantle);padding:6px;display:flex;flex-direction:column;gap:4px;}
  .cb-groupe{font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--overlay0);padding:8px 8px 3px;}
  .cb-pin{display:flex;align-items:center;gap:6px;border:1px solid transparent;border-radius:8px;padding:2px 4px;}
  .cb-pin:hover{background:var(--surface0);border-color:var(--surface1);}
  .cb-pin-main{display:flex;align-items:center;gap:9px;flex:1;min-width:0;text-align:left;background:none;
  border:none;padding:5px 4px;cursor:pointer;color:var(--text);font:inherit;}
  .cb-pin.absente{opacity:.85;}
  .cb-pin i{flex:0 0 auto;width:11px;height:11px;border-radius:50%;box-shadow:0 0 0 1px var(--surface1);}
  .cb-pin i.vide{background:transparent!important;border:1px dashed var(--overlay0);box-shadow:none;}
  .cb-pin span{min-width:0;flex:1;}
  .cb-pin b{display:block;font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  .cb-pin small{display:block;font-family:var(--mono);font-size:10.5px;color:var(--overlay1);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  .cb-pin em{font-style:normal;font-family:var(--mono);font-size:10px;color:var(--peach);white-space:nowrap;}
  .cb-grappe-wrap{background:none;border:none;overflow:visible;}
  .cb-detail{position:absolute;left:100%;top:50%;transform:translateY(-50%);margin-left:5px;
  display:flex;flex-direction:column;gap:2px;pointer-events:none;}
  .cb-detail b{font-family:var(--mono);font-size:9.5px;font-weight:700;line-height:1;color:#11111b;
  background:var(--c);border:1px solid var(--base);border-radius:4px;padding:2px 4px;white-space:nowrap;}
  .cb-grappe{display:flex;align-items:center;justify-content:center;border-radius:50%;
  background:var(--c);color:#11111b;font-family:var(--mono);font-size:14px;font-weight:700;
  border:3px solid var(--base);box-shadow:0 0 0 2px var(--c),0 3px 10px var(--ombre);cursor:pointer;}
  .cb-grappe:hover{transform:scale(1.08);}
  .cb-legende{display:flex;flex-wrap:wrap;gap:10px 16px;margin-top:10px;font-family:var(--mono);font-size:11px;color:var(--overlay1);}
  .cb-legende i{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:5px;vertical-align:-1px;
  border:2px solid var(--base);box-shadow:0 0 0 1px var(--surface1);}
  .cb-case{display:flex;align-items:center;gap:7px;margin-right:auto;font-size:13.5px;color:var(--subtext0);cursor:pointer;}
  .cb-case input{accent-color:var(--mauve);}
  .cb-hist{max-height:60vh;overflow:auto;display:flex;flex-direction:column;gap:6px;}
  .cb-jour{font-family:var(--mono);font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--overlay0);
  border-bottom:1px solid var(--surface0);padding:10px 0 5px;margin-top:6px;}
  .cb-ent{background:var(--mantle);border:1px solid var(--surface0);border-radius:9px;padding:9px 12px;font-size:13.5px;}
  .cb-ent b{font-weight:620;}
  .cb-ent i{font-style:normal;color:var(--overlay1);font-family:var(--mono);font-size:11px;}
  .cb-ent .av{color:var(--overlay1);text-decoration:line-through;}
  .cb-ent .ap{color:var(--green);}
  #diag{font-family:var(--mono);font-size:11px;white-space:pre-wrap;word-break:break-all;color:var(--subtext0);
  background:var(--crust);border:1px solid var(--surface0);border-radius:8px;padding:10px;max-height:220px;overflow:auto;margin-top:10px;}
  .cb-toast{position:fixed;bottom:80px;left:50%;transform:translate(-50%,12px);background:var(--surface0);border:1px solid var(--surface1);
  color:var(--text);padding:9px 16px;border-radius:999px;font-size:13.5px;opacity:0;pointer-events:none;transition:opacity .18s,transform .18s;z-index:99999;max-width:90vw;text-align:center;}
  .cb-toast.show{opacity:1;transform:translate(-50%,0);}

  /* Coquille mobile : barre de navigation, bouton flottant, menu « Plus ».
     Définie ici mais masquée — seul le @media sous 620px l'active. Le desktop ne la voit jamais. */
  .cb-nav,.cb-fab,.cb-menu-ov{display:none;}
  .cb-nav{position:fixed;left:0;right:0;bottom:0;z-index:9997;background:var(--mantle);border-top:1px solid var(--surface0);
  grid-template-columns:repeat(4,1fr);padding-bottom:env(safe-area-inset-bottom);box-shadow:0 -4px 20px var(--ombre);}
  .cb-nav-i{position:relative;background:none;border:none;color:var(--overlay1);font:inherit;font-size:10.5px;font-weight:560;
  display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;padding:9px 4px 8px;min-height:58px;cursor:pointer;}
  .cb-nav-i svg{width:22px;height:22px;stroke:currentColor;fill:none;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round;}
  .cb-nav-i[aria-selected="true"]{color:var(--mauve);}
  .cb-nav-badge{font-style:normal;position:absolute;top:6px;left:calc(50% + 7px);min-width:16px;height:16px;padding:0 4px;
  background:var(--surface1);color:var(--text);border-radius:8px;font-family:var(--mono);font-size:9.5px;line-height:16px;text-align:center;}
  .cb-nav-badge:empty{display:none;}
  .cb-nav-i[aria-selected="true"] .cb-nav-badge{background:var(--mauve);color:var(--base);}
  .cb-fab{position:fixed;right:15px;z-index:9997;width:56px;height:56px;border-radius:50%;border:none;
  background:var(--mauve);color:var(--base);font-size:30px;line-height:0;cursor:pointer;box-shadow:0 6px 20px var(--ombre);
  align-items:center;justify-content:center;padding:0 0 3px;}
  .cb-fab:active{transform:scale(.93);}
  .cb-menu-ov{position:fixed;inset:0;background:var(--voile);z-index:99998;align-items:flex-end;justify-content:center;}
  .cb-menu-ov.open{display:flex;}
  .cb-menu{background:var(--base);border:1px solid var(--surface1);border-bottom:none;border-radius:16px 16px 0 0;
  width:100%;max-width:560px;padding:6px 8px calc(10px + env(safe-area-inset-bottom));display:flex;flex-direction:column;gap:1px;
  animation:cb-monte .2s ease;}
  @keyframes cb-monte{from{transform:translateY(30px);opacity:.4;}to{transform:translateY(0);opacity:1;}}
  .cb-menu-h{font-family:var(--mono);font-size:10.5px;text-transform:uppercase;letter-spacing:.09em;color:var(--overlay0);
  padding:12px 12px 9px;text-align:center;}
  .cb-menu-i{display:flex;align-items:center;gap:13px;background:none;border:none;color:var(--text);font:inherit;font-size:15px;
  text-align:left;padding:14px 13px;border-radius:11px;cursor:pointer;min-height:52px;width:100%;}
  .cb-menu-i:active{background:var(--surface0);}
  .cb-menu-i span{font-size:17px;width:22px;text-align:center;flex:0 0 auto;color:var(--subtext0);}
  a:focus-visible,button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible{outline:2px solid var(--lavender);outline-offset:1px;}
  @media(max-width:620px){
  .cb-app{padding:14px 12px calc(112px + env(safe-area-inset-bottom));}

  /* Coquille native : la barre du bas et le bouton flottant remplacent
     la barre d'actions, les onglets du haut et le bouton « + » en ligne. */
  .cb-nav{display:grid;}
  .cb-fab{display:flex;bottom:calc(70px + env(safe-area-inset-bottom));}
  .cb-onglets,.cb-sync,.cb-outils-head,#btn-ajouter{display:none!important;}

  /* En-tête : titre seul ; les outils sont passés dans la barre du bas et le menu « Plus ». */
  .cb-head{gap:10px;margin-bottom:12px;}
  .cb-head h1{font-size:21px;}
  .cb-file{font-size:10.5px;}
  .cb-theme{padding:8px 10px;min-height:40px;}
  .cb-tech{display:none;}
  *{-webkit-tap-highlight-color:transparent;}

  /* Compteurs : six cellules sur deux rangées, aucune hors écran. */
  .cb-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:5px;width:100%;}
  .cb-stat{min-width:0;padding:6px 2px;text-align:center;}
  .cb-stat b{font-size:16px;}
  .cb-stat small{font-size:8px;letter-spacing:.01em;display:block;overflow:hidden;text-overflow:ellipsis;}

  /* Filtres : bande défilante horizontale, indépendante du reste. */
  .cb-chips{flex-wrap:nowrap;overflow-x:auto;width:100%;padding-bottom:2px;scrollbar-width:none;}
  .cb-chips::-webkit-scrollbar{display:none;}
  .cb-chip{flex:0 0 auto;min-height:36px;}
  .cb-tools{gap:9px;}
  .cb-search{flex:1 1 100%;}
  #btn-ajouter{width:100%;min-height:46px;font-size:15px;}

  /* Cibles tactiles : rien en dessous de 44 px de haut. */
  .cb-search input,.cb-f input,.cb-f select,.cb-f textarea{font-size:16px;min-height:44px;}
  .cb-f textarea{min-height:80px;}
  .cb-btn{min-height:42px;}
  .cb-ico{min-height:38px;padding:8px 11px;font-size:12px;}
  .cb-pin-main{padding:9px 4px;}
  .cb-mini{min-height:34px;padding:6px 10px;}

  .cb-card{grid-template-columns:1fr;padding:12px 13px;}
  .cb-side{flex-direction:row;align-items:center;justify-content:space-between;width:100%;margin-top:10px;
  border-top:1px solid var(--surface0);padding-top:9px;}
  .cb-age{display:flex;align-items:baseline;gap:5px;flex:0 0 auto;}
  .cb-age b{font-size:16px;white-space:nowrap;}
  .cb-age small{white-space:nowrap;}
  .cb-gauge{flex:1 1 auto;width:auto;min-width:24px;max-width:90px;}
  .cb-side>div:last-child{flex:0 0 auto;display:flex;gap:6px;flex-wrap:nowrap;}
  .cb-t b{font-size:15px;}

  /* Fenêtres : feuilles pleine hauteur, actions collées en bas et atteignables au pouce. */
  .cb-ov{padding:0;align-items:stretch;}
  .cb-sheet,.cb-sheet.large,.cb-sheet.xlarge{max-width:none;border-radius:0;border:none;min-height:100%;
  padding:16px 14px calc(16px + env(safe-area-inset-bottom));display:flex;flex-direction:column;}
  .cb-sheet h2{font-size:16px;position:sticky;top:0;background:var(--base);padding:4px 0 10px;margin:0 0 10px;z-index:2;}
  .cb-grid{grid-template-columns:1fr;flex:1;align-content:start;}
  .cb-sf{position:sticky;bottom:0;background:var(--base);border-top:1px solid var(--surface0);
  padding:10px 0 0;margin-top:14px;gap:7px;}
  .cb-sf .cb-btn{flex:1;min-width:0;}
  .cb-sf .cb-btn.d{flex:0 0 100%;margin-right:0;order:5;}
  .cb-hist{max-height:none;}

  /* Carte : empilée sous le panneau, avec assez de hauteur pour rester manipulable. */
  .cb-zone{grid-template-columns:1fr;}
  #map{height:46vh;min-height:260px;}
  .cb-panneau{height:auto;max-height:32vh;min-height:0;}
  .cb-legende{font-size:10px;gap:6px 12px;}
  .cb-case{flex:0 0 100%;margin:0 0 4px;order:-1;min-height:40px;}

  .cb-radar{padding:10px 13px;margin-bottom:13px;}
  .cb-radar-b{font-size:13px;}
  .cb-notes{font-size:12.5px;}
  .cb-meta{font-size:10.5px;gap:6px;}
  .cb-list{gap:7px;}
  .cb-empty{padding:40px 16px;}
  .cb-ov{overscroll-behavior:contain;}
  .cb-lieu{font-size:11px;padding:2px 6px;}

  /* Barre d'enregistrement : remontée au-dessus de la barre de navigation. */
  .cb-bar{left:10px;right:10px;bottom:calc(74px + env(safe-area-inset-bottom));transform:translate(0,180%);max-width:none;}
  .cb-bar.on{transform:translate(0,0);}
  .cb-bar p b{font-size:13.5px;}
  .cb-bar p{font-size:12px;}
  .cb-bar .cb-btn{flex:0 0 auto;}
  .cb-toast{left:10px;right:10px;bottom:calc(132px + env(safe-area-inset-bottom));transform:translate(0,12px);max-width:none;}
  .cb-toast.show{transform:translate(0,0);}
  }
  @media(prefers-reduced-motion:reduce){*{transition:none!important;}}`;

  /* ================= structure ================= */
  const HTML = `
  <div class="cb-app">
    <div class="cb-head">
      <div>
        <h1>Carnet de <span>bord</span></h1>
        <p class="cb-file" id="etat-fichier"></p>
      </div>
      <div class="cb-stats" id="stats"></div>
      <div class="cb-outils-head">
        <button class="cb-theme" id="btn-recharger" title="Recharger la page" aria-label="Recharger la page">⟳</button>
        <button class="cb-theme" id="btn-theme" title="Changer de thème" aria-label="Changer de thème"></button>
      </div>
    </div>
    <div class="cb-sync">
      <span id="bandeau"></span>
      <button class="cb-btn g" id="btn-charger">Charger</button>
      <button class="cb-btn g" id="btn-carte">Carte</button>
      <button class="cb-btn g" id="btn-historique">Historique</button>
      <button class="cb-btn g" id="btn-nettoyer">Nettoyer les textes</button>
      <button class="cb-btn g" id="btn-reglages">Réglages</button>
      <button class="cb-btn p" id="btn-enregistrer">Enregistrer</button>
    </div>
    <div class="cb-onglets" role="tablist">
      <button class="cb-onglet" data-onglet="candidatures" aria-selected="true">Candidatures<i id="n-cand"></i></button>
      <button class="cb-onglet" data-onglet="prospects" aria-selected="false">À postuler<i id="n-prosp"></i></button>
    </div>
    <div class="cb-radar" id="radar"></div>
    <div class="cb-tools">
      <div class="cb-search"><input id="q" type="search" placeholder="Chercher une entreprise, un poste, une note…"></div>
      <div class="cb-chips" id="chips"></div>
      <button class="cb-btn p" id="btn-ajouter">+ Candidature</button>
    </div>
    <ul class="cb-list" id="list"></ul>
  </div>

  <div class="cb-bar" id="bar">
    <p><b id="bar-titre"></b><span id="bar-texte"></span></p>
    <button class="cb-btn g" id="bar-recharger">Recharger</button>
    <button class="cb-btn p" id="bar-save">Enregistrer sur kDrive</button>
  </div>

  <div class="cb-ov" id="ov-fiche"><div class="cb-sheet" role="dialog" aria-modal="true">
    <h2 id="fiche-titre">Nouvelle candidature</h2>
    <div class="cb-grid">
      <div class="cb-f w"><label for="f-entreprise">Nom entreprise</label><input id="f-entreprise"></div>
      <div class="cb-f w"><label for="f-poste">Poste</label><input id="f-poste"></div>
      <div class="cb-f w"><label for="f-adresse">Adresse</label><input id="f-adresse" placeholder="21 rue des Glénan, 35850 Gévezé, France"></div>
      <div class="cb-f w"><label for="f-ville">Localisation</label>
        <select id="f-ville"></select>
        <input id="f-ville-new" placeholder="Nom de la nouvelle ville" style="display:none">
      </div>
      <div class="cb-f"><label for="f-type">Sujet</label><input id="f-type" list="l-types"><datalist id="l-types"></datalist></div>
      <div class="cb-f"><label for="f-etat">Etat</label><select id="f-etat"></select></div>
      <div class="cb-f"><label for="f-date">Premier envoi</label><input id="f-date" type="date"></div>
      <div class="cb-f"><label for="f-relance-add">Ajouter une relance</label><input id="f-relance-add" type="date"></div>
      <div class="cb-f w"><label for="f-source">Source</label><input id="f-source" placeholder="https://"></div>
      <div class="cb-f w"><label for="f-notes">Notes externes</label><textarea id="f-notes"></textarea></div>
      <div class="cb-f w"><label>Relances enregistrées</label><div class="cb-rel" id="f-relances"></div></div>
    </div>
    <div class="cb-sf">
      <button class="cb-btn d" id="f-suppr">Supprimer</button>
      <button class="cb-btn g" id="f-annuler">Annuler</button>
      <button class="cb-btn p" id="f-valider">Enregistrer</button>
    </div>
  </div></div>

  <div class="cb-ov" id="ov-prospect"><div class="cb-sheet" role="dialog" aria-modal="true">
    <h2 id="prospect-titre">Entreprise à postuler</h2>
    <div class="cb-grid">
      <div class="cb-f w"><label for="p-entreprise">Entreprise</label><input id="p-entreprise"></div>
      <div class="cb-f w"><label for="p-poste">Poste visé</label><input id="p-poste" placeholder="Développeur fullstack"></div>
      <div class="cb-f w"><label for="p-ville">Localisation</label>
        <select id="p-ville"></select>
        <input id="p-ville-new" placeholder="Nom de la nouvelle ville" style="display:none">
      </div>
      <div class="cb-f w"><label for="p-adresse">Adresse</label><input id="p-adresse" placeholder="21 rue des Glénan, 35850 Gévezé, France"></div>
      <div class="cb-f w"><label for="p-source">Source</label><input id="p-source" placeholder="https://"></div>
      <div class="cb-f"><label for="p-reperee">Repérée le</label><input id="p-reperee" type="date"></div>
      <div class="cb-f w"><label for="p-notes">Notes</label><textarea id="p-notes" placeholder="Pourquoi elle t'intéresse, qui contacter, ce qu'il reste à préparer…"></textarea></div>
    </div>
    <div class="cb-sf">
      <button class="cb-btn d" id="p-suppr">Supprimer</button>
      <button class="cb-btn g" id="p-annuler">Annuler</button>
      <button class="cb-btn p" id="p-valider">Enregistrer</button>
    </div>
  </div></div>

  <div class="cb-ov" id="ov-carte"><div class="cb-sheet xlarge" role="dialog" aria-modal="true">
    <h2>Carte des candidatures</h2>
    <div class="cb-zone">
      <div id="map"></div>
      <div class="cb-panneau" id="m-liste"></div>
    </div>
    <div class="cb-legende" id="m-legende"></div>
    <p class="cb-rel" id="m-info" style="margin:10px 0 0"></p>
    <div class="cb-sf">
      <label class="cb-case"><input type="checkbox" id="m-noms"> Afficher les noms</label>
      <button class="cb-btn g" id="m-geo">Localiser les adresses</button>
      <button class="cb-btn g" id="m-fermer">Fermer</button>
    </div>
  </div></div>

  <div class="cb-ov" id="ov-hist"><div class="cb-sheet large" role="dialog" aria-modal="true">
    <h2>Historique des modifications</h2>
    <div class="cb-hist" id="hist"></div>
    <div class="cb-sf"><button class="cb-btn g" id="h-fermer">Fermer</button></div>
  </div></div>

  <div class="cb-ov" id="ov-net"><div class="cb-sheet large" role="dialog" aria-modal="true">
    <h2>Nettoyage des textes</h2>
    <div class="cb-hist" id="net"></div>
    <div class="cb-sf">
      <button class="cb-btn g" id="n-annuler">Annuler</button>
      <button class="cb-btn p" id="n-appliquer">Appliquer</button>
    </div>
  </div></div>

  <div class="cb-ov" id="ov-reglages"><div class="cb-sheet" role="dialog" aria-modal="true">
    <h2>Réglages kDrive</h2>
    <div class="cb-grid">
      <div class="cb-f w"><label for="c-token">Token API (scopes drive + user_info)</label><input id="c-token" type="password" placeholder="collez le token"></div>
      <div class="cb-f"><label for="c-drive">drive_id</label><input id="c-drive"></div>
      <div class="cb-f"><label for="c-file">file_id</label><input id="c-file"></div>
      <div class="cb-f"><label for="c-parent">directory_id (dossier parent)</label><input id="c-parent"></div>
      <div class="cb-f"><label for="c-nom">Nom du fichier</label><input id="c-nom"></div>
      <div class="cb-f w"><label for="c-mode">Mode d'envoi</label><select id="c-mode">
        <option value="fileid">file_id seul (recommandé)</option>
        <option value="replace">directory_id + file_name + conflict=version</option>
      </select></div>
      <div class="cb-f w"><div id="diag">Le diagnostic affichera ici la réponse brute de l'API.</div></div>
    </div>
    <div class="cb-sf">
      <button class="cb-btn g" id="c-diag">Diagnostic</button>
      <button class="cb-btn g" id="c-sonde">Tester l'envoi</button>
      <button class="cb-btn g" id="c-annuler">Fermer</button>
      <button class="cb-btn p" id="c-valider">Enregistrer et charger</button>
    </div>
  </div></div>

  <nav class="cb-nav" id="cb-nav">
    <button class="cb-nav-i" data-onglet="candidatures" aria-selected="true">
      <svg viewBox="0 0 24 24"><path d="M8 6h12M8 12h12M8 18h12M4 6h.01M4 12h.01M4 18h.01"/></svg>
      <span>Candidatures</span><i class="cb-nav-badge" id="nb-cand"></i>
    </button>
    <button class="cb-nav-i" data-onglet="prospects" aria-selected="false">
      <svg viewBox="0 0 24 24"><path d="M6 3h12a1 1 0 0 1 1 1v17l-7-4-7 4V4a1 1 0 0 1 1-1Z"/></svg>
      <span>À postuler</span><i class="cb-nav-badge" id="nb-prosp"></i>
    </button>
    <button class="cb-nav-i" id="nav-carte">
      <svg viewBox="0 0 24 24"><path d="M12 21s7-6.3 7-11a7 7 0 1 0-14 0c0 4.7 7 11 7 11Z"/><circle cx="12" cy="10" r="2.6"/></svg>
      <span>Carte</span>
    </button>
    <button class="cb-nav-i" id="nav-plus">
      <svg viewBox="0 0 24 24"><circle cx="5" cy="12" r="1.4"/><circle cx="12" cy="12" r="1.4"/><circle cx="19" cy="12" r="1.4"/></svg>
      <span>Plus</span>
    </button>
  </nav>

  <button class="cb-fab" id="cb-fab" aria-label="Ajouter">+</button>

  <div class="cb-menu-ov" id="cb-menu-ov">
    <div class="cb-menu" role="menu">
      <div class="cb-menu-h">Actions</div>
      <button class="cb-menu-i" data-act="save" role="menuitem"><span>⭱</span>Enregistrer sur kDrive</button>
      <button class="cb-menu-i" data-act="charger" role="menuitem"><span>↻</span>Charger depuis kDrive</button>
      <button class="cb-menu-i" data-act="historique" role="menuitem"><span>≡</span>Historique des modifications</button>
      <button class="cb-menu-i" data-act="nettoyer" role="menuitem"><span>✦</span>Nettoyer les textes</button>
      <button class="cb-menu-i" data-act="reglages" role="menuitem"><span>⚙</span>Réglages kDrive</button>
      <button class="cb-menu-i" data-act="theme" role="menuitem"><span>◐</span>Basculer le thème</button>
      <button class="cb-menu-i" data-act="recharger" role="menuitem"><span>⟳</span>Recharger la page</button>
    </div>
  </div>

  <div class="cb-toast" id="toast"></div>`;

  document.head.innerHTML = '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">';
  document.title = 'Carnet de bord';
  const style = document.createElement('style');
  style.textContent = CSS;
  document.head.appendChild(style);

  // Leaflet fournit sa CSS via @resource : on l'injecte avant la nôtre pour pouvoir la surcharger.
  try {
    const cssLeaflet = GM_getResourceText('leafletcss');
    if (cssLeaflet) {
      const sl = document.createElement('style');
      sl.textContent = cssLeaflet;
      document.head.insertBefore(sl, style);
    }
  } catch (e) { /* la carte restera fonctionnelle mais sans styles Leaflet */ }

  document.body.innerHTML = HTML;
  const $ = s => document.querySelector(s);

  function appliquerTheme() {
    document.documentElement.setAttribute('data-theme', theme);
    $('#btn-theme').textContent = theme === 'dark' ? '☀' : '☾';
  }
  appliquerTheme();

  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove('show'), 4000);
  }
  const majBandeau = txt => { $('#bandeau').textContent = txt; };

  /* ================= rendu ================= */
  function visibles() {
    const q = recherche.trim().toLowerCase();
    return items
      .filter(it => filtre === 'toutes' || (filtre === 'actives' ? estActif(it) : it.etat === filtre))
      .filter(it => !q || [it.entreprise, it.poste, it.notes, it.type, it.etat, it.ville].join(' ').toLowerCase().includes(q))
      .sort((a, b) => ((aRelancer(b) ? 1 : 0) - (aRelancer(a) ? 1 : 0)) || (b.date || '').localeCompare(a.date || ''));
  }

  function rendre() {
    $('#etat-fichier').innerHTML = charge
      ? `${esc(cfg.fileName)}<span class="cb-tech"> · drive ${esc(cfg.driveId)} · fichier ${esc(cfg.fileId)}</span>${modifie ? ' · <b>non enregistré</b>' : ' · à jour'}`
      : 'fichier non chargé';
    const heure = dernierControle ? new Date(dernierControle).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : null;
    majBandeau(charge
      ? (modifie ? 'Modifications en attente d\'envoi.' : `Synchronisé avec kDrive${heure ? ' · vérifié à ' + heure : ''}.`)
      : 'Chargez le fichier pour commencer.');
    $('#btn-enregistrer').disabled = !charge;
    $('#btn-nettoyer').disabled = !charge;

    const barre = $('#bar');
    barre.classList.toggle('on', modifie || conflit);
    barre.classList.toggle('conflit', conflit);
    $('#bar-recharger').style.display = conflit ? '' : 'none';
    if (conflit) {
      $('#bar-titre').textContent = 'Le fichier a changé sur kDrive';
      $('#bar-texte').textContent = modifie
        ? 'Un autre appareil l\'a modifié. Enregistrer écrasera ces changements.'
        : 'Rechargez pour récupérer la version à jour.';
      $('#bar-save').textContent = 'Enregistrer quand même';
    } else {
      $('#bar-titre').textContent = 'Modifications non enregistrées';
      $('#bar-texte').textContent = 'Elles seront perdues si vous fermez l\'onglet.';
      $('#bar-save').textContent = 'Enregistrer sur kDrive';
    }

    $('#n-cand').textContent = items.length;
    $('#n-prosp').textContent = prospects.length;
    document.querySelectorAll('.cb-onglet').forEach(b =>
      b.setAttribute('aria-selected', b.dataset.onglet === onglet));
    // Barre de navigation mobile : même état actif, badges de comptage.
    document.querySelectorAll('.cb-nav-i[data-onglet]').forEach(b =>
      b.setAttribute('aria-selected', b.dataset.onglet === onglet));
    $('#nb-cand').textContent = items.length || '';
    $('#nb-prosp').textContent = prospects.length || '';
    const prospectsVue = onglet === 'prospects';
    $('#chips').style.display = prospectsVue ? 'none' : '';
    $('#radar').style.display = prospectsVue ? 'none' : '';
    $('#btn-ajouter').textContent = prospectsVue ? '+ Entreprise à postuler' : '+ Candidature';
    $('#q').placeholder = prospectsVue
      ? 'Chercher une entreprise, un poste, une note…'
      : 'Chercher une entreprise, un poste, une note…';

    const relances = items.filter(aRelancer).length;
    $('#stats').innerHTML = [
      ['Envoyées', items.length, 'var(--text)'],
      ['En cours', items.filter(estActif).length, 'var(--yellow)'],
      ['Entretiens', items.filter(it => it.etat === 'Entretien').length, 'var(--sky)'],
      ['À relancer', relances, relances ? 'var(--peach)' : 'var(--overlay0)'],
      ['Refus', items.filter(it => it.etat === 'Refus').length, 'var(--red)'],
      ['À postuler', prospects.length, prospects.length ? 'var(--lavender)' : 'var(--overlay0)']
    ].map(([l, v, c]) => `<div class="cb-stat"><b style="color:${c}">${v}</b><small>${l}</small></div>`).join('');

    const cibles = items.filter(aRelancer).sort((a, b) => attente(b) - attente(a));
    const r = $('#radar');
    if (!cibles.length) {
      r.className = 'cb-radar quiet';
      r.innerHTML = `<div class="cb-radar-t">Relances</div><div class="cb-radar-b">Rien à relancer. Une candidature remonte ici après ${SEUIL} jours sans nouvelle.</div>`;
    } else {
      r.className = 'cb-radar';
      r.innerHTML = `<div class="cb-radar-t">À relancer — ${cibles.length}</div><div class="cb-radar-b">`
        + cibles.slice(0, 6).map(it => `<button data-relance="${it.id}">${esc(it.entreprise)} · ${attente(it)} j</button>`).join(' &nbsp;· ')
        + (cibles.length > 6 ? ` et ${cibles.length - 6} autre(s)` : '') + '</div>';
    }

    const defs = [['toutes', 'Toutes'], ['actives', 'En cours'], ...Object.keys(ETATS).map(e => [e, e])];
    $('#chips').innerHTML = defs.map(([k, l]) => {
      const n = k === 'toutes' ? items.length : k === 'actives' ? items.filter(estActif).length : items.filter(it => it.etat === k).length;
      if (!n && k !== 'toutes' && k !== 'actives') return '';
      return `<button class="cb-chip" data-filtre="${esc(k)}" aria-pressed="${filtre === k}">${esc(l)}<i>${n}</i></button>`;
    }).join('');

    if (prospectsVue) { rendreProspects(); return; }

    const liste = visibles();
    $('#list').innerHTML = liste.length ? liste.map(it => {
      const conf = ETATS[it.etat] || { tone: 'var(--overlay0)', actif: false };
      const j = attente(it), actif = conf.actif;
      const age = actif ? tonAge(j) : 'var(--overlay0)';
      let host = '';
      if (it.source) { try { host = new URL(it.source).hostname.replace(/^www\./, ''); } catch (e) { host = 'lien'; } }
      return `<li class="cb-card" data-id="${it.id}" style="--tone:${conf.tone};--age:${age}">
        <div>
          <div class="cb-t"><b>${esc(it.entreprise) || 'Sans nom'}</b>${it.ville ? `<span class="cb-lieu">${esc(it.ville)}</span>` : ''}${it.source ? `<a href="${esc(it.source)}" target="_blank" rel="noopener">↗ ${esc(host)}</a>` : ''}</div>
          <div class="cb-poste">${esc(it.poste)}</div>
          ${it.notes ? `<div class="cb-notes">${esc(it.notes)}</div>` : ''}
          <div class="cb-meta">
            <span class="cb-tag cb-etat">${esc(it.etat)}</span>
            ${it.type ? `<span class="cb-tag">${esc(it.type)}</span>` : ''}
            <span>envoyée le ${dateFr(it.date)}</span>
            ${it.relances.length ? `<span>· ${it.relances.length} relance${it.relances.length > 1 ? 's' : ''}, dernière le ${dateFr(it.relances[it.relances.length - 1])}</span>` : ''}
          </div>
        </div>
        <div class="cb-side">
          <div class="cb-age"><b>${j == null ? '—' : j + ' j'}</b><small>${actif ? 'sans réponse' : 'clos'}</small></div>
          ${actif ? `<div class="cb-gauge"><span style="width:${Math.min(100, Math.round((j ?? 0) / 21 * 100))}%"></span></div>` : ''}
          <div>${actif ? `<button class="cb-ico" data-relance="${it.id}">+ relance</button>` : ''}<button class="cb-ico" data-edit="${it.id}">modifier</button></div>
        </div>
      </li>`;
    }).join('') : `<li class="cb-empty"><b>${charge ? (items.length ? 'Aucun résultat' : 'Le carnet est vide') : 'Rien de chargé'}</b>${charge ? (items.length ? 'Changez de filtre ou videz la recherche.' : 'Ajoutez votre première candidature.') : 'Cliquez sur « Charger ».'}</li>`;
  }

  /* ================= carte ================= */
  // Requête sans en-tête d'autorisation : le token kDrive ne doit jamais partir ailleurs.
  function requeteLibre(url) {
    return new Promise((resolve, reject) => {
      XHR({
        method: 'GET', url,
        headers: {
          'Accept': 'application/json',
          'Accept-Language': 'fr',
          // La politique d'usage de Nominatim impose une application identifiable.
          'User-Agent': 'CarnetDeBord/2.10 (userscript personnel de suivi de candidatures)'
        },
        timeout: 20000,
        onload: r => resolve(r),
        onerror: () => reject(new Error('Requête échouée')),
        ontimeout: () => reject(new Error('Délai dépassé'))
      });
    });
  }

  const requete = it => (it.adresse || '').trim() || (it.ville || '').trim();
  const coords = it => geo[requete(it)] || null;
  const aGeocoder = () => Array.from(new Set(items.map(requete).filter(q => q && !geo[q])));

  // Nominatim découpe mal une adresse sans virgules, et bute sur les codes de site
  // internes aux entreprises. On dérive donc plusieurs formulations, de la plus
  // complète à la plus large, et on s'arrête à la première qui répond.
  function variantesRequete(q) {
    const base = String(q).replace(/\s+/g, ' ').trim();
    const out = [];
    const push = v => {
      v = String(v).replace(/\s*,\s*/g, ', ').replace(/\s+/g, ' ').replace(/^[,\s]+|[,\s]+$/g, '');
      if (v && !out.includes(v)) out.push(v);
    };
    const VOIRIE = /\b(\d+\s+)?(rue|route|avenue|av|boulevard|bd|chemin|impasse|place|all[ée]e|quai|cours|square|zone|za|zi|zac|parc|lieu[- ]dit)\b/i;

    push(base);

    const voie = base.match(VOIRIE);
    const debutVoie = voie && voie.index > 0 ? voie.index : 0;
    if (debutVoie) push(base.slice(debutVoie));

    const cp = base.match(/\b\d{5}\b/);
    if (cp) {
      const iCp = base.indexOf(cp[0]);
      push(base.slice(debutVoie, iCp) + ', ' + base.slice(iCp));
      push(base.slice(iCp));
    }
    return out.map(v => /\bfrance\b/i.test(v) ? v : v + ', France');
  }

  let traceGeo = [];

  async function geocoderUn(q) {
    const base = 'https://nominatim.openstreetmap.org/search?format=json&limit=1';
    traceGeo = [];

    // Une ville seule part en requête structurée : Nominatim cherche alors une commune,
    // et non un lieu-dit ou une église qui porterait le même nom ailleurs.
    if (villes.includes(q)) {
      try {
        const r = await requeteLibre(`${base}&city=${encodeURIComponent(q)}&country=France`);
        const d = r.status === 200 ? JSON.parse(r.responseText) : [];
        traceGeo.push(`ville « ${q} » → HTTP ${r.status}, ${d.length} résultat(s)`);
        if (d.length) return { lat: parseFloat(d[0].lat), lon: parseFloat(d[0].lon) };
      } catch (e) { traceGeo.push(`ville « ${q} » → ${e.message}`); }
      return null;
    }

    // Chaque tentative est isolée : une erreur réseau sur l'une ne doit pas
    // interrompre la cascade, sinon les formulations suivantes ne sont jamais testées.
    const essais = variantesRequete(q);
    for (let i = 0; i < essais.length; i++) {
      if (i) await new Promise(r => setTimeout(r, 1200)); // une requête par seconde maximum
      try {
        const r = await requeteLibre(`${base}&q=${encodeURIComponent(essais[i])}`);
        const d = r.status === 200 ? JSON.parse(r.responseText) : [];
        traceGeo.push(`${i + 1}. « ${essais[i]} » → HTTP ${r.status}, ${d.length} résultat(s)`);
        if (d.length) return { lat: parseFloat(d[0].lat), lon: parseFloat(d[0].lon), via: essais[i] };
      } catch (e) {
        traceGeo.push(`${i + 1}. « ${essais[i]} » → ${e.message}`);
      }
    }
    return null;
  }

  async function geocoderTout() {
    if (geocodageEnCours) return;
    const files = aGeocoder();
    if (!files.length) { toast('Toutes les adresses sont déjà localisées.'); return; }
    geocodageEnCours = true;
    const btn = $('#m-geo');
    let trouves = 0;
    const echecs = [];
    for (let i = 0; i < files.length; i++) {
      btn.textContent = `Localisation ${i + 1}/${files.length}…`;
      try {
        const c = await geocoderUn(files[i]);
        if (c) { geo[files[i]] = { lat: c.lat, lon: c.lon }; trouves++; }
        else echecs.push({ q: files[i], trace: traceGeo.slice() });
      } catch (e) { echecs.push({ q: files[i], trace: [e.message] }); }
      // Nominatim impose au plus une requête par seconde.
      if (i < files.length - 1) await new Promise(r => setTimeout(r, 1200));
    }
    geocodageEnCours = false;
    if (trouves) { toucher(); rendre(); }
    rendreMarqueurs();
    majBoutonGeo();
    if (echecs.length) afficherEchecs(echecs);
    toast(trouves ? `${trouves} adresse${trouves > 1 ? 's' : ''} localisée${trouves > 1 ? 's' : ''} — pensez à enregistrer.` : 'Aucune adresse n\'a pu être localisée.');
  }

  function afficherTrace() {
    if (!traceGeo.length) return;
    $('#m-info').innerHTML = 'Tentatives de localisation :<br>' + traceGeo.map(esc).join('<br>');
  }

  function afficherEchecs(echecs) {
    $('#m-info').innerHTML = echecs.map(e =>
      `<strong>${esc(e.q)}</strong><br>` + e.trace.map(esc).join('<br>')
    ).join('<br><br>');
  }

  function majBoutonGeo() {
    const n = aGeocoder().length;
    const btn = $('#m-geo');
    btn.disabled = !n || geocodageEnCours;
    btn.textContent = n ? `Localiser ${n} adresse${n > 1 ? 's' : ''}` : 'Tout est localisé';
  }

  const SEUIL_PX = 54;      // supérieur au diamètre d'une pastille, sinon elles se chevauchent
  const LETTRES = {
    'En attente': 'A', 'Relancé': 'L', 'Entretien': 'E',
    'Offre reçue': 'O', 'Refus': 'R', 'Offre retirée': 'X'
  };
  const ECART_PX = 13;      // écart appliqué aux candidatures d'une même adresse
  let eclates = new Set();  // adresses dépliées manuellement

  function rendreMarqueurs(ajuster) {
    if (!coucheMarqueurs) return;
    coucheMarqueurs.clearLayers();
    marqueurs = new Map();
    const contour = theme === 'dark' ? '#11111b' : '#ffffff';
    const etiquettes = $('#m-noms').checked;

    // 1. Regrouper les candidatures partageant exactement les mêmes coordonnées.
    const lieux = new Map();
    items.forEach(it => {
      const c = coords(it);
      if (!c) return;
      const cle = c.lat.toFixed(5) + ',' + c.lon.toFixed(5);
      if (!lieux.has(cle)) lieux.set(cle, { cle, lat: c.lat, lon: c.lon, liste: [] });
      lieux.get(cle).liste.push(it);
    });

    // 2. Agréger les lieux trop proches, en fusionnant de proche en proche jusqu'à
    //    ce qu'aucune paire de paquets ne soit plus assez rapprochée pour se recouvrir.
    const centroide = liste => {
      const pts = liste.map(l => carte.latLngToLayerPoint([l.lat, l.lon]));
      return {
        x: pts.reduce((s, p) => s + p.x, 0) / pts.length,
        y: pts.reduce((s, p) => s + p.y, 0) / pts.length
      };
    };
    const paquets = Array.from(lieux.values()).map(l => ({
      lieux: [l], pt: carte.latLngToLayerPoint([l.lat, l.lon])
    }));
    let fusionne = true;
    while (fusionne) {
      fusionne = false;
      for (let i = 0; i < paquets.length && !fusionne; i++) {
        for (let j = i + 1; j < paquets.length; j++) {
          if (Math.hypot(paquets[i].pt.x - paquets[j].pt.x, paquets[i].pt.y - paquets[j].pt.y) < SEUIL_PX) {
            paquets[i].lieux = paquets[i].lieux.concat(paquets[j].lieux);
            paquets[i].pt = centroide(paquets[i].lieux);
            paquets.splice(j, 1);
            fusionne = true;
            break;
          }
        }
      }
    }

    const points = [];

    const marqueurSimple = (it, lat, lon) => {
      const couleur = TEINTES[it.etat] || '#6c7086';
      const j = attente(it);
      if (aRelancer(it)) {
        L.circleMarker([lat, lon], { radius: 19, stroke: false, fillColor: couleur, fillOpacity: .18, interactive: false })
          .addTo(coucheMarqueurs);
      }
      const m = L.circleMarker([lat, lon], {
        radius: 11, color: contour, weight: 3, opacity: .9, fillColor: couleur, fillOpacity: .95
      }).bindPopup(
        `<strong>${esc(it.entreprise)}</strong><br>${esc(it.poste) || 'poste non précisé'}<br>`
        + `<span style="color:${couleur}">${esc(it.etat)}</span>`
        + `${estActif(it) && j != null ? ` · ${j} j sans réponse` : ''}<br>`
        + `${esc(it.adresse || it.ville)}`
        + (it.source ? `<br><a href="${esc(it.source)}" target="_blank" rel="noopener">Voir l'offre</a>` : '')
        + `<br><button class="cb-mini" data-repos="${esc(requete(it))}">Repositionner</button>`
        + `<button class="cb-mini" data-regeo="${esc(requete(it))}">Recalculer</button>`
      ).addTo(coucheMarqueurs);
      m._libelle = esc(it.entreprise);
      m.bindTooltip(m._libelle, { permanent: etiquettes, direction: 'top', offset: [0, -12], className: 'cb-tip' });
      marqueurs.set(it.id, m);
      points.push([lat, lon]);
    };

    // Candidatures d'une même adresse : écart en pixels, pour qu'aucune ne soit masquée.
    const deplier = lieu => {
      lieu.liste.forEach((it, i) => {
        let lat = lieu.lat, lon = lieu.lon;
        if (lieu.liste.length > 1) {
          const angle = (2 * Math.PI * i) / lieu.liste.length;
          const centre = carte.latLngToLayerPoint([lat, lon]);
          const d = carte.layerPointToLatLng([
            centre.x + ECART_PX * Math.cos(angle),
            centre.y + ECART_PX * Math.sin(angle)
          ]);
          lat = d.lat; lon = d.lng;
        }
        marqueurSimple(it, lat, lon);
      });
    };

    paquets.forEach(p => {
      const total = p.lieux.reduce((n, l) => n + l.liste.length, 0);
      const memeAdresse = p.lieux.length === 1;

      // Un seul point, ou une adresse que l'on a demandé à déplier : affichage détaillé.
      if (total === 1 || (memeAdresse && eclates.has(p.lieux[0].cle))) {
        p.lieux.forEach(deplier);
        return;
      }

      // Sinon, une pastille unique portant le nombre de candidatures.
      const lat = p.lieux.reduce((s, l) => s + l.lat, 0) / p.lieux.length;
      const lon = p.lieux.reduce((s, l) => s + l.lon, 0) / p.lieux.length;
      const membres = p.lieux.flatMap(l => l.liste);
      const etats = new Set(membres.map(it => it.etat));
      const couleur = etats.size === 1 ? (TEINTES[membres[0].etat] || '#6c7086') : '#b4befe';
      const taille = total > 9 ? 40 : 34;

      // Une pastille lavande mélange plusieurs états : on détaille le compte à côté.
      let detail = '';
      if (etats.size > 1) {
        detail = '<span class="cb-detail">' + Object.keys(TEINTES)
          .map(e => [e, membres.filter(it => it.etat === e).length])
          .filter(([, n]) => n)
          .map(([e, n]) => `<b style="--c:${TEINTES[e]}">${n}${LETTRES[e] || '?'}</b>`)
          .join('') + '</span>';
      }

      const grappe = L.marker([lat, lon], {
        icon: L.divIcon({
          className: 'cb-grappe-wrap',
          html: `<div class="cb-grappe" style="--c:${couleur};width:${taille}px;height:${taille}px">${total}</div>${detail}`,
          iconSize: [taille, taille],
          iconAnchor: [taille / 2, taille / 2]
        })
      }).addTo(coucheMarqueurs);

      grappe.bindTooltip(
        membres.slice(0, 6).map(it => esc(it.entreprise)).join('<br>')
        + (membres.length > 6 ? `<br>et ${membres.length - 6} autre(s)` : '')
        + `<br><em>${memeAdresse ? 'même adresse — cliquer pour déplier' : 'cliquer pour zoomer'}</em>`,
        { direction: 'top', offset: [0, -taille / 2], className: 'cb-tip' }
      );

      grappe.on('click', () => {
        if (memeAdresse) {
          eclates.add(p.lieux[0].cle);
          rendreMarqueurs(false);
        } else {
          carte.fitBounds(p.lieux.map(l => [l.lat, l.lon]), {
            padding: [60, 60], maxZoom: Math.min(18, carte.getZoom() + 4)
          });
        }
      });

      points.push([lat, lon]);
      membres.forEach(it => marqueurs.set(it.id, grappe));
    });

    $('#m-info').textContent = lieux.size
      ? `${items.filter(coords).length} candidature${items.filter(coords).length > 1 ? 's' : ''} placée${items.filter(coords).length > 1 ? 's' : ''} sur ${items.length}.`
      : 'Aucune candidature localisée pour l\'instant.';
    if (ajuster && points.length) carte.fitBounds(points, { padding: [50, 50], maxZoom: 12 });
    rendrePanneau();
  }

  // Panneau latéral : toutes les candidatures, placées ou non, pour qu'aucune
  // ne devienne inaccessible faute de point sur la carte.
  function rendrePanneau() {
    const el = $('#m-liste');
    const placees = [], absentes = [];
    items.forEach(it => (coords(it) ? placees : absentes).push(it));
    const ligne = (it, place) => {
      const couleur = TEINTES[it.etat] || '#6c7086';
      const lieu = it.adresse || it.ville || 'aucune adresse';
      const q = requete(it);
      return `<div class="cb-pin${place ? '' : ' absente'}">
        <button class="cb-pin-main" data-focus="${it.id}">
          <i class="${place ? '' : 'vide'}" style="background:${couleur}"></i>
          <span><b>${esc(it.entreprise) || 'Sans nom'}</b><small>${esc(lieu)} · ${esc(it.etat)}</small></span>
          ${place ? '' : '<em>à placer</em>'}
        </button>
        ${q ? `<button class="cb-mini" data-regeo="${esc(q)}" title="Relancer la recherche d'adresse">↻</button>` : ''}
      </div>`;
    };
    el.innerHTML =
      (absentes.length ? `<div class="cb-groupe">Non placées — ${absentes.length}</div>` + absentes.map(it => ligne(it, false)).join('') : '')
      + (placees.length ? `<div class="cb-groupe">Sur la carte — ${placees.length}</div>` + placees.map(it => ligne(it, true)).join('') : '')
      + (items.length ? '' : '<div class="cb-groupe">Aucune candidature</div>');
  }

  function viser(id) {
    const it = items.find(x => x.id === id);
    if (!it) return;
    const c = coords(it);
    if (c) {
      eclates.add(c.lat.toFixed(5) + ',' + c.lon.toFixed(5));
      carte.setView([c.lat, c.lon], Math.max(carte.getZoom(), 15), { animate: true });
      rendreMarqueurs(false);
      setTimeout(() => {
        const m = marqueurs.get(id);
        if (m && m.openPopup) m.openPopup();
      }, 350);
      return;
    }
    const q = requete(it);
    if (!q) { toast('Renseignez d\'abord une adresse ou une localisation pour cette candidature.'); return; }
    repositionner = q;
    document.getElementById('map').style.cursor = 'crosshair';
    toast(`Cliquez sur la carte pour placer ${it.entreprise}.`);
  }

  // Bascule des étiquettes sans reconstruire la carte : on se contente de rebrancher
  // les infobulles des marqueurs individuels, les pastilles gardant la leur au survol.
  function basculerNoms() {
    const actif = $('#m-noms').checked;
    new Set(marqueurs.values()).forEach(m => {
      if (!m._libelle) return;
      m.unbindTooltip();
      m.bindTooltip(m._libelle, { permanent: actif, direction: 'top', offset: [0, -12], className: 'cb-tip' });
    });
  }

  function ouvrirCarte() {
    $('#ov-carte').classList.add('open');
    if (!carte) {
      carte = L.map('map', { scrollWheelZoom: true }).setView([48.11, -1.68], 8);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19, attribution: '© OpenStreetMap'
      }).addTo(carte);
      coucheMarqueurs = L.layerGroup().addTo(carte);
      carte.on('zoomend', () => rendreMarqueurs(false));
      carte.on('click', e => {
        if (!repositionner) return;
        geo[repositionner] = { lat: e.latlng.lat, lon: e.latlng.lng };
        const q = repositionner;
        repositionner = null;
        document.getElementById('map').style.cursor = '';
        toucher(); rendre(); rendreMarqueurs();
        toast(`Position de « ${q} » corrigée — pensez à enregistrer.`);
      });
    }
    majBoutonGeo();
    $('#m-legende').innerHTML = Object.keys(TEINTES)
      .filter(e => items.some(it => it.etat === e))
      .map(e => `<span><i style="background:${TEINTES[e]}"></i>${esc(e)}</span>`).join('')
      + '<span><i style="background:transparent;box-shadow:0 0 0 1px var(--peach)"></i>halo = à relancer</span>'
      + '<span>pastille lavande = états mélangés, détail à droite (A attente, L relancé, E entretien, O offre, R refus, X retirée)</span>';
    setTimeout(() => { carte.invalidateSize(); rendreMarqueurs(true); }, 80);
  }

  function rendreProspects() {
    const q = recherche.trim().toLowerCase();
    const liste = prospects
      .filter(p => !q || [p.entreprise, p.poste, p.notes, p.ville].join(' ').toLowerCase().includes(q))
      .sort((a, b) => (b.reperee || '').localeCompare(a.reperee || ''));

    $('#list').innerHTML = liste.length ? liste.map(p => {
      let host = '';
      if (p.source) { try { host = new URL(p.source).hostname.replace(/^www\./, ''); } catch (e) { host = 'lien'; } }
      const j = jours(p.reperee);
      return `<li class="cb-card cb-prospect" data-idp="${p.id}">
        <div>
          <div class="cb-t"><b>${esc(p.entreprise) || 'Sans nom'}</b>${p.ville ? `<span class="cb-lieu">${esc(p.ville)}</span>` : ''}${p.source ? `<a href="${esc(p.source)}" target="_blank" rel="noopener">↗ ${esc(host)}</a>` : ''}</div>
          <div class="cb-poste">${esc(p.poste) || 'poste à préciser'}</div>
          ${p.notes ? `<div class="cb-notes">${esc(p.notes)}</div>` : ''}
          <div class="cb-meta">
            ${p.reperee ? `<span>repérée le ${dateFr(p.reperee)}${j != null && j > 0 ? ` · il y a ${j} j` : ''}</span>` : ''}
            ${p.adresse ? `<span>· ${esc(p.adresse)}</span>` : ''}
          </div>
        </div>
        <div class="cb-side">
          <div>
            <button class="cb-ico" data-postuler="${p.id}">postuler</button>
            <button class="cb-ico" data-editp="${p.id}">modifier</button>
          </div>
        </div>
      </li>`;
    }).join('') : `<li class="cb-empty"><b>${prospects.length ? 'Aucun résultat' : 'Aucune entreprise en attente'}</b>${prospects.length ? 'Videz la recherche.' : 'Notez ici les entreprises repérées, avant d\'avoir postulé.'}</li>`;
  }

  /* ================= historique ================= */
  function ouvrirHistorique() {
    const enAttente = charge ? diff() : [];
    const tout = journal.concat(enAttente.map(e => Object.assign({ attente: true }, e)));
    const el = $('#hist');
    if (!tout.length) {
      el.innerHTML = '<div class="cb-empty"><b>Aucune modification consignée</b>Le journal se remplit à chaque enregistrement sur kDrive.</div>';
    } else {
      const parJour = new Map();
      tout.slice().reverse().forEach(e => {
        const j = (e.at || '').slice(0, 10) || 'Date inconnue';
        if (!parJour.has(j)) parJour.set(j, []);
        parJour.get(j).push(e);
      });
      el.innerHTML = Array.from(parJour.entries()).map(([j, entrees]) => {
        const titre = j === today() ? 'Aujourd\'hui' : dateFr(j);
        return `<div class="cb-jour">${esc(titre)}</div>` + entrees.map(e => `
          <div class="cb-ent">
            <b>${esc(e.entreprise) || '—'}</b> <i>${esc(e.champ)}${e.attente ? ' · non enregistré' : ''}</i><br>
            ${e.avant ? `<span class="av">${esc(e.avant)}</span> → ` : ''}<span class="ap">${esc(e.apres)}</span>
          </div>`).join('');
      }).join('');
    }
    $('#ov-hist').classList.add('open');
  }

  /* ================= nettoyage ================= */
  let propsEnCours = [];
  function ouvrirNettoyage() {
    propsEnCours = propositions();
    const el = $('#net');
    if (!propsEnCours.length) {
      el.innerHTML = '<div class="cb-empty"><b>Rien à corriger</b>Les libellés sont déjà homogènes.</div>';
      $('#n-appliquer').disabled = true;
    } else {
      $('#n-appliquer').disabled = false;
      el.innerHTML = propsEnCours.map(x => `
        <div class="cb-ent">
          <b>${esc(x.entreprise) || '—'}</b><br>
          ${x.changes.map(c => `<i>${esc(c.champ)}</i> <span class="av">${esc(c.avant) || '(vide)'}</span> → <span class="ap">${esc(c.apres)}</span>`).join('<br>')}
        </div>`).join('');
    }
    $('#ov-net').classList.add('open');
  }
  function appliquerNettoyage() {
    const par = new Map(propsEnCours.map(x => [x.id, x.p]));
    items = items.map(it => par.has(it.id) ? Object.assign({}, it, par.get(it.id)) : it);
    toucher(); $('#ov-net').classList.remove('open'); rendre();
    toast(`${propsEnCours.length} fiche${propsEnCours.length > 1 ? 's' : ''} corrigée${propsEnCours.length > 1 ? 's' : ''} — pensez à enregistrer.`);
  }

  /* ================= fiche ================= */
  function rendreRelances() {
    $('#f-relances').innerHTML = brouillonRelances.length
      ? brouillonRelances.map((d, i) => `${dateFr(d)} <button class="cb-ico" data-rmrel="${i}">retirer</button>`).join(' · ')
      : 'Aucune.';
  }
  function rendreVilles(courante, prefixe) {
    prefixe = prefixe || 'f';
    if (courante && !villes.includes(courante)) { villes.push(courante); trierVilles(); }
    const sel = $('#' + prefixe + '-ville');
    sel.innerHTML = '<option value="">— non précisée —</option>'
      + villes.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('')
      + '<option value="__new">+ Ajouter une ville…</option>';
    sel.value = courante || '';
    basculerVille(prefixe);
  }
  function basculerVille(prefixe) {
    prefixe = prefixe || 'f';
    const neuf = $('#' + prefixe + '-ville').value === '__new';
    const champ = $('#' + prefixe + '-ville-new');
    champ.style.display = neuf ? '' : 'none';
    champ.style.marginTop = '6px';
    if (neuf) { champ.value = ''; champ.focus(); }
  }
  function villeSaisie(prefixe) {
    let v = $('#' + prefixe + '-ville').value;
    if (v === '__new') v = $('#' + prefixe + '-ville-new').value.replace(/\s+/g, ' ').trim();
    if (v && !villes.includes(v)) { villes.push(v); trierVilles(); }
    return v;
  }

  function ouvrirFiche(id) {
    editId = id;
    const it = id ? items.find(x => x.id === id) : null;
    brouillonRelances = it ? it.relances.slice() : [];
    $('#fiche-titre').textContent = it ? 'Modifier la candidature' : 'Nouvelle candidature';
    $('#f-entreprise').value = it?.entreprise || '';
    $('#f-poste').value = it?.poste || '';
    $('#f-adresse').value = it?.adresse || '';
    rendreVilles(it?.ville || '');
    $('#f-type').value = it?.type || SUJETS[1];
    $('#f-date').value = it?.date || today();
    $('#f-source').value = it?.source || '';
    $('#f-notes').value = it?.notes || '';
    $('#f-relance-add').value = '';
    const etats = Array.from(new Set([...Object.keys(ETATS), ...items.map(x => x.etat).filter(Boolean)]));
    $('#f-etat').innerHTML = etats.map(e => `<option>${esc(e)}</option>`).join('');
    $('#f-etat').value = it?.etat || 'En attente';
    $('#l-types').innerHTML = Array.from(new Set([...SUJETS, ...items.map(x => x.type).filter(Boolean)]))
      .map(t => `<option value="${esc(t)}">`).join('');
    $('#f-suppr').style.display = it ? '' : 'none';
    rendreRelances();
    $('#ov-fiche').classList.add('open');
    $('#f-entreprise').focus();
  }
  const fermerFiche = () => { $('#ov-fiche').classList.remove('open'); editId = null; };

  function validerFiche() {
    const entreprise = $('#f-entreprise').value.trim();
    if (!entreprise) { $('#f-entreprise').focus(); toast('Il manque le nom de l\'entreprise.'); return; }
    const ville = villeSaisie('f');
    if ($('#f-ville').value === '__new' && !ville) {
      $('#f-ville-new').focus(); toast('Saisissez le nom de la nouvelle ville, ou choisissez « non précisée ».'); return;
    }
    const ajout = $('#f-relance-add').value;
    if (ajout && !brouillonRelances.includes(ajout)) brouillonRelances.push(ajout);
    brouillonRelances.sort();
    const data = {
      entreprise,
      poste: $('#f-poste').value.trim(),
      ville: ville,
      adresse: $('#f-adresse').value.replace(/\s+/g, ' ').trim(),
      type: $('#f-type').value.trim(),
      etat: $('#f-etat').value,
      date: $('#f-date').value,
      source: $('#f-source').value.trim(),
      notes: $('#f-notes').value.trim(),
      relances: brouillonRelances.slice()
    };
    if (editId) items = items.map(it => it.id === editId ? Object.assign({}, it, data) : it);
    else items.unshift(Object.assign({ id: uid() }, data));
    toucher(); fermerFiche(); rendre();
  }

  /* ================= entreprises à postuler ================= */
  let editProspect = null;

  function ouvrirProspect(id) {
    editProspect = id;
    const p = id ? prospects.find(x => x.id === id) : null;
    $('#prospect-titre').textContent = p ? 'Modifier l\'entreprise' : 'Entreprise à postuler';
    $('#p-entreprise').value = p?.entreprise || '';
    $('#p-poste').value = p?.poste || '';
    $('#p-adresse').value = p?.adresse || '';
    $('#p-source').value = p?.source || '';
    $('#p-reperee').value = p?.reperee || today();
    $('#p-notes').value = p?.notes || '';
    rendreVilles(p?.ville || '', 'p');
    $('#p-suppr').style.display = p ? '' : 'none';
    $('#ov-prospect').classList.add('open');
    $('#p-entreprise').focus();
  }

  function validerProspect() {
    const entreprise = $('#p-entreprise').value.trim();
    if (!entreprise) { $('#p-entreprise').focus(); toast('Il manque le nom de l\'entreprise.'); return; }
    const data = {
      entreprise,
      poste: $('#p-poste').value.trim(),
      ville: villeSaisie('p'),
      adresse: $('#p-adresse').value.replace(/\s+/g, ' ').trim(),
      source: $('#p-source').value.trim(),
      reperee: $('#p-reperee').value,
      notes: $('#p-notes').value.trim()
    };
    if (editProspect) prospects = prospects.map(p => p.id === editProspect ? Object.assign({}, p, data) : p);
    else prospects.unshift(Object.assign({ id: uid() }, data));
    toucher();
    $('#ov-prospect').classList.remove('open');
    editProspect = null;
    rendre();
  }

  // Passage de la liste d'intentions aux candidatures réelles : la fiche s'ouvre
  // aussitôt pour compléter ce qui ne se devine pas, à commencer par le sujet.
  function convertir(id) {
    const p = prospects.find(x => x.id === id);
    if (!p) return;
    const nouvelle = {
      id: uid(), entreprise: p.entreprise, poste: p.poste, ville: p.ville, adresse: p.adresse,
      source: p.source, notes: p.notes, type: 'Candidature spontanée', etat: 'En attente',
      date: today(), relances: []
    };
    items.unshift(nouvelle);
    prospects = prospects.filter(x => x.id !== id);
    onglet = 'candidatures';
    toucher(); rendre();
    ouvrirFiche(nouvelle.id);
    toast('Passée en candidature — vérifiez le sujet et la date.');
  }

  /* ================= réglages ================= */
  function ouvrirReglages() {
    $('#c-token').value = cfg.token;
    $('#c-drive').value = cfg.driveId;
    $('#c-file').value = cfg.fileId;
    $('#c-parent').value = cfg.parentId;
    $('#c-nom').value = cfg.fileName;
    $('#c-mode').value = cfg.uploadMode;
    $('#ov-reglages').classList.add('open');
  }
  function lireReglages() {
    cfg.token = $('#c-token').value.trim();
    cfg.driveId = $('#c-drive').value.trim();
    cfg.fileId = $('#c-file').value.trim();
    cfg.parentId = $('#c-parent').value.trim();
    cfg.fileName = $('#c-nom').value.trim();
    cfg.uploadMode = $('#c-mode').value;
  }

  /* ================= événements ================= */
  document.addEventListener('click', e => {
    const t = e.target.closest('[data-relance],[data-edit],[data-filtre],[data-rmrel],[data-repos],[data-regeo],[data-focus],[data-onglet],[data-editp],[data-postuler]');
    if (!t) return;
    if (t.dataset.onglet) { onglet = t.dataset.onglet; recherche = ''; $('#q').value = ''; rendre(); return; }
    if (t.dataset.editp) { ouvrirProspect(t.dataset.editp); return; }
    if (t.dataset.postuler) { convertir(t.dataset.postuler); return; }
    if (t.dataset.focus) { viser(t.dataset.focus); return; }
    if (t.dataset.repos) {
      repositionner = t.dataset.repos;
      carte.closePopup();
      document.getElementById('map').style.cursor = 'crosshair';
      toast('Cliquez sur la carte à l\'emplacement exact.');
      return;
    }
    if (t.dataset.regeo) {
      const q = t.dataset.regeo;
      delete geo[q];
      carte.closePopup();
      toast('Nouvelle recherche en cours…');
      geocoderUn(q).then(c => {
        if (c) {
          geo[q] = { lat: c.lat, lon: c.lon };
          toucher(); rendre(); rendreMarqueurs();
          toast(c.via && c.via !== q ? `Placée via « ${c.via} » — vérifiez la position.` : 'Position recalculée.');
        }
        else {
          rendreMarqueurs();
          afficherTrace();
          toast('Aucune formulation n\'a abouti — voir le détail sous la carte.');
        }
      }).catch(e => { afficherTrace(); toast('Recherche impossible : ' + e.message); });
      return;
    }
    if (t.dataset.relance) {
      items = items.map(it => it.id === t.dataset.relance
        ? Object.assign({}, it, { relances: it.relances.concat(today()), etat: it.etat === 'En attente' ? 'Relancé' : it.etat })
        : it);
      toucher(); rendre(); toast('Relance notée.');
    } else if (t.dataset.edit) ouvrirFiche(t.dataset.edit);
    else if (t.dataset.filtre) { filtre = t.dataset.filtre; rendre(); }
    else if (t.dataset.rmrel) { brouillonRelances.splice(+t.dataset.rmrel, 1); rendreRelances(); }
  });

  $('#btn-recharger').addEventListener('click', () => {
    if (modifie && !confirm('Des modifications ne sont pas encore envoyées sur kDrive. Recharger la page les perdra. Continuer ?')) return;
    location.reload();
  });
  $('#btn-theme').addEventListener('click', () => {
    theme = theme === 'dark' ? 'light' : 'dark';
    GM_setValue(THEME_KEY, theme);
    appliquerTheme();
    if (carte) rendreMarqueurs();
  });
  $('#q').addEventListener('input', e => { recherche = e.target.value; rendre(); });
  $('#btn-ajouter').addEventListener('click', () => onglet === 'prospects' ? ouvrirProspect(null) : ouvrirFiche(null));
  $('#p-annuler').addEventListener('click', () => { $('#ov-prospect').classList.remove('open'); editProspect = null; });
  $('#p-valider').addEventListener('click', validerProspect);
  $('#p-suppr').addEventListener('click', () => {
    const p = prospects.find(x => x.id === editProspect);
    if (!p || !confirm(`Retirer ${p.entreprise} de la liste à postuler ?`)) return;
    prospects = prospects.filter(x => x.id !== editProspect);
    toucher(); $('#ov-prospect').classList.remove('open'); editProspect = null; rendre();
  });
  $('#btn-charger').addEventListener('click', () => {
    if (modifie && !confirm('Des modifications ne sont pas encore envoyées sur kDrive. Recharger les écrasera. Continuer ?')) return;
    charger();
  });
  $('#btn-enregistrer').addEventListener('click', () => enregistrer());
  $('#bar-save').addEventListener('click', () => enregistrer(conflit === true));
  $('#bar-recharger').addEventListener('click', () => {
    if (modifie && !confirm('Vos modifications non enregistrées seront perdues. Recharger quand même ?')) return;
    charger();
  });
  $('#btn-carte').addEventListener('click', ouvrirCarte);
  $('#m-fermer').addEventListener('click', () => $('#ov-carte').classList.remove('open'));
  $('#m-noms').addEventListener('change', basculerNoms);
  $('#m-geo').addEventListener('click', geocoderTout);
  $('#btn-historique').addEventListener('click', ouvrirHistorique);
  $('#h-fermer').addEventListener('click', () => $('#ov-hist').classList.remove('open'));
  $('#btn-nettoyer').addEventListener('click', ouvrirNettoyage);
  $('#n-annuler').addEventListener('click', () => $('#ov-net').classList.remove('open'));
  $('#n-appliquer').addEventListener('click', appliquerNettoyage);
  $('#btn-reglages').addEventListener('click', ouvrirReglages);
  $('#f-ville').addEventListener('change', () => basculerVille('f'));
  $('#p-ville').addEventListener('change', () => basculerVille('p'));
  $('#f-annuler').addEventListener('click', fermerFiche);
  $('#f-valider').addEventListener('click', validerFiche);
  $('#f-suppr').addEventListener('click', () => {
    const it = items.find(x => x.id === editId);
    if (!it || !confirm(`Supprimer la candidature chez ${it.entreprise} ?`)) return;
    items = items.filter(x => x.id !== editId);
    toucher(); fermerFiche(); rendre();
  });
  $('#c-annuler').addEventListener('click', () => $('#ov-reglages').classList.remove('open'));
  $('#c-diag').addEventListener('click', () => { lireReglages(); $('#diag').textContent = 'Test en cours…'; diagnostic(); });
  $('#c-sonde').addEventListener('click', () => { lireReglages(); $('#diag').textContent = 'Sonde en cours…'; sonder(); });
  $('#c-valider').addEventListener('click', () => {
    lireReglages(); saveCfg();
    $('#ov-reglages').classList.remove('open');
    charger();
  });

  /* ---- coquille mobile ----
     La barre du bas (onglets Candidatures / À postuler) réutilise la délégation data-onglet ;
     seuls Carte, le bouton flottant, le menu « Plus » et le tap-sur-carte sont câblés ici.
     Chaque action délègue aux fonctions existantes : aucune logique dupliquée. */
  const menuOv = $('#cb-menu-ov');
  const fermerMenu = () => menuOv.classList.remove('open');
  $('#nav-carte').addEventListener('click', ouvrirCarte);
  $('#nav-plus').addEventListener('click', () => menuOv.classList.add('open'));
  $('#cb-fab').addEventListener('click', () => onglet === 'prospects' ? ouvrirProspect(null) : ouvrirFiche(null));
  menuOv.addEventListener('click', e => {
    if (e.target === menuOv) { fermerMenu(); return; }        // tap sur le voile
    const b = e.target.closest('[data-act]');
    if (!b) return;
    fermerMenu();
    switch (b.dataset.act) {
      case 'save': enregistrer(conflit === true); break;
      case 'charger':
        if (modifie && !confirm('Des modifications ne sont pas encore envoyées sur kDrive. Recharger les écrasera. Continuer ?')) return;
        charger(); break;
      case 'historique': ouvrirHistorique(); break;
      case 'nettoyer': ouvrirNettoyage(); break;
      case 'reglages': ouvrirReglages(); break;
      case 'theme': $('#btn-theme').click(); break;
      case 'recharger': $('#btn-recharger').click(); break;
    }
  });
  // Sur mobile, taper une carte (hors lien ou bouton) ouvre la fiche.
  $('#list').addEventListener('click', e => {
    if (window.innerWidth > 620 || e.target.closest('a,button')) return;
    const c = e.target.closest('.cb-card');
    if (!c) return;
    if (c.dataset.id) ouvrirFiche(c.dataset.id);
    else if (c.dataset.idp) ouvrirProspect(c.dataset.idp);
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') document.querySelectorAll('.cb-ov.open,.cb-menu-ov.open').forEach(o => o.classList.remove('open'));
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && $('#ov-fiche').classList.contains('open')) validerFiche();
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && $('#ov-prospect').classList.contains('open')) validerProspect();
    if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); enregistrer(); }
  });
  window.addEventListener('beforeunload', e => { if (modifie) { e.preventDefault(); e.returnValue = ''; } });

  if (typeof GM_registerMenuCommand === 'function') GM_registerMenuCommand('Réglages kDrive', ouvrirReglages);

  rendre();
  if (cfg.token) charger(true); else ouvrirReglages();

  setInterval(() => { if (!document.hidden) verifier(); }, INTERVALLE);
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden && Date.now() - dernierControle > 60000) verifier();
  });
})();
