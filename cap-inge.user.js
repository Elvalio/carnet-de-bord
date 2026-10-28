// ==UserScript==
// @name         Cap ingé — écoles et entreprises (kDrive)
// @namespace    harry.cap-inge
// @version      1.0
// @description  Écoles d'ingé en alternance et entreprises robotique, lues et réécrites dans un classeur sur kDrive
// @author       Harry
// @updateURL    https://raw.githubusercontent.com/Elvalio/carnet-de-bord/main/cap-inge.user.js
// @downloadURL  https://raw.githubusercontent.com/Elvalio/carnet-de-bord/main/cap-inge.user.js
// @match        https://example.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @connect      api.infomaniak.com
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @run-at       document-idle
// ==/UserScript==

/*
  Ouvrir  https://example.com/#cap
  Page indépendante du carnet de bord : son propre classeur kDrive, ses propres réglages.
  Token API Infomaniak : scope drive. Il reste dans le stockage local de Violentmonkey.
  Le fichier kDrive ne doit pas être ouvert dans OnlyOffice au moment d'enregistrer.
*/

(function () {
  'use strict';
  // Le carnet de bord répond à #carnet sur ce même domaine : chacun son fragment.
  if (!location.hash.startsWith('#cap')) return;

  const XHR = typeof GM_xmlhttpRequest === 'function' ? GM_xmlhttpRequest : GM.xmlHttpRequest;

  /* ================= configuration ================= */
  const CFG_KEY = 'cap.config';
  const CACHE_KEY = 'cap.cache';
  // Même kDrive et même dossier que le carnet ; file_id vide tant que le fichier n'est pas créé.
  const DEFAULTS = { token: '', driveId: '2606717', fileId: '', parentId: '289', fileName: 'Cap ingé.xlsx' };
  let cfg = Object.assign({}, DEFAULTS, JSON.parse(GM_getValue(CFG_KEY, '{}')));
  const saveCfg = () => GM_setValue(CFG_KEY, JSON.stringify(cfg));

  /* ================= modèle ================= */
  // Une vue par feuille du classeur. L'ordre des champs est celui des colonnes.
  const GEII = { ok: 'Accepté', check: 'À confirmer', non: 'Non accepté' };
  const PROCHE = { k: 'proche', col: 'Proche de Bruz', w: 15, type: 'oui', libelle: 'Proche de Bruz (moins d\'1 h 30)' };
  const SHORTLIST = { k: 'shortlist', col: 'Shortlist', w: 10, type: 'oui', libelle: 'Dans ma shortlist ★' };
  const VUES = {
    ecoles: {
      feuille: 'Écoles', nouveau: 'Nouvelle école', modifier: 'Modifier l\'école', ajouter: '+ Ajouter une école',
      champs: [
        { k: 'nom', col: 'École', w: 26, plein: true, ph: 'ECAM Rennes' },
        { k: 'formation', col: 'Formation', w: 50, plein: true, ph: 'Mécatronique · ITII Bretagne' },
        { k: 'lieu', col: 'Lieu', w: 22, ph: 'Bruz (35)' },
        { k: 'places', col: 'Places', w: 16, ph: '30 (2026)' },
        { k: 'geii', col: 'BUT GEII', w: 14, type: 'geii' },
        { k: 'site', col: 'Site', w: 34, type: 'url' },
        { k: 'interet', col: 'Ce qui t\'intéresse', w: 50, type: 'long' },
        { k: 'savoir', col: 'À savoir', w: 60, type: 'long' },
        PROCHE, SHORTLIST
      ]
    },
    entreprises: {
      feuille: 'Entreprises', nouveau: 'Nouvelle entreprise', modifier: 'Modifier l\'entreprise', ajouter: '+ Ajouter une entreprise',
      champs: [
        { k: 'nom', col: 'Entreprise', w: 26, plein: true, ph: 'BA Systèmes' },
        { k: 'activite', col: 'Activité', w: 40, plein: true, ph: 'Robots mobiles, AGV' },
        { k: 'lieu', col: 'Lieu', w: 22, ph: 'Mordelles (35)' },
        { k: 'taille', col: 'Taille', w: 20, ph: 'PME, startup, labo…' },
        { k: 'site', col: 'Site', w: 34, type: 'url', plein: true },
        { k: 'pourquoi', col: 'Pourquoi elle', w: 70, type: 'long' },
        PROCHE, SHORTLIST
      ]
    }
  };
  const vierge = vue => Object.fromEntries(VUES[vue].champs.map(c =>
    [c.k, c.type === 'oui' ? false : c.type === 'geii' ? 'check' : '']));

  let donnees = { ecoles: [], entreprises: [] };
  let autresFeuilles = [];   // feuilles ajoutées à la main dans OnlyOffice : recopiées telles quelles
  let vue = 'ecoles', fNear = false, fStar = false;
  let editId = null, editVue = 'ecoles';
  let charge = false, modifie = false, envoiEnCours = false;
  let versionDistante = null, conflit = false, controleEnCours = false, dernierControle = 0;
  const INTERVALLE = 5 * 60 * 1000;

  /* ================= liste de départ ================= */
  // Recherche initiale, vérifiée le 27 oct. 2026. Versée dans le classeur à sa création.
  const DEPART = {
    ecoles: [
      { nom: 'ECAM Rennes', formation: 'Systèmes numériques pour l\'innovation industrielle · ITII Bretagne', lieu: 'Bruz (35)', proche: true, geii: 'ok', places: '80 (2026)',
        interet: 'Numérique, IA, cobots, outils industriels', savoir: 'Privée (EESPIG), frais pris en charge par le contrat d\'apprentissage. Sur ton campus.', site: 'https://www.ecam-rennes.fr' },
      { nom: 'ENSIBS', formation: 'Mécatronique · ITII Bretagne', lieu: 'Lorient (56)', proche: true, geii: 'ok', places: '30 (2026)',
        interet: 'Mécanique + électronique + info, BE et R&D', savoir: 'Publique. Embauche moyenne 35,7 k€ brut. 500 entreprises partenaires (Naval Group, Thales…).', site: 'https://www.ensibs.fr' },
      { nom: 'INSA Hauts-de-France', formation: 'Génie électrique et informatique industrielle (FISA)', lieu: 'Valenciennes (59)', geii: 'ok',
        interet: 'Axe 4A sur atelier reconfigurable : robotique mobile et collaborative', savoir: 'La continuité la plus directe du BUT GEII. 12 semaines min. à l\'étranger.', site: 'https://www.insa-hautsdefrance.fr' },
      { nom: 'Polytech Dijon', formation: 'Robotique, parcours Cobotique (FISA)', lieu: 'Le Creusot (71)', geii: 'ok',
        interet: 'Robotique industrielle, cobotique, vision', savoir: 'Pôle d\'excellence robotique et vision. Un dossier par spécialité, admission confirmée à la signature du contrat.', site: 'https://polytech.ube.fr' },
      { nom: 'UPSSITECH (Toulouse III)', formation: 'Robotique, parcours Systèmes robotiques et interactifs', lieu: 'Toulouse (31)', geii: 'check',
        interet: 'Robotique + IA, robotique logicielle, startups', savoir: 'Ouvert en alternance depuis sept. 2026 : toute nouvelle filière, peu de recul.', site: 'https://www.univ-tlse3.fr' },
      { nom: 'Cnam × ITII Île-de-France', formation: 'Automatique et robotique (apprentissage)', lieu: 'Île-de-France', geii: 'check',
        interet: 'Automatique, robotique', savoir: 'Existe aussi avec l\'ITII Picardie (Hauts-de-France). Bassin parisien = startups robotique.', site: 'https://www.cnam.fr' },
      { nom: 'INSA Strasbourg', formation: 'Mécatronique · ITII Alsace', lieu: 'Strasbourg (67)', geii: 'check',
        interet: 'Mécatronique, conception à industrialisation', savoir: 'Embauche moyenne 43,1 k€ brut (2025). Parcours franco-allemand possible.', site: 'https://www.insa-strasbourg.fr' },
      { nom: 'Arts et Métiers Châlons', formation: 'Mécanique et robotique (apprentissage)', lieu: 'Châlons-en-Ch. (51)', geii: 'check',
        interet: 'Architecture de systèmes robotiques, lignes automatisées, intégration au SI de production',
        savoir: 'Rythme 4 sem. entreprise / 5 sem. école. Dossier + entretien + tests maths et anglais. 12 sem. à l\'étranger en 3e année. Contact : scolarite-fis.chalons@ensam.eu', site: 'https://lise.ensam.eu' },
      { nom: 'Arts et Métiers', formation: 'Mécanique et mécatronique (apprentissage)', lieu: 'Metz (57) · St-Étienne (42)', geii: 'check', places: '24 à Metz (2025)',
        interet: 'Machines spéciales, systèmes mécatroniques et robotiques', savoir: 'Priorité aux BUT GMP/GIM, autres BUT acceptés. Embauche moyenne 40,2 k€ à Metz.', site: 'https://artsetmetiers.fr' },
      { nom: 'Polytech Annecy-Chambéry', formation: 'Mécanique, parcours Conception & mécatronique (FISA)', lieu: 'Annecy (74)', geii: 'check',
        interet: 'Conception, intégration mécatronique', savoir: 'Plus orienté mécanique que robotique pure.' },
      { nom: 'ENSIL-ENSCI', formation: 'Mécatronique (filière apprentissage)', lieu: 'Limoges (87)', geii: 'check',
        interet: 'Mécatronique, commande, robotique', savoir: 'Filière ouverte en 2024. Même diplôme que la voie étudiante.', site: 'https://www.ensil-ensci.unilim.fr' }
    ],
    entreprises: [
      { nom: 'BA Systèmes', activite: 'Robots mobiles, AGV, cobot mobile', lieu: 'Mordelles (35)', proche: true, taille: 'PME industrielle',
        pourquoi: 'Intralogistique (Nestlé, Airbus…), médical, portuaire. Forte culture R&D et lien historique avec l\'IUT de Rennes.' },
      { nom: 'Haption', activite: 'Retour d\'effort, téléopération', lieu: 'Soulgé-sur-Ouette (53)', proche: true, taille: 'PME deeptech',
        pourquoi: 'Interfaces haptiques pour nucléaire, médical, auto et aéro. Croise robotique et VR.', site: 'https://www.haption.com' },
      { nom: 'Inria / IRISA – équipe Rainbow', activite: 'Labo de recherche robotique', lieu: 'Rennes (35)', proche: true, taille: 'Labo public',
        pourquoi: 'Asservissement visuel, multi-robots, contrôle partagé. Alternance ingé à demander directement.', site: 'https://team.inria.fr/rainbow/' },
      { nom: 'Sepro Robotique', activite: 'Robots pour presses à injection', lieu: 'La Roche-sur-Yon (85)', taille: 'ETI ~650 pers.',
        pourquoi: 'Robots cartésiens et 6 axes, cellules d\'automatisation, 85 % à l\'export.', site: 'https://www.sepro-group.com' },
      { nom: 'Wandercraft', activite: 'Exosquelettes + humanoïde Calvin', lieu: 'Paris', taille: 'Scale-up',
        pourquoi: 'Renault au capital et premier client de Calvin-40. Marche bipède, contrôle, méca.' },
      { nom: 'Enchanted Tools', activite: 'Robot social Mirokaï', lieu: 'Paris', taille: 'Startup',
        pourquoi: 'Robot d\'assistance pour l\'hôpital, partenariat NVIDIA. Premières livraisons visées fin 2026.', site: 'https://enchanted.tools' },
      { nom: 'Inbolt', activite: 'Vision 3D + IA pour bras industriels', lieu: 'Paris', taille: 'Startup',
        pourquoi: '11 M€ levés en sept. 2026. Clients Stellantis, GM, Toyota. Compatible FANUC, KUKA, ABB, UR.', site: 'https://www.inbolt.com' },
      { nom: 'Isybot', activite: 'Robotique collaborative', lieu: 'Voir site', taille: 'Startup',
        pourquoi: '7 M€ levés en fév. 2026. Cobots d\'assemblage pour l\'auto et l\'aéro.' },
      { nom: 'Exotec', activite: 'Robots d\'entrepôt Skypod', lieu: 'Croix (59)', taille: 'Licorne industrielle',
        pourquoi: 'Plus de 10 000 robots déployés. Publie des offres d\'alternance (mise en service, achats…).', site: 'https://www.exotec.com' },
      { nom: 'Pollen Robotics (Hugging Face)', activite: 'Humanoïdes open source Reachy', lieu: 'Bordeaux (33)', taille: 'Équipe de ~20–30',
        pourquoi: 'Rachetée par Hugging Face en 2025. Reachy 2 et Reachy Mini, très orienté IA incarnée.' },
      { nom: 'Naïo', activite: 'Robots agricoles autonomes', lieu: 'Escalquens (31)', taille: 'PME reprise',
        pourquoi: 'Repris fin 2025, 24 salariés, partenariat Kioti. Navigation autonome, embarqué.', site: 'https://www.naio-technologies.com' },
      { nom: 'Nomad Robotics', activite: 'Intégration robotique, industrie 4.0', lieu: 'Laval (53)', proche: true, taille: 'Jeune PME',
        pourquoi: 'Prend déjà des alternants (côté commercial). Candidature spontanée côté technique.' }
    ]
  };
  const listeDeDepart = () => ({
    ecoles: DEPART.ecoles.map(r => Object.assign({ id: uid() }, vierge('ecoles'), r)),
    entreprises: DEPART.entreprises.map(r => Object.assign({ id: uid() }, vierge('entreprises'), r))
  });

  /* ================= utilitaires ================= */
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const uid = () => 'r' + Math.random().toString(36).slice(2, 9);
  const sansAccents = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const estOui = s => { const n = sansAccents(s); return !!n && !['non', 'no', '0', 'faux', 'false'].includes(n); };
  function versGeii(s) {
    const n = sansAccents(s);
    if (n.includes('non') || n.includes('refus')) return 'non';
    if (n.includes('accept') || n === 'ok' || n === 'oui') return 'ok';
    return 'check';
  }
  // Seuls les liens http(s) deviennent cliquables : le classeur peut contenir n'importe quoi.
  const lienValide = s => /^https?:\/\//i.test(String(s ?? '').trim());
  function normUrl(s) {
    s = String(s ?? '').trim();
    return !s || lienValide(s) ? s : 'https://' + s.replace(/^\/+/, '');
  }

  /* ================= API kDrive ================= */
  // Toujours GM_xmlhttpRequest : l'API kDrive ne renvoie aucun en-tête CORS.
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

  const UPLOAD = () => `https://api.infomaniak.com/3/drive/${cfg.driveId}/upload`;

  async function envoyer(url, buf) {
    const r = await api({
      method: 'POST', url,
      headers: { 'Content-Type': 'application/octet-stream' },
      data: new Blob([buf], { type: 'application/octet-stream' })
    });
    let corps = null;
    try { corps = JSON.parse(r.responseText); } catch (e) { /* réponse non JSON */ }
    const ok = r.status >= 200 && r.status < 300 && corps?.result !== 'error';
    return { ok, status: r.status, corps, brut: r.responseText || '' };
  }

  // Réécriture : file_id seul. Y ajouter conflict renvoie HTTP 422.
  async function ecrireFichier(buf) {
    const r = await envoyer(`${UPLOAD()}?file_id=${encodeURIComponent(cfg.fileId)}&total_size=${buf.byteLength}`, buf);
    if (r.ok) return;
    if (r.corps?.error?.code === 'lock_error') {
      throw new Error('Fichier verrouillé : il est ouvert dans l\'éditeur OnlyOffice de kDrive. Ferme l\'onglet, attends une minute, puis réessaie. Tes modifications sont conservées ici.');
    }
    throw new Error(`Envoi refusé (HTTP ${r.status}) — ${r.brut.slice(0, 600)}`);
  }

  // Création : dossier + nom. conflict=error pour ne jamais écraser un fichier du même nom.
  async function creerFichierDistant(buf) {
    const r = await envoyer(`${UPLOAD()}?directory_id=${encodeURIComponent(cfg.parentId)}`
      + `&file_name=${encodeURIComponent(cfg.fileName)}&total_size=${buf.byteLength}&conflict=error`, buf);
    const id = r.corps?.data?.id;
    if (r.ok && id) return String(id);
    const deja = r.status === 409 || /conflict|already/i.test(r.brut);
    throw new Error((deja ? `Un fichier « ${cfg.fileName} » existe déjà dans ce dossier : colle son file_id, ou change le nom. ` : '')
      + `Création refusée (HTTP ${r.status}) — ${r.brut.slice(0, 400)}`);
  }

  /* ================= xlsx ================= */
  function parser(buf) {
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
    const lire = vue => {
      const v = VUES[vue], ws = wb.Sheets[v.feuille];
      if (!ws) return [];
      return XLSX.utils.sheet_to_json(ws, { defval: '', raw: false })
        .filter(l => v.champs.some(c => String(l[c.col] ?? '').trim()))
        .map(l => {
          const r = { id: uid() };
          v.champs.forEach(c => {
            const s = String(l[c.col] ?? '').trim();
            r[c.k] = c.type === 'oui' ? estOui(s) : c.type === 'geii' ? versGeii(s) : s;
          });
          return r;
        });
    };
    const nos = Object.values(VUES).map(v => v.feuille);
    return {
      donnees: { ecoles: lire('ecoles'), entreprises: lire('entreprises') },
      autres: wb.SheetNames.filter(n => !nos.includes(n) && wb.Sheets[n]['!ref']).map(n => [n, wb.Sheets[n]])
    };
  }

  function feuille(vue, liste) {
    const v = VUES[vue];
    const aoa = [v.champs.map(c => c.col)];
    liste.forEach(r => aoa.push(v.champs.map(c =>
      c.type === 'oui' ? (r[c.k] ? 'oui' : '')
        : c.type === 'geii' ? (GEII[r[c.k]] || GEII.check)
        : (r[c.k] ?? ''))));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = v.champs.map(c => ({ wch: c.w }));
    return ws;
  }

  function construire(d) {
    const wb = XLSX.utils.book_new();
    Object.keys(VUES).forEach(vue => XLSX.utils.book_append_sheet(wb, feuille(vue, d[vue]), VUES[vue].feuille));
    autresFeuilles.forEach(([nom, ws]) => XLSX.utils.book_append_sheet(wb, ws, nom));
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  }

  /* ================= actions ================= */
  const memoriser = () => GM_setValue(CACHE_KEY, JSON.stringify({ donnees, at: Date.now() }));
  function toucher() { modifie = true; memoriser(); }

  async function charger(silencieux) {
    if (!cfg.token || !cfg.fileId) { ouvrirReglages(); return; }
    majEtat('Lecture du fichier sur kDrive…');
    try {
      const r = parser(await lireFichier());
      donnees = r.donnees; autresFeuilles = r.autres;
      charge = true; modifie = false; conflit = false;
      try { versionDistante = await metaFichier(); dernierControle = Date.now(); } catch (e) { /* sans gravité */ }
      memoriser(); rendre();
      if (!silencieux) toast(`${donnees.ecoles.length} écoles et ${donnees.entreprises.length} entreprises chargées.`);
    } catch (e) {
      charge = false;
      const cache = JSON.parse(GM_getValue(CACHE_KEY, 'null'));
      const restaure = !!(cache && cache.donnees) && !donnees.ecoles.length && !donnees.entreprises.length;
      if (restaure) donnees = cache.donnees;
      rendre();
      toast('Lecture impossible : ' + e.message + (restaure ? ' — copie locale affichée.' : ''));
    }
  }

  async function enregistrer(force) {
    if (!charge) { toast('Charge d\'abord le fichier depuis kDrive.'); return; }
    if (envoiEnCours) return;
    envoiEnCours = true;
    try {
      if (force !== true) {
        try {
          const d = await metaFichier();
          dernierControle = Date.now();
          if (d && versionDistante && d > versionDistante) {
            conflit = true; rendre();
            if (!confirm('Le fichier a été modifié sur kDrive depuis ton chargement, sans doute depuis un autre appareil.\n\nEnregistrer maintenant écrasera ces changements. Continuer ?')) return;
          }
        } catch (e) { /* si la vérification échoue, on tente l'envoi quand même */ }
      }
      majEtat('Envoi vers kDrive…');
      await ecrireFichier(construire(donnees));
      modifie = false; conflit = false;
      try { versionDistante = await metaFichier(); dernierControle = Date.now(); } catch (e) { /* sans gravité */ }
      rendre();
      toast('Enregistré sur kDrive.');
    } catch (e) {
      rendre();
      toast(e.message);
    } finally { envoiEnCours = false; }
  }

  // Contrôle périodique : rien n'est téléchargé tant que la date de modification n'a pas bougé.
  // Reporté si une fenêtre est ouverte, pour ne pas recharger sous une saisie en cours.
  async function verifier() {
    if (!charge || !cfg.token || controleEnCours || document.querySelector('.ov.open')) return;
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
          toast('Liste actualisée depuis kDrive.');
        }
      } else rendre();
    } catch (e) { /* hors ligne ou API indisponible : on réessaiera */ }
    finally { controleEnCours = false; }
  }

  // Ce qui est à l'écran part dans le nouveau fichier ; s'il n'y a rien, la liste de départ.
  async function creerFichier() {
    lireReglages();
    if (!cfg.token || !cfg.driveId || !cfg.parentId || !cfg.fileName) {
      info('Il faut le token, le drive_id, le directory_id et un nom de fichier.'); return;
    }
    if (cfg.fileId && !confirm(`Un fichier est déjà relié (file_id ${cfg.fileId}). En créer un nouveau à la place ?`)) return;
    const contenu = donnees.ecoles.length || donnees.entreprises.length ? donnees : listeDeDepart();
    info('Création du fichier sur kDrive…');
    try {
      cfg.fileId = await creerFichierDistant(construire(contenu));
      saveCfg();
      $('#c-file').value = cfg.fileId;
      donnees = contenu;
      charge = true; modifie = false; conflit = false;
      try { versionDistante = await metaFichier(); dernierControle = Date.now(); } catch (e) { /* sans gravité */ }
      memoriser(); rendre();
      info(`Fichier créé, file_id ${cfg.fileId}. Sur un autre appareil, colle ce numéro dans file_id.`);
    } catch (e) { info(e.message); }
  }

  /* ================= styles ================= */
  const CSS = `
  :root{
    --bg:#f5f7f9;--surface:#ffffff;--ink:#15202b;--muted:#5a6977;--line:#d9e0e6;
    --accent:#1f5fbf;--accent-soft:#e3ecfa;--near:#f2b705;--near-soft:#fff5d1;--warn:#b4540a;
    --danger:#b42318;--voile:rgba(21,32,43,.45);--ombre:rgba(21,32,43,.16);
    --f-display:"Archivo","Arial Narrow",system-ui,sans-serif;
    --f-body:"IBM Plex Sans",system-ui,-apple-system,"Segoe UI",sans-serif;
    --f-mono:"IBM Plex Mono",ui-monospace,Menlo,monospace;
    color-scheme:light;
  }
  @media (prefers-color-scheme:dark){:root{
    --bg:#0f151b;--surface:#16202a;--ink:#e6edf3;--muted:#93a3b2;--line:#2a3846;
    --accent:#7aa8f0;--accent-soft:#1c2c44;--near:#f2c94c;--near-soft:#3a3214;--warn:#f0a35e;
    --danger:#f97066;--voile:rgba(5,8,12,.72);--ombre:rgba(0,0,0,.5);color-scheme:dark;}}
  *,*::before,*::after{box-sizing:border-box;}
  html,body{margin:0!important;padding:0!important;background:var(--bg)!important;}
  body{color:var(--ink);font-family:var(--f-body);font-size:15px;line-height:1.5;-webkit-font-smoothing:antialiased;}
  .wrap{max-width:1180px;margin:0 auto;padding-block:32px 120px;padding-left:max(16px,env(safe-area-inset-left,0px));
  padding-right:max(16px,env(safe-area-inset-right,0px));display:flex;flex-direction:column;gap:24px;}
  header{display:flex;flex-direction:column;gap:8px;}
  .eyebrow{font-family:var(--f-mono);font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);}
  h1{font-family:var(--f-display);font-weight:800;font-size:clamp(28px,4.4vw,42px);line-height:1.05;margin:0;text-wrap:balance;letter-spacing:-.01em;}
  .lede{max-width:68ch;color:var(--muted);margin:0;}
  .facts{display:flex;flex-wrap:wrap;gap:8px 20px;font-family:var(--f-mono);font-size:13px;color:var(--muted);}
  .facts b{color:var(--ink);font-weight:500;}
  .sync{display:flex;flex-wrap:wrap;gap:8px;align-items:center;background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:8px 10px 8px 14px;}
  .sync span{font-family:var(--f-mono);font-size:13px;color:var(--muted);margin-right:auto;min-width:0;}
  .timeline{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:16px 18px;display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;}
  .timeline div{min-width:0;display:flex;flex-direction:column;gap:2px;}
  .timeline .when{font-family:var(--f-mono);font-size:12px;color:var(--accent);letter-spacing:.04em;text-transform:uppercase;}
  .timeline .what{font-size:14px;}
  .bar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between;}
  .tabs{display:flex;gap:4px;background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:4px;}
  .tabs button{font:600 14px var(--f-body);border:0;background:transparent;color:var(--muted);padding:8px 14px;border-radius:6px;cursor:pointer;}
  .tabs button[aria-selected="true"]{background:var(--accent);color:var(--surface);}
  .tabs button span{font:500 12px var(--f-mono);margin-left:4px;opacity:.8;}
  .filters{display:flex;flex-wrap:wrap;gap:8px;align-items:center;}
  .chip{font:500 13px var(--f-body);border:1px solid var(--line);background:var(--surface);color:var(--ink);padding:6px 12px;border-radius:999px;cursor:pointer;}
  .chip[aria-pressed="true"]{border-color:var(--accent);background:var(--accent-soft);color:var(--accent);}
  .btn{font:600 14px var(--f-body);border:1px solid var(--line);background:var(--surface);color:var(--ink);padding:8px 14px;border-radius:8px;cursor:pointer;min-height:38px;}
  .btn:hover{border-color:var(--muted);}
  .btn.p{background:var(--accent);border-color:var(--accent);color:var(--surface);}
  .btn.p:hover{filter:brightness(1.08);}
  .btn.g{background:transparent;color:var(--muted);}
  .btn.g:hover{color:var(--ink);}
  .btn.d{background:transparent;color:var(--danger);}
  a:focus-visible,button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible{outline:2px solid var(--accent);outline-offset:2px;}
  .scroller{overflow-x:auto;background:var(--surface);border:1px solid var(--line);border-radius:10px;}
  table{border-collapse:collapse;width:100%;min-width:900px;font-size:14px;}
  th{font:600 12px var(--f-mono);letter-spacing:.06em;text-transform:uppercase;color:var(--muted);text-align:left;padding:12px 14px;border-bottom:1px solid var(--line);background:var(--surface);}
  td{padding:14px;border-bottom:1px solid var(--line);vertical-align:top;}
  tr:last-child td{border-bottom:0;}
  tr.near td:first-child{box-shadow:inset 4px 0 0 var(--near);}
  tbody tr[data-id]{cursor:pointer;}
  tbody tr[data-id]:hover td{background:var(--bg);}
  td.c-star,td.c-edit{width:1%;white-space:nowrap;}
  .c-long{white-space:pre-line;}
  .name{font-family:var(--f-display);font-weight:700;font-size:16px;line-height:1.2;}
  .sub{color:var(--muted);font-size:13px;margin-top:2px;}
  .loc{font-family:var(--f-mono);font-size:13px;white-space:nowrap;}
  .tag{display:inline-block;font:500 11px var(--f-mono);padding:2px 7px;border-radius:4px;margin:4px 4px 0 0;letter-spacing:.03em;}
  .tag.near{background:var(--near-soft);color:var(--ink);}
  .tag.check{background:transparent;border:1px solid var(--warn);color:var(--warn);}
  .tag.ok{background:var(--accent-soft);color:var(--accent);}
  .tag.non{background:transparent;border:1px solid var(--line);color:var(--muted);}
  .num{font-family:var(--f-mono);font-variant-numeric:tabular-nums;font-size:13px;}
  a{color:var(--accent);}
  .star{border:1px solid var(--line);background:transparent;color:var(--muted);width:34px;height:34px;border-radius:6px;cursor:pointer;font-size:16px;line-height:1;}
  .star[aria-pressed="true"]{background:var(--near-soft);border-color:var(--near);color:var(--ink);}
  .mod{font:500 12px var(--f-mono);background:transparent;border:1px solid transparent;color:var(--muted);border-radius:6px;padding:5px 8px;cursor:pointer;}
  tbody tr:hover .mod{border-color:var(--line);color:var(--ink);}
  .vide-ligne td{color:var(--muted);text-align:center;padding:32px 16px;}
  .vide-ligne .btn{margin-top:12px;}
  .notes{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;}
  .note{min-width:0;}
  .note h2{font-family:var(--f-display);font-size:18px;margin:0 0 6px;}
  .note p,.note ul{margin:0;color:var(--muted);font-size:14px;}
  .note ul{padding-left:18px;display:flex;flex-direction:column;gap:4px;}
  footer{font-size:12px;color:var(--muted);font-family:var(--f-mono);}

  .barre{position:fixed;left:50%;bottom:calc(16px + env(safe-area-inset-bottom));transform:translate(-50%,160%);visibility:hidden;
  display:flex;align-items:center;gap:12px;background:var(--surface);border:1px solid var(--near);border-radius:12px;
  padding:10px 12px 10px 16px;box-shadow:0 8px 28px var(--ombre);z-index:50;width:max-content;max-width:calc(100vw - 24px);
  transition:transform .2s ease,visibility .2s;}
  .barre.on{transform:translate(-50%,0);visibility:visible;}
  .barre.conflit{border-color:var(--danger);}
  .barre p{margin:0;font-size:13.5px;color:var(--muted);min-width:0;}
  .barre p b{display:block;color:var(--ink);font-size:14px;}

  .ov{position:fixed;inset:0;background:var(--voile);display:none;align-items:flex-start;justify-content:center;
  padding:40px 16px;overflow-y:auto;z-index:100;overscroll-behavior:contain;}
  .ov.open{display:flex;}
  .sheet{background:var(--surface);border:1px solid var(--line);border-radius:12px;width:100%;max-width:620px;padding:20px 22px;box-shadow:0 24px 60px var(--ombre);}
  .sheet h2{font-family:var(--f-display);font-weight:700;font-size:21px;margin:0 0 16px;}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
  .f{display:flex;flex-direction:column;gap:5px;min-width:0;}
  .w{grid-column:1/-1;}
  .f label{font:500 11px var(--f-mono);letter-spacing:.06em;text-transform:uppercase;color:var(--muted);}
  .f input,.f select,.f textarea{font:15px var(--f-body);color:var(--ink);background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:9px 11px;width:100%;min-width:0;}
  .f textarea{resize:vertical;min-height:76px;}
  .case{display:flex;align-items:center;gap:10px;min-height:42px;font-size:14px;cursor:pointer;}
  .case input{width:20px;height:20px;margin:0;accent-color:var(--accent);}
  .aide{margin:0;font-size:13px;color:var(--muted);}
  .aide b{font-family:var(--f-mono);font-weight:500;color:var(--ink);}
  .info{margin:0;font-family:var(--f-mono);font-size:12px;color:var(--ink);white-space:pre-wrap;word-break:break-word;
  background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:9px 11px;}
  .info:empty{display:none;}
  .version{margin:14px 0 0;font-family:var(--f-mono);font-size:12px;color:var(--muted);}
  .actions{display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;margin-top:18px;}
  .actions .d,.actions .gauche{margin-right:auto;}

  .toast{position:fixed;left:50%;bottom:calc(88px + env(safe-area-inset-bottom));transform:translate(-50%,10px);
  background:var(--ink);color:var(--surface);padding:9px 16px;border-radius:12px;font-size:14px;opacity:0;pointer-events:none;
  transition:opacity .18s,transform .18s;z-index:200;width:max-content;max-width:calc(100vw - 24px);text-align:center;}
  .toast.show{opacity:1;transform:translate(-50%,0);}

  @media (max-width:620px){
  .wrap{padding-block:20px calc(130px + env(safe-area-inset-bottom));gap:18px;}
  *{-webkit-tap-highlight-color:transparent;}
  .sync span{flex:1 1 100%;margin:0 0 2px;}
  .sync .btn{flex:1;}
  .tabs{width:100%;}
  .tabs button{flex:1;min-height:42px;}
  .filters{width:100%;}
  .chip{flex:1;min-height:40px;}
  #btn-ajouter{flex:1 1 100%;min-height:46px;font-size:15px;}
  .btn{min-height:44px;}

  /* Tableau : sous 620 px, chaque ligne devient une carte. Rien ne déborde hors de l'écran. */
  .scroller{overflow:visible;background:none;border:0;border-radius:0;}
  #tbl,#tbl tbody{display:block;min-width:0;}
  #tbl thead{display:none;}
  #tbl tbody{display:flex;flex-direction:column;gap:10px;}
  #tbl tr{display:flex;flex-wrap:wrap;align-items:center;gap:6px 10px;background:var(--surface);border:1px solid var(--line);
  border-radius:10px;padding:12px 12px 12px 16px;}
  #tbl td{display:block;padding:0;border:0;background:none!important;}
  #tbl tr.near{box-shadow:inset 4px 0 0 var(--near);}
  #tbl tr.near td:first-child{box-shadow:none;}
  .c-nom{order:1;flex:1 1 calc(100% - 60px);min-width:0;}
  .c-star{order:2;flex:0 0 auto;width:auto!important;}
  .c-lieu,.c-geii,.c-places,.c-taille{order:3;flex:0 0 auto;}
  .c-lieu{display:flex!important;align-items:center;gap:6px;}
  .c-lieu br{display:none;}
  .c-lieu .tag,.c-geii .tag{margin:0;}
  .c-places::before{content:"Places : ";color:var(--muted);}
  .c-taille{font-size:13px;margin:0;}
  .c-long{order:4;flex:1 1 100%;font-size:14px;}
  .c-long::before{content:attr(data-l);display:block;font:500 11px var(--f-mono);letter-spacing:.06em;text-transform:uppercase;color:var(--muted);margin-bottom:1px;}
  .c-places.vide,.c-taille.vide,.c-long.vide,.c-edit{display:none!important;}
  .star{width:42px;height:42px;font-size:18px;}
  .vide-ligne{justify-content:center;}
  .vide-ligne td{flex:1 1 100%;}

  /* Fenêtres : feuilles pleine hauteur, titre et actions collés. */
  .ov{padding:0;align-items:stretch;}
  .sheet{max-width:none;border-radius:0;border:0;min-height:100%;padding:16px 14px calc(16px + env(safe-area-inset-bottom));display:flex;flex-direction:column;}
  .sheet h2{position:sticky;top:0;background:var(--surface);padding:4px 0 10px;margin:0 0 10px;z-index:2;font-size:19px;}
  .grid{grid-template-columns:1fr;flex:1;align-content:start;}
  .f input,.f select,.f textarea{font-size:16px;min-height:44px;}
  .f textarea{min-height:96px;}
  .actions{position:sticky;bottom:0;background:var(--surface);border-top:1px solid var(--line);padding:10px 0 0;margin-top:14px;}
  .actions .btn{flex:1;min-width:0;}
  .actions .d,.actions .gauche{flex:0 0 100%;margin-right:0;order:5;}

  .barre{left:10px;right:10px;width:auto;max-width:none;transform:translateY(160%);flex-wrap:wrap;}
  .barre.on{transform:none;}
  .barre p{flex:1 1 100%;}
  .barre .btn{flex:1;}
  .toast{left:12px;right:12px;width:auto;max-width:none;bottom:calc(140px + env(safe-area-inset-bottom));transform:translateY(10px);}
  .toast.show{transform:none;}
  }
  @media (prefers-reduced-motion:reduce){*{transition:none!important;}}`;

  /* ================= structure ================= */
  const HTML = `
  <div class="wrap">
    <header>
      <span class="eyebrow">Après le BUT GEII · cycle ingénieur 3 ans en apprentissage</span>
      <h1>Écoles d'ingé en alternance &amp; entreprises robotique</h1>
      <p class="lede">Formations sous statut apprenti (FISA) orientées robotique, mécatronique et automatique, puis des entreprises de R&amp;D robotique où chercher le contrat. Les lignes surlignées en jaune sont à moins d'une heure et demie de Bruz.</p>
      <div class="facts"><span>Rentrée visée <b>sept. 2027</b></span><span>Admission définitive <b>= contrat signé</b></span></div>
    </header>

    <div class="sync">
      <span id="etat"></span>
      <button class="btn g" id="btn-charger">Charger</button>
      <button class="btn g" id="btn-reglages">Réglages</button>
    </div>

    <section class="timeline" aria-label="Calendrier type">
      <div><span class="when">Nov → janv.</span><span class="what">Shortlist d'écoles, CV et lettre orientés robotique, premières candidatures spontanées en entreprise.</span></div>
      <div><span class="when">Janv. → avril</span><span class="what">Dossiers école (réseau Polytech : candidatures entre janvier et avril), oraux.</span></div>
      <div><span class="when">Mars → août</span><span class="what">Entretiens entreprises, signature du contrat d'apprentissage de 3 ans.</span></div>
      <div><span class="when">Sept. 2027</span><span class="what">Entrée en 1re année du cycle ingénieur.</span></div>
    </section>

    <div class="bar">
      <div class="tabs" role="tablist">
        <button role="tab" data-vue="ecoles" aria-selected="true">Écoles<span id="n-ecoles"></span></button>
        <button role="tab" data-vue="entreprises" aria-selected="false">Entreprises<span id="n-entreprises"></span></button>
      </div>
      <div class="filters">
        <button class="chip" id="f-near" aria-pressed="false">Proche de Bruz</button>
        <button class="chip" id="f-star" aria-pressed="false">Ma shortlist ★</button>
        <button class="btn p" id="btn-ajouter">+ Ajouter</button>
      </div>
    </div>

    <div class="scroller"><table id="tbl"><thead></thead><tbody></tbody></table></div>

    <section class="notes">
      <div class="note">
        <h2>Comment lire la colonne GEII</h2>
        <p>« Accepté » veut dire que la fiche officielle cite le BUT GEII. « À confirmer » veut dire que la fiche parle de BUT mécanique ou « autres BUT acceptés » : un mail au responsable de formation suffit pour trancher. « Non accepté » : la réponse a été non.</p>
      </div>
      <div class="note">
        <h2>Trouver l'entreprise</h2>
        <ul>
          <li>Les écoles ITII (ECAM, ENSIBS, INSA Strasbourg) ont des réseaux d'entreprises partenaires : demande-leur la liste.</li>
          <li>Pour les startups, la candidature spontanée marche mieux que l'offre publiée : vise le responsable technique.</li>
          <li>La Bonne Alternance (gouv) liste aussi les entreprises qui ont déjà pris des alternants.</li>
        </ul>
      </div>
      <div class="note">
        <h2>À éviter</h2>
        <p>Aldebaran (NAO, Pepper) a été liquidé. Naïo a été repris fin 2025 avec une petite équipe : intéressant mais peu de places.</p>
      </div>
    </section>

    <footer>Liste de départ vérifiée le 27 oct. 2026. Sources : Onisep, France Compétences (RNCP), sites des écoles, presse économique (Maddyness, Usine Digitale, The Robot Report). Les places et salaires sont ceux publiés pour 2025–2026.</footer>
  </div>

  <div class="barre" id="barre">
    <p><b id="barre-titre"></b><span id="barre-texte"></span></p>
    <button class="btn g" id="barre-recharger">Recharger</button>
    <button class="btn p" id="barre-save">Enregistrer sur kDrive</button>
  </div>

  <div class="ov" id="ov-fiche"><div class="sheet" role="dialog" aria-modal="true" aria-labelledby="fiche-titre">
    <h2 id="fiche-titre"></h2>
    <div class="grid" id="fiche-champs"></div>
    <div class="actions">
      <button class="btn d" id="f-suppr">Supprimer</button>
      <button class="btn g" id="f-annuler">Annuler</button>
      <button class="btn p" id="f-valider">Enregistrer</button>
    </div>
  </div></div>

  <div class="ov" id="ov-reglages"><div class="sheet" role="dialog" aria-modal="true" aria-labelledby="reglages-titre">
    <h2 id="reglages-titre">Réglages kDrive</h2>
    <div class="grid">
      <div class="f w"><label for="c-token">Token API Infomaniak (scope drive)</label><input id="c-token" type="password" placeholder="colle le token" autocomplete="off"></div>
      <div class="f"><label for="c-drive">drive_id</label><input id="c-drive" inputmode="numeric"></div>
      <div class="f"><label for="c-file">file_id</label><input id="c-file" inputmode="numeric" placeholder="vide si pas encore créé"></div>
      <div class="f"><label for="c-parent">directory_id (dossier)</label><input id="c-parent" inputmode="numeric"></div>
      <div class="f"><label for="c-nom">Nom du fichier</label><input id="c-nom"></div>
      <p class="aide w">Les numéros se lisent dans l'adresse kDrive : …/drive/<b>drive_id</b>/files/<b>directory_id</b>/preview/spreadsheet/<b>file_id</b>. Pas encore de fichier ? Renseigne le dossier : « Créer le fichier » l'y dépose avec la liste de départ.</p>
      <p class="info w" id="c-info"></p>
    </div>
    <p class="version" id="app-version"></p>
    <div class="actions">
      <button class="btn g gauche" id="c-creer">Créer le fichier</button>
      <button class="btn g" id="c-fermer">Fermer</button>
      <button class="btn p" id="c-valider">Enregistrer et charger</button>
    </div>
  </div></div>

  <div class="toast" id="toast" role="status"></div>`;

  document.documentElement.lang = 'fr';
  document.head.innerHTML = '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">'
    + '<link rel="preconnect" href="https://fonts.googleapis.com">'
    + '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@500;700;800&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">';
  document.title = 'Cap ingé';
  const style = document.createElement('style');
  style.textContent = CSS;
  document.head.appendChild(style);
  document.body.innerHTML = HTML;
  const $ = s => document.querySelector(s);
  const ouvrir = s => $(s).classList.add('open');
  const fermer = s => $(s).classList.remove('open');

  // Version lue depuis l'en-tête via GM_info, jamais recopiée à la main.
  try {
    const v = (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version) || '';
    $('#app-version').textContent = v ? 'Cap ingé — version ' + v : '';
  } catch (e) { /* GM_info indisponible : la ligne reste vide */ }

  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove('show'), 4500);
  }
  const majEtat = txt => { $('#etat').textContent = txt; };
  const info = txt => { $('#c-info').textContent = txt; };

  /* ================= rendu ================= */
  const etoile = r => `<button class="star" data-star="${r.id}" aria-pressed="${!!r.shortlist}" aria-label="Shortlist : ${esc(r.nom)}">${r.shortlist ? '★' : '☆'}</button>`;
  function nomLie(r) {
    const n = esc(r.nom) || 'Sans nom';
    return lienValide(r.site) ? `<a href="${esc(r.site.trim())}" target="_blank" rel="noopener">${n}</a>` : n;
  }
  const celluleLieu = r => `<td class="c-lieu"><span class="loc">${esc(r.lieu) || '—'}</span>${r.proche ? '<br><span class="tag near">Proche</span>' : ''}</td>`;
  const celluleLongue = (texte, libelle, cls) => `<td class="c-long${cls ? ' ' + cls : ''}${texte ? '' : ' vide'}" data-l="${esc(libelle)}">${esc(texte)}</td>`;
  const celluleModif = r => `<td class="c-edit"><button class="mod" aria-label="Modifier ${esc(r.nom)}">modifier</button></td>`;

  const TABLES = {
    ecoles: {
      entetes: ['', 'Formation', 'Lieu', 'BUT GEII', 'Places', 'Ce qui t\'intéresse', 'À savoir', ''],
      ligne: r => `<td class="c-star">${etoile(r)}</td>
        <td class="c-nom"><div class="name">${nomLie(r)}</div>${r.formation ? `<div class="sub">${esc(r.formation)}</div>` : ''}</td>
        ${celluleLieu(r)}
        <td class="c-geii"><span class="tag ${r.geii}">${GEII[r.geii]}</span></td>
        <td class="c-places num${r.places ? '' : ' vide'}">${esc(r.places) || '—'}</td>
        ${celluleLongue(r.interet, 'Ce qui t\'intéresse')}${celluleLongue(r.savoir, 'À savoir', 'sub')}${celluleModif(r)}`
    },
    entreprises: {
      entetes: ['', 'Entreprise', 'Lieu', 'Taille', 'Pourquoi elle', ''],
      ligne: r => `<td class="c-star">${etoile(r)}</td>
        <td class="c-nom"><div class="name">${nomLie(r)}</div>${r.activite ? `<div class="sub">${esc(r.activite)}</div>` : ''}</td>
        ${celluleLieu(r)}
        <td class="c-taille sub${r.taille ? '' : ' vide'}">${esc(r.taille)}</td>
        ${celluleLongue(r.pourquoi, 'Pourquoi elle')}${celluleModif(r)}`
    }
  };

  function messageVide(tout) {
    if (!charge && !tout.length) return cfg.fileId ? 'Fichier non chargé : utilise « Charger ».' : 'Aucun fichier relié : ouvre les Réglages pour le créer ou le relier.';
    if (!donnees.ecoles.length && !donnees.entreprises.length) {
      return `Le fichier est vide.<br><button class="btn p" data-depart>Ajouter la liste de départ (${DEPART.ecoles.length} écoles, ${DEPART.entreprises.length} entreprises)</button>`;
    }
    if (!tout.length) return vue === 'ecoles' ? 'Aucune école pour l\'instant.' : 'Aucune entreprise pour l\'instant.';
    return 'Rien à afficher. Ajoute des lignes à ta shortlist avec ☆ ou retire un filtre.';
  }

  function rendre() {
    const heure = dernierControle ? new Date(dernierControle).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '';
    majEtat(charge
      ? `${cfg.fileName} · ${modifie ? 'modifications non enregistrées' : 'à jour' + (heure ? ', vérifié à ' + heure : '')}`
      : (cfg.fileId ? 'Fichier non chargé.' : 'Aucun fichier relié.'));

    const barre = $('#barre');
    barre.classList.toggle('on', modifie || conflit);
    barre.classList.toggle('conflit', conflit);
    $('#barre-recharger').style.display = conflit ? '' : 'none';
    $('#barre-titre').textContent = conflit ? 'Le fichier a changé sur kDrive' : 'Modifications non enregistrées';
    $('#barre-texte').textContent = conflit
      ? (modifie ? 'Un autre appareil l\'a modifié. Enregistrer écrasera ces changements.' : 'Recharge pour récupérer la version à jour.')
      : 'Elles seront perdues si tu fermes l\'onglet.';
    $('#barre-save').textContent = conflit ? 'Enregistrer quand même' : 'Enregistrer sur kDrive';

    document.querySelectorAll('.tabs [data-vue]').forEach(b => b.setAttribute('aria-selected', b.dataset.vue === vue));
    $('#n-ecoles').textContent = donnees.ecoles.length;
    $('#n-entreprises').textContent = donnees.entreprises.length;
    $('#f-near').setAttribute('aria-pressed', fNear);
    $('#f-star').setAttribute('aria-pressed', fStar);
    $('#btn-ajouter').textContent = VUES[vue].ajouter;

    const t = TABLES[vue], tout = donnees[vue];
    const liste = tout.filter(r => (!fNear || r.proche) && (!fStar || r.shortlist));
    $('#tbl thead').innerHTML = '<tr>' + t.entetes.map(h => `<th>${esc(h)}</th>`).join('') + '</tr>';
    $('#tbl tbody').innerHTML = liste.length
      ? liste.map(r => `<tr class="${r.proche ? 'near' : ''}" data-id="${r.id}">${t.ligne(r)}</tr>`).join('')
      : `<tr class="vide-ligne"><td colspan="${t.entetes.length}">${messageVide(tout)}</td></tr>`;
  }

  /* ================= fiche ================= */
  function ouvrirFiche(id) {
    const v = VUES[vue];
    const r = id ? donnees[vue].find(x => x.id === id) : null;
    editId = r ? r.id : null; editVue = vue;
    $('#fiche-titre').textContent = r ? v.modifier : v.nouveau;
    $('#fiche-champs').innerHTML = v.champs.map(c => {
      const cid = 'f-' + c.k;
      if (c.type === 'oui') return `<label class="case"><input type="checkbox" id="${cid}"> ${esc(c.libelle)}</label>`;
      const ph = c.ph ? ` placeholder="${esc(c.ph)}"` : '';
      const champ = c.type === 'long' ? `<textarea id="${cid}"${ph}></textarea>`
        : c.type === 'geii' ? `<select id="${cid}">${Object.entries(GEII).map(([k, l]) => `<option value="${k}">${esc(l)}</option>`).join('')}</select>`
        : c.type === 'url' ? `<input id="${cid}" type="url" inputmode="url" placeholder="https://">`
        : `<input id="${cid}"${ph}>`;
      return `<div class="f${c.plein || c.type === 'long' ? ' w' : ''}"><label for="${cid}">${esc(c.col)}</label>${champ}</div>`;
    }).join('');
    const base = r || vierge(vue);
    v.champs.forEach(c => {
      const el = $('#f-' + c.k);
      if (c.type === 'oui') el.checked = !!base[c.k]; else el.value = base[c.k] ?? '';
    });
    $('#f-suppr').style.display = r ? '' : 'none';
    ouvrir('#ov-fiche');
    if (!r) $('#f-nom').focus();
  }

  function validerFiche() {
    const v = VUES[editVue];
    const data = {};
    v.champs.forEach(c => {
      const el = $('#f-' + c.k);
      data[c.k] = c.type === 'oui' ? el.checked
        : c.type === 'url' ? normUrl(el.value)
        : c.type === 'long' ? el.value.trim()
        : el.value.replace(/\s+/g, ' ').trim();
    });
    if (!data.nom) { $('#f-nom').focus(); toast('Il manque le nom.'); return; }
    if (editId) donnees[editVue] = donnees[editVue].map(r => r.id === editId ? Object.assign({}, r, data) : r);
    else {
      donnees[editVue] = [Object.assign({ id: uid() }, data)].concat(donnees[editVue]);
      if ((fNear && !data.proche) || (fStar && !data.shortlist)) toast('Ajoutée, mais masquée par un filtre actif.');
    }
    toucher(); fermer('#ov-fiche'); rendre();
  }

  function basculerShortlist(id) {
    donnees[vue] = donnees[vue].map(r => r.id === id ? Object.assign({}, r, { shortlist: !r.shortlist }) : r);
    toucher(); rendre();
  }

  /* ================= réglages ================= */
  function ouvrirReglages() {
    $('#c-token').value = cfg.token;
    $('#c-drive').value = cfg.driveId;
    $('#c-file').value = cfg.fileId;
    $('#c-parent').value = cfg.parentId;
    $('#c-nom').value = cfg.fileName;
    info('');
    ouvrir('#ov-reglages');
  }
  function lireReglages() {
    cfg.token = $('#c-token').value.trim();
    cfg.driveId = $('#c-drive').value.trim();
    cfg.fileId = $('#c-file').value.trim();
    cfg.parentId = $('#c-parent').value.trim();
    cfg.fileName = $('#c-nom').value.trim();
  }

  /* ================= événements ================= */
  document.querySelector('.tabs').addEventListener('click', e => {
    const b = e.target.closest('[data-vue]');
    if (b) { vue = b.dataset.vue; rendre(); }
  });
  $('#f-near').addEventListener('click', () => { fNear = !fNear; rendre(); });
  $('#f-star').addEventListener('click', () => { fStar = !fStar; rendre(); });
  $('#btn-ajouter').addEventListener('click', () => ouvrirFiche(null));
  // Un clic sur une ligne ouvre la fiche, sauf sur l'étoile ou un lien.
  $('#tbl').addEventListener('click', e => {
    if (e.target.closest('[data-depart]')) {
      donnees = listeDeDepart(); toucher(); rendre();
      toast('Liste de départ ajoutée — pense à enregistrer.');
      return;
    }
    const s = e.target.closest('[data-star]');
    if (s) { basculerShortlist(s.dataset.star); return; }
    if (e.target.closest('a')) return;
    const tr = e.target.closest('tr[data-id]');
    if (tr) ouvrirFiche(tr.dataset.id);
  });

  $('#f-annuler').addEventListener('click', () => fermer('#ov-fiche'));
  $('#f-valider').addEventListener('click', validerFiche);
  $('#f-suppr').addEventListener('click', () => {
    const r = donnees[editVue].find(x => x.id === editId);
    if (!r || !confirm(`Supprimer ${r.nom} ?`)) return;
    donnees[editVue] = donnees[editVue].filter(x => x.id !== editId);
    toucher(); fermer('#ov-fiche'); rendre();
  });

  $('#btn-charger').addEventListener('click', () => {
    if (modifie && !confirm('Des modifications ne sont pas encore envoyées sur kDrive. Recharger les écrasera. Continuer ?')) return;
    charger();
  });
  $('#btn-reglages').addEventListener('click', ouvrirReglages);
  $('#barre-save').addEventListener('click', () => enregistrer(conflit === true));
  $('#barre-recharger').addEventListener('click', () => {
    if (modifie && !confirm('Tes modifications non enregistrées seront perdues. Recharger quand même ?')) return;
    charger();
  });

  $('#c-fermer').addEventListener('click', () => fermer('#ov-reglages'));
  $('#c-creer').addEventListener('click', creerFichier);
  $('#c-valider').addEventListener('click', () => {
    lireReglages();
    if (!cfg.fileId) { info('Renseigne le file_id, ou crée le fichier avec « Créer le fichier ».'); return; }
    if (modifie && !confirm('Des modifications ne sont pas encore envoyées sur kDrive. Recharger les écrasera. Continuer ?')) return;
    saveCfg(); fermer('#ov-reglages'); charger();
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') document.querySelectorAll('.ov.open').forEach(o => o.classList.remove('open'));
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && $('#ov-fiche').classList.contains('open')) validerFiche();
    if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); enregistrer(); }
  });
  window.addEventListener('beforeunload', e => { if (modifie) { e.preventDefault(); e.returnValue = ''; } });

  if (typeof GM_registerMenuCommand === 'function') GM_registerMenuCommand('Réglages kDrive', ouvrirReglages);

  rendre();
  if (cfg.token && cfg.fileId) charger(true); else ouvrirReglages();

  setInterval(() => { if (!document.hidden) verifier(); }, INTERVALLE);
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden && Date.now() - dernierControle > 60000) verifier();
  });
})();
