# Carnet de bord — suivi de candidatures en alternance

> Si `SETUP.md` est présent à la racine, la mise en place initiale n'a pas encore été faite : le lire et l'exécuter en premier.

## Ce que c'est

Userscript Violentmonkey en un seul fichier (`carnet-de-bord.user.js`), sans build, sans dépendances installées localement. Il lit et réécrit un classeur Excel stocké sur kDrive Infomaniak, qui tient lieu de base de données. Utilisé sur Firefox desktop et Firefox Android, à la même adresse et avec les mêmes données.

Points d'entrée : `https://internsearch.com` (domaine parqué, l'app remplace le contenu de la page ; la racine redirige vers `/lander`, sans conséquence) et `https://example.com/#carnet` en secours. Le fragment `#carnet` n'est exigé que sur le second.

Dépendances chargées à l'exécution via `@require` et `@resource` : SheetJS 0.18.5 pour le format xlsx, Leaflet 1.9.4 pour la carte.

## Contraintes établies empiriquement — ne pas « corriger » sans test

**CORS.** L'API kDrive ne renvoie aucun en-tête `Access-Control-Allow-Origin`. Une page web ordinaire ne peut donc pas l'appeler, quel que soit le code. Tous les accès passent obligatoirement par `GM_xmlhttpRequest`, qui s'exécute dans le contexte de l'extension et échappe à cette règle. Ne jamais remplacer ces appels par `fetch`. Tout domaine contacté doit figurer en `@connect`.

**Upload kDrive.** `POST /3/drive/{drive}/upload?file_id={file}&total_size={octets}`, corps binaire en `application/octet-stream`. Les paramètres `conflict` et `file_id` s'excluent mutuellement : les fournir ensemble renvoie HTTP 422 avec `validation_rule_prohibits`. Les valeurs acceptées de `conflict` sont `rename`, `version`, `error` — pas `replace`. Le paramètre `last_modified_at` déclenchait également un 422 et a été retiré. Ces trois points ont été trouvés en interrogeant l'API réelle, la documentation publique étant illisible pour un agent.

**Verrou OnlyOffice.** Si le fichier est ouvert dans l'éditeur de kDrive, l'upload renvoie HTTP 403 avec `lock_error`. Le cas est intercepté et traduit en message clair.

**Nominatim.** Une requête par seconde au maximum, en-tête `User-Agent` identifiant obligatoire sous peine de 403. Le géocodage dérive plusieurs formulations d'une même adresse et les essaie dans l'ordre, de la plus complète à la plus large : suppression du préfixe précédant le premier mot de voirie (les codes de site internes aux entreprises n'existent dans aucune base cartographique), insertion d'une virgule avant le code postal, repli sur le code postal seul. Chaque tentative est isolée dans son propre `try` — sans cela, une erreur réseau sur la première interrompt toute la cascade, ce qui a été un vrai bug. Les villes seules partent en requête structurée `city=&country=France`, sinon une commune peut être confondue avec un homonyme ailleurs. Les résultats sont mis en cache dans le classeur et jamais redemandés.

**Aucun secret dans le code.** Le token API vit dans le stockage de Violentmonkey via `GM_setValue`. Il ne doit jamais apparaître dans le fichier, ni partir ailleurs que vers `api.infomaniak.com` : les appels à Nominatim passent par `requeteLibre`, une fonction distincte qui n'ajoute pas l'en-tête d'autorisation. C'est ce qui permet de publier le code sur un dépôt public.

## Structure du classeur

Cinq feuilles, toutes réécrites intégralement à chaque enregistrement.

- Feuille principale (nom repris du fichier existant, ne pas l'imposer) : `Sujet`, `Premier envoi`, `Etat`, `Nom entreprise`, `Localisation`, `Adresse`, `Poste`, `Source`, `Notes externes`, `Relances`. La colonne de date contient de vraies dates au format `dd/mm/yyyy`.
- `A postuler` : entreprises repérées mais pas encore contactées. `Entreprise`, `Poste visé`, `Localisation`, `Adresse`, `Source`, `Repérée le`, `Notes`.
- `Journal` : historique des modifications, une ligne par champ modifié, plafonné à 600 entrées, plus récent en premier.
- `Villes` : localisations connues, alimentant les menus déroulants. Une ville y reste même si plus aucune candidature ne l'utilise, pour permettre de préparer des cibles.
- `Geo` : cache de géocodage, `Requête` vers latitude et longitude. Contient aussi les positions corrigées à la main depuis la carte.

La réécriture perd les listes déroulantes et la mise en forme conditionnelle d'origine. Aucune bibliothèque JS ne sait les préserver ; c'est un compromis accepté.

## Concurrence entre appareils

Le fichier est édité depuis un PC et un téléphone. Le motif retenu : mémoriser `last_modified_at` au chargement, le revérifier toutes les cinq minutes, au retour sur l'onglet, et juste avant chaque écriture. En cas de divergence sans modification locale, rechargement silencieux. Avec modifications locales, la barre du bas passe en rouge et l'utilisateur choisit entre recharger et écraser. Aucune fusion automatique — ce serait fragile et opaque.

## Carte

Regroupement des points calculé en distance d'écran et non en distance réelle, recalculé à chaque changement de zoom : c'est ce qui fait que les pastilles se séparent naturellement en zoomant. Le seuil doit rester supérieur au diamètre d'une pastille, sinon elles se chevauchent — erreur commise puis corrigée. La fusion est itérative jusqu'à stabilisation, en recalculant le centre à chaque étape ; une passe simple laissait des paquets voisins non fusionnés. Les candidatures partageant des coordonnées identiques ne pourront jamais être séparées par le zoom : leur pastille se déplie sur place au clic. Un panneau latéral liste toutes les candidatures, placées ou non, pour qu'aucune ne devienne inaccessible faute de point sur la carte.

## Mobile

L'interface est utilisée quotidiennement sur Firefox Android, écran d'environ 380 pixels. Règles retenues : aucune bande défilante horizontale pour un contenu essentiel — l'utilisateur ne devine pas qu'une zone coupée au bord de l'écran défile ; les compteurs passent en grille, tout visible d'un coup. Aucune cible tactile sous 38 pixels de haut. Champs de saisie à 16 pixels au minimum, sinon Firefox zoome automatiquement au focus et casse la mise en page. Les fenêtres deviennent des feuilles pleine hauteur, titre collé en haut et actions collées en bas. Les encoches sont prises en compte via `env(safe-area-inset-bottom)`.

**Coquille native mobile (v4.0).** Sous 620 px, la navigation passe en barre fixe collée en bas — `.cb-nav` avec quatre entrées : Candidatures, À postuler (les deux portent `data-onglet` et réutilisent donc la délégation de clic existante), Carte, et Plus. Un bouton flottant `.cb-fab` reprend la logique de `#btn-ajouter` (candidature ou prospect selon l'onglet). Les actions rares — Enregistrer, Charger, Historique, Nettoyer, Réglages, Thème, Recharger — sont regroupées dans un menu-feuille `.cb-menu-ov` ouvert par « Plus » ; chaque entrée délègue aux fonctions ou boutons existants, **aucune logique n'est dupliquée**. L'ancien chrome (barre `.cb-sync`, onglets `.cb-onglets`, outils `.cb-outils-head`, bouton `+` en ligne) est masqué en `!important` dans le `@media` — le desktop est strictement inchangé. L'état actif de la barre du bas et ses badges de comptage sont synchronisés dans `rendre()` (`.cb-nav-i[data-onglet]`, `#nb-cand`, `#nb-prosp`). Les cartes portent `data-id`/`data-idp` et un écouteur sur `#list` ouvre la fiche au tap n'importe où sur la carte, sauf sur un lien ou un bouton (`closest('a,button')`) et seulement sous 620 px. La barre d'enregistrement flottante et le toast sont remontés au-dessus de la barre de navigation. Contrainte à tenir : toute nouvelle action mobile doit **déléguer** au câblage existant, jamais le réimplémenter.

**Rendu « cockpit » et ajout unifié (v4.1).** Sous 620 px, les cartes passent en lignes denses (`grid-template-columns:1fr auto`) : nom + poste à gauche, métrique de jours et jauge à droite. Les notes sont masquées et le bouton « modifier » aussi (`.cb-side button[data-edit]{display:none}`), le tap sur la carte ouvrant déjà la fiche ; « + relance » reste. C'est un choix mobile-only, le desktop garde ses cartes complètes. Le bouton d'ajout ne dépend plus de l'onglet : le « + » (flottant sur mobile, `#btn-ajouter` sur desktop) ouvre un choix `#ov-ajout` « Candidature / À postuler » (`ouvrirChoixAjout`, feuille montante sur mobile, carte centrée sur desktop) — indispensable puisque le sélecteur peut être masqué. Justement : **quand `prospects.length === 0`, le sélecteur Candidatures / À postuler disparaît** (desktop *et* mobile) et la vue est ramenée sur les candidatures — géré dans `rendre()` : `.cb-onglets` en `display:none` inline, `.cb-nav` reçoit la classe `sans-prospects` (trois colonnes, entrée « À postuler » masquée), et `onglet` est forcé à `candidatures`. Le sélecteur réapparaît dès qu'un prospect existe. Le choix visuel « cockpit » vient de la piste 03 d'une exploration de maquettes ; l'accent reste `--mauve` pour la cohérence avec le desktop (la maquette montrait du bleu — swap trivial si un jour souhaité).

## Publier une version

Incrémenter `@version` dans l'en-tête, sinon Violentmonkey ne proposera jamais la mise à jour. Vérifier la syntaxe avec `node --check carnet-de-bord.user.js`. Puis commit et push : les appareils se mettent à jour seuls depuis l'URL déclarée en `@updateURL`. Le dépôt doit rester public, `raw.githubusercontent.com` exigeant un token sur un dépôt privé, ce que Violentmonkey ne sait pas fournir.

## Travail sur deux systèmes

Le projet est développé sur la même machine, alternativement sous Linux et sous Windows. L'historique des conversations de Claude Code est stocké localement dans `~/.claude/` et ne suit pas d'un système à l'autre : `/resume` ne retrouvera rien de ce qui a été fait sous l'autre. Il en va de même des réglages et des permissions accordées.

Conséquence pratique, à respecter systématiquement : toute décision, contrainte découverte ou piste écartée qui compte pour la suite doit être écrite dans ce fichier au moment où elle est prise, et non laissée dans le fil de conversation. Quand une session aboutit à un choix structurant, l'ajouter ici avant de terminer.

Et faire `git pull` en début de session, `git push` en fin de session, sans exception. Deux versions divergentes du même fichier unique produiraient un conflit pénible à résoudre à la main.

## Conventions

Interface, commentaires et messages en français. Ton sobre, sans emphase ni jargon. Le code reste volontairement en un seul fichier, sans framework ni étape de build : tout découpage en modules casserait la distribution par userscript, qui est la raison d'être du projet.
