# Carnet de bord

Suivi de candidatures en alternance, sous forme d'un userscript unique. Les données vivent dans un classeur Excel sur kDrive Infomaniak, que l'application lit et réécrit directement. Aucun serveur, aucune base de données, aucun hébergement.

## Ce que ça fait

Liste des candidatures avec état, compteur de jours sans réponse et relances. Bandeau des candidatures à relancer au-delà de douze jours. Onglet « À postuler » pour les entreprises repérées mais pas encore contactées, convertibles en candidature en un clic. Carte des candidatures géocodées, avec regroupement automatique des points proches et correction manuelle des positions. Historique des modifications. Nettoyage automatique des libellés. Thème clair et sombre.

Le classeur reste exploitable tel quel dans OnlyOffice ou Excel.

## Installation

1. Installer [Violentmonkey](https://violentmonkey.github.io/) sur Firefox, desktop et Android.
2. Ouvrir l'URL brute du fichier `carnet-de-bord.user.js` de ce dépôt. Violentmonkey propose l'installation.
3. Créer un token API dans le Manager Infomaniak : profil, section Développeur, Tokens API. Scopes `drive` et `user_info`, durée illimitée.
4. Ouvrir l'application et coller le token dans les réglages.

Le token reste dans le stockage local de Violentmonkey, sur chaque appareil. Il n'est jamais écrit dans le code ni transmis ailleurs qu'à l'API Infomaniak.

## Configuration

Les identifiants du fichier se lisent dans l'URL kDrive :

```
ksuite.infomaniak.com/{compte}/kdrive/app/drive/{drive_id}/files/{directory_id}/preview/spreadsheet/{file_id}
```

Ils sont pré-remplis dans le code et modifiables depuis les réglages. La fenêtre des réglages contient également un diagnostic, qui teste la lecture, et une sonde qui vérifie les formes d'appel de l'écriture sur un fichier jetable.

## Cap ingé

Deuxième userscript du dépôt, `cap-inge.user.js`, indépendant du carnet : une page de recherche d'écoles d'ingénieur en alternance et d'entreprises robotique, avec son propre classeur kDrive (`Cap ingé.xlsx`, feuilles `Écoles` et `Entreprises`). Shortlist, filtre « proche de Bruz », ajout, modification et suppression depuis la page.

1. Ouvrir l'URL brute de `cap-inge.user.js` pour l'installer dans Violentmonkey.
2. Ouvrir `https://example.com/#cap`.
3. Dans les réglages, coller un token API (scope `drive`), vérifier le `drive_id` et le `directory_id` du dossier, puis « Créer le fichier » : il est déposé sur kDrive avec la liste de départ, et son `file_id` est rempli automatiquement.
4. Sur un autre appareil, saisir le même token et ce `file_id`, puis « Enregistrer et charger ».

## Points de vigilance

Fermer l'éditeur OnlyOffice de kDrive avant d'enregistrer, faute de quoi le fichier est verrouillé.

Charger avant d'éditer quand on change d'appareil. L'application vérifie périodiquement si le fichier a changé et prévient en cas de conflit, mais ne fusionne rien.

Rien n'est envoyé automatiquement : une barre apparaît en bas de l'écran tant que des modifications n'ont pas été enregistrées.
