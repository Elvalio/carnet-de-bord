# Mise en place — à faire une seule fois, puis supprimer ce fichier

Ce dossier vient d'être extrait d'une archive. Il n'est pas encore un dépôt git et n'est publié nulle part. Voici ce qu'il reste à faire.

## Contexte utile

Compte GitHub : `Elvalio`. Le projet vit actuellement dans un gist (`f9e30da71f5c4f58c2d3d36b907b229d`) dont les URL sont encore déclarées en `@updateURL` et `@downloadURL` dans l'en-tête du userscript. On migre vers un dépôt nommé `carnet-de-bord`.

Le dépôt doit être **public**. `raw.githubusercontent.com` exige un token d'authentification sur un dépôt privé, et Violentmonkey ne sait pas en fournir : la mise à jour automatique cesserait de fonctionner. Le code ne contient aucun secret, seuls des identifiants de drive inertes sans le token API, qui lui reste dans le stockage local de l'extension.

## Étapes

1. Initialiser le dépôt sur la branche `main`, premier commit avec l'ensemble des fichiers.
2. Créer le dépôt distant public et pousser. Utiliser `gh repo create carnet-de-bord --public --source=. --push` si l'outil `gh` est disponible et authentifié ; sinon demander à l'utilisateur de créer le dépôt à la main et ajouter le remote.
3. Remplacer dans `carnet-de-bord.user.js` les deux occurrences de l'URL du gist par `https://raw.githubusercontent.com/Elvalio/carnet-de-bord/main/carnet-de-bord.user.js`, dans les lignes `@updateURL` et `@downloadURL`.
4. Incrémenter `@version` — sans cela Violentmonkey ne proposera pas la mise à jour.
5. Lancer `./verifier.sh`, qui contrôle la syntaxe et vérifie que la version a bien changé.
6. Commit et push.
7. Signaler à l'utilisateur qu'il doit réinstaller le script **une fois sur chaque appareil** depuis la nouvelle URL brute, l'ancienne pointant toujours sur le gist. Son token et ses réglages ne seront pas perdus.
8. Supprimer ce fichier et le commiter, la mise en place étant terminée.

## Ensuite

Lire `CLAUDE.md` avant toute modification : il contient les contraintes de l'API kDrive découvertes empiriquement, la structure du classeur, et la discipline à tenir entre les deux systèmes d'exploitation utilisés.
