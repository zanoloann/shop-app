# V2.1 — Double authentification (TOTP)

## Ce qui change par rapport à V2

Le « code unique » utilisé pour enrôler un nouvel appareil (première connexion, ou après révocation) n'est plus un code stocké dans le Google Sheet et affiché via un écran « Réglages » de l'appli. C'est désormais un vrai **TOTP** (Time-based One-Time Password) : le même principe que la double authentification d'un compte Google, compatible avec Google Authenticator, Authy, etc.

**Pourquoi ce changement :**
- Un admin n'a plus jamais besoin d'ouvrir le Google Sheet ou l'éditeur Apps Script pour donner un code à un artisan.
- Le code à 6 chiffres change automatiquement toutes les 30 secondes et se lit directement dans l'appli d'authentification, sur le téléphone de l'admin — aucun écran « Afficher le code » dans nos deux applis (supprimé).
- Comme le secret est le même sur les 3 téléphones admin, ils affichent tous le même code au même moment : n'importe lequel des 3 peut enrôler un artisan, sans coordination entre eux.
- Le token d'appareil (`deviceToken`) et la révocation via suppression de ligne dans l'onglet « Accès » ne changent pas.

## Ce que tu dois configurer de ton côté (une seule fois)

1. **Redéployer le script** : colle les 4 fichiers `.gs` de ce dossier dans ton projet Apps Script (remplace les anciens), puis Déployer > Gérer les déploiements > ✏️ > Nouvelle version.
2. **Générer le secret partagé** : dans l'éditeur Apps Script, sélectionne la fonction `printTotpSetupInfo_` dans le menu déroulant en haut, clique ▶ Exécuter, puis ouvre Affichage > Journaux (ou Ctrl/Cmd+Entrée). Tu obtiens :
   - un `secret` (à saisir manuellement dans l'appli d'authentification si elle ne lit pas de QR code), et
   - une URL `otpauth://...` (à transformer en QR code via n'importe quel générateur en ligne, pour un scan direct).
3. **Ajouter le compte sur les 3 téléphones admin** : dans Google Authenticator (ou équivalent) sur chacun des 3 téléphones, scanne le QR code (ou saisis le secret manuellement). Un compte « Atelier des Artisans » apparaît, avec un code à 6 chiffres qui se renouvelle toutes les 30 secondes — identique sur les 3 téléphones.
4. **Supprimer l'ancien déclencheur** (si tu en avais créé un pour `rotateAccessCode`) : icône ⏰ Déclencheurs dans l'éditeur Apps Script > supprime-le, il n'est plus utilisé.
5. Vérifie que `script-url.txt` pointe toujours vers la même URL de déploiement (inchangé si tu redéployais une nouvelle version du même déploiement).

Ensuite, pour enrôler un artisan : un admin ouvre son appli d'authentification, lit le code du moment, le communique à l'artisan qui le saisit dans l'écran de connexion — comme avant, mais sans jamais passer par le Sheet.
