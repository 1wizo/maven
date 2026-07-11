# Maven — Un guide clair pour Office LTSC + activation

Soyons honnêtes sur ce qu'est ce guide : un parcours sans détour pour installer Microsoft Office LTSC sur Windows sans payer d'abonnement, puis activer Windows et Office pour qu'ils arrêtent de vous harceler. On utilise des outils publiés par Microsoft eux-mêmes, plus un projet open-source appelé MAS qui gère l'activation. Pas d'ISO louches, pas de cracks téléchargés sur des forums random, pas d'« outils d'activation » qui sont en fait des malwares.

J'ai aussi inclus une section pour les utilisateurs Linux à la fin, parce que tout le monde ne tourne pas sous Windows et que la situation Office de ce côté-là est franchement confuse.

---

## Ce qu'il vous faut

Avant de commencer, assurez-vous d'avoir :

- **Windows 10 ou 11** — n'importe quelle édition fonctionne, y compris LTSC. Windows 7/8.1 fonctionnent techniquement pour l'activation mais pas pour Office LTSC 2021+.
- **Environ 5 Go d'espace libre** — Office lui-même fait 2–4 Go, plus la place pour l'installateur et les fichiers temporaires.
- **Une connexion internet stable** — vous téléchargez Office directement depuis le CDN de Microsoft, donc la vitesse dépend de votre ligne.
- **Un accès administrateur** — nécessaire pour l'installateur Office et le script d'activation.
- **PowerShell 5.1 ou plus récent** — déjà installé sur toutes les machines Windows 10/11, vous n'avez probablement pas à vous en soucier.

Temps total : environ 15–30 minutes, essentiellement à attendre que le téléchargement Office se termine.

---

## Partie 1 — Installer Office LTSC

L'approche peut paraître un peu détournée, mais c'est en fait la façon la plus propre de faire. Au lieu de télécharger un ISO Office cracké, on va utiliser l'Office Deployment Tool (ODT) de Microsoft — le même outil que les départements IT utilisent pour déployer Office sur des centaines de machines. La différence, c'est qu'on le pointe vers le canal Volume License, ce qui nous donne la build LTSC que MAS peut activer.

### Étape 1 — Récupérer l'Office Deployment Tool

Allez sur la page de téléchargement officielle de Microsoft :

```
https://www.microsoft.com/download/details.aspx?id=49117
```

Cliquez sur **Download**, lancez le fichier `officedeploymenttool_*.exe` qui arrive, acceptez le contrat de licence, et extrayez tout dans un dossier dont vous souviendrez — `C:\OfficeInstall\` fait très bien l'affaire.

Ce que vous obtiendrez : `setup.exe` (c'est le fichier important — gardez-le) et une poignée de fichiers XML d'exemple que vous pouvez ignorer. On va générer notre propre XML à l'étape suivante.

### Étape 2 — Créer votre configuration

C'est là que vous dites à l'installateur exactement ce que vous voulez. Microsoft a un outil web pour ça, et il est franchement bien fait :

```
https://config.office.com/deploymentsettings
```

Connectez-vous avec n'importe quel compte Microsoft — un compte Outlook.com gratuit fait l'affaire, pas besoin de rien de spécial. Une fois dedans, vous verrez un assistant. Voici ce qu'il faut régler :

**Dans « Office Products » :**
- **Products :** Choisissez « Office LTSC Professional Plus 2021 - Volume License » (ou 2024 si vous voulez la version plus récente)
- **Update channel :** Ça c'est critique — mettez « Office LTSC 2021 Perpetual Enterprise ». Si vous choisissez le mauvais canal, vous aurez la build grand public au lieu de la build Volume License, et MAS ne pourra pas l'activer.
- **Apps :** Cochez celles que vous voulez vraiment. Word, Excel, PowerPoint, Outlook et OneNote sont les cinq de base. Visio et Project sont dispo aussi si vous en avez besoin. Décochez les indésirables — Clipchamp, Bing, et tout ce que Microsoft colle dedans ces temps-ci.

**Dans « Language » :**
Prenez ce que vous voulez. English (US) est la valeur sûre par défaut, mais Français, Arabe, Espagnol, Allemand et des dizaines d'autres sont tous là.

**Dans « Settings » :**
- Installation platform : **64-bit** (y a pas de bonne raison de prendre 32-bit en 2025)
- Update automatically : **Yes**
- Pin to taskbar : **No** (personne veut ça)
- Remove older versions : **Yes** (ça désinstalle proprement un éventuel Office existant avant d'installer LTSC)

**Dans « Licensing and Activation » :**
- Accept EULA : **Yes**
- Product key : **laisser vide** (la build Volume License n'a pas besoin de clé intégrée dans la config)
- AUTOACTIVATE : **FALSE** — on gérera l'activation séparément avec MAS

Une fois que tout est bon, cliquez sur **Export** en haut à droite, choisissez **XML** comme format, et enregistrez le fichier sous `configuration.xml` dans le même dossier que `setup.exe`.

### Étape 3 — Lancer l'installation

Ouvrez l'Invite de commandes en tant qu'administrateur (menu Démarrer → tapez `cmd` → clic droit → Exécuter en tant qu'administrateur), naviguez jusqu'à votre dossier Office, et lancez l'installateur :

```batch
cd /d C:\OfficeInstall
setup.exe /configure configuration.xml
```

Une petite fenêtre Office va s'ouvrir et commencer le téléchargement. C'est le moment d'aller faire un café — ça tire 2–4 Go directement depuis le CDN de Microsoft, donc selon votre connexion, comptez entre 5 et 20 minutes. L'installation tourne silencieusement en arrière-plan ; pas besoin de cliquer dans des assistants.

Quand c'est fini, vous verrez Word, Excel, PowerPoint et les autres dans votre menu Démarrer. Ils se lancent sans souci, mais vous remarquerez une bannière « Activer Office » — c'est normal, on s'en occupe dans la Partie 3.

### Étape 4 — Vérification rapide

Ouvrez Word. Il devrait se lancer sans demander de clé produit, et vous devriez pouvoir créer et enregistrer des documents. La bannière d'activation sera là, mais l'appli elle-même est pleinement fonctionnelle. Si Word ne se lance pas du tout, l'installation a raté — vérifiez que `setup.exe` s'est terminé avec le code 0 dans votre invite de commande.

<details>
<summary>Envie d'écrire le XML à la main ? Voici les IDs de produit</summary>

```xml
<Configuration>
  <Add OfficeClientEdition="64-bit" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="en-us" />
    </Product>
    <Product ID="Visio2021Volume">
      <Language ID="en-us" />
    </Product>
    <Product ID="Project2021Volume">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="FALSE" />
</Configuration>
```

| ID de produit | Ce qu'il installe |
|---|---|
| `ProPlus2021Volume` | Office LTSC Pro Plus 2021 |
| `Standard2021Volume` | Office LTSC Standard 2021 |
| `Visio2021Volume` | Visio LTSC 2021 |
| `Project2021Volume` | Project LTSC 2021 |
| `ProPlus2024Volume` | Office LTSC Pro Plus 2024 |
| `Visio2024Volume` | Visio LTSC 2024 |
| `Project2024Volume` | Project LTSC 2024 |

Pour les versions 2024, remplacez `Channel="PerpetualVL2021"` par `"PerpetualVL2024"` et utilisez les IDs de produit 2024.
</details>

---

## Partie 2 — Activer Windows

Parlons activation Windows. L'outil qu'on utilise s'appelle MAS (Microsoft Activation Scripts). C'est un projet open-source qui existe depuis des années, maintenu activement sur GitHub, et qui fonctionne en générant une licence numérique légitime liée au matériel — le même type que vous auriez si vous achetiez Windows chez Microsoft et vous connectiez à votre compte Microsoft.

La méthode qu'on va utiliser s'appelle HWID. Elle crée une licence numérique permanente liée à votre matériel, donc elle survit aux réinstallations, aux mises à jour de fonctionnalités, et même à la plupart des changements de hardware. C'est le même type d'activation que Microsoft donne aux gens qui ont mis à niveau depuis Windows 7/8 gratuitement.

### Étape 1 — Ouvrir PowerShell en tant qu'administrateur

Appuyez sur la touche Windows, tapez `powershell`, clic droit sur **Windows PowerShell**, et choisissez **Exécuter en tant qu'administrateur**. Cliquez sur **Oui** à l'invite UAC.

### Étape 2 — Lancer MAS

Collez cette ligne unique et appuyez sur Entrée :

```powershell
irm https://get.activated.win | iex
```

Si `get.activated.win` est bloqué par votre DNS ou pare-feu (certains réseaux d'entreprise font ça), utilisez le miroir :

```powershell
irm https://massgrave.dev/get | iex
```

Un menu textuel va apparaître. Ça ressemble à un programme DOS de 1995, mais ça marche parfaitement.

### Étape 3 — Choisir la méthode d'activation

Tapez `1` et Entrée pour sélectionner « Activate Windows ». Vous verrez alors trois méthodes :

1. **HWID** — C'est celle qu'il vous faut. Elle crée une licence numérique permanente liée à votre matériel. Idéal pour Windows 10 et 11. Utilisez celle-ci.
2. **KMS38** — Valide jusqu'en 2038. Utilisez-la si HWID échoue pour une raison quelconque, ou si vous êtes sur Windows Server.
3. **Online KMS** — Se renouvelle tous les 180 jours. À utiliser uniquement si vous êtes sur Windows 7 ou 8.1, qui ne supportent pas HWID.

Choisissez l'option 1 (HWID), confirmez avec `Y`, et attendez. Ça prend généralement 10–30 secondes. Vous verrez un message « Activation successful » quand c'est fait.

### Étape 4 — Vérifier

Tapez `5` dans le menu MAS et Entrée pour vérifier le statut. Ça devrait dire « Windows is permanently activated ».

C'est tout. La licence est liée à votre carte mère et à la signature de votre CPU, donc elle survit aux mises à jour Windows, aux montées de version (comme 23H2 → 24H2), et même à une réinstall propre sur le même matériel.

---

## Partie 3 — Activer Office

Office est installé mais pas encore activé. On va utiliser une autre méthode MAS appelée Ohook pour celle-ci. Ohook fonctionne en patchant la vérification d'activation d'Office au niveau binaire — c'est permanent, ça marche hors-ligne, et ça survit aux mises à jour Office. C'est la méthode d'activation Office la plus fiable que MAS propose.

### Étape 1 — Fermer toutes les applis Office

Ça c'est important. Assurez-vous que Word, Excel, PowerPoint, Outlook, Visio et Project sont tous complètement fermés. Vérifiez la zone de notification (en bas à droite de la barre des tâches) et quittez les icônes Office qui s'y cachent. Ohook doit modifier des fichiers qu'Office verrouille pendant qu'il tourne, donc si quelque chose est ouvert, l'activation échouera.

### Étape 2 — Relancer MAS

Même lanceur qu'avant :

```powershell
irm https://get.activated.win | iex
```

### Étape 3 — Activer

Dans le menu MAS :
- Tapez `2` et Entrée (Activate Office)
- Tapez `1` et Entrée (Ohook)
- Tapez `Y` pour confirmer

Attendez « Activation successful ». Ça prend généralement 5–15 secondes.

### Étape 4 — Vérifier

Tapez `5` dans le menu MAS pour vérifier le statut. Tous vos produits Office devraient afficher « LICENSED ». Ouvrez Word, allez dans Fichier → Compte, et vous verrez « Microsoft Office LTSC Professional Plus 2021 » sans bannière d'activation. C'est fini.

---

## Partie 4 — Alternatives Office sur Linux

Si vous lisez ça sur Linux, la situation est différente. Microsoft Office ne tourne pas nativement sur Linux, et honnêtement, ça n'arrivera probablement jamais. Mais vous avez deux options solides selon ce dont vous avez réellement besoin.

**OnlyOffice** est une suite office Linux native, gratuite, open-source, qui lit et écrit les fichiers `.docx` mieux que tout ce que j'ai essayé. Si vous avez juste besoin d'écrire des documents, faire des tableurs et des présentations — et que vous voulez qu'ils rendent correctement quand vos collègues sous Windows les ouvrent — c'est celle-là.

**LinOffice** c'est autre chose. C'est un installateur en 1 clic qui monte une VM Windows dans un conteneur et fait tourner le vrai Microsoft Office 2024 dedans, puis présente chaque appli Office comme si c'était une fenêtre Linux native. C'est plus lourd (nécessite 8 Go RAM et KVM), et il vous faut quand même une licence Office, mais vous avez une compatibilité MS Office 100% parfaite avec le support complet des macros VBA.

### Laquelle choisir, concrètement ?

Honnêtement, pour 90% des utilisateurs Linux, **OnlyOffice** est la bonne réponse. C'est léger, c'est gratuit, ça ne nécessite pas de VM, et la fidélité `.docx` est excellente. La seule raison de passer à LinOffice, c'est si vous avez absolument besoin du vrai Microsoft Office — peut-être que votre boulot exige des macros VBA qu'OnlyOffice ne peut pas lancer, ou vous avez besoin d'un rendu parfait de modèles Excel complexes, ou il vous faut Outlook pour l'intégration Exchange.

| | OnlyOffice | LinOffice |
|---|---|---|
| **Ce que c'est** | Suite office Linux native | Vrai MS Office 2024 dans une VM Windows |
| **Licence** | AGPL v3 (gratuit, open-source) | AGPL v3 (gratuit, open-source) — fork de WinApps |
| **Fidélité MS Office** | Excellente — OOXML natif | Parfaite — c'EST Microsoft Office |
| **Inclut** | Équivalents Word, Excel, PowerPoint | Vrai Word, Excel, PowerPoint, OneNote, Outlook |
| **Utilisation ressources** | Léger — environ 450 Mo | Lourd — 8 Go RAM, 64 Go disque, VM Windows complète |
| **Matériel** | Tous (x86_64, ARM, etc.) | x86_64 uniquement + virtualisation KVM |
| **Activation requise** | Non — totalement gratuit | Oui — clé Office 2024 ou abo M365 (MAS fonctionne aussi) |
| **Intégration Linux** | Native — associations de fichiers, notifications, tout OK | Bonne via FreeRDP ; quelques soucis sur Wayland/multi-écran |

### Installer OnlyOffice

Choisissez la méthode qui correspond à votre distro :

```bash
# Snap (fonctionne sur toute distro qui a snapd)
sudo snap install onlyoffice-desktopeditors

# Flatpak (fonctionne sur toute distro qui a Flatpak)
flatpak install flathub org.onlyoffice.desktopeditors

# Ubuntu / Debian (.deb — téléchargez depuis onlyoffice.com/download-desktop.aspx)
sudo apt install ./onlyoffice-desktopeditors-amd64.deb

# Arch (AUR)
yay -S onlyoffice-bin
```

### Installer LinOffice

LinOffice a un script quickstart qui gère tout automatiquement — il détecte votre distro, installe toutes les dépendances, télécharge l'installateur, et lance la GUI :

```bash
curl -sSL https://github.com/eylenburg/linoffice/raw/refs/heads/main/quickstart.sh -o quickstart.sh \
  && chmod +x quickstart.sh && ./quickstart.sh
```

Si vous préférez installer les dépendances manuellement :

```bash
# Ubuntu 24.04+ / Debian 13+
sudo apt install podman podman-compose freerdp3 python3 \
  python3-pyside6 python3-pyside6.qtwidgets python3-pyside6.qtuitools

# Fedora
sudo dnf install podman podman-compose freerdp python3 python-pyside6

# Arch
sudo pacman -Syu podman podman-compose freerdp python pyside6

# openSUSE
sudo zypper install podman podman-compose freerdp python3 python3-pyside6
```

Ensuite téléchargez le repo, extrayez-le, et lancez soit `./setup.sh` (CLI) soit `python3 gui/linoffice.py` (GUI).

**Attention à la config matérielle :** LinOffice a besoin d'un CPU x86_64 avec support KVM (vérifiez votre BIOS), d'au moins 8 Go de RAM, et d'environ 64 Go d'espace disque libre. Le téléchargement initial depuis Microsoft fait autour de 8 Go, et toute l'installation prend environ 15 minutes sur une connexion correcte. Aussi, évitez FreeRDP version 3.23 — il a un bug connu. Prenez 3.22 ou attendez 3.24.

Une fois installé, vous avez une commande `linoffice` :

```bash
linoffice word           # lancer Word
linoffice excel          # lancer Excel
linoffice powerpoint     # lancer PowerPoint
linoffice onenote        # lancer OneNote
linoffice outlook        # lancer Outlook
linoffice windows        # bureau Windows complet via RDP
linoffice update         # mettre à jour Windows + Office
linoffice stopcontainer  # arrêter la VM (garde vos données)
```

L'installateur crée aussi des entrées `.desktop`, donc Word, Excel, etc. apparaissent dans votre lanceur d'applis comme des applis Linux normales. Les fichiers dans `/home` sont accessibles depuis la VM, donc vous pouvez ouvrir et enregistrer des documents dans votre dossier home Linux habituel.

### Et Wine/Bottles/CrossOver alors ?

Vous pouvez techniquement faire tourner MS Office via Wine ou Bottles, mais honnêtement, c'est fragile. Chaque mise à jour Office a une chance de casser quelque chose, le support des macros est inégal, et l'activation devient bizarre. CrossOver est l'option la plus aboutie mais elle est payante. LinOffice est la solution gratuite la plus propre si vous avez vraiment besoin du vrai Office — c'est essentiellement WinApps (son projet amont) avec une couche d'automatisation par-dessus.

---

## Aide-mémoire rapide

Si vous avez déjà fait ça et que vous voulez juste les commandes :

**Installer Office :**
```batch
cd /d C:\OfficeInstall
setup.exe /configure configuration.xml
```

**Activer Windows :**
```powershell
irm https://get.activated.win | iex
```
→ `1` → `1` (HWID) → `Y`

**Activer Office :**
```powershell
irm https://get.activated.win | iex
```
→ `2` → `1` (Ohook) → `Y`

**Les deux en même temps (silencieux) :**
```powershell
& ([ScriptBlock]::Create((irm https://get.activated.win))) -HWID -Ohook -S
```

**Vérifier le statut d'activation :**
```batch
slmgr /xpr                    :: Windows
cscript ospp.vbs /dstatus     :: Office (lancé depuis le dossier Office)
```

---

## Dépannage

De vrais problèmes que j'ai réellement rencontrés, et comment les corriger :

**L'installateur Office plante** — Assurez-vous que `configuration.xml` est dans le même dossier que `setup.exe`. Si le XML est mal formé, réexportez-le depuis config.office.com. La cause la plus courante est une coquille dans l'ID produit ou le nom du canal.

**L'installation bloque à « Downloading »** — Généralement un problème réseau. Désactivez tout VPN ou proxy, vérifiez que `*.officecdn.microsoft.com` n'est pas bloqué par votre pare-feu, et que vous avez bien 5+ Go de libre sur votre disque système.

**Vous avez eu Office 365 au lieu de LTSC** — Vous avez choisi le mauvais canal de mise à jour dans la config. Retournez sur config.office.com et mettez « Office LTSC 2021 Perpetual Enterprise » — pas « Current Channel » ou « Monthly Enterprise ».

**Visio ou Project ne s'est pas installé** — Ils doivent être ajoutés comme entrées `<Product>` séparées dans le XML, pas juste cochés dans la liste des applis. Voir l'exemple XML dans la Partie 1.

**« Impossible d'installer Office 64-bit avec la 32-bit déjà installée »** — Mettez « Remove older versions = Yes » dans la config, ou désinstallez manuellement l'ancien Office 32-bit depuis Paramètres → Applications d'abord.

**MAS ne peut pas résoudre `get.activated.win`** — Votre DNS le bloque. Utilisez le miroir : `irm https://massgrave.dev/get | iex`

**Windows Defender supprime le script MAS** — C'est attendu. MAS déclenche des faux positifs dans la plupart des antivirus. Désactivez temporairement la protection en temps réel, ajoutez le dossier aux exclusions, lancez MAS, puis réactivez Defender. C'est safe — MAS est open-source et audité.

**HWID échoue avec l'erreur 0xC004C003** — Essayez KMS38 à la place (option 2). C'est valide jusqu'en 2038 et ça marche sur les mêmes builds Windows. Vous pouvez aussi lancer le dépanneur intégré : menu MAS → `[6] Troubleshoot` → `Fix Licensing`, puis réessayez.

**MAS ne détecte pas Office** — Vous avez installé la build grand public (Click-to-Run) au lieu de la build Volume License. La solution est de reconstruire votre config avec le canal `PerpetualVL2021` et de réinstaller.

**L'activation dit « 180 jours » au lieu de permanent** — Vous avez eu Online KMS au lieu de HWID/Ohook. Relancez MAS et choisissez spécifiquement HWID pour Windows (option 1) et Ohook pour Office (option 1).

**L'activation a disparu après un changement de matériel** — Les licences HWID sont liées à votre carte mère et CPU. Si vous remplacez les deux, vous devrez peut-être relancer HWID. Il se reconnectera automatiquement au nouveau matériel.

**Erreurs réseau (0x80072EE2, 0x80072EFD)** — Ce sont des problèmes de connectivité. Coupez le VPN/proxy, autorisez `*.windows.com` et `*.microsoft.com` dans votre pare-feu, et essayez un autre résolveur DNS (1.1.1.1 ou 8.8.8.8 marchent souvent mieux que le DNS du FAI).

---

## FAQ

**Puis-je faire ça sur une installation Windows fraîche ?**
Oui. L'ordre n'a pas d'importance — vous pouvez activer Windows d'abord et installer Office après, ou inversement. Les deux fonctionnent sur un Windows 10/11 propre sans configuration supplémentaire.

**Ai-je besoin d'un compte Microsoft ?**
Uniquement pour vous connecter à config.office.com et créer le XML de configuration. L'installateur Office lui-même ne nécessite pas de connexion — il télécharge directement depuis le CDN de Microsoft sans authentification. Vous pouvez utiliser un compte Outlook.com jetable pour l'outil de config.

**Office se mettra-t-il à jour tout seul ?**
Oui. LTSC reçoit les correctifs de sécurité via Windows Update. Le « LTSC » veut dire « Long-Term Servicing Channel » — vous avez les correctifs de sécurité mais pas les mises à jour de fonctionnalités. L'interface ne changera pas tous les mois comme avec Office 365. C'est tout l'intérêt.

**L'activation survit-elle aux mises à jour de fonctionnalités Windows (comme 23H2 → 24H2) ?**
Oui. HWID crée une licence numérique liée au matériel qui persiste à travers les montées de version. Ohook patche les fichiers d'activation d'Office, qui survivent aussi aux mises à jour. Vous ne devriez pas avoir à réactiver après une mise à jour Windows normale.

**Puis-je installer juste Visio et Project sans Word/Excel ?**
Oui. Dans config.office.com, décochez tout sauf Visio et Project dans la section Apps. Ou si vous écrivez le XML à la main, incluez juste `Visio2021Volume` et `Project2021Volume` comme seules entrées `<Product>`.

**Comment passer d'Office 365 à LTSC ?**
Mettez « Remove older versions = Yes » dans la config (ou ajoutez `<RemoveMSI />` au XML). Quand vous lancez `setup.exe /configure`, ça désinstalle 365 et installe LTSC en une seule passe. Vos documents ne seront pas touchés.

**Est-ce que ça marche sur Windows 7 ou 8.1 ?**
MAS fonctionne sur Windows 7/8.1 pour l'activation Windows, mais Office LTSC 2021/2024 nécessite Windows 10 ou plus. Si vous êtes sur une vieille version de Windows, votre meilleur pari est Office 2016 + activation Online KMS.

**Où sont les logs MAS ?**
`C:\ProgramData\MAS\` — ce dossier contient les logs d'activation, les infos de diagnostic, et les fichiers de ticket. Regardez ici si quelque chose tourne mal et que vous voulez déboguer.

**Comment désinstaller Office plus tard ?**
Paramètres → Applications → Microsoft Office → Désinstaller. Ou vous pouvez utiliser l'ODT avec un XML de désinstallation — voir la doc Microsoft pour le schéma.

**Puis-je activer Office 365 Family/Personal avec MAS ?**
Ohook fonctionne sur les builds Office 365 grand public. KMS non — les builds 365 grand public ne sont pas Volume License, donc KMS38 et Online KMS ne les activeront pas. Si vous avez un abo 365 grand public, Ohook est votre seule option MAS.

**La commande `irm | iex` est-elle sûre à copier/coller ?**
Oui. `get.activated.win` est le raccourci officiel maintenu par le projet MAS. Il redirige vers le script hébergé dans le dépôt GitHub `massgravel/Microsoft-Activation-Scripts`. Vérifiez toujours que vous utilisez l'URL officielle — il y a des imposteurs dehors.

**Comment mettre à jour MAS vers une version plus récente ?**
Lancez simplement la même commande `irm https://get.activated.win | iex` à nouveau. Les nouvelles versions remplacent automatiquement les anciennes. Il n'y a rien à désinstaller — MAS détecte et écrase la version précédente.

---

## Liens utiles

- **Page du projet MAS :** https://massgrave.dev
- **Dépôt GitHub MAS :** https://github.com/massgravel/Microsoft-Activation-Scripts
- **Dernière version MAS (.zip) :** https://github.com/massgravel/Microsoft-Activation-Scripts/releases/latest
- **Office Deployment Tool :** https://www.microsoft.com/download/details.aspx?id=49117
- **Office Customization Tool :** https://config.office.com/deploymentsettings

---

*Ce guide est à but éducatif. Tirez toujours les scripts depuis le dépôt GitHub officiel pour garantir l'authenticité. L'auteur et le projet MAS ne sont pas affiliés à Microsoft.*
