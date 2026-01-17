

# 📦 Gestion de stock – Projet d’initiation au VBA

Ce projet est une application simple de **gestion de stock** destinée à un service informatique, développée sous **Microsoft Excel 2024** à l’aide du **VBA** et de l’environnement **VBE**.  
Il s’agit avant tout d’un **projet d’apprentissage**, conçu pour découvrir le VBA, créer des formulaires, manipuler des données Excel et comprendre la logique des macros.

---

## 🎯 Objectifs du projet

- Découvrir l’environnement **VBE** (UserForms, modules, console de débogage).  
- Comprendre l’interaction entre VBA et les feuilles Excel.  
- Utiliser l’**enregistreur de macros** pour générer du code (tri, filtres).  
- Mettre en place des **fonctions CRUD** simples.  
- Structurer un classeur Excel propre et modulaire.  
- Explorer le VBA sans objectif d’expertise, uniquement pour le plaisir d’apprendre.

---

## 🧩 Fonctionnalités principales

- Interface utilisateur complète pour gérer le stock.  
- Ajout, modification, suppression et consultation d’éléments.  
- Gestion des mouvements (entrées / sorties).  
- Tri automatique des tableaux.
- Filtre par mot clé ou selon le seuil d'alerte
- Actualisation dynamique de l’interface.  
- Gestion d’erreurs robuste.  
- Organisation claire du classeur et des données.

---

## 👥 Contributions

### Développeur
- Idée du projet.  
- Construction et développement complet de l'interface utilisateur.  
- Enregistrement des macros de tri et de filtre.  
- Mise en place de l’architecture globale (modules, constantes, séparation des responsabilités).  
- Développement initial des fonctions CRUD.  
- Structuration du classeur Excel.

### Assistant IA – Copilot
- Développement avancé des fonctions CRUD.  
- Gestion de l’actualisation des données dans l’interface.  
- Mise en place d’une gestion d’erreurs robuste.  
- Documentation complète du code (commentaires, explications).  
- Rédaction du README complet.

---

## 🛠️ Technologies utilisées

- **Microsoft Excel 2024**  
- **VBE (Visual Basic Editor)**  
- **VBA (Visual Basic for Applications)**  

---

# 🧭 Installation & Mise en place de l’environnement Excel

Cette section permet de **recréer exactement le classeur** nécessaire au fonctionnement de l’application.

---

## 1️⃣ Créer le classeur Excel

1. Ouvrir Excel.  
2. Créer un nouveau classeur.  
3. L’enregistrer immédiatement au format **.xlsm** (macro-enabled).  
4. Nom conseillé :  
   **`stock_service_informatique.xlsm`**

---

## 2️⃣ Créer les feuilles nécessaires

Créer **trois feuilles** avec les noms suivants :

- `stock`
- `movement`
- `configuration`

---

# 3️⃣ Feuille **stock**

Créer un **tableau structuré** nommé **`stock`**, à partir de la cellule **A1**.

### Colonnes (dans cet ordre) :

| Colonne | Type | Notes |
|--------|------|-------|
| libellé | texte | **tri A → Z** |
| stock | nombre | nombre entier | 
| catégorie | texte | |
| maj | date courte | date de mise à jour |
| seuil | nombre | nombre entier |
| sous-catégorie | texte | |
| commentaire | texte | |
| ligne_tableau | nombre | utilisé par le code |
| ligne_feuille | nombre | utilisé par le code |

---

# 4️⃣ Feuille **movement**

Créer un tableau structuré nommé **`movement`**, à partir de **A1**.

### Colonnes :

| Colonne | Type | Notes |
|--------|------|-------|
| date | date courte | tri Z → A |
| type | texte | entrée / sortie |
| valeur | nombre | nombre entier |
| description | texte | |
| matériel | texte | correspond au libellé du stock |

---

# 5️⃣ Feuille **configuration**

Cette feuille contient **tous les tableaux de configuration**, chacun trié **A → Z**, et chacun portant un nom spécifique.

Chaque tableau occupe **une seule colonne**, et commence dans une colonne différente :  
**A, C, E, G, I, K, M, O, Q, S, U, W**.

Tous les tableaux doivent être créés en tant que **tableaux structurés Excel**, avec les noms suivants :

- `category`
- `office_equipment`
- `printer_scanner`
- `internal_component`
- `peripheral`
- `network_hardware`
- `storage`
- `connector_cabling`
- `accessorie`
- `consumable`
- `software_licence`
- `mobile_hardware`

---

## 📋 Données complètes à insérer dans les tableaux  

### 🔹 Tableau `category`
```
accessoire
composant interne
connectique/câblage
consommable
imprimante/scanner
logiciel/licence
matériel de bureau
matériel mobile
matériel réseau
périphérique
stockage
```

### 🔹 Tableau `office_equipment`
```
écran/moniteur
ordinateur fixe
ordinateur portable
station de travail
vidéoprojecteur
```

### 🔹 Tableau `printer_scanner`
```
imprimante jet d'encre
imprimante laser
imprimante multifonction
scanner
```

### 🔹 Tableau `internal_component`
```
alimentation électrique
boîtier
carte graphique
carte mère
mémoire vive (RAM)
processeur (CPU)
```

### 🔹 Tableau `peripheral`
```
casque
clavier
microphone
souris
webcam
```

### 🔹 Tableau `network_hardware`
```
carte réseau
commutateur
concentrateur
point d'accès
routeur
```

### 🔹 Tableau `storage`
```
carte mémoire
carte SD
clé USB
disque externe
disque HDD interne
disque SSD interne
serveur NAS
```

### 🔹 Tableau `connector_cabling`
```
adaptateur et convertisseur
câble audio
câble de données
câble réseau
câble vidéo
```

### 🔹 Tableau `accessorie`
```
batterie et chargeur
onduleur (UPS)
outil et kit de nettoyage
pile et accumulateur
sacoche
tapis de souris
```

### 🔹 Tableau `consumable`
```
cartouche d'encre
papier
toner
```

### 🔹 Tableau `software_licence`
```
logiciel de sécurité
logiciel métier
suite bureautique
système d'exploitation
```

### 🔹 Tableau `mobile_hardware`
```
smartphone
smartwatche
tablette
```

---

# ▶️ Utilisation

1. Ouvrir le fichier Excel.  
2. Activer les macros.  
3. À l’ouverture, **le classeur se masque automatiquement** et **l’application (interface utilisateur) s’affiche seule**.  
4. Avant utilisation, **importer les fichiers fournis dans le dépôt GitHub** :  
   - Modules (`.bas`)  
   - Formulaires (`.frm`)  
   - Classes (`.cls`)  
   via **VBE → Fichier → Importer un fichier…**  
5. Utiliser l’application pour :  
   - Ajouter un matériel  
   - Modifier une entrée  
   - Supprimer un élément  
   - Enregistrer un mouvement (entrée/sortie)  
6. À la fermeture de l’interface, une fenêtre propose :  
   - **Fermer complètement l’application et le classeur**, ou  
   - **Fermer uniquement l’application et afficher le classeur Excel**.  
7. Les tableaux se mettent automatiquement à jour selon les actions effectuées.