# AEOS Import Automation

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE?logo=powershell&logoColor=white)
![SQL Server](https://img.shields.io/badge/SQL%20Server-Staging%20%2B%20SP-CC2927?logo=microsoftsqlserver&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

Script d'automatisation **PowerShell de production** pour l'import de personnel prestataire accrédité dans le système de contrôle d'accès **Nedap AEOS**. Gère le parsing CSV, le staging SQL Server, la gestion des vendors via SOAP, et un audit logging complet.

## Fonctionnalités

- **Parsing CSV robuste** — Gestion multi-encodages (UTF-8, Windows-1252, UTF-16), détection BOM, réparation mojibake
- **Mapping intelligent des en-têtes** — Détection automatique des colonnes quelle que soit la variante de nommage
- **Intégration SQL Server** — Introspection de schéma, staging BulkCopy, exécution de procédure stockée
- **Gestion SOAP des vendors** — Création/mise à jour automatique des vendors AEOS avec logique de retry
- **Cycle de vie vendor** — Déblocage des vendors en base, polling de synchronisation
- **Piste d'audit complète** — Logging détaillé avec suivi UPSERT/BLOCK par personne
- **Identifiants sécurisés** — Utilisation de credentials chiffrés PowerShell CliXml (Windows DPAPI)
- **Configurable** — Tous les paramètres externalisés dans un fichier JSON

## Architecture

```
Fichier CSV (prestataires.csv)
    │
    ▼
┌──────────────────────┐
│  Parsing & validation│  ← Détection d'encodage, mapping des en-têtes
│  (PowerShell)        │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Staging SQL Server  │  ← BulkCopy vers PJ_PRESTATAIRES_IMPORT_STAGE
│  (SqlBulkCopy)       │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Synchro SOAP Vendor │  ← Ajout/modification vendors via API SOAP AEOS
│  (Optionnel)         │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Procédure stockée   │  ← dbo.sp_LoadAccreditesToImport
│  (SQL Server)        │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Archivage & audit   │  ← Déplacement CSV, logging UPSERT/BLOCK
└──────────────────────┘
```

## Stack technique

| Composant | Technologie |
|-----------|------------|
| Langage | PowerShell 5.1+ |
| Base de données | SQL Server (SqlClient) |
| SOAP | HTTP natif (Invoke-WebRequest) |
| Identifiants | PowerShell CliXml (chiffrement DPAPI) |
| Planification | Planificateur de tâches Windows |

## Installation

```bash
git clone https://github.com/jmanu1983/aeos-import-automation.git
cd aeos-import-automation
```

## Configuration

1. Copier la configuration exemple :
   ```powershell
   Copy-Item config\accredites.config.json.example config\accredites.config.json
   ```

2. Modifier le fichier de config avec vos paramètres (SQL Server, chemins, URL SOAP).

3. Créer les fichiers d'identifiants chiffrés :
   ```powershell
   # Identifiants SQL Server
   Get-Credential | Export-Clixml -Path config\sql_cred.xml

   # Identifiants SOAP AEOS
   Get-Credential | Export-Clixml -Path secrets\aeos-soap.cred.clixml
   ```

## Utilisation

```powershell
# Exécuter avec le chemin de config par défaut
.\bin\01-Import-Accredites.ps1

# Exécuter avec une config personnalisée
.\bin\01-Import-Accredites.ps1 -ConfigPath "D:\chemin\vers\config.json"
```

### Exécution planifiée

Configurer une tâche dans le Planificateur de tâches Windows :

```powershell
powershell.exe -ExecutionPolicy Bypass -File "D:\chemin\vers\bin\01-Import-Accredites.ps1"
```

## Structure du projet

```
aeos-import-automation/
├── bin/
│   └── 01-Import-Accredites.ps1    # Script d'import principal
├── config/
│   ├── accredites.config.json.example  # Modèle de configuration
│   └── PJ_PRESTATAIRES_IMPORT_LOAD.sql # Procédure stockée
├── secrets/                         # Identifiants chiffrés (hors VCS)
├── logs/                            # Logs d'exécution (hors VCS)
└── README.md
```

## Licence

Ce projet est sous licence MIT.
