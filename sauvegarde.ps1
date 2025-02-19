<#
.SYNOPSIS
  Script de sauvegarde compressée avec limitation à 10 sauvegardes.

.DESCRIPTION
  - Sauvegarde le dossier source vers un dossier de destination en ZIP.
  - Nomme l’archive avec la date et l’heure (format "yyyy-MM-dd_HH-mm").
  - Conserve uniquement les 10 plus récentes sauvegardes.
  - Écrit des logs dans un fichier dédié.
  - Peut être exécuté en mode silencieux (sans boîtes de dialogue).

.PARAMETER Silent
  Lance la sauvegarde en mode silencieux (pas de boîtes de dialogue).
#>
# Sauvegarde ce fichier en UTF-8 (avec BOM)
Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.MessageBox]::Show(
    "Échec ou Succès ? Vérifiez les accents : é, è, à, ù.",
    "Test Accents",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
# ------------- GESTION DE L'ENCODAGE (Optionnel) -------------
# Si tu as une erreur sur CodePagesEncodingProvider, commente ou supprime ces lignes :
try {
    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
} catch {
    # Simplement ignorer si la version de .NET/PowerShell ne supporte pas
    Write-Host "Encodage UTF-8 non configuré (version de .NET insuffisante ?) : $($_.Exception.Message)"
}
# --------------------------------------------------------------

param(
    [switch]$Silent  # Si -Silent est spécifié, pas de boîtes de dialogue
)

# Pour les boîtes de dialogue Windows
Add-Type -AssemblyName System.Windows.Forms

# Fichier de log
$logFile = "\\alphatech.local\data\partage\IT\Logs\sauvegarde.log"

# Fonction pour écrire dans le log avec horodatage
function Write-Log($message) {
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value "$timestamp $message"
}

# -------------------------------------------------------------------------------------
# 1. Si on n’est pas en mode silencieux, on affiche une boîte de dialogue de confirmation
# -------------------------------------------------------------------------------------
if (-not $Silent) {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous lancer la sauvegarde ?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::No) {
        Write-Host "Sauvegarde annulée par l'utilisateur."
        Write-Log  "Sauvegarde annulée par l'utilisateur."
        exit
    }
    Write-Host "Sauvegarde confirmée par l'utilisateur."
    Write-Log  "Sauvegarde confirmée par l'utilisateur."
}
else {
    Write-Host "Sauvegarde lancée en mode silencieux (WSL/Automatique)."
    Write-Log  "Sauvegarde lancée en mode silencieux (WSL/Automatique)."
}

# -------------------------------------------------------------------------------------
# 2. Paramètres de la sauvegarde
# -------------------------------------------------------------------------------------
$source      = "\\alphatech.local\data\partage"
$destination = "\\10.11.11.201\recovery"
# On ajoute HH-mm (heure-minute) au format
$date        = Get-Date -Format "yyyy-MM-dd_HH-mm"
$backupName  = "backup_$date.zip"
$backupPath  = Join-Path $destination $backupName

Write-Host "Démarrage de la sauvegarde..."
Write-Log  "Démarrage de la sauvegarde..."

# -------------------------------------------------------------------------------------
# 3. Création de l’archive ZIP + Limitation du nombre de sauvegardes
# -------------------------------------------------------------------------------------
try {
    # a) Création de l’archive ZIP
    Compress-Archive -Path "$source\*" -DestinationPath $backupPath -Force

    # b) Vérification de la création
    if (Test-Path $backupPath) {
        Write-Host "Sauvegarde créée : $backupPath"
        Write-Log  "Sauvegarde créée : $backupPath"

        # c) Limiter à 10 sauvegardes
        $allBackups = Get-ChildItem -Path $destination -Filter "backup_*.zip" -File |
                      Sort-Object LastWriteTime -Descending

        if ($allBackups.Count -gt 10) {
            # On conserve les 10 plus récentes, on supprime le reste
            $toRemove = $allBackups | Select-Object -Skip 10
            foreach ($old in $toRemove) {
                Write-Host "Suppression de l'ancienne sauvegarde : $($old.Name)"
                Write-Log  "Suppression de l'ancienne sauvegarde : $($old.FullName)"
                Remove-Item $old.FullName -Force
            }
        }

        # d) Message de succès si pas en mode silencieux
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show(
                "Sauvegarde terminée avec succès !",
                "Succès",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    else {
        throw "Le fichier de sauvegarde n'a pas été créé."
    }
}
catch {
    $errMsg = "Échec de la sauvegarde : $($_.Exception.Message)"
    Write-Host $errMsg
    Write-Log  $errMsg

    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show(
            $errMsg,
            "Erreur",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}
