<#
.SYNOPSIS
  Script de sauvegarde compressée avec limitation à 10 sauvegardes.

.DESCRIPTION
  - Sauvegarde le dossier source vers un dossier de destination en ZIP.
  - Nomme l'archive avec la date et l'heure (format "yyyy-MM-dd_HH-mm").
  - Conserve uniquement les 10 plus récentes sauvegardes.
  - Écrit des logs dans un fichier dédié.
  - Conçu pour être exécuté automatiquement via le planificateur de tâches.
#>

# ------------- GESTION DE L'ENCODAGE (Optionnel) -------------
try {
    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
} catch {
    Write-Host "Encodage UTF-8 non configuré : $($_.Exception.Message)"
}
# -------------------------------------------------------------

# Fichier de log
$logFile = "\\alphatech.local\data\partage\IT\Logs\sauvegarde.log"

# Fonction pour écrire dans le log avec horodatage
function Write-Log($message) {
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value "$timestamp $message"
}

# Début de la sauvegarde en mode automatique (sans boîtes de dialogue)
Write-Host "Démarrage de la sauvegarde automatique..."
Write-Log "Démarrage de la sauvegarde automatique..."

# Paramètres de la sauvegarde
$source      = "\\alphatech.local\data\partage"
$destination = "\\10.11.11.201\recovery"
$date        = Get-Date -Format "yyyy-MM-dd_HH-mm"
$backupName  = "backup_$date.zip"
$backupPath  = Join-Path $destination $backupName

Write-Host "Création de l'archive ZIP..."
Write-Log "Création de l'archive ZIP..."

try {
    # Création de l'archive ZIP
    Compress-Archive -Path "$source\*" -DestinationPath $backupPath -Force

    # Vérification de la création
    if (Test-Path $backupPath) {
        Write-Host "Sauvegarde créée : $backupPath"
        Write-Log "Sauvegarde créée : $backupPath"

        # Limiter à 10 sauvegardes
        $allBackups = Get-ChildItem -Path $destination -Filter "backup_*.zip" -File |
                      Sort-Object LastWriteTime -Descending

        if ($allBackups.Count -gt 10) {
            $toRemove = $allBackups | Select-Object -Skip 10
            foreach ($old in $toRemove) {
                Write-Host "Suppression de l'ancienne sauvegarde : $($old.Name)"
                Write-Log "Suppression de l'ancienne sauvegarde : $($old.FullName)"
                Remove-Item $old.FullName -Force
            }
        }
    }
    else {
        throw "Le fichier de sauvegarde n'a pas été créé."
    }
}
catch {
    $errMsg = "Échec de la sauvegarde : $($_.Exception.Message)"
    Write-Host $errMsg
    Write-Log $errMsg
    exit 1
}

Write-Log "Fin du script."
exit 0
