# -- Encodage UTF-8 pour gérer les accents --
[System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

param(
    [switch]$Silent  # Si -Silent est spécifié, on n'affiche pas de boîtes de dialogue
)

Add-Type -AssemblyName System.Windows.Forms

# Fichier de log
$logFile = "\\alphatech.local\data\partage\IT\Logs\sauvegarde.log"

# Fonction pour écrire dans le log avec horodatage
function Write-Log($message) {
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value "$timestamp $message"
}

# Mode manuel : on demande confirmation par boîte de dialogue
if (-not $Silent) {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous lancer la sauvegarde ?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::No) {
        Write-Host "Sauvegarde annulée par l'utilisateur."
        Write-Log "Sauvegarde annulée par l'utilisateur."
        exit
    }
    Write-Host "Sauvegarde confirmée par l'utilisateur."
    Write-Log  "Sauvegarde confirmée par l'utilisateur."
}
else {
    # Mode silencieux
    Write-Host "Sauvegarde lancée en mode silencieux (WSL/Automatique)."
    Write-Log  "Sauvegarde lancée en mode silencieux (WSL/Automatique)."
}

# Définition des chemins
$source      = "\\alphatech.local\data\partage"
$destination = "\\10.11.11.201\recovery"
$date        = Get-Date -Format "yyyy-MM-dd"  # ex: 2025-02-19
$backupName  = "backup_$date.zip"
$backupPath  = Join-Path $destination $backupName

# Création de la sauvegarde
Write-Host "Démarrage de la sauvegarde..."
Write-Log  "Démarrage de la sauvegarde..."

try {
    # 1) Créer l’archive ZIP
    Compress-Archive -Path "$source\*" -DestinationPath $backupPath -Force

    # 2) Vérifier la création
    if (Test-Path $backupPath) {
        Write-Host "Sauvegarde créée : $backupPath"
        Write-Log  "Sauvegarde créée : $backupPath"
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

    # 3) Limiter à 10 sauvegardes
    $allBackups = Get-ChildItem -Path $destination -Filter "backup_*.zip" -File `
                  | Sort-Object LastWriteTime -Descending

    if ($allBackups.Count -gt 10) {
        # On conserve les 10 plus récentes, on supprime le reste
        $toRemove = $allBackups | Select-Object -Skip 10
        foreach ($old in $toRemove) {
            Write-Host "Suppression de l'ancienne sauvegarde : $($old.Name)"
            Write-Log  "Suppression de l'ancienne sauvegarde : $($old.FullName)"
            Remove-Item $old.FullName -Force
        }
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
