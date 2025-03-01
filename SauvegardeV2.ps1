<#
.SYNOPSIS
  Script de sauvegarde compressee avec limitation a 10 sauvegardes.

.DESCRIPTION
  - Sauvegarde le dossier source vers un dossier de destination en ZIP.
  - Nomme larchive avec la date et lheure (format "yyyy-MM-dd_HH-mm").
  - Conserve uniquement les 10 plus recentes sauvegardes.
  - Ecrit des logs dans un fichier dedie.
  - Peut etre execute en mode silencieux (sans boites de dialogue).

.PARAMETER Silent
  Lance la sauvegarde en mode silencieux (pas de boites de dialogue).
#>

# ------------- GESTION DE L'ENCODAGE (Optionnel) -------------
# Si tu as une erreur sur CodePagesEncodingProvider, commente ou supprime ce bloc.
try {
    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
} catch {
    Write-Host "Encodage UTF-8 non configure (version de .NET insuffisante ?) : $($_.Exception.Message)"
}
# -------------------------------------------------------------

param(
    [switch]$Silent  # Si -Silent est specifie, pas de boites de dialogue
)

Add-Type -AssemblyName System.Windows.Forms

# Fichier de log
$logFile = "\\alphatech.local\data\partage\IT\Logs\sauvegarde.log"

# Fonction pour ecrire dans le log avec horodatage
function Write-Log($message) {
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value "$timestamp $message"
}

# -------------------------------------------------------------------------------------
# 1. Si on n’est pas en mode silencieux, on affiche une boite de dialogue de confirmation
# -------------------------------------------------------------------------------------
if (-not $Silent) {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous lancer la sauvegarde ?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::No) {
        Write-Host "Sauvegarde annulee par lutilisateur."
        Write-Log  "Sauvegarde annulee par lutilisateur."
        exit 0
    }
    Write-Host "Sauvegarde confirmee par lutilisateur."
    Write-Log  "Sauvegarde confirmee par lutilisateur."
}
else {
    Write-Host "Sauvegarde lancee en mode silencieux (WSL/Automatique)."
    Write-Log  "Sauvegarde lancee en mode silencieux (WSL/Automatique)."
}

# -------------------------------------------------------------------------------------
# 2. Parametres de la sauvegarde
# -------------------------------------------------------------------------------------
$source      = "\\alphatech.local\data\partage"
$destination = "\\10.11.11.201\recovery"
# Ajout de HH-mm (heure-minute) au format
$date        = Get-Date -Format "yyyy-MM-dd_HH-mm"
$backupName  = "backup_$date.zip"
$backupPath  = Join-Path $destination $backupName

Write-Host "Demarrage de la sauvegarde..."
Write-Log  "Demarrage de la sauvegarde..."

# -------------------------------------------------------------------------------------
# 3. Creation de larchive ZIP + Limitation du nombre de sauvegardes
# -------------------------------------------------------------------------------------
try {
    # a) Creation de larchive ZIP
    Compress-Archive -Path "$source\*" -DestinationPath $backupPath -Force

    # b) Verification de la creation
    if (Test-Path $backupPath) {
        Write-Host "Sauvegarde creee : $backupPath"
        Write-Log  "Sauvegarde creee : $backupPath"

        # c) Limiter a 10 sauvegardes
        $allBackups = Get-ChildItem -Path $destination -Filter "backup_*.zip" -File |
                      Sort-Object LastWriteTime -Descending

        if ($allBackups.Count -gt 10) {
            # On conserve les 10 plus recentes, on supprime le reste
            $toRemove = $allBackups | Select-Object -Skip 10
            foreach ($old in $toRemove) {
                Write-Host "Suppression de lancienne sauvegarde : $($old.Name)"
                Write-Log  "Suppression de lancienne sauvegarde : $($old.FullName)"
                Remove-Item $old.FullName -Force
            }
        }

        # d) Message de succes si pas en mode silencieux
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show(
                "Sauvegarde terminee avec succes !",
                "Succes",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    else {
        throw "Le fichier de sauvegarde na pas ete cree."
    }
}
catch {
    $errMsg = "Echec de la sauvegarde : $($_.Exception.Message)"
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
    exit 1
}

# -------------------------------------------------------------------------------------
# 4. Fin du script
# -------------------------------------------------------------------------------------
Write-Log "Fin du script."
exit 0
