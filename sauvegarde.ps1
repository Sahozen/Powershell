# Charge l'assembly pour utiliser les boîtes de dialogue
Add-Type -AssemblyName System.Windows.Forms

# Boîte de dialogue de confirmation
$confirmation = [System.Windows.Forms.MessageBox]::Show(
    "Voulez-vous lancer la sauvegarde ?", 
    "Confirmation", 
    [System.Windows.Forms.MessageBoxButtons]::YesNo, 
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($confirmation -eq [System.Windows.Forms.DialogResult]::No) {
    # Si l'utilisateur clique sur "Non", on quitte le script
    Write-Host "Sauvegarde annulée."
    exit
}

# Définition des chemins
$source      = "\\alphatech.local\data\partage"
$destination = "\\10.11.11.201\recovery"
$date        = Get-Date -Format "yyyy-MM-dd"    # Format de la date (ex: 2025-02-19)
$backupName  = "backup_$date.zip"
$backupPath  = Join-Path $destination $backupName

# Supprimer les anciennes sauvegardes (celles qui datent d'avant aujourd'hui)
Write-Host "Recherche et suppression des anciennes sauvegardes..."
$oldBackups = Get-ChildItem -Path $destination -Filter "backup_*.zip" -File -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).Date }

foreach ($backup in $oldBackups) {
    Remove-Item $backup.FullName -Force
    Write-Host "Supprimé : $($backup.FullName)"
}

# Vérifier si la sauvegarde du jour existe déjà
if (Test-Path $backupPath) {
    Write-Host "Une sauvegarde existe déjà pour aujourd'hui. Suppression..."
    Remove-Item $backupPath -Force
}

# Création de l'archive ZIP
Write-Host "Création de la sauvegarde compressée..."
try {
    Compress-Archive -Path "$source\*" -DestinationPath $backupPath -Force
    if (Test-Path $backupPath) {
        Write-Host "Sauvegarde réussie : $backupPath"
        [System.Windows.Forms.MessageBox]::Show(
            "Sauvegarde terminée avec succès !", 
            "Succès", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } else {
        throw "Le fichier de sauvegarde n'a pas été créé."
    }
}
catch {
    Write-Host "Échec de la sauvegarde : $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show(
        "Échec de la sauvegarde : $($_.Exception.Message)", 
        "Erreur", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
