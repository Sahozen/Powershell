# Import du module Active Directory
Import-Module ActiveDirectory

# Chargement de l'assembly System.Windows.Forms pour l'explorateur de fichiers
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Fonction de génération d'un mot de passe aléatoire
function New-RandomPassword {
    param(
        [int]$length = 16
    )
    $chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+'
    $password = -join ((1..$length) | ForEach-Object {
        $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)]
    })
    return $password
}

# -------------------------------------------------------------------
# 1. Sélection du fichier CSV via l'explorateur
# -------------------------------------------------------------------

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = "Sélectionnez un fichier CSV"
$openFileDialog.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
$openFileDialog.InitialDirectory = "C:\"

$dialogResult = $openFileDialog.ShowDialog()

if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $cheminCsv = $openFileDialog.FileName
    Write-Host "Fichier sélectionné : $cheminCsv" -ForegroundColor Cyan
} else {
    Write-Warning "Opération annulée par l'utilisateur."
    return
}

# -------------------------------------------------------------------
# 2. Préparation du fichier de log
# -------------------------------------------------------------------

# Chemin du fichier de log pour enregistrer les login et mots de passe
$logFile = "O:\Direction\RH\ImportationRH\UserCreationLog.txt"

# Si le fichier log existe déjà, on le supprime pour repartir d'un log vierge
if (Test-Path $logFile) {
    Remove-Item $logFile
    Write-Host "Fichier de log supprimé : $logFile"
}

# -------------------------------------------------------------------
# 3. Paramètres de l'AD
# -------------------------------------------------------------------

# OU de destination dans l'AD
$ouBase = "OU=Utilisateurs,DC=alphatech,DC=local"

# -------------------------------------------------------------------
# 4. Lecture du CSV et création des utilisateurs
# -------------------------------------------------------------------

$utilisateurs = Import-Csv -Path $cheminCsv

foreach ($utilisateur in $utilisateurs) {
    try {
        # Vérifier si l'utilisateur existe déjà
        $existingUser = Get-ADUser -Filter "SamAccountName -eq '$($utilisateur.SamAccountName)'" `
                                   -SearchBase $ouBase `
                                   -SearchScope Subtree `
                                   -ErrorAction SilentlyContinue

        if ($existingUser) {
            Write-Host "L'utilisateur '$($utilisateur.Name)' existe déjà (SamAccountName = $($utilisateur.SamAccountName))." -ForegroundColor Yellow
            continue
        }

        # Génération d'un mot de passe aléatoire
        $passwordPlain = New-RandomPassword -length 16
        $motDePasse = ConvertTo-SecureString $passwordPlain -AsPlainText -Force

        # Création du compte utilisateur dans l'Active Directory
        New-ADUser `
            -Name $utilisateur.Name `
            -DisplayName $utilisateur.DisplayName `
            -GivenName $utilisateur.GivenName `
            -Surname $utilisateur.Surname `
            -SamAccountName $utilisateur.SamAccountName `
            -UserPrincipalName $utilisateur.UserPrincipalName `
            -EmailAddress $utilisateur.EmailAddress `
            -Department $utilisateur.Department `
            -Title $utilisateur.Title `
            -OfficePhone $utilisateur.TelephoneNumber `
            -StreetAddress $utilisateur.StreetAddress `
            -POBox $utilisateur.POBox `
            -PostalCode $utilisateur.PostalCode `
            -State $utilisateur.StateOrProvince `
            -Country $utilisateur.Country `
            -AccountPassword $motDePasse `
            -Enabled $true `
            -ChangePasswordAtLogon $true `
            -Path $ouBase `
            -Description "Utilisateur importé par script CSV - $($utilisateur.Title), $($utilisateur.Department)"

        Write-Host "Création réussie : $($utilisateur.Name)" -ForegroundColor Green
        Write-Host "  -> Login : $($utilisateur.SamAccountName)" -ForegroundColor Green
        Write-Host "  -> Mot de passe initial : $passwordPlain" -ForegroundColor Green

        # Ajouter l'utilisateur au groupe correspondant à son service
        Add-ADGroupMember -Identity $utilisateur.Department -Members $utilisateur.SamAccountName
        Write-Host "Ajout au groupe '$($utilisateur.Department)' effectué." -ForegroundColor Green

        # Écriture dans le fichier log
        $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Login: $($utilisateur.SamAccountName) | Mot de passe: $passwordPlain"
        Add-Content -Path $logFile -Value $logEntry
    }
    catch {
        Write-Error "Erreur pour l'utilisateur '$($utilisateur.Name)' : $_"
    }

    Start-Sleep -Seconds 1
}

Write-Host "Traitement terminé. Consultez le fichier de log : $logFile" -ForegroundColor Cyan
