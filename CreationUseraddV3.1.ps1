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
$openFileDialog.InitialDirectory = "O:\Direction\RH\ImportationRH"

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
# Horodatage : AAAAMMJJ_HHMM
$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'

# Chemin du fichier de log (ex. UserCreationLog_20250610_1432.txt)
$logFile = "O:\Direction\RH\ImportationRH\UserCreationLog_$timestamp.txt"

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

# --- Si l'utilisateur existe déjà -----------------------------------
if ($existingUser) {
    Write-Host "L'utilisateur '$($utilisateur.Name)' existe déjà (SamAccountName = $sam)." -ForegroundColor Green
    continue
}

...

# --- Ajout au groupe ------------------------------------------------
if ($grp) {
    Add-ADGroupMember -Identity $grp -Members $adUser
    Write-Host "Ajout au groupe '$($grp.Name)' effectué." -ForegroundColor Green
} else {
    Write-Host "Groupe '$($utilisateur.Department)' introuvable : pas d'ajout." -ForegroundColor Green   # ← vert au lieu de jaune
}


foreach ($utilisateur in $utilisateurs) {
    try {
        # --- Variables locales ---
        $sam = $utilisateur.SamAccountName.Trim()
        $expirationDate = (Get-Date).AddYears(1)

        # --- L'utilisateur existe déjà ? ---
        $existingUser = Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue
        if ($existingUser) {
            Write-Host "L'utilisateur '$($utilisateur.Name)' existe déjà (SamAccountName = $sam)." -ForegroundColor Green
            continue
        }

        # --- Paramètres de création ---
        $params = @{
            Name                  = $utilisateur.Name
            DisplayName           = $utilisateur.DisplayName
            GivenName             = $utilisateur.GivenName
            Surname               = $utilisateur.Surname
            SamAccountName        = $sam
            UserPrincipalName     = $utilisateur.UserPrincipalName
            EmailAddress          = $utilisateur.EmailAddress
            Department            = $utilisateur.Department
            Title                 = $utilisateur.Title
            OfficePhone           = $utilisateur.TelephoneNumber
            StreetAddress         = $utilisateur.StreetAddress
            POBox                 = $utilisateur.POBox
            PostalCode            = $utilisateur.PostalCode
            State                 = $utilisateur.StateOrProvince
            Country               = $utilisateur.Country
            AccountPassword       = $motDePasse
            Enabled               = $true
            ChangePasswordAtLogon = $true
            AccountExpirationDate = $expirationDate
            Path                  = $ouBase
            Description           = "Poste : $($utilisateur.Title) | Service : $($utilisateur.Department) | Import CSV $(Get-Date -Format 'd')"
        }

        # --- Création du compte (+ PassThru) ---
        $adUser = New-ADUser @params -PassThru

        Write-Host "Création réussie : $($utilisateur.Name)" -ForegroundColor Green
        Write-Host "  -> Login : $sam" -ForegroundColor Green
        Write-Host "  -> Mot de passe initial : $passwordPlain" -ForegroundColor Green
        Write-Host "  -> Expire le : $($expirationDate.ToShortDateString())" -ForegroundColor Green

        # --- Ajout au groupe ---
        $grp = Get-ADGroup -Filter "Name -eq '$($utilisateur.Department.Trim())'" -ErrorAction SilentlyContinue
        if ($grp) {
            Add-ADGroupMember -Identity $grp -Members $adUser
            Write-Host "Ajout au groupe '$($grp.Name)' effectué." -ForegroundColor Green
        } else {
            Write-Host "Groupe '$($utilisateur.Department)' introuvable : pas d'ajout." -ForegroundColor Green
        }

        # --- Log ---
        $logEntry = "[{0}] Login: {1} | Mot de passe: {2}" -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'), $sam, $passwordPlain
        Add-Content -Path $logFile -Value $logEntry
    }
    catch {
        Write-Error "Erreur pour l'utilisateur '$($utilisateur.Name)' : $_"
    }

    Start-Sleep -Seconds 1
}
Write-Host "Traitement terminé. Consultez le fichier de log : $logFile" -ForegroundColor Cyan
