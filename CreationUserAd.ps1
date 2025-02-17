# Import du module Active Directory
Import-Module ActiveDirectory

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
# 1. Variables de configuration
# -------------------------------------------------------------------

# Chemin complet vers le fichier CSV
$cheminCsv = "O:\Direction\RH\ImportationRH\Testimportation.csv"

# Chemin du fichier de log pour enregistrer les login et mots de passe
$logFile = "O:\Direction\RH\ImportationRH\UserCreationLog.txt"

# OU de destination dans l'AD
$ouBase = "OU=Utilisateurs,DC=alphatech,DC=local"

# -------------------------------------------------------------------
# 2. Préparation du fichier de log
# -------------------------------------------------------------------

# Si le fichier log existe déjà, on le supprime pour repartir d'un log vierge
if (Test-Path $logFile) {
    Remove-Item $logFile
    Write-Host "Fichier de log supprimé : $logFile"
}

# -------------------------------------------------------------------
# 3. Importation du CSV et boucle principale
# -------------------------------------------------------------------

# Lecture du fichier CSV
$utilisateurs = Import-Csv -Path $cheminCsv

foreach ($utilisateur in $utilisateurs) {
    try {
        # Vérifier si l'utilisateur existe déjà (recherche par SamAccountName dans l'OU spécifiée)
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

    # Petite pause pour éviter d'éventuels conflits/latences
    Start-Sleep -Seconds 1
}

Write-Host "Traitement terminé. Consultez le fichier de log : $logFile" -ForegroundColor Cyan
