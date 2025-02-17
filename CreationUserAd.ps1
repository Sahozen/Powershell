# Import du module Active Directory
Import-Module ActiveDirectory

# Fonction de génération d'un mot de passe aléatoire
function New-RandomPassword {
    param(
        [int]$length = 16
    )
    $chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+'
    $password = -join ((1..$length) | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
    return $password
}

# Chemin complet du fichier CSV contenant les informations des utilisateurs
$cheminCsv = "O:\Direction\RH\ImportationRH\Testimportation.csv"

# Chemin du fichier de log pour enregistrer les login et mot de passe
$logFile = "O:\Direction\RH\ImportationRH\UserCreationLog.txt"

# Si le fichier log existe déjà, on le supprime pour repartir d'un log vierge
if (Test-Path $logFile) {
    Remove-Item $logFile
}

# Importation du fichier CSV
$utilisateurs = Import-Csv -Path $cheminCsv

# OU de destination dans l'AD
$ouBase = "OU=Utilisateurs,DC=alphatech,DC=local"

foreach ($utilisateur in $utilisateurs) {
    try {
        # Vérifier si l'utilisateur existe déjà en recherchant par SamAccountName
        $existingUser = Get-ADUser -Filter { SamAccountName -eq $($utilisateur.SamAccountName) } -SearchBase $ouBase -ErrorAction SilentlyContinue
        if ($existingUser) {
            Write-Host "L'utilisateur '$($utilisateur.Name)' existe déjà." -ForegroundColor Yellow
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

        # Enregistrement dans le fichier log
        $logEntry = "Login: $($utilisateur.SamAccountName) - Mot de passe: $passwordPlain"
        Add-Content -Path $logFile -Value $logEntry
    }
    catch {
        Write-Error "Erreur pour l'utilisateur '$($utilisateur.Name)' : $_"
    }
    
    # Petite pause pour éviter d'éventuels conflits/latences
    Start-Sleep -Seconds 1
}
