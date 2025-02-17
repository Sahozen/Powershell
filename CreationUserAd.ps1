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

# Indiquez le chemin vers votre fichier CSV
$cheminCsv = "C:\Import\Testimportation.csv"

# Importation du fichier CSV
$utilisateurs = Import-Csv -Path $cheminCsv

# OU de destination dans l'AD
$ouBase = "OU=Utilisateurs,DC=alphatech,DC=local"

foreach ($utilisateur in $utilisateurs) {
    try {
        # Génération d'un mot de passe aléatoire
        $passwordPlain = New-RandomPassword -length 16
        $motDePasse = ConvertTo-SecureString $passwordPlain -AsPlainText -Force
        
        # Création du compte utilisateur
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
    }
    catch {
        Write-Error "Erreur pour l'utilisateur '$($utilisateur.Name)' : $_"
    }
    
    # Petite pause (optionnelle) pour éviter d'éventuels conflits/latences
    Start-Sleep -Seconds 1
}
