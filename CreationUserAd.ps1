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
$cheminCsv = "O:\chemin\vers\nouveaux_utilisateurs.csv"

# Importation du fichier CSV
$utilisateurs = Import-Csv -Path $cheminCsv

# Parcours de chaque utilisateur du CSV
foreach ($utilisateur in $utilisateurs) {
    try {
        # Création du nom complet et de l'adresse e-mail
        $prenom = $utilisateur.FirstName.Trim()
        $nom = $utilisateur.LastName.Trim()
        $nomComplet = "$prenom $nom"
        $email = "$prenom.$nom@alphatech.local".ToLower()

        # Génération d'un mot de passe aléatoire de 16 caractères
        $passwordPlain = New-RandomPassword -length 16
        $motDePasse = ConvertTo-SecureString $passwordPlain -AsPlainText -Force
        
        # Définition de l'OU de destination (à adapter selon votre AD)
        $cheminOU = "OU=Utilisateurs,DC=alphatech,DC=local"

        # Création du compte utilisateur dans l'Active Directory
        New-ADUser `
            -Name $nomComplet `
            -GivenName $prenom `
            -Surname $nom `
            -SamAccountName $utilisateur.Username `
            -UserPrincipalName $email `
            -AccountPassword $motDePasse `
            -Enabled $true `
            -EmailAddress $email `
            -ChangePasswordAtLogon $true `
            -Path $cheminOU

        # Confirmation de la création du compte
        Write-Host "L'utilisateur '$nomComplet' a été créé avec succès." -ForegroundColor Green
        Write-Host "  -> Login : $($utilisateur.Username)" -ForegroundColor Green
        Write-Host "  -> Mot de passe initial : $passwordPlain" -ForegroundColor Green

        # Affectation de l'utilisateur à un groupe correspondant à son service
        $nomGroupe = $utilisateur.Service.Trim()
        Add-ADGroupMember -Identity $nomGroupe -Members $utilisateur.Username
        Write-Host "L'utilisateur '$nomComplet' a été ajouté au groupe '$nomGroupe'." -ForegroundColor Green
    }
    catch {
        Write-Error "Erreur lors de la création ou de l'affectation de l'utilisateur '$nomComplet' : $_"
    }
}
