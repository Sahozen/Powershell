# Import du module Active Directory
Import-Module ActiveDirectory

# Chemin complet du fichier CSV contenant les informations des utilisateurs
$cheminCsv = "C:..."

# Importation du fichier CSV
$utilisateurs = Import-Csv -Path $chemincsv

# Parcours de chaque utilisateur du CSV
foreach ($utilisateur in $utilisateur) {
  # Création du nom complet
  $nomComplet = "$($utilisateur.FirstName) $($utilisateur.LastName)"

   # Conversion du mot de passe en SecureString requis par New-ADUser
   $motDePasse = ConvertTo-SecureString $utilisateur.Password -AsPlainText -Force
   # Définition de l'OU de destination (à adapter selon votre AD)
      $cheminOU = "OU=Utilisateurs,DC=alphatech,DC=local"

    # Création du compte utilisateur dans l'Active Directory
    New-ADUser `
        -Name $nomComplet `
        -GivenName $utilisateur.FirstName `
        -Surname $utilisateur.LastName `
        -SamAccountName $utilisateur.Username `
        -UserPrincipalName "$($utilisateur.Username)@alpahtech.local" `
        -AccountPassword $motDePasse `
        -Enabled $true `
        -EmailAddress $utilisateur.Email `
        -ChangePasswordAtLogon $true `
        -Path $cheminOU
}
