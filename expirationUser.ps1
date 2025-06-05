$expiration = (Get-Date).AddYears(1)      # exactement dans 12 mois

Get-ADUser -Filter * -SearchBase "OU=Utilisateurs,DC=alphatech,DC=local" |
    Set-ADUser -AccountExpirationDate $expiration
