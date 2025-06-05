<# =====================================================================
.SYNOPSIS
    Importation massive d’utilisateurs AD depuis un CSV.

.DESCRIPTION
    - Sélectionne un fichier CSV (GUI ou paramètre -CsvPath),
    - Crée les comptes dans l’OU cible,
    - Génère mot de passe complexe + expiration (par défaut 1 an),
    - Ajoute l’utilisateur à un groupe portant le nom de son service,
    - Loggue dans un fichier texte (et JSON facultatif),
    - Affiche une barre de progression et la durée totale d’exécution.

.PARAMETER CsvPath
    Chemin du fichier CSV. Si absent, une boîte de dialogue s’ouvre (sauf -NoGui).

.PARAMETER OUBase
    Chemin LDAP de l’OU cible (par défaut "OU=Utilisateurs,DC=alphatech,DC=local").

.PARAMETER ExpirationYears
    Durée de vie du compte en années (défaut 1).

.PARAMETER LogFile
    Chemin du fichier log texte (défaut "O:\Direction\RH\ImportationRH\UserCreationLog.txt").

.PARAMETER JsonLog
    Chemin d’un log JSON optionnel (défaut : aucun).

.PARAMETER Delimiter
    Délimiteur du CSV (défaut ';').

.PARAMETER NoGui
    Supprime toute interaction graphique (nécessite -CsvPath sinon erreur).

.EXAMPLE
    .\Import-ADUsersV3.ps1 -CsvPath ".\nouveaux.csv" -OUBase "OU=Stagiaires,DC=alphatech,DC=local" -ExpirationYears 2

.EXAMPLE
    .\Import-ADUsersV3.ps1 -NoGui       # s’attend à recevoir -CsvPath, sinon erreur
===================================================================== #>

param(
    [string]$CsvPath,
    [string]$OUBase          = "OU=Utilisateurs,DC=alphatech,DC=local",
    [int]   $ExpirationYears = 1,
    [string]$LogFile         = "O:\Direction\RH\ImportationRH\UserCreationLog.txt",
    [string]$JsonLog,
    [char]  $Delimiter       = ';',
    [switch]$NoGui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#---------------------------------------------------------------
function New-RandomPassword {
    param([int]$Length = 16)
    $chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+'
    -join (1..$Length | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
}

#---------------------------------------------------------------
function Select-CsvFile {
    if ($NoGui) { throw "Aucun -CsvPath fourni et -NoGui spécifié. Opération annulée." }

    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.Application]::EnableVisualStyles()
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Title            = "Sélectionnez un fichier CSV"
    $dlg.Filter           = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $dlg.InitialDirectory = "C:\"
    return if ($dlg.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) { $dlg.FileName } else { throw "Opération annulée par l'utilisateur." }
}

#---------------------------------------------------------------
function Start-TranscriptSafe {
    try   { $scriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)
            Start-Transcript -Path "$env:TEMP\$scriptName-$(Get-Date -Format yyyyMMdd_HHmmss).log" -NoClobber }
    catch { Write-Warning "Transcript non initialisé : $_" }
}

#---------------------------------------------------------------
function Validate-CsvHeaders {
    param($Headers, [string[]]$Required)
    $missing = $Required | Where-Object { $_ -notin $Headers }
    if ($missing) { throw "Le CSV est invalide ; en-têtes manquants : $($missing -join ', ')" }
}

#---------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path $LogFile -Value "[$stamp][$Level] $Message"
}

#---------------------------------------------------------------
function Write-JsonLog {
    param($Object)
    if ($JsonLog) {
        $Object | ConvertTo-Json -Depth 5 | Add-Content -Path $JsonLog
    }
}

#---------------------------------------------------------------
function Import-AdUsers {
    param(
        [string]$CsvPath,
        [string]$OUBase,
        [int]   $ExpirationYears
    )

    #-- Lecture CSV
    Write-Host "Lecture du fichier CSV : $CsvPath" -ForegroundColor Cyan
    $csv = Import-Csv -Path $CsvPath -Delimiter $Delimiter
    Validate-CsvHeaders -Headers $csv[0].PsObject.Properties.Name -Required @(
        'Name','DisplayName','GivenName','Surname',
        'SamAccountName','UserPrincipalName','EmailAddress',
        'Department','Title','TelephoneNumber','StreetAddress',
        'POBox','PostalCode','StateOrProvince','Country'
    )

    $total = $csv.Count
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()

    for ($i = 0; $i -lt $total; $i++) {
        $u = $csv[$i]
        Write-Progress -Activity "Création comptes AD" -Status "$($i+1)/$total : $($u.SamAccountName)" -PercentComplete (($i+1)/$total*100)

        try {
            if (Get-ADUser -Filter "SamAccountName -eq '$($u.SamAccountName)'" -SearchBase $OUBase -ErrorAction SilentlyContinue) {
                Write-Log "Utilisateur $($u.SamAccountName) déjà existant – ignoré." 'WARN'
                continue
            }

            $plainPwd  = New-RandomPassword
            $securePwd = ConvertTo-SecureString $plainPwd -AsPlainText -Force
            $expiry    = (Get-Date).AddYears($ExpirationYears)

            $newParams = @{
                Name                  = $u.Name
                DisplayName           = $u.DisplayName
                GivenName             = $u.GivenName
                Surname               = $u.Surname
                SamAccountName        = $u.SamAccountName
                UserPrincipalName     = $u.UserPrincipalName
                EmailAddress          = $u.EmailAddress
                Department            = $u.Department
                Title                 = $u.Title
                OfficePhone           = $u.TelephoneNumber
                StreetAddress         = $u.StreetAddress
                POBox                 = $u.POBox
                PostalCode            = $u.PostalCode
                State                 = $u.StateOrProvince
                Country               = $u.Country
                AccountPassword       = $securePwd
                Enabled               = $true
                ChangePasswordAtLogon = $true
                AccountExpirationDate = $expiry
                Path                  = $OUBase
                Description           = "$($u.Title), $($u.Department) Import CSV"
            }

            New-ADUser @newParams

            # Ajout au groupe
            Add-ADGroupMember -Identity $u.Department -Members $u.SamAccountName

            Write-Log "Création $($u.SamAccountName) OK (expire $($expiry.ToShortDateString()))"
            Write-JsonLog @{
                SamAccountName = $u.SamAccountName
                Password       = $plainPwd
                Expires        = $expiry
                Status         = 'Created'
            }
        }
        catch {
            Write-Log "Erreur pour $($u.SamAccountName) : $_" 'ERROR'
            Write-JsonLog @{
                SamAccountName = $u.SamAccountName
                Error          = $_.Exception.Message
                Status         = 'Failed'
            }
        }
    }

    $stopwatch.Stop()
    Write-Host "Traitement terminé en $([Math]::Round($stopwatch.Elapsed.TotalSeconds,2)) s – consultez $LogFile" -ForegroundColor Green
}

#===================== MAIN ==========================================
try {
    Start-TranscriptSafe

    if (-not $CsvPath) { $CsvPath = Select-CsvFile }

    # Nettoyage éventuel du log précédent
    if (Test-Path $LogFile) { Remove-Item $LogFile -Force }
    if ($JsonLog)           { if (Test-Path $JsonLog) { Remove-Item $JsonLog -Force } }

    Import-AdUsers -CsvPath $CsvPath -OUBase $OUBase -ExpirationYears $ExpirationYears
}
finally {
    Stop-Transcript | Out-Null
}

