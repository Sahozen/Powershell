<# =====================================================================
.SYNOPSIS
    Importation massive d’utilisateurs AD depuis un CSV.

.DESCRIPTION
    • Sélectionne un fichier CSV (GUI ou paramètre -CsvPath),
    • Crée les comptes dans l’OU cible,
    • Génère un mot de passe complexe + définition d’une date d’expiration (par défaut +1 an),
    • Ajoute l’utilisateur à un groupe dont le nom correspond à son service,
    • Journalise dans un fichier texte (et un JSON optionnel),
    • Affiche une barre de progression et la durée totale d’exécution.

.EXAMPLE
    .\Import-ADUsersV3.ps1 -CsvPath "C:\temp\nouveaux.csv" -OUBase "OU=Stagiaires,DC=alphatech,DC=local" -ExpirationYears 2
===================================================================== #>

# Prérequis : PowerShell 5.1 + module RSAT ActiveDirectory
# ----------------------------------------------------------------------------
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [string]$OUBase          = "OU=Utilisateurs,DC=alphatech,DC=local",

    [ValidateRange(1, 10)]
    [int]$ExpirationYears    = 1,

    [string]$LogFile         = "O:\Direction\RH\ImportationRH\UserCreationLog.txt",

    [string]$JsonLog,

    [char]$Delimiter         = ';',

    [switch]$NoGui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Module AD
Import-Module ActiveDirectory -ErrorAction Stop

# ----------------------------------------------------------------------------
function New-RandomPassword {
    param([int]$Length = 16)

    $chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+'
    -join (1..$Length | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
}

# ----------------------------------------------------------------------------
function Select-CsvFile {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        [System.Windows.Forms.Application]::EnableVisualStyles()

        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title            = "Sélectionnez un fichier CSV"
        $dlg.Filter           = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
        $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dlg.FileName
        }
        throw "Opération annulée par l'utilisateur."
    }
    catch {
        throw "Impossible d’ouvrir la boîte de dialogue de sélection : $_"
    }
}


# ----------------------------------------------------------------------------
function Start-TranscriptSafe {
    try {
        $scriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)
        Start-Transcript -Path "$env:TEMP\$scriptName-$(Get-Date -Format yyyyMMdd_HHmmss).log" -NoClobber
    }
    catch {
        Write-Warning "Transcript non initialisé : $_"
    }
}

# ----------------------------------------------------------------------------
function Validate-CsvHeaders {
    param(
        [string[]]$Headers,
        [string[]]$Required
    )

    $missing = $Required | Where-Object { $_ -notin $Headers }
    if ($missing) {
        throw "Le CSV est invalide ; en-têtes manquants : $($missing -join ', ')"
    }
}

# ----------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path $LogFile -Value "[$stamp][$Level] $Message"
}

# ----------------------------------------------------------------------------
function Write-JsonLog {
    param($Object)

    if ($JsonLog) {
        $Object | ConvertTo-Json -Depth 5 | Add-Content -Path $JsonLog
    }
}

# ----------------------------------------------------------------------------
function Import-AdUsers {
    param(
        [string]$CsvPath,
        [string]$OUBase,
        [int]$ExpirationYears
    )

    Write-Host "\nLecture du fichier CSV : $CsvPath\n" -ForegroundColor Cyan

    $csv = Import-Csv -Path $CsvPath -Delimiter $Delimiter
    Validate-CsvHeaders -Headers $csv[0].PSObject.Properties.Name -Required @(
        'Name','DisplayName','GivenName','Surname',
        'SamAccountName','UserPrincipalName','EmailAddress',
        'Department','Title','TelephoneNumber','StreetAddress',
        'POBox','PostalCode','StateOrProvince','Country'
    )

    $total = $csv.Count
    $sw = [Diagnostics.Stopwatch]::StartNew()

    for ($i = 0; $i -lt $total; $i++) {
        $u = $csv[$i]
        Write-Progress -Activity "Création de comptes Active Directory" `
                       -Status "$($i + 1)/$total : $($u.SamAccountName)" `
                       -PercentComplete ((($i + 1) / $total) * 100)

        try {
            if (Get-ADUser -Filter "SamAccountName -eq '$($u.SamAccountName)'" `
                           -SearchBase $OUBase -ErrorAction SilentlyContinue) {
                Write-Log "Utilisateur $($u.SamAccountName) déjà existant – ignoré." 'WARN'
                continue
            }

            $plainPwd  = New-RandomPassword
            $securePwd = ConvertTo-SecureString $plainPwd -AsPlainText -Force
            $expiry    = (Get-Date).AddYears($ExpirationYears)

            # -- Création du compte -------------------------------------------------
            New-ADUser @{
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
                Description           = "$($u.Title), $($u.Department) – Import CSV"
            }

            # -- Ajout au groupe ---------------------------------------------------
            try {
                Add-ADGroupMember -Identity $u.Department -Members $u.SamAccountName -ErrorAction Stop
            }
            catch {
                Write-Log "Le groupe $($u.Department) est introuvable." 'WARN'
            }

            Write-Log "Création $($u.SamAccountName) OK (expire $($expiry.ToShortDateString()))"
            Write-JsonLog @{
                SamAccountName = $u.SamAccountName
                Password       = $plainPwd
                Expires        = $expiry
                Status         = 'Created'
            }
        }
        catch {
            Write-Log "Erreur pour $($u.SamAccountName) : $_" 'ERROR'
            Write-JsonLog @{
                SamAccountName = $u.SamAccountName
                Error          = $_.Exception.Message
                Status         = 'Failed'
            }
        }
    }

    $sw.Stop()
    Write-Host "\nTraitement terminé en $([Math]::Round($sw.Elapsed.TotalSeconds, 2)) s – consultez : `n -> $LogFile" -ForegroundColor Green
}

# =====================  MAIN  ======================================
try {
    Start-TranscriptSafe

    if (-not $CsvPath) {
        $CsvPath = Select-CsvFile
    }

    # Purge éventuelle d'anciens logs
    if (Test-Path $LogFile) { Remove-Item $LogFile -Force }
    if ($JsonLog -and (Test-Path $JsonLog)) { Remove-Item $JsonLog -Force }

    Import-AdUsers -CsvPath $CsvPath -OUBase $OUBase -ExpirationYears $ExpirationYears
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
}
