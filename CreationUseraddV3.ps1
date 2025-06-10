<# =====================================================================
.SYNOPSIS
    Importation massive d’utilisateurs Active Directory depuis un CSV.

.DESCRIPTION
    • Sélectionne un fichier CSV (GUI ou -CsvPath),
    • Crée le compte dans l’OU cible,
    • Génère un mot de passe complexe + date d’expiration (par défaut +1 an),
    • Ajoute l’utilisateur au groupe portant le nom de son service (créé si absent),
    • Journalise dans un fichier texte et/ou JSON,
    • Affiche une barre de progression et la durée totale d’exécution.

.EXAMPLE
    # Mode silencieux (pas de GUI)
    .\Import-ADUsers.ps1 -CsvPath "C:\temp\nouveaux.csv" -OUBase "OU=Stagiaires,DC=alphatech,DC=local" -ExpirationYears 2

.EXAMPLE
    # Lancement interactif : une boîte de dialogue vous laisse choisir le CSV
    .\Import-ADUsers.ps1
===================================================================== #>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [string]$OUBase          = 'OU=Utilisateurs,DC=alphatech,DC=local',

    [ValidateRange(1, 10)]
    [int]$ExpirationYears    = 1,

    [string]$LogFile         = 'O:\Direction\RH\ImportationRH\UserCreationLog.txt',

    [string]$JsonLog,

    [char]$Delimiter         = ';',

    [switch]$NoGui
)

# ------------------------------ Pré-requis ------------------------------
# 1. Module RSAT AD obligatoire
Import-Module ActiveDirectory -ErrorAction Stop

# 2. StrictMode pour assainir la syntaxe
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 3. Vérifie/applique le mode STA quand la GUI est demandée
if (-not $PSCmdlet.MyInvocation.ExpectingInput -and
    -not [threading.thread]::CurrentThread.GetApartmentState() -eq 'STA' -and
    -not $NoGui) {

    Write-Verbose 'Redémarrage du script en STA pour la boîte de dialogue…'
    & powershell.exe -STA -ExecutionPolicy Bypass `
        -File $PSCommandPath @PSBoundParameters
    return
}

# ------------------------- Fonctions utilitaires ------------------------

function New-RandomPassword {
    param([int]$Length = 16)
    $chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+'
    -join (1..$Length | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

function Select-CsvFile {
    if ($NoGui) {
        throw 'Le paramètre -NoGui est actif sans -CsvPath : impossible de continuer.'
    }
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.Application]::EnableVisualStyles()

    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Title            = 'Sélectionnez un fichier CSV'
    $dlg.Filter           = 'Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*'
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')

    if ($dlg.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        return $dlg.FileName
    }
    throw 'Opération annulée par utilisateur.'
}

function Test-CsvHeaders {
    param(
        [string[]]$Headers,
        [string[]]$Required
    )
    $missing = $Required | Where-Object { $_ -notin $Headers }
    if ($missing) {
        throw "En-têtes CSV manquants : $($missing -join ', ')"
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    $stamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    Add-Content -Path $LogFile -Value "[$stamp][$Level] $Message"
}

function Write-JsonLog {
    param($Object)
    if ($JsonLog) {
        $Object | ConvertTo-Json -Depth 5 | Add-Content -Path $JsonLog
    }
}

# --------------------------- Fonction principale ------------------------

function Import-AdUsers {
    param(
        [string]$CsvPath,
        [string]$OUBase,
        [int]$ExpirationYears
    )

    Write-Host "`nLecture du fichier CSV : $CsvPath`n" -ForegroundColor Cyan
    $csv = Import-Csv -Path $CsvPath -Delimiter $Delimiter
    Test-CsvHeaders -Headers $csv[0].PSObject.Properties.Name -Required @(
        'Name','DisplayName','GivenName','Surname',
        'SamAccountName','UserPrincipalName','EmailAddress',
        'Department','Title','TelephoneNumber','StreetAddress',
        'POBox','PostalCode','StateOrProvince','Country'
    )

    $total = $csv.Count
    $sw = [Diagnostics.Stopwatch]::StartNew()

    for ($i = 0; $i -lt $total; $i++) {
        $u = $csv[$i]
        Write-Progress -Activity 'Création des comptes AD' `
                        -Status ("{0}/{1} : {2}" -f ($i+1),$total,$u.SamAccountName) `
                        -PercentComplete ((($i+1)/$total)*100)

        try {
            # --- Existence du compte ------------------------------------------------
            if (Get-ADUser -Filter "SamAccountName -eq '$($u.SamAccountName)'" `
                            -SearchBase $OUBase -ErrorAction SilentlyContinue) {
                Write-Log "Utilisateur $($u.SamAccountName) déjà existant – ignoré." 'WARN'
                continue
            }

            # --- Préparation des attributs -----------------------------------------
            $plainPwd  = New-RandomPassword
            $securePwd = ConvertTo-SecureString $plainPwd -AsPlainText -Force
            $expiry    = (Get-Date).AddYears($ExpirationYears)

            $params = @{
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
                Description           = "$($u.Title), $($u.Department) • Import CSV"
            }

            # --- Création de l’utilisateur -----------------------------------------
            New-ADUser @params

            # --- Gestion du groupe --------------------------------------------------
            $grp = Get-ADGroup -Filter "Name -eq '$($u.Department)'" -ErrorAction SilentlyContinue
            if (-not $grp) {
                $grp = New-ADGroup -Name $u.Department -GroupScope Global -Path $OUBase `
                                    -Description "Groupe créé automatiquement pour le service $($u.Department)"
                Write-Log "Groupe $($u.Department) créé." 'INFO'
            }
            Add-ADGroupMember -Identity $grp -Members $u.SamAccountName

            # --- Logs ---------------------------------------------------------------
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
        
    
        $sw.Stop()
            Write-Host ("`nTraitement terminé en {0}s – consultez :`n -> {1}" `
                        -f ([math]::Round($sw.Elapsed.TotalSeconds,2)), $LogFile) -ForegroundColor Green
        }
    }


# =========================== MAIN SCRIPT ===============================
try {
    # Transcription (log détaillé dans %TEMP%)
    Start-Transcript -Path "$env:TEMP\Import-ADUsers-$(Get-Date -Format yyyyMMdd_HHmmss).log" -NoClobber

    if (-not $CsvPath) { $CsvPath = Select-CsvFile }

    # Nettoyage des anciens logs
    if (Test-Path $LogFile) { Remove-Item $LogFile -Force }
    if ($JsonLog  -and (Test-Path $JsonLog)) { Remove-Item $JsonLog -Force }

    Import-AdUsers -CsvPath $CsvPath -OUBase $OUBase -ExpirationYears $ExpirationYears
}
finally {
    Stop-Transcript | Out-Null
}
