<# =====================================================================
.SYNOPSIS
    Importation manuelle d’utilisateurs AD à partir d’un CSV.
.DESCRIPTION
    • Affiche une boîte de dialogue pour choisir le CSV si -CsvPath absent.
    • Crée le compte dans l’OU cible et génère un mot de passe complexe.
    • Ajoute l’utilisateur au groupe portant le nom de son service
      (créé automatiquement si nécessaire).
    • Journalise chaque action dans un fichier texte basique.
.EXAMPLE
    # Lancement interactif (GUI)
    .\Import-ADUsers.ps1
.EXAMPLE
    # Lancement direct sans GUI
    .\Import-ADUsers.ps1 -CsvPath "C:\temp\nouveaux.csv"
===================================================================== #>

[CmdletBinding()]
param(
    [string]$CsvPath,
    [string]$OUBase = 'OU=Utilisateurs,DC=alphatech,DC=local',
    [ValidateRange(1,10)]
    [int]   $ExpirationYears = 1,
    [string]$LogFile = 'O:\Direction\RH\ImportationRH\UserCreationLog.txt',
    [char]  $Delimiter = ';'
)

#----------------------------- Pré-requis ------------------------------
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory -ErrorAction Stop      # RSAT requis

# Redémarrage éventuel en STA pour la GUI
if (-not [Threading.Thread]::CurrentThread.ApartmentState -eq 'STA' -and -not $CsvPath) {
    Write-Verbose 'Passage en STA pour la boîte de dialogue…'
    & powershell.exe -STA -ExecutionPolicy Bypass -File $PSCommandPath @PSBoundParameters
    return
}

#--------------------------- Fonctions ---------------------------------
function New-RandomPassword {
    param([int]$Length = 16)
    $c = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+'
    -join (1..$Length | ForEach-Object { $c[(Get-Random -Maximum $c.Length)] })
}

function Select-CsvFile {
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.Application]::EnableVisualStyles()
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Title  = 'Sélectionnez un fichier CSV'
    $dlg.Filter = 'Fichier CSV (*.csv)|*.csv'
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.FileName }
    throw 'Opération annulée par l’utilisateur.'
}

function Write-Log ([string]$Msg,[string]$Lvl='INFO') {
    '{0} [{1}] {2}' -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'),$Lvl,$Msg |
        Add-Content -Path $LogFile
}

function Validate-CsvHeaders ($Headers) {
    $required = 'Name','DisplayName','GivenName','Surname',
                'SamAccountName','UserPrincipalName','EmailAddress',
                'Department','Title','TelephoneNumber','StreetAddress',
                'POBox','PostalCode','StateOrProvince','Country'
    $missing = $required | Where-Object { $_ -notin $Headers }
    if ($missing) { throw "En-têtes manquants : $($missing -join ', ')" }
}

#--------------------------- Import AD ---------------------------------
function Import-AdUsers {
    param($Csv)

    $total = $Csv.Count
    $sw    = [Diagnostics.Stopwatch]::StartNew()

    for ($i = 0; $i -lt $total; $i++) {
        $u = $Csv[$i]
        Write-Progress -Activity 'Import AD' -Status "$($i+1)/$total : $($u.SamAccountName)" `
            -PercentComplete ((($i+1)/$total)*100)

        try {
            if (Get-ADUser -Filter "SamAccountName -eq '$($u.SamAccountName)'" `
                           -SearchBase $OUBase -EA SilentlyContinue) {
                Write-Log "Déjà existant : $($u.SamAccountName)" 'WARN'; continue
            }

            $pwdPlain  = New-RandomPassword
            $pwdSecure = ConvertTo-SecureString $pwdPlain -AsPlainText -Force
            $expiry    = (Get-Date).AddYears($ExpirationYears)

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
                AccountPassword       = $pwdSecure
                Enabled               = $true
                ChangePasswordAtLogon = $true
                AccountExpirationDate = $expiry
                Path                  = $OUBase
                Description           = "$($u.Title), $($u.Department)"
            }

            # Groupe (création si absent)
            $grp = Get-ADGroup -Filter "Name -eq '$($u.Department)'" -EA SilentlyContinue
            if (-not $grp) {
                $grp = New-ADGroup -Name $u.Department -GroupScope Global -Path $OUBase `
                       -Description "Groupe auto ($($u.Department))"
            }
            Add-ADGroupMember -Identity $grp -Members $u.SamAccountName

            Write-Log "Créé : $($u.SamAccountName) | exp. $($expiry.ToShortDateString())"
        }
        catch {
            Write-Log "Erreur : $($u.SamAccountName) -> $_" 'ERROR'
        }
    }

    $sw.Stop()
    Write-Host ("`nTerminé en {0}s (voir {1})" -f [math]::Round($sw.Elapsed.TotalSeconds,1),$LogFile) `
        -ForegroundColor Green
}

#------------------------------ MAIN ------------------------------------
try {
    Start-Transcript -Path "$env:TEMP\ImportAD-$(Get-Date -f yyyyMMdd_HHmmss).log" -NoClobber

    if (-not $CsvPath) { $CsvPath = Select-CsvFile }
    if (Test-Path $LogFile) { Remove-Item $LogFile -Force }

    $csv = Import-Csv -Path $CsvPath -Delimiter $Delimiter
    Validate-CsvHeaders $csv[0].PSObject.Properties.Name

    Import-AdUsers -Csv $csv
}
finally { Stop-Transcript | Out-Null }
