powershell.exe


-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "\\alphatech.local\data\Partage\IT\Scripts\Sauvegarde.ps1"

\\alphatech.local\data\Partage\IT\Scripts\

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "\\alphatech.local\data\Partage\IT\Scripts\Sauvegarde.ps1"

0 2 * * * powershell.exe -File /mnt/o/NT/Scripts/Sauvegarde.ps1 >> /home/<votre_user>/sauvegarde.log 2>&1


