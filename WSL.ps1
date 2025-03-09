powershell.exe


-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "\\alphatech.local\data\Partage\IT\Scripts\Sauvegarde.ps1"

\\alphatech.local\data\Partage\IT\Scripts\

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

-d Debian -- sudo service cron start


powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "\\alphatech.local\data\Partage\IT\Scripts\Sauvegarde.ps1"

0 2 * * * powershell.exe -File "C:\Scripts\Sauvegarde.ps1" >> /home/<user>/sauvegarde.log 2>&1

0 2 * * * /mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe -File /mnt/c/Scripts/Sauvegarde.ps1 >> /home/<user>/sauvegarde.log 2>&1



