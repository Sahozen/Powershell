# Backlog en attente entre les deux DC
Get-DfsrBacklog -GroupName "Domain System Volume" `
                -FolderName "SYSVOL Share" `
                -SourceComputerName WIN-PDC `
                -DestinationComputerName WIN-ADC


wmic /namespace:\\root\microsoftdfs path dfsrMachineConfig get MaxOfflineTimeInDays
wmic /namespace:\\root\microsoftdfs path dfsrMachineConfig set MaxOfflineTimeInDays=700


Restart-Service DFSR -Force
