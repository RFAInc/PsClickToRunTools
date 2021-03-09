# PsClickToRunTools
Toolset for working with Office Click-to-Run updates

# Examples

## Audit your AD
```
# Import standard modules
ipmo ActiveDirectory
# Import my modules
$web=New-Object Net.WebClient
$web.DownloadString('https://raw.githubusercontent.com/RFAInc/PsClickToRunTools/main/PsClickToRunTools.psm1')|iex;
$web.DownloadString('https://raw.githubusercontent.com/tonypags/PsWinAdmin/master/Get-InstalledSoftware.ps1')|iex;
# Get computer list
$PCs=Get-AdComputer -f * | % Name
# Get software list from computers
$Report=Get-InstalledSoftware -ComputerName $PCs |
    Where {$_.Name -like 'Microsoft 365*'} |
    Test-Ms365RequiresUpdate -Channel 'Monthly Enterprise Channel'
# Export your data however you want, ex: CSV
$Report | Export-CSV $env:temp\ms365-version-qa.csv -notype
```
