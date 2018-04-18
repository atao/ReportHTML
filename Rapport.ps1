# region Variables
$HTMLReport = "$PSScriptRoot\Report.html"
$ReportTitle = "Services"
# endregion Variables

# region Main
#Suppression du rapport HTML précédent
Remove-Item -Path $HTMLReport

# Collect Data
# $ResultSet = Import-Csv C:\Temp\Rapport.csv -Delimiter ";" | Sort-Object -Property Lieu, Utilisateur | Select-Object @{N="Nom PC";E={$_.Nom}}, Type, "Système d'exploitation", fabricant, lieu, utilisateur |
$ResultSet = Get-Service | Select-Object Name, DisplayName, Status, CanPauseAndContinue, CanShutdown, CanStop | Sort-Object Name |
ConvertTo-Html -CssUri "$PSScriptRoot\Rapport.css" -PreContent '<script src="sorttable.js"></script>' -Title $ReportTitle -Body "<h1>$ReportTitle</h1>`n<h5>Updated: on $(Get-Date)</h5>"
$ResultSet = $ResultSet -replace "<table>", '<table class="sortable">'
# Write Content to Report.
Add-Content $HTMLReport $ResultSet
# Call the results or open the file.
Invoke-Item $HTMLReport
# endregion Main