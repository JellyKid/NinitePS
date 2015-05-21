

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
tr:nth-child(odd) {background: #CCC}
</style>
"@
$date = Get-Date
$pre = @"
<h1>
3rd Party Update Report 
</h1>
<h2>
Run on $date
</h2>
"@
$cnu = 4
$Post = @"
<h2>
$cnu computers need updates
</h2>
"@

$MyReport = Import-Csv .\ComputerList.csv | Sort-Object UpToDate

$MyReport | Select 'Name','UpToDate','Pingable','LastContact','Needed' | ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | Out-File UpdateReport.html
