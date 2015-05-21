

#Get Computer Not Updated Count
$cnu = 0




$date = Get-Date

$pre = @"

<style>
.ninitereport TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
.ninitereport TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
.ninitereport TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.ninitereport tr:nth-child(odd) {background: #CCC}
</style>
<div class="ninitereport">


<h1>
3rd Party Update Report 
</h1>
<h2>
Run on $date
</h2>
"@

$Post = @"
<h2>
$cnu computers need updates
</h2>
</div>
"@

$MyReport | Sort-Object UpToDate | Select 'Name','UpToDate','Pingable','LastContact','Needed' | ConvertTo-HTML -PreContent $Pre -PostContent $Post | Out-File UpdateReport.html
