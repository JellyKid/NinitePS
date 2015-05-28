param(
		
		[Parameter(ParameterSetName='install')]
		[Switch]$install,
		[Parameter(ParameterSetName='uninstall')]
		[Switch]$uninstall,
		[Parameter(ParameterSetName='audit')]
		[Switch]$audit,
		[Parameter(ParameterSetName='update')]
		[Switch]$update,
		[Parameter(ValueFromRemainingArguments=$true,Position=0)]
		[string[]]$product,
        [Switch]$FullReports,
		[string]$Machine
)


#Get all machines in AD and test their connectivity

Push-Location $PsScriptRoot

Import-Module ActiveDirectory
 

if($machine) {
	$ADList = Get-ADComputer -Filter {(cn -eq $Machine)}
} else {
	$ADList = Get-ADComputer -Filter '*'
}

 
#--- Start setting job options
if($install){
	$job = @(
		"/disableshortcuts",
		"/disableautoupdate"
	)
	$ReportTitle = 'Install Report'
	$jobname = 'Install'
}	

if($uninstall){
	$job = @(
		"/uninstall"		
	)
	$ReportTitle = 'Uninstall Report'
	$jobname = 'Uninstall'
}

if($update){
	$job = @(
		"/updateonly",
		"/disableshortcuts",
		"/disableautoupdate"
	)
	$ReportTitle = '3rd Party Update Report'
	$jobname = 'Update'
}

if($audit){
	$job = @(
		"/audit"
	)
	$ReportTitle = 'Software Audit Report'
}

$job += @("/silent",".")

if($product){
    $job += @("/select")
    $job += $product
}

#--- End setting job options
 
#--- Start helper functions
function combine($string1, $string2) {
	$old = {$string1 -split ','}.Invoke()
	$new = {$string2 -split ','}.Invoke()
	foreach ($item in $old) {
		if ($new -notcontains $item) {
			$new.add($item)
		}
	}
	return ($new -join ',')
}	

function create_hash ([array] $doublearray) {
	$keys = $doublearray[0].split(",")
	$values = $doublearray[1].split(",")
    $h = @{}
    if ($keys.Length -ne $values.Length) {
        Write-Error -Message "Array lengths do not match" `
                    -Category InvalidData `
                    -TargetObject $values
    } else {
        for ($i = 0; $i -lt $keys.Length; $i++) {
            $h[$keys[$i]] = $values[$i]
        }
    }
    return $h
}


function parse-results($resulthash) {
	$needed = @()
    $installed = @()
	$errors = ''
	
	foreach($item in $resulthash.GetEnumerator()) { 
		if ($item.Value.Contains('Update')) {
			$needed += $item.Name
		}
        if ($item.Value.Contains('OK')) {
            $installed += $item.Name
        }
		if ($item.Value.Contains('program running')) {
			$needed += $item.Name
			$errors += $item.Name
			$errors += " - "
			$errors += $item.Value
			$errors += ";`n"
		}
	}
	
	$returnarray = New-Object string[] 3
    $returnarray[0] = $needed -Join ','
    $returnarray[1] = $installed -Join ','
	$returnarray[2] = $errors
	return $returnarray
}

#--- End helper functions

function Call-Ninite {

	param
	(
		[Parameter(Mandatory=$True,position=0,HelpMessage="audit/install/upgrade")]
		$job,
		[Parameter(Mandatory=$True,position=1,HelpMessage="What machine?")]
		$computer
	)

	
	if(Test-Path Ninite.exe) {
		$job += @("/remote",$computer)
		Write-Host $job
		$status = & .\Ninite.exe $job | Write-Output
        
		$status = $status[1..($status.Count - 1)]  
		return $status
	} else {
		Write-Error "Error Ninite.exe doesn't exist in this path"
		return $null
	}
}

function BuildReport ($MyReport,$title) {

$footer = ""
$header = ""

#Get Computer Not Updated Count
if ($audit) {
	$cnu = 0
	
	foreach ($comp in $MyReport){
		if ($comp.UpdatesNeeded -and ($comp.UpdatesNeeded -ne 'Never Checked')){
			$cnu++
		}
	}
	
	$footer = "$cnu computers need updates"
} 

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
$title
</h1>

<h2>
Run on $date
</h2>

"@


$Post = @"

<h2>
$footer
</h2>

</div>

"@

$MyReport | ConvertTo-HTML -PreContent $Pre -PostContent $Post | Out-File Report.html
}

 

#Check and see if machines have already been audited and recorded
$ComputerStats = if (Test-Path ComputerStats.csv) {@(Import-Csv ComputerStats.csv)} else {,@()}

#Create object to store working machines in
$CurrentList = @()

$CompObj = @{
	'Name'			= '';
	'Connectivity'	= '';
	'LastContact'	= 'None';
	'UpToDate'		= 'Unknown';
	'UpdatesNeeded'	= 'Never Checked';
	'Error'			= '';
}

if (!$audit) {
	$CompObj.Add('JobStatus','')
}

if ($product) {
	$CompObj.Add('Product','')
}
 
#--- Start Main Logic
foreach ($computer in $ADList) {
	 
	if ($computer.Enabled){	
		
		#Create a new computer object and see if we can ping it
		$NewCompObj = New-Object -TypeName PSObject -Property $CompObj
		$NewCompObj.Name = $computer.Name
		$NewCompObj.Connectivity = Test-Connection -ComputerName $NewCompObj.Name -Count 1 -Quiet
						
		if ($NewCompObj.Connectivity) {
			
            $NewCompObj.LastContact = Get-Date
			
			#Execute Ninite with specified parameters and return raw results
            $results = Call-Ninite $job $NewCompObj.Name
            
          	#This next if/else block parses and stores the results in the 
			#newly created computer object
			if ($results) {	
                
				$ResultsHash = create_hash($results)
				
				#Change Status to Success so that results can be parsed properly
				if($ResultsHash.Status -eq 'OK') {$ResultsHash.Status = 'Success'}
				
				$ParsedResults = parse-results $ResultsHash
                
				$NewCompObj.UpdatesNeeded = $ParsedResults[0]
				$NewCompObj.UpToDate = $ParsedResults[1]
				$NewCompObj.Error = $ParsedResults[2]
				
				if (!$audit) {
					$NewCompObj.JobStatus = $ResultsHash.Status
				}

				if ($product) {
					$NewCompObj.Product = $NewCompObj.UpToDate
				}
									
			} else {
				Write-Host 'No results?'
				$NewCompObj.Error = 'No Results from last job'
			}
			
			#Write out raw ninite reports, mainly for debugging purposes			
			if($FullReports -and $results){
                    $date = Get-Date -UFormat "%m%d%Y-%H%M"
                    $compname = $NewCompObj.Name
                    if(!(Test-Path ".\FullReports")){New-Item -ItemType directory 'FullReports'}
                    ConvertFrom-Csv $results | Export-Csv ".\FullReports\$compname-$date.csv" -NoTypeInformation
            }
			
		}
		
		#Add the current computer object to the list for reports
		$CurrentList += $NewCompObj
		
		#This next if/else block compares current computer information to stored
		#and puts the information back into stored
		
		if ($ComputerStats -and $ComputerStats.Name.Contains($NewCompObj.Name)) {
			
				$i = $ComputerStats.Name.IndexOf($NewCompObj.Name)
				$ComputerStats[$i].Connectivity = $NewCompObj.Connectivity
				$ComputerStats[$i].LastContact = $NewCompObj.LastContact
				$ComputerStats[$i].Error = $NewCompObj.Error
				$ComputerStats[$i].UpToDate = combine $ComputerStats[$i].UpToDate $NewCompObj.UpToDate
				$ComputerStats[$i].UpdatesNeeded = combine $ComputerStats[$i].UpdatesNeeded $NewCompObj.UpdatesNeeded
				
		} else {
			$ComputerStats += $NewCompObj | Select 'Name','Connectivity','LastContact','UpToDate','UpdatesNeeded','Error'
		}	
	}
}

#--- End Main Logic

if (!$ComputerStats -or !$CurrentList) {
	Write-Error 'Missing Job information. Something went wrong.'
	break
}
#Export known and updated computer status list to CSV for future retrieval
$ComputerStats | Sort-Object 'Name' | Select 'Name','Connectivity','LastContact','UpdatesNeeded','UpToDate','Error' | Export-CSV ComputerStats.csv

#Build reports based on job

if ($audit) {
	$MyReport = $ComputerStats | Sort-Object 'Name' | Select 'Name','Connectivity','LastContact','UpdatesNeeded','UpToDate','Error' 
}

if ($update) {
	$MyReport = $CurrentList | Sort-Object 'JobStatus' -Descending | Select 'Name','JobStatus','Connectivity','UpToDate','UpdatesNeeded','Error'
}

if ($install -or $uninstall) {
	$MyReport = $CurrentList | Sort-Object 'JobStatus' -Descending | Select 'Name','JobStatus','Connectivity','Product','Error'
}

BuildReport $MyReport $ReportTitle
 
Pop-Location




