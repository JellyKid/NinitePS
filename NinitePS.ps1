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
		[Parameter(Mandatory=$true,ParameterSetName='install')]
		[Parameter(Mandatory=$true,ParameterSetName='uninstall')]
		[Parameter(Mandatory=$false,ParameterSetName='update')]
		[string[]]$product,
		[Parameter(Mandatory=$false,ParameterSetName='install')]
		[Parameter(Mandatory=$false,ParameterSetName='uninstall')]
		[Parameter(Mandatory=$false,ParameterSetName='update')]
        [Switch]$FullReports,
		[Parameter(Mandatory=$false,ParameterSetName='install')]
		[Parameter(Mandatory=$false,ParameterSetName='uninstall')]
		[Parameter(Mandatory=$false,ParameterSetName='update')]
		[Parameter(Mandatory=$false,ParameterSetName='audit')]
		[string]$Machine,
		[Parameter(Mandatory=$true,ParameterSetName='reportonly')]
		[Switch]$ReportOnly
)


If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    exit
}

if(($install -or $uninstall) -and !$product) {
	Write-Warning "You must specify a product when installing or uninstalling"
	exit
}


#Get all machines in AD and test their connectivity

Push-Location $PsScriptRoot

 
#region setting job options
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

if($audit -or $reportonly){
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

#endregion setting job options
 
#region helper functions
function include($string1, $string2) {
	$array1 = {$string1 -split ', '}.Invoke()
	$array2 = {$string2 -split ', '}.Invoke()
	foreach ($item in $array2) {
		if ($array1 -notcontains $item) {
			$array1.add($item)
		}
	}
	return ($array1 -join ', ')
}

function exclude($string1, $string2) {
	$array1 = {$string1 -split ', '}.Invoke()
	$array2 = {$string2 -split ', '}.Invoke()
	foreach ($item in $array2) {
		if ($array1 -contains $item) {
			$array1.remove($item) | out-null
		}
	}
	return ($array1 -join ', ')
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
		if($item.name -notcontains 'Computer' -and $item.name -notcontains 'Status') {
			switch -wildcard ($item.value) {
				"Update*"{$needed += $item.Name}
				"Partial*"{$needed += $item.Name}
				"OK*"{$installed += $item.Name}
				"Skipped -*"{$installed += $item.Name}
				"Skipped (no updates)"{}
				"Skipped (up to date)"{}
				"Skipped (not installed)"{}
				"Not installed"{}
				"Success*"{}
				default {
					$needed += $item.Name
					$errors += $item.Name
					$errors += " - "
					$errors += $item.Value
					$errors += ";`n"
				}
			}
		}
	}
	
	$returnarray = New-Object string[] 3
    $returnarray[0] = $needed -join ', '
    $returnarray[1] = $installed -join ', '
	$returnarray[2] = $errors
	return $returnarray
}

#endregion helper functions

function Call-Ninite {

	param
	(
		[Parameter(Mandatory=$True,position=0,HelpMessage="audit/install/upgrade")]
		$job,
		[Parameter(Mandatory=$True,position=1,HelpMessage="What machine?")]
		$computer
	)

	
	if(Test-Path ninitepro.exe) {
		$job += @("/remote",$computer)
		Write-Host $job
		$status = & .\ninitepro.exe $job | Write-Output
	} else {
		Write-Error "Error ninitepro.exe doesn't exist in this path"
		return $null
	}
	
	return $status
	
}

function BuildReport ($MyReport,$title) {

$footer = ""
$header = ""
$cnu = 0
$joberrors = 0

foreach ($comp in $MyReport){
		if ($comp.UpdatesNeeded -and ($comp.UpdatesNeeded -ne 'Never Checked')){
			$cnu++
		}
		if ($comp.Error){
			$joberrors ++
		}
		
	}

#Create footer and header
if ($audit -or $ReportOnly) {	
	$header = "$cnu computers need updates. $joberrors errors. Please see below."
} else {
	$header += "$jobname - <span class=`"niniteproduct`">$product</span>"
		
	$footer = "Job had $joberrors errors, please see above"
}



$date = Get-Date

$pre = @"

<style>
.ninitereport TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
.ninitereport TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
.ninitereport TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.ninitereport tr:nth-child(odd) {background: #CCC}
.niniteproduct {text-transform: capitalize;}
</style>

<div class="ninitereport">

<h1>
$title
</h1>

<h2>
$header
</h2>

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

if($ComputerStats -and $ReportOnly){
	$MyReport = $ComputerStats | Sort-Object 'Name' | Select 'Name','Connectivity','LastContact','UpdatesNeeded','UpToDate','Error' 
	BuildReport $MyReport $ReportTitle
	exit
}

#Create object to store working machines in
$CurrentList = @()

$CompObj = @{
	'Name'			= '';
	'Connectivity'	= '';
	'LastContact'	= 'None';
	'UpToDate'		= 'Unknown';
	'UpdatesNeeded'	= 'Never Checked';
	'Error'			= '';
    'JobStatus'     = '';
}



Import-Module ActiveDirectory
 

if($machine) {
	$ADList = Get-ADComputer -Filter {(cn -eq $Machine)}
} else {
	$ADList = Get-ADComputer -Filter '*'
}

 
#region Main Logic
foreach ($computer in $ADList) {
	 
	if ($computer.Enabled){	
		
		#Create a new computer object and see if we can ping it
		$NewCompObj = New-Object -TypeName PSObject -Property $CompObj
		$NewCompObj.Name = $computer.Name
		
		#Check if machine can be pinged and TermService up and running
		if(Test-Connection -ComputerName $NewCompObj.Name -Count 1 -Quiet) {
			if((Get-Service -ComputerName $NewCompObj.Name | Where-Object {$_.Name -eq 'RpcSs'}).Status -eq 'Running'){
				$NewCompObj.Connectivity = $true
			} else {
				$NewCompObj.Error = "RPC Services not running on remote machine"
			}
		} else {
			$NewCompObj.Error = "Cannot ping machine"
		}
						
		if ($NewCompObj.Connectivity) {
			
            $NewCompObj.LastContact = Get-Date
			
			#Execute Ninite with specified parameters and return raw results
            $results = Call-Ninite $job $NewCompObj.Name
			
			#Check if job was successful 
			if ($results) {
				if ($results[0] -contains 'OK') {
					$success = $results[0]
					$results = $results[1..($results.Count - 1)]
					
				} else {
					$success = $results[0]
					foreach ($line in $results) {write-warning $line}	
				} 
			} else {
				Write-Host 'No results?'
				$NewCompObj.Error = 'No Results from last job'
			}
            
          	#Parse and store the results in the 
			#newly created computer object
			if ($success -eq 'OK') {
                				
				
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
									
			} else {
				$NewCompObj.Error = $success
				$NewCompObj.JobStatus = "Failed"
			}
			
			
			
			#Write out raw ninite reports, mainly for debugging purposes			
			if($FullReports -and $results){
                    $date = Get-Date -UFormat "%m%d%Y-%H%M"
                    $compname = $NewCompObj.Name
                    if(!(Test-Path ".\FullReports")){New-Item -ItemType directory 'FullReports'}
                    ConvertFrom-Csv $results | Export-Csv ".\FullReports\$compname-$date.csv" -NoTypeInformation
            }
			
		}

		#This next if/else block compares current computer information to stored
		#and puts the information back into stored
		
		if ($ComputerStats -and $ComputerStats.Name.Contains($NewCompObj.Name)) {
		
			if (($NewCompObj.UpdatesNeeded -eq 'Never Checked') -and $NewCompObj.connectivity){$NewCompObj.UpdatesNeeded = ''}
		
			$i = $ComputerStats.Name.IndexOf($NewCompObj.Name)
			$ComputerStats[$i].Connectivity = $NewCompObj.Connectivity
			if($NewCompObj.LastContact -ne 'None'){$ComputerStats[$i].LastContact = $NewCompObj.LastContact}
			$ComputerStats[$i].Error = $NewCompObj.Error
			
			if ($uninstall) {
				$ComputerStats[$i].UpToDate = exclude $ComputerStats[$i].UpToDate $NewCompObj.UpToDate
				$ComputerStats[$i].UpdatesNeeded = exclude $ComputerStats[$i].UpdatesNeeded $NewCompObj.UpToDate
			}
			if ($install -or $update) {
				$ComputerStats[$i].UpToDate = include $ComputerStats[$i].UpToDate $NewCompObj.UpToDate
				$ComputerStats[$i].UpdatesNeeded = exclude $ComputerStats[$i].UpdatesNeeded $NewCompObj.UpToDate
			}
			if ($audit -and $NewCompObj.Connectivity) {
				$ComputerStats[$i].UpToDate = $NewCompObj.UpToDate
				$ComputerStats[$i].UpdatesNeeded = $NewCompObj.UpdatesNeeded
			}	
				
		} else {
			$ComputerStats += $NewCompObj | Select 'Name','Connectivity','LastContact','UpToDate','UpdatesNeeded','Error'
		}
		
		#Add content to make reports look better
		
		if (!$audit -and $NewCompObj.Connectivity) {
			if (!$NewCompObj.UpToDate -or ($NewCompObj.UpToDate -eq 'Unknown')){$NewCompObj.UpToDate = "None"}
			if (!$NewCompObj.UpdatesNeeded){$NewCompObj.UpdatesNeeded = "None"}
		} 
		if (!$audit -and !$NewCompObj.Connectivity) {
			if ($NewCompObj.UpdatesNeeded -eq 'Never Checked'){$NewCompObj.UpdatesNeeded = "Unknown"}
			if ($NewCompObj.UpToDate -eq 'Unknown'){$NewCompObj.UpToDate = "None"}
		}
		
		#Add the current computer object to the list for reports
		$CurrentList += $NewCompObj
		
	} else {
	#if computer is not enabled remove from computerstats.csv list
		if ($audit -and $ComputerStats) {
			if($ComputerStats.Name -Contains $computer.name){
				write-host "Removing disabled computer " $computer.name
				$ComputerStats = $ComputerStats | where {$_.Name -ne $computer.Name}
			}
		}
	}
	
	if (!$ComputerStats -or !$CurrentList) {
		Write-Error 'Missing Job information. Something went wrong.'
		break
	}
	#Export known and updated computer status list to CSV for future retrieval
	$ComputerStats | Sort-Object 'Name' | Select 'Name','Connectivity','LastContact','UpdatesNeeded','UpToDate','Error' | Export-CSV ComputerStats.csv

} 
#endregion Main Logic



#Add product to make reports look better

if (!$product) {
	$product = "All Software"
}

#Build reports based on job


if ($audit) {
	$MyReport = $ComputerStats | Sort-Object 'Name' | Select 'Name','Connectivity','LastContact','UpdatesNeeded','UpToDate','Error' 
}

if ($update) {
	$MyReport = $CurrentList | Sort-Object 'JobStatus' -Descending | Select 'Name','JobStatus','Connectivity',@{Name='Installed';Expression={$_.UpToDate}},'UpdatesNeeded','Error'
}

if ($install) {
	$MyReport = $CurrentList | Sort-Object 'JobStatus' -Descending | Select 'Name','JobStatus','Connectivity',@{Name='Installed';Expression={$_.UpToDate}},'Error'
}

if ($uninstall) {
	$MyReport = $CurrentList | Sort-Object 'JobStatus' -Descending | Select 'Name','JobStatus','Connectivity',@{Name='Uninstalled';Expression={$_.UpToDate}},'Error'
}

BuildReport $MyReport $ReportTitle
 
Pop-Location




