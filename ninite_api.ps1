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

Import-Module ActiveDirectory
 

if($machine) {
	$ADList = Get-ADComputer -Filter {(cn -eq $Machine)}
} else {
	$ADList = Get-ADComputer -Filter '*'
}

 

if($install){
	$job = @(
		"/disableshortcuts",
		"/disableautoupdate"
	)
	$ReportTitle = 'Install Report'
}	

if($uninstall){
	$job = @(
		"/uninstall"		
	)
	$ReportTitle = 'Uninstall Report'
}

if($update){
	$job = @(
		"/updateonly",
		"/disableshortcuts",
		"/disableautoupdate"
	)
	$ReportTitle = '3rd Party Update Report'
	$import = $true
}

if($audit){
	$job = @(
		"/audit"
	)
	$ReportTitle = 'Weekly Software Audit Report'
	$import = $true
}


$job += @("/silent",".")

if($product){
    $job += @("/select")
    $job += $product
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


function parse-results($resulthash,$current) {
	$needed = @()
    $installed = @()
	$errors = ''
	
	#Parse Results and add to needed pile or installed
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
	
	if ($current -eq 'Unknown') {$current = ''}
	#Compare machines old need list to newly parsed and add or remove as needed
	if ($current) {
		foreach($item in $current.split(',')) {
			if(!$needed.Contains($item)){
				$needed += $item
			}
		}
	} 
	
	
	$returnarray = New-Object string[] 3
    $returnarray[0] = $needed -Join ','
    $returnarray[1] = $installed -Join ','
	$returnarray[2] = $errors
	return $returnarray
}

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


#Get Computer Not Updated Count
if ($import) {
	$cnu = 0

	if ($MyReport.Needed.GetType().Name -eq 'String') {
		if($MyReport.Needed -ne 'Unknown' -and $MyReport){
			$cnu++
		} 
	} else {
		$MyReport.Needed.ForEach({
			if($_ -ne 'Unknown' -and $_){
			$cnu++
			}
		})
	}
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

if ($import) {
$Post = @"
<h2>
$cnu computers need updates
</h2>
</div>
"@
} else {
if (!$NewCompObj.Errors){
$Post = @"
<h2>
Job Completed Successfully!
</h2>
</div>
"@
} else {
$Post = @"
<h2>
Job completed with $joberror errors, see above.
</h2>
</div>
"@
}
}


$MyReport | ConvertTo-HTML -PreContent $Pre -PostContent $Post | Out-File Report.html
}

 

if ($import -and (Test-Path ComputerList.csv)) {
	$CompList = @(Import-Csv ComputerList.csv)
} else {
	$CompList = @()
}

$CompObj = @{
	'Name'			= '';
	'Pingable'		= '';	
}
 
if ($audit) {
	$CompObj.Add('Installed','')
}

if ($update) {
	$CompObj.Add('Error','')
}
if ($import) {
	$CompObj.Add('UpToDate','Unknown')
	$CompObj.Add('Needed','Unknown')
	$CompObj.Add('LastContact','')
}

if (!$import) {
	$CompObj.Add('JobStatus','')
	$CompObj.Add('Success','')
	$CompObj.Add('Errors','')
}
 
#START LOGIC!!!!
foreach ($computer in $ADList) {
	 
	if ($computer.Enabled){	
		

		$NewCompObj = New-Object -TypeName PSObject -Property $CompObj
		$NewCompObj.Name = $computer.Name
		$NewCompObj.Pingable = Test-Connection -ComputerName $NewCompObj.Name -Count 1 -Quiet
						
		if ($NewCompObj.Pingable) {			
			
            
            $results = Call-Ninite $job $NewCompObj.Name
            
            

          			
			if ($results) {						#PARSE RESULTS
                
				$reshash = create_hash($results)
				
								
                if($FullReports){
                    $date = Get-Date -UFormat "%m%d%Y-%H%M"
                    $compname = $NewCompObj.Name
                    if(!(Test-Path ".\FullReports")){New-Item -ItemType directory 'FullReports'}
                    ConvertFrom-Csv $results | Export-Csv ".\FullReports\$compname-$date.csv" -NoTypeInformation
                    
                }
				
				if($reshash.Status -eq 'OK') {$reshash.Status = 'Success'} #Change Status to Success so that results can be parsed properly
                
				if ($import) { #If audit or update
					 
					$ParsedResults = parse-results $reshash $NewCompObj.Needed
					
					$NewCompObj.Needed = $ParsedResults[0]
					if ($NewCompObj.Needed -eq 'Unknown') {$NewCompObj.Needed = ''}
					
					if ($audit) {
						$NewCompObj.Installed = $ParsedResults[1]						
					} else {
						$NewCompObj.Error = $ParsedResults[2]
					}

					if($NewCompObj.Needed -eq '') {
						$NewCompObj.UpToDate = 'Yes'
					} elseif ($NewCompObj.Needed -ne 'Unknown'){
						$NewCompObj.UpToDate = 'No'
					}
					
					$NewCompObj.LastContact = Get-Date
					
				} else { #if Install or uninstall

					$NewCompObj.JobStatus = $reshash.Status
					
					foreach($item in $reshash.GetEnumerator()) { 
						if ($item.Value.Contains('OK')) {
							$NewCompObj.Success += $item.Name
							$NewCompObj.Success += "; "
						} else {
							if (!$item.Name.Equals('Computer') -and !$item.Name.Equals('Status')) {
								$NewCompObj.Errors += $item.Name
								$NewCompObj.Errors += " - "
								$NewCompObj.Errors += $item.Value
								$NewCompObj.Errors += ";`n"
							}
						}
					}
					 
				}
					
			} else {Write-Host 'No results?'}
					
			
		}
        else {
			if ($import) {$NewCompObj.UpToDate = 'Unknown'}
        }
		
		if ($CompList -and $import){
				if ($CompList.Name.Contains($NewCompObj.Name)) {
					$index = $CompList.Name.IndexOf($NewCompObj.Name)
					if ($NewCompObj.Pingable) {
						$CompList[$index].Pingable = $NewCompObj.Pingable
						$CompList[$index].UpToDate = $NewCompObj.UpToDate
						$CompList[$index].Needed = $NewCompObj.Needed
						$CompList[$index].LastContact = $NewCompObj.LastContact
						if ($audit) {
							$CompList[$index] | Add-Member -NotePropertyName 'Installed' -NotePropertyValue $NewCompObj.Installed -Force
						} else {
							$CompList[$index] | Add-Member -NotePropertyName 'Error' -NotePropertyValue $NewCompObj.Error -Force
						}
					} else {
						$CompList[$index].Pingable = $NewCompObj.Pingable
					}
				} else {
					$CompList += $NewCompObj
				}
			} else {
					$CompList += $NewCompObj
			}
		
	}
}
 if ($CompList) {
	if (!$import) {
		$CompList = $CompList | Sort-Object UpToDate | Select 'Name','JobStatus','Pingable','Success','Errors'
		BuildReport $CompList $ReportTitle
	} 
	
	if ($audit)
	{
		$CompList = $CompList | Sort-Object UpToDate | Select 'Name','UpToDate','Pingable','LastContact','Needed','Installed'
		BuildReport $CompList $ReportTitle
		$CompList | Select 'Name','UpToDate','Pingable','LastContact','Needed' | Export-CSV ComputerList.csv -NoTypeInformation
	}
	
	if ($update)
	{
		$CompList = $CompList | Sort-Object UpToDate | Select 'Name','UpToDate','Pingable','LastContact','Needed','Error'
		BuildReport $CompList $ReportTitle		
	}
	
}




