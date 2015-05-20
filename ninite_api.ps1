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
		[string[]]$product
)
	

if($install){
	$job = @(
		"/disableshortcuts",
		"/disableautoupdate"
	)
}	

if($uninstall){
	$job = @(
		"/uninstall"		
	)
}

if($update){
	$job = @(
		"/updateonly",
		"/disableshortcuts",
		"/disableautoupdate"
	)
}

if($audit){
	$job = @(
		"/audit"
	)
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

function parse-results($resulthash) {
	$needed = @()
	foreach($item in $resulthash.GetEnumerator()) {
		if ($item.Value.Contains('Update')) {
			$needed += $item.Name
		}		
	}
	$returnstring = $needed -Join ','
	return $returnstring 
}

function Call-Ninite {

	param
	(
		[Parameter(Mandatory=$True,position=0,HelpMessage="audit/install/upgrade")]
		$job,
		[Parameter(Mandatory=$True,position=1,HelpMessage="What machine?")]
		$computer
	)

	$exists = Test-Path Ninite.exe
	if($exists) {
		Write-Host $job $computer
		$status = & .\Ninite.exe $job | Write-Output
		$status = $status[1..($status.Count - 1)]
		return $status
	} else {
		echo "Error Ninite doesn't exist in this path"
		return $null
	}
}


$exists = Test-Path ComputerList.csv
if ($exists) {
	$CompList = @(Import-Csv ComputerList.csv)
} else {
	$CompList = @()
}

$CompObj = @{
	'Name'			= $null;
	'Pingable'		= $null;
	'UpToDate'		= $null;
	'Needed'		= $null;
	'LastContact'	= $null;
}




#Get all machines in AD and test their connectivity

Import-Module ActiveDirectory
#$ADList = Get-ADComputer -Filter '*' 
$ADList = Get-ADComputer -Filter {(cn -eq "JS-MSI")}


foreach ($computer in $ADList) {
	if ($computer.Enabled){					#PING MACHINES TO TEST CONNECTIVITY
		$NewCompObj = New-Object -TypeName PSObject -Property $CompObj
		$NewCompObj.Name = $computer.Name
		$NewCompObj.Pingable = Test-Connection -ComputerName $NewCompObj.Name -Count 1 -Quiet
		
		
		if ($NewCompObj.Pingable) {			#AUDIT MACHINE WITH NINITE
			
			$results = Call-Ninite $job $NewCompObj.Name
		
		
			if ($results) {						#PARSE RESULTS
				$reshash = create_hash($results)
				$NewCompObj.UpToDate = $reshash.Status
				if ($NewCompObj.UpToDate -ne 'OK'){
					$NewCompObj.Needed  = parse-results($reshash)
				}
				$NewCompObj.LastContact = Get-Date
			}
					
			
		}
		
		if ($CompList){
				if ($CompList.Name.Contains($NewCompObj.Name)) {
					$index = $CompList.Name.IndexOf($NewCompObj.Name)
					if ($NewCompObj.Pingable) {
						$CompList[$index].Pingable = $NewCompObj.Pingable
						$CompList[$index].UpToDate = $NewCompObj.UpToDate
						$CompList[$index].Needed = $NewCompObj.Needed
						$CompList[$index].LastContact = $NewCompObj.LastContact	
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
	$CompList | Sort-Object Name | Export-CSV ComputerList.csv -NoTypeInformation
}
