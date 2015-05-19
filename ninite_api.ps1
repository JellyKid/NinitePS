function parse-status($status) {
	$needed = @()
	foreach($item in $status) {
		$item
		if ($item.Contains('Update') {
			$needed += $item
		}		
	}
	return $needed
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
		$status = .\Ninite.exe /remote $computer /$job /silent . | Write-Output
		$status = $status[1..($status.Count - 1)]
		<# $status = ConvertFrom-Csv -InputObject $status #>
		return $status
	} else {
		echo "Error Ninite doesn't exist in this path"
		return $null
	}
}


$exists = Test-Path ComputerList.csv
if ($exists) {
	$CompList = Import-Csv ComputerList.csv
} else {
	$CompList = @()
}

$CompObj = @{
	'Name'			= $null;
	'Pingable'		= $null;
	'UpDated'		= $null;
	'Needed'		= $null;
	'LastContact'	= $null;
}
$computer = New-Object -TypeName PSObject -Property $CompObj



#Get all machines in AD and test their connectivity

Import-Module ActiveDirectory
#$ADList = Get-ADComputer -Filter '*' 
$ADList = Get-ADComputer -Filter {(cn -eq "JS-MSI")}


foreach ($computer in $ADList) {
	if ($computer.Enabled){					#PING MACHINES TO TEST CONNECTIVITY
		$computer.Pingable = Test-Connection -ComputerName $computer.Name -Count 1 -Quiet
		
		
		if ($computer.Pingable) {			#AUDIT MACHINE WITH NINITE
			$results = Call-Ninite 'audit' $computer.Name
		}
		
		if ($results) {						#PARSE RESULTS
			$computer.UpDated = $results.Status
			$computer.Needed  = parse-status($results)
			$computer.LastContact = Get-Date
		}
		
		<# if ($testcon) {
			$computer.Pingable = $true
		} else {
			$computer.Pingable = $false
		} #>
		
		if ($CompList.Name.Contains($computer.Name)) {
			$CompList[$CompList.Name.IndexOf($computer.Name)].Pingable = $computer.Pingable
		} else {
			$CompList += $computer
		}
		
		
	}
}

$CompList | Sort-Object Name | Export-CSV ComputerList.csv -NoTypeInformation
