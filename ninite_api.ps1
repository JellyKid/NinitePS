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
        [Switch]$FullReports
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

	
	if(Test-Path Ninite.exe) {
		Write-Host $job $computer
        $job += @("/remote",$computer)
		$status = & .\Ninite.exe $job | Write-Output
        
		$status = $status[1..($status.Count - 1)]  
		return $status
	} else {
		Write-Error "Error Ninite.exe doesn't exist in this path"
		return $null
	}
}

function BuildReport ($MyReport) {

#Get Computer Not Updated Count
$cnu = 0
$MyReport.Needed.ForEach({
    if($_ -ne '' -and $_ -ne $null){
    $cnu++
    }
})

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

$Post = @"
<h2>
$cnu computers need updates
</h2>
"@

$MyReport | Sort-Object UpToDate | Select 'Name','UpToDate','Pingable','LastContact','Needed' | ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | Out-File UpdateReport.html
}



if (Test-Path ComputerList.csv) {
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
	if ($computer.Enabled){					
		$NewCompObj = New-Object -TypeName PSObject -Property $CompObj
		$NewCompObj.Name = $computer.Name
		$NewCompObj.Pingable = Test-Connection -ComputerName $NewCompObj.Name -Count 1 -Quiet
		
		
		if ($NewCompObj.Pingable) {			
			
            
            if (!$audit) {
                Call-Ninite $job $NewCompObj.Name
            }

            #Ninite needs to be called /w the audit command no matter what to get the results

            $job = @("/audit","/silent",".")
            
            
			$results = Call-Ninite $job $NewCompObj.Name
		
		
			if ($results) {						#PARSE RESULTS
                
				$reshash = create_hash($results)

                if($FullReports){
                    $date = Get-Date -UFormat "%m%d%Y-%H%M"
                    $compname = $NewCompObj.Name
                    if(!(Test-Path ".\FullReports")){New-Item -ItemType directory 'FullReports'}

                    #$reshash | Export-Csv -Path ".\FullReports\$compname-$date.csv"
                }
                
				$NewCompObj.UpToDate = $reshash.Status
				if ($NewCompObj.UpToDate -ne 'OK'){
					$NewCompObj.Needed  = parse-results($reshash)
                    $cnu ++
				}
				$NewCompObj.LastContact = Get-Date
			} else {Write-Host 'No results?'}
					
			
		}
        else {
            $NewCompObj.UpToDate = 'UnKnown'
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
    BuildReport($CompList)
}
