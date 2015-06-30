# NinitePS

is a Powershell wrapper for [Ninite Pro](http://www.ninite.com/pro) that will search your AD structure for enabled machines, test their connectivity, make specified software changes, store those changes in a CSV and create a pretty HTML report. 

##Purpose
I really wanted to keep tabs on machines that may have been offline or unreachable since Ninite was last run. This tool accomplishes my objective by keeping a running record of all of the machines it has contacted, last date of contact and current "up to date" status. I also wanted better automation for my active directory domain and easier to read reports.

##Install

Just include NinitePS.ps1 in the same directory as ninitepro.exe

##Switches

####-Audit
Scans all machines in your environment for Ninite supported software for out of date and installed software. It then stores it's findings in ComputerStats.csv in the same directory and creates an HTML report called Report.html

####-Update
Scans and updates all machines in your environment for out of date software and updates with /disableshortcuts and /disableautoupdate. It then updates the ComputerStats.csv with any software successfully updated.

####-Install
Installs specified software with /disableshortcuts and /disableautoupdate on all machines in environment. It then updates the ComputerStats.csv with any software successfully installed. Requires the -product argument.

####-Uninstall
Uninstalls specified software on all machines in environment. It then updates the ComputerStats.csv with any software successfully uninstalled. Requires the -product argument.

####-Product
Specifies which software to work with, otherwise it includes all supported software. Install and uninstall arguments must have products specified. For a list of product commandline arguments check out Ninite's site [here](https://ninite.com/applist/pro.html)

####-Machine
Specify which machine to work with, otherwise it includes all machines found in AD. Right now this argument only takes one machine.

####-FullReports
Creates a CSV under the reports directory of Ninite output. Useful for debugging.

####-ReportOnly
Creates an HTML report based on content already in CSV. 

##Examples
Audit all machines in the domain

*.\NinitePS.ps1 -audit*


Install Greenshot on all machines in the domain

*.\NinitePS.ps1 -install -product greenshot*


Uninstall Adobe Reader on PC1

*.\NinitePS.ps1 -uninstall -product reader -machine PC1*


Update Oracle Java and Adobe Flash on PC2

*.\NinitePS.ps1 -update -product java flash -machine PC2*


Audit PC3 and update info in ComputerStats.csv

*.\NinitePS.ps1 -audit -machine PC3*

##Personal Usage

In my environment I have 2 scheduled tasks. The first task is running 4 times a day and is using the update switch. 

>"powershell -command c:\locationtoscript\NinitePS.ps1 -update"

This searches for all active computers in my domain, tests connectivity, updates any software out of date(audits & updates if machine isn't in csv already), and stores it in the CSV. The second task I run at the end of the day is the -ReportOnly switch followed by an emailreports.ps1 script

>"powershell -command c:\locationtoscript\NinitePS.ps1 -ReportOnly"
>"powershell -command c:\locationtoscript\emailreports.ps1"

####emailreports.ps1
```powershell
$mailprefs = @{
            'To'           = 'email1@domain.com','email2@domain.com';
            'Subject'      = 'Ninite Reports';
            'Body'         = 'No Report Found!';
            'From'         = 'ninite@domain.com';
            'SmtpServer'   = 'smtp.server.domain.com'
}

Push-Location $PsScriptRoot

if(Test-Path .\Report.html) {
    $body = Get-Content .\Report.html | Out-String
    $mailprefs.Set_Item('Body',$body)
    $mailprefs.Add('BodyAsHtml',$true)
}

Send-MailMessage @mailprefs

Pop-Location
``` 


##Output
Everytime you run NinitePS an html report.html is create with job status and ComputerStats.csv is update/created with the latest status. 



A great way to use this tool is to have it create HTML reports and mail those reports through a simple script as follows


