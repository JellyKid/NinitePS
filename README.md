# NinitePS

Powershell wrapper for [Ninite Pro](http://www.ninite.com/pro) It keeps track of currently installed software and makes pretty reports for email.

##Install

Just include NinitePS.ps1 in the same directory as ninitepro.exe

##Usage

NinitePS is a wrapper that will search your AD structure for enabled machines, test their connectivity, make specified software changes, store those changes in a CSV and create a job completion report. The following are script arguments to make the whole thing work.

./NinitePS.ps1 -Audit
...Scans all machines in your environment for Ninite supported software for out of date and installed software. It then stores it's findings in ComputerStats.csv in the same directory and creates an HTML report called Report.html

./NinitePS.ps1 -Update
...Scans and updates all machines in your environment for out of date software and updates with /disableshortcuts and /disableautoupdate. It then updates the ComputerStats.csv with any software successfully updated.

./NinitePS.ps1 -Install
...Installs specified software with /disableshortcuts and /disableautoupdate on all machines in environment. It then updates the ComputerStats.csv with any software successfully installed. Requires the -product argument.

./NinitePS.ps1 -Uninstall
...Uninstalls specified software on all machines in environment. It then updates the ComputerStats.csv with any software successfully uninstalled. Requires the -product argument.

./NinitePS.ps1 -Product
...Specifies which software to work with, otherwise it includes all supported software. Install and uninstall arguments must have products specified. For a list of product commandline arguments check out Ninite's site [here](https://ninite.com/applist/pro.html)





