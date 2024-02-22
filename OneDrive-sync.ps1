
###---Set the script path to current location
$scriptPath = split-path -Parent $MyInvocation.MyCommand.Definition


###---Function that discoveres available drives
function Get-AvailableDriveLetter
{
    $allDriveLetters = 90..68
    $currentDriveLetters = Get-CimInstance -ClassName win32_LogicalDisk
    foreach ($Letter in $allDriveLetters)
    {
        $driveLetter = [String][Char]$Letter + ":"
        $listcheck = $currentDriveLetters.DeviceID -contains $driveLetter
        $testpath = Test-Path $driveLetter
        if ($listcheck -eq $false -and $testpath -eq $false)
        {
            return $driveLetter
        }
    } 
}

###---Create input box
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Move (H) to OneDrive'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter email handle of the target end-user:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $emailID = $textBox.Text
    $emailID | Out-Null
}


###---Convert email handle to full email address and then retreive HomeDir value
$WB = "[domain]"
$email = $emailID + $WB
$homedrive = Get-ADUser -Filter {UserPrincipalName -eq $email} -Properties HomeDirectory
if ($homedrive -ne $null)
{
	$source = $homedrive.HomeDirectory
}
else
{
	Write-Host "Error. Please enter valid email handle."
}


###---Retreive OneDrive URL based on email address and select a 'Personal Space' only
$userName = "[Username]"
$password = "[Password]"
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $(convertto-securestring $password -AsPlainText -Force)
Connect-SPOService -Url [Admin Sharpoint URL] -Credential $creds
$sODsite = (Get-SPOSite -IncludePersonalSite $true -Filter "owner -eq $email").url
if ($sODsite[1])
{
    foreach ($sODsubsite in $sODsite)
    {
        if ($sODsubsite.Contains("personal"))
        {
            $sODsite = $sODsubsite
        }
    }
}


###---Launch background process of IE11 with SPO URL which resolves the authentication/access issue
$ie = New-Object -ComObject "InternetExplorer.Application"
$ie.navigate("$sODsite")
Start-Sleep -Seconds 10
get-process iexplore | Stop-Process


###---Extend URL variable, get available drive letter, map OneDrive URL as network drive, set destination variable
$OneDriveURL = $sODsite +"/Documents/"
$aDriveLetter1 = Get-AvailableDriveLetter
$network = New-Object -ComObject WScript.Network
$network.MapNetworkDrive("$aDriveLetter1",$OneDriveURL)
$destination = New-PSDrive -Name "OneDrive" -Root $aDriveLetter1 -PSProvider filesystem


###---Send kick off email
.$scriptPath\Send-Email.ps1 -To "[EMAIL]" -Subject "Start H: drive migration for user '$email'" -Body "***This mailbox is not monitored. Do not reply.***"


###---Remove Read-Only attribute from all files
$allFiles = Get-ChildItem -Path $source -Recurse
foreach ($file in $allFiles)
{
    if ($file.isReadOnly){
        $file.isReadOnly = $false
    }
}


###---Remove Archive attribute from all files
$attribute = [io.fileattributes]::archive
foreach ($file in $allFiles)
{
	if ((Get-ItemProperty -Path $file.fullname).attributes -band $attribute)
	{
    	Set-ItemProperty -Path $file.fullname -Name attributes -Value ((Get-ItemProperty $file.fullname).attributes -BXOR $attribute) -ErrorAction SilentlyContinue
    }
}


###---Start Robocopy from Home drive to OneDrive
<#

Copy Options
/e copies subs including empty directories
/mt creates multithreaded copies with n threads (default is 8)
/copy: specifies which file properties to copy. the valid values for this option are (default option is DAT):
D-Data A-Attributes T-Time stamps S-NTFS access control list (ACL) O-Owner information U-Auditing information
/dcopy:	Specifies what to copy in directories. The valid values for this option are (default is DA):
D-Data A-Attributes T-Time stamps
/fft assumes FAT file times (should help when copying between file systems)
/compress requests network compression during file transfer, if applicable
/256 turns off support for paths longer than 256 characters (should help with SPO compatibility)
/tbd specifies that the system will wait for share names to be defined (retry error 67)

Exclusion Options
/xn excludes existing files newer than the copy in the source directory (Robocopy normally overwrites those)
/xo excludes existing files older than the copy in the source directory (Robocopy normally overwrites those)
/xc excludes changed files.
/xf [...] excludes files that match specified names (wildcard supported)
/xd [...] excludes directories that match specified names (wildcard supported)
/xa: excludes files for which any of the specified attributes are set. The valid values for this option are:
R-Read only A-Archive S-System H-Hidden C-Compressed N-Not content indexed E-Encrypted T-Temporary O-Offline

Retry Options
/r:<n> specifies number of retries on failed copies (default is 1 million)
/w:<n> specifies wait time between retries in seconds (default is 30 seconds)

Logging Options
/np	specifies that the progress of the copying operation (the number of files or directories copied so far) will not be displayed
/tee writes the status output to the console window, as well as to the log file
/v produces verbose output, and shows all skipped files
/log:<logfile> writes status output to the log file (overwrites existing log)
#>

.$scriptPath\Get-FailedFiles.ps1

$checkPath = "C:\Blair\Logs\OneDrive_Migration"
if (!(Test-Path $checkPath))
{
      New-Item -ItemType Directory -Force -Path $checkPath
}

#Robcopy and logging settings
$logPath = "[LogPath]"
$migrationLog = "$emailID"+"_"+(Get-Date -Format "MM-dd-yyyy_hh-mm-ss")+".log"
$failedFilesLog = "$emailID"+"_Failed_"+(Get-Date -Format "MM-dd-yyyy_hh-mm-ss")+".csv"
$skippedFilesLog = "$emailID"+"_Skipped_"+(Get-Date -Format "MM-dd-yyyy_hh-mm-ss")+".csv"
$copystatsLog = "$emailID" + "_Stats_" + (Get-Date -Format "MM-dd-yyyy_hh-mm-ss") + ".csv"
#Basic settings
$optionsA = @("/e", "/mt", "/r:3", "/w:5", "/v", "/log:$logPath\$migrationLog")
#Advanced/performance settings
$optionsB = @("/e", "/xn", "/mt:4", "/copy:DT", "/fft", "/compress", "/r:2", "/w:5", "/v", "/np", "/tee", "/log:$logPath\$migrationLog")
$exclusions = @("/xa:h", "/xf", '"*.one"', '"*.onetoc2"', '"*.exe"', '"*.bin"', '"*.msi"', '"*.pst"', "/xd", '$RECYCLE.BIN', "System Volume Information")
$cmdArgs1 = @("$source","$aDriveLetter1",$optionsB,$exclusions)

#Start files Robocopy
robocopy @cmdArgs1
Get-FailedFiles -RoboLog $logPath\$migrationLog -FailedFilesCSV $logPath\$failedFilesLog -CopyStatsCSV $logPath\$copystatsLog -SkippedFilesCSV $logPath\$skippedFilesLog


###---Get another available drive letter, map network share path for logs as network drive, copy logs from C: drive to S: drive
$aDriveLetter2 = Get-AvailableDriveLetter
$logShare = '[LogShareLocation]'
$Network = New-Object -ComObject WScript.Network
$Network.MapNetworkDrive("$aDriveLetter2",$logShare)
$optionsC = @("/e", "/compress", "/np", "/r:2", "/w:2")
$cmdArgs2 = @("$logPath","$aDriveLetter2",$optionsC)

#Start logs Robocopy
robocopy @cmdArgs2


###---Reset script environment and remove mapped drive
Set-Location $scriptPath -PassThru | Out-Null
Get-PSDrive OneDrive | Remove-PSDrive -force | Out-Null
NET USE $aDriveLetter1 /DELETE | Out-Null
NET USE $aDriveLetter2 /DELETE | Out-Null


###---Send completion email
.$scriptPath\Send-Email.ps1 -To "[EMAIL]" -Subject "End H: drive migration for user '$email'" -Body "***This mailbox is not monitored. Do not reply.***"