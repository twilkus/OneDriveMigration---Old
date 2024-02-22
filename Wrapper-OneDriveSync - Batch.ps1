$username = "[UserName]"
$password = "[Password]"
$securePW = $password | ConvertTo-SecureString -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $userName,$securePw

$arg = '-executionpolicy bypass -file "[LocationofOneDrive-syncScript]"'

Start-Process -FilePath "powershell.exe" -ArgumentList $arg -Credential $creds