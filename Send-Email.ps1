param([string]$to,
[string]$subject,
[string]$body
)

$smtpServer = "[RelayServer]"
$smtpFrom = "[EmailAddressFrom]"
$smtpTo = $to
$messageSubject = $subject
$messageBody = $body

$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($smtpFrom,$smtpTo,$messagesubject,$messagebody)