$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://<EXCHANGESERVER>/PowerShell/"; -Authentication Kerberos
Import-PSSession $s -commandname Get-Mailbox,New-MailboxExportRequest | Out-Null
$Export = Get-Mailbox
$Export | %{New-MailboxExportRequest -Malibox $_ -FilePath "\\server\pst\$($_.alias).pst"}
