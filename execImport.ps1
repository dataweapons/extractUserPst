

#Outlook COM Object how I connect to outlook  

 

[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null

$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]

$OlClass = "Microsoft.Office.Interop.Outlook.OlObjectClass" -as [type]

$OlSaveAs = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]

$OlBodyFormat = "Microsoft.Office.Interop.Outlook.OlBodyFormat" -as [type]

 

$Outlook = new-object -comobject outlook.application

$NameSpace = $Outlook.GetNameSpace("MAPI")

 

[ENUM]::GetNames($olFolders)

 

$NameSpace.Folders | Select Name

 

#Sets the number of days to check email to move to PST

 

$30days =(Get-Date).adddays(-3)

 

#RegEx pattern to verify date format in user description field.

 

$RegEx = '^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](20)\d\d$'

 

#User will not be able to see error messages

 

$ErrorActionPreference = "SilentlyContinue" 

 

# User email address and source folder 

# if you wanted to add multiple folders 

#$Source = $NameSpace.Folders["MyName.Domain.com"].Folders.Item("Inbox").Folders.Item("Subfolder1").Folders.Item("SubFolder2")

 

$Source = $NameSpace.Folders["name@email.com"].Folders.Item("Inbox")

 

#Messages that we are moving that are older then 30days 

 

$Messages = $Source.Items | Where{$(get-date $_.ReceivedTime) -Lt $30days}

 

# Destention PST file name and name of folder that the user see in Outlook 

 

$PST = $NameSpace.Folders["April 2014"].Folders.Item("Inbox")

 

foreach ($msg in $Messages) {

    "Moving $($msg.Subject) from $($msg.SenderName) sent on $($msg.SentOn) to the PST file."

   $msg.Move($PST) 

}

 

 

#Message that the user know how many emails have been copied to spicfic PST file 

 

Write-host Complete  -ForegroundColor  Red
