cls

#access the mailbox
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# counter to control how much we return
# iterate through items in inbox, get the pro rata emails, split each email into substrings wherein the second substring has the info we want
# split again to clean up the end of the email, print the desired substring to console, increment and loop
$n = 0
foreach($item in $inbox.Items){
   if($n -lt 85){
    if($item.subject -like "*Pro Rata*"){
        write-host $item.ReceivedTime -ForegroundColor Yellow
        $a,$b = $item.Body -split "Personnel"
        $c,$d = $b -split "mailto"
        $c
        write-host
        $n++
    }
   }
}
