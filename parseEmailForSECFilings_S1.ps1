
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")


#### S-1 ####

$unread = $namespace.Folders.Item(1).Folders.Item(23).Folders.Item(1).items | where {$_.UnRead -eq "True"}
$urls = @()
foreach($message in $unread){
   $a,$b = $message.Body -split "<"
   $url,$d = $b -split ">"
   $urls += $url 
   $message.UnRead = "False"
}

foreach($url in $urls){
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.navigate2($url)
    while($ie.busy){
        Start-Sleep -Seconds 1
    }

$link = $ie.Document.getElementsByTagName("a") | where {$_.href -like "*/Archives/edgar/data*"} | select -Index 0
$link.click()
}


