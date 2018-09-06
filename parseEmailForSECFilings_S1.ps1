
# this section is contructing the outlook object 
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")


#### S-1 ####
# here we navigate the outlook object to find items that are unread and assign them to the $unread variable. The values in parenthesis will vary.
$unread = $namespace.Folders.Item(1).Folders.Item(23).Folders.Item(1).items | where {$_.UnRead -eq "True"}
# make a new array, $urls
$urls = @()
#loop through our $unread messages and do some clunky parsing for the link text. Add that link text to the $urls array. Mark the message as read.
foreach($message in $unread){
   $a,$b = $message.Body -split "<"
   $url,$d = $b -split ">"
   $urls += $url 
   $message.UnRead = "False"
}

#loop through the $urls array and create a new browser window for each
foreach($url in $urls){
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.navigate2($url)
    while($ie.busy){
        Start-Sleep -Seconds 1
    }
# in each window, we traverse the DOM and find the appropriate link to click and stash it in a variable $link
$link = $ie.Document.getElementsByTagName("a") | where {$_.href -like "*/Archives/edgar/data*"} | select -Index 0
# here we click that link to open the document
$link.click()
}


