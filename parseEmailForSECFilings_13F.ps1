
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")


#### 13F #### 

$unread = $namespace.Folders.Item(1).Folders.Item(23).Folders.Item(4).items | where {$_.UnRead -eq "True"}
$urls = @()
foreach($message in $unread){
    $a,$b = $message.body -split "<"
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

$link = $ie.Document.getElementsByTagName("a") | where {$_.href -like "*xslForm13F_X01*"} | select -Index 1 
$link.click()
}


############################ TEST SECTION ####################################
<#
#retain this line to see what number the folder to target is
#$namespace.Folders.Item(1).Folders | select FolderPath
#$namespace.Folders.Item(1).Folders.item(2).Items | select SentOn, SenderEmailAddress, Subject -Last 20 | Sort-Object SentOn


#>