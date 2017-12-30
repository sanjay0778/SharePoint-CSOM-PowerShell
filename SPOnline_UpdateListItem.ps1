#SharePoint Online : CSOM : PowerShell : Update List Item
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$siteURL = ""
$userId = ""
$pwd = Read-Host -Prompt "Enter password" -AsSecureString
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
$ctx.credentials = $creds
try{
    $lists = $ctx.web.Lists
    $list = $lists.GetByTitle("TestList")
    $listItem = $list.GetItemById(1)
    $listItem["Title"] = "aa"
    $listItem.Update()
    $ctx.load($listItem)   
    $ctx.executeQuery()
}
catch{
    write-host "$($_.Exception.Message)" -foregroundcolor red
} 