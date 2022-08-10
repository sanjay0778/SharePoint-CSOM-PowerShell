#Set Variables
$SiteURL = "https://abc.sharepoint.com/sites/TestSite"
$ListName = "Test List"
$OldContentTypeName = "Test Content Type"
$NewContentTypeName = "Item"
 
#Connect to PNP Online
Connect-PnPOnline -Url $SiteURL -Interactive 
 
#Get the New Content Type from the List
$NewContentType = Get-PnPContentType -List $ListName | Where {$_.Name -eq $NewContentTypeName}
 
#Get List Items of Old content Type
$ListItems = Get-PnPListItem -List $ListName -Query "<Query><Where><Eq><FieldRef Name='ContentType'/><Value Type='Computed'>$OldContentTypeName</Value></Eq></Where></Query>"
Write-host "Total Number of Items with Old Content Type:"$ListItems.count
 
ForEach($Item in $ListItems)
{
    #Change the Content Type of the List Item
    Set-PnPListItem -List $ListName -Identity $Item -ContentType $NewContentType
}