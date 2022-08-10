#Set Variables
$SiteURL = "https://abc.sharepoint.com/sites/TestSite/subsite1"

 
#Connect to PNP Online
Connect-PnPOnline -Url $SiteURL -Interactive 

$ListName = "TestList"
Get-PnPList -Identity $ListName | Get-PnPListItem -PageSize 100 -ScriptBlock { 
    Param($items) Invoke-PnPQuery } | ForEach-Object {$_.Recycle()
}

