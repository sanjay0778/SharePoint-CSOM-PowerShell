Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# SharePoint Variables      
$siteURL = "https://abc.sharepoint.com/sites/test"  
$ListTitle= "TestList"
$userId = "sanjay.pathak@xyz.com"  
$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL) 
$ctx.RequestTimeOut = 5000*10000 
$ctx.credentials = $creds  
try{
    
    $lists = $ctx.web.Lists  
    $list = $lists.GetByTitle($ListTitle)
    
    $ListItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()) 
    $ctx.Load($ListItems)
    $ctx.ExecuteQuery() 
    write-host "Total Number of List Items found:"$ListItems.Count -ForegroundColor Green   
    #Delete all list items
    if ($ListItems.Count -gt 0)
    {
        $counter = $ListItems.Count
        While ($counter -gt 1 )
        {
        $g = 0
            #Loop through each item and delete in batch of 100 items
            For ($i = $counter-1; $i -ge 0 -and $g -lt 100; $i--)
            {
                $ListItems[$i].DeleteObject()
                $g = $g + 1
                Write-Host "Item left to be deleted = " $i -ForegroundColor Green
            }        
            $ctx.ExecuteQuery() 
            $counter = $counter - 100
        }
        Write-Host "All List Items deleted Successfully!" -ForegroundColor Green
    }    
           
}  
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
}  
