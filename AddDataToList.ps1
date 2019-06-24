Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

  
$siteURL = "https://abc.sharepoint.com/sites/test"  
$ListTitle= "TestList"
$userId = "sanjay.pathak@xyz.com"  
$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds  
try{ 

    # Import the .csv file, and specify manually the headers, without column name in the file 
    $contents = Import-CSV "C:\PROJECTS\ExcelAndScript\testList.csv" -header("Number", "Title","PlanDate")  
    
    $lists = $ctx.web.Lists  
    $list = $lists.GetByTitle($ListTitle)
    
    # Iterate for each list column
    $count = 0
    foreach ($row in $contents )
    {
    #if ($count -eq 10){
     #  break }
      
      #else{
        $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $listItem = $list.AddItem($listItemInfo)
        $listItem["ProjectNumber"] = $row.Number  
        $listItem["ProjectTitle"] = $row.Title          
        $DateFormat = "dd-MMM-yy"
        if ($row.PlanDate)
        {
            $listItem["PlanDate"] = [DateTime]::ParseExact($row.PlanDate,$DateFormat,[System.Globalization.DateTimeFormatInfo]::InvariantInfo,[System.Globalization.DateTimeStyles]::None)
        }
        else
        {
            Write-Host "PlanDate is blank for Title - " $row.Title -foregroundcolor red -backgroundcolor white
        }
        $listItem.Update()  
        Write-Host "Item Added with Title - " $row.Title -foregroundcolor black -backgroundcolor yellow
        #}
        $count = $count+1; 
        Write-Host "Number of items added successfully = " $count -ForegroundColor Green       
        
    }     
    
    $ctx.load($list)      
    $ctx.executeQuery()  
    Write-Host "All List Items Added Successfully!" -ForegroundColor Green                    
}  
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
}  
