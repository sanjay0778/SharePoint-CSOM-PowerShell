Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# Specify the path to the Excel file and the WorkSheet Name
$FilePath = "C:\MigrationExcelAndScript\Github\TestListExcel.xlsx"
$SheetName = "TestListExcel"
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
    $DateFormat = "dd-MMM-yy"
    $count = 0
    # Excel operations
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    $WorkBook = $objExcel.Workbooks.Open($FilePath)
    # Load the WorkSheet
    $WorkSheet = $WorkBook.sheets.item($SheetName)

    $intRowMax = ($WorkSheet.UsedRange.Rows).count
    Write-Host "Number of rows =" $intRowMax
    #$intRowCount = $intRowMax
    # Column numbers in excel
    $intNumber = 1
    $intTitle = 2
    $intPlanDate = 3
    # start index for forloop
    $startIndex = 2
    $intRow = $startIndex
    While ($intRow -le $intRowMax)
    {
    $g = 0
        for($intRow; $intRow -le $intRowMax -and $g -lt 100; $intRow++)
        {    
            $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $listItem = $list.AddItem($listItemInfo)
            $Number  = $WorkSheet.cells.item($intRow,$intNumber).value2
            Write-Host $Number -ForegroundColor Green
            if ($Number)
            {
                $listItem["Number"] = $Number  
            }
        
            $Title  = $WorkSheet.cells.item($intRow,$intTitle).value2        
            if ($Title)
            {
                $listItem["Title"] = $Title  
            }
        
            $PlanDate  = $WorkSheet.cells.item($intRow,$intPlanDate).text
            if ($PlanDate)
            {
                $listItem["PlanDate"] = [DateTime]::ParseExact($PlanDate,$DateFormat,[System.Globalization.DateTimeFormatInfo]::InvariantInfo,[System.Globalization.DateTimeStyles]::None)
            }
            else
            {
                Write-Host "PlanDate is blank for Title - " $Title -foregroundcolor red -backgroundcolor white
            }
            $listItem.Update()  
            Write-Host "Item Added with Title - " $Title -foregroundcolor black -backgroundcolor yellow
            #}
            $count = $count+1; 
            Write-Host "Number of items added successfully = " $count -ForegroundColor Green 
            $g = $g + 1
        }
        Write-Host "Please wait while we are commiting changes in SharePoint"
        $ctx.load($list)
        #commiting items in SharePoint list in batch of 100 items at a time      
        $ctx.executeQuery()         
        $startIndex = $startIndex + 100
    } 
    Write-Host "All List Items Added Successfully!" -ForegroundColor Green                    
    $WorkBook.close()
    $objExcel.quit() 
}  
catch{  
    #write-host "$($_.Exception.Message)" -foregroundcolor red  
    echo $_.Exception|format-list -force
    $WorkBook.close()
    $objExcel.quit()
}