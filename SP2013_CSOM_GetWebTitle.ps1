#SharePoint 2013 CSOM - PowerShell script to get web title
#clear
#Download SharePoint 2013 CSOM from http://www.microsoft.com/en-us/download/details.aspx?id=35585
$showResultsOnly = $false;
$rowLimit = 10;
$siteUrl = "https://abc.xyz.net/"

#$isapi15 = "C:\Users\ssspp\Documents\csom15APIs"
$isapi15 = "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI"

Import-Module "$isapi15\Microsoft.SharePoint.Client.dll"
Import-Module "$isapi15\Microsoft.SharePoint.Client.Runtime.dll"

$execTime = $(get-date -Format "yyyyMMMdd-hhmm-sstt")
write-host "Execution Time: $execTime, $PSScriptRoot"

function main()
{
    $username = [System.Environment]::UserName
    write-host "Using default credentials for: $username" -f Yellow
    $cred = [System.Net.CredentialCache]::DefaultNetworkCredentials;
    $cacheCred = $cred
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $ctx.Credentials = $cacheCred;

    if ($ctx)
    {
        Write-Host "Opening Web..."
        $web = $ctx.Site.RootWeb
        $ctx.Load($web);
        $ctx.ExecuteQuery();
        $title = $web.Title;
        write-host "Opened Web: $($web.Title) from $siteUrl" -f Cyan;

        if ([String]::IsNullOrEmpty($title))
        {
            write-host "Unable to open web: check your site url or user permissions" -f Red
            return;
        }
    }
}

main
#pause