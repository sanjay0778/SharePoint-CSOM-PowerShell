<#
/* ************************************************************************************************************ /
 Objective  : Creates CustomList with data 
/**************************************************************************************************************/
#>

clear
#Download SharePoint 2013 CSOM from http://www.microsoft.com/en-us/download/details.aspx?id=35585
Import-Module "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Import-Module "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$execTime = $(get-date -Format 'yyyyMMMdd-hhmm-sstt')
write-host "Execution Time: $execTime, $PSScriptRoot"

$Global:ctx = $null;
$Global:web = $null;

$Global:ListName = "ListName"
$Global:ListFields = "ListFields"
$Global:SiteColumnName = "SiteColumnName"
$Global:SiteColumnTitle = "SiteColumnTitle"
$Global:FilterLookupValues = "FilterLookupValues"
$Global:filterJoin = "@@"
$Global:delmiterTab = "`t"
$Global:delmiter = ","

function main()
{
    $siteUrl = "https://xyz/"

    $username = [System.Environment]::UserName
    write-host "Using default credentials for: $username" -f Yellow
    $cred = [System.Net.CredentialCache]::DefaultNetworkCredentials;
    $cacheCred = $cred
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $ctx.Credentials = $cacheCred;

    if ($ctx)
    {
        $web = $ctx.Site.RootWeb;

        $Global:ctx = $ctx;
        $Global:web = $web;

        $ctx.Load($web);
        $ctx.ExecuteQuery();
        $title = $web.Title;

        write-host "Opened Web: $($web.Title) from $siteUrl" -f Cyan;

        if ([String]::IsNullOrEmpty($title))
        {
            write-host "Unable to open web: check your site url or user permissions" -f Red
            return;
        }
        

        $file_path = "$psscriptroot\CustomList.csv";
        write-host "Loading CSV file for bulk update:" -fc green
        write-host $file_path

        if (-NOT (Test-Path $file_path))
        {
            write-host "Unable to find file" -fc Red 
            $proceed = $false;
            return;
        }
            
        $metadatas = Import-Csv $file_path -Delimiter $Global:delmiter -Encoding UTF8

        foreach($meta in $metadatas)
        {

            #get guid of current list
            $guid = "?-?-?-?-?"

            write-host "Working on List: $($meta.$Global:ListName)" -f Yellow
            if ($($meta.$Global:ListName) -ne "")
            {
                write-host "List Fields: $($meta.$Global:ListFields)" -f DarkYellow

                $meta.$Global:FilterLookupValues

                updateListData $meta.$Global:ListName $meta.$Global:ListFields $meta.$Global:SiteColumnTitle $meta.$Global:FilterLookupValues

                if ($meta.$Global:ListName -ne "")
                {
                    write-host "Working on site column: $($meta.$Global:SiteColumnName)" -f DarkYellow
                    if ($($meta.$Global:SiteColumnName) -ne "")
                    {
                        write-host "Column Title:$($meta.$Global:SiteColumnTitle)" -f DarkYellow            

                        updateSiteColumn $meta.$Global:ListName $meta.$Global:SiteColumnName $meta.$Global:SiteColumnTitle
                        
                        write-host " - Done!!" -f DarkRed
                    }
                }
            }

        } 
    }
}

function updateListData([string]$listName, [string]$listFields, [string]$lookupTitle, [string]$lookupFilter)
{
    $tab = "`t";
    $lists = $Global:web.Lists
    $Global:ctx.Load($lists);
    $Global:ctx.ExecuteQuery();
    $filter = $false;
    $filterColumnName = ""
    #$exists = $Global:web.Lists | ? {$_.Title -eq $csvFileName}
    #write-host $lists.Count;
    $list = $lists | ? {$_.Title -eq $listName}
        
    if ("$list" -ne "")
    {
        write-host "Deleting lists: $($list.Title)" -NoNewline -f Red
        
        $list.DeleteObject();
        $global:ctx.ExecuteQuery();

        write-host " - Done!!" -f DarkRed
    }

    write-host "Creating list: $($list.Title)" -NoNewline -f Yellow
    $newList = New-Object Microsoft.SharePoint.Client.ListCreationInformation
    $newList.Url = $listName.Replace(" ","")
    $newList.Title = $listName;
    $newList.TemplateType = 100;
                
    $list = $Global:web.Lists.Add($newList);
    $list.Description = $listName;
    $list.Update();
    $Global:ctx.ExecuteQuery();
    write-host " - Done!!" -f DarkYellow

    $list = $Global:web.Lists.GetByTitle($listName);
    $Global:ctx.Load($list);
    $Global:ctx.ExecuteQuery();

    #hide title field
    #$titleField = $list.Fields.GetByTitle("Title");
    $global:cached_lookupValues.clear();

    foreach($listField in $listFields.Split(','))
    {
        write-host "Column data: $listField"
        $fld_guid = [System.Guid]::NewGuid();
        #$fld_name = $listField.Split(':')[0]
        $fld_dn = $listField.Split(':')[0]
        $fld_name = $fld_dn.Replace(" ","");
        $fld_type = $listField.Split(':')[1]
        
        $fieldXML = "";
        if ($fld_type -clike "Text*")
        {
            $fieldXML = "<Field ID='{$fld_guid}' Name='$fld_name' StaticName='$fld_name' DisplayName='$fld_dn' Type='Text' Required='True' ></Field>"
        }else
        {
            $fld_lookup_isMult = $false;
            #type can be like Lookup,ListName
            if ($fld_type -contains "multi")
            {
                $fld_lookup_isMult = $true;
            }

            $fld_lookup = $fld_type.Split(';')[1]
            $fld_lookup_name = $fld_lookup.Split('@')[0]
            $fld_lookup_showfield = $fld_lookup.Split('@')[1]
            
            $fld_lookup_list = $Global:web.Lists.GetByTitle($fld_lookup_name);
            $global:ctx.Load($fld_lookup_list);
            $global:ctx.ExecuteQuery();

            $spfld_lookup_showfield = $fld_lookup_list.Fields.GetByTitle($fld_lookup_showfield);
            $global:ctx.Load($spfld_lookup_showfield);
            $global:ctx.ExecuteQuery();

            $fld_lookup_showfield = $spfld_lookup_showfield.InternalName;

            $filter = $false;
            write-host "Lookup filter:$lookupFilter"
            if (-NOT [String]::IsNullOrEmpty($lookupFilter))
            {
                $filterListTitle = $lookupFilter.Split('@')[0]; 
                $filterColumnName = $lookupFilter.Split('@')[1];
                if ($filterListTitle -eq $fld_lookup_list.Title)
                { 
                    $filter = $true;
                }

            }

            if ($global:cached_lookupValues.Count -eq 0)
            {
                cacheList $fld_lookup_showfield $filter $filterColumnName $fld_lookup_list $global:cached_lookupValues
            }

            $fld_lookup_listid = $fld_lookup_list.Id;
            $field_lookup_webid = $global:web.Id.ToString();
            $fieldXML = "<Field ID='{$fld_guid}' Name='$fld_name' StaticName='$fld_name' DisplayName='$fld_dn' Type='Lookup' Required='True' List='{$fld_lookup_listid}' Mult='$fld_lookup_isMult' ShowField='$fld_lookup_showfield' WebId='{$field_lookup_webid}'></Field>"
        }
        #$list.Fields.ad
        write-host "Adding columns to list:$fld_name  {fieldtype:$fld_type}" -f Green
        write-host $fieldXML -f DarkGreen
        
        $newfield = $list.Fields.AddFieldAsXml($fieldXML, $true,  [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint);
        $newfield.UpdateAndPushChanges($false);
        $Global:ctx.ExecuteQuery();

    }

    #<#
    $titleField = $list.Fields.GetByTitle("Title");
    $global:ctx.Load($titleField);
    $global:ctx.ExecuteQuery();

    if ($titleField)
    {
        $titleField.Required = $false;
        #$titleField.Hidden = $true;
        $titleField.Update();
        $global:ctx.ExecuteQuery();

        $defaultView = $list.DefaultView
        $viewFields = $defaultView.ViewFields;

        $global:ctx.Load($defaultView);
        $global:ctx.Load($viewFields);
        $global:ctx.ExecuteQuery();

        $defaultView.ViewFields.Remove("LinkTitle");
        $defaultView.Update();
        $global:ctx.ExecuteQuery();

    }
    #>
    write-host $global:cached_lookupValues.Count

    $datafilePath = "$PSScriptRoot\lists\$listName.txt"
    write-host "Loading file for list data: $datafilePath" -f Green

    #return;
    
    #add content to list
    $items = Import-Csv $datafilePath -Delimiter $Global:delmiterTab -Encoding UTF8
    $pending = $items.Count;

    foreach($item in $items)
    {
        write-host "Adding new item to List: ($pending pending)" -f Yellow
        write-host $item -f DarkYellow
        $pending--;
        $newItem = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        #$newItem["Title"] = $item;
        $spItem = $list.AddItem($newItem);

        foreach($listField in $listFields.Split(','))
        {
            $fld_name = $listField.Split(':')[0]
            write-host "****    $fld_name : {$($item.$fld_name)}    *****"

            $fld = $list.Fields.GetByTitle($fld_name);
            #write-host $fld.InternalName -f Magenta
            $Global:ctx.Load($fld);
            $Global:ctx.ExecuteQuery();

            $fld_value = $item.$fld_name;
            $fld_value_key = $fld_value;
            if ($fld.TypeAsString -eq "Lookup")
            {
                if ($filter)
                {
                    $fld_filter = $item.$filterColumnName
                    $fld_value_key = "$fld_Value$Global:filterJoin$fld_filter"
                    write-host "---------"
                }
                $fld_value = getLookupValues $fld_value_key $global:cached_lookupValues
                $fld.ValidateSetValue($spItem, $fld_value);    
            }else
            {
                $spItem[$fld.InternalName] =$fld_value;
                #$spItem["Title"] = $fld_value;
            }
            $spItem.Update();
            $Global:ctx.ExecuteQuery();
        }
    }
}

$global:cached_lookupValues = New-Object "System.Collections.Generic.Dictionary[string,string]" ([System.StringComparer]::OrdinalIgnoreCase)
function cacheList([string]$columnName,  [bool]$filter, [string]$filterColumnName, [Microsoft.SharePoint.Client.List]$spList, [System.Collections.Generic.Dictionary[string,string]]$collection)
{
    $collection.Clear();
    write-host "Filter: $filter column name:$filterColumnName" -f Green
    $query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery();
    #$query = New-Object Microsoft.SharePoint.Client.CamlQuery
    #$query.ViewXml = "<View><RowLimit>1000</RowLimit></View>"
    $items = $spList.GetItems($query);
            
    $Global:ctx.Load($items);
    $Global:ctx.ExecuteQuery();
    Write-Host "Caching $($spList.Title) ColumnName:$columnName Items count:$($items.Count)" -f Yellow
    
	foreach($item in $items)
	{
        
        $Global:ctx.Load($item);
        $Global:ctx.ExecuteQuery();
        $key=""
		$column =$item[$columnName]
        if (-NOT $filter)
        {
            $key = "$column"
        }
        else
        {
            $columnFilter = $item[$filterColumnName].LookupValue
            $key = "$column$Global:filterJoin$columnFilter"
        }
        Write-Host "ID: $($item.Id) Key:$key" -f DarkYellow
        if (-Not $collection.ContainsKey($key))
        {
		    $lookupValue = [string]::Format("{0};#{1};#", $item.ID,$column);
		    $collection.Add($key, $lookupValue);
        }
	}

    Write-Host "Cached: $($global:cached_lookupValues.Count)" -f Yellow
}

function getLookupValues([System.String]$ForLookupValues, [System.Collections.Generic.Dictionary[string,string]]$collection)
{
    $values = "";
    $LookupValues = $ForLookupValues.Split(";");
    $count = $LookupValues.Count;

    foreach ($value in $LookupValues)
    {
        if ($collection.ContainsKey($value))
        {
            $values += $collection[$value] 
            $count--;
        }
    }
    $values = $values.TrimEnd(";#");
    Write-Host "Lookup Value: $values ($count missing) for input: $ForLookupValues" -f DarkYellow
    if ($count -ne 0)
    {
        $values = "";
        Write-Host "Not able to find all lookup values!!" -f Red
    }
    return $values;
}

function updateSiteColumn([string]$listName, [string]$columnName, [string]$titleField)
{

    $list = $Global:web.Lists.GetByTitle($listName);
    $Global:ctx.Load($list);
    $Global:ctx.ExecuteQuery();

    write-host $list.Id $titleField
    $field = $Global:web.Fields.GetByInternalNameOrTitle($columnName);
    $Global:ctx.Load($field);
    $Global:ctx.ExecuteQuery();

    $schema = $field.SchemaXml
    $xml = [xml]$schema
    #write-host $xml.OuterXml

    $listValue = $xml.Field.Attributes["List"].Value
    $showFieldValue = $xml.Field.Attributes["ShowField"].Value
    write-host "Before Update: ListId=$listValue ShowField=$showFieldValue" -f Cyan 

    write-host "Updating schemaxml";
    $xml.Field.Attributes["List"].Value = $list.Id;
    $xml.Field.Attributes["ShowField"].Value = $titleField;
    $field.SchemaXml = $xml.OuterXml;
    $field.UpdateAndPushChanges($true);
    $Global:ctx.ExecuteQuery();

    $listValue = $xml.Field.Attributes["List"].Value
    $showFieldValue = $xml.Field.Attributes["ShowField"].Value
    write-host "After Update: ListId=$listValue ShowField=$showFieldValue" -f Cyan 
    
}

function deleteListData()
{
    #$query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery();
    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml = "<View><RowLimit>1000</RowLimit></View>"
    $items = $list.GetItems($query);
            
    $Global:ctx.Load($items);
    $Global:ctx.ExecuteQuery();

    $count = $items.Count
            
    #foreach($item in $items)
    if ($count -gt 0)
    {
        write-host "Deleting all items:$count" -f red
        for($i=0; $i -le $count-1; $i++)
        {
            write-host "Deleting item:$i ID: $($items[$i].Id)" -f red
            #$items[$i].DeleteObject();
            $list.GetItemById($items[$i].Id).DeleteObject();
        }
        $Global:ctx.ExecuteQuery();
        #return;
    }
}


main
#pause
