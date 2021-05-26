<#
.SYNOPSIS
    .
.DESCRIPTION
    .
.PARAMETER Path
    The path to the .
.PARAMETER LiteralPath
    Specifies a path to one or more locations. Unlike Path, the value of 
    LiteralPath is used exactly as it is typed. No characters are interpreted 
    as wildcards. If the path includes escape characters, enclose it in single
    quotation marks. Single quotation marks tell Windows PowerShell not to 
    interpret any characters as escape sequences.
#>


# powershell_sharepoint_folder_export.ps1
[CmdletBinding()]

PARAM ( 
    [string] $Sharepoint_Site_List_Full_URL_Encoded = $(throw "-Sharepoint_Site_List_Full_URL_Encoded is required."),
    [string] $Sharepoint_Login_Email = $(throw "-Sharepoint_Login_Email is required."),
    [string] $CSV_Export_Filename = $(throw "-CSV_Export_Filename is required."),
    [string] $Sharepoint_Last_Modified_Start_Date = $(throw "-Sharepoint_Last_Modified_Start_Date is required."),
    [string] $Sharepoint_Last_Modified_End_Date = $(throw "-Sharepoint_Last_Modified_End_Date is required.")
)



#----------------[ Declarations ]----------------

# Set Error Action
$ErrorActionPreference = "Continue"

#----------------[ Functions ]------------------
Function URL_Decode_String{
    Param(
        [string] $URL_Encoded_String
    )
    
    $URL_Decoded_String = [System.Web.HttpUtility]::UrlDecode($URL_Encoded_String)
    return [string] $URL_Decoded_String
}

Function Get_Sharepoint_Site_URL{
    Param(
        [string] $string_obj_sharepoint_site_list_full_url_decoded
    )
    $system_uri_obj_sharepoint_site_list_full_url_decoded = [System.Uri]$string_obj_sharepoint_site_list_full_url_decoded
    
    $string_obj_sharepoint_site_full_url_decoded = $system_uri_obj_sharepoint_site_list_full_url_decoded.Scheme + "://" + $system_uri_obj_sharepoint_site_list_full_url_decoded.Host + "/" + $system_uri_obj_sharepoint_site_list_full_url_decoded.Segments[1] + $system_uri_obj_sharepoint_site_list_full_url_decoded.Segments[2]
    
    return [string] $string_obj_sharepoint_site_full_url_decoded
}

Function Get_Sharepoint_Folder_Site_Relative_URL{
    Param(
        [string] $string_obj_sharepoint_site_list_full_url_decoded
    )
    
    
    $system_uri_obj_sharepoint_site_list_full_url_decoded = [System.Uri]$string_obj_sharepoint_site_list_full_url_decoded
    
    $string_obj_sharepoint_site_full_url_decoded = "/" + $system_uri_obj_sharepoint_site_list_full_url_decoded.Segments[1] + $system_uri_obj_sharepoint_site_list_full_url_decoded.Segments[2]
    
    $uri_obj_sharepoint_site_list_full_url_decoded = [System.Uri] $string_obj_sharepoint_site_list_full_url_decoded
    
    $system_web_httputility_obj_sharepoint_site_list_full_url_decoded_parsed_query_string = [System.Web.HttpUtility]::ParseQueryString($uri_obj_sharepoint_site_list_full_url_decoded.Query)
    
    $string_obj_sharepoint_site_list_full_url_decoded_id = $system_web_httputility_obj_sharepoint_site_list_full_url_decoded_parsed_query_string['id']
    
    if ([string]::IsNullOrEmpty($string_obj_sharepoint_site_list_full_url_decoded_id)) {
        $string_obj_sharepoint_folder_site_relative_url_decoded = "/" + $system_uri_obj_sharepoint_site_list_full_url_decoded.Segments[3]
        $string_obj_sharepoint_folder_site_relative_url_decoded = $string_obj_sharepoint_folder_site_relative_url_decoded.Substring(0,$string_obj_sharepoint_folder_site_relative_url_decoded.Length-1)
    } else {
        $string_obj_sharepoint_folder_site_relative_url_decoded = "/" + $string_obj_sharepoint_site_list_full_url_decoded_id -replace $string_obj_sharepoint_site_full_url_decoded,""
    }
    
    $string_obj_sharepoint_folder_site_relative_url_decoded = URL_Decode_String -URL_Encoded_String $string_obj_sharepoint_folder_site_relative_url_decoded
    
    return [string] $string_obj_sharepoint_folder_site_relative_url_decoded
}

Function Sharepoint_Authentication{
    Param(
        [string] $string_obj_sharepoint_site_url
    )
    Connect-PnPOnline -Url $string_obj_sharepoint_site_url -UseWebLogin -ForceAuthentication
    return
}

Function Sharepoint_Folder_Export{
    Param(
        $pnpcontext_obj_current_pnp_context,
        [string] $Sharepoint_Login_Email,
        [string] $string_obj_sharepoint_folder_site_relative_url,
        [string] $string_obj_csv_export_filename,
        [string] $string_obj_sharepoint_last_modified_start_date,
        [string] $string_obj_sharepoint_last_modified_end_date        
    )
    
    # Write-Host -ForegroundColor Red $string_obj_sharepoint_folder_site_relative_url
    
    $folder_obj_sharepoint_folder_items = Get-PnPFolderItem -FolderSiteRelativeUrl $string_obj_sharepoint_folder_site_relative_url
    
    ForEach($item_obj_sharepoint_folder_item in $folder_obj_sharepoint_folder_items)
    {
        
        $string_obj_sharepoint_item_name = [string] $item_obj_sharepoint_folder_item.Name
        $string_obj_sharepoint_item_type = [string] $item_obj_sharepoint_folder_item.GetType()
        
        if ($string_obj_sharepoint_item_type -eq "Microsoft.SharePoint.Client.Folder") {
            
            $string_obj_sharepoint_item_last_modified_date = $item_obj_sharepoint_folder_item.TimeLastModified.toString("dd/MM/yyyy HH:mm:ss")
            
            
            $date_obj_sharepoint_last_modified_start_date = (Get-Date $string_obj_sharepoint_last_modified_start_date).toString("dd/MM/yyyy HH:mm:ss")
            $date_obj_sharepoint_last_modified_end_date = (Get-Date $string_obj_sharepoint_last_modified_end_date).toString("dd/MM/yyyy HH:mm:ss")
            $date_obj_sharepoint_last_modified_date = (Get-Date $string_obj_sharepoint_item_last_modified_date).toString("dd/MM/yyyy HH:mm:ss")
            
            
            
            # Write-Host $date_obj_sharepoint_last_modified_start_date
            # Write-Host $date_obj_sharepoint_last_modified_end_date
            # Write-Host $date_obj_sharepoint_last_modified_date
        
            if($date_obj_sharepoint_last_modified_date -ge $date_obj_sharepoint_last_modified_start_date -and $date_obj_sharepoint_last_modified_date -le $date_obj_sharepoint_last_modified_end_date){
                Write-Host "Name" : $item_obj_sharepoint_folder_item.Name
                Write-Host "UniqueId" : $item_obj_sharepoint_folder_item.UniqueId
                Write-Host "Type" : $item_obj_sharepoint_folder_item.GetType()
                Write-Host "Last Modified" : $item_obj_sharepoint_folder_item.TimeLastModified
                Write-Host "---------------------------" 
                
                $string_obj_sharepoint_item_name + "," + $string_obj_sharepoint_item_last_modified_date | Out-File -Append $CSV_Export_Filename -Encoding UTF8
                
            }
            
            $string_obj_sharepoint_folder_site_new_relative_url = $string_obj_sharepoint_folder_site_relative_url + "/" + $string_obj_sharepoint_item_name
            Sharepoint_Folder_Export -pnpcontext_obj_current_pnp_context $pnpcontext_obj_current_pnp_context -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_folder_site_relative_url $string_obj_sharepoint_folder_site_new_relative_url -string_obj_csv_export_filename $CSV_Export_Filename -string_obj_sharepoint_last_modified_start_date $Sharepoint_Last_Modified_Start_Date -string_obj_sharepoint_last_modified_end_date $Sharepoint_Last_Modified_End_Date
        } 
    }
    
    
    return
}

Function Main{
    Param(
        [string] $string_obj_sharepoint_site_list_full_url_encoded,
        [string] $Sharepoint_Login_Email,
        [string] $string_obj_csv_export_filename,
        [string] $string_obj_sharepoint_last_modified_start_date,
        [string] $string_obj_sharepoint_last_modified_end_date
    )
    
    try {  
        # URL Decode Sharepoint URL
        $string_obj_sharepoint_site_list_full_url_decoded = URL_Decode_String -URL_Encoded_String $string_obj_sharepoint_site_list_full_url_encoded
        
        # Get Sharepoint Site URL
        $string_obj_sharepoint_site_url = Get_Sharepoint_Site_URL -string_obj_sharepoint_site_list_full_url_decoded $string_obj_sharepoint_site_list_full_url_decoded
        
        Write-Host -ForegroundColor Green "Sharepoint Site URL" :  $string_obj_sharepoint_site_url
        
        # Get Sharepoint Folder Site Relative URL
        $string_obj_sharepoint_folder_site_relative_url = Get_Sharepoint_Folder_Site_Relative_URL -string_obj_sharepoint_site_list_full_url_decoded $string_obj_sharepoint_site_list_full_url_decoded
        
        Write-Host -ForegroundColor Green "Sharepoint Relative URL" : $string_obj_sharepoint_folder_site_relative_url
        
        # Authenticate to Sharepoint
        Sharepoint_Authentication -string_obj_sharepoint_site_url $string_obj_sharepoint_site_url
        
        # Get Current PNP Context
        $pnpcontext_obj_current_pnp_context = Get-PnPContext
        
        # Sharepoint Folder Export
        Write-Host "---------------------------" 
        Sharepoint_Folder_Export -pnpcontext_obj_current_pnp_context $pnpcontext_obj_current_pnp_context -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_folder_site_relative_url $string_obj_sharepoint_folder_site_relative_url -string_obj_csv_export_filename $CSV_Export_Filename -string_obj_sharepoint_last_modified_start_date $Sharepoint_Last_Modified_Start_Date -string_obj_sharepoint_last_modified_end_date $Sharepoint_Last_Modified_End_Date
        

    } catch {  
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        $msg = $e.Message 

        Write-Host -ForegroundColor Red "caught exception: $e at $line" 
    }
}

#----------------[ Imports ]---------------
if (!(Get-Module "PnP.PowerShell")) {
    Install-Module SharePointPnPPowerShellOnline
}
Add-Type -AssemblyName System.Web

#----------------[ Main Execution ]---------------
Main -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_site_list_full_url_encoded $Sharepoint_Site_List_Full_URL_Encoded -string_obj_csv_export_filename $CSV_Export_Filename -string_obj_sharepoint_last_modified_start_date $Sharepoint_Last_Modified_Start_Date -string_obj_sharepoint_last_modified_end_date $Sharepoint_Last_Modified_End_Date
