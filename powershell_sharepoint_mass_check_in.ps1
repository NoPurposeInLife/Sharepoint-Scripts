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


# powershell_sharepoint_mass_check_in.ps1
[CmdletBinding()]

# https://ss64.com/ps/syntax-template.html
# https://www.c-sharpcorner.com/blogs/mfa-multi-factor-authentication-authentication-using-powershell-in-sharepoint-online
# https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps#installation

# https://guimatheus92.medium.com/get-sharepoint-items-from-lists-by-using-powershell-8126dfbb0c43

# https://www.c-sharpcorner.com/blogs/retrieve-sharepoint-list-items-using-pnp-powershell

# Install-Module SharePointPnPPowerShellOnline

# https://www.sharepointdiary.com/2018/03/sharepoint-online-powershell-to-get-folder-in-document-library.html

# https://www.sharepointdiary.com/2016/10/sharepoint-online-how-to-check-in-document-using-powershell.html

# https://sharepoint.stackexchange.com/questions/103610/how-to-checkin-all-checked-out-files-via-powershell/103622


PARAM ( 
    [string] $Sharepoint_Site_List_Full_URL_Encoded = $(throw "-Sharepoint_Site_List_Full_URL_Encoded is required."),
    [string] $Sharepoint_Login_Email = $(throw "-Sharepoint_Login_Email is required.")
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
    
    # https://stackoverflow.com/questions/53766303/how-do-i-split-parse-a-url-string-into-an-object
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

Function Sharepoint_Mass_Check_In{
    Param(
        $pnpcontext_obj_current_pnp_context,
        [string] $Sharepoint_Login_Email,
        [string] $string_obj_sharepoint_folder_site_relative_url    
    )
    
    # Write-Host -ForegroundColor Red $string_obj_sharepoint_folder_site_relative_url
    
    # https://www.sharepointdiary.com/2018/03/sharepoint-online-powershell-to-get-folder-in-document-library.html
    $folder_obj_sharepoint_folder_items = Get-PnPFolderItem -FolderSiteRelativeUrl $string_obj_sharepoint_folder_site_relative_url
    
    ForEach($item_obj_sharepoint_folder_item in $folder_obj_sharepoint_folder_items)
    {
        
        # $item_obj_sharepoint_folder_item | Get-Member
        #Exit 1
        Write-Host "Name" : $item_obj_sharepoint_folder_item.Name
        Write-Host "UniqueId" : $item_obj_sharepoint_folder_item.UniqueId
        Write-Host "Type" : $item_obj_sharepoint_folder_item.GetType()
        #Write-Host "ServerRelativePath" : $item_obj_sharepoint_folder_item.ServerRelativePath
        #Write-Host "ServerRelativeURL" : $item_obj_sharepoint_folder_item.ServerRelativeUrl
        
        $string_obj_sharepoint_item_name = [string] $item_obj_sharepoint_folder_item.Name
        $string_obj_sharepoint_item_type = [string] $item_obj_sharepoint_folder_item.GetType()
        
        
        if ($string_obj_sharepoint_item_type -eq "Microsoft.SharePoint.Client.Folder") {
            
            Write-Host -ForegroundColor Red "Result" : "None"
            Write-Host "---------------------------" 
            $string_obj_sharepoint_folder_site_new_relative_url = $string_obj_sharepoint_folder_site_relative_url + "/" + $string_obj_sharepoint_item_name
            Sharepoint_Mass_Check_In -pnpcontext_obj_current_pnp_context $pnpcontext_obj_current_pnp_context -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_folder_site_relative_url $string_obj_sharepoint_folder_site_new_relative_url
        } else {
            
            # https://sharepoint.stackexchange.com/questions/103610/how-to-checkin-all-checked-out-files-via-powershell/103622
            # https://github.com/pnp/PnP-PowerShell/issues/1370
            
            
            $user_obj_sharepoint_folder_item_checked_out_by_user = $item_obj_sharepoint_folder_item.CheckedOutByUser
            
            $pnpproperty_obj_sharepoint_folder_item_check_out_user_email = (Get-PnPProperty -ClientObject $user_obj_sharepoint_folder_item_checked_out_by_user -Property "Email" 2> $null)
            
            $string_obj_sharepoint_folder_item_check_out_user_email = [string] $pnpproperty_obj_sharepoint_folder_item_check_out_user_email
                
            # $pnpproperty_obj_sharepoint_folder_item_check_out_user_email
            
            # Check if the file is checked out
            If (($item_obj_sharepoint_folder_item.CheckOutType -ne "None") -And ($string_obj_sharepoint_folder_item_check_out_user_email -eq $Sharepoint_Login_Email)) {
                # https://www.sharepointdiary.com/2016/10/sharepoint-online-how-to-check-in-document-using-powershell.html
                
                $item_obj_sharepoint_folder_item.CheckIn("File Checked In.",[Microsoft.SharePoint.Client.CheckinType]::MinorCheckIn)
                $pnpcontext_obj_current_pnp_context.ExecuteQuery()
                Write-Host -ForegroundColor Green "Result" : "File Checked In."
                Write-Host "---------------------------" 
            } else {
                Write-Host -ForegroundColor Red "Result" : "File Not Checked Out or Login Email Does Not Match Checked Out To Email"
                Write-Host "---------------------------" 
            }
        }
    }
    
    
    return
}

Function Main{
    Param(
        [string] $string_obj_sharepoint_site_list_full_url_encoded,
        [string] $Sharepoint_Login_Email
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
        
        # Sharepoint Mass Check In
        Write-Host "---------------------------" 
        Sharepoint_Mass_Check_In -pnpcontext_obj_current_pnp_context $pnpcontext_obj_current_pnp_context -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_folder_site_relative_url $string_obj_sharepoint_folder_site_relative_url
        

    } catch {  
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        $msg = $e.Message 

        Write-Host -ForegroundColor Red "caught exception: $e at $line" 
    }
}

#----------------[ Imports ]---------------
if (!(Get-Module "PnP.PowerShell")) {
    Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser
}
Add-Type -AssemblyName System.Web

#----------------[ Main Execution ]---------------
Main -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_site_list_full_url_encoded $Sharepoint_Site_List_Full_URL_Encoded