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

PARAM ( 
    [string] $Sharepoint_Site_List_Full_URL_Encoded = $(throw "-Sharepoint_Site_List_Full_URL_Encoded is required."),
    [string] $Sharepoint_Login_Email = $(throw "-Sharepoint_Login_Email is required."),
    [string] $Local_Directory_Full_Path = $(throw "-Local_Directory_Full_Path is required.")
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


Function Get_Sharepoint_Folder_Site_Full_Relative_URL{
    Param(
        [string] $string_obj_sharepoint_site_list_full_url_decoded
    )
    
    
    $system_uri_obj_sharepoint_site_list_full_url_decoded = [System.Uri]$string_obj_sharepoint_site_list_full_url_decoded
    
    $string_obj_sharepoint_site_full_url_decoded = "/" + $system_uri_obj_sharepoint_site_list_full_url_decoded.Segments[1] + $system_uri_obj_sharepoint_site_list_full_url_decoded.Segments[2]
    
    $uri_obj_sharepoint_site_list_full_url_decoded = [System.Uri] $string_obj_sharepoint_site_list_full_url_decoded
    
    $system_web_httputility_obj_sharepoint_site_list_full_url_decoded_parsed_query_string = [System.Web.HttpUtility]::ParseQueryString($uri_obj_sharepoint_site_list_full_url_decoded.Query)
    
    $string_obj_sharepoint_folder_site_full_relative_url_decoded = $system_web_httputility_obj_sharepoint_site_list_full_url_decoded_parsed_query_string['id']
    
    return [string] $string_obj_sharepoint_folder_site_full_relative_url_decoded
}


Function Sharepoint_Authentication{
    Param(
        [string] $string_obj_sharepoint_site_url
    )
    Connect-PnPOnline -Url $string_obj_sharepoint_site_url -UseWebLogin -ForceAuthentication
    return
}

function Set-EscapeCharacters_Powershell {
    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $string
    )
    
    # https://stackoverflow.com/questions/57965466/how-to-escape-special-characters-in-powershell
    
    $string = $string -replace '\*', '`*'
    #$string = $string -replace '\\', '`\'
    $string = $string -replace '\~', '`~'
    $string = $string -replace '\;', '`;'
    $string = $string -replace '\(', '`('
    $string = $string -replace '\%', '%%'
    $string = $string -replace '\?', '`?'
    #$string = $string -replace '\.', '`.'
    #$string = $string -replace '\:', '`:'
    $string = $string -replace '\@', '`@'
    $string = $string -replace '\/', '`/'
    #$string = $string -replace ' ', '` '
    $string = $string -replace '\[', '`['
    $string = $string -replace '\]', '`]'
    #$string = $string -replace '\-', '`-'
    #$string = $string -replace '\#', '`#'
    return [string] $string
}

function Set-EscapeCharacters_Sharepoint {
    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $string
    )
    $string = $string -replace '\#', '`#'
    return [string] $string
}

function Set-EscapeCharacters_Custom {
    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $string
    )
    $string = $string -replace '%', '%25'
    return [string] $string
}

Function Sharepoint_Mass_File_Upload{
    Param(
        $pnpcontext_obj_current_pnp_context,
        [string] $Sharepoint_Login_Email,
        [string] $string_obj_sharepoint_folder_site_full_relative_url,
        [string] $string_obj_local_directory_full_path
    )
        
    # https://stackoverflow.com/questions/39825440/check-if-a-path-is-a-folder-or-a-file-in-powershell
    $item_obj_files_and_folders = Get-ChildItem -LiteralPath $string_obj_local_directory_full_path -Force -Include *

    foreach ($f in $item_obj_files_and_folders) {
        $string_obj_item_full_path = $f.FullName
        $string_obj_item_short_name = $f.Name
        
        if (Test-Path -LiteralPath $string_obj_item_full_path -PathType Container) {
            # https://veronicageek.com/microsoft-365/sharepoint-online/create-a-folder-structure-in-sharepoint-online-using-powershell-pnp-from-file-shares/2019/02/
            # https://sharepoint.stackexchange.com/questions/151074/list-files-in-a-specific-folder-in-a-document-library
            # https://stackoverflow.com/questions/63825327/add-pnpfolder-fails-for-folder-names-containing-a-hash
            
            Write-Host "Name" : $string_obj_item_short_name
            Write-Host "Type" : "Folder"
            Write-Host "Sharepoint Path" : ($string_obj_sharepoint_folder_site_full_relative_url + '/' + $string_obj_item_short_name)  
            

            $web_obj_current_pnp_context_web = $pnpcontext_obj_current_pnp_context.Web
            $pnpcontext_obj_current_pnp_context.Load($web_obj_current_pnp_context_web)
                        
            # https://sharepoint.stackexchange.com/questions/292665/spo-csom-powershell-manage-folders-with-and-in-the-name            
            $folder_obj_target_sharepoint_folder = $web_obj_current_pnp_context_web.GetFolderByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($string_obj_sharepoint_folder_site_full_relative_url))
            
            $pnpcontext_obj_current_pnp_context.Load($folder_obj_target_sharepoint_folder)
            $pnpcontext_obj_current_pnp_context.ExecuteQuery()
       
            $folder_obj_target_sharepoint_folder.Folders.Add($string_obj_item_short_name) > $null
            $folder_obj_target_sharepoint_folder.Context.ExecuteQuery()

            # Add-PnPFolder -Name $string_obj_item_short_name -Folder $string_obj_sharepoint_folder_site_full_relative_url > $null
            
            
                      
            Write-Host -ForegroundColor Green "Result" : "Folder Created."
            Write-Host "---------------------------" 
            
            $string_obj_sharepoint_folder_site_new_full_relative_url = ($string_obj_sharepoint_folder_site_full_relative_url + '/' + $string_obj_item_short_name)
                        
            
            $string_obj_local_directory_new_full_path = ($string_obj_local_directory_full_path + "\" + $string_obj_item_short_name)
                        
            Sharepoint_Mass_File_Upload -pnpcontext_obj_current_pnp_context $pnpcontext_obj_current_pnp_context -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_folder_site_full_relative_url $string_obj_sharepoint_folder_site_new_full_relative_url -string_obj_local_directory_full_path $string_obj_local_directory_new_full_path
        } else {            
        
            
            # https://sharepoint.stackexchange.com/questions/159085/upload-files-to-sharepoint-intranet-site-using-powershell
            # https://www.sharepointdiary.com/2018/01/sharepoint-online-upload-file-to-folder-using-powershell.html\
            # https://www.sharepointdiary.com/2020/05/upload-large-files-to-sharepoint-online-using-powershell.html
            
            Write-Host "Name" : $string_obj_item_short_name
            Write-Host "Type" : "File"
            Write-Host "Sharepoint Path" : ($string_obj_sharepoint_folder_site_full_relative_url + '/' + $string_obj_item_short_name)
            
            #Get the Target Folder to upload
            $web_obj_current_pnp_context_web = $pnpcontext_obj_current_pnp_context.Web
            $pnpcontext_obj_current_pnp_context.Load($web_obj_current_pnp_context_web)
            
            $folder_obj_target_sharepoint_folder = $web_obj_current_pnp_context_web.GetFolderByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($string_obj_sharepoint_folder_site_full_relative_url))
            
            $pnpcontext_obj_current_pnp_context.Load($folder_obj_target_sharepoint_folder)
            $pnpcontext_obj_current_pnp_context.ExecuteQuery()
            
            
            #Get the source file from disk
            $stream_obj_source_file_stream = ([System.IO.FileInfo] ($f)).OpenRead()
            #Get File Name from source file path
            $string_obj_sharepoint_site_url = $string_obj_sharepoint_folder_site_full_relative_url+"/"+$string_obj_item_short_name
         
            #Upload the File to SharePoint Library Folder

            # https://www.sharepointdiary.com/2016/10/check-if-file-exists-in-document-library-using-powershell-csom.html
            $array_obj_file_site_relative_url = $string_obj_sharepoint_folder_site_full_relative_url.split("/")
            
            $string_obj_file_site_relative_url_front = ("/" + $array_obj_file_site_relative_url[1] +"/" + $array_obj_file_site_relative_url[2])

            $string_obj_file_site_relative_url = $string_obj_sharepoint_site_url.replace($string_obj_file_site_relative_url_front,"")
            
            $sharepoint_file_obj_target_file = Get-PnPFile -Url $string_obj_file_site_relative_url  -ErrorAction SilentlyContinue
            
            if ($sharepoint_file_obj_target_file) {
                try {
                    $sharepoint_file_obj_target_file.CheckOut()
                } catch {
                }
            }
            
            if ($pnpcontext_obj_current_pnp_context.HasPendingRequest) {
                try {
                    $pnpcontext_obj_current_pnp_context.ExecuteQuery()
                } catch {
                }
            }
            
            
            $pnpcontext_obj_current_pnp_context.RequestTimeout = [System.Threading.Timeout]::Infinite
            [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($pnpcontext_obj_current_pnp_context, $string_obj_sharepoint_site_url, $stream_obj_source_file_stream,$true)
                    

            $sharepoint_file_obj_uploaded_file = Get-PnPFile -Url $string_obj_file_site_relative_url  -ErrorAction SilentlyContinue
            
            $sharepoint_file_obj_uploaded_file.CheckIn("File Checked In.",[Microsoft.SharePoint.Client.CheckinType]::MinorCheckIn)
            $pnpcontext_obj_current_pnp_context.ExecuteQuery()
            
            
            Write-Host -ForegroundColor Green "Result" : "File Uploaded and Checked In."
            Write-Host "---------------------------" 
        }
    }
      
    
    return
}

Function Main{
    Param(
        [string] $string_obj_sharepoint_site_list_full_url_encoded,
        [string] $Sharepoint_Login_Email,
        [string] $string_obj_local_directory_full_path
    )
    
    try {  
        # URL Decode Sharepoint URL
        $string_obj_sharepoint_site_list_full_url_decoded = URL_Decode_String -URL_Encoded_String $string_obj_sharepoint_site_list_full_url_encoded
        
        # Get Sharepoint Site URL
        $string_obj_sharepoint_site_url = Get_Sharepoint_Site_URL -string_obj_sharepoint_site_list_full_url_decoded $string_obj_sharepoint_site_list_full_url_decoded
        
        Write-Host -ForegroundColor Green "Sharepoint Site URL" :  $string_obj_sharepoint_site_url
        
        # Get Sharepoint Folder Site Relative URL
        $string_obj_sharepoint_folder_site_full_relative_url = Get_Sharepoint_Folder_Site_Full_Relative_URL -string_obj_sharepoint_site_list_full_url_decoded $string_obj_sharepoint_site_list_full_url_decoded
        
        Write-Host -ForegroundColor Green "Sharepoint Full Relative URL" : $string_obj_sharepoint_folder_site_full_relative_url
        
        # Authenticate to Sharepoint
        Sharepoint_Authentication -string_obj_sharepoint_site_url $string_obj_sharepoint_site_url
        
        # Get Current PNP Context
        $pnpcontext_obj_current_pnp_context = Get-PnPContext
        
        # Sharepoint Mass File Upload
        Write-Host "---------------------------" 
        Sharepoint_Mass_File_Upload -pnpcontext_obj_current_pnp_context $pnpcontext_obj_current_pnp_context -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_folder_site_full_relative_url $string_obj_sharepoint_folder_site_full_relative_url -string_obj_local_directory_full_path $string_obj_local_directory_full_path
        

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
Main -Sharepoint_Login_Email $Sharepoint_Login_Email -string_obj_sharepoint_site_list_full_url_encoded $Sharepoint_Site_List_Full_URL_Encoded -string_obj_local_directory_full_path $Local_Directory_Full_Path
