# Sharepoint Scripts

## Usage for Sharepoint-Mass-File-Check-In
### Description: The Sharepoint Mass File Check In will ***recursively*** check in all checked-out files in the specified directory which will also recursively check in files in sub-directories. (Note: The checked-out file has to be checked-out to the logged-in user.)
1. In Sharepoint > Click the Site > Documents > *Navigate to the path (It should have AllItems.aspx in the URL)*
1. .\powershell_sharepoint_mass_check_in.ps1 -Sharepoint_Login_Email "username@example.com" -Sharepoint_Site_List_Full_URL_Encoded "https://example.sharepoint.com/sites/Example-Site/Shared%20Documents/Forms/AllItems.aspx?viewid=EXAMPLE&id=%2Fsites%2FExample%2DSite%2FShared%20Documents%2FPersonal%2FTEST%20USER%2FExample%5FExample%2FExample%20%2D%20Work%20Projects%2F%5BExample%5D%20Clients"

![image](https://user-images.githubusercontent.com/30126475/119512830-4385e000-bda6-11eb-8929-a023f54f24a7.png)


## Usage for Sharepoint-Folder-Export
### Description: The Sharepoint Folder Export will ***recursively*** export folders that is last modified within the specified start date and end date into a CSV file with the folder name and last modified date.
1. In Sharepoint > Click the Site > Documents > *Navigate to the path (It should have AllItems.aspx in the URL)*
1. .\powershell_sharepoint_mass_folder_export.ps1 -Sharepoint_Login_Email "username@example.com" -Sharepoint_Site_List_Full_URL_Encoded "https://example.sharepoint.com/sites/Example-Site/Shared%20Documents/Forms/AllItems.aspx?viewid=EXAMPLE&id=%2Fsites%2FExample%2DSite%2FShared%20Documents%2FPersonal%2FTEST%20USER%2FExample%5FExample%2FExample%20%2D%20Work%20Projects%2F%5BExample%5D%20Clients" -CSV_Export_Filename "Example.csv" -Sharepoint_Last_Modified_Start_Date "25/5/2021 1:41:05 PM" -Sharepoint_Last_Modified_End_Date "25/5/2021 1:42:13 PM"

![image](https://user-images.githubusercontent.com/30126475/119519486-0ae90500-bdac-11eb-8d26-1442fe4e1c9d.png)
![image](https://user-images.githubusercontent.com/30126475/119519392-f573db00-bdab-11eb-983c-5f98c3649d64.png) 

## Usage for Sharepoint-Mass-File-Upload
### Description: The Sharepoint Mass File Upload will ***recursively*** upload files from the specified local directory into Sharepoint while also checking-out & checking-in applicable files.
1. In Sharepoint > Click the Site > Documents > *Navigate to the path (It should have AllItems.aspx in the URL)*
2. .\powershell_sharepoint_mass_check_in.ps1 -Sharepoint_Login_Email "username@example.com" -Sharepoint_Site_List_Full_URL_Encoded "https://example.sharepoint.com/sites/Example-Site/Shared%20Documents/Forms/AllItems.aspx?viewid=<REDACTED>&id=%2Fsites%2FExample%2DSite%2FShared%20Documents%2FPersonal%2FTEST%20USER%2FExample%5FExample%2FExample%20%2D%20Work%20Projects%2F%5BExample%5D%20Clients" -Local_Directory_Full_Path 'D:\Desktop\EXAMPLE\FOLDER\'
  
  ![image](https://user-images.githubusercontent.com/30126475/119648189-c153f500-be53-11eb-9409-1ca9456a281c.png)
