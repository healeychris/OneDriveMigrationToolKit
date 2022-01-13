<#	
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	06/01/2022
    Created by:   	Chris Healey
    Organization: 	
    Filename:     	MainController.ps1
    Project path:   https://github.com/healeychris/OneDriveMigrationToolKit
    ===========================================================================
    .DESCRIPTION
    This script is used discover and assist with OneDrive Migrations
    .NOTES

#>


####### Variable list #######
$Version                                                = "0.1"                                          # Version of script
Clear-Host                                                                                               # Clear screen
$host.ui.RawUI.WindowTitle                              = 'OneDrive Migration Tool Kit'                  # Set name in task bar
$DBName                                                 = "OneDriveDiscovery.db"                         # Database Name used to store and create if not existing
$ExcludeListFile                                        ='.\ExcludeUsers.txt'                            # List of Users to exclude by SAmaccountName
$ExcludeGroupFile                                       ='.\ExcludeGroups.txt'                           # Lit of AD Groups by DN to exclude if a users is a member of


$MFALoginRequired                                       = $false                                         # Is MFA required to login
$RequireConnectMicrosoftOnline                          = $true                                          # Only required for MSol related discovery
$RequireConnectMicrosoftSharePoint                      = $true                                          # Only needed to connect to SharePoint Online Services
$PreloadOneDriveSites                                   = $true                                          # Used to preload OneDrive sites and not perform single lookups
$AdminSiteURL                                           = "https://cnainsurance-admin.sharepoint.com"    # Admin URL for Tenant
$PermissionsGroup                                       = "Domain Admins"                                # Permissions object to add to ACL permissions
$AdminPermissions                                       = "CNAW2K\Domain Admins"                         # Permissions to be check on directory when using get-acl
$IgnoreDirectories                                      = "Windows PowerShell|.cache"                    # Directories to be excluded from checking when running get-acl
$PreloadOneDriveSites                                   = $true                                          # Preload all OneDrive details and not one at a time
$DatabaseConnected                                      = $false                                         # Preset Database connection to false, reset on DB connection function





###### Functions ######
function DisplayExtendedInfo () {

    # Display to notify the operator before running
    Clear-Host
    Write-Host 
    Write-Host 
    Write-Host  '-------------------------------------------------------------------------------'	
	Write-Host  '                   OneDrive Migration Tool Kit                                 '   -ForegroundColor Green
	Write-Host  '-------------------------------------------------------------------------------'
    Write-Host  '                                                                               '
    Write-Host  '  This Tool kit is used to help identify and resolve issues for OneDrive       '   -ForegroundColor YELLOW
    Write-Host  '  Migrations.                                                                  '   -ForegroundColor YELLOW
    Write-Host  "                                                              version: $version"   -ForegroundColor YELLOW
    Write-Host  '-------------------------------------------------------------------------------'
    Write-Host 
}


# FUNCTION - Request to continue the operation
function ShowMenuSelectOperations {
    
    Write-host 
    Write-Host "(0) - Select Single User from Active Directory"
    Write-host "(1) - Get AD Users from CSV File"
    Write-host "(2) - Check a User for SharePoint Licence & Provisioned"
    Write-host "(3) - Check Home Drives for Permission issues"
    Write-host "(4) - Perform Full Pre-migration Checks"
    Write-host "(5) - Find PST Files on Home Drives"
    Write-host "(6) - Perform a User Home Drive permissions ADD to Domain Admins"
    Write-host 


    $input = read-host " *** Please select a number to perform the operation *** "

     switch ($input) `
    {
    '0' {
        # Select Single User from Active Directory
        CheckforPSliteDBModule
        ImportPSliteDBModule
        DatabaseConnection
        ImportActiveDirectoryModule
        SingleADUser

    }
    
    '1' {
        # Get Users from CSV File
        CheckforPSliteDBModule
        ImportPSliteDBModule
        DatabaseConnection
        ImportActiveDirectoryModule
         
    }

    '2' {
        # Check a User for SharePoint Licence & Provisioned
        CheckforPSliteDBModule
        ImportPSliteDBModule
        DatabaseConnection
        ImportActiveDirectoryModule
        SingleADUser
        CheckMSonlineModule
        SharePointModule
        AskForAdminCreds
        ConnectMicrosoftOnline
        ConnectMicrosoftSharePoint
        FindAdminLogonID
        CheckPermissionsOnline
        SharePointCheckLicence
        $PreloadOneDriveSites = $false
        GetOneDriveDetails
        

	}

    '3' {
        # Check Home Drives for Permission issues

	}

    '4' {
        # Perform Full Pre-migration Checks

	}

    '5' {
        # Find PST Files on Home Drives

	}

    '6' {
        # Perform a User Home Drive permissions ADD to Domain Admins

	}

    default {
        write-host 'You may only answer MigrateUser or AssignPolicies, please try again.'
        ShowMenu
    }   
    }
}



###
# FUNCTION - Request to continue the operation
function ShowMenuSelectUsers {
    
    Write-host 
    Write-Host "(0) - Select Single User from Active Directory"
    Write-Host "(1) - Load from CSV File (SamaccountNames)"
    Write-Host "(2) - Run custom Get-Aduser command"

    Write-host 


    $input = read-host " *** Please select a number to perform the operation *** "

     switch ($input) `
    {
    '0' {
        # Select Single User from Active Directory
        CheckforPSliteDBModule 
        ImportPSliteDBModule
        DatabaseConnection
        ImportActiveDirectoryModule
        SingleADUser
        ShowMenuSelectOperations

    }
    
    '1' {
        # Load from CSV File (SamaccountNames)

    }

    '2' {
        # Run custom Get-Aduser command
       

	}

    '3' {
        # Check Home Drives for Permission issues

	}

    

    default {
        write-host 'You may only answer MigrateUser or AssignPolicies, please try again.'
        ShowMenu
    }   
    }
}

###

# Run Order Functions
function FunctionRunOrderList () {

    DisplayExtendedInfo
    ShowMenu
    CheckforPSliteDBModule
    ImportPSliteDBModule
    DatabaseConnection
    CheckMSonlineModule
    SharePointModule
    ImportActiveDirectoryModule
    AskForAdminCreds
    FindAdminLogonID
    ConnectMicrosoftOnline
    ConnectMicrosoftSharePoint
    CheckPermissionsOnline
    CheckExcludeListFile
    CheckExcludeGroupFile
    ImportExcludeList
    ImportExclusionGroupsFile
    PreloadOneDriveSites
    GetServerShareSplit
    TestConnection
    TestPathHomeDrive
    GetHomeDirectorySize
    CheckAccessDirectories
    ProvisionOneDriveUser
    GetOneDriveDetails
    CheckForExcludedUser
    Takedown
    IcaclsPermissions
    SharePointCheckLicence

}


# FUNCTION - Setup PSlteDB Database
function DatabaseConnection () {

    if ($PSliteModuleImported -eq $true){

        # Test if database exists
        $TestDBExists = Test-Path $DBName

        # Checks if database exists and then creates if not found
        if ($TestDBExists){ 
            try {Open-LiteDBConnection $DBname -Mode shared | Out-Null ; $Global:DatabaseConnected = $true
            WriteTransactionsLogs -Task "Connected to Database $DBName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False}
            Catch {WriteTransactionsLogs -Task "Connection to database Failed" -Result Error -ErrorMessage "Connection Error:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
        }
        Else {Try {New-LiteDBDatabase -Path $DBname | Out-Null
            WriteTransactionsLogs -Task "Creating Database $DBname" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData False
            Open-LiteDBConnection $DBName -Mode shared | Out-Null ; $Global:DatabaseConnected = $true} 
            catch {WriteTransactionsLogs -Task "Failed to Create Database $DBname" -Result Information -ErrorMessage "Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
        }
    
        if ($DatabaseConnected -eq $true){
    
            # Create Collections in Database
            WriteTransactionsLogs -Task "Checking for Database Collections" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False
            New-LiteDBCollection Transactions -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 
            New-LiteDBCollection UserReports -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 

            
    
        }
    }

}


# FUNCTION - Check for NoSQL Database Module
function CheckforPSliteDBModule () {

    # Find is PSliteDB module is installed
    if (Get-Module -ListAvailable -Name PSLiteDB) {WriteTransactionsLogs -Task "Found PSliteDB Module via ListAvailable" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False
        $Global:PSliteDBModuleLocation = 'ListAvailable'}

    # Check of the PSliteDB module is located in the script directory
    Elseif (Test-Path .\PSliteDB\module\PSLiteDB.psd1) {WriteTransactionsLogs -Task "PSLiteDB Module found in script directory" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False
        $Global:PSliteDBModuleLocation = 'Directory'}

    Else {WriteTransactionsLogs -Task "No Database Module found, install the PSliteDB from 'https://github.com/v2kiran/PSLiteDB', application will now close" -Result Information -ErrorMessage "Missing Database Module" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False -ExportData False
    TerminateScript}
}


# FUNCTION -  Check MSol Module is installed
function CheckMSonlineModule () {

    WriteTransactionsLogs -Task "Checking Microsoft Online Module"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   
    if (Get-Module -ListAvailable -Name MSonline) {
        WriteTransactionsLogs -Task "Found Microsoft Online Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Microsoft Online Module, it needs to be installed" -Result Error -ErrorMessage "Online Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    TerminateScript	
    }
}


# FUNCTION -  Check SharePoint Module is installed
function SharePointModule () {

    WriteTransactionsLogs -Task "Checking Microsoft SharePoint Module"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   
    if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell ) {
        WriteTransactionsLogs -Task "Found Microsoft SharePoint Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Microsoft SharePoint Module, it needs to be installed" -Result Error -ErrorMessage "SharePoint Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    TerminateScript	
    }
}


# FUNCTION -  Check Active Directory Module is installed
function ImportActiveDirectoryModule () {

    WriteTransactionsLogs -Task "Checking Active Directory Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        WriteTransactionsLogs -Task "Found ActiveDirectory Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Import-Module ActiveDirectory 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Active Directory Module" -Result Error -ErrorMessage "Active Directory Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false 
        TerminateScript	
    }
    
}

# FUNCTION - Import PSlite Module 
function ImportPSliteDBModule () {

    # Import module from List Available
    if ($PSliteDBModuleLocation -eq "ListAvailable") {
        WriteTransactionsLogs -Task "Importing PSliteDB Module via ListAvailable" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData False
        try {Import-Module PSliteDB -ErrorAction Stop; $Global:PSliteModuleImported = $true}
        Catch {WriteTransactionsLogs -Task "Error loading module from ListAvailable" -Result Information -ErrorMessage "Import Failed" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False}
    }

    # import module from Directory
    if ($PSliteDBModuleLocation -eq "Directory") {
        WriteTransactionsLogs -Task "Importing PSliteDB Module via Directory" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData False
        try {Import-Module .\PSliteDB\module\PSLiteDB.psd1 -ErrorAction Stop ; $Global:PSliteModuleImported = $true}
        Catch {WriteTransactionsLogs -Task "Error Importing module from Directory" -Result Information -ErrorMessage "Import Failed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
    }

    if ($null -eq $PSliteModuleImported ){TerminateScript}
}


# FUNCTION - Terminate script
function TerminateScript () {

    Write-Host `n
    Write-Host 'This script has closed due to the error above' -ForegroundColor WHITE -backgroundcolor RED
    Write-Host `n
    Write-Host "Press any key to end..."

    # Create pause like wait...
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Exit
}

# FUNCTION - Get Credentions
function AskForAdminCreds () {

    # Asking for creds if they don't exist 
    if ($MFALoginRequired -eq $false) {
        WriteTransactionsLogs -Task "Asking for Service Account Credentials"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $Global:credentials = Get-Credential -Message "Enter Login details for Office 365"
    }
}

# FUNCTION - Connect to 365 Microsoft Online
function ConnectMicrosoftOnline () {

    # Check Connection to 365 or Connect if not already
    if ($RequireConnectMicrosoftOnline -eq $true) {

        try {
         try { Get-MsolCompanyInformation -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing Msol Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
         catch {
                WriteTransactionsLogs -Task "Not Connected to Msole" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                if ($MFALoginRequired -eq $True){Connect-MsolService  -ErrorAction Stop | Out-Null}
                if ($MFALoginRequired -eq $False){Connect-MsolService -Credential $Global:credentials  -ErrorAction Stop | Out-Null}
            }
        }  
        Catch {
            WriteTransactionsLogs -Task "Unable to Connect to Microsoft Online" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
	        TerminateScript
        }
    }
}


# FUNCTION - Connect to SharePoint Online
function ConnectMicrosoftSharePoint () {

    # Check Connection to SharePoint or Connect if not already
    if ($RequireConnectMicrosoftSharePoint -eq $true) {

        try {
         try { Get-SPOTenant -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing SharePoint Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
         catch {
                WriteTransactionsLogs -Task "Not Connected to SharePoint" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                if ($MFALoginRequired -eq $True){$Global:SharePointlogon = Connect-SPOService -Url $AdminSiteURL  -ErrorAction Stop | Out-Null}
                if ($MFALoginRequired -eq $False){Connect-SPOService -Url $AdminSiteURL -credential $credentials  -ErrorAction Stop | Out-Null}
            }
        }  
        Catch {
            WriteTransactionsLogs -Task "Unable to Connect to Microsoft SharePoint" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
	        TerminateScript
        }
    }
}

# FUNCTION -  Check User Permissions
function CheckPermissionsOnline () {

    $ValidPermissions = ''
    WriteTransactionsLogs -Task "Checking Online Permissions" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
   if ($ValidPermissions -like $null){try {$ValidPermissions = Get-MsolRoleMember -RoleObjectId 62e90394-69f5-4237-9190-012177145e10 | Where-Object {$_.emailaddress -eq $userAdminID}; WriteTransactionsLogs -Task "Found Admin in Global Administrators" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false} catch {WriteTransactionsLogs -Task "Permissions Error" -Result Information -ErrorMessage "Error happened searching Rbac Group" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false }}
   if ($ValidPermissions -like $null){try {$ValidPermissions = Get-MsolRoleMember -RoleObjectId f28a1f50-f6e7-4571-818b-6a12f2af6b6c | Where-Object {$_.emailaddress -eq $userAdminID}; WriteTransactionsLogs -Task "Found Admin in SharePoint Service Administrator" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false} catch {WriteTransactionsLogs -Task "Permissions Error" -Result Information -ErrorMessage "Error happened searching Rbac Group" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false}}   
   if ($ValidPermissions -like $null) {
       WriteTransactionsLogs -Task "Current user has no Permissions to perform the required actions" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    TerminateScript	}
}

# FUNCTION - Get Userlogon Identity
function FindAdminLogonID () {

    if ($MFALoginRequired -eq $false) {$Global:userAdminID = $credentials.username}
    if ($MFALoginRequired -eq $true) {$Global:userAdminID = $SharePointlogon.account}
}


# FUNCTION - Check Exclude User list
function CheckExcludeListFile () {

    WriteTransactionsLogs -Task "Checking For Exclude List File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (! (Test-Path $ExcludeListFile)) {
	    WriteTransactionsLogs -Task "Exclude List File Check" -Result Information -ErrorMessage "Exclude File Not found in expected location" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        $Global:ExcludeListFileNotFound = $false
    } else {
        WriteTransactionsLogs -Task "Exclude List File Check Located..........."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    }
}

# FUNCTION - Import Exclude User list
function ImportExcludeList () {

    if ($null -eq $ExcludeListFileNotFound){
       WriteTransactionsLogs -Task "Importing Exclude List File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        
        try {$ExcludeListUsers =  Get-content $ExcludeListFile
            $ExcludeListUsersCount = $ExcludeListUsers.count
            WriteTransactionsLogs -Task "Imported Exclude List File and has $ExcludeListUsersCount Users listed!"    -Result Warning -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        }
        Catch {WriteTransactionsLogs -Task "Imported Exclude List Failed, Job will Continue"    -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
            $ExcludeListFileNotFound = $false
        }
    }
}


# FUNCTION - Check Exclude group File
function CheckExcludeGroupFile () {

    WriteTransactionsLogs -Task "Checking For Exclude Group File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (! (Test-Path $ExcludeGroupFile)) {
	    WriteTransactionsLogs -Task "Exclude Group File Check" -Result Information -ErrorMessage "Exclude File Not found for groups in expected location" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        $Global:ExcludeGroupFileNotFound = $false
    } else {
        WriteTransactionsLogs -Task "Exclude Group File Check Located..........."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    }
}


# FUNCTION - Import Exclusion Groups
function ImportExclusionGroupsFile () {

    if ($null -eq $ExcludeGroupFileNotFound){
        WriteTransactionsLogs -Task "Importing Exclude Group File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
         
        try {$ExcludeGroups =  Get-content $ExcludeGroupFile
            $ExcludeGroupCount = $ExcludeGroups | Measure-Object | Select-Object -ExpandProperty Count
            WriteTransactionsLogs -Task "Imported Exclude Group File and has $ExcludeGroupCount Groups listed!"    -Result Warning -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        }
        Catch {WriteTransactionsLogs -Task "Imported Exclude Group List Failed, Job will Continue"    -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
            $ExcludeGroupFileNotFound = $false
        }
    }

}


# FUNCTION - PreLoad OneDrive sites for checking
function PreloadOneDriveSites () {
    
    If ($true -eq $PreloadOneDriveSites){

        try {WriteTransactionsLogs -Task "PreLoading OneDrive Sites...Please wait"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            $Global:OneDriveSites = Get-SPOSite -Template "SPSPERS" -Limit ALL -includepersonalsite $True -ea stop}
        Catch {WriteTransactionsLogs -Task "Failed to get OneDrive Sites"    -Result Error -ErrorMessage 'Error:' -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
               TerminateScript}
    }

}

# FUNCTION - WriteTransaction Log function    
function WriteTransactionsLogs  {

    #WriteTransactionsLogs -Task 'Creating folder' -Result information  -ScreenMessage true -ShowScreenMessage true exit #Writes to file and screen, basic display
          
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError true #Writes to file and screen and system "error[0]" is recorded
         
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError false  #Writes to file and screen but no system "error[0]" is recorded
         


    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$Task,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('Information','Warning','Error','Completed','Processing')]
        [string]$Result,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [string]$ErrorMessage,
    
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('True','False')]
        [string]$ShowScreenMessage,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [string]$ScreenMessageColour,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$IncludeSysError,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$ExportData
)
 
    process {
 
        # Stores Variables
        #$LogsFolder           = 'Logs'
 
        # Date
        $DateNow = Get-Date -f g    
        
        # Error Message
        $SysErrorMessage = $error[0].Exception.message
 
  
 
        $TransactionLogScreen = [pscustomobject][ordered]@{}
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Date"-Value $DateNow 
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Task" -Value $Task
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Error" -Value $ErrorMessage
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "SystemError" -Value $SysErrorMessage
        
       
        # Output to screen
       
        if  ($Result -match "Information|Warning" -and $ShowScreenMessage -eq "$true"){
 
        Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
        Write-host " | " -NoNewline
        Write-Host $TransactionLogScreen.Task  -NoNewline
        Write-host " | " -NoNewline
        Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour 
        }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$false"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage  -ForegroundColor $ScreenMessageColour
       }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$true"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage -NoNewline -ForegroundColor $ScreenMessageColour
       if (!$SysErrorMessage -eq $null) {Write-Host " | " -NoNewline}
       Write-Host $SysErrorMessage -ForegroundColor $ScreenMessageColour
       Write-Host
       }
   
        # Build PScustomObject
        $TransactionLogFile = [pscustomobject][ordered]@{}
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Date"-Value "$datenow"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Task"-Value "$task"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Result"-Value "$result"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Error"-Value "$ErrorMessage"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "SystemError"-Value "$SysErrorMessage"
 
        # Connect to Database
        if ($Global:DatabaseConnected -eq $true){Open-LiteDBConnection $DBName -Mode shared | Out-Null}

        # Export data if NOT specified
        if(!($ExportData)){$TransactionLogFile |  ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection Transactions}
        
        
 
 
        # Clear Error Messages
        $error.clear()
    }   
 
}

# FUNCTION - Feed the script AD users information from Get-ADuser or CSV file
function GetADUsersData () {

    # Pull users from AD via script / search 

    #$Users = Get-ADUser CAE1823 -Properties * | Where-Object {($_.enabled -eq "$true" -and $_.HomeDirectory -match "\\")} | Select-Object Displayname,Mail,HomeDirectory,Samaccountname,UserPrincipalName,AccountExpirationDate,accountExpires
    #$Users = Get-ADUser CAE0542 -Properties * | Where-Object {($_.enabled -eq "$true" -and $_.HomeDirectory -match "\\")} | Select-Object Displayname,Mail,HomeDirectory,Samaccountname,UserPrincipalName,AccountExpirationDate,accountExpires,Enabled
    #$Users = Get-ADUser cab1245 -Properties * | Where-Object {($_.enabled -eq "$true" -and $_.HomeDirectory -match "\\")} | Select-Object Displayname,Mail,HomeDirectory,Samaccountname,UserPrincipalName,AccountExpirationDate,accountExpires



    # Get Users via CSV File

    # Variables needed in script
    $HomeDrivePath # users homedrive path
    $SamaccountName
    $Mail 
}

# FUNCTION - Ask for individual User
function SingleADUser () {
    # Get a single AD user from AD via search
    $SingleUser = Read-Host -Prompt "Enter the SamaccountName of the User"
    Write-Host `n
    if ($SingleUser -eq "") {WriteTransactionsLogs -Task "No Username was entered" -Result Error -ErrorMessage "No ID Entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        Write-Host `n
        ShowMenu
    }

    Try {$Global:ADUserInfo = Get-ADuser -identity $SingleUser -properties * -EA Stop
        $ADUserInfoDisplayName = $ADUserInfo.DisplayName
        WriteTransactionsLogs -Task "Found AD User $ADUserInfoDisplayName" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False    
    }
    Catch {WriteTransactionsLogs -Task "User was not found" -Result Error -ErrorMessage "Not Found in AD" -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError True}



}



# FUNCTION Calculate bytes to KB/MB/GB....
Function Format-Bytes {
    Param
    (
        [Parameter(
            ValueFromPipeline = $true
        )]
        [ValidateNotNullOrEmpty()]
        [float]$number
    )
    Begin{
        $sizes = 'KB','MB','GB','TB','PB'
    }
    Process {
        # New for loop
        for($x = 0;$x -lt $sizes.count; $x++){
            if ($number -lt "1$($sizes[$x])"){
                if ($x -eq 0){
                    return "$number B"
                } else {
                    $num = $number / "1$($sizes[$x-1])"
                    $num = "{0:N2}" -f $num
                    return "$num $($sizes[$x-1])"
                }
            }
        }

    }
    End{}
}




######################## Admin functions #####################################
# FUNCTION - Get Server name and Share name from HomeDirectory Path
function GetServerShareSplit () {

    # Find Server name in Path
    [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
    
    #Find Share name in Path
    [string]$HomedriveShare = $HomeDrivePath -split '\\' | select-object -Skip 1 -Last 1
    
    # Write data to database 
    $Report | Add-Member -MemberType NoteProperty -Name HostingServer -Value $HomedriveServer -Force
    $Report | Add-Member -MemberType NoteProperty -Name HostingShare -Value $HomedriveShare -Force

}


# FUNCTION -  Test Connection to server using $homeDriveServer variable
function TestConnection () {
    
    # Test connection to server 
    Try {$TestConnection = Test-connection -Count 2 -ComputerName $HomedriveServer -ResolveDestination -EA Stop  
        WriteTransactionsLogs -Task "Successfully connected to $HomedriveServer for $SamaccountName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False
        $Report | Add-Member -MemberType NoteProperty -Name TestConnection -Value "Passed" -Force

        # Find Server domain name
        $Testconnection = $Testconnection | Select-Object -First 1
        $Dot = $Testconnection.Destination.IndexOf(".")
        $ServerDomain = $testconnection.Destination.Substring(1+$dot)

        $Report | Add-Member -MemberType NoteProperty -Name ServerDomain -Value "$ServerDomain" -Force 
    }
    
    Catch{WriteTransactionsLogs -Task "Failed to connect to $HomedriveServer for $SamaccountName " -Result ERROR -ErrorMessage "Test-Connection Failed:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        $Report | Add-Member -MemberType NoteProperty -Name TestConnection -Value "Failed" -Force    

    }
}

# FUNCTION - TestPath of Home Drive Directory
function TestPathHomeDrive () {

    # Test if HomeDirectory path is vaild
    Try {$Testpath = Test-path -Path $HomeDrivePath -EA Stop 
            
        If ($Testpath -eq $true){WriteTransactionsLogs -Task "Successful Testpath to $HomeDrivePath for $SamaccountName " -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False
            $Report | Add-Member -MemberType NoteProperty -Name TestPath -Value "Passed" -Force  
        }
        If ($Testpath -eq $false){WriteTransactionsLogs -Task "Failed Testpath to $HomeDrivePath for $SamaccountName " -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
            $Report | Add-Member -MemberType NoteProperty -Name TestPath -Value "Failed" -Force
        }
    }
        
    Catch {WriteTransactionsLogs -Task "Failed to Test path $HomeDrivePath for $SamaccountName " -Result ERROR -ErrorMessage "Test-Path Failed:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
    }

}

# FUNCTION - Get HomeDrive Size
function GetHomeDirectorySize () {

    WriteTransactionsLogs -Task "Calculating HomeDirectory Size.....Please Wait" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        
    if ($HomeDirectoryVisable -eq $True) { # Change later in script as var not set anywhere
            
        # Use Robocopy and get size
        $HomeDriveSize = (robocopy.exe $ADUserInfo.HomeDirectory c:\fakepathduh /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64)[-4] -replace '\D+(\d+).*','$1'
        $TotalSize = Format-Bytes $HomeDriveSize

        WriteTransactionsLogs -Task "$Displayname has the following stats TotalSize $TotalSize " -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectorySize -Value $TotalSize -Force
    }
    Else {
        WriteTransactionsLogs -Task "Failed to get the HomeDirectory details" -Result Error -ErrorMessage "No Access or other error" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectorySize -Value 'No Data' -Force
    }      
}


# FUNCTION - Check access using Get-ACL on directories to see access denid errors
function CheckAccessDirectories () {

    WriteTransactionsLogs -Task "Finding Directories for $SamaccountName on $HomeDrivePath... please wait" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    $DirList = robocopy $HomeDrivePath NULL /L /S /NJH /NJS /NFL /FP /NC /NS /XJ /R:0 /W:0
    $DirList = $DirList.Trim()
    $DirList = $DirList | Where-Object {$_ -ne ""}
    $DirListCount = $DirList | Measure-Object | Select-Object -ExpandProperty Count
    WriteTransactionsLogs -Task "Found $DirListCount Directories for $SamaccountName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    $Report | Add-Member -MemberType NoteProperty -Name DirectoriesFound -Value "$DirListCount" -Force

       
    # Update Status bar with progress
    $Processedirs = 0
    $ProcessFailed = 0
    WriteTransactionsLogs -Task "Performing ACL Check on Directories.. Please wait" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    Foreach ($Dir in $DirList){

    # Reset Catch trigger to filter duplicate checking
    $TriggedCatch = $false

        
    # First attemp to get Permissions
    Try {$FullPermissions = Get-Acl -ea stop $Dir ;$Permissions = $FullPermissions.access | Where-Object {$_.IdentityReference -match $PermissionsGroup} | Select-Object IdentityReference,FileSystemRights
        $Processedirs ++
    }
    Catch {$ErrorMessageDetails = $error[0].Exception.message
        #WriteTransactionsLogs -Task "Failed $SamaccountName on $Dir" -Result ERROR -ErrorMessage "Get-ACL Failed:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        $FailedDirs | Add-Member -MemberType NoteProperty -Name SamaccountName -Value "$SamaccountName" -Force
        $FailedDirs | Add-Member -MemberType NoteProperty -Name BadDirectory -Value "$Dir" -Force
        $FailedDirs | Add-Member -MemberType NoteProperty -Name ErrorDetails -Value "$ErrorMessageDetails" -Force
        $FailedDirs | Add-Member -MemberType NoteProperty -Name AdminFixTried -Value "$true" -Force
        $FailedDirs | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection FailedDirectories
        $ProcessFailed ++
        $TriggedCatch = $true
    }

    If ($TriggedCatch -eq $false){
        If ($FullPermissions.Path -match $IgnoreDirectories){}
            Else {
            
            If ($Permissions){ }
                Else {$FailedDirs | Add-Member -MemberType NoteProperty -Name Displayname -Value "$Displayname" -Force
                $FailedDirs | Add-Member -MemberType NoteProperty -Name SamaccountName -Value "$SamaccountName" -Force
                $FailedDirs | Add-Member -MemberType NoteProperty -Name BadDirectory -Value "$Dir" -Force
                $FailedDirs | Add-Member -MemberType NoteProperty -Name ErrorDetails -Value "Missing $PermissionsGroup Access" -Force
                $FailedDirs | Add-Member -MemberType NoteProperty -Name AdminFixTried -Value "$false" -Force
                $FailedDirs | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection FailedDirectories
                $ProcessFailed ++ 
               }
            }
        }
    }

}

# FUNCTION - SharePoint Licence Check via SKU
function SharePointCheckLicence () {

    WriteTransactionsLogs -Task "Checking for SharePoint SKU" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    if ($MsolUser.Licenses.ServiceStatus.ServicePlan.ServiceName -eq "SHAREPOINTENTERPRISE"){$SharePointLicFound = $true
       WriteTransactionsLogs -Task "User is assigned a SharePoint SKU" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
       $Report | Add-Member -MemberType NoteProperty -Name SharePointLicence -Value $SharePointLicFound -Force

   } Else {
       WriteTransactionsLogs -Task "Check for SharePoint SKU Failed" -Result information -ErrorMessage "User has not SharePoint SKU or Failed check" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
       $Report | Add-Member -MemberType NoteProperty -Name SharePointLicence -Value "Not Assigned or has error" -Force
   }
}


# FUNCTION - Provision User with OneDrive
function ProvisionOneDriveUser () {

    WriteTransactionsLogs -Task "Requesting OneDrive Provisioning" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
    try {Request-SPOPersonalSite -UserEmails $Mail -NoWait
        WriteTransactionsLogs -Task "Requesting Completed" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $ADUserInfo | Add-Member -MemberType NoteProperty -Name OneDriveProvisioned -Value $true -Force}
    Catch {WriteTransactionsLogs -Task "Request OneDrive Provisioning Failed" -Result Information -ErrorMessage "Error:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
    }
}


# FUNCTION - GetOneDrive Details from SharePoint Online
function GetOneDriveDetails () {

    WriteTransactionsLogs -Task "Checking OneDrive for Provisioned Container.... Please Wait" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        
    if ($PreloadOneDriveSites -eq $false) {$OneDriveDetails = Get-SPOSite -Template "SPSPERS" -Limit ALL -includepersonalsite $True -Filter "owner -eq $mail"}
    if ($PreloadOneDriveSites -eq $true) {$OneDriveDetails = $OneDriveSites | where-object {$_.owner -eq $mail}}
     
       
    If ($OneDriveDetails -like $null){WriteTransactionsLogs -Task "OneDrive not provisioned for $mail" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false

    } Else {
        WriteTransactionsLogs -Task "$mail is Provisioned for OneDrive" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $Report | Add-Member -MemberType NoteProperty -Name OneDriveProvisioned -Value $true -Force
        $OneDriveCurrentSize = $OneDriveDetails.StorageUsageCurrent
        $OneDriveURL = $OneDriveDetails.url
        $Report | Add-Member -MemberType NoteProperty -Name OneDriveCurrentSize -Value $OneDriveCurrentSize -Force
        $Report | Add-Member -MemberType NoteProperty -Name OneDriveURL -Value $OneDriveURL -Force
        WriteTransactionsLogs -Task "Current OneDrive Size is $OneDriveCurrentSize MB" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    }
}


# FUNCTION - Get Msol User Details
function GetMsolUser () {

    try {$MsolUser = Get-MsolUser -UserPrincipalName $Mail -ErrorAction Stop
        WriteTransactionsLogs -Task "Found $Mail in Azure " -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
     }
    Catch {WriteTransactionsLogs -Task "Failed to locate $Mail in Azure" -Result Error -ErrorMessage "$Mail was not found in Msol" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    }
}

# FUNCTION - Check for excluded user
function CheckForExcludedUser () {

    if ($null -eq $ExcludeListFileNotFound){
            
        # Compare exclude list with current user
        If ($ExcludeListUsers -contains "$UPN") {$UserValid = $false
            WriteTransactionsLogs -Task "$UPN found in exclude list and Remove from processing" -Result information -ShowScreenMessage true -ScreenMessageColour Yellow -IncludeSysError false
        }
    }
}

# FUNCTION - Reset Ownship on Folders and Files
function Takedown () {

    WriteTransactionsLogs -Task "Running Takeown on $HomeDrivePath for $UPN Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    takeown /F $HomeDrivePath\Data /R /A /D Y
    WriteTransactionsLogs -Task "Finished running Takeown on $HomeDrivePath for $UPN" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
}

# FUNCTION - Add icacls Domain Admins Permssions
function IcaclsPermissions () {

    WriteTransactionsLogs -Task "Running icacls on $HomeDrivePath Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    icacls $HomeDrivePath\Data /grant Domain Admins:(OI)(CI)F /T

    # Check if test folders exist for get-acl
    $Desktop        = Test-Path $HomeDrivePath\DATA\Desktop
    $Documents      = Test-Path $HomeDrivePath\DATA\Documents
    
    try {if ($Desktop -eq $true)
        {$FullPermissions = Get-Acl -ea stop "$HomeDrivePath\DATA\Desktop" ;$Permissions = $FullPermissions.access | Where-Object {$_.IdentityReference -match $PermissionsGroup} | Select-Object IdentityReference,FileSystemRights}
        Else {WriteTransactionsLogs -Task "No Access to $HomeDrivePath\DATA\Desktop for $UPN, Investigating Needed" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false}
    }
    Catch {$ErrorMessageDetails = $error[0].Exception.message
        WriteTransactionsLogs -Task "Failed Takeown on $HomeDrivePath on $UPN" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    }

    try {if ($Documents -eq $true)
        {$FullPermissions = Get-Acl -ea stop "$HomeDrivePath\DATA\Documents" ;$Permissions = $FullPermissions.access | Where-Object {$_.IdentityReference -match $PermissionsGroup} | Select-Object IdentityReference,FileSystemRights}
        Else {WriteTransactionsLogs -Task "No Access to $HomeDrivePath\DATA\Documents for $UPN, Investigating Needed" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false}
    }
    Catch {$ErrorMessageDetails = $error[0].Exception.message
        WriteTransactionsLogs -Task "Failed Takeown on $HomeDrivePath on $UPN" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    }

}





# Run main function to call others
DisplayExtendedInfo
ShowMenuSelectUsers

