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
 ##

####### Variable list #######
$Version                                                = "0.5"                                          # Version of script
Clear-Host                                                                                               # Clear screen
$host.ui.RawUI.WindowTitle                              = 'OneDrive Migration Tool Kit'                  # Set name in task bar
$DBName                                                 = "OneDriveDiscovery.db"                         # Database Name used to store and create if not existing
$ExcludeListFile                                        ='.\ExcludeUsers.txt'                            # List of Users to exclude by SAmaccountName
$ExcludeGroupFile                                       ='.\ExcludeGroups.txt'                           # List of AD Groups by DN to exclude if a users is a member of
$LicenceGroupsFile                                      ='.\LicenceGroups.txt'                           # List of AD Groups by DN to that contains E3/licences
$CSVDataFile                                            ='.\ADUsers.csv'                                 # File for AD users when importing via CSV - use SamaccountName only


$MFALoginRequired                                       = $true                                          # Is MFA required to login
$RequireConnectMicrosoftOnline                          = $true                                          # Only required for MSol related discovery
$RequireConnectMicrosoftSharePoint                      = $true                                          # Only needed to connect to SharePoint Online Services
$AdminSiteURL                                           = "https://cnainsurance-admin.sharepoint.com"    # Admin URL for Tenant
$PermissionsGroup                                       = "Domain Admins"                                # Permissions object to add to ACL permissions
$AdminPermissions                                       = "CNAW2K\Domain Admins"                         # Permissions to be check on directory when using get-acl
$IgnoreDirectories                                      = "Windows PowerShell|.cache"                    # Directories to be excluded from checking when running get-acl
$PreloadOneDriveSites                                   = $true                                          # Preload all OneDrive details and not one at a time
$DatabaseConnected                                      = $false                                         # Preset Database connection to false, reset on DB connection function
$ThrottleLimit                                          = '20'                                           # Throttle limit when using multi-threaded Jobs




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


###
# FUNCTION - Request to continue the operation
function ShowMenuSelectUsers {
    
    Write-host 
    Write-Host "(0) - Select Single User from Active Directory"
    Write-Host "(1) - Load from CSV File (SamaccountNames)"
    Write-Host "(2) - Run custom Get-Aduser command"
    Write-Host "(3) - Get OneDrive Users Information"
    Write-Host "(4) - Reports"

    Write-host 


    $input = read-host " *** Please select a number to perform the operation *** "

     switch ($input) `
    {
    '0' {
        # Select Single User from Active Directory
        CheckforPSliteDBModule 
        ImportPSliteDBModule
        DatabaseConnection
        #ImportActiveDirectoryModule
        SingleADUser
        ShowMenuSelectOperations

    }
    
    '1' {
        # Load from CSV File (SamaccountNames)
        CheckforPSliteDBModule 
        ImportPSliteDBModule
        DatabaseConnection
        #ImportActiveDirectoryModule
        ImportCSVData
        ShowMenuSelectOperations
        

    }

    '2' {
        # Run custom Get-Aduser command
        CheckforPSliteDBModule 
        ImportPSliteDBModule
        DatabaseConnection
        Write-Host `n
        Write-host "Enter Get-ADUser Search filter to find users" -ForegroundColor GREEN
        Write-Host `n
        Write-Host "Example: Displayname -eq ""Chris Healey""" -ForegroundColor DarkYellow
        Write-Host "Example: Country -ne ""UK""'"               -ForegroundColor DarkYellow
        Write-Host "Example: mail -like ""*@company.com"""     -ForegroundColor DarkYellow
        Write-Host "Example: mail -like ""*@company.com"" -and Department -eq ""Sales"""  -ForegroundColor DarkYellow
        Write-Host `n
        [string]$script:CustomADSearch = Read-Host -Prompt "Enter filter"

        if ($script:CustomADSearch -eq "") {WriteTransactionsLogs -Task "No filter was entered" -Result Error -ErrorMessage "No Filter" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData false
        ShowMenuSelectUsers}
        Else{

        ADFilterSearch}
    
	}

    '3' {
        # Get OneDrive Users Information
        CheckforPSliteDBModule
        ImportPSliteDBModule
        DatabaseConnection
        SharePointModule
        CheckMSonlineModule
        ConnectMicrosoftOnline
        AskForAdminCreds
        FindAdminLogonID
        ConnectMicrosoftSharePoint
        CheckPermissionsOnline
        PreloadOneDriveSites


	}

    '4' {
        # Open Report Menu
        CheckforPSliteDBModule
        ImportPSliteDBModule
        DatabaseConnection
        ShowReportMenu
	}

    

    default {
        write-host 'You may select one of the options'
        ShowMenuSelectUsers
    }   
    }
}

# FUNCTION - Request to continue the operation
function ShowMenuSelectOperations {
    
    Write-host '--------------------------------------------------------------------------------' -ForegroundColor DarkBlue
    Write-Host "(00) - PreMigration Check - PST/AppMember/Size/Item/Licence - SingleThread"
    Write-host "(01) - Check Home Drives for Permission Access issues - SingleThread"
    Write-host "(02) - Find Files on Home Drives - SingleThread"
    Write-host "(03) - Get All OneDrive Users - SingleThread - in Dev"
    Write-host "(04) - Perform a User Home Drive permissions ADD using Takeowns - SingleThread"
    Write-host "(05) - Find the size of a users HomeDrive - SingleThread"
    Write-host "(06) - Slow takeown and Icals - SingleThread"
    Write-host "(07) - Find the AppData Directory Size - SingleThread"
    Write-host "(08) - Make users HomeDrive Read Only (ACL) - SingleThread"
    Write-host "(09) - Create CSV Migration File for SharePoint Online"
    Write-host "(10) - Rename Users HomeDrive Share to Removed Access"


    Write-host 
    Write-host "(20) - Find Files on Home Drives - MultiThread"
    Write-host "(21) - Perform a User Home Drive permissions ADD using Takeowns - MultiThread"
    Write-host "(22) - Find the size of a users HomeDrive - MultiThread"
    Write-host "(23) - Check Home Drives for Permission Access issues - MultiThread"
    Write-host "(24) - Perform Slow Takeown & Icals - MultiThread"
    Write-host "(25) - PreMigration Check - PST/AppMember/Size/Item/Licence - MultiThread"  
    Write-host "(26) - Find the AppData Directory Size - MultiThread"  
    Write-host "(27) - Make users HomeDrive Read Only (ACL) - MultiThread"  
    Write-host '--------------------------------------------------------------------------------' -ForegroundColor DarkBlue

    Write-Host `n
    $input = read-host " *** Please select a number to perform the operation *** "
   
     switch ($input) `
    {
'00' {
         $Script:JobType = "PreMigration Check - PST/AppMember/Size/Item/Licence - SingleThread"
         Splitline
         BatchInformation
         CreateOperationCollector
         CheckLicenceGroupFile
         ImportLicenceGroupsFile
         AskForFileType
               
         Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            $Mail                    = $ADUser.Mail
            $UserDistinguishedName   = $ADUser.DistinguishedName

            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Functions 
            SingleThreadUpdate
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            CheckADGroupMembershipForLicence
            FindFiles
            CheckADGroupMembership           
            GetOneDriveDetails
            GetHomeDirectorySize
            SaveStoredResults
            Splitline
            }




    }
    
'01' {
        $Script:JobType = 'Check Home Drives for Permission Access issues - SingleThread'
         # Functions
         Splitline
         BatchInformation
         CreateOperationCollector
         Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Functions
            SingleThreadUpdate
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            Splitline 
            SaveStoredResults

            }

         
    }

'02' {
        $Script:JobType = 'Find Files on Home Drives - SingleThread'

            # Functions
            Splitline
            BatchInformation
            CreateOperationCollector
            AskForFileType
            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Functions
            SingleThreadUpdate
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            FindFiles
            SaveStoredResults
            Splitline 
            }
        

	}

'03' {
        $Script:JobType = "Get All OneDrive Users - SingleThread"
         Splitline


	}

'04' {
        # 
        $Script:JobType = 'Perform a User Home Drive permissions ADD using Takeowns - SingleThread'
         # Functions
         Splitline
         BatchInformation
         CreateOperationCollector
            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Functions 
            SingleThreadUpdate        
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            WriteTransactionsLogs -Task "Running Takeown on $HomeDrivePath for $UPN Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            TakeownTool
            WriteTransactionsLogs -Task "Finished running Takeown on $HomeDrivePath for $UPN" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            IcaclsPermissions
            SaveStoredResults
            }

	}

'05' {
        # Find the size of a users HomeDrive - SingleThread
         $Script:JobType = 'Find the size of a users HomeDrive - SingleThread'

         # Functions
         Splitline
         BatchInformation
         CreateOperationCollector
         Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
                        
            # Functions
            SingleThreadUpdate
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            GetHomeDirectorySize
            SaveStoredResults
            Splitline
            }


	}

'06' {
        # Slow takeown and Icals - SingleThread
         $Script:JobType = 'Perform Slow Takeown & Icals - SingleThread'

         # Functions
         Splitline
         BatchInformation
         CreateOperationCollector
         Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
                        
            # Functions
            SingleThreadUpdate
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            SlowTakeOwnIcals
            SaveStoredResults
            Splitline
        }

    

	}

'07' {
        $Script:JobType = 'Find the AppData Directory Size - SingleThread'

            # Functions
            Splitline
            BatchInformation
            CreateOperationCollector
            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Functions
            SingleThreadUpdate
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            GetHomeAppDataDirectorySize
            SaveStoredResults
            Splitline 
            
        }   

	}

'08' {
        $Script:JobType = 'Make users HomeDrive Read Only - SingleThread'

            # Functions
            Splitline
            BatchInformation
            CreateOperationCollector
            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Functions
            SingleThreadUpdate
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            MakeHomeDriveReadOnly
            SaveStoredResults
            Splitline 
            
        }   

	}

'09' {
        $Script:JobType = 'Create CSV Migration File for SharePoint Online"'

            # Functions
            Splitline
            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            $Mail                    = $ADUser.Mail
            
            # Functions
            SingleThreadUpdate
            ExportSharePointCSVList
            Splitline 
            
        }   

	}

'10' {
        $Script:JobType = 'Rename Users HomeDrive Share to Removed Access'

            # Functions
            Splitline
            BatchInformation
            CreateOperationCollector
            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Functions
            SingleThreadUpdate
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            RenameUserHomeFolder
            SaveStoredResults
            Splitline 
            
        }   

	}




'20' {
        # 
         $Script:JobType = 'Find Files on Home Drives - MultiThread'
         # Functions
         Splitline
         BatchInformation
         AskForFileType

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
            
            # Run Multithread call
            Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            $script:FileType = $using:FileType
            $Script:JobType = $using:JobType
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            FindFiles
            SaveStoredResults
            } | Out-Null

            } # Close foreach
            MultithreadGetEvents
            WriteTransactionsLogs -Task "Finished running PST Search" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            
           
        
        }

'21' {
        # 
         $Script:JobType = 'Perform a User Home Drive permissions ADD using Takeowns - MultiThread'
         # Functions
         Splitline
         BatchInformation

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1

            
            # Run Multithread call
            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            $PermissionsGroup = $using:PermissionsGroup
            $Script:JobType = $using:JobType
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck 
            WriteTransactionsLogs -Task "Running Takeown and Icacl on $HomeDrivePath for $UPN Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            TakeownTool
            IcaclsPermissions
            SaveStoredResults
            } | Out-Null


            } # Close foreach
            MultithreadGetEvents
            WriteTransactionsLogs -Task "Finished $Script:JobType" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   


	}

'22' {
        # Find the size of a users HomeDrive - MultiThread
        $Script:JobType = 'Find the size of a users HomeDrive - MultiThread'
         # Functions
         Splitline
         BatchInformation

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1

            
            # Run Multithread call
            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $Script:JobType = $using:JobType
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            GetHomeDirectorySize
            SaveStoredResults
            } | Out-Null


            } # Close foreach
            MultithreadGetEvents
            WriteTransactionsLogs -Task "Finished getting Directory homedrive sizes" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            


	}

'23' {
        #
        $Script:JobType = 'Check Home Drives for Permission Access issues - MultiThread'
         # Functions
         Splitline
         BatchInformation

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1

            
            # Run Multithread call
            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $Script:JobType = $using:JobType
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            $PermissionsGroup = $using:PermissionsGroup
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            SaveStoredResults
            } | Out-Null


            } # Close foreach
            MultithreadGetEvents
            WriteTransactionsLogs -Task "Finished $Script:JobType" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            


	}

'24' {
        # 
         $Script:JobType = 'Perform Slow Takeown & Icals - MultiThread'
         # Functions
         Splitline
         BatchInformation

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1

            
            # Run Multithread call
            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            $PermissionsGroup = $using:PermissionsGroup
            $Script:JobType = $using:JobType
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            SlowTakeOwnIcals
            #IcaclsPermissionsFix
            SaveStoredResults
            } | Out-Null


            } # Close foreach
            MultithreadGetEvents
            WriteTransactionsLogs -Task "Finished $Script:JobType" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   


	}

'25' {
        # 
         $Script:JobType = 'PreMigration Check - PST/AppMember/Size/Item/Licence - MultiThread'
         # Functions
         Splitline
         BatchInformation
         AskForFileType

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            $Mail                    = $ADUser.Mail
            $UserDistinguishedName   = $ADUser.DistinguishedName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1

            
            # Run Multithread call
            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $mail = $using:mail
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            $Script:JobType = $using:JobType
            $UserDistinguishedName = $using:UserDistinguishedName
            $ExcludeGroupFile = $using:ExcludeGroupFile
            $LicenceGroupsFile = $using:LicenceGroupsFile
            $script:FileType = $using:FileType

            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CheckLicenceGroupFile
            ImportLicenceGroupsFile
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            CheckADGroupMembershipForLicence
            FindFiles
            CheckExcludeGroupFile
            ImportExclusionGroupsFile
            CheckADGroupMembership           
            GetOneDriveDetails
            GetHomeDirectorySize
            SaveStoredResults


            } | Out-Null


            } # Close foreach
            MultithreadGetEvents
            WriteTransactionsLogs -Task "Finished $Script:JobType" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   


	}

'26' {
     # Find the size of a users AppData HomeDrive - MultiThread
        $Script:JobType = 'Find the size of a users AppData HomeDrive - MultiThread'
         # Functions
         Splitline
         BatchInformation

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1

            
            # Run Multithread call
            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $Script:JobType = $using:JobType
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            GetHomeAppDataDirectorySize
            SaveStoredResults
            } | Out-Null


            } # Close foreach
            MultithreadGetEvents       


	}

'27' {
     # Make users HomeDrive Read Only - MultiThread
        $Script:JobType = 'Make users HomeDrive Read Only - MultiThread'
         # Functions
         Splitline
         BatchInformation

            Foreach ($ADUser in $ADUserInfo){
            
            # Build Strings
            $HomeDrivePath           = $ADUser.HomeDirectory
            $UPN                     = $ADUser.UserPrincipalName
            $SamaccountName          = $ADUser.SamaccountName
            [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1

            
            # Run Multithread call
            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit  -InitializationScript $export_functions -ScriptBlock {
            $Script:JobType = $using:JobType
            $nextbatch = $using:nextbatch
            $HomeDrivePath = $using:HomeDrivePath
            $UPN = $using:UPN
            $SamaccountName = $Using:SamaccountName
            $HomedriveServer = $Using:HomeDriveServer
            $DBName = $using:DBName
            CheckforPSliteDBModule
            ImportPSliteDBModule
            DatabaseConnection
            CreateOperationCollector
            GetServerShareSplit
            TestConnection
            TestPathHomeDrive
            SimplePermissionsCheck
            MakeHomeDriveReadOnly
            SaveStoredResults
            } | Out-Null


            } # Close foreach
            MultithreadGetEvents       
	}

    default {
        write-host 'You may only answer MigrateUser or AssignPolicies, please try again.'
        ShowMenuSelectOperations
    }   
    }
}

function ShowReportMenu {
    
    Write-host 
    Write-Host "(0) - Show All Batch Information"
    Write-Host "(1) - Export Batch information By ID"
    Write-Host "(2) - Export OneDrive Information"
    Write-Host "(3) - Export File Search Results by Batch ID"
    Write-Host "(4) - Search Transactions by SamaccountName"

    Write-host 


    $input = read-host " *** Please select a number to perform the operation *** "

     switch ($input) `
    {
    '0' {
        # Show All Batch Information
        Splitline
        Open-LiteDBConnection $DBName -Mode shared | Out-Null
        Find-LiteDBDocument -Collection 'BatchInformation' | Select-Object _ID,JobType,DateJobCreation,UserCount | Sort-Object DateJobCreation | Format-Table -AutoSize
        Splitline
        ShowReportMenu

    }
    
    '1' {
        # Export Batch information By ID
        AskForBatchID
        ExportOperationsByBatchID
        AskToExportReport
        ShowReportMenu
       

    }

    '2' {
        # Export OneDrive Information
        Splitline
        Open-LiteDBConnection $DBName -Mode shared | Out-Null
        $ReportData = Find-LiteDBDocument -Collection 'OneDriveDetails' -limit 999999 
        Splitline
        $SaveMyFile = Get-Filename
        $ReportData | Export-Csv -Path $SaveMyFile
        ShowReportMenu

    
	}

    '3' {
        # Export File Search Results by Batch ID
        AskForBatchID
        ExportFileSearchByBatchID
        AskToExportReport       
        ShowReportMenu


	}

    '4' {
        # Search Transactions by SamaccountName
        AskForSamaccountID
        SearchTransationBySamaccountName
        AskToExportReport
        ShowReportMenu


	}   

    default {
        write-host 'You may select one of the options'
        ShowReportMenu
    }   
    }
}

# FUNCTION - Update Status Bar with Progress
function SingleThreadUpdate () {

    #Count objects in ADUserInfo Array
    $ADUserInfoCount = $ADUserInfo | Measure-Object | Select-Object -ExpandProperty Count
    
    # Create Counter var if not existing and add 1
    if (!($script:SingleADUserProcessedCount)){$script:SingleADUserProcessedCount = 1}
    Else {$script:SingleADUserProcessedCount ++}
    

    # Update Status bar
    $host.ui.RawUI.WindowTitle = "$Script:JobType | $script:SingleADUserProcessedCount out of $ADUserInfoCount  | BatchID $nextbatch"  
}



# FUNCTION - Check the status of multithreaded operations
function MultithreadGetEvents () {

Start-Sleep -Seconds 5
$Check = $true
$JobStatus = Get-Job
$TotalJobs = $JobStatus | where-object {$_.State -eq "NotStarted"} | Measure-Object | select -ExpandProperty count
 While($Check -eq $true){

            $JobStatus = Get-Job

            # Build running counts
            $RunningJobs = $JobStatus | where-object {$_.State -eq "Running"} | Measure-Object | select -ExpandProperty count
            $NotStartedJobs = $JobStatus | where-object {$_.State -eq "NotStarted"} | Measure-Object | select -ExpandProperty count
            $CompletedJobs = $JobStatus | where-object {$_.State -eq "Completed"} | Measure-Object | select -ExpandProperty count


            Get-Job | Receive-Job

            # Update Status bar with latest results
            Start-Sleep 1
            $host.ui.RawUI.WindowTitle = "Running Jobs: $RunningJobs | Not Started: $NotStartedJobs | Completed Jobs: $CompletedJobs | BatchID $nextbatch"
    
            # Check if still running and close jobs when last one completes
            #$RunningJobs = $JobStatus | where-object {$_.State -eq "Running"} | Measure-Object | select -ExpandProperty count
            #if ($CompletedJobs -eq $TotalJobs -and $RunningJobs -eq '0'){
            if ($RunningJobs -eq '0'){
                $host.ui.RawUI.WindowTitle = "Running Jobs: $RunningJobs | Not Started: $NotStartedJobs | Completed Jobs: $CompletedJobs  | BatchID $nextbatch - COMPLETED"
                Write-Host ''
                Start-Sleep -Seconds 2
                Get-Job | Remove-Job
                $Check = $false
                


            }
    }


}


# Run Order Functions
function FunctionList {

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
    Takeown
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
            try {Open-LiteDBConnection $DBname -Mode shared | Out-Null ; $script:DatabaseConnected = $true
            WriteTransactionsLogs -Task "Connected to Database $DBName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False}
            Catch {WriteTransactionsLogs -Task "Connection to database Failed" -Result Error -ErrorMessage "Connection Error:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
        }
        Else {Try {New-LiteDBDatabase -Path $DBname | Out-Null
            WriteTransactionsLogs -Task "Creating Database $DBname" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData False
            Open-LiteDBConnection $DBName -Mode shared | Out-Null ; $script:DatabaseConnected = $true} 
            catch {WriteTransactionsLogs -Task "Failed to Create Database $DBname" -Result Information -ErrorMessage "Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
        }
    
        if ($script:DatabaseConnected -eq $true){
    
            # Create Collections in Database
            WriteTransactionsLogs -Task "Checking for Database Collections" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False
            New-LiteDBCollection Transactions -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 
            New-LiteDBCollection BatchInformation -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            New-LiteDBCollection FileSessionDetails -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            New-LiteDBCollection Operations -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            New-LiteDBCollection OneDriveDetails -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

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
    WriteTransactionsLogs -Task "Importing Microsoft Online Module"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   
    Try {Import-Module Msonline -UseWindowsPowerShell -ea Stop -WarningAction SilentlyContinue
    }
    Catch {WriteTransactionsLogs -Task "Failed Importing Microsoft Online Module, it needs to be installed" -Result Error -ErrorMessage "Online Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        TerminateScript}
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
        Catch {WriteTransactionsLogs -Task "Error loading module from ListAvailable" -Result Information -ErrorMessage "Import Failed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False -ExportData False}
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
    #$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
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

    # Import SharePoint Module
    Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -WarningAction SilentlyContinue

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
   if ($ValidPermissions -like $null){try {$ValidPermissions = Get-MsolRoleMember -RoleObjectId 62e90394-69f5-4237-9190-012177145e10 | Where-Object {$_.emailaddress -eq $userAdminID}; WriteTransactionsLogs -Task "Found Admin in Global Administrators" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false} catch {WriteTransactionsLogs -Task "Permissions Error" -Result ERROR -ErrorMessage "Error happened searching Rbac Group - NOT A GLOBAL ADMIN" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false }}
   if ($ValidPermissions -like $null){try {$ValidPermissions = Get-MsolRoleMember -RoleObjectId f28a1f50-f6e7-4571-818b-6a12f2af6b6c | Where-Object {$_.emailaddress -eq $userAdminID}; WriteTransactionsLogs -Task "Found Admin in SharePoint Service Administrator" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false} catch {WriteTransactionsLogs -Task "Permissions Error" -Result ERROR -ErrorMessage "Error happened searching Rbac Group - NOT A SHAREPOINT ADMIN" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false}}   
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

    if (Test-Path $ExcludeGroupFile) {
	    WriteTransactionsLogs -Task "Exclude Group File Check Located..........."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $script:ExcludeGroupFileFound = $true
    } else {
        WriteTransactionsLogs -Task "Exclude Group File Check" -Result Information -ErrorMessage "Exclude File Not found for groups in expected location" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false 
    }
}


# FUNCTION - Check Licence group File exists
function CheckLicenceGroupFile () {

    if (Test-Path $LicenceGroupsFile) {
	    WriteTransactionsLogs -Task "Licence Group File Check Located..........."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $script:LicenceGroupsFileFound = $true
    } else {
        WriteTransactionsLogs -Task "Licence Group File Check" -Result Information -ErrorMessage "Exclude File Not found for groups in expected location" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false 
    }
}



# FUNCTION - Import Exclusion Groups
function ImportExclusionGroupsFile () {

    if ($true -eq $script:ExcludeGroupFileFound){
                 
        try {$script:ExcludeGroups =  Get-content $ExcludeGroupFile
            $ExcludeGroupCount = $ExcludeGroups | Measure-Object | Select-Object -ExpandProperty Count
            WriteTransactionsLogs -Task "Imported Exclude Group File and has $ExcludeGroupCount Groups listed!"    -Result Warning -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        }
        Catch {WriteTransactionsLogs -Task "Imported Exclude Group List Failed, Job will Continue"    -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        }
    }

}


# FUNCTION - Import Licence Groups File
function ImportLicenceGroupsFile () {

    if ($true -eq $script:LicenceGroupsFileFound){
                 
        try {$script:LicenceGroups =  Get-content $LicenceGroupsFile
            $LicenceGroupsFileCount = $script:LicenceGroups | Measure-Object | Select-Object -ExpandProperty Count
            WriteTransactionsLogs -Task "Imported Licence Group File and has $LicenceGroupsFileCount Groups listed!"    -Result Warning -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
            $script:LicenceGroupFileLoaded = $true
        }
        Catch {WriteTransactionsLogs -Task "Imported Licence Group List Failed, Job will Continue"    -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        }
    }

}



# FUNCTION - PreLoad OneDrive sites for checking
function PreloadOneDriveSites () {
    
     try {WriteTransactionsLogs -Task "PreLoading OneDrive Sites...Please wait"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
          $Global:OneDriveSites = Get-SPOSite -Template "SPSPERS" -Limit ALL -includepersonalsite $True -ea stop | Select Owner,StorageUsageCurrent,URL
          
          WriteTransactionsLogs -Task "Recording OneDrive Sites into Database...Please wait"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
          # Load data into Database
          Open-LiteDBConnection $DBName -Mode shared | Out-Null

          # Clean up existing data if required
          Remove-LiteDBCollection -Collection OneDriveDetails -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -confirm:$false
          # Create new collection
          New-LiteDBCollection -Collection OneDriveDetails 

          
          $OneDriveSitesArray = $Global:OneDriveSites | ConvertTo-LiteDbBSON -as array 
          Add-LiteDBDocument 'OneDriveDetails' -BsonDocumentArray $OneDriveSitesArray -BatchSize 1000 -BulkInsert
          
          Close-liteDBConnection
          WriteTransactionsLogs -Task "Completed PreLoading OneDrive Sites"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false


          }
      Catch {WriteTransactionsLogs -Task "Failed to get OneDrive Sites"    -Result Error -ErrorMessage 'Error:' -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
             TerminateScript}
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
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "SamAccountName" -Value $samaccountname
        
       
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
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "SamaccountName"-Value "$SamaccountName"
 
        # Connect to Database
        if ($script:DatabaseConnected -eq $true){Open-LiteDBConnection $DBName -Mode shared | Out-Null}

        # Export data if NOT specified
        if(!($ExportData)){$TransactionLogFile |  ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection Transactions}
        
        
 
 
        # Clear Error Messages
        $error.clear()
    }   
 
}



# FUNCTION - Get CSV Data from File
function ImportCSVData () {

    WriteTransactionsLogs -Task "Importing Data file................$CSVDataFile"   -Result Information none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    try {$Global:OneDriveUsers = Import-Csv ".\$CSVDataFile" -Delimiter "," -ea stop
        WriteTransactionsLogs -Task "Loaded Users Data"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError true
    } 
    catch {WriteTransactionsLogs -Task "Error loading Users data File" -Result Error -ErrorMessage "An error happened importing the data file, Please Check File" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
          TerminateScript
    }

    if ($Global:OneDriveUsers){
        
            if ($ADUserDetails){Remove-Variable -Scope Global -Name ADuserInfo}
            
            # Build PScustomObject
            $Global:ADuserInfo = @()
            Foreach ($User in $OneDriveUsers) {

                # Build String
                $SamAccountName = $User.SAmaccountName
            
                Try {$Global:ADuserInfo += Get-ADuser -identity $SamAccountName -Properties Displayname,Mail,HomeDirectory -EA Stop
                
                 }
                Catch {WriteTransactionsLogs -Task "NOT Found User $SamAccountName" -Result Information -ErrorMessage "User Not Found" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
                $ADUsersNotFound ++
            }
        }

    }
    if (!($ADUsersNotFound)){$ADUsersNotFoundCount = 0}
    $ADUsersNotFoundCount = $ADUsersNotFound 
    $ADUserInfoCount = $Global:ADUserInfo | Measure-Object | Select-Object -ExpandProperty Count
    WriteTransactionsLogs -Task "Active Directory Users Found $ADUserInfoCount from CSV, Not Found $ADUsersNotFoundCount" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False
}



# FUNCTION - Ask for individual User
function SingleADUser () {
    # Get a single AD user from AD via search
    $SingleUser = Read-Host -Prompt "Enter the SamaccountName of the User"
    Write-Host `n
    if ($SingleUser -eq "") {WriteTransactionsLogs -Task "No Username was entered" -Result Error -ErrorMessage "No ID Entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        Write-Host `n
    }

    Try {$Global:ADUserInfo = Get-ADuser -identity $SingleUser -properties * -EA Stop
        $ADUserInfoDisplayName = $ADUserInfo.DisplayName
        WriteTransactionsLogs -Task "Found AD User $ADUserInfoDisplayName" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False    
    }
    Catch {WriteTransactionsLogs -Task "User was not found" -Result Error -ErrorMessage "Not Found in AD" -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError True
        SingleADUser}



}


# FUNCTION - Customer AD Filter Search
function ADFilterSearch () {

    WriteTransactionsLogs -Task "Running customer AD Filter Search..... Please Wait" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False

    try{$ADUserInfo = Get-ADuser -filter $script:CustomADSearch -Properties * -ea stop}
    catch {WriteTransactionsLogs -Task "There was an error with the search filter, please check." -Result ERROR -ErrorMessage "ERROR : " -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError true
        pause
        ShowMenuSelectUsers}

        # Calculate how many users are found and display results
        $ADUserInfoCount = $ADUserInfo.samaccountname | Measure-Object | Select-Object -ExpandProperty count

        WriteTransactionsLogs -Task "AD user(s) found $ADUserInfoCount from custom filter. Displaying first 5 results:" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $ADUserInfo | Select SamaccountName,Displayname,mail | Out-GridView
        Write-host 
        Write-host 
        ShowMenuSelectOperations
}



# FUNCTION Calculate bytes to KB/MB/GB....
Function FormatBytes {
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


# FUNCTION - Ask for file type for search
function AskForFileType () {

    $script:FileType = Read-Host -Prompt "Enter File Type to Search for: e.g. *.pst"
    Write-Host `n
    if ($script:FileType -eq "") {WriteTransactionsLogs -Task "File type was entered" -Result Error -ErrorMessage "No file type entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        Write-Host `n
        AskForFileType 
    }
   
}



######################## Admin functions #####################################
# FUNCTION - Get Server name and Share name from HomeDirectory Path
function GetServerShareSplit () {

       
    # Find Server name in Path
    [string]$HomedriveServer = $HomeDrivePath -split '\\' | select-object -Skip 2 -Last 1
    
    #Find Share name in Path
    [string]$HomedriveShare = $HomeDrivePath -split '\\' | select-object -Skip 1 -Last 1

    
    # Write data to database 
    $script:Operations | Add-Member -MemberType NoteProperty -Name HostingServer -Value $HomedriveServer -Force
    $script:Operations | Add-Member -MemberType NoteProperty -Name HostingShare -Value $HomedriveShare -Force
    


}


# FUNCTION -  Test Connection to server using $homeDriveServer variable
function TestConnection () {    
         
    # Test connection to server 
    Try {$TestConnection = Test-connection -Count 2 -ComputerName $HomedriveServer -ResolveDestination -EA Stop  
        WriteTransactionsLogs -Task "Successfully connected to $HomedriveServer for $SamaccountName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False
        $script:Operations | Add-Member -MemberType NoteProperty -Name TestConnection -Value "Passed" -Force

        # Find Server domain name
        $Testconnection = $Testconnection | Select-Object -First 1
        $Dot = $Testconnection.Destination.IndexOf(".")
        $ServerDomain = $testconnection.Destination.Substring(1+$dot)

        $script:Operations | Add-Member -MemberType NoteProperty -Name ServerDomain -Value "$ServerDomain" -Force 
    }
    
    Catch{WriteTransactionsLogs -Task "Failed to connect to $HomedriveServer for $SamaccountName " -Result ERROR -ErrorMessage "Test-Connection Failed:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        $script:Operations | Add-Member -MemberType NoteProperty -Name TestConnection -Value "Failed" -Force    

    }
    
}

# FUNCTION - TestPath of Home Drive Directory
function TestPathHomeDrive () {

    # Test if HomeDirectory path is vaild
    Try {$Testpath = Test-path -Path $HomeDrivePath -EA Stop 
            
        If ($Testpath -eq $true){WriteTransactionsLogs -Task "Successful Testpath to $HomeDrivePath for $SamaccountName " -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False
            $script:HomeDirectoryValid = $True
            $script:Operations| Add-Member -MemberType NoteProperty -Name TestPath -Value "Passed" -Force  
        }
        If ($Testpath -eq $false){WriteTransactionsLogs -Task "Failed Testpath to $HomeDrivePath for $SamaccountName " -Result Error -ErrorMessage "Possible drive not existing" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
            $script:Operations | Add-Member -MemberType NoteProperty -Name TestPath -Value "Failed" -Force
            $script:HomeDirectoryValid = $false
        }
    }
        
    Catch {WriteTransactionsLogs -Task "Failed to Test path $HomeDrivePath for $SamaccountName " -Result ERROR -ErrorMessage "Test-Path Failed:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
    }
    

}

# FUNCTION - Get HomeDrive Size
function GetHomeDirectorySize () {

    if ($Global:AllowDriveAccess -eq $true){

    WriteTransactionsLogs -Task "Calculating HomeDirectory Size.....Please Wait" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        
    if ($HomeDirectoryValid -eq $True) { # Change later in script as var not set anywhere
         
        # Use Robocopy and get size
        $script:HomeDriveDetails = (robocopy.exe $HomeDrivePath c:\fakepathduh /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64)
           
        # Files
        $HomeDriveFileCount = $script:HomeDriveDetails[-5] -replace '\D+(\d+).*','$1'
        # Directories
        $HomeDriveDirCount = $script:HomeDriveDetails[-6] -replace '\D+(\d+).*','$1'
        # Size
        $HomeDriveSize = $script:HomeDriveDetails[-4] -replace '\D+(\d+).*','$1'
        # Create total size and use function to count bytes
        $TotalSize = FormatBytes $HomeDriveSize
              
    }
    Else {
        WriteTransactionsLogs -Task "Failed to get the HomeDirectory details" -Result Error -ErrorMessage "No Access or other error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
    } 
    
    if ($HomeDriveSize -ge '0'){WriteTransactionsLogs -Task "$UPN HomeDrive Size: $TotalSize | Directories:$HomeDriveDirCount| Files:$HomeDriveFileCount" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeDirectoryFileCount"-Value $HomeDriveFileCount
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeDirectoryDirsCount"-Value $HomeDriveDirCount
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeDirectorySize"-Value $TotalSize
    } 
        
   }
   Else {WriteTransactionsLogs -Task "Skipped Drive Size Check for $SamaccountName on $HomeDrivePath as failed simple drive access check" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeDirectoryInfo"-Value 'Failed simple drive access check'
    }

}


# FUNCTION - Get HomeDrive Size
function GetHomeAppDataDirectorySize () {

    if ($Global:AllowDriveAccess -eq $true){

    WriteTransactionsLogs -Task "Calculating AppData Size.....Please Wait" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        
    if ($HomeDirectoryValid -eq $True) { # Change later in script as var not set anywhere
         
        # Use Robocopy and get size
        $script:HomeDriveDetails = (robocopy.exe $HomeDrivePath\Data\AppData c:\fakepathduh /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64)
           
        # Files
        $HomeDriveFileCount = $script:HomeDriveDetails[-5] -replace '\D+(\d+).*','$1'
        # Directories
        $HomeDriveDirCount = $script:HomeDriveDetails[-6] -replace '\D+(\d+).*','$1'
        # Size
        $HomeDriveSize = $script:HomeDriveDetails[-4] -replace '\D+(\d+).*','$1'
        # Create total size and use function to count bytes
        $TotalSize = FormatBytes $HomeDriveSize
              
    }
    Else {
        WriteTransactionsLogs -Task "Failed to get the AppData HomeDirectory details" -Result Error -ErrorMessage "No Access or other error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
    } 
    
    if ($HomeDriveSize -ge '0'){WriteTransactionsLogs -Task "$UPN AppData HomeDrive Size: $TotalSize | Directories:$HomeDriveDirCount| Files:$HomeDriveFileCount" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeAppDataDirectoryFileCount"-Value $HomeDriveFileCount
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeAppDataDirectoryDirsCount"-Value $HomeDriveDirCount
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeAppDataDirectorySize"-Value $TotalSize
    } 
        
   }
   Else {WriteTransactionsLogs -Task "Skipped Drive Size Check for $SamaccountName on $HomeDrivePath as failed simple drive access check" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $script:Operations | Add-Member -MemberType NoteProperty -Name "HomeDirectoryInfo"-Value 'Failed simple drive access check'
    }

}



# FUNCTION - Check if a user is part of a group
function CheckADGroupMembership () {

    if ($true -eq $script:ExcludeGroupFileFound){

        # Create GroupContainer
        $GroupmemberShip = ''

        Foreach ($group in $script:ExcludeGroups){

            # Get Group Displayname
            try {$GroupDetails = Get-ADobject $Group -properties member -ea stop

                $Groupname = $GroupDetails.Name
            }
            Catch{WriteTransactionsLogs -Task "Unable to find $group" -Result ERROR -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true -ExportData false
            }

            # Find User in group if member
            if ($GroupDetails | Where-Object {$_.member -contains $UserDistinguishedName}){
                WriteTransactionsLogs -Task "Found User $SamaccountName in $groupName" -Result information -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
            
                # Build groups user is a member of
                $GroupmemberShip += $Groupname+";"
            }
            Else {WriteTransactionsLogs -Task "$SamaccountName Not a member of $groupName" -Result information -ShowScreenMessage true -ScreenMessageColour Green -IncludeSysError false
            }
        # Record values 
        $script:Operations | Add-Member -MemberType NoteProperty -Name GroupMemberShip -Value $GroupmemberShip -Force
        }
    }   
}

# FUNCTION - Check if a user is part of the licence group
function CheckADGroupMembershipForLicence () {

    if ($true -eq $script:LicenceGroupFileLoaded){

        # Create GroupContainer
        $GroupmemberShip = ''

        Foreach ($group in $script:LicenceGroups){

            # Get Group Displayname
            try {$GroupDetails = Get-ADobject $Group -properties member -ea stop

                $Groupname = $GroupDetails.Name
            }
            Catch{WriteTransactionsLogs -Task "Unable to find $group" -Result ERROR -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true -ExportData false
            }

            # Find User in group if member
            if ($GroupDetails | Where-Object {$_.member -contains $UserDistinguishedName}){
                WriteTransactionsLogs -Task "Found User $SamaccountName in $groupName" -Result information -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
            
                # Build groups user is a member of
                $GroupmemberShip += $Groupname+";"
            }
            Else {WriteTransactionsLogs -Task "$SamaccountName Not a member of $groupName" -Result information -ShowScreenMessage true -ScreenMessageColour Green -IncludeSysError false
            }
        # Record values 
        $script:Operations | Add-Member -MemberType NoteProperty -Name LicenseGroupMemberShip -Value $GroupmemberShip -Force
        }
    }   
}



# FUNCTION - Check access using Get-ACL on directories to see access denied errors
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


# FUNCTION - Rename Users Home Drive Folder
function RenameUserHomeFolder () {

    # Create var to host replacement char
    $DateRename = Get-date -f yyyymmdd
    $ReplaceChar = '-DEL-'
    # Set th Error ActionPerference as Rename-item does not handle failed events
    $ErrorActionPreference = 'Stop'

    WriteTransactionsLogs -Task "Performing Rename of $HomeDrivePath on user $samAccountName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
    
    try {
        Rename-Item $HomeDrivePath $HomeDrivePath$ReplaceChar$DateRename -ErrorAction Stop -Force
        WriteTransactionsLogs -Task "Rename Completed for $HomeDrivePath to $HomeDrivePath$ReplaceChar$DateRename" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false}

    Catch{WriteTransactionsLogs -Task "Rename Failed for $HomeDrivePath to $HomeDrivePath$ReplaceChar$DateRename" -Result ERROR -ErrorMessage "UnableToProcess: " -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true}
}


# FUNCTION - Get Directories force takeown and Icals
function SingleDirectoryForceTakeOwnIcals () {

     # Build string
     $switches = ':(OI)(CI)F'
       
    WriteTransactionsLogs -Task "Performing takeown and Icacls on $HomeDrivePath" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
         
    takeown /F  $HomeDrivePath /R /A /D Y
    icacls $HomeDrivePath /grant "Domain Admins:(OI)(CI)F" /C /t /grant "$SamaccountName$switches" /t /C
    
    
}
          

# FUNCTION - SharePoint Licence Check via SKU
function SharePointCheckLicence () {

    WriteTransactionsLogs -Task "Checking for SharePoint SKU" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    if ($MsolUser.Licenses.ServiceStatus.ServicePlan.ServiceName -eq "SHAREPOINTENTERPRISE"){$SharePointLicFound = $true
       WriteTransactionsLogs -Task "User is assigned a SharePoint SKU" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
       $script:Operations | Add-Member -MemberType NoteProperty -Name SharePointLicence -Value $SharePointLicFound -Force

   } Else {
       WriteTransactionsLogs -Task "Check for SharePoint SKU Failed" -Result ERROR -ErrorMessage "User has not SharePoint SKU or Failed check" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
       $script:Operations | Add-Member -MemberType NoteProperty -Name SharePointLicence -Value "Not Assigned or has error" -Force
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

    WriteTransactionsLogs -Task "Checking OneDrive database for existing entires.... Please Wait" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false

    # Open database and check for existing user
    Open-LiteDBConnection $DBName -Mode shared | Out-Null
    $SearchResultsOneDrive = Find-LiteDBDocument -Collection 'OneDriveDetails' -limit 999999  | Where-Object {$_.owner -eq $Mail}
    Close-liteDBConnection

    # Check for data if not perform a lookup of the user via SharePoint
    if ($SearchResultsOneDrive) { 

        # Build Strings
        [string]$OneDriveCurrentSize = $SearchResultsOneDrive.StorageUsageCurrent
        $OneDriveURL = $SearchResultsOneDrive.url
        
        WriteTransactionsLogs -Task "$Mail is Provisioned for OneDrive" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $script:Operations  | Add-Member -MemberType NoteProperty -Name OneDriveProvisioned -Value $true -Force
        $script:Operations  | Add-Member -MemberType NoteProperty -Name OneDriveCurrentSize -Value $OneDriveCurrentSize -Force
        $script:Operations  | Add-Member -MemberType NoteProperty -Name OneDriveURL -Value $OneDriveURL -Force
        WriteTransactionsLogs -Task "Current OneDrive Size is $OneDriveCurrentSize MB" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        
    }
    Else {
        
        WriteTransactionsLogs -Task "$Mail OneDrive Data not found in Database, Ensure OneDrive data is up to data, run option 3" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false 
        $script:Operations  | Add-Member -MemberType NoteProperty -Name OneDriveProvisioned -Value $false -Force
        $script:Operations  | Add-Member -MemberType NoteProperty -Name OneDriveCurrentSize -Value '0' -Force
        $script:Operations  | Add-Member -MemberType NoteProperty -Name OneDriveURL -Value 'Not Provisioned' -Force
        }
}


# FUNCTION - Create Export SharePoint Migrations
function ExportSharePointCSVList () {

    WriteTransactionsLogs -Task "Generating CSV Migration Export for $Mail" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    
    # Open database and check for existing user
    Open-LiteDBConnection $DBName -Mode shared | Out-Null
    $SearchResultsOneDrive = Find-LiteDBDocument -Collection 'OneDriveDetails' -limit 999999 | Where-Object {$_.owner -eq $Mail}
    Close-liteDBConnection

    # Build strings for data export
    $OneDriveURL  = $SearchResultsOneDrive.url

    # Check if file exists before been created
    $MigrationResultsFile  = ".\SharePointMigration_$((get-date).ToString('yyyyMMdd')).csv"        # Results file from Output
    $MigrationResultsFileCheck =  Test-path $MigrationResultsFile 


    # Create File to export results
    if ($MigrationResultsFileCheck -eq $true){

        # Export to File
        Add-Content "$HomeDrivePath,,,$OneDriveURL,Documents," -path $MigrationResultsFile
    }
    Else {
        # Export to File to create headers in file
        Add-Content -Value "FileSharePath,,,SharePointSite,DocLibrary," -path $MigrationResultsFile
        Add-Content "$HomeDrivePath,,,$OneDriveURL,Documents," -path $MigrationResultsFile
    }

}


# FUNCTION - Get Msol User Details
function GetMsolUser () {

    try {$MsolUser = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
        WriteTransactionsLogs -Task "Found $UPN in Azure " -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
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
function TakeownTool () {

    if ($script:HomeDirectoryValid -eq $true){

        if ($Global:AllowDriveAccess -eq $false){
    
            takeown /F $HomeDrivePath\Data /R /A /D Y >$null

            # Build PScustomObject
            $script:Operations | Add-Member -MemberType NoteProperty -Name "TakeownRun"-Value $True
            $script:TakeownCompleted = $true
        }
        Else {WriteTransactionsLogs -Task "Takeown Not required as passed SimplePermissionsCheck on $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        $script:Operations | Add-Member -MemberType NoteProperty -Name "TakeownRun"-Value $false}

    } Else {WriteTransactionsLogs -Task "Cannot run Takeown as Testpath failed to validate $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
            $script:Operations | Add-Member -MemberType NoteProperty -Name "TakeownRun"-Value $false
            }
}


# FUNCTION - Add icacls Domain Admins Permssions
function IcaclsPermissions () {

    if ($script:TakeownCompleted -eq $true){

        # Build string
        $switches = ':(OI)(CI)F'
    
        WriteTransactionsLogs -Task "Running icacls on $HomeDrivePath Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
        $icaclStatus = icacls $HomeDrivePath\Data\ /grant "Domain Admins:(OI)(CI)F" /C /T /grant "$SamaccountName$switches" /T /C /Q


        # Build PScustomObjec
        $script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclsPermermissionsRun"-Value $True
        $script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclStatusResults"-Value $icaclStatus

        
    }
    Else {WriteTransactionsLogs -Task "Icacls Not required as Takeown task not completed on $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
         $script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclsPermermissionsRun"-Value $false   }
     
        
}


# FUNCTION - Add icacls Domain Admins Permssions
function IcaclsPermissionsfix () {

    # Build string
    $switches = ':(OI)(CI)F'
    
    WriteTransactionsLogs -Task "Running icacls on $HomeDrivePath Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    icacls $HomeDrivePath\Data\ /grant "Domain Admins:(OI)(CI)F" /C /T /grant "$SamaccountName$switches" /T /C



    # Build PScustomObjec
    $script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclsPermermissionsRun"-Value $True
    #$script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclStatusResults"-Value $icaclStatus
}


# FUNCTION - Make Home Drive Read only
function MakeHomeDriveReadOnly () {

    # Build string
    $switches = ':(OI)(CI)RX'
    
    WriteTransactionsLogs -Task "Running icacls on $HomeDrivePath to make Read Only...Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    icacls $HomeDrivePath\  /grant:r "$SamaccountName$switches" /T /C /Q

    # Build PScustomObject
    $script:Operations | Add-Member -MemberType NoteProperty -Name "MakeHomeDriveReadOnly"-Value $True
    WriteTransactionsLogs -Task "Finished making Drive Read Only for $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
}


# FUNCTION - None impacting Takeown and ICALS
function SlowTakeOwnIcals () {

    if ($script:HomeDirectoryValid -eq $true){

        if ($Global:AllowDriveAccess -eq $false){

            # Get Directories from the HomeDrivePath
            WriteTransactionsLogs -Task "Finding Directories and Files for $SamaccountName on $HomeDrivePath... Please wait" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            $Objects = robocopy $HomeDrivePath\DATA NULL /L /S /NJH /NJS /FP /NC /NS /XJ /R:0 /W:0
            $Objects = $Objects | Where-Object {$_ -ne ""}
            $Objects = $Objects.Trim()

            $ObjectsCount = $Objects | Measure-Object | Select-Object -ExpandProperty Count
            WriteTransactionsLogs -Task "Found $ObjectsCount objects for $SamaccountName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            $script:Operations | Add-Member -MemberType NoteProperty -Name DirectoriesFound -Value "$ObjectsCount" -Force

            # Build string
            $switches = ':(OI)(CI)F'

            $Counter = 0
            WriteTransactionsLogs -Task "Performing Takeown and Icals on $HomeDrivePath... Please wait" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            Foreach ($Object in $Objects){
    
                $Counter ++
                Write-Progress -Activity "Processing Files and Folders on $HomeDrivePath" -CurrentOperation $Object -PercentComplete (($Counter / $Objects.count) *100)

                Takeown.exe /f $Object /A
                #Takeown.exe /f $Object /A >$Null
          
                #icacls $Object /q /grant "Domain Admins$switches" /C /grant "$SamaccountName$switches" /C >$null
                icacls $Object /grant "cnasurety.net\Domain Admins$switches" /C /grant "$SamaccountName$switches" /C /grant "cna.com\Domain Admins$switches" /C 

         
    }
    # Build PScustomObject
    $script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclsPermermissionsRun"-Value $True
    $script:Operations | Add-Member -MemberType NoteProperty -Name "TakeownRun"-Value $true
    WriteTransactionsLogs -Task "Finished Processing File and Folders on $SamaccountName on $HomeDrivePath" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    
    }Else {WriteTransactionsLogs -Task "Takeown Not required as passed SimplePermissionsCheck on $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        $script:Operations | Add-Member -MemberType NoteProperty -Name "TakeownRun"-Value $false
        $script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclsPermermissionsRun"-Value $false}


        }Else {WriteTransactionsLogs -Task "Cannot run Takeown as Testpath failed to validate $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
            $script:Operations | Add-Member -MemberType NoteProperty -Name "TakeownRun"-Value $false
            $script:Operations | Add-Member -MemberType NoteProperty -Name "IcaclsPermermissionsRun"-Value $false
            }

    # Clear data
    $Objects = ''
}


# FUNCTION - Find Files on the users HomeDirectory
function FindFiles () {

    if ($Global:AllowDriveAccess -eq $true){

    WriteTransactionsLogs -Task "Starting $script:FileType Search for $SamaccountName on $HomeDrivePath" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false

    # Find PST files via RoboCopy
    $FilesRAW = Invoke-Expression -Command "Robocopy $HomeDrivePath NULL /L /S /NJH /NJS /FP /NC /NDL /NS /XJ /R:0 /W:0 $script:FileType"
    $FilesRaw = $FilesRaw.Trim()
    $FilesList =  $FilesRaw.Split('',[System.StringSplitOptions]::RemoveEmptyEntries)
    $FileCount = $FilesList | Measure-Object | Select-Object -ExpandProperty count
    #$PSTFilesList

    WriteTransactionsLogs -Task "Completed $script:FileType Search for $SamaccountName on $HomeDrivePath and found $FileCount $script:FileType's" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    $JobCurrentStatus = 'Completed'
    $script:Operations | Add-Member -MemberType NoteProperty -Name "FilesFound"-Value $FileCount -force
    
   
    if ($FilesList){
        foreach ($File in $FilesList){

        try{$FileData = Get-ChildItem -path $File -ea SilentlyContinue |  Select-Object Name,CreationTime,LastWriteTime,LastAccessTime,fullname,Length}
        catch{WriteTransactionsLogs -Task "Failed getting $script:FileType File for $SamaccountName on $HomeDrivePath" -Result ERROR -ErrorMessage "Error:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true}
        
        $Name             =       [string]$FileData.Name
        $CreationTime     =       $FileData.CreationTime
        $LastWriteTime    =       $FileData.LastWriteTime
        $LastAccessTime   =       $FileData.LastAccessTime
        $FullName         =       [string]$FileData.FullName
        $Length           =       $FileData.Length


        # Work out size of file and convert into KB/MB/GB/TB...
        $FileSize = FormatBytes $Length
     

        # Date
        $DateNow = Get-Date -f g  


        # Construct AD Searcher
        $ADSearcher = New-Object DirectoryServices.DirectorySearcher -Property @{Filter = "(samaccountname=$SamaccountName)"}
        $ADResult = $ADSearcher.Findone()
        [string]$SMTPAddress = $ADResult.Properties.mail 

        # If not address found build string to say not found
        If (!$SMTPAddress){$SMTPAddress = 'No Address Found for User'}

         # Build PScustomObject

        $FileSessionDetails  = @()
        $FileSessionDetails = [pscustomobject][ordered]@{}
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "BatchID" -Value "$nextbatch" -force
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "Samaccountname"-Value $Samaccountname
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "HomeDirectory"-Value $HomeDrivePath
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "Name"-Value $name
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "CreationTime"-Value $CreationTime
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "LastWriteTime"-Value $LastWriteTime
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "LastAccessTime"-Value $LastAccessTime
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "fullname"-Value $fullname
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "FileSize"-Value $FileSize
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "ServerName"-Value $HomedriveServer
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "WhenDiscovered"-Value $DateNow
        $FileSessionDetails | Add-Member -MemberType NoteProperty -Name "OwnerSMTPaddress"-Value $SMTPAddress

        # Add Batch Details to Database 
        Open-LiteDBConnection $DBName -Mode shared | Out-Null
        $FileSessionDetails | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection FileSessionDetails
        Close-liteDBConnection

        }
        
   }
   $script:Operations | Add-Member -MemberType NoteProperty -Name "FileSearchInformation"-Value "Completed for $script:FileType" -force

   }
   Else {WriteTransactionsLogs -Task "Skipped $script:FileType Search for $SamaccountName on $HomeDrivePath as failed simple drive access check" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    $script:Operations | Add-Member -MemberType NoteProperty -Name "FileSearchInformation"-Value 'Skipped due to Drive Access Check Fail' -force}
}


# FUNCTION - Simple check for access permission issues in DATA directories 
function SimplePermissionsCheck () {

    if ($script:HomeDirectoryValid -eq $True){

    WriteTransactionsLogs -Task "Running simple permissions check on $HomeDrivePath Please wait..." -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    

    # Reset Global Variable
    $Global:AllowDriveAccess = $true

    # Does Directory have any content
    try {$DirectoryInformation = Get-ChildItem $HomeDrivePath\DATA -ea Stop | Measure-Object}
    Catch {WriteTransactionsLogs -Task "DATA Directory does not exist on $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false}

    if ($DirectoryInformation.count -ge 1){

        # Check if test folders exist for get-acl
        try {$Desktop        = Test-Path $HomeDrivePath\DATA\Desktop -ea stop} catch {}
        try {$Documents      = Test-Path $HomeDrivePath\DATA\Documents -ea stop} catch {}
    
        if ($Desktop -eq $true){

            try{$FullPermissions = Get-Acl -ea stop "$HomeDrivePath\DATA\Desktop" ;$Permissions = $FullPermissions.access | Where-Object {$_.IdentityReference -match $PermissionsGroup} | Select-Object IdentityReference,FileSystemRights
                              
                if (!($Permissions)) {WriteTransactionsLogs -Task "No Access to $HomeDrivePath\DATA\Desktop for $UPN, Investigating Needed" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                    $Global:AllowDriveAccess = $false}
                    }
            catch {WriteTransactionsLogs -Task "No Access  to $HomeDrivePath\DATA\Desktop for $UPN, Investigating Needed" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                $Global:AllowDriveAccess = $false}
        }

        if ($Documents -eq $true){

            try{$FullPermissions = Get-Acl -ea stop "$HomeDrivePath\DATA\Documents" ;$Permissions = $FullPermissions.access | Where-Object {$_.IdentityReference -match $PermissionsGroup} | Select-Object IdentityReference,FileSystemRights
                if (!($Permissions)) {WriteTransactionsLogs -Task "No Access to $HomeDrivePath\DATA\Documents for $UPN, Investigating Needed" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                    $Global:AllowDriveAccess = $false}
                    }
            catch {WriteTransactionsLogs -Task "No Access  to $HomeDrivePath\DATA\Documents for $UPN, Investigating Needed" -Result information -ErrorMessage "ERROR" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                $Global:AllowDriveAccess = $false}
        }
    
        
    }



    if ($Global:AllowDriveAccess -eq $true){WriteTransactionsLogs -Task "Passed simple permissions check on $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
    if ($Global:AllowDriveAccess -eq $false){WriteTransactionsLogs -Task "Failed simple permissions check on $HomeDrivePath" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false}

     }Else {
     $Global:AllowDriveAccess = $false}
     
     # Build PScustomObject
     $Operations | Add-Member -MemberType NoteProperty -Name "AllowDriveAccess" -Value $Global:AllowDriveAccess -force
        

}


# FUNCTION - Split line function 
function Splitline (){


Write-host "-----------------------------------------------------------------------------------------" -ForegroundColor DarkBlue

}


function BatchInformation () {
    
    # Get-date
    $datenow = Get-Date -f g

    # Open Database Connection
    Open-LiteDBConnection $DBName -Mode shared | Out-Null

    # Clean up variable
    Remove-Variable nextbatch -Scope Global -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
 
    # Find if any batches have been created before
    $AnyExistingBatches = Find-LiteDBDocument -Collection 'batchinformation'

    if ($AnyExistingBatches){
     
        # Find last batch ID
        [int]$lastbatch = Find-LiteDBDocument -Collection 'batchinformation' | Select-Object -ExpandProperty  _id | Measure-Object -Maximum | Select -ExpandProperty Count

        # New batch ID
        [int]$Global:nextbatch = $lastbatch +1

        }

        # Build Pre information for report
        $ADUserInfoCount = $ADUserInfo.count
        $AdminRunBy = whoami


        # Build PScustomObject
        $BatchReport = @()
        $BatchReport = [pscustomobject][ordered]@{}

        if ($Global:nextbatch) {$BatchReport | Add-Member -MemberType NoteProperty -Name "_id" -Value "$nextbatch" -force}
        if (!($Global:nextbatch)){$BatchReport | Add-Member -MemberType NoteProperty -Name "_id" -Value "1001" -force ;[int]$global:nextbatch = '1001' }
        
        $BatchReport | Add-Member -MemberType NoteProperty -Name "DateJobCreation" -Value "$datenow" -force
        $BatchReport | Add-Member -MemberType NoteProperty -Name "UserCount" -Value $ADUserInfoCount -force
        $BatchReport | Add-Member -MemberType NoteProperty -Name "JobType" -Value $script:JobType -force
        $BatchReport | Add-Member -MemberType NoteProperty -Name "ScriptVersion" -Value $Version -force
        $BatchReport | Add-Member -MemberType NoteProperty -Name "AdminRunBy" -Value $AdminRunBy -force

        # Add Batch Details to Database 
        $BatchReport | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection batchinformation
        Close-liteDBConnection

}


# FUNCTION - Create operation collector
function CreateOperationCollector () {


    # Build PScustomObject for operations events
    $Script:Operations = @()
    $Script:Operations = [pscustomobject][ordered]@{}


}


# FUNCTION - Check Powershell version 
function CheckPSVersion () {

    if ($PSVersionTable.PSVersion.Major -ge '7') {}
    Else {
    
    Write-host
    Write-host
    Write-host
    Write-host "Powershell 7 and above is required to run this script. This script will now close." -ForegroundColor YELLOW 
    Write-host
    Write-host

    Exit
           
    }

}


# FUNCTION - Store Collected Data
function SaveStoredResults () {

    $datenow = Get-Date -f g

    $script:Operations  | Add-Member -MemberType NoteProperty -Name "BatchID" -Value "$nextbatch" -force
    $script:Operations  | Add-Member -MemberType NoteProperty -Name "UPN" -Value "$UPN" -force
    $script:Operations  | Add-Member -MemberType NoteProperty -Name "SamAccountName"-Value $SamaccountName -force
    $script:Operations  | Add-Member -MemberType NoteProperty -Name "HomeDrivePath" -Value $HomeDrivePath -force
    $script:Operations  | Add-Member -MemberType NoteProperty -Name "JobType" -Value $JobType -force
    $script:Operations  | Add-Member -MemberType NoteProperty -Name "Date" -Value $datenow -force
    $script:Operations  | Add-Member -MemberType NoteProperty -Name "Mail" -Value $mail -force

    # Add Batch Details to Database 
    Open-LiteDBConnection $DBName -Mode shared | Out-Null
    $script:Operations | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection Operations
    Close-liteDBConnection

    # Clear existing data
    Remove-Variable Operations -Scope script -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

    
    # Build PScustomObject for operations events
    $Script:Operations = @()
    $Script:Operations = [pscustomobject][ordered]@{}

}

##################### Report Functions #######################

function AskForBatchID () {

    Splitline
    $Script:BatchIDValue = Read-Host -Prompt "Enter the BATCH ID of the Job to Export"
    Splitline
    Write-Host `n
    if ($Script:BatchIDValue -eq "") {WriteTransactionsLogs -Task "BATCH ID not entered" -Result Error -ErrorMessage "No batch ID entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        Write-Host `n
        AskForBatchID 
    }
   
}
# FUNCTION ASk for SAmaccount ID
function AskForSamaccountID () {

    Splitline
    $script:SamaccountNameValue = Read-Host -Prompt "Enter the SAmaccountName ID"
    Splitline
    Write-Host `n
    if ($script:SamaccountNameValue -eq "") {WriteTransactionsLogs -Task "SamaccountName not entered" -Result Error -ErrorMessage "No Sam ID entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        Write-Host `n
        AskForSamaccountID 
    }
   
}

# FUNCTION - Export Operations based on BatchID
function ExportOperationsByBatchID () {

    # Show report data and tee-object to var for export
    Open-LiteDBConnection $DBName -Mode shared | Out-Null
    Find-LiteDBDocument -Collection 'Operations' -Limit 999999 | Where-Object {$_.BatchID -eq $script:BatchIDValue} | Select-Object *
    $ReportData = Find-LiteDBDocument -Collection 'Operations' -Limit 999999 | Where-Object {$_.BatchID -eq $script:BatchIDValue}

    if ($ReportData) {
        # Store the report data in script var
        $script:ReportData = $ReportData}
    Else {WriteTransactionsLogs -Task "No Batch data found for $script:BatchIDValue " -Result Warning -ErrorMessage "No Batch DataFound" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        AskForBatchID
    }

}

# FUNCTION - Export File Search Results based on BatchID
function ExportFileSearchByBatchID () {

    # Show report data and tee-object to var for export
    Open-LiteDBConnection $DBName -Mode shared | Out-Null
    Find-LiteDBDocument -Collection 'FileSessionDetails' -Limit 999999 | Where-Object {$_.BatchID -eq $script:BatchIDValue} | Select-Object *
    $ReportData = Find-LiteDBDocument -Collection 'FileSessionDetails' -Limit 999999 | Where-Object {$_.BatchID -eq $script:BatchIDValue}

    if ($ReportData) {
        # Store the report data in script var
        $script:ReportData = $ReportData}
    Else {WriteTransactionsLogs -Task "No Batch data found for $script:BatchIDValue " -Result Warning -ErrorMessage "No Batch DataFound" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        AskForBatchID
    }

}


# FUNCTION - Search by Samaccount ID for transactions
function SearchTransationBySamaccountName () {

    # Show report data and tee-object to var for export
    Open-LiteDBConnection $DBName -Mode shared | Out-Null
    Find-LiteDBDocument -Collection 'Transactions' -Limit 999999 | Where-Object {$_.SamaccountName -eq $script:SamaccountNameValue} | Select-Object *
    $ReportData = Find-LiteDBDocument -Collection 'Transactions' -Limit 999999 | Where-Object {$_.SamaccountName -eq $script:SamaccountNameValue} 

    if ($ReportData) {
        # Store the report data in script var
        $script:ReportData = $ReportData}
    Else {WriteTransactionsLogs -Task "No Batch data found for $script:SamaccountnameValue " -Result Warning -ErrorMessage "No Batch DataFound" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        AskForSamaccountID 
    }
}


# FUNCTION - Export Report data stored in the $script:ReportData Var
function AskToExportReport () {

    $Input = Read-Host -Prompt "Do you want to export the current report data to File" 
    If ($Input -match 'No') {ShowReportMenu}
    if ($Input -match 'Yes') { 
            
        $SaveMyFile = Get-Filename
        $script:ReportData | Export-Csv -Path $SaveMyFile
        
    }

}

# FUNCTION - Save File GUI
Function Get-FileName($initialDirectory) {   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = "CSV file (*.csv)|*.csv| All Files (*.*)|*.*";
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFileDialog.filename
}


# Export functions for multithread scripts
$export_functions = [scriptblock]::Create(@" 

function GetServerShareSplit {$function:GetServerShareSplit}
function TestConnection {$function:TestConnection}
function TestPathHomeDrive  {$function:TestPathHomeDrive}
function FindFiles {$function:FindFiles}
function WriteTransactionsLogs {$function:WriteTransactionsLogs}
function FormatBytes{$function:FormatBytes}
function DatabaseConnection{$function:DatabaseConnection}
function ImportPSliteDBModule{$function:ImportPSliteDBModule}
function TerminateScript{$function:TerminateScript}
function CheckforPSliteDBModule{$function:CheckforPSliteDBModule}
function TakeownTool{$function:TakeownTool}
function CheckMSonlineModule{$function:CheckMSonlineModule}
function SharePointModule{$function:SharePointModule}
function AskForAdminCreds{$function:AskForAdminCreds}
function ConnectMicrosoftOnline{$function:ConnectMicrosoftOnline}
function ConnectMicrosoftSharePoint{$function:ConnectMicrosoftSharePoint}
function FindAdminLogonID{$function:FindAdminLogonID}
function CheckPermissionsOnline{$function:CheckPermissionsOnline}
function ProvisionOneDriveUser{$function:ProvisionOneDriveUser}
function SharePointCheckLicence{$function:SharePointCheckLicence}
function SimplePermissionsCheck{$function:SimplePermissionsCheck}
function IcaclsPermissions{$function:IcaclsPermissions}
function IcaclsPermissionsFix{$function:IcaclsPermissionsFix}
function GetHomeDirectorySize{$function:GetHomeDirectorySize}
function BatchInformation{$function:BatchInformation}
function SaveStoredResults{$function:SaveStoredResults}
function CreateOperationCollector{$function:CreateOperationCollector}
function AskForFileType{$function:AskForFileType}
function SlowTakeOwnIcals{$function:SlowTakeOwnIcals}
function CheckADGroupMembership{$function:CheckADGroupMembership}
function ImportExclusionGroupsFile{$function:ImportExclusionGroupsFile}
function CheckExcludeGroupFile{$function:CheckExcludeGroupFile}
function GetOneDriveDetails{$function:GetOneDriveDetails}
function CheckLicenceGroupFile{$function:CheckLicenceGroupFile}
function CheckADGroupMembershipForLicence{$function:CheckADGroupMembershipForLicence}
function ImportLicenceGroupsFile{$function:ImportLicenceGroupsFile}
function GetHomeAppDataDirectorySize{$function:GetHomeAppDataDirectorySize}
function MakeHomeDriveReadOnly{$function:MakeHomeDriveReadOnly}
"@)


# Run main function to call others
DisplayExtendedInfo
CheckPSVersion
ShowMenuSelectUsers

