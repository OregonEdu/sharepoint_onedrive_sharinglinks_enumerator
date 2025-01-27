## SharePoint/OneDrive User Enumerator - v2.0
## Benjamin Barshaw <benjamin.barshaw@ode.oregon.gov> - IT Operations & Support Network Team Lead - Oregon Department of Education
#
#  Requirements: PnP PowerShell Module
#                SharePoint Admin  
#
#  This script will export by default all true external "guest" users for a SiteCollection or all external SharingLinks for a SiteCollection(s) within an M365 tenant. I have added options to toggle INVITED external users (users with #ext# 
#  in their UPN in Entra) as well as internal users to your tenant. In v1.0 of the script the OneDrive SharingLinks were built out from Get-PnPUser because the Groups the user object returns contain references to the SharingLink 
#  themselves. Then during testing I made the unfortunate finding that when using the built-in functionality to "Send" an e-mail with the share it does NOT generate a SharingLink! I have modified the way OneDrive shares are handled in
#  this version of the script to accomodate for this but without a SharingLink it is impossible to determine when the link was shared. For these types the "InvitedOn" column is simply set to "N/A". For SharePoint sites I have the libraries
#  you want to export set in the USER-DEFINED VARIABLES section below which you can add/remove to suit your needs.  This script came about by me noticing thata SharingLink came in a format of (for example):
#
#  SharingLinks.af100805-bd06-4c03-b9a0-f9506f6e8d57.Flexible.47d674bc-3399-41d3-a5fb-f84b2c04df52
#
#  By digging through the PnP cmdlets I was able to determine that the first long hexadecimal string correlated with the UniqeId of a file in a Document library. So it is the file/folder being shared! In further digging I was able
#  to determine that the second long hexadecimal string correlated to the SharingLink itself and you could do lookups on each by using values embedded in each object. This was not fun nor easy. :)
#
#  It should be noted that even with SharePoint Admin you need to manually add your PA account as a SiteCollectionAdmin -- SharePoint Admin does NOT give you carte blanche.
#
#  To make it easier to delineate between the comments and actual code, PowerShell ISE or Visual Studio Code is recommended for editing/reading.
#
#  If you have any questions/comments, please feel free to reach out to me via Teams or e-mail.
#
#  P.S. - One of my closest personal friends Sean McArdle makes fun of how "ANSI C" my PowerShell is. I learned programming from "Programming in ANSI C" by Stephen Kochan -- so this checks out! :)

# Create our class for the SharingLink CSV export
class SharingLink
{
    # URL of the SharePoint/OneDrive site
    [string]$SiteURL
    # DisplayName of the external guest
    [string]$Title
    # LoginName of the external guest
    [string]$LoginName
    # Email of the external guest
    [string]$Email
    # Date they were invited
    [string]$InvitedOn
    # Show if it's a file or folder that was shared
    [string]$ObjectType
    # Show what was shared
    [string]$Object
    # Permission level
    [string]$Permission      
}

# Create our class for our ExternalUser CSV export
class ExternalUser
{
    # URL of the SharePoint/OneDrive site
    [string]$SiteURL
    # DisplayName of the external guest
    [string]$Title
    # LoginName of the external guest
    [string]$LoginName
    # Email of the external guest
    [string]$Email
}

## This is a method to import the SharePoint module in PowerShell 7 and have authentication work -- I removed ALL the SharePoint module code but left this commented out to help anyone else wanting to use the module in PS7
# Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell

### USER DEFINED VARIABLES SECTION ###
#
# PA account to add to user's OneDrive as SiteCollectionAdmin
$paAccount = "pa_ode_bbarshaw@odemail.onmicrosoft.com"
# Change this to the admin portal of your SharePoint
$spoAdmin = "https://odemail-admin.sharepoint.com"
# Change this to the ClientID of the PnP Entra application
$pnpAppClientId = "ab26adab-5273-4711-ad54-7314e685d34f"
# Change this to your TenantID
$tenantId = "b4f51418-b269-49a2-935a-fa54bf584fc8"
# Change this to the SharePoint Libraries you want to traverse
$sharePointLibraries = @('Documents', 'Site Pages')
### END USER DEFINED VARIABLES SECTION ###

## GLOBAL VARIABLES SECTION - DON'T TOUCH PLEASE ###
#
# Version of the script 
$version = "v2.0"
# Get today's date formatted for a filename
$getDate = Get-Date -Format MMddyyyy
# Toggle for SiteCollectionAdmin check
$siteCollectionAdminFlag = 0
# Toggle for External Invited users -- on by default as we want ALL external
$externalInvitedFlag = 1
# Toggle for Internal users
$internalInvitedFlag = 0
# Tally of sites found
$getSPOSites = 0
### END GLOBAL "DON'T TOUCH PLEASE" VARIABLES SECTION ###

# Display our menu 
function showMenu
{
    Write-Host -ForegroundColor Cyan "SharePoint/OneDrive User Enumerator $($version) - Benjamin Barshaw <benjamin.barshaw@ode.oregon.gov>"    
    Write-Host -ForegroundColor Yellow "Please make your selection:"    
    Write-Host -ForegroundColor DarkYellow " 1) Load SharePoint sites for tenant"
    Write-Host -ForegroundColor DarkYellow " 2) Load OneDrive sites for tenant"
    Write-Host -ForegroundColor DarkYellow " 3) Load SharePoint & OneDrive sites for tenant"
    Write-Host -ForegroundColor DarkYellow " 4) Load SharePoint or OneDrive for a single site/user"
    Write-Host -ForegroundColor DarkYellow " 5) Load SharePoint or OneDrive sites from a CSV (NOTE: Must have Url header!)"
    Write-Host -ForegroundColor DarkYellow " 6) Toggle check to see if PA account is SiteCollectionAdmin for loaded sites (Will export to SiteCollectionAdmin-<date>.csv if not already admin)"
    Write-Host -ForegroundColor DarkYellow " 7) Add PA account as SiteCollectionAdmin from SiteCollectionAdmin CSV"
    Write-Host -ForegroundColor DarkYellow " 8) Add PA account as SiteCollectionAdmin for ALL loaded sites"
    Write-Host -ForegroundColor DarkYellow " 9) Remove PA account as SiteCollectionAdmin from SiteCollectionAdmin CSV"
    Write-Host -ForegroundColor DarkYellow "10) Remove PA account as SiteCollectionAdmin for ALL loaded sites"
    Write-Host -ForegroundColor DarkYellow "11) Toggle checking for external INVITED users (i.e.: has #ext# in UPN in Entra)"
    Write-Host -ForegroundColor DarkYellow "12) Toggle checking for internal users"
    Write-Host -ForegroundColor DarkYellow "13) Export external users for loaded sites (Exports to .\External_Users_Export-<todays_date>.csv)"
    Write-Host -ForegroundColor DarkYellow "14) Export SharingLinks for loaded sites (Exports to .\SharingLinks_Exports-<todays_date>.csv)"
    Write-Host -ForegroundColor DarkYellow "15) Cleanup CSV's"
    Write-Host -ForegroundColor DarkYellow "16) Show menu"    
    Write-Host -ForegroundColor DarkYellow "17) Exit"
    # Display our toggle flags
    Write-Host -ForegroundColor Cyan -NoNewline "SiteCollectionAdmin Flag: "
    Write-Host -ForegroundColor Magenta "$([bool]$siteCollectionAdminFlag)"
    Write-Host -ForegroundColor Cyan -NoNewline "ExternalInvitedFlag: "
    Write-Host -ForegroundColor Magenta "$([bool]$externalInvitedFlag)"
    Write-Host -ForegroundColor Cyan -NoNewline "InternalInvitedFlag: "
    Write-Host -ForegroundColor Magenta "$([bool]$internalInvitedFlag)"
    # Display count of how many sites are loaded
    Write-Host -ForegroundColor Cyan -NoNewline "Loaded sites: "
    If (! $getSPOSites)
    {
        Write-Host -ForegroundColor Magenta "0"
    }
    Else
    {
        Write-Host -ForegroundColor Magenta "$($getSPOSites.Count)"
    }
}

# This functions checks to see if your PA is already a SiteCollectionAdmin and if so will not add it to the CSV we create to track the sites you added your PA to. The logic here being that there could be site collections that your PA
# SHOULD be SiteCollectionAdmin for and is already. By handling it this way, we won't remove your PA from these sites when/if you remove yourself later from the CSV.
function checkSiteCollectionAdmin($getSPOSites)
{    
    ForEach ($spoSite in $getSPOSites)
    {
        Connect-PnPOnline -Url $spoSite -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                
        If ((Get-PnPSiteCollectionAdmin -ErrorAction Ignore).Email -contains $paAccount) 
        { 
            Write-Host -ForegroundColor DarkYellow "$($paAccount) was already a SiteCollectionAdmin for $($spoSite)! Not exporting to the CSV."
        }
        Else
        {                                                                
            Write-Host -ForegroundColor Cyan "Adding $($spoSite) to SiteCollectionAdmin-$($getDate).csv"
            [PSCustomObject]@{Url = $spoSite} | Export-Csv -NoTypeInformation -Append -Path ".\SiteCollectionAdmin-$getDate.csv"
        }
    }
}

# Add our PA as SiteCollectionAdmin for the sites specified -- this is needed to do ANY kind of work on them
function addSiteCollectionAdmin($getSPOSites)
{
    try
    {
        Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
        Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
        exit
    }   

    ForEach ($spoSite in $getSPOSites)
    {
        try
        {            
            Set-PnPTenantSite -Identity $spoSite -Owners $paAccount
            Write-Host -ForegroundColor Cyan "Successfully added $($paAccount) as SiteCollectionAdmin to $($spoSite)!"
        }
        catch
        {
            Write-Host -ForegroundColor Red "Could not add PA account to $($spoSite)!"
        }
    }
}

# Remove our PA as SiteCollectionAdmin for the sites specified
function removeSiteCollectionAdmin($getSPOSites)
{
    ForEach ($spoSite in $getSPOSites)
    {
        Connect-PnPOnline -Url $spoSite -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
        Remove-PnPSiteCollectionAdmin -Owners $paAccount
        Write-Host -ForegroundColor Cyan "Successfully removed $($paAccount) as SiteCollectionAdmin for $($spoSite)!"
    }
}

# The meat of the script. I had an internal argument about splitting up the OneDrive/SharePoint exports into 2 seperate functions. The problem with that is though on each subsequent function call we would need to re-authenticate on
# the Connect-PnPOnline! If we keep it all in the same memory space we don't have that issue. 
function getSharingLinks($getSPOSites)
{
    ForEach ($spoSite in $getSPOSites)
    {        
        # OneDrive has some nuances that differ slightly from SharePoint -- notably the way to extract when the SharingLink was created
        $oneDriveFlag = 0
        # Set this to 1 (true) and 0 (false) if they don't exist which is entirely possible if it's a new SiteCollection
        $sharingLinksFlag = 1

        # If the URL contains *-my.sharepoint.com/personal* we know it's a OneDrive and set the flag accordingly
        If ($spoSite -like "*-my.sharepoint.com/personal*")
        {
            $oneDriveFlag = 1            
        }
        
        Connect-PnPOnline -Url $spoSite -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                
        If ($oneDriveFlag)
        {
            Write-Host -ForegroundColor Yellow "Compiling PnP List of Documents for OneDrive site $($spoSite)..."
            # Unlike SharePoint -- we are ONLY concerned with the "Documents" library for a OneDrive as it is where all the files are located
            $getPnPListItems = Get-PnPListItem -List "Documents" -PageSize 5000
            Write-Host -ForegroundColor Yellow "Comping PnP List of SharingLinks for OneDrive site $($spoSite)..."
            # Try to get all the SharingLinks objects -- if it fails we set the flag so we know not to do work on the object
            try
            {
                $getPnPSharingLinks = Get-PnPListItem -List "Sharing Links" -PageSize 5000
            }
            catch
            {
                Write-Host -ForegroundColor Yellow "SharingLinks doesn't exist for $($spoSite)!"
                $sharingLinksFlag = 0
            }
                
            # We by default get all external SPO guests but we consider external invited users to be "internal" -- this flag will add external invited users to our array of objects that we will export to our report
            If ($externalInvitedFlag)
            {
                Write-Host -ForegroundColor Yellow "ExternalInvitedFlag set! Adding external invited users..."
            }

            # Again, by default we skip internal users -- this flag was added by a request from my good friend Vahan Michaelian "just in case" -- I would NOT recommend setting it for the entire tenant unless you got a week to spare :)
            If ($internalInvitedFlag)
            {
                Write-Host -ForegroundColor Yellow "InternalInvitedFlag set! Adding internal users..."
            }

            # Originally we cycled through USERS instead of the Document library but I had to change it to match SharePoint's method since SharingLinks aren't generated for e-mail invites! BOO!!
            ForEach ($pnpListItem in $getPnPListItems)
            {
                # Check to see if the file/folder has unique permissions -- unique being permissions created outside of the inherited permissions
                If (($getUniquePermissions = Get-PnPProperty -ClientObject $pnpListItem -Property "HasUniqueRoleAssignments") -eq $true)
                {
                    # We initialize the RoleAssignments collection which we will need to get permissions given to the external user
                    $getRoleAssignments = Get-PnPProperty -ClientObject $pnpListItem -Property RoleAssignments -ErrorAction SilentlyContinue
                    # We will cycle through each role assignment to extract the information we need
                    ForEach ($roleAssignment in $getRoleAssignments)
                    {
                        # We initialize the RoleDefinitionBindings and Member collections using the RoleAssignment as the object
                        $getMembers = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings, Member -ErrorAction SilentlyContinue
                        # Finally we initialize the Users collection to get the users associated with the particular file/folder
                        $getUsers = Get-PnPProperty -ClientObject $roleAssignment.Member -Property Users -ErrorAction SilentlyContinue
                        # Here we store the permission itself (Read/Contribute) -- they are organized by permission level
                        $getRoleDefinitionBindings = $roleAssignment.RoleDefinitionBindings.Name
                        
                        # This section is for if a person did a "Copy link" instead of a "Send" (email) which does NOT generate a SharingLink!
                        If ($roleAssignment.Member.Title -like "SharingLinks.*")
                        {
                            # Grab the second long hexadecimal string -- we don't need the first one since we are already working on the file/folder (document) itself
                            $getSharingLinkUniqueId = $roleAssignment.Member.Title.Split('.')[3]

                            try
                            {
                                # OneDrive method of extracting InvitedOn date
                                $getSharingLinkInvitedDate = Get-Date -Date (Get-Date -Date (($getPnPSharingLinks.FieldValues.AvailableLinks | ConvertFrom-Json | Where-Object { $_.ShareId -eq $getSharingLinkUniqueId }).Invitees.InvitedOn | Select-Object -First 1).ToLocalTime()) -Format "MM/dd/yyyy hh:mm:sstt"
                            }
                            catch
                            {
                                # If no SharingLink set it to "N/A"
                                $getSharingLinkInvitedDate = "N/A"
                            }

                            # If it was a SharingLink we'll have user information -- not if sent via e-mail (see below)
                            If ($getUsers)
                            {
                                ForEach ($spoUser in $getUsers)
                                {
                                    # Set a flag that gets set if any of our conditions are met
                                    $processFlag = 0                                        
                                    
                                    # We can use the same flag for all conditions -- if they are met we want them output just the same
                                    If ($externalInvitedFlag -and $spoUser.LoginName -like "*#ext#*")                                           
                                    {                                            
                                        $processFlag++
                                    }
                                    ElseIf ($internalInvitedFlag -and (($spoUser.LoginName -notlike "*#ext#*") -or ($spoUser.LoginName -notlike "*urn%3aspo%3aguest#*")))
                                    {                                                                                        
                                        $processFlag++
                                    }
                                    ElseIf ($spoUser.LoginName -like "*urn%3aspo%3aguest#*")
                                    {                                            
                                        $processFlag++
                                    }
                                    
                                    # This is the same export using the SharingLink class as the OneDrive section
                                    If ($processFlag)
                                    {
                                        Write-Host -ForegroundColor Cyan "$($spoUser.Title): $($pnpListItem.FileSystemObjectType): $($pnpListItem.FieldValues.FileRef)"
                                        $exportMe = [SharingLink]::new()
                                        $exportMe.SiteURL = $spoSite
                                        $exportMe.Title = $spoUser.Title
                                        $exportMe.LoginName = $spoUser.LoginName
                                        $exportMe.Email = $spoUser.Email
                                        $exportMe.InvitedOn = $getSharingLinkInvitedDate
                                        $exportMe.ObjectType = $pnpListItem.FileSystemObjectType
                                        $exportMe.Object = $pnpListItem.FieldValues.FileRef
                                        $exportMe.Permission = $getRoleDefinitionBindings
                           
                                        $exportMe | EXport-Csv -NoTypeInformation -Append -Path ".\SharingLinks_Export-$getDate.csv"
                                    }
                                }
                             }                             
                        }
                        # This is our emailed invite section! Here we set an ElseIf there is a permission not like the default "Limited Access*"
                        ElseIf ($roleAssignment.Member.Title -notlike "Limited Access*")
                        {                            
                            # The rest of this section is the same except we work off the RoleAssignment collection instead of the Users
                            $processFlag = 0                                        
                                                        
                            If ($externalInvitedFlag -and $roleAssignment.Member.LoginName -like "*#ext#*")                                           
                            {                                                                            
                                $processFlag++
                            }
                            ElseIf ($internalInvitedFlag -and (($roleAssignment.Member.LoginName -notlike "*#ext#*") -or ($roleAssignment.Member.LoginName -notlike "*urn%3aspo%3aguest#*")))
                            {                                                                                                                      
                                $processFlag++
                            }
                            ElseIf ($roleAssignment.Member.LoginName -like "*urn%3aspo%3aguest#*")
                            {                                                                            
                                $processFlag++
                            }

                            If ($processFlag)
                            {
                                Write-Host -NoNewline -ForegroundColor Magenta "E-mail invite detected! "
                                Write-Host -ForegroundColor Cyan "$($roleAssignment.Member.Title): $($pnpListItem.FileSystemObjectType): $($pnpListItem.FieldValues.FileRef)"
                                $exportMe = [SharingLink]::new()
                                $exportMe.SiteURL = $spoSite
                                $exportMe.Title = $roleAssignment.Member.Title
                                $exportMe.LoginName = $roleAssignment.Member.LoginName
                                $exportMe.Email = $roleAssignment.Member.Email
                                $exportMe.InvitedOn = "N/A" # No SharingLink since it was an e-mail invite
                                $exportMe.ObjectType = $pnpListItem.FileSystemObjectType
                                $exportMe.Object = $pnpListItem.FieldValues.FileRef
                                $exportMe.Permission = $getRoleDefinitionBindings
                   
                                $exportMe | EXport-Csv -NoTypeInformation -Append -Path ".\SharingLinks_Export-$getDate.csv"
                            }
                        }                                                     
                    }
                }
            }
        }
        Else
        {
            # Unlike OneDrive, a SharePoint site has MANY libraries we could be interested in knowing if they have been shared out so we cycle through each one of them defined in the $sharePointLibraries array at the top of the script
            ForEach ($sharePointLibrary in $sharePointLibraries)
            {            
                Write-Host -ForegroundColor Yellow "Compiling PnP List of $($sharePointLibrary) for SharePoint site $($spoSite)..."
                # Get the items for the library we are currently working on
                $getPnPListItems = Get-PnPListItem -List $sharePointLibrary -PageSize 5000
                Write-Host -ForegroundColor Yellow "Comping PnP List of Sharing Links for $($sharePointLibrary) for SharePoint site $($spoSite)..."
                try
                {
                    $getPnPSharingLinks = Get-PnPListItem -List "Sharing Links" -PageSize 5000
                }
                catch
                {
                    Write-Host -ForegroundColor Yellow "SharingLinks doesn't exist for $($spoSite)!"
                    $sharingLinksFlag = 0
                }            
                                
                If ($externalInvitedFlag)
                {
                    Write-Host -ForegroundColor Yellow "ExternalInvitedFlag set! Adding external invited users..."
                }

                If ($internalInvitedFlag)
                {
                    Write-Host -ForegroundColor Yellow "InternalInvitedFlag set! Adding internal users..."
                }

                Write-Host -ForegroundColor Yellow "Adding SPO guest users..."

                # SharePoint doesn't add SharingLinks to the Groups of a user so we have to cycle through the file/folders themselves
                ForEach ($pnpListItem in $getPnPListItems)
                {
                    # The rest of this is virtually identical to the OneDrive method                    
                    If (($getUniquePermissions = Get-PnPProperty -ClientObject $pnpListItem -Property "HasUniqueRoleAssignments") -eq $true)
                    {                        
                        $getRoleAssignments = Get-PnPProperty -ClientObject $pnpListItem -Property RoleAssignments -ErrorAction SilentlyContinue
                        ForEach ($roleAssignment in $getRoleAssignments)
                        {                            
                            $getMembers = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings, Member -ErrorAction SilentlyContinue
                            $getUsers = Get-PnPProperty -ClientObject $roleAssignment.Member -Property Users -ErrorAction SilentlyContinue
                            $getRoleDefinitionBindings = $roleAssignment.RoleDefinitionBindings.Name
                            
                            If ($roleAssignment.Member.Title -like "SharingLinks.*")
                            {                                
                                $getSharingLinkUniqueId = $roleAssignment.Member.Title.Split('.')[3]                            
                            
                                try
                                {
                                    # For a SharePoint site SharingLink we use the CreatedDate property instead of the InvitedOn OneDrive uses
                                    $getSharingLinkInvitedDate = Get-Date -Date (Get-Date -Date (($getPnPSharingLinks.FieldValues.AvailableLinks | ConvertFrom-Json | Where-Object { $_.ShareId -eq $getSharingLinkUniqueId }).CreatedDate).ToLocalTime()) -Format "MM/dd/yyyy hh:mm:sstt"
                                }
                                catch
                                {
                                    $getSharingLinkInvitedDate = "N/A"
                                }

                                If ($getUsers)
                                {
                                    ForEach ($spoUser in $getUsers)
                                    {                                        
                                        $processFlag = 0                                        
                                                                                
                                        If ($externalInvitedFlag -and $spoUser.LoginName -like "*#ext#*")                                           
                                        {                                            
                                            $processFlag++
                                        }
                                        ElseIf ($internalInvitedFlag -and (($spoUser.LoginName -notlike "*#ext#*") -or ($spoUser.LoginName -notlike "*urn%3aspo%3aguest#*")))
                                        {                                                                                        
                                            $processFlag++
                                        }
                                        ElseIf ($spoUser.LoginName -like "*urn%3aspo%3aguest#*")
                                        {                                            
                                            $processFlag++
                                        }
                                        
                                        # This is the same export using the SharingLink class as the OneDrive section
                                        If ($processFlag)
                                        {
                                            Write-Host -ForegroundColor Cyan "$($spoUser.Title): $($pnpListItem.FileSystemObjectType): $($pnpListItem.FieldValues.FileRef)"
                                            $exportMe = [SharingLink]::new()
                                            $exportMe.SiteURL = $spoSite
                                            $exportMe.Title = $spoUser.Title
                                            $exportMe.LoginName = $spoUser.LoginName
                                            $exportMe.Email = $spoUser.Email
                                            $exportMe.InvitedOn = $getSharingLinkInvitedDate
                                            $exportMe.ObjectType = $pnpListItem.FileSystemObjectType
                                            $exportMe.Object = $pnpListItem.FieldValues.FileRef
                                            $exportMe.Permission = $getRoleDefinitionBindings

                                            $exportMe | EXport-Csv -NoTypeInformation -Append -Path ".\SharingLinks_Export-$getDate.csv"
                                        }
                                    }
                                }
                            }
                        }                    
                    }
                }                       
            }
        }            
    }  
}

# I thought about modifying this function to add the logic for the $externalInvitedFlag and $internalInvitedFlag but the whole point of the menu option is to get a top-level external user breakdown so left it alone
function getExternalUsers($getSPOSites)
{
    ForEach ($spoSite in $getSPOSites)
    {
        Connect-PnPOnline -Url $spoSite -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
        Write-Host -ForegroundColor Yellow "Compiling list of external users for $($spoSite)..."
        $getAllUsers = Get-PnPUser
        ForEach ($spoUser in $getAllUsers)
        {
            If (($spoUser.LoginName -like "*#ext#*") -or ($spoUser.LoginName -like "*urn%3aspo%3aguest#*"))
            {
                Write-Host -ForegroundColor Cyan "Found external user: $($spoUser.Email)"
                $exportMe = [ExternalUser]::new()
                $exportMe.SiteURL = $spoSite
                $exportMe.Title = $spoUser.Title
                $exportMe.LoginName = $spoUser.LoginName
                $exportMe.Email = $spoUser.Email                
            
                $exportMe | Export-Csv -NoTypeInformation -Append -Path ".\External_Users_Export-$getDate.csv"
            }
        }
    }
}

# Display our menu
showMenu

# Start a do/while loop that does a switch() on the integer value menu input
do
{
    [int]$menuResponse = $(Read-Host "Choice (16 to re-display menu; 17 to exit)")

    switch($menuResponse)
    {
        # Grab all SharePoint sites -- we omit the -IncludeOneDriveSites:$true from Get-PnPTenantSite to only get true SharePoint sites
        1
        {
            try
            {
                Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
                exit
            }

            Write-Host -ForegroundColor Magenta "Collecting all SharePoint sites in tenant..."
            $getSPOSites = Get-PnPTenantSite | Select-Object -ExpandProperty Url
            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) SharePoint sites!"

            # Check if the toggle to check SiteCollectionAdmin is set
            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # This time we include -IncludeOneDriveSites:$true but also -Filter on the return to ONLY get OneDrive sites
        2
        {            
            try
            {
                Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
                exit
            }

            Write-Host -ForegroundColor Magenta "Collecting all OneDrive sites in tenant..."
            $getSPOSites = Get-PnPTenantSite -IncludeOneDriveSites:$true -Filter "Url -like '-my.sharepoint.com/personal/'" | Select-Object -ExpandProperty Url
            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) OneDrive sites!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # We want both SharePoint/OneDrive so we let Get-PnPTenantSite rip!
        3
        {
            try
            {
                Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
                exit
            }

            Write-Host -ForegroundColor Magenta "Collecting all SharePoint & OneDrive sites in tenant..."
            $getSPOSites = Get-PnPTenantSite -IncludeOneDriveSites:$true | Select-Object -ExpandProperty Url
            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) SharePoint & OneDrive sites!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # Used this the most in testing -- just a single site input at the prompt
        4
        {
            $getSPOSites = $(Read-Host -Prompt "URL")
            Write-Host -ForegroundColor Cyan "Loading $($getSPOSites.Count) site!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # This would be used coupled with a Get-ADUser command such as: Get-ADUser -Filter "Department -like 'IT Oper*'" | Select-Object -ExpandProperty UserPrincipalName | ForEach-Object { $oneDriveUser = $_ -replace '[^a-zA-Z0-9]','_'; "https://odemail-my.sharepoint.com/personal/$($oneDriveUser)" }
        5
        {
            $csvFile = $(Read-Host -Prompt "CSV File (Must have Url header)")
            $importCsv = Import-Csv $csvFile

            $getSPOSites = $importCsv.Url

            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) sites in CSV!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }            
        # Toggle the check for SiteCollectionAdmin
        6
        {
            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Magenta "Setting siteCollectionAdminFlag to False..."
                $siteCollectionAdminFlag = 0
            }
            Else
            {
                Write-Host -ForegroundColor Magenta "Setting siteCollectionAdminFlag to True..."
                $siteCollectionAdminFlag = 1
            }

            showMenu
        }
        # This will blindly add whatever account as the SiteCollectionAdmin from a CSV without checking -- be careful!
        7
        {
            $csvFile = $(Read-Host -Prompt "CSV File (Must have Url header)")
            $importCsv = Import-Csv $csvFile

            $getSPOSites = $importCsv.Url

            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) sites in CSV! Loading..."
            Write-Host -ForegroundColor Yellow "This will add $($paAccount) as SiteCollectionAdmin to all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                addSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
        # Same as 7 but from loaded sites
        8
        {
            Write-Host -ForegroundColor Yellow "This will add $($paAccount) as SiteCollectionAdmin to all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                addSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
        # Remove the SiteCollectionAdmin from a CSV
        9
        {
            $csvFile = $(Read-Host -Prompt "CSV File (Must have Url header)")
            $importCsv = Import-Csv $csvFile

            $getSPOSites = $importCsv.Url

            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) sites in CSV! Loading..."
            Write-Host -ForegroundColor Yellow "This will remove $($paAccount) as SiteCollectionAdmin for all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                removeSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
       # Same as 9 but for loaded sites
       10
        {
            Write-Host -ForegroundColor Yellow "This will remove $($paAccount) as SiteCollectionAdmin for all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                removeSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
       # This toggle determines if we want to include the "*#ext#*" users 
       11
        {
            If ($externalInvitedFlag)
            {
                Write-Host -ForegroundColor Magenta "Setting ExternalInvitedFlag to False..."
                $externalInvitedFlag = 0
            }
            Else
            {
                Write-Host -ForegroundColor Magenta "Setting ExternalInvitedFlag to True..."
                $externalInvitedFlag = 1
            }

            showMenu
        }
       # This toggle determines if we want to include internal users 
       12
        {
            If ($internalInvitedFlag)
            {
                Write-Host -ForegroundColor Magenta "Setting InternalInvitedFlag to False..."
                $internalInvitedFlag = 0
            }
            Else
            {
                Write-Host -ForegroundColor Magenta "Setting siteCollectionAdminFlag to True..."
                $internalInvitedFlag = 1
            }

            showMenu
       }            
       # This does the top-level external users export
       13
        {
            Write-Host -ForegroundColor Yellow "This will export external users for all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                getExternalUsers $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
       # This does the SharingLinks export
       14
        {
            Write-Host -ForegroundColor Yellow "This will export SharingLinks for all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                getSharingLinks $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            } 
        }                                            
       # I used this in debugging and left it in because why not? :)
       15
        {            
            try
            {
                Remove-Item -Force -Path ".\SharingLinks_Export-$getDate.csv"
            }
            catch
            {
                Write-Host -ForegroundColor Red "SharingLinks_Export-$($getDate).csv doesn't exist!"
            }
            try
            {
                Remove-Item -Force -Path ".\SiteCollectionAdmin-$getDate.csv"
            }
            catch
            {
                Write-Host -ForegroundColor Red "SiteCollectionAdmin-$($getDate).csv doesn't exist!"
            }
            try
            {
                Remove-Item -Force -Path ".\External_Users_Export-$getDate.csv"
            }
            catch
            {
                Write-Host -ForegroundColor Red "External_Users_Export-$($getDate).csv doesn't exist!"
            }
        }
       # Re-display the menu
       16
        {
            showMenu
        }
    }     
}
while ($menuResponse -ne 17)  # Quit on 17!
# FIN