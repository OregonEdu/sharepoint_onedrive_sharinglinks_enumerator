# SharePoint/OneDrive SharingLinks Enumerator - v2.0
Benjamin Barshaw <<benjamin.barshaw@ode.oregon.gov>> - IT Operations & Support Network Team Lead - Oregon Department of Education

Requirements: PnP PowerShell Module
              SharePoint Admin  

This script will export by default all true external "guest" users for a SiteCollection or all external SharingLinks for a SiteCollection(s) within an M365 tenant. I have added options to toggle INVITED external users (users with #ext# 
in their UPN in Entra) as well as internal users to your tenant. In v1.0 of the script the OneDrive SharingLinks were built out from Get-PnPUser because the Groups the user object returns contain references to the SharingLink 
themselves. Then during testing I made the unfortunate finding that when using the built-in functionality to "Send" an e-mail with the share it does NOT generate a SharingLink! I have modified the way OneDrive shares are handled in
this version of the script to accomodate for this but without a SharingLink it is impossible to determine when the link was shared. For these types the "InvitedOn" column is simply set to "N/A". For SharePoint sites I have the libraries
you want to export set in the USER-DEFINED VARIABLES section below which you can add/remove to suit your needs.  This script came about by me noticing thata SharingLink came in a format of (for example):

SharingLinks.af100805-bd06-4c03-b9a0-f9506f6e8d57.Flexible.47d674bc-3399-41d3-a5fb-f84b2c04df52

By digging through the PnP cmdlets I was able to determine that the first long hexadecimal string correlated with the UniqeId of a file in a Document library. So it is the file/folder being shared! In further digging I was able
to determine that the second long hexadecimal string correlated to the SharingLink itself and you could do lookups on each by using values embedded in each object. This was not fun nor easy. :)

It should be noted that even with SharePoint Admin you need to manually add your PA account as a SiteCollectionAdmin -- SharePoint Admin does NOT give you carte blanche.

To make it easier to delineate between the comments and actual code, PowerShell ISE or Visual Studio Code is recommended for editing/reading.

If you have any questions/comments, please feel free to reach out to me via Teams or e-mail.

P.S. - One of my closest personal friends Sean McArdle makes fun of how "ANSI C" my PowerShell is. I learned programming from "Programming in ANSI C" by Stephen Kochan -- so this checks out! :)
