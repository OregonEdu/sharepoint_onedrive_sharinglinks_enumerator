# SharePoint/OneDrive SharingLinks Enumerator - v1.5
Benjamin Barshaw <<benjamin.barshaw@ode.oregon.gov>> - IT Operations & Support Network Team Lead - Oregon Department of Education

Requirements: PnP PowerShell Module & SharePoint Admin  

This script will export by default all true external "guest" users for a SiteCollection or all external SharingLinks for a SiteCollection(s) within a tenant. I have added options to toggle INVITED external users (users with #ext# in
their UPN in Entra) as well as internal users to your tenant. For OneDrive SharingLinks we are able to build out our report from Get-PnPUser because the Groups the user object returns contain references to the SharingLink themselves. 
This is NOT the case for SharePoint sites proper and requires you build out from the documents/pages/etc. The libraries you want defined are below in the user defined variables section. This script came about by me noticing that
a SharingLink came in a format of (for example):

SharingLinks.af100805-bd06-4c03-b9a0-f9506f6e8d57.Flexible.47d674bc-3399-41d3-a5fb-f84b2c04df52

By digging through the PnP cmdlets I was able to determine that the first long hexadecimal string correlated with the UniqeId of a file in a Document library. So it is the file/folder being shared! In further digging I was able
to determine that the second long hexadecimal string correlated to the SharingLink itself and you could do lookups on each by using values embedded in each object. This was not fun nor easy. :)

It should be noted that even with SharePoint Admin you need to manually add your PA account as a SiteCollectionAdmin -- SharePoint Admin does NOT give you carte blanche.

P.S. - One of my closest personal friends Sean McArdle makes fun of how "ANSI C" my PowerShell is. I learned programming from "Programming in ANSI C" by Stephen Kochan -- so this checks out! :)
