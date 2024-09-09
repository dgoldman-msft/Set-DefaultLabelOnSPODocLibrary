# Set-DefaultLabelOnSPODocLibrary
Set a sensitivity label on a SharePoint Library

## Getting Started with Set-DefaultLabelOnSPODocLibrary

This script will configure your tenant to allow Sites and Groups access for stamping sensitivity labels

### DESCRIPTION

You have a choice of selecting 'All' which will create 3 Hyper-V machines for a domain controller, windows server and a windows 11 client as well as a Kali Linux machine, just the 3 windows machines or just the 1 kali linux machine.

### Examples

- EXAMPLE PS Set-DefaultLabelOnSPODocLibrary -UserPrincipalName "admin@YourTenant.onmicrosoft.com" -TenantName YourTenant -SharePointSiteUrl YourSPOSite"

This will connect to the Security and Compliance center and SharePoint online, read all of the sensitivity labels and let you choose one to stamp on a SharePoint library.