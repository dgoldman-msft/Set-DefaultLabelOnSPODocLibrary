function Set-DefaultLabelOnSPODocLibrary {
    <#
        .SYNOPSIS
            Configures a default sensitivity label for a SharePoint document library.

        .DESCRIPTION
            This function installs necessary modules, connects to Exchange Online and SharePoint Online, creates a sensitivity label if it doesn't exist, and sets it as the default for a specified SharePoint site.

        .PARAMETER CreateLabel
            Indicate that you want to create a default label

        .PARAMETER DisableNameChecking
            Turn off PowerShell import cmdlet warnings

        .PARAMETER LogFile
            Location for logging file

        .PARAMETER SPOSiteName
            The name of the SharePoint site where the default sensitivity label will be set.

        .PARAMETER TenantName
            Your Entra tenant name

        .PARAMETER UserPrincipalName
            The user principal name (UPN) of the admin account.

        .EXAMPLE
            Set-DefaultSensitivityLabel -UserPrincipalName "admin@YourTenant.onmicrosoft.com" -TenantName YourTenant -SharePointSiteUrl YourSPOSite"

        .NOTES
            https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-default-label
            https://learn.microsoft.com/en-us/purview/set-up-irm-in-sp-admin-center#irm-enable-sharepoint-document-libraries-and-lists
            https://learn.microsoft.com/en-us/purview/apply-irm-to-a-list-or-library
            https://learn.microsoft.com/en-us/compliance/assurance/assurance-data-classification-and-labels
            Ensure you have the necessary permissions to run this script.
    #>

    [CmdletBinding(SupportsShouldProcess = $True)]
    param (
        [bool]
        $DisableNameChecking = $True,

        [Parameter(Mandatory = $false)]
        [string]
        $LogFile = "C:\Logs\Set-DefaultSensitivityLabel.log",

        [Parameter(Mandatory = $true)]
        [string]
        $SPOSiteName,

        [Parameter(Mandatory = $true)]
        [string]
        $TenantName,

        [Parameter(Mandatory = $true)]
        [string]
        $UserPrincipalName
    )

    $modules = @('ExchangeOnlineManagement', 'Microsoft.Online.SharePoint.PowerShell')

    # Error logging setup
    Write-Output "Setting up logging"
    if (-not (Test-Path -Path (Split-Path -Path $LogFile))) {
        New-Item -ItemType Directory -Path (Split-Path -Path $LogFile) -Force
    }
    Start-Transcript -Path $LogFile -Append

    try {
        # Check for an install necessary modules
        foreach ($module in $modules) {
            Write-Output "Checking for $module"
            if (-NOT (Get-Module -Name $module -ListAvailable | Select-Object -First 1)) {
                Write-Output "Installing $module"
                Install-Module -Name $module -Force -ErrorAction Stop
            }
            else {
                Write-Verbose "$module already installed! importing module"
                Import-Module -Name $module -DisableNameChecking:$DisableNameChecking -ErrorAction Stop
            }
        }
    }
    catch {
        Write-Output "A logging error occurred: $_. Exiting!"
    }

    try {
        # Connect to Exchange and SharePoint Online
        Write-Output "Connecting to Exchange Online"
        Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
        Write-Output "Connecting to SharePoint Online"
        Connect-SPOService -Url "https://$TenantName-admin.sharepoint.com/" -ErrorAction Stop
    }
    catch {
        Write-Error "An connection error occurred: $_"
    }

    try {
        Write-Output "Checking to see if AIP Integration has been enabled for tenant $($TenantName)."
        if (-Not (Get-SPOTenant).EnableAIPIntegration) {
            $response = Read-Host "AIP Integration is disabled in your tenant $($TenantName) for SharePoint Online and OneDrive for Business. Would you like to enable it? (y/n)"
            if ($response -eq "Y" -or $response -eq "y") {
                Set-SPOTenant -EnableAIPIntegration $true -ErrorAction Stop

                if (-Not (Get-SPOTenant).EnableAIPIntegration) {
                    Throw "EnableAIPIntegration failed to set. Please check the logs for more information."
                }
                else {
                    Write-Output "AIP Integration has been enabled in tenant $($TenantName)."
                }
            }
            else {
                Write-Output "AIP Integration remains disabled. Exiting"
                return
            }
        }

        Write-Output "Checking to see if EnableSensitivityLabelforPDF support has been enabled for tenant $($TenantName)."
        if (-Not (Get-SPOTenant).EnableSensitivityLabelforPDF ) {
            $response = Read-Host "EnableSensitivityLabelforPDF support is disabled in your tenant $($TenantName) for SharePoint Online and OneDrive for Business. Would you like to enable it? (y/n)"
            if ($response -eq "Y" -or $response -eq "y") {
                Set-SPOTenant -EnableSensitivityLabelforPDF $true -ErrorAction Stop

                if (-Not (Get-SPOTenant).EnableSensitivityLabelforPDF ) {
                    Throw "EnableSensitivityLabelforPDF failed to set. Please check the logs for more information."
                }
                else {
                    Write-Output "EnableSensitivityLabelforPDF has been enabled in tenant $($TenantName)."
                }
            }
            else {
                Write-Output "EnableSensitivityLabelforPDF remains disabled. Exiting"
                return
            }
        }
    }
    catch {
        Write-Output "API integration error occurred. Error: $_"
        return
    }

    try {
        # Get labels
        Write-Output "Getting labels from the $($TenantName)"
        $labelArray = @()
        $counter = 1
        $labels = Get-Label
        if ($labels.Count -eq 0) {
            Write-Output "No labels found in $($TenantName). Unable to set sensitivity label on SPOSite: $($SPOSiteName). You must create a label in the Security and Compliance center before you can use this feature."
            return
        }
        else {
            Write-Output "\nLabels found"
            foreach ($label in $labels) {
                $labelArray += $label
                Write-output "[$counter] - $($label.DisplayName) | Content Type: $($label.ContentType)"
                $counter++
            }

            $choice = Read-Host -Prompt "Your choice? (Enter the number corresponding to the label, Default = Exit)"
            if ($choice -match '^\d+$' -and [int]$choice -le $labels.Count) {
                $selectedLabel = $labelArray[[int]$choice - 1]

                # Get the label ID
                Write-Output "Getting $($selectedLabel.DisplayName)'s identifier"
                $selectedLabelId = (Get-Label -Identity $selectedLabel).Id
            }
            else {
                Write-Output "Invalid choice or exit selected."
                return
            }
        }
    }
    catch {
        Write-Error "A label error occurred: $_"
    }

    try {
        Write-Output "You selected: $($selectedLabel.DisplayName). Setting this label on the following SPOSite: $($SPOSiteName)"
        Set-SPOSite -Identity "https://$TenantName.sharepoint.com/sites/$SPOSiteName" -SensitivityLabel $selectedLabelId -ErrorAction Stop
    }
    catch {
        Write-Error "An error occurred: $_"
    }

    # Stop Logging and disconnect from exchange
    Stop-Transcript
}