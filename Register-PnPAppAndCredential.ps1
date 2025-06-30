<#
.SYNOPSIS
    Registers an Azure AD application and stores its credential for PnP.PowerShell automation.

.DESCRIPTION
    This script automates the registration of an Azure AD application, creates a client secret, and stores the credential in the Windows Credential Manager for use with PnP.PowerShell.
    It checks and installs required modules, allows overwriting existing applications, and validates the entire process.
    The script does NOT assign permissions to the application—this must be done separately if needed.

.PARAMETER AppName
    The name of the Azure AD application to register. Default: "PnP-Automation".

.PARAMETER TenantName
    The tenant name (e.g., contoso) where the application will be registered.

.PARAMETER ForceReplace
    Switch. If set, removes any existing application with the same name before creating a new one.

.EXAMPLE
    .\Register-PnPAppAndCredential.ps1 -TenantName "contoso"

.EXAMPLE
    .\Register-PnPAppAndCredential.ps1 -AppName "MyApp" -TenantName "contoso" -ForceReplace

.INPUTS
    [String] Application name and tenant name.

.OUTPUTS
    [PSCustomObject] Information about the created application and credential.

.NOTES
    Author      : Kenlar Taj
    Version     : 1.0
    License     : GNU General Public License v3.0
    Date        : 2025-06-30
    Tested on   : Windows 11 Pro 26100.4349 with PowerShell 7.5.2
    Repository  : https://github.com/KenlarTaj/powershel
    Contact     : kenlartaj@gmail.com
    Requirements: Permission to register apps in Azure AD and install PowerShell modules.
                  Powershell 7.4+ (required for PnP.PowerShell).
    Note        : The script uses Write-Host for colored messages to improve interactive experience.
#>
param(
    [string]$AppName = "PnP-Automation",
    [Parameter(Mandatory)]
    [string]$TenantName,
    [switch]$ForceReplace
)

function Write-Msg {
    param(
        [string]$Type = "INFO",
        [switch]$BreakLine = $false,
        [string]$Msg
    )
    if ($BreakLine) {
        Write-Host ""
    }
    switch ($Type) {
        "INFO" { Write-Host "[$Type] $Msg" -ForegroundColor Cyan }
        "SUCC" { Write-Host "[$Type] $Msg" -ForegroundColor Green }
        "WARN" { Write-Host "[$Type] $Msg" -ForegroundColor Yellow }
        "ERRO" { Write-Host "[$Type] $Msg" -ForegroundColor Red }
        default { Write-Host "[$Type] $Msg" -ForegroundColor Cyan }
    }
}

function Get-ResourceAppName {
    param([string]$ResourceAppId)
    $sp = Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -ErrorAction SilentlyContinue
    if ($sp) { return $sp.DisplayName }
    else { return $ResourceAppId }
}

function Get-ResourceAccessName {
    param(
        [string]$ResourceAppId,
        [string]$ResourceAccessId
    )
    $sp = Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -ErrorAction SilentlyContinue
    if ($sp) {
        $perm = $sp.AppRoles | Where-Object { $_.Id -eq $ResourceAccessId }
        if ($perm) { return $perm.Value }
        $perm = $sp.Oauth2PermissionScopes | Where-Object { $_.Id -eq $ResourceAccessId }
        if ($perm) { return $perm.Value }
    }
    return $ResourceAccessId
}

function Register-PSGalleryIfNeeded {
    Write-Msg -Type "INFO" -BreakLine -Msg "Checking if PowerShell Gallery is registered..."
    # Check if the PSGallery repository is already registered
    $errorGallery = $false
    try {
        $psGallery = Get-PSRepository -Name "PSGallery" -ErrorAction Stop
    } catch {
        $errorGallery = $true
    }
    if ($errorGallery) {
        Write-Msg -Type "WARN" -Msg "Error trying to get PSGallery repository. Checking PowerShellGet module..."
        try {
            $psGetModule = Get-Module -Name "PowerShellGet" -ListAvailable -ErrorAction Stop
        } catch {
            throw "Error retrieving PowerShellGet module: $($_.Exception.Message)"
        }
        if (-not $psGetModule) {
            throw "PowerShellGet module is required to register the PowerShell Gallery repository. Please install it before proceeding."
        }
        else {
            Write-Msg -Type "SUCC" -Msg "PowerShellGet module is installed."
            Write-Msg -Type "INFO" -Msg "Importing PowerShellGet module..."
            try {
                Import-Module -Name "PowerShellGet" -Force -ErrorAction Stop
                Write-Msg -Type "SUCC" -Msg "PowerShellGet module imported successfully."
            } catch {
                throw "Error importing PowerShellGet module: $($_.Exception.Message)"
            }
        }

        try {
            $psGallery = Get-PSRepository -Name "PSGallery" -ErrorAction Stop
        } catch {
            throw "Error trying to get PSGallery repository. Cannot proceed with registration."
        }
    }
    if (-not $psGallery) {
        Write-Msg -Type "WARN" -Msg "PowerShell Gallery repository is not registered. Proceeding to register it..."
        try {
            Register-PSRepository -Default -ErrorAction Stop
            Write-Msg -Type "SUCC" -Msg "PowerShell Gallery repository registered successfully."
        } catch {
            throw "Failed to register PowerShell Gallery repository: $($_.Exception.Message)"
        }
    }
    if ($psGallery) {
        Write-Msg -Type "SUCC" -Msg "PowerShell Gallery repository is already registered."
        return
    }
}
function Install-ModuleIfNeeded {
    param (
        [string]$ModuleName
    )
    try {
        Write-Msg -Type "INFO" -BreakLine -Msg "Checking if module '$ModuleName' is installed..."
        $module = Get-Module -ListAvailable -Name $ModuleName
        if ($module) {
            Write-Msg -Type "SUCC" -Msg "Module '$ModuleName' is installed."
        } else {
            Write-Msg -Type "WARN" -Msg "Module '$ModuleName' is not installed." 
            # Search for the module in PowerShell Gallery
            Write-Msg -Type "INFO" -Msg "Searching for module '$ModuleName' in PowerShell Gallery..."
            $moduleSearch = Find-Module -Name $ModuleName -ErrorAction Stop
            if (-not $moduleSearch) {
                # If the module is not found, throw an error
                throw "Module '$ModuleName' not found in PowerShell Gallery."
            } else {
                # If the module is found, proceed with installation
                Write-Msg -Type "SUCC" -Msg "Module '$ModuleName' found in PowerShell Gallery."
                Write-Msg -Type "INFO" -Msg "Installing module '$ModuleName'..."
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -ErrorAction Stop
                Write-Msg -Type "SUCC" -Msg "Module '$ModuleName' installed successfully."

                # Import the module after installation
                Write-Msg -Type "INFO" -Msg "Importing module '$ModuleName'..."
                Import-Module -Name $ModuleName
                Write-Msg -Type "SUCC" -Msg "Module '$ModuleName' imported successfully."
                
                # Verify if the module is available
                Write-Msg -Type "INFO" -Msg "Verifying module '$ModuleName'..."
                if (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue) {
                    Write-Msg -Type "SUCC" -Msg "Module '$ModuleName' is available for use."
                } else {
                    throw "Module '$ModuleName' could not be imported."
                }
            }
        }
        return
    } catch {
        throw "Error checking or installing module '$ModuleName': $($_.Exception.Message)"
    }
}

function Get-ApplicationInfo {
    param(
        [string]$AppName
    )
    try {
        Write-Msg -Type "INFO" -BreakLine -Msg "Checking if the application '$AppName' already exists..."
        $app = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction Stop
        if ($app) {
            $appPermissions = $app.RequiredResourceAccess
            if (-not $appPermissions) {
                Write-Msg -Type "SUCC" -Msg "Application '$AppName' already exists with ID $($app.AppId), no permissions assigned."
            }
            else {
                $permissionObjects = @()
                foreach ($perm in $appPermissions) {
                    $apiName = Get-ResourceAppName $perm.ResourceAppId
                    foreach ($ra in $perm.ResourceAccess) {
                        $permName = Get-ResourceAccessName $perm.ResourceAppId $ra.Id
                        $permissionObjects += [PSCustomObject]@{
                            ApiName   = $apiName
                            PermName  = $permName
                            ApiId     = $perm.ResourceAppId
                            PermId    = $ra.Id
                            Type      = $ra.Type
                        }
                    }
                }
                Write-Msg -Type "SUCC" -Msg "Application '$AppName' already exists with ID $($app.AppId) and the following permissions:`n"
                $permissionObjects | Format-Table -AutoSize | Out-Host
           }
            return $true
        } else {
            return $false
        }
    } catch {
        throw "Error retrieving application '$AppName': $($_.Exception.Message)"
    }
}

function Install-Application {
    param(
        [string]$AppName
    )
    try {
        Write-Msg -Type "INFO" -BreakLine -Msg "Registering new Azure AD application '$AppName'..."
        $newApp = New-MgApplication -DisplayName $AppName
        if (-not $newApp) {
            throw "Failed to register Azure AD application '$AppName'."
        }
        Write-Msg -Type "SUCC" -Msg "Application '$AppName' registered successfully with ID $($newApp.AppId)."

        $newClientId = $newApp.AppId
        # Create a client secret
        Write-Msg -Type "INFO" -BreakLine -Msg "Creating client secret for the application..."
        $newSecret = Add-MgApplicationPassword -ApplicationId $newApp.Id
        if (-not $newSecret) {
            throw "Failed to create client secret for application '$AppName'."
        }
        $newSecureSecret = ConvertTo-SecureString $newSecret.SecretText -AsPlainText -Force
        $newSecretId = $newSecret.KeyId
        Write-Msg -Type "SUCC" -Msg "Client secret created successfully for application '$AppName'."
        
        Write-Msg -Type "INFO" -BreakLine -Msg "Verifying the application registration..."
        $app = Get-MgApplication -ApplicationId $newApp.Id
        if (-not $app) {
            throw "Failed to retrieve the registered application '$AppName'."
        }
        $secret = $app.PasswordCredentials | Where-Object { $_.KeyId -eq $newSecret.KeyId }
        if (-not $secret) {
            throw "Failed to retrieve the client secret for application '$AppName'."
        }
        Write-Msg -Type "SUCC" -Msg "Application '$AppName' registered successfully."

        Write-Msg -Type "INFO" -Msg "Client ID: '$newClientId'"
        Write-Msg -Type "INFO" -Msg "Secret ID: '$newSecretId'.
        "
        return [PSCustomObject]@{
            ClientId     = $newClientId
            SecretId     = $newSecretId
            SecureSecret = $newSecureSecret
        }
    } catch {
        throw "Error registering application '$AppName': $($_.Exception.Message)"
    }
}

function Remove-Application {
    param(
        [string]$AppName
    )
    try {
        Write-Msg -Type "INFO" -BreakLine -Msg "Removing application '$AppName'..."
        $app = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction Stop
        if (-not $app) {
            throw "Application '$AppName' not found."
        }
        Remove-MgApplication -ApplicationId $app.Id -Confirm:$false
        Write-Msg -Type "SUCC" -Msg "Application '$AppName' removed successfully."
    } catch {
        throw "Error removing application '$AppName': $($_.Exception.Message)"
    }
}
function Undo-Credential {
    param(
        [string]$AppName,
        [string]$TenantName
    )

    try {
        $CredentialName = "https://${TenantName}-admin.sharepoint.com"
        Write-Msg -Type "INFO" -BreakLine -Msg "Rolling back credential installation for application '$AppName'..."
        Write-Msg -Type "INFO" -Msg "Checking if the credential '$CredentialName' exists for cleanup..."
        # If the credential exists, remove it
        if(Get-PnPStoredCredential -Name $CredentialName) {
            Write-Msg -Type "INFO" -Msg "Credential '$CredentialName' exists. Proceeding to remove it..."
            Remove-PnPStoredCredential -Name $CredentialName -ErrorAction SilentlyContinue
            Write-Msg -Type "SUCC" -Msg "Credential '$CredentialName' cleaned up successfully."
        } else {
            Write-Msg -Type "INFO" -Msg "Credential '$CredentialName' does not exist, no cleanup needed."
        }
    } catch {
        throw "Error during rollback of credential installation: $($_.Exception.Message)"
    }
}
function Install-Credential {
    param(
        [string]$ClientId,
        [SecureString]$ClientSecret
    )

    try {
        if ([string]::IsNullOrWhiteSpace($ClientId)) {
            throw "Client ID is null, empty, or whitespace."
        }
        if (-not $ClientSecret) {
            throw "Secure string for client secret is null."
        }
        $CredentialName = "https://${TenantName}-admin.sharepoint.com"
        Write-Msg -Type "INFO" -BreakLine -Msg "Trying to install credential '$CredentialName' in Windows Credential Manager..."
        Write-Msg -Type "INFO" -Msg "Creating PSCredential object for client ID '$ClientId'..."
        $newcredential = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
        if (-not $newcredential) {
            throw "Failed to create PSCredential object."
        }
        Write-Msg -Type "SUCC" -Msg "PSCredential object created successfully."
        Write-Msg -Type "INFO" -Msg "Adding credential '$CredentialName' to PnPStoredCredential..."
        # Check if the credential already exists
        if ((Get-PnPStoredCredential -Name $CredentialName -ErrorAction SilentlyContinue)) {
            Write-Msg -Type "WARN" -Msg "Credential '$CredentialName' already exists. Trying to remove it."
            Remove-PnPStoredCredential -Name $CredentialName
            Write-Msg -Type "SUCC" -Msg "Credential '$CredentialName' removed successfully."
        }
        # Create the new credential
        Write-Msg -Type "INFO" -BreakLine -Msg "Credential '$CredentialName' does not exist. Proceeding to create it."
        Write-Msg -Type "INFO" -Msg "Creating new credential '$CredentialName'..."
        Add-PnPStoredCredential -Name $CredentialName -Username $ClientId -Password $ClientSecret
        Write-Msg -Type "SUCC" -Msg "Credential '$CredentialName' created successfully."
        Write-Msg -Type "INFO" -Msg "Verifying the saved credential '$CredentialName'..."
        $testCredential = Get-PnPStoredCredential -Name $CredentialName
        if (-not $testCredential) {
            throw "Credential '$CredentialName' not found."
        } else {
            Write-Msg -Type "INFO" -Msg "Credential '$CredentialName' found."
            if ($newcredential.UserName -ne $testCredential.UserName -or
                $newcredential.GetNetworkCredential().Password -ne $testCredential.GetNetworkCredential().Password) {
                throw "Stored credential $CredentialName does not match the expected values."
            } else {
                Write-Msg -Type "INFO" -Msg "Stored credential matches the expected values."
            }
        }
        Write-Msg -Type "SUCC" -Msg "Credential '$CredentialName' created and verified successfully."
    } catch {
        throw "Error creating or storing credential '$CredentialName': $($_.Exception.Message)"
    }

    try {
        Write-Msg -Type "INFO" -BreakLine -Msg "Disposing of the secure string for client secret..."
        $ClientSecret.Dispose()
        Write-Msg -Type "SUCC" -Msg "Secure string disposed successfully."
    } catch {
        Write-Msg -Type "WARN" -Msg "Error disposing secure string for client secret: $($_.Exception.Message)"
    }
}
function Main {
    param(
        [string]$AppName,
        [string]$TenantName
    )

    if (-not $AppName) {
        throw "Application name is required."
    }
    if (-not $TenantName) {
        throw "Tenant name is required."
    }

    Write-Msg -Type "INFO" -BreakLine -Msg "Starting the registration process for application '$AppName' in tenant '$TenantName'..."

    try {
        Install-ModuleIfNeeded -ModuleName "PowerShellGet"
        Install-ModuleIfNeeded -ModuleName "Microsoft.Graph.Applications"
        Install-ModuleIfNeeded -ModuleName "PnP.PowerShell"
    } catch {
        throw $($_.Exception.Message)
    }

    try {
        Write-Msg -Type "INFO" -BreakLine -Msg "Authenticating to Microsoft Graph with user consent..."
        Connect-MgGraph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","Directory.ReadWrite.All" -NoWelcome
        Write-Msg -Type "SUCC" -Msg "Connected to Microsoft Graph successfully."
    } catch {
        throw "Error connecting to Microsoft Graph: $($_.Exception.Message)"
    }

    # Check if the application already exists
    try {
        $appExists = Get-ApplicationInfo -AppName $AppName
    } catch {
        throw $($_.Exception.Message)
    }

    if ($appExists) {
        if ($ForceReplace) {
            Write-Msg -Type "WARN" -BreakLine -Msg "Parameter -ForceReplace is set! Proceeding to remove application '$AppName'."
            try {
                Remove-Application -AppName $AppName
            } catch {
                throw "Error removing application '$AppName': $($_.Exception.Message)"
            }
        } else {
            Write-Msg -Type "INFO" -Msg "You can use this application for your automation scripts."
            return
        }
    }

    Write-Msg -Type "WARN" -BreakLine -Msg "Application '$AppName' does not exist. Proceeding to create a new application."

    try {
        $app = Install-Application -AppName $AppName
    } catch {
        throw $($_.Exception.Message)
    }

    if (-not $app) {
        throw "Failed to create application '$AppName'."
    } else {
        $ClientId = $app.ClientId
        $ClientSecret = $app.SecureSecret
    }

    try {
        Install-Credential -ClientId $ClientId -ClientSecret $ClientSecret

        $ClientSecret.Dispose()
    } catch {
        Undo-Credential -AppName $AppName -TenantName $TenantName
        throw $($_.Exception.Message)
    }
    $CredentialName = "https://${TenantName}-admin.sharepoint.com"
    Write-Msg -Type "SUCC" -BreakLine -Msg "Credential '$CredentialName' is ready for use in your automation scripts!"
    Write-Msg -Type "SUCC" -BreakLine -Msg "You can now use the following commands in your scripts:"
    Write-Host -ForegroundColor Yellow @"

`$tenantURL    = "$CredentialName"
`$credential   = Get-PnPStoredCredential -Name `$tenantURL
`$clientID     = `$credential.UserName
`$clientSecret = `$credential.GetNetworkCredential().Password
"@
}

Write-Msg -Type "INFO" -BreakLine -Msg "----- START -----"

# Run the main function with provided parameters
try {
    Main -AppName $AppName -TenantName $TenantName
} catch {
    throw $($_.Exception.Message)
}

Write-Msg -Type "INFO" -BreakLine -Msg "----- END -----"