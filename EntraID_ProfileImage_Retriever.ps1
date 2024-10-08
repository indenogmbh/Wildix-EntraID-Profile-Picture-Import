# Function to check if a module is installed, if not, install and import it
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName not found. Installing..."
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
    }
    
    Import-Module $ModuleName
    Write-Host "$ModuleName is installed and imported."
}

# Ensure necessary modules are installed and imported
$requiredModules = @("Az.Accounts", "Az.Storage", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users")

foreach ($module in $requiredModules) {
    Ensure-Module -ModuleName $module
}

# Disable WAM-based login for Azure
Update-AzConfig -EnableLoginByWam $false

# Authenticate with Microsoft Graph (only required scopes)
try {
    Connect-MgGraph -Scopes "User.Read.All", "ProfilePhoto.Read.All" -ErrorAction Stop
    Write-Host "Successfully authenticated with Microsoft Graph."
} catch {
    Write-Host "Error: Failed to authenticate with Microsoft Graph. Exiting script."
    exit
}

# Directory for temporary storage of profile pictures in the system's Temp directory
$profilePicFolder = [System.IO.Path]::Combine($env:TEMP, "ProfilePictures")
if (-not (Test-Path -Path $profilePicFolder)) {
    New-Item -Path $profilePicFolder -ItemType Directory
}

# Prompt the user for Azure Storage Account and Blob Container information
$storageAccountName = Read-Host "Please enter the Azure Storage Account name"
$resourceGroupName = Read-Host "Please enter the Resource Group name"
$containerName = Read-Host "Please enter the Blob Container name (e.g., 'profilepictures')"

# Authenticate with Azure and get the Storage Account context
Connect-AzAccount
$storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName

# Create the Blob service context
$ctx = $storageAccount.Context

# Create the container if it doesn't exist
$container = Get-AzStorageContainer -Context $ctx -Name $containerName -ErrorAction SilentlyContinue
if (-not $container) {
    $container = New-AzStorageContainer -Context $ctx -Name $containerName -Permission Off
    Write-Host "Blob container '$containerName' has been created."
} else {
    Write-Host "Blob container '$containerName' already exists."
}

# Generate a SAS token for the Blob container
$expiryDate = (Get-Date).AddYears(15)
$sasToken = New-AzStorageContainerSASToken -Context $ctx -Name $containerName -Permission r -ExpiryTime $expiryDate -Protocol HttpsOnly

Write-Host "SAS token generated: $sasToken"

# Retrieve all cloud users from Entra ID (Azure AD), filtering to only include members
$users = Get-MgUser -All -Filter "UserType eq 'Member'" -ConsistencyLevel eventual -CountVariable CountVar

# Loop through all filtered users and process their profile pictures
foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    Write-Host "Processing user: $upn"

    # Attempt to retrieve the profile picture
    try {
        $photoPath = [System.IO.Path]::Combine($profilePicFolder, "$upn.jpg")
        Get-MgUserPhotoContent -UserId $user.Id -outfile $photoPath -ErrorAction Stop
        Write-Host "Profile picture saved: $photoPath"

        # Upload the image directly to the Azure Blob container
        $blobPath = "$upn.jpg"
        $blob = Set-AzStorageBlobContent -File $photoPath -Container $containerName -Blob $blobPath -Context $ctx -ErrorAction Stop
        Write-Host "Profile picture uploaded to Azure Blob: $blobPath"

        # Delete the temporary file
        Remove-Item $photoPath -Force
    } catch {
        # Skip users without a profile picture
        if ($_.Exception -match "ImageNotFoundException") {
            Write-Host "No profile picture found for $($user.DisplayName), skipping user."
        } else {
            Write-Host "An error occurred while processing $($user.DisplayName): $_"
        }
    }
}

# Prompt user for CSV output path, default is the current script directory
$defaultCsvPath = [System.IO.Path]::Combine((Get-Location).Path, "profile_picture_urls.csv")
$csvPath = Read-Host "Please enter the output path for the CSV file (default: $defaultCsvPath)"
if (-not $csvPath) {
    $csvPath = $defaultCsvPath
}

# Base URL for the Azure Storage Blob (without file name)
$baseUrl = "https://$($storageAccountName).blob.core.windows.net/$containerName"

# List to store UPNs and URLs
$filesAndUrls = @()

# Retrieve all blobs in the container
$blobs = Get-AzStorageBlob -Container $containerName -Context $ctx

foreach ($blob in $blobs) {
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($blob.Name)
    
    # Replace "@" with "%40" in the blob name part of the URL only
    $encodedBlobName = $blob.Name -replace "@", "%40"
    $fullUrl = "$baseUrl/$encodedBlobName?$sasToken"

    # Create object with UPN and URL
    $obj = [PSCustomObject]@{
        UPN = $fileNameWithoutExtension
        URL = $fullUrl
    }

    # Add the object to the list
    $filesAndUrls += $obj
}

# Write the CSV file
$filesAndUrls | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "CSV file has been successfully created: $csvPath"

# Disconnect from Microsoft Graph and Azure
Disconnect-MgGraph
Disconnect-AzAccount
