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
$requiredModules = @("Az.Accounts", "Az.Storage", "Microsoft.Graph")

foreach ($module in $requiredModules) {
    Ensure-Module -ModuleName $module
}

# Authenticate with Microsoft Graph
Connect-MgGraph -Scopes "ProfilePhoto.ReadWrite.All, User.Read.All, User.ReadWrite.All"

# Directory for temporary storage of profile pictures in the system's Temp directory
$profilePicFolder = [System.IO.Path]::Combine($env:TEMP, "ProfilePictures")
if (-not (Test-Path -Path $profilePicFolder)) {
    New-Item -Path $profilePicFolder -ItemType Directory
}

# Prompt the user for Azure Storage Account and Blob Container information
$storageAccountName = Read-Host "Please enter the Azure Storage Account name"
$resourceGroupName = Read-Host "Please enter the Resource Group name"
$containerName = Read-Host "Please enter the Blob Container name (e.g., 'profilepictures')"

# Prompt the user for the folder name inside the Blob container (default: 'profilepictures')
$blobFolderName = Read-Host "Please enter the folder name inside the Blob container (default: 'profilepictures')" 
if (-not $blobFolderName) {
    $blobFolderName = "profilepictures"
}

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

# Generate a SAS token for the folder in the Blob container
$expiryDate = (Get-Date).AddYears(15)
$sasToken = New-AzStorageContainerSASToken -Context $ctx -Name $containerName -Permission r -ExpiryTime $expiryDate -Protocol HttpsOnly

Write-Host "SAS token generated: $sasToken"

# Retrieve all users from Entra ID (Azure AD)
$users = Get-MgUser -All

# Loop through all users and process their profile pictures
foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    Write-Host "Processing user: $upn"

    # Attempt to retrieve the profile picture
    try {
        $photoPath = [System.IO.Path]::Combine($profilePicFolder, "$upn.jpg")
        Get-MgUserPhotoContent -UserId $user.Id -outfile $photoPath -ErrorAction Stop
        Write-Host "Profile picture saved: $photoPath"

        # Upload the image to the specified Azure Blob folder
        $blobPath = "$blobFolderName/$upn.jpg"
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

# Disconnect from Microsoft Graph and Azure
Disconnect-MgGraph
Disconnect-AzAccount
