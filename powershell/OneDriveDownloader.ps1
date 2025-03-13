# Import required libraries
Add-Type -AssemblyName "System.Windows.Forms"
Add-Type -AssemblyName "System.Drawing"
Add-Type -TypeDefinition @"
using System;
using System.Net.Http;
using System.Threading.Tasks;
"@

# Function to load configuration from JSON file
Function Load-Config {
    $configPath = "config.json"
    if (Test-Path -Path $configPath) {
        $config = Get-Content -Path $configPath | ConvertFrom-Json
        return $config
    }
    else {
        Write-Error "Config file not found!"
        return $null
    }
}

# Function to get OAuth access token using client credentials flow
Function Get-AccessToken {
    $config = Load-Config
    if ($config -eq $null) {
        Write-Error "Failed to load configuration. Cannot continue."
        return
    }

    $tenantId = $config.tenantId
    $clientId = $config.clientId
    $clientSecret = $config.clientSecret
    $scope = "https://graph.microsoft.com/.default"
    $url = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = $scope
    }

    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body
        return $response.access_token
    }
    catch {
        $errorMessage = "Authentication failed: $_"
        Write-Error $errorMessage
        return $errorMessage
    }
}

# Function to get the tenant's organization information
Function Get-TenantInfo {
    param (
        [string]$accessToken
    )

    $url = "https://graph.microsoft.com/v1.0/organization"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        return $response.value[0].displayName
    }
    catch {
        Write-Error "Failed to retrieve tenant information: $_"
        return "Unknown Tenant"
    }
}

# Function to show an error message box with detailed error
Function Show-ErrorMessage {
    param (
        [string]$message
    )

    [System.Windows.Forms.MessageBox]::Show($message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Function to get OneDrive items (files/folders) for a specific user
Function Get-OneDriveItems {
    param (
        [string]$email,
        [string]$accessToken,
        [string]$parentId = "root"
    )

    $url = "https://graph.microsoft.com/v1.0/users/$email/drive/items/$parentId/children"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }

    $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
    return $response.value
}

# Function to download a file
Function Download-File {
    param (
        [string]$downloadUrl,
        [string]$destination
    )

    $client = New-Object System.Net.Http.HttpClient
    $response = $client.GetAsync($downloadUrl).Result
    $fileStream = [System.IO.File]::Create($destination)
    $response.Content.CopyToAsync($fileStream).Wait()
    $fileStream.Close()
}

# Function to recursively download files and folders
Function Download-AllFiles {
    param (
        [string]$email,
        [string]$accessToken,
        [string]$parentId,
        [string]$downloadPath,
        [System.Windows.Forms.TextBox]$logTextBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        [ref]$totalItems
    )

    $items = Get-OneDriveItems -email $email -accessToken $accessToken -parentId $parentId

    $totalItems.Value = $items.Count

    if ($totalItems.Value -gt 0) {
        $currentItem = 0
        foreach ($item in $items) {
            $itemName = $item.name
            $itemId = $item.id
            $itemDownloadUrl = $item."@microsoft.graph.downloadUrl"
            $itemIsFolder = $item.folder -ne $null

            $currentItem++

            if ($totalItems.Value -gt 0) {
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, ($currentItem / $totalItems.Value) * 100))
            }

            if ($itemIsFolder) {
                $logTextBox.AppendText("Found folder: $itemName`r`n")
                $folderPath = Join-Path -Path $downloadPath -ChildPath $itemName
                if (-not (Test-Path -Path $folderPath)) {
                    New-Item -ItemType Directory -Path $folderPath
                }
                Download-AllFiles -email $email -accessToken $accessToken -parentId $itemId -downloadPath $folderPath -logTextBox $logTextBox -progressBar $progressBar -totalItems $totalItems
            }
            else {
                $logTextBox.AppendText("Downloading file: $itemName`r`n")
                $filePath = Join-Path -Path $downloadPath -ChildPath $itemName
                Download-File -downloadUrl $itemDownloadUrl -destination $filePath
            }
        }
    }
    else {
        $logTextBox.AppendText("No items to download.`r`n")
    }
}

# Function to zip all files in a directory
Function Zip-Folder {
    param (
        [string]$folderPath,
        [string]$zipPath
    )

    [System.IO.Compression.ZipFile]::CreateFromDirectory($folderPath, $zipPath)
}

# Create the Windows Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "OneDrive Downloader"
$form.Size = New-Object System.Drawing.Size(600, 550)  
$form.FormBorderStyle = 'FixedDialog'   
$form.MaximizeBox = $false              
$form.MinimizeBox = $false             

# Tenant Name Label
$tenantNameLabel = New-Object System.Windows.Forms.Label
$tenantNameLabel.Text = "Tenant: $(Get-TenantInfo -accessToken (Get-AccessToken))"
$tenantNameLabel.Location = New-Object System.Drawing.Point(10, 10)
$tenantNameLabel.Size = New-Object System.Drawing.Size(580, 20)
$form.Controls.Add($tenantNameLabel)

# User Email Input
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "Enter the user's email address:"
$emailLabel.Location = New-Object System.Drawing.Point(10, 40)
$emailLabel.Size = New-Object System.Drawing.Size(580, 20)  
$form.Controls.Add($emailLabel)

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Location = New-Object System.Drawing.Point(10, 60)
$emailTextBox.Width = 400
$form.Controls.Add($emailTextBox)

# Download Path Input
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Enter the download path:"
$pathLabel.Location = New-Object System.Drawing.Point(10, 100)
$pathLabel.Size = New-Object System.Drawing.Size(580, 20)  
$form.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(10, 130)
$pathTextBox.Width = 400
$form.Controls.Add($pathTextBox)

# Browse button for selecting the download folder
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse"
$browseButton.Location = New-Object System.Drawing.Point(420, 130)
$form.Controls.Add($browseButton)

$browseButton.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.ShowNewFolderButton = $true
        if ($folderDialog.ShowDialog() -eq 'OK') {
            $pathTextBox.Text = $folderDialog.SelectedPath
        }
    })

# Zip Output Path Input
$zipPathLabel = New-Object System.Windows.Forms.Label
$zipPathLabel.Text = "Select the output folder for the ZIP file:"
$zipPathLabel.Location = New-Object System.Drawing.Point(10, 170)
$zipPathLabel.Size = New-Object System.Drawing.Size(580, 20)  
$form.Controls.Add($zipPathLabel)

$zipPathTextBox = New-Object System.Windows.Forms.TextBox
$zipPathTextBox.Location = New-Object System.Drawing.Point(10, 200)
$zipPathTextBox.Width = 400
$form.Controls.Add($zipPathTextBox)

# Browse button for selecting the zip output folder
$zipBrowseButton = New-Object System.Windows.Forms.Button
$zipBrowseButton.Text = "Browse"
$zipBrowseButton.Location = New-Object System.Drawing.Point(420, 200)
$form.Controls.Add($zipBrowseButton)

$zipBrowseButton.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.ShowNewFolderButton = $true
        if ($folderDialog.ShowDialog() -eq 'OK') {
            $zipPathTextBox.Text = $folderDialog.SelectedPath
        }
    })

# Log Textbox to display download progress
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 250)  
$logTextBox.Size = New-Object System.Drawing.Size(560, 150)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = 'Vertical'
$form.Controls.Add($logTextBox)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 420)  
$progressBar.Size = New-Object System.Drawing.Size(560, 30)
$form.Controls.Add($progressBar)

# Create the Download and Zip button
$downloadAndZipButton = New-Object System.Windows.Forms.Button
$downloadAndZipButton.Text = "Download and Zip"
$downloadAndZipButton.Location = New-Object System.Drawing.Point(10, 460)  
$form.Controls.Add($downloadAndZipButton)

# Button Events
$downloadAndZipButton.Add_Click({
        # Clear previous logs
        $logTextBox.Clear()
        $progressBar.Value = 0

        # Get the access token
        $accessToken = Get-AccessToken

        # Get the OneDrive items
        $email = $emailTextBox.Text
        $downloadPath = $pathTextBox.Text
        $zipOutputPath = $zipPathTextBox.Text

        if (![string]::IsNullOrEmpty($email) -and ![string]::IsNullOrEmpty($downloadPath) -and ![string]::IsNullOrEmpty($zipOutputPath)) {
            # Log starting message
            $logTextBox.AppendText("Starting download process...`r`n")

            # Create a reference variable for total items
            $totalItems = 0

            # Download all files from OneDrive recursively
            Download-AllFiles -email $email -accessToken $accessToken -parentId "root" -downloadPath $downloadPath -logTextBox $logTextBox -progressBar $progressBar -totalItems ([ref]$totalItems)

            # Log completion message
            $logTextBox.AppendText("Download complete. Zipping files...`r`n")

            # Get today's date
            $todayDate = Get-Date -Format "yyyy-MM-dd"

            # Generate the zip file path with today's date
            $zipFileName = "$($email)_OneDrive_Backup_$todayDate.zip"
            $zipPath = Join-Path -Path $zipOutputPath -ChildPath $zipFileName

            # Zip the downloaded folder
            Zip-Folder -folderPath $downloadPath -zipPath $zipPath

            # Log success
            $logTextBox.AppendText("Files downloaded and zipped successfully!`r`n")
        }
        else {
            $logTextBox.AppendText("Please fill in all fields before proceeding.`r`n")
        }
    })

# Show the form
$form.ShowDialog()
