# Import required libraries
Add-Type -AssemblyName "System.Windows.Forms"
Add-Type -AssemblyName "System.Drawing"
Add-Type -TypeDefinition @"
using System;
using System.Net.Http;
using System.Threading.Tasks;
"@

# Function to get OAuth access token using client credentials flow
Function Get-AccessToken {
    $tenantId = ""
    $clientId = ""
    $clientSecret = ""
    $scope = "https://graph.microsoft.com/.default"
    $url = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = $scope
    }
    
    $response = Invoke-RestMethod -Uri $url -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body
    return $response.access_token
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

    # Initialize totalItems only when items are retrieved
    $totalItems.Value = $items.Count

    # Ensure the totalItems is greater than 0 before starting the download
    if ($totalItems.Value -gt 0) {
        $currentItem = 0
        foreach ($item in $items) {
            $itemName = $item.name
            $itemId = $item.id
            $itemDownloadUrl = $item."@microsoft.graph.downloadUrl"
            $itemIsFolder = $item.folder -ne $null
            
            $currentItem++
            
            # Update progress bar only if totalItems is greater than 0
            if ($totalItems.Value -gt 0) {
                # Clamp the value between 0 and 100 to prevent exceeding the valid range
                $progressBar.Value = [Math]::Min(100, [Math]::Max(0, ($currentItem / $totalItems.Value) * 100))
            }
            
            if ($itemIsFolder) {
                # Log folder
                $logTextBox.AppendText("Found folder: $itemName`r`n")
                
                # Create the folder in the download path if it doesn't exist
                $folderPath = Join-Path -Path $downloadPath -ChildPath $itemName
                if (-not (Test-Path -Path $folderPath)) {
                    New-Item -ItemType Directory -Path $folderPath
                }
                
                # Recursively download files inside this folder
                Download-AllFiles -email $email -accessToken $accessToken -parentId $itemId -downloadPath $folderPath -logTextBox $logTextBox -progressBar $progressBar -totalItems $totalItems
            }
            else {
                # Log file being downloaded
                $logTextBox.AppendText("Downloading file: $itemName`r`n")
                
                # Download the file
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
    
    # Create the zip file
    [System.IO.Compression.ZipFile]::CreateFromDirectory($folderPath, $zipPath)
}

# Create the Windows Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "OneDrive Downloader"
$form.Size = New-Object System.Drawing.Size(600, 550)  # Reduced height
$form.FormBorderStyle = 'FixedDialog'   # Make the window non-resizable
$form.MaximizeBox = $false              # Disable maximize button
$form.MinimizeBox = $false             # Disable minimize button

# User Email Input
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "Enter the user's email address:"
$emailLabel.Location = New-Object System.Drawing.Point(10, 20)
$emailLabel.Size = New-Object System.Drawing.Size(580, 20)  # Adjusted width of label to avoid clipping
$form.Controls.Add($emailLabel)

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Location = New-Object System.Drawing.Point(10, 60)  # Increased padding from top
$emailTextBox.Width = 400
$form.Controls.Add($emailTextBox)

# Download Path Input
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Enter the download path:"
$pathLabel.Location = New-Object System.Drawing.Point(10, 100)  # Increased padding from top
$pathLabel.Size = New-Object System.Drawing.Size(580, 20)  # Adjusted width of label to avoid clipping
$form.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(10, 130)  # Increased padding from top
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
$zipPathLabel.Location = New-Object System.Drawing.Point(10, 170)  # Increased padding from top
$zipPathLabel.Size = New-Object System.Drawing.Size(580, 20)  # Adjusted width of label to avoid clipping
$form.Controls.Add($zipPathLabel)

$zipPathTextBox = New-Object System.Windows.Forms.TextBox
$zipPathTextBox.Location = New-Object System.Drawing.Point(10, 200)  # Increased padding from top
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
$logTextBox.Location = New-Object System.Drawing.Point(10, 250)  # Adjusted Y value for better spacing
$logTextBox.Size = New-Object System.Drawing.Size(560, 150)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = 'Vertical'
$form.Controls.Add($logTextBox)

# Progress Bar to display download progress
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 420)  # Adjusted Y value for better spacing
$progressBar.Size = New-Object System.Drawing.Size(560, 30)
$form.Controls.Add($progressBar)

# Start Download and Zip Button
$downloadAndZipButton = New-Object System.Windows.Forms.Button
$downloadAndZipButton.Text = "Download and Zip"
$downloadAndZipButton.Location = New-Object System.Drawing.Point(10, 460)  # Adjusted Y value for better spacing
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
        
        # Zip the downloaded folder
        $zipPath = Join-Path -Path $zipOutputPath -ChildPath "$($email)_OneDrive_Backup.zip"
        Zip-Folder -folderPath $downloadPath -zipPath $zipPath
        
        # Log success
        $logTextBox.AppendText("Files downloaded and zipped successfully!`r`n")
    } else {
        $logTextBox.AppendText("Please enter all required fields.`r`n")
    }
})

# Run the Form
$form.ShowDialog()
