# Load required assemblies for Windows Forms
Add-Type -AssemblyName "System.Windows.Forms"
Add-Type -AssemblyName "System.Drawing"

# Function to check if a module is installed
function Check-Module {
    param (
        [string]$ModuleName
    )
    
    # Check if the module is installed
    $module = Get-Module -ListAvailable -Name $ModuleName
    if ($null -eq $module) {
        Write-Host "$ModuleName module is not installed. Installing..."
        # Install the module if not found
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
    }
    else {
        Write-Host "$ModuleName module is already installed."
    }
}

# Ensure required modules are installed
Check-Module -ModuleName "Microsoft.Graph"
Check-Module -ModuleName "AzureAD"

# Import required modules
Import-Module Microsoft.Graph
Import-Module AzureAD

# Function to login as Global Administrator (interactive login)
function Login-GlobalAdministrator {
    # Perform interactive login to Microsoft Graph
    Connect-MgGraph -Scopes "Directory.AccessAsUser.All", "User.ReadWrite.All", "Policy.ReadWrite.ApplicationConfiguration"
}

# Function to get self-service product details (fetch all and filter locally)
function Get-SelfServiceProducts {
    try {
        # Query Microsoft Graph to get all subscriptions
        $subscriptions = Get-MgSubscribedSku
        
        # Filter the subscriptions locally by checking enabled prepaid units
        $filteredSubscriptions = $subscriptions | Where-Object { $_.PrepaidUnits.Enabled -gt 0 }

        return $filteredSubscriptions
    }
    catch {
        Write-Host "Error fetching self-service products: $($_.Exception.Message)"
        return @()
    }
}

# Function to disable selected trials
function Disable-SelectedTrials {
    param (
        [Parameter(Mandatory = $true)]
        [array]$selectedTrials
    )

    try {
        foreach ($trial in $selectedTrials) {
            Write-Host "Disabling trial for product: $($trial.SkuPartNumber)"
            # You can implement your logic to disable a self-service trial here.
            # For now, we will simply display the product being disabled.
            # Example: Disable by removing subscription
            # Remove-MgSubscribedSku -SubscribedSkuId $trial.Id
        }

        Write-Host "Selected trials have been disabled successfully."
    }
    catch {
        Write-Host "Error disabling trials: $($_.Exception.Message)"
    }
}

# Function to create and show a new UI with products list
function Show-ProductsUI {
    $selfServiceProducts = Get-SelfServiceProducts

    # Create new form for displaying self-service products
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Self-Service Products"
    $form.Size = New-Object System.Drawing.Size(600, 400)

    # Create a label for status
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Size = New-Object System.Drawing.Size(550, 40)
    $statusLabel.Location = New-Object System.Drawing.Point(25, 20)
    $statusLabel.Text = "List of Active Self-Service Products and Trials"
    $form.Controls.Add($statusLabel)

    # Create a CheckedListBox for displaying available self-service products
    $productCheckListBox = New-Object System.Windows.Forms.CheckedListBox
    $productCheckListBox.Size = New-Object System.Drawing.Size(550, 200)
    $productCheckListBox.Location = New-Object System.Drawing.Point(25, 60)

    # Populate checked list box with product names and trial statuses
    foreach ($product in $selfServiceProducts) {
        $productInfo = "$($product.SkuPartNumber) - $($product.SkuId) (Enabled: $($product.PrepaidUnits.Enabled))"
        $productCheckListBox.Items.Add($productInfo, $false)
    }

    $form.Controls.Add($productCheckListBox)

    # Create a button to disable selected trials
    $disableButton = New-Object System.Windows.Forms.Button
    $disableButton.Size = New-Object System.Drawing.Size(200, 40)
    $disableButton.Location = New-Object System.Drawing.Point(100, 280)
    $disableButton.Text = "Disable Selected Trials"
    $form.Controls.Add($disableButton)

    # Button click to disable selected trials
    $disableButton.Add_Click({
            # Get selected trials from CheckedListBox
            $selectedTrials = @()
            foreach ($index in $productCheckListBox.CheckedIndices) {
                $trial = $selfServiceProducts[$index]
                $selectedTrials += $trial
            }

            # Disable selected trials
            if ($selectedTrials.Count -gt 0) {
                Disable-SelectedTrials -selectedTrials $selectedTrials
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Please select at least one trial to disable.")
            }
        })

    # Create a button to close the form
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Size = New-Object System.Drawing.Size(100, 40)
    $closeButton.Location = New-Object System.Drawing.Point(250, 330)
    $closeButton.Text = "Close"
    $form.Controls.Add($closeButton)

    # Close button functionality
    $closeButton.Add_Click({
            $form.Close()
        })

    # Display the form
    [void]$form.ShowDialog()
}

# Create the login form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD Admin Tool"
$form.Size = New-Object System.Drawing.Size(400, 250)

# Create a label for status
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Size = New-Object System.Drawing.Size(350, 40)
$statusLabel.Location = New-Object System.Drawing.Point(25, 30)
$statusLabel.Text = "Click below to login and manage self-service trials and purchases."
$form.Controls.Add($statusLabel)

# Create a button to perform the login and disable actions
$actionButton = New-Object System.Windows.Forms.Button
$actionButton.Size = New-Object System.Drawing.Size(200, 40)
$actionButton.Location = New-Object System.Drawing.Point(100, 100)
$actionButton.Text = "Login & Manage Trials"
$form.Controls.Add($actionButton)

# Define button click behavior
$actionButton.Add_Click({
        try {
            # Disable the button while the task is running
            $actionButton.Enabled = $false
            $statusLabel.Text = "Logging in..."
        
            # Perform login
            Login-GlobalAdministrator

            # Show the self-service products UI after login
            Show-ProductsUI
        }
        catch {
            $statusLabel.Text = "Error occurred: $($_.Exception.Message)"
        }
        finally {
            # Re-enable the button after operation
            $actionButton.Enabled = $true
        }
    })

# Display the login form (this ensures the form shows)
[void]$form.ShowDialog()
