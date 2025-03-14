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
    try {
        # Disconnect any existing sessions first (force re-login)
        Disconnect-MgGraph
        
        # Perform interactive login to Microsoft Graph with specific permissions requested
        Connect-MgGraph -Scopes "Directory.Read.All", "User.Read.All", "Organization.Read.All"
        
        Write-Host "Successfully logged in as Global Administrator."
    }
    catch {
        Write-Host "Login failed: $($_.Exception.Message)"
        return $false
    }
    return $true
}

# Function to get self-service product details (fetch all and filter locally)
function Get-SelfServiceProducts {
    try {
        # Query Microsoft Graph to get all subscriptions
        $subscriptions = Get-MgSubscribedSku
        
        # Filter for self-service products based on ServiceName or SKU Name
        $selfServiceSubscriptions = $subscriptions | Where-Object { 
            $_.ServicePlans -match "Self-service" -or $_.SkuPartNumber -like "*Trial*" 
        }

        return $selfServiceSubscriptions
    }
    catch {
        Write-Host "Error fetching self-service products: $($_.Exception.Message)"
        return @()
    }
}

# Function to get users assigned to a specific SKU
function Get-UsersAssignedToSelfServiceLicense {
    param (
        [Parameter(Mandatory = $true)]
        [string]$skuId
    )
    
    try {
        $users = Get-MgUser -Select "Id,DisplayName,AssignedLicenses"
        $assignedUsers = @()
        
        foreach ($user in $users) {
            foreach ($license in $user.AssignedLicenses) {
                if ($license.SkuId -eq $skuId) {
                    # Add the user and license assignment to the array
                    $assignedUsers += [PSCustomObject]@{
                        UserName = $user.DisplayName
                        UserId   = $user.Id
                    }
                }
            }
        }
        
        return $assignedUsers
    }
    catch {
        Write-Host "Error fetching users assigned to self-service license: $($_.Exception.Message)"
        return @()
    }
}

# Function to create and show a new UI with products list and users in a table
function Show-ProductsWithUsersUI {
    $selfServiceProducts = Get-SelfServiceProducts

    # Retrieve the tenant's name
    $tenantInfo = Get-MgOrganization
    $tenantName = $tenantInfo.DisplayName

    # Create new form for displaying self-service products
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Self-Service Products and Assigned Users"
    $form.Size = New-Object System.Drawing.Size(800, 600)

    # Create a label for tenancy name
    $tenantLabel = New-Object System.Windows.Forms.Label
    $tenantLabel.Size = New-Object System.Drawing.Size(750, 30)
    $tenantLabel.Location = New-Object System.Drawing.Point(25, 25)
    $tenantLabel.Text = "Tenant: $tenantName"
    $form.Controls.Add($tenantLabel)

    # Create a DataGridView to display the table
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Size = New-Object System.Drawing.Size(750, 400)
    $dataGridView.Location = New-Object System.Drawing.Point(25, 60)
    
    # Set DataGridView properties
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.AllowUserToDeleteRows = $false
    $dataGridView.ReadOnly = $true
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    
    # Define columns for the DataGridView (removed Date Purchased column)
    $dataGridView.Columns.Add("LicenseName", "Self-Service License Display Name")
    $dataGridView.Columns.Add("User", "User")

    # Populate the DataGridView with self-service products and their assigned users
    foreach ($product in $selfServiceProducts) {
        $assignedUsers = Get-UsersAssignedToSelfServiceLicense -skuId $product.SkuId
        
        foreach ($user in $assignedUsers) {
            # Add row to DataGridView (no SKU column or Date Purchased column)
            $dataGridView.Rows.Add($product.SkuPartNumber, $user.UserName)
        }
    }

    # Add DataGridView to form
    $form.Controls.Add($dataGridView)

    # Create a button to close the form
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Size = New-Object System.Drawing.Size(100, 40)
    $closeButton.Location = New-Object System.Drawing.Point(350, 500)
    $closeButton.Text = "Close"
    $form.Controls.Add($closeButton)

    # Close button functionality
    $closeButton.Add_Click({
            $form.Close()
        })

    # Display the form
    [void]$form.ShowDialog()
}

# Create the main form with login button
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD Admin Tool"
$form.Size = New-Object System.Drawing.Size(400, 250)

# Create a label for status
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Size = New-Object System.Drawing.Size(350, 40)
$statusLabel.Location = New-Object System.Drawing.Point(25, 30)
$statusLabel.Text = "Click below to login and manage self-service trials and purchases."
$form.Controls.Add($statusLabel)

# Create a button to perform the login and display table
$loginButton = New-Object System.Windows.Forms.Button
$loginButton.Size = New-Object System.Drawing.Size(200, 40)
$loginButton.Location = New-Object System.Drawing.Point(100, 100)
$loginButton.Text = "Login & Show Trials"
$form.Controls.Add($loginButton)

# Define login button click behavior
$loginButton.Add_Click({
        try {
            # Disable the button while the task is running
            $loginButton.Enabled = $false
            $statusLabel.Text = "Logging in..."
        
            # Perform login
            $loginSuccessful = Login-GlobalAdministrator

            if ($loginSuccessful) {
                # Hide the main form (Azure AD Admin Tool)
                $form.Close()

                # Show the self-service products UI after login
                Show-ProductsWithUsersUI
            }
            else {
                $statusLabel.Text = "Login failed. Please try again."
            }
        }
        catch {
            $statusLabel.Text = "Error occurred: $($_.Exception.Message)"
        }
        finally {
            # Re-enable the button after operation
            $loginButton.Enabled = $true
        }
    })

# Display the main form (this ensures the form shows)
[void]$form.ShowDialog()
