<#
.SYNOPSIS
This script generates a password using predefined words and displays it in a Windows Forms GUI. It also provides options to copy the password to the clipboard or generate a new password.

.DESCRIPTION
The script creates a GUI that shows a generated password, allows the user to copy it to the clipboard, and generate new passwords on-demand. Optionally, it can also compile the script into an EXE using the PS2EXE module.

.PARAMETER generateEXE
If provided as "true", the script will generate an executable (EXE) file from the PowerShell script using PS2EXE.

.EXAMPLE
.\GeneratePasswordGUI.ps1 -generateEXE "true"
This will generate an EXE file without showing the GUI.

.EXAMPLE
.\GeneratePasswordGUI.ps1
This will run the script normally, displaying the password generator GUI.

.NOTES
File version: 1.0
Author: [Your Name]
Date: 2025-03-17
#>

param (
    [string]$generateEXE = "false"  # Flag to generate EXE file, default is false
)

# Check if PS2EXE module is installed and install if missing
if (-not (Get-Module -ListAvailable -Name PS2EXE)) {
    Write-Host "PS2EXE module not found. Installing..."
    Install-Module -Name PS2EXE -Force -Scope CurrentUser
}

# Import PS2EXE module
Import-Module PS2EXE

# If EXE generation is requested, only generate the EXE and exit
if ($generateEXE -eq "true") {
    Write-Host "Generating EXE file..."
    # Compile the script into an EXE
    Invoke-PS2EXE -inputFile $MyInvocation.MyCommand.Path -outputFile "GeneratePasswordGUI.exe"
    Write-Host "EXE file generated successfully."
    return  # Exit the script to prevent GUI from being displayed
}

# Define a function to generate a random number between 10 and 100
function GenerateRandomNumber {
    return Get-Random -Minimum 10 -Maximum 100
}

# Define a list of words for password generation
$wordList = @(
    "computer", "school", "teacher", "student", "pen", "pencil", "desk", "chair", "paper", "eraser",
    "ruler", "math", "science", "art", "music", "play", "friend", "happy", "sad", "fun",
    "game", "park", "color", "red", "blue", "green", "yellow", "purple", "orange", "pink",
    "black", "white", "brown", "gray", "shoes", "socks", "shirt", "pants", "hat", "jacket",
    "sweater", "dress", "shorts", "skirt", "glasses", "hat", "gloves", "scarf", "boots", "backpack",
    "lunchbox", "bedroom", "kitchen", "bathroom", "livingroom", "bed", "table", "chair", "sofa", "TV",
    "computer", "phone", "door", "window", "floor", "fruit", "vegetable", "pizza", "cake", "ice cream",
    "candy", "cookie", "sandwich", "juice", "milk", "water", "bread", "cheese", "chicken", "pasta",
    "rice", "soup", "salad", "burger", "fries", "pizza", "spaghetti", "pancake", "waffle", "grapes",
    "melon", "strawberry", "carrot", "broccoli", "potato", "tomato", "onion", "lettuce", "banana", "apple",
    "orange", "pear", "peach", "grapefruit", "lemon", "watermelon", "pineapple", "cherry", "blueberry", "raspberry",
    "peas", "corn", "beans", "pumpkin", "cucumber"
)

# Function to generate a new password
function GenerateNewPassword {
    $word1 = $wordList | Get-Random
    $word2 = $wordList | Get-Random

    # Capitalize the first word
    $word1 = $word1.Substring(0, 1).ToUpper() + $word1.Substring(1)

    $number = GenerateRandomNumber
    $symbols = "!", "?", "*"
    $symbol = $symbols | Get-Random

    return "$word1$number$symbol$word2"
}

# Define a global variable to store the password
$global:password = GenerateNewPassword

# Create Form (only if not generating EXE)
$form = New-Object System.Windows.Forms.Form
$form.Text = "Password Generator"
$form.Size = New-Object System.Drawing.Size(300, 200) # Adjusted height
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog" # Set form to not resizable

# Create TextBox to display password
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Location = New-Object System.Drawing.Point(10, 20)
$textbox.Width = $form.ClientSize.Width - 20 # Adjust width of textbox to fit the form
$textbox.Text = $global:password
$textbox.ReadOnly = $true

# Create Button to copy password to clipboard
$buttonCopy = New-Object System.Windows.Forms.Button
$buttonCopy.Location = New-Object System.Drawing.Point(20, 60) # Adjusted location
$buttonCopy.Size = New-Object System.Drawing.Size(120, 40)
$buttonCopy.Text = "Copy to Clipboard"

$buttonCopy.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($global:password)
        [System.Windows.Forms.MessageBox]::Show("Password copied to clipboard.", "Copy Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

# Create Button to generate a new password
$buttonNew = New-Object System.Windows.Forms.Button
$buttonNew.Location = New-Object System.Drawing.Point(160, 60) # Adjusted location
$buttonNew.Size = New-Object System.Drawing.Size(120, 40)
$buttonNew.Text = "Generate New Password"

$buttonNew.Add_Click({
        $global:password = GenerateNewPassword
        $textbox.Text = $global:password
    })

# Add controls to form
$form.Controls.Add($textbox)
$form.Controls.Add($buttonCopy)
$form.Controls.Add($buttonNew)

# Define the form close event to kill the application
$form.Add_Closing({
        [System.Windows.Forms.Application]::Exit()
    })

# Display the form
$form.ShowDialog()
