# PowerShell Script with GUI to Display, Add, Remove, and Save Users and Departments to CSV File

# Load Windows Forms and Drawing assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Unassigned User List Editor'
$form.Size = New-Object System.Drawing.Size(500,700)
$form.StartPosition = 'CenterScreen'

# Create a list box to display users and departments
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,10)
$listBox.Size = New-Object System.Drawing.Size(460,500)
$listBox.SelectionMode = 'MultiExtended'

# Load users from CSV file
$csvFilePath = "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Ref\users.csv"
if(Test-Path $csvFilePath) {
    $userData = Import-Csv $csvFilePath
    foreach($user in $userData) {
        $listBox.Items.Add("$($user.Email) - $($user.Department)")
    }
}

# Create input fields for new user
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Location = New-Object System.Drawing.Point(10,520)
$emailLabel.Size = New-Object System.Drawing.Size(120,20)
$emailLabel.Text = 'New User Email:'

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Location = New-Object System.Drawing.Point(130,520)
$emailTextBox.Size = New-Object System.Drawing.Size(240,20)

# Create input fields for department
$departmentLabel = New-Object System.Windows.Forms.Label
$departmentLabel.Location = New-Object System.Drawing.Point(10,550)
$departmentLabel.Size = New-Object System.Drawing.Size(120,20)
$departmentLabel.Text = 'Department:'

$departmentTextBox = New-Object System.Windows.Forms.TextBox
$departmentTextBox.Location = New-Object System.Drawing.Point(130,550)
$departmentTextBox.Size = New-Object System.Drawing.Size(240,20)

# Create a button to add new user
$addButton = New-Object System.Windows.Forms.Button
$addButton.Location = New-Object System.Drawing.Point(90,580)
$addButton.Size = New-Object System.Drawing.Size(120,30)
$addButton.Text = 'Add User'
$addButton.Add_Click({
    if($emailTextBox.Text -ne '' -and $departmentTextBox.Text -ne '') {
        $department = $departmentTextBox.Text.ToUpper() 
        $listBox.Items.Add("$($emailTextBox.Text) - $department")
        $emailTextBox.Text = ''
        $departmentTextBox.Text = ''
    }
})
# Create a button to remove selected user
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Location = New-Object System.Drawing.Point(290,580)
$removeButton.Size = New-Object System.Drawing.Size(120,30)
$removeButton.Text = 'Remove Selected'
$removeButton.Add_Click({
    $selectedItems = $listBox.SelectedItems
    if($selectedItems.Count -gt 0) {
        $selectedItems | ForEach-Object { $listBox.Items.Remove($_) }
    }
})

# Create a button to save the current list to CSV
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Location = New-Object System.Drawing.Point(150,620)
$saveButton.Size = New-Object System.Drawing.Size(200,30)
$saveButton.Text = 'Save Changes'
$saveButton.Add_Click({
    $updatedUserData = $listBox.Items | ForEach-Object { 
        $parts = $_ -split ' - '
        [PSCustomObject]@{Email = $parts[0]; Department = $parts[1].ToUpper()} # Ensure department is in upper case
    }
    $updatedUserData | Export-Csv $csvFilePath -NoTypeInformation -Force
    [System.Windows.Forms.MessageBox]::Show("Changes saved to CSV file.")
    $form.Close() # This line closes the form
})

# Add controls to the form
$form.Controls.Add($listBox)
$form.Controls.Add($emailLabel)
$form.Controls.Add($emailTextBox)
$form.Controls.Add($departmentLabel)
$form.Controls.Add($departmentTextBox)
$form.Controls.Add($addButton)
$form.Controls.Add($removeButton)
$form.Controls.Add($saveButton)

# Show the form
$form.ShowDialog()
