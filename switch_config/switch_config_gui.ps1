# Script for switch configuration GUI
$pythonpath = "C:\path\to\python.exe"

$scriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition 

$env:Path += ";$Pythonpath"
Add-Type -AssemblyName System.Windows.Forms

# Create form
$formOptions = New-Object System.Windows.Forms.Form
$formOptions.Text = "Switch Configuration"
$formOptions.Width = 600
$formOptions.Height = 450
$formOptions.AutoSize = $true
$formOptions.BackColor = [System.Drawing.Color]::White

$Labelpr = New-Object System.Windows.Forms.Label
$Labelpr.Text = "Select a Switch to Configure:"
$Labelpr.Left = 200
$Labelpr.Top = 100
$Labelpr.Width = 200
$labelpr.Font = New-Object System.Drawing.Font("Arial", 10)
$formOptions.Controls.Add($Labelpr)

$buttonOption1 = New-Object System.Windows.Forms.Button
$buttonOption1.Location = New-Object System.Drawing.Point(200, 150)
$buttonOption1.Size = New-Object System.Drawing.Size(200, 30)
$buttonOption1.Text = "GRS105/106"
$buttonOption1.Add_Click({
    # Code to open Form for Option 1 (FormOption1)
    $formOptions.Hide()  # Hide the current form
    Show-FormOption1    # Show Form for Option 1
    $formOptions.Close() # Close the current form after Option 1 form is closed
})

$buttonOption2 = New-Object System.Windows.Forms.Button
$buttonOption2.Location = New-Object System.Drawing.Point(200, 200)
$buttonOption2.Size = New-Object System.Drawing.Size(200, 30)
$buttonOption2.Text = "GRS1042"
$buttonOption2.Add_Click({
    # Code to open Form for Option 1 (FormOption1)
    $formOptions.Hide()  # Hide the current form
    Show-FormOption2    # Show Form for Option 1
    $formOptions.Close() # Close the current form after Option 1 form is closed
})

$buttonOption3 = New-Object System.Windows.Forms.Button
$buttonOption3.Location = New-Object System.Drawing.Point(200, 250)
$buttonOption3.Size = New-Object System.Drawing.Size(200, 30)
$buttonOption3.Text = "RSPE35"
$buttonOption3.Add_Click({
    # Code to open Form for Option 1 (FormOption1)
    $formOptions.Hide()  # Hide the current form
    Show-FormOption3    # Show Form for Option 1
    $formOptions.Close() # Close the current form after Option 1 form is closed
})

# Add buttons to Option Selection form
$formOptions.Controls.Add($buttonOption1)
$formOptions.Controls.Add($buttonOption2)
$formOptions.Controls.Add($buttonOption3)

##### OPTION 1 ######

function Show-FormOption1{
    $Form1 = New-Object System.Windows.Forms.Form
    $Form1.Text = "GRS105/106 Configuration"
    $Form1.Size = New-Object System.Drawing.Size(600, 450)
    $form1.AutoSize = $true
    $form1.BackColor = [System.Drawing.Color]::White

    # Add Window Title
    $LabelWindow = New-Object System.Windows.Forms.Label
    $LabelWindow.Text = "GRS-105 Configuration"
    $LabelWindow.Left = 10 
    $LabelWindow.Top = 10
    $LabelWindow.Width = 600
    $LabelWindow.Height = 30
    $labelWindow.Font = New-Object System.Drawing.Font("Arial", 18)
    $form1.Controls.Add($LabelWindow)

    # Add project name
    $LabelName = New-Object System.Windows.Forms.Label
    $LabelName.Text = "Project Name:"
    $LabelName.Left = 10 
    $LabelName.Top = 50
    $LabelName.Width = 130
    $form1.Controls.Add($LabelName)
    $txtboxName = New-Object System.Windows.Forms.TextBox
    $txtboxName.Left = 150
    $txtboxName.Width = 250
    $txtboxName.Top = 50
    $form1.Controls.Add($txtboxName)

    # Add customer name
    $LabelCustomer = New-Object System.Windows.Forms.Label
    $LabelCustomer.Text = "Customer Name:"
    $LabelCustomer.Left = 10 
    $LabelCustomer.Top = 80
    $LabelCustomer.Width = 130
    $form1.Controls.Add($LabelCustomer)
    $txtboxCustomer = New-Object System.Windows.Forms.TextBox
    $txtboxCustomer.Left = 150
    $txtboxCustomer.Width = 250
    $txtboxCustomer.Top = 80
    $form1.Controls.Add($txtboxCustomer)


    # Add username
    $LabelUser = New-Object System.Windows.Forms.Label
    $LabelUser.Text = "Admin Username:"
    $LabelUser.Left = 10 
    $LabelUser.Top = 110
    $LabelUser.Width = 130
    $form1.Controls.Add($LabelUser)
    $txtboxUser = New-Object System.Windows.Forms.TextBox
    $txtboxUser.Left = 150
    $txtboxUser.Width = 250
    $txtboxUser.Top = 110
    $form1.Controls.Add($txtboxUser)

    # Add remote password
    $LblRemotePass = New-Object System.Windows.Forms.Label
    $LblRemotePass.Text = "Remote Password:"
    $LblRemotePass.Left = 10
    $LblRemotePass.Top = 140 
    $LblRemotePass.Width = 130
    $form1.Controls.Add($LblRemotePass)
    $txtboxRemotePass = New-Object System.Windows.Forms.TextBox
    $txtboxRemotePass.Left = 150
    $txtboxRemotePass.Width = 250 
    $txtboxRemotePass.Top = 140
    $form1.Controls.Add($txtboxRemotePass)

    # Add Admin Password
    $LblPass = New-Object System.Windows.Forms.Label
    $LblPass.Text = "Admin Password:"
    $LblPass.Left = 10
    $LblPass.Top = 170
    $LblPass.Width = 130
    $form1.Controls.Add($LblPass)
    $txtboxPass = New-Object System.Windows.Forms.TextBox
    $txtboxPass.Left = 150 
    $txtboxPass.Width = 250 
    $txtboxPass.Top = 170
    $form1.Controls.Add($txtboxPass)

    # Add user Password
    $LblUserPass = New-Object System.Windows.Forms.Label
    $LblUserPass.Text = "User Password:"
    $LblUserPass.Left = 10
    $LblUserPass.Top = 200
    $LblUserPass.Width = 130
    $form1.Controls.Add($LblUserPass)
    $txtboxUserPass = New-Object System.Windows.Forms.TextBox
    $txtboxUserPass.Left = 150 
    $txtboxUserPass.Width = 250 
    $txtboxUserPass.Top = 200
    $form1.Controls.Add($txtboxUserPass)

    # add what tab in excel it is
    $LblSheetName = New-Object System.Windows.Forms.Label
    $LblSheetName.Text = "Network Sheet Name:"
    $LblSheetName.Left = 10
    $LblSheetName.Top = 230
    $LblSheetName.Width = 130
    $form1.Controls.Add($LblSheetName)
    $txtboxSheetName = New-Object System.Windows.Forms.TextBox
    $txtboxSheetName.Left = 150 
    $txtboxSheetName.Width = 250 
    $txtboxSheetName.Top = 230
    $form1.Controls.Add($txtboxSheetName)

    # add max temp
    $LblMaxTemp = New-Object System.Windows.Forms.Label
    $LblMaxTemp.Text = "Max Temp:"
    $LblMaxTemp.Left = 10 
    $LblMaxTemp.Top = 290
    $LblMaxTemp.Width = 130
    $form1.Controls.Add($LblMaxTemp)
    $txtboxMaxTemp = New-Object System.Windows.Forms.TextBox
    $txtboxMaxTemp.Left = 150
    $txtboxMaxTemp.Width = 50 
    $txtboxMaxTemp.Top = 290
    $form1.Controls.Add($txtboxMaxTemp)

    #add min temp
    $LblMinTemp = New-Object System.Windows.Forms.Label
    $LblMinTemp.Text = "Min Temp:"
    $LblMinTemp.Left = 10
    $LblMinTemp.Top = 260
    $LblMinTemp.Width = 130
    $form1.Controls.Add($LblMinTemp)
    $txtboxMinTemp = New-Object System.Windows.Forms.TextBox
    $txtboxMinTemp.Left = 150 
    $txtboxMinTemp.Width = 50
    $txtboxMinTemp.Top = 260
    $form1.Controls.Add($txtboxMinTemp)

    # Find excel file
    $LblExcel = New-Object System.Windows.Forms.Label
    $Lblexcel.Text = "Select Excel File:"
    $Lblexcel.Left = 10
    $Lblexcel.Top = 320
    $Lblexcel.Width = 130
    $form1.Controls.Add($LblExcel)
    $textboxFile = New-Object System.Windows.Forms.TextBox
    $textboxFile.Location = New-Object System.Drawing.Point(150, 320)
    $textboxFile.Size = New-Object System.Drawing.Size(250, 20)
    $form1.Controls.Add($textboxFile)
    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Location = New-Object System.Drawing.Point(410, 320)
    $buttonBrowse.Size = New-Object System.Drawing.Size(90, 20)
    $buttonBrowse.Text = "Browse"
    $buttonBrowse.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select an Excel file"  # Dialog title
        $fileDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"  # File filter
        $dialogResult = $fileDialog.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $fileDialog.FileName
            $textboxFile.Text = $selectedFile
        }
    })
    $form1.Controls.Add($buttonBrowse)
    
    $script:arguments = @()

    # Add button
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(200, 360)
    $button.Size = New-Object System.Drawing.Size(200, 25)
    $button.Text = "Generate Switch .config Files"
    $button.Add_Click({
        # Run the Python executable with arguments
        $script:arguments += $txtboxName.Text.Trim()
        $script:arguments += $txtboxCustomer.Text.Trim()
        $script:arguments += $txtboxUser.Text.Trim()
        $script:arguments += $txtboxRemotePass.Text.Trim()
        $script:arguments += $txtboxPass.Text.Trim()
        $script:arguments += $txtboxUserPass.Text.Trim()
        $script:arguments += $txtboxMinTemp.Text.Trim()
        $script:arguments += $txtboxMaxTemp.Text.Trim()
        $script:arguments += $txtboxSheetName.Text.Trim()
        $script:arguments += $textboxFile.Text
        $form1.Dispose()
        
    })
    $form1.Controls.Add($button) 

    # Show the form
    $form1.ShowDialog() | Out-Null

    $grspath = Join-Path -Path $ScriptDir -Childpath "switch_config_GRS105.py"

    & python $grspath $arguments
}
##### OPTION 2 ######

function Show-FormOption2{
    $Form2 = New-Object System.Windows.Forms.Form
    $Form2.Text = "GRS1042 Configuration"
    $Form2.Size = New-Object System.Drawing.Size(600, 450)
    $form2.AutoSize = $true
    $form2.BackColor = [System.Drawing.Color]::White

    # Add Window Title
    $LabelWindow = New-Object System.Windows.Forms.Label
    $LabelWindow.Text = "GRS-1042 Configuration"
    $LabelWindow.Left = 10 
    $LabelWindow.Top = 10
    $LabelWindow.Width = 600
    $LabelWindow.Height = 30
    $labelWindow.Font = New-Object System.Drawing.Font("Arial", 18)
    $form2.Controls.Add($LabelWindow)

    # Add project name
    $LabelName = New-Object System.Windows.Forms.Label
    $LabelName.Text = "Project Name:"
    $LabelName.Left = 10 
    $LabelName.Top = 50
    $LabelName.Width = 130
    $form2.Controls.Add($LabelName)
    $txtboxName = New-Object System.Windows.Forms.TextBox
    $txtboxName.Left = 150
    $txtboxName.Width = 250
    $txtboxName.Top = 50
    $form2.Controls.Add($txtboxName)

    # Add customer name
    $LabelCustomer = New-Object System.Windows.Forms.Label
    $LabelCustomer.Text = "Customer Name:"
    $LabelCustomer.Left = 10 
    $LabelCustomer.Top = 80
    $LabelCustomer.Width = 130
    $form2.Controls.Add($LabelCustomer)
    $txtboxCustomer = New-Object System.Windows.Forms.TextBox
    $txtboxCustomer.Left = 150
    $txtboxCustomer.Width = 250
    $txtboxCustomer.Top = 80
    $form2.Controls.Add($txtboxCustomer)


    # Add username
    $LabelUser = New-Object System.Windows.Forms.Label
    $LabelUser.Text = "Admin Username:"
    $LabelUser.Left = 10 
    $LabelUser.Top = 110
    $LabelUser.Width = 130
    $form2.Controls.Add($LabelUser)
    $txtboxUser = New-Object System.Windows.Forms.TextBox
    $txtboxUser.Left = 150
    $txtboxUser.Width = 250
    $txtboxUser.Top = 110
    $form2.Controls.Add($txtboxUser)

    # Add remote password
    $LblRemotePass = New-Object System.Windows.Forms.Label
    $LblRemotePass.Text = "Remote Password:"
    $LblRemotePass.Left = 10
    $LblRemotePass.Top = 140 
    $LblRemotePass.Width = 130
    $form2.Controls.Add($LblRemotePass)
    $txtboxRemotePass = New-Object System.Windows.Forms.TextBox
    $txtboxRemotePass.Left = 150
    $txtboxRemotePass.Width = 250 
    $txtboxRemotePass.Top = 140
    $form2.Controls.Add($txtboxRemotePass)

    # Add Admin Password
    $LblPass = New-Object System.Windows.Forms.Label
    $LblPass.Text = "Admin Password:"
    $LblPass.Left = 10
    $LblPass.Top = 170
    $LblPass.Width = 130
    $form2.Controls.Add($LblPass)
    $txtboxPass = New-Object System.Windows.Forms.TextBox
    $txtboxPass.Left = 150 
    $txtboxPass.Width = 250 
    $txtboxPass.Top = 170
    $form2.Controls.Add($txtboxPass)

    # Add user Password
    $LblUserPass = New-Object System.Windows.Forms.Label
    $LblUserPass.Text = "User Password:"
    $LblUserPass.Left = 10
    $LblUserPass.Top = 200
    $LblUserPass.Width = 130
    $form2.Controls.Add($LblUserPass)
    $txtboxUserPass = New-Object System.Windows.Forms.TextBox
    $txtboxUserPass.Left = 150 
    $txtboxUserPass.Width = 250 
    $txtboxUserPass.Top = 200
    $form2.Controls.Add($txtboxUserPass)

    # add what tab in excel it is
    $LblSheetName = New-Object System.Windows.Forms.Label
    $LblSheetName.Text = "Network Sheet Name:"
    $LblSheetName.Left = 10
    $LblSheetName.Top = 230
    $LblSheetName.Width = 130
    $form2.Controls.Add($LblSheetName)
    $txtboxSheetName = New-Object System.Windows.Forms.TextBox
    $txtboxSheetName.Left = 150 
    $txtboxSheetName.Width = 250 
    $txtboxSheetName.Top = 230
    $form2.Controls.Add($txtboxSheetName)

    # add max temp
    $LblMaxTemp = New-Object System.Windows.Forms.Label
    $LblMaxTemp.Text = "Max Temp:"
    $LblMaxTemp.Left = 10 
    $LblMaxTemp.Top = 290
    $LblMaxTemp.Width = 130
    $form2.Controls.Add($LblMaxTemp)
    $txtboxMaxTemp = New-Object System.Windows.Forms.TextBox
    $txtboxMaxTemp.Left = 150
    $txtboxMaxTemp.Width = 50 
    $txtboxMaxTemp.Top = 290
    $form2.Controls.Add($txtboxMaxTemp)

    #add min temp
    $LblMinTemp = New-Object System.Windows.Forms.Label
    $LblMinTemp.Text = "Min Temp:"
    $LblMinTemp.Left = 10
    $LblMinTemp.Top = 260
    $LblMinTemp.Width = 130
    $form2.Controls.Add($LblMinTemp)
    $txtboxMinTemp = New-Object System.Windows.Forms.TextBox
    $txtboxMinTemp.Left = 150 
    $txtboxMinTemp.Width = 50
    $txtboxMinTemp.Top = 260
    $form2.Controls.Add($txtboxMinTemp)

    # Find excel file
    $LblExcel = New-Object System.Windows.Forms.Label
    $Lblexcel.Text = "Select Excel File:"
    $Lblexcel.Left = 10
    $Lblexcel.Top = 320
    $Lblexcel.Width = 130
    $form2.Controls.Add($LblExcel)
    $textboxFile = New-Object System.Windows.Forms.TextBox
    $textboxFile.Location = New-Object System.Drawing.Point(150, 320)
    $textboxFile.Size = New-Object System.Drawing.Size(250, 20)
    $form2.Controls.Add($textboxFile)
    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Location = New-Object System.Drawing.Point(410, 320)
    $buttonBrowse.Size = New-Object System.Drawing.Size(90, 20)
    $buttonBrowse.Text = "Browse"
    $buttonBrowse.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select an Excel file"  # Dialog title
        $fileDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"  # File filter
        $dialogResult = $fileDialog.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $fileDialog.FileName
            $textboxFile.Text = $selectedFile
        }
    })
    $form2.Controls.Add($buttonBrowse)

    $script:arguments = @()

    # Add button
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(200, 330)
    $button.Size = New-Object System.Drawing.Size(200, 25)
    $button.Text = "Generate Switch .config Files"
    $button.Add_Click({
        # Run the Python executable with arguments
        $script:arguments += $txtboxName.Text.Trim()
        $script:arguments += $txtboxCustomer.Text.Trim()
        $script:arguments += $txtboxUser.Text.Trim()
        $script:arguments += $txtboxRemotePass.Text.Trim()
        $script:arguments += $txtboxPass.Text.Trim()
        $script:arguments += $txtboxUserPass.Text.Trim()
        $script:arguments += $txtboxMinTemp.Text.Trim()
        $script:arguments += $txtboxMaxTemp.Text.Trim()
        $script:arguments += $txtboxSheetName.Text.Trim()
        $script:arguments += $txtboxFile.Text.Trim()
        $form2.Dispose()
    })
    $form2.Controls.Add($button) 

    # Show the form
    $form2.ShowDialog() | Out-Null
    
    $grspath = Join-Path -Path $ScriptDir -Childpath "switch_config_GRS1042.py"

    & python $grspath $arguments
}

##### OPTION 3 ######
function Show-FormOption3{
    $Form3 = New-Object System.Windows.Forms.Form
    $Form3.Text = "RSPE35 Configuration"
    $Form3.Size = New-Object System.Drawing.Size(600, 450)
    $form3.AutoSize = $true
    $form3.BackColor = [System.Drawing.Color]::White

    # Add Window Title
    $LabelWindow = New-Object System.Windows.Forms.Label
    $LabelWindow.Text = "RSPE35 Configuration"
    $LabelWindow.Left = 10 
    $LabelWindow.Top = 10
    $LabelWindow.Width = 600
    $LabelWindow.Height = 30
    $labelWindow.Font = New-Object System.Drawing.Font("Arial", 18)
    $form3.Controls.Add($LabelWindow)

    # Add project name
    $LabelName = New-Object System.Windows.Forms.Label
    $LabelName.Text = "Project Name:"
    $LabelName.Left = 10 
    $LabelName.Top = 50
    $LabelName.Width = 130
    $form3.Controls.Add($LabelName)
    $txtboxName = New-Object System.Windows.Forms.TextBox
    $txtboxName.Left = 150
    $txtboxName.Width = 250
    $txtboxName.Top = 50
    $form3.Controls.Add($txtboxName)

    # Add customer name
    $LabelCustomer = New-Object System.Windows.Forms.Label
    $LabelCustomer.Text = "Customer Name:"
    $LabelCustomer.Left = 10 
    $LabelCustomer.Top = 80
    $LabelCustomer.Width = 130
    $form3.Controls.Add($LabelCustomer)
    $txtboxCustomer = New-Object System.Windows.Forms.TextBox
    $txtboxCustomer.Left = 150
    $txtboxCustomer.Width = 250
    $txtboxCustomer.Top = 80
    $form3.Controls.Add($txtboxCustomer)


    # Add username
    $LabelUser = New-Object System.Windows.Forms.Label
    $LabelUser.Text = "Admin Username:"
    $LabelUser.Left = 10 
    $LabelUser.Top = 110
    $LabelUser.Width = 130
    $form3.Controls.Add($LabelUser)
    $txtboxUser = New-Object System.Windows.Forms.TextBox
    $txtboxUser.Left = 150
    $txtboxUser.Width = 250
    $txtboxUser.Top = 110
    $form3.Controls.Add($txtboxUser)

    # Add remote password
    $LblRemotePass = New-Object System.Windows.Forms.Label
    $LblRemotePass.Text = "Remote Password:"
    $LblRemotePass.Left = 10
    $LblRemotePass.Top = 140 
    $LblRemotePass.Width = 130
    $form3.Controls.Add($LblRemotePass)
    $txtboxRemotePass = New-Object System.Windows.Forms.TextBox
    $txtboxRemotePass.Left = 150
    $txtboxRemotePass.Width = 250 
    $txtboxRemotePass.Top = 140
    $form3.Controls.Add($txtboxRemotePass)

    # Add Admin Password
    $LblPass = New-Object System.Windows.Forms.Label
    $LblPass.Text = "Admin Password:"
    $LblPass.Left = 10
    $LblPass.Top = 170
    $LblPass.Width = 130
    $form3.Controls.Add($LblPass)
    $txtboxPass = New-Object System.Windows.Forms.TextBox
    $txtboxPass.Left = 150 
    $txtboxPass.Width = 250 
    $txtboxPass.Top = 170
    $form3.Controls.Add($txtboxPass)

    # Add user Password
    $LblUserPass = New-Object System.Windows.Forms.Label
    $LblUserPass.Text = "User Password:"
    $LblUserPass.Left = 10
    $LblUserPass.Top = 200
    $LblUserPass.Width = 130
    $form3.Controls.Add($LblUserPass)
    $txtboxUserPass = New-Object System.Windows.Forms.TextBox
    $txtboxUserPass.Left = 150 
    $txtboxUserPass.Width = 250 
    $txtboxUserPass.Top = 200
    $form3.Controls.Add($txtboxUserPass)

    # add what tab in excel it is
    $LblSheetName = New-Object System.Windows.Forms.Label
    $LblSheetName.Text = "Network Sheet Name:"
    $LblSheetName.Left = 10
    $LblSheetName.Top = 230
    $LblSheetName.Width = 130
    $form3.Controls.Add($LblSheetName)
    $txtboxSheetName = New-Object System.Windows.Forms.TextBox
    $txtboxSheetName.Left = 150 
    $txtboxSheetName.Width = 250 
    $txtboxSheetName.Top = 230
    $form3.Controls.Add($txtboxSheetName)

    # add max temp
    $LblMaxTemp = New-Object System.Windows.Forms.Label
    $LblMaxTemp.Text = "Max Temp:"
    $LblMaxTemp.Left = 10 
    $LblMaxTemp.Top = 290
    $LblMaxTemp.Width = 130
    $form3.Controls.Add($LblMaxTemp)
    $txtboxMaxTemp = New-Object System.Windows.Forms.TextBox
    $txtboxMaxTemp.Left = 150
    $txtboxMaxTemp.Width = 50 
    $txtboxMaxTemp.Top = 290
    $form3.Controls.Add($txtboxMaxTemp)

    #add min temp
    $LblMinTemp = New-Object System.Windows.Forms.Label
    $LblMinTemp.Text = "Min Temp:"
    $LblMinTemp.Left = 10
    $LblMinTemp.Top = 260
    $LblMinTemp.Width = 130
    $form3.Controls.Add($LblMinTemp)
    $txtboxMinTemp = New-Object System.Windows.Forms.TextBox
    $txtboxMinTemp.Left = 150 
    $txtboxMinTemp.Width = 50
    $txtboxMinTemp.Top = 260
    $form3.Controls.Add($txtboxMinTemp)

    # Find excel file
    $LblExcel = New-Object System.Windows.Forms.Label
    $Lblexcel.Text = "Select Excel File:"
    $Lblexcel.Left = 10
    $Lblexcel.Top = 320
    $Lblexcel.Width = 130
    $form2.Controls.Add($LblExcel)
    $textboxFile = New-Object System.Windows.Forms.TextBox
    $textboxFile.Location = New-Object System.Drawing.Point(150, 320)
    $textboxFile.Size = New-Object System.Drawing.Size(250, 20)
    $form2.Controls.Add($textboxFile)
    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Location = New-Object System.Drawing.Point(410, 320)
    $buttonBrowse.Size = New-Object System.Drawing.Size(90, 20)
    $buttonBrowse.Text = "Browse"
    $buttonBrowse.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select an Excel file"  # Dialog title
        $fileDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"  # File filter
        $dialogResult = $fileDialog.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $fileDialog.FileName
            $textboxFile.Text = $selectedFile
        }
    })
    $form2.Controls.Add($buttonBrowse)

    $script:argumentss = @()

    # Add button
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(200, 330)
    $button.Size = New-Object System.Drawing.Size(200, 25)
    $button.Text = "Generate Switch .config Files"
    $button.Add_Click({
        # Run the Python executable with arguments
        $script:arguments += $txtboxName.Text.Trim()
        $script:arguments += $txtboxCustomer.Text.Trim()
        $script:arguments += $txtboxUser.Text.Trim()
        $script:arguments += $txtboxRemotePass.Text.Trim()
        $script:arguments += $txtboxPass.Text.Trim()
        $script:arguments += $txtboxUserPass.Text.Trim()
        $script:arguments += $txtboxMinTemp.Text.Trim()
        $script:arguments += $txtboxMaxTemp.Text.Trim()
        $script:arguments += $txtboxSheetName.Text.Trim()
        $script:arguments += $txtboxFile.Text.Trim()
        $form3.Dispose()
    })
    $form3.Controls.Add($button) 

    # Show the form
    $form3.ShowDialog() | Out-Null
    
    $grspath = Join-Path -Path $ScriptDir -Childpath "switch_config_RSPE35.py"

    & python $grspath $arguments
}

$formOptions.ShowDialog() | Out-Null





