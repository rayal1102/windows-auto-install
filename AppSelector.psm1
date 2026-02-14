# ===================================================================
# GUI APP SELECTOR - SU DUNG WINDOWS FORMS
# ===================================================================

function Show-AppSelector {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Apps,
        [Parameter(Mandatory=$true)]
        [string]$Title,
        [string]$Description = ""
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Tao form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(600, 700)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(10, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(560, 30)
    $titleLabel.Text = $Title
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)
    
    # Description label
    if ($Description) {
        $descLabel = New-Object System.Windows.Forms.Label
        $descLabel.Location = New-Object System.Drawing.Point(10, 45)
        $descLabel.Size = New-Object System.Drawing.Size(560, 20)
        $descLabel.Text = $Description
        $descLabel.ForeColor = [System.Drawing.Color]::Gray
        $form.Controls.Add($descLabel)
    }
    
    # Search box
    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Location = New-Object System.Drawing.Point(10, 75)
    $searchLabel.Size = New-Object System.Drawing.Size(100, 20)
    $searchLabel.Text = "Tim kiem:"
    $form.Controls.Add($searchLabel)
    
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(90, 73)
    $searchBox.Size = New-Object System.Drawing.Size(380, 25)
    $form.Controls.Add($searchBox)
    
    # Buttons: Select All, Deselect All, Clear Search
    $selectAllBtn = New-Object System.Windows.Forms.Button
    $selectAllBtn.Location = New-Object System.Drawing.Point(475, 72)
    $selectAllBtn.Size = New-Object System.Drawing.Size(100, 25)
    $selectAllBtn.Text = "Chon tat ca"
    $form.Controls.Add($selectAllBtn)
    
    # CheckedListBox for apps
    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(10, 110)
    $checkedListBox.Size = New-Object System.Drawing.Size(565, 480)
    $checkedListBox.CheckOnClick = $true
    $checkedListBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($checkedListBox)
    
    # Populate apps
    $script:allApps = @()
    foreach ($app in $Apps) {
        $displayText = $app.name
        if ($app.description) {
            $displayText += " - $($app.description)"
        }
        $index = $checkedListBox.Items.Add($displayText)
        $script:allApps += @{
            Index = $index
            Name = $app.name
            Id = $app.id
            Description = $app.description
            DisplayText = $displayText
        }
    }
    
    # Selected count label
    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Location = New-Object System.Drawing.Point(10, 600)
    $countLabel.Size = New-Object System.Drawing.Size(300, 20)
    $countLabel.Text = "Da chon: 0/$($Apps.Count)"
    $countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($countLabel)
    
    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(380, 595)
    $okButton.Size = New-Object System.Drawing.Size(90, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton
    
    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(480, 595)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.Text = "Huy"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton
    
    # Update count function
    $updateCount = {
        $selectedCount = $checkedListBox.CheckedItems.Count
        $countLabel.Text = "Da chon: $selectedCount/$($Apps.Count)"
    }
    
    # Select All button click
    $selectAllBtn.Add_Click({
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            $checkedListBox.SetItemChecked($i, $true)
        }
        & $updateCount
    })
    
    # Search box text changed
    $searchBox.Add_TextChanged({
        $searchText = $searchBox.Text.ToLower()
        $checkedListBox.Items.Clear()
        
        if ([string]::IsNullOrWhiteSpace($searchText)) {
            # Show all
            foreach ($app in $script:allApps) {
                $checkedListBox.Items.Add($app.DisplayText) | Out-Null
            }
        } else {
            # Filter
            foreach ($app in $script:allApps) {
                if ($app.DisplayText.ToLower().Contains($searchText)) {
                    $checkedListBox.Items.Add($app.DisplayText) | Out-Null
                }
            }
        }
    })
    
    # Item check changed
    $checkedListBox.Add_ItemCheck({
        # Delay update to after check state changes
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 10
        $timer.Add_Tick({
            & $updateCount
            $timer.Stop()
            $timer.Dispose()
        })
        $timer.Start()
    })
    
    # Show form
    $result = $form.ShowDialog()
    
    # Get selected apps
    $selectedApps = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($checkedItem in $checkedListBox.CheckedItems) {
            $itemText = $checkedItem.ToString()
            $app = $script:allApps | Where-Object { $_.DisplayText -eq $itemText }
            if ($app) {
                $selectedApps += @{
                    name = $app.Name
                    id = $app.Id
                }
            }
        }
    }
    
    $form.Dispose()
    return $selectedApps
}

# Export function
Export-ModuleMember -Function Show-AppSelector
