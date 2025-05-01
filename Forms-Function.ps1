function Show-UserSelectionForm {
    param (
        [Parameter(Mandatory=$true)]
        [array]$UserList
    )

    # Fix 1: Add error handling for icon loading
    $Micon = $null
    try {
        $Micon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Microsoft_logo.ico")
    } catch {
        Write-Verbose "Unable to load icon: $_"
    }
    
    # Add isExternal property to users
    foreach ($user in $UserList) {
        # Check if the user is external by looking for #EXT# in userPrincipalName
        $user | Add-Member -NotePropertyName 'isExternal' -NotePropertyValue ($user.userPrincipalName -match '#EXT#')
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Users for Shortcut Creation"
    $form.Size = New-Object System.Drawing.Size(700, 520)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    if ($Micon) { $form.Icon = $Micon }
    $form.Font = New-Object System.Drawing.Font("Arial", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

    # Create a consistent style for buttons
    function Add-StyledButton {
        param (
            [string]$Text, 
            [int]$X, 
            [int]$Y, 
            [int]$Width = 100, 
            [int]$Height = 30,
            [bool]$Primary = $false
        )
        
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Size = New-Object System.Drawing.Size($Width, $Height)
        $button.Location = New-Object System.Drawing.Point($X, $Y)
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.Cursor = [System.Windows.Forms.Cursors]::Hand
        
        if ($Primary) {
            $button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            $button.ForeColor = [System.Drawing.Color]::White
            $button.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $button.FlatAppearance.BorderSize = 0
            $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
            $button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 90, 180)
        } else {
            $button.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
            $button.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
            $button.Font = New-Object System.Drawing.Font("Arial", 9)
            $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
            $button.FlatAppearance.BorderSize = 1
            $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
            $button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
        }
        
        return $button
    }

    # Create panel to hold the search elements (for better visual grouping)
    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Size = New-Object System.Drawing.Size(660, 40)
    $searchPanel.Location = New-Object System.Drawing.Point(10, 10)
    $searchPanel.BackColor = [System.Drawing.Color]::White
    $searchPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Create search icon using PictureBox
    $searchIconBox = New-Object System.Windows.Forms.PictureBox
    $searchIconBox.Size = New-Object System.Drawing.Size(20, 20)
    $searchIconBox.Location = New-Object System.Drawing.Point(15, 10)
    $searchIconBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

    # Load the search icon with error handling
    try {
    $searchIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Search_icon.ico")
    $searchIconBox.Image = $searchIcon.ToBitmap()
    $searchPanel.Controls.Add($searchIconBox)
    } catch {
    Write-Verbose "Unable to load search icon: $_"
    # If icon fails to load, don't add anything to maintain clean layout
    }

    # Enhanced search box with adjusted position
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Size = New-Object System.Drawing.Size(600, 25)
    $searchBox.Location = New-Object System.Drawing.Point(45, 8)  # Adjusted position for icon
    $searchBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $searchBox.BackColor = [System.Drawing.Color]::White
    $searchBox.ForeColor = [System.Drawing.Color]::Gray
    $searchBox.Text = "Search by name or email..."
    $searchPanel.Controls.Add($searchBox)

    # Add event handlers for placeholder behavior
    $searchBox.Add_GotFocus({
        if ($this.Text -eq "Search by name or email...") {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
        }
    })

    $searchBox.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = "Search by name or email..."
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $form.Controls.Add($searchPanel)

    # Improved DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Size = New-Object System.Drawing.Size(660, 370)
    $dataGridView.Location = New-Object System.Drawing.Point(10, 60)
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.AllowUserToDeleteRows = $false
    $dataGridView.SelectionMode = 'FullRowSelect'
    $dataGridView.MultiSelect = $false
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.AutoSizeColumnsMode = 'Fill'
    $dataGridView.ScrollBars = 'Vertical'
    $dataGridView.BackgroundColor = [System.Drawing.Color]::White
    $dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $dataGridView.Font = New-Object System.Drawing.Font("Arial", 9)
    $dataGridView.GridColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $dataGridView.RowTemplate.Height = 30
    $dataGridView.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $dataGridView.RowsDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 249, 255)
    $dataGridView.RowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    $dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)

    # Create and configure columns
    $checkColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $checkColumn.HeaderText = "Select"
    $checkColumn.Width = 60
    $checkColumn.Name = "Include"
    $checkColumn.ReadOnly = $false  # Only this column is editable
    $checkColumn.FillWeight = 15
    $dataGridView.Columns.Add($checkColumn) | Out-Null

    # Helper function to create read-only columns
    function Add-ReadOnlyColumn {
        param (
            [string]$Name,
            [string]$HeaderText,
            [int]$FillWeight = 30
        )
        
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $Name
        $column.HeaderText = $HeaderText
        $column.ReadOnly = $true
        $column.FillWeight = $FillWeight
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $column.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
        return $column
    }

    $nameColumn = Add-ReadOnlyColumn -Name "DisplayName" -HeaderText "Display Name" -FillWeight 40
    $mailColumn = Add-ReadOnlyColumn -Name "Mail" -HeaderText "Email" -FillWeight 45
    # Fix 3: Change FillWeight from 0 to a small value (1)
    $idColumn = Add-ReadOnlyColumn -Name "Id" -HeaderText "ID" -FillWeight 1
    
    # Add external indicator column
    $externalColumn = Add-ReadOnlyColumn -Name "External" -HeaderText "External" -FillWeight 15
    
    $dataGridView.Columns.Add($nameColumn) | Out-Null
    $dataGridView.Columns.Add($mailColumn) | Out-Null
    $dataGridView.Columns.Add($externalColumn) | Out-Null
    $dataGridView.Columns.Add($idColumn) | Out-Null
    
    # Hide the ID column as it's not typically needed for user display
    $dataGridView.Columns[4].Visible = $false

    # Style the header
    $dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $dataGridView.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dataGridView.ColumnHeadersHeight = 35
    $dataGridView.EnableHeadersVisualStyles = $false

    $form.Controls.Add($dataGridView)

    # Create styled buttons
    $selectAllButton = Add-StyledButton -Text "Select All" -X 10 -Y 440
    $form.Controls.Add($selectAllButton)

    $deselectAllButton = Add-StyledButton -Text "Deselect All" -X 120 -Y 440
    $form.Controls.Add($deselectAllButton)

    # Add a count indicator label
    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Size = New-Object System.Drawing.Size(130, 23)
    $countLabel.Location = New-Object System.Drawing.Point(230, 444)
    $countLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($countLabel)

    # Add toggle external users button
    $toggleExternalButton = Add-StyledButton -Text "Hide External Users" -X 400 -Y 440 -Width 150
    $form.Controls.Add($toggleExternalButton)
    # Variable to track if external users are visible
    $script:showExternalUsers = $true

    $okButton = Add-StyledButton -Text "OK" -X 580 -Y 440 -Width 100 -Height 35 -Primary $true
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $checkedUsers = @{}

    # Function to update selection count
    function Update-SelectionCount {
        $selectedCount = ($checkedUsers.Values | Where-Object { $_ -eq $true }).Count
        $totalCount = $dataGridView.Rows.Count
        $countLabel.Text = "$selectedCount of $totalCount users shown"
    }

    # Add all users to the DataGridView with filtering
    function Set-AllUsers {
        param (
            [string]$SearchText = "",
            [bool]$ShowExternal = $true
        )
        
        $dataGridView.SuspendLayout()
        $dataGridView.Rows.Clear()
        
        $filteredUsers = $UserList
        
        # Apply search filter if search text exists and isn't the placeholder
        if ($SearchText -ne "" -and $SearchText -ne "Search by name or email...") {
            $searchText = $SearchText.ToLower()
            $filteredUsers = $filteredUsers | Where-Object { 
                $_.displayName -like "*$searchText*" -or 
                $_.mail -like "*$searchText*" 
            }
        }
        
        # Apply external user filter
        if (-not $ShowExternal) {
            $filteredUsers = $filteredUsers | Where-Object { -not $_.isExternal }
        }
        
        foreach ($user in $filteredUsers) {
            $isChecked = $false
            if ($checkedUsers.ContainsKey($user.id)) {
                $isChecked = $checkedUsers[$user.id]
            }
            
            # Display external indicator
            $externalIndicator = if ($user.isExternal) { "Yes" } else { "" }
            
            $rowIdx = $dataGridView.Rows.Add($isChecked, $user.displayName, $user.mail, $externalIndicator, $user.id)
            
            # Highlight external users with a different color
            if ($user.isExternal) {
                $dataGridView.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 244, 230)
            }
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    }

    # Event handlers with improved performance
    $dataGridView.add_CellValueChanged({
        param($sender, $e)
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $row = $dataGridView.Rows[$e.RowIndex]
            $userId = $row.Cells[4].Value  # Updated index for ID
            $checkedUsers[$userId] = $row.Cells[0].Value
            Update-SelectionCount
        }
    })

    $dataGridView.add_CurrentCellDirtyStateChanged({
        if ($dataGridView.IsCurrentCellDirty -and $dataGridView.CurrentCell.ColumnIndex -eq 0) {
            $dataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    # Enhanced the checkbox appearance using CellPainting event
    $dataGridView.add_CellPainting({
        param($sender, $e)
        
        # Only customize the checkbox column
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $e.PaintBackground($e.CellBounds, $true)
            
            $checkboxSize = 16
            $x = $e.CellBounds.X + ($e.CellBounds.Width - $checkboxSize) / 2
            $y = $e.CellBounds.Y + ($e.CellBounds.Height - $checkboxSize) / 2
            
            $checkboxRect = New-Object System.Drawing.Rectangle($x, $y, $checkboxSize, $checkboxSize)
            
            $isChecked = $e.Value -eq $true
            
            $borderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
            $fillColor = [System.Drawing.Color]::White
            $checkColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            
            if ($isChecked) {
                $fillColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
                $borderColor = [System.Drawing.Color]::FromArgb(0, 100, 200)
            }
            
            # Draw checkbox
            $pen = New-Object System.Drawing.Pen($borderColor, 1)
            $brush = New-Object System.Drawing.SolidBrush($fillColor)
            
            $e.Graphics.FillRectangle($brush, $checkboxRect)
            $e.Graphics.DrawRectangle($pen, $checkboxRect)
            
            # Draw check mark if checked
            if ($isChecked) {
                $checkPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2)
                $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
                $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
                
                # Draw checkmark
                $e.Graphics.DrawLine($checkPen, 
                    $x + 3, $y + 8, 
                    $x + 6, $y + 11)
                $e.Graphics.DrawLine($checkPen, 
                    $x + 6, $y + 11, 
                    $x + 13, $y + 4)
            }
            
            # Cleanup
            $pen.Dispose()
            $brush.Dispose()
            if ($isChecked) { $checkPen.Dispose() }
            
            $e.Handled = $true
        }
    })

    # Select All button event
    $selectAllButton.Add_Click({
        $dataGridView.SuspendLayout()
        
        # Check all visible items
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $dataGridView.Rows[$i].Cells[0].Value = $true
            $userId = $dataGridView.Rows[$i].Cells[4].Value  # Updated index for ID
            $checkedUsers[$userId] = $true
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    })

    # Deselect All button event
    $deselectAllButton.Add_Click({
        $dataGridView.SuspendLayout()
        
        # Uncheck all visible items
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $dataGridView.Rows[$i].Cells[0].Value = $false
            $userId = $dataGridView.Rows[$i].Cells[4].Value  # Updated index for ID
            $checkedUsers[$userId] = $false
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    })

# Toggle external users button event
$toggleExternalButton.Add_Click({
    # Toggle the visibility state
    $script:showExternalUsers = -not $script:showExternalUsers
    
    # Update button text based on state
    if ($script:showExternalUsers) {
        $toggleExternalButton.Text = "Hide External Users"
    } else {
        $toggleExternalButton.Text = "Show External Users"
    }
    
    # Force button to stay enabled and visible
    $toggleExternalButton.Enabled = $true
    $toggleExternalButton.Visible = $true
    
    # Apply background color (ensure it's not gray)
    if ($script:showExternalUsers) {
        $toggleExternalButton.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    } else {
        # Slight color difference when showing state
        $toggleExternalButton.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    }
    
    # Refresh user list with current search and external filter settings
    $currentSearch = if ($searchBox.Text -eq "Search by name or email...") { "" } else { $searchBox.Text }
    Set-AllUsers -SearchText $currentSearch -ShowExternal $script:showExternalUsers
    
    # Force immediate redraw of the button
    $toggleExternalButton.Refresh()
})

    # Add search functionality - Modified to handle placeholder text
    $searchBox.add_TextChanged({
        # Skip search when showing placeholder text
        if ($searchBox.Text -eq "Search by name or email..." -or $null -eq $searchBox.Text) {
            Set-AllUsers -ShowExternal $showExternalUsers
            return
        }
        
        Set-AllUsers -SearchText $searchBox.Text -ShowExternal $showExternalUsers
    })

    # Add double-click to toggle checkbox
    $dataGridView.add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -gt 0) {
            $row = $dataGridView.Rows[$e.RowIndex]
            $currentValue = $row.Cells[0].Value
            $row.Cells[0].Value = -not $currentValue
        }
    })

    # Initialize with all users
    Set-AllUsers -ShowExternal $showExternalUsers

    $result = $form.ShowDialog()
    $selectedUsers = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($user in $UserList) {
            if ($checkedUsers.ContainsKey($user.id) -and $checkedUsers[$user.id]) {
                $selectedUsers += [PSCustomObject]@{
                    DisplayName = $user.displayName
                    Mail        = $user.mail
                    Id          = $user.id
                    IsExternal  = $user.isExternal
                }
            }
        }
    }
    return $selectedUsers
}

function Show-SiteSelectionForm {
    param (
        [Parameter(Mandatory=$true)]
        [array]$SiteList
    )

    # Try to load icon with error handling
    $Micon = $null
    try {
        $Micon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Microsoft_logo.ico")
    } catch {
        Write-Verbose "Unable to load icon: $_"
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select SharePoint Sites for Shortcut Creation"
    $form.Size = New-Object System.Drawing.Size(700, 520)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    if ($Micon) { $form.Icon = $Micon }
    $form.Font = New-Object System.Drawing.Font("Arial", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

    # Create a consistent style for buttons
    function Add-StyledButton {
        param (
            [string]$Text, 
            [int]$X, 
            [int]$Y, 
            [int]$Width = 100, 
            [int]$Height = 30,
            [bool]$Primary = $false
        )
        
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Size = New-Object System.Drawing.Size($Width, $Height)
        $button.Location = New-Object System.Drawing.Point($X, $Y)
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.Cursor = [System.Windows.Forms.Cursors]::Hand
        
        if ($Primary) {
            $button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            $button.ForeColor = [System.Drawing.Color]::White
            $button.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $button.FlatAppearance.BorderSize = 0
            $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
            $button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 90, 180)
        } else {
            $button.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
            $button.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
            $button.Font = New-Object System.Drawing.Font("Arial", 9)
            $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
            $button.FlatAppearance.BorderSize = 1
            $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
            $button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
        }
        
        return $button
    }

    # Create panel to hold the search elements (for better visual grouping)
    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Size = New-Object System.Drawing.Size(660, 40)
    $searchPanel.Location = New-Object System.Drawing.Point(10, 10)
    $searchPanel.BackColor = [System.Drawing.Color]::White
    $searchPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Create search icon using PictureBox
    $searchIconBox = New-Object System.Windows.Forms.PictureBox
    $searchIconBox.Size = New-Object System.Drawing.Size(20, 20)
    $searchIconBox.Location = New-Object System.Drawing.Point(15, 10)
    $searchIconBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

    # Load the search icon with error handling
    try {
        $searchIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Search_icon.ico")
        $searchIconBox.Image = $searchIcon.ToBitmap()
        $searchPanel.Controls.Add($searchIconBox)
    } catch {
        Write-Verbose "Unable to load search icon: $_"
        # If icon fails to load, don't add anything to maintain clean layout
    }

    # Enhanced search box with adjusted position
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Size = New-Object System.Drawing.Size(600, 25)
    $searchBox.Location = New-Object System.Drawing.Point(45, 8)  # Adjusted position for icon
    $searchBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $searchBox.BackColor = [System.Drawing.Color]::White
    $searchBox.ForeColor = [System.Drawing.Color]::Gray
    $searchBox.Text = "Search by site name..."
    $searchPanel.Controls.Add($searchBox)

    # Add event handlers for placeholder behavior
    $searchBox.Add_GotFocus({
        if ($this.Text -eq "Search by site name...") {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
        }
    })

    $searchBox.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = "Search by site name..."
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $form.Controls.Add($searchPanel)

    # Improved DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Size = New-Object System.Drawing.Size(660, 370)
    $dataGridView.Location = New-Object System.Drawing.Point(10, 60)
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.AllowUserToDeleteRows = $false
    $dataGridView.SelectionMode = 'FullRowSelect'
    $dataGridView.MultiSelect = $false
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.AutoSizeColumnsMode = 'Fill'
    $dataGridView.ScrollBars = 'Vertical'
    $dataGridView.BackgroundColor = [System.Drawing.Color]::White
    $dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $dataGridView.Font = New-Object System.Drawing.Font("Arial", 9)
    $dataGridView.GridColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $dataGridView.RowTemplate.Height = 30
    $dataGridView.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $dataGridView.RowsDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 249, 255)
    $dataGridView.RowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    $dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)

    # Create and configure columns
    $checkColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $checkColumn.HeaderText = "Select"
    $checkColumn.Width = 60
    $checkColumn.Name = "Include"
    $checkColumn.ReadOnly = $false  # Only this column is editable
    $checkColumn.FillWeight = 15
    $dataGridView.Columns.Add($checkColumn) | Out-Null

    # Helper function to create read-only columns
    function Add-ReadOnlyColumn {
        param (
            [string]$Name,
            [string]$HeaderText,
            [int]$FillWeight = 30
        )
        
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $Name
        $column.HeaderText = $HeaderText
        $column.ReadOnly = $true
        $column.FillWeight = $FillWeight
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $column.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
        return $column
    }

    $nameColumn = Add-ReadOnlyColumn -Name "DisplayName" -HeaderText "Site Name" -FillWeight 40

    # Create URL column as link column
    $urlColumn = New-Object System.Windows.Forms.DataGridViewLinkColumn
    $urlColumn.Name = "WebUrl"
    $urlColumn.HeaderText = "URL"
    $urlColumn.ReadOnly = $true
    $urlColumn.FillWeight = 60
    $urlColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $urlColumn.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    # Set link appearance
    $urlColumn.LinkColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $urlColumn.ActiveLinkColor = [System.Drawing.Color]::FromArgb(204, 0, 0)
    $urlColumn.VisitedLinkColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $urlColumn.TrackVisitedState = $false
    # Important: This needs to be FALSE for the URL text to appear
    $urlColumn.UseColumnTextForLinkValue = $false

    $idColumn = Add-ReadOnlyColumn -Name "Id" -HeaderText "ID" -FillWeight 1
    
    $dataGridView.Columns.Add($nameColumn) | Out-Null
    $dataGridView.Columns.Add($urlColumn) | Out-Null
    $dataGridView.Columns.Add($idColumn) | Out-Null
    
    # Hide the ID column as it's not typically needed for user display
    $dataGridView.Columns[3].Visible = $false

    # Style the header
    $dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $dataGridView.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dataGridView.ColumnHeadersHeight = 35
    $dataGridView.EnableHeadersVisualStyles = $false

    $form.Controls.Add($dataGridView)

    # Create styled buttons
    $selectAllButton = Add-StyledButton -Text "Select All" -X 10 -Y 440
    $form.Controls.Add($selectAllButton)

    $deselectAllButton = Add-StyledButton -Text "Deselect All" -X 120 -Y 440
    $form.Controls.Add($deselectAllButton)

    # Add a count indicator label
    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Size = New-Object System.Drawing.Size(130, 23)
    $countLabel.Location = New-Object System.Drawing.Point(230, 444)
    $countLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($countLabel)

    $okButton = Add-StyledButton -Text "OK" -X 580 -Y 440 -Width 100 -Height 35 -Primary $true
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $checkedSites = @{}

    # Function to update selection count
    function Update-SelectionCount {
        $selectedCount = ($checkedSites.Values | Where-Object { $_ -eq $true }).Count
        $totalCount = $dataGridView.Rows.Count
        $countLabel.Text = "$selectedCount of $totalCount sites"
    }

    # Add all sites to the DataGridView with filtering
    function Set-AllSites {
        param (
            [string]$SearchText = ""
        )
        
        $dataGridView.SuspendLayout()
        $dataGridView.Rows.Clear()
        
        $filteredSites = $SiteList
        
        # Apply search filter if search text exists and isn't the placeholder
        if ($SearchText -ne "" -and $SearchText -ne "Search by site name...") {
            $searchText = $SearchText.ToLower()
            $filteredSites = $filteredSites | Where-Object { 
                $_.displayName -like "*$searchText*" -or 
                $_.name -like "*$searchText*" 
            }
        }
        
        foreach ($site in $filteredSites) {
            $isChecked = $false
            if ($checkedSites.ContainsKey($site.id)) {
                $isChecked = $checkedSites[$site.id]
            }
            
            # Create a row directly
            $row = New-Object System.Windows.Forms.DataGridViewRow
            $row.CreateCells($dataGridView)
            
            # Set values for each cell
            $row.Cells[0].Value = $isChecked
            $row.Cells[1].Value = $site.displayName
            $row.Cells[2].Value = $site.webUrl
            $row.Cells[3].Value = $site.id
            
            $rowIdx = $dataGridView.Rows.Add($row)
            
            # Apply styling for root sites if needed
            if ($null -ne $site.root) {
                $dataGridView.Rows[$rowIdx].DefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
                $dataGridView.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 247, 255)
            }
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    }

    # Event handlers with improved performance
    $dataGridView.add_CellValueChanged({
        param($sender, $e)
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $row = $dataGridView.Rows[$e.RowIndex]
            $siteId = $row.Cells[3].Value  # Updated index for ID
            $checkedSites[$siteId] = $row.Cells[0].Value
            Update-SelectionCount
        }
    })

    $dataGridView.add_CurrentCellDirtyStateChanged({
        if ($dataGridView.IsCurrentCellDirty -and $dataGridView.CurrentCell.ColumnIndex -eq 0) {
            $dataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    # Enhanced checkbox appearance using CellPainting event
    $dataGridView.add_CellPainting({
        param($sender, $e)
        
        # Only customize the checkbox column
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $e.PaintBackground($e.CellBounds, $true)
            
            $checkboxSize = 16
            $x = $e.CellBounds.X + ($e.CellBounds.Width - $checkboxSize) / 2
            $y = $e.CellBounds.Y + ($e.CellBounds.Height - $checkboxSize) / 2
            
            $checkboxRect = New-Object System.Drawing.Rectangle($x, $y, $checkboxSize, $checkboxSize)
            
            $isChecked = $e.Value -eq $true
            
            $borderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
            $fillColor = [System.Drawing.Color]::White
            $checkColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            
            if ($isChecked) {
                $fillColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
                $borderColor = [System.Drawing.Color]::FromArgb(0, 100, 200)
            }
            
            # Draw checkbox
            $pen = New-Object System.Drawing.Pen($borderColor, 1)
            $brush = New-Object System.Drawing.SolidBrush($fillColor)
            
            $e.Graphics.FillRectangle($brush, $checkboxRect)
            $e.Graphics.DrawRectangle($pen, $checkboxRect)
            
            # Draw check mark if checked
            if ($isChecked) {
                $checkPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2)
                $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
                $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
                
                # Draw checkmark
                $e.Graphics.DrawLine($checkPen, 
                    $x + 3, $y + 8, 
                    $x + 6, $y + 11)
                $e.Graphics.DrawLine($checkPen, 
                    $x + 6, $y + 11, 
                    $x + 13, $y + 4)
            }
            
            # Cleanup
            $pen.Dispose()
            $brush.Dispose()
            if ($isChecked) { $checkPen.Dispose() }
            
            $e.Handled = $true
        }
    })

    # Add event handler for cell content click to handle URL clicks
    $dataGridView.add_CellContentClick({
        param($sender, $e)
        
        # Check if the clicked cell is in the URL column
        if ($e.ColumnIndex -eq 2 -and $e.RowIndex -ge 0) {
            try {
                $url = $dataGridView.Rows[$e.RowIndex].Cells[2].Value
                
                # Verify URL isn't empty
                if (-not [string]::IsNullOrEmpty($url)) {
                    # Open the URL in the default browser
                    Write-Verbose "Opening URL: $url"
                    Start-Process $url
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Could not open the URL: $($_.Exception.Message)",
                    "Error Opening URL",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    })

    # Select All button event
    $selectAllButton.Add_Click({
        $dataGridView.SuspendLayout()
        
        # Check all visible items
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $dataGridView.Rows[$i].Cells[0].Value = $true
            $siteId = $dataGridView.Rows[$i].Cells[3].Value  # Updated index for ID
            $checkedSites[$siteId] = $true
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    })

    # Deselect All button event
    $deselectAllButton.Add_Click({
        $dataGridView.SuspendLayout()
        
        # Uncheck all visible items
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $dataGridView.Rows[$i].Cells[0].Value = $false
            $siteId = $dataGridView.Rows[$i].Cells[3].Value  # Updated index for ID
            $checkedSites[$siteId] = $false
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    })

    # Add search functionality
    $searchBox.add_TextChanged({
        # Skip search when showing placeholder text
        if ($searchBox.Text -eq "Search by site name..." -or $null -eq $searchBox.Text) {
            Set-AllSites
            return
        }
        
        Set-AllSites -SearchText $searchBox.Text
    })

    # Add double-click to toggle checkbox
    $dataGridView.add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -gt 0) {
            $row = $dataGridView.Rows[$e.RowIndex]
            $currentValue = $row.Cells[0].Value
            $row.Cells[0].Value = -not $currentValue
        }
    })

    # Initialize with all sites
    Set-AllSites

    $result = $form.ShowDialog()
    $selectedSites = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($site in $SiteList) {
            if ($checkedSites.ContainsKey($site.id) -and $checkedSites[$site.id]) {
                $selectedSites += [PSCustomObject]@{
                    DisplayName = $site.displayName
                    Name        = $site.name
                    WebUrl      = $site.webUrl
                    Id          = $site.id
                }
            }
        }
    }
    return $selectedSites
}