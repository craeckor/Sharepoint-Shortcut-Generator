function Show-UserSelectionForm {
    [CmdletBinding()]
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
    
    # Filter out external users immediately
    $filteredUserList = $UserList | Where-Object { -not $_.isExternal }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Users for Shortcut Creation"
    $form.Size = New-Object System.Drawing.Size(700, 530)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
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

    # Add warning message about external users
    $warningPanel = New-Object System.Windows.Forms.Panel
    $warningPanel.Size = New-Object System.Drawing.Size(660, 30)
    $warningPanel.Location = New-Object System.Drawing.Point(10, 60)
    $warningPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 240)
    $warningPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $warningLabel = New-Object System.Windows.Forms.Label
    $warningLabel.Size = New-Object System.Drawing.Size(640, 25)
    $warningLabel.Location = New-Object System.Drawing.Point(10, 2)
    $warningLabel.Text = "External Users are not supported and therefore filtered out."
    $warningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $warningLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $warningLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $warningPanel.Controls.Add($warningLabel)

    $form.Controls.Add($warningPanel)

    # Improved DataGridView - adjust position to accommodate the warning panel
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Size = New-Object System.Drawing.Size(660, 340) # Reduced height to fit the warning panel
    $dataGridView.Location = New-Object System.Drawing.Point(10, 100) # Adjusted position below warning panel
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
    $idColumn = Add-ReadOnlyColumn -Name "Id" -HeaderText "ID" -FillWeight 1
    
    $dataGridView.Columns.Add($nameColumn) | Out-Null
    $dataGridView.Columns.Add($mailColumn) | Out-Null
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
    $selectAllButton = Add-StyledButton -Text "Select All" -X 10 -Y 450
    $form.Controls.Add($selectAllButton)

    $deselectAllButton = Add-StyledButton -Text "Deselect All" -X 120 -Y 450
    $form.Controls.Add($deselectAllButton)

    # Add a count indicator label
    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Size = New-Object System.Drawing.Size(280, 23)
    $countLabel.Location = New-Object System.Drawing.Point(230, 454)
    $countLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($countLabel)

    $okButton = Add-StyledButton -Text "OK" -X 580 -Y 450 -Width 100 -Height 35 -Primary $true
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
            [string]$SearchText = ""
        )
        
        $dataGridView.SuspendLayout()
        $dataGridView.Rows.Clear()
        
        $displayUsers = $filteredUserList
        
        # Apply search filter if search text exists and isn't the placeholder
        if ($SearchText -ne "" -and $SearchText -ne "Search by name or email...") {
            $searchText = $SearchText.ToLower()
            $displayUsers = $displayUsers | Where-Object { 
                $_.displayName -like "*$searchText*" -or 
                $_.userPrincipalName -like "*$searchText*" 
            }
        }
        
        foreach ($user in $displayUsers) {
            $isChecked = $false
            if ($checkedUsers.ContainsKey($user.id)) {
                $isChecked = $checkedUsers[$user.id]
            }
            
            $rowIdx = $dataGridView.Rows.Add($isChecked, $user.displayName, $user.userPrincipalName, $user.id)
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    }

    # Event handlers with improved performance
    $dataGridView.add_CellValueChanged({
        param($sender, $e)
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $row = $dataGridView.Rows[$e.RowIndex]
            $userId = $row.Cells[3].Value  # Updated index for ID
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
            $userId = $dataGridView.Rows[$i].Cells[3].Value  # Updated index for ID
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
            $userId = $dataGridView.Rows[$i].Cells[3].Value  # Updated index for ID
            $checkedUsers[$userId] = $false
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    })

    # Add search functionality - Modified to handle placeholder text
    $searchBox.add_TextChanged({
        # Skip search when showing placeholder text
        if ($searchBox.Text -eq "Search by name or email..." -or $null -eq $searchBox.Text) {
            Set-AllUsers
            return
        }
        
        Set-AllUsers -SearchText $searchBox.Text
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
    Set-AllUsers

    $result = $form.ShowDialog()
    $selectedUsers = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($user in $filteredUserList) {
            if ($checkedUsers.ContainsKey($user.id) -and $checkedUsers[$user.id]) {
                $selectedUsers += [PSCustomObject]@{
                    DisplayName = $user.displayName
                    Mail        = $user.userPrincipalName
                    Id          = $user.id
                    IsExternal  = $false # We know they're not external since we filtered
                }
            }
        }
    }
    return $selectedUsers
}

function Show-SiteSelectionForm {
    [CmdletBinding()]
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
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
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

function Show-FolderSelectionForm {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [array]$DriveList,
        
        [Parameter(Mandatory=$true)]
        [string]$AccessToken
    )

    # State variables for navigation and selection
    $script:currentPath = @{}
    $script:navigationHistory = @{}
    $script:checkedFolders = @{}
    $script:isLoading = $false
    $script:currentDriveId = $null
    $script:currentItemId = $null
    $script:currentItems = @()
    $script:selectedDrive = $null

    # Try to load icon with error handling
    $Micon = $null
    try {
        $Micon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Microsoft_logo.ico")
    } catch {
        Write-Verbose "Unable to load icon: $_"
    }

    # Try to load folder icon
    $folderIcon = $null
    try {
        $folderIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Folder_icon.ico")
    } catch {
        Write-Verbose "Unable to load folder icon: $_"
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select SharePoint Folders"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
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
            [bool]$Primary = $false,
            [bool]$Disabled = $false
        )
        
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Size = New-Object System.Drawing.Size($Width, $Height)
        $button.Location = New-Object System.Drawing.Point($X, $Y)
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.Cursor = [System.Windows.Forms.Cursors]::Hand
        $button.Enabled = -not $Disabled
        
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
        
        # If disabled, make it look disabled
        if ($Disabled) {
            $button.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
            $button.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
        }
        
        return $button
    }

    # Create panel for the navigation elements
    $navPanel = New-Object System.Windows.Forms.Panel
    $navPanel.Size = New-Object System.Drawing.Size(760, 40)
    $navPanel.Location = New-Object System.Drawing.Point(10, 10)
    $navPanel.BackColor = [System.Drawing.Color]::White
    $navPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Create drive selection dropdown
    $driveLabel = New-Object System.Windows.Forms.Label
    $driveLabel.Text = "Drive:"
    $driveLabel.Size = New-Object System.Drawing.Size(40, 24)
    $driveLabel.Location = New-Object System.Drawing.Point(10, 10)
    $driveLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $navPanel.Controls.Add($driveLabel)

    $driveDropdown = New-Object System.Windows.Forms.ComboBox
    $driveDropdown.Size = New-Object System.Drawing.Size(300, 24)
    $driveDropdown.Location = New-Object System.Drawing.Point(60, 8)
    $driveDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $driveDropdown.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $navPanel.Controls.Add($driveDropdown)

    # Add drives to the dropdown
    foreach ($drive in $DriveList) {
        $driveName = if ($drive.name) { $drive.name } else { "Documents" }
        $siteDisplayName = if ($drive.siteDisplayName) { $drive.siteDisplayName } else { "SharePoint" }
        $displayText = "$siteDisplayName - $driveName"
        
        # Add the drive ID as Tag to retrieve it later
        $driveItem = New-Object System.Windows.Forms.ComboBox
        $driveItem = $displayText
        $driveDropdown.Items.Add($driveItem) | Out-Null
    }

    # Back button with proper Unicode character
    $backButton = Add-StyledButton -Text "Back" -X 370 -Y 6 -Width 80 -Height 25 -Disabled $true
    $navPanel.Controls.Add($backButton)

    # Refresh button with proper Unicode character
    $refreshButton = Add-StyledButton -Text "Refresh" -X 460 -Y 6 -Width 80 -Height 25
    $navPanel.Controls.Add($refreshButton)

    # Add a button to select the current folder
    $selectCurrentButton = Add-StyledButton -Text "Select Current" -X 550 -Y 6 -Width 110 -Height 25
    $navPanel.Controls.Add($selectCurrentButton)

    $form.Controls.Add($navPanel)

    # Create a new path panel below the navigation panel
    $pathPanel = New-Object System.Windows.Forms.Panel
    $pathPanel.Size = New-Object System.Drawing.Size(760, 30)
    $pathPanel.Location = New-Object System.Drawing.Point(10, 60) # Place it right below the nav panel
    $pathPanel.BackColor = [System.Drawing.Color]::White
    $pathPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Add a more descriptive path label
    $pathDescLabel = New-Object System.Windows.Forms.Label
    $pathDescLabel.Text = "Current Path:"
    $pathDescLabel.Size = New-Object System.Drawing.Size(90, 24)
    $pathDescLabel.Location = New-Object System.Drawing.Point(10, 3)
    $pathDescLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $pathDescLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $pathPanel.Controls.Add($pathDescLabel)

    # Create a better path display with more room
    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Size = New-Object System.Drawing.Size(650, 24)
    $pathLabel.Location = New-Object System.Drawing.Point(100, 3)
    $pathLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $pathLabel.Text = "/ (Root)"
    $pathLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $pathPanel.Controls.Add($pathLabel)

    $form.Controls.Add($pathPanel)

    # Add search panel to folder selection form, after the path panel
    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Size = New-Object System.Drawing.Size(760, 40)
    $searchPanel.Location = New-Object System.Drawing.Point(10, 100) # Position below path panel
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
    $searchBox.Size = New-Object System.Drawing.Size(700, 25)
    $searchBox.Location = New-Object System.Drawing.Point(45, 8) # Adjusted position for icon
    $searchBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $searchBox.BackColor = [System.Drawing.Color]::White
    $searchBox.ForeColor = [System.Drawing.Color]::Gray
    $searchBox.Text = "Search folders..."
    $searchPanel.Controls.Add($searchBox)

    # Add event handlers for placeholder behavior
    $searchBox.Add_GotFocus({
        if ($this.Text -eq "Search folders...") {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
        }
    })

    $searchBox.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = "Search folders..."
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    # Add search functionality - handles filtering the visible folders
    $searchBox.add_TextChanged({
        # Skip search when showing placeholder text or during loading
        if ($script:isLoading -or ($searchBox.Text -eq "Search folders..." -or $null -eq $searchBox.Text)) {
            return
        }
        
        $searchText = $searchBox.Text.ToLower()
        $dataGridView.SuspendLayout()
        $dataGridView.Rows.Clear()
        
        # Get folder icon
        $folderBitmap = $null
        if ($folderIcon) {
            $folderBitmap = $folderIcon.ToBitmap()
        }
        
        # Filter current items based on search text
        $filteredItems = @()
        if ($searchText -ne "") {
            # Apply filter
            $filteredItems = $script:currentItems | Where-Object { 
                $null -ne $_.folder -and $_.name -like "*$searchText*" 
            }
        } else {
            # No filter - show all folders
            $filteredItems = $script:currentItems | Where-Object { $null -ne $_.folder }
        }
        
        # Add filtered items to grid
        foreach ($item in $filteredItems) {
            $isChecked = $false
            $folderKey = "$script:currentDriveId|$($item.id)"
            
            # Check if this folder is already selected
            if ($script:checkedFolders.ContainsKey($folderKey)) {
                $isChecked = $script:checkedFolders[$folderKey]
            }
            
            # Add folder to grid
            $childCount = if ($item.folder.childCount) { $item.folder.childCount } else { 0 }
            $size = if ($item.size) { Format-FileSize -Size $item.size } else { "0 B" }
            
            $row = $dataGridView.Rows.Add(
                $isChecked,
                $folderBitmap,
                $item.name,
                "Folder",
                $size,
                $childCount,
                $item.id
            )
            
            # Store checkbox state
            $script:checkedFolders[$folderKey] = $isChecked
        }
        
        $dataGridView.ResumeLayout()
        
        # Update selection count for currently visible items
        Update-SelectionCount
    })

    $form.Controls.Add($searchPanel)

    # Create loading panel with animation
    $loadingPanel = New-Object System.Windows.Forms.Panel
    $loadingPanel.Size = New-Object System.Drawing.Size(760, 350)
    $loadingPanel.Location = New-Object System.Drawing.Point(10, 150)
    $loadingPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $loadingPanel.Visible = $false

    # For the loading indicator, let's use a simple label-based approach instead of a GIF
    $loadingLabel = New-Object System.Windows.Forms.Label
    $loadingLabel.Text = "Loading folders..."
    $loadingLabel.Size = New-Object System.Drawing.Size(760, 400)
    $loadingLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $loadingLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $loadingLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $loadingPanel.Controls.Add($loadingLabel)
    $form.Controls.Add($loadingPanel)

    # Create the DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Size = New-Object System.Drawing.Size(760, 350)
    $dataGridView.Location = New-Object System.Drawing.Point(10, 150)
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
    $dataGridView.RowTemplate.Height = 36
    $dataGridView.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $dataGridView.RowsDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 249, 255)
    $dataGridView.RowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    $dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)

    # Create columns
    $checkColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $checkColumn.HeaderText = "Select"
    $checkColumn.Width = 60
    $checkColumn.Name = "Include"
    $checkColumn.ReadOnly = $false
    $checkColumn.FillWeight = 10
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

    # Create image column for folder icon
    $iconColumn = New-Object System.Windows.Forms.DataGridViewImageColumn
    $iconColumn.HeaderText = ""
    $iconColumn.Width = 30
    $iconColumn.Name = "Icon"
    $iconColumn.ReadOnly = $true
    $iconColumn.FillWeight = 5
    $iconColumn.ValueType = [System.Drawing.Image]
    $dataGridView.Columns.Add($iconColumn) | Out-Null

    $nameColumn = Add-ReadOnlyColumn -Name "Name" -HeaderText "Folder Name" -FillWeight 40
    $typeColumn = Add-ReadOnlyColumn -Name "Type" -HeaderText "Type" -FillWeight 15
    $sizeColumn = Add-ReadOnlyColumn -Name "Size" -HeaderText "Size" -FillWeight 15
    $itemCountColumn = Add-ReadOnlyColumn -Name "Items" -HeaderText "Items" -FillWeight 10
    $idColumn = Add-ReadOnlyColumn -Name "Id" -HeaderText "ID" -FillWeight 1
    
    $dataGridView.Columns.Add($nameColumn) | Out-Null
    $dataGridView.Columns.Add($typeColumn) | Out-Null
    $dataGridView.Columns.Add($sizeColumn) | Out-Null
    $dataGridView.Columns.Add($itemCountColumn) | Out-Null
    $dataGridView.Columns.Add($idColumn) | Out-Null
    
    # Hide the ID column
    $dataGridView.Columns[6].Visible = $false

    # Style the header
    $dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $dataGridView.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dataGridView.ColumnHeadersHeight = 35
    $dataGridView.EnableHeadersVisualStyles = $false

    $form.Controls.Add($dataGridView)

    # Create the bottom panel with buttons
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Size = New-Object System.Drawing.Size(760, 50)
    $bottomPanel.Location = New-Object System.Drawing.Point(10, 500)
    $bottomPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

    # Create buttons
    $selectAllButton = Add-StyledButton -Text "Select All" -X 10 -Y 10
    $bottomPanel.Controls.Add($selectAllButton)

    $deselectAllButton = Add-StyledButton -Text "Deselect All" -X 120 -Y 10
    $bottomPanel.Controls.Add($deselectAllButton)

    # Add a count indicator label
    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Size = New-Object System.Drawing.Size(280, 23)
    $countLabel.Location = New-Object System.Drawing.Point(230, 14)
    $countLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $countLabel.Text = "0 of 0 folders selected"
    $bottomPanel.Controls.Add($countLabel)

    $okButton = Add-StyledButton -Text "OK" -X 640 -Y 10 -Width 100 -Height 35 -Primary $true
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $bottomPanel.Controls.Add($okButton)

    $form.Controls.Add($bottomPanel)

    # Function to update selection count
    function Update-SelectionCount {
        # Count currently visible folders that are selected
        $visibleFoldersCount = $dataGridView.Rows.Count
        $visibleSelectedCount = 0
        
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $isChecked = $dataGridView.Rows[$i].Cells[0].Value -eq $true
            if ($isChecked) {
                $visibleSelectedCount++
            }
        }
        
        # Count total selected folders across all views
        $totalSelectedCount = ($script:checkedFolders.Values | Where-Object { $_ -eq $true }).Count
        
        # Check if the current folder is selected but not visible in grid (e.g., root folder)
        $currentFolderKey = "$script:currentDriveId|$script:currentItemId"
        $currentFolderSelected = $script:checkedFolders.ContainsKey($currentFolderKey) -and $script:checkedFolders[$currentFolderKey] -eq $true
        
        # Update the label to show both visible folders and total selected
        $countLabel.Text = "$visibleSelectedCount of $visibleFoldersCount folders shown | $totalSelectedCount total selected"
    }

    # Function to show loading overlay
    function Show-Loading {
        param([string]$message = "Loading folders...")
        
        $script:isLoading = $true
        $loadingLabel.Text = $message
        $loadingPanel.Visible = $true
        $dataGridView.Visible = $false
        $form.Refresh()
    }

    # Function to hide loading overlay
    function Hide-Loading {
        $script:isLoading = $false
        $loadingPanel.Visible = $false
        $dataGridView.Visible = $true
        $form.Refresh()
    }

    # Helper function to convert bytes to human-readable format
    function Format-FileSize {
        param([long]$Size)
        
        if ($Size -lt 1KB) {
            return "$Size B"
        }
        elseif ($Size -lt 1MB) {
            return "{0:N1} KB" -f ($Size / 1KB)
        }
        elseif ($Size -lt 1GB) {
            return "{0:N1} MB" -f ($Size / 1MB)
        }
        else {
            return "{0:N2} GB" -f ($Size / 1GB)
        }
    }

    # Function to load folder items
    function Sync-FolderItems {
        param(
            [string]$DriveId,
            [string]$ItemId,
            [bool]$IsRoot = $false
        )
        
        Show-Loading -message "Loading folders..."

        $searchBox.Text = "Search folders..."
        $searchBox.ForeColor = [System.Drawing.Color]::Gray
        
        try {
            $script:currentDriveId = $DriveId
            
            if ($IsRoot) {
                # Initialize or clear navigation history for this drive
                if (-not $script:navigationHistory.ContainsKey($DriveId)) {
                    $script:navigationHistory[$DriveId] = @()
                }
                
                # Get root folder ID
                $rootUrl = "$graphEndpoint/drives/$DriveId/root"
                $rootResponse = Send-GraphRequest -Method GET -Uri $rootUrl -AccessToken $AccessToken
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Failed to get root folder: $rootResponse"
                }
                $rootObject = $rootResponse | ConvertFrom-Json
                $script:currentItemId = $rootObject.id
                
                # Update path tracking
                $script:currentPath[$DriveId] = @{
                    "path" = "/"
                    "id" = $rootObject.id
                    "name" = "Root"
                }
                
                # Clear navigation history for this drive and add root
                $script:navigationHistory[$DriveId] = @(@{
                    "id" = $rootObject.id
                    "name" = "Root"
                    "path" = "/"
                })
                
                # Set the path label
                $pathLabel.Text = "/ (Root)"
                
                # Load children of root
                $childrenUrl = "$graphEndpoint/drives/$DriveId/items/$($rootObject.id)/children"
                $ItemId = $rootObject.id
            } else {
                # Load children of specified item
                $childrenUrl = "$graphEndpoint/drives/$DriveId/items/$ItemId/children"
                
                # Get parent item details for navigation path
                $itemUrl = "$graphEndpoint/drives/$DriveId/items/$ItemId"
                $itemResponse = Send-GraphRequest -Method GET -Uri $itemUrl -AccessToken $AccessToken
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Failed to get item details: $itemResponse"
                }
                $itemObject = $itemResponse | ConvertFrom-Json
                
                # If navigating to a new folder (not refreshing)
                if ($script:currentItemId -ne $ItemId) {
                    # Update current path - completely replacing the logic to prevent duplication
                    $pathParts = @()
                    
                    # Get the complete path from the parent reference if available
                    if ($itemObject.parentReference -and $itemObject.parentReference.path) {
                        try {
                            # The path often comes in format like "/drives/id/root:/path/to/folder"
                            # Split by ":" and take the second part if it exists
                            $parentPath = $itemObject.parentReference.path
                            if ($parentPath -match ":/(.+)") {
                                $pathParts += $Matches[1].Split("/") | Where-Object { $_ -ne "" }
                            }
                        } catch {
                            # If we can't parse the path, just use what we know
                            Write-Verbose "Could not parse parent path: $_"
                        }
                    }
                    
                    # Add the current folder name
                    $pathParts += $itemObject.name
                    
                    # Reconstruct the path
                    $newPath = "/" + ($pathParts -join "/")
                    
                    $script:currentPath[$DriveId] = @{
                        "path" = $newPath
                        "id" = $ItemId
                        "name" = $itemObject.name
                    }
                    
                    # Add to navigation history
                    $alreadyInHistory = $false
                    foreach ($item in $script:navigationHistory[$DriveId]) {
                        if ($item.id -eq $ItemId) {
                            $alreadyInHistory = $true
                            break
                        }
                    }
                    
                    if (-not $alreadyInHistory) {
                        $script:navigationHistory[$DriveId] += @{
                            "id" = $ItemId
                            "name" = $itemObject.name
                            "path" = $newPath
                        }
                    }
                }
                
                # Update path label
                $pathLabel.Text = $script:currentPath[$DriveId].path
            }
            
            # Store current item ID
            $script:currentItemId = $ItemId
            
            # Get children
            $response = Send-GraphRequest -Method GET -Uri $childrenUrl -AccessToken $AccessToken
            if ($LASTEXITCODE -ne 0) {
                Write-Error "Failed to get children: $response"
            }
            $items = ($response | ConvertFrom-Json).value
            
            # Store current items
            $script:currentItems = $items
            
            # Update back button state - this is the key fix
            # Since we store history in an array, we need at least 2 items (root + current) to enable back
            $backButton.Enabled = ($script:navigationHistory[$DriveId].Count -gt 1)
            # Force a refresh of the back button's appearance
            $backButton.BackColor = if ($backButton.Enabled) {
                [System.Drawing.Color]::FromArgb(255, 255, 255)  # Normal color when enabled
            } else {
                [System.Drawing.Color]::FromArgb(230, 230, 230)  # Disabled color
            }
            $backButton.ForeColor = if ($backButton.Enabled) {
                [System.Drawing.Color]::FromArgb(60, 60, 60)     # Normal text color
            } else {
                [System.Drawing.Color]::FromArgb(150, 150, 150)  # Disabled text color
            }
            $backButton.Refresh()
            
            # Clear and populate the grid
            $dataGridView.SuspendLayout()
            $dataGridView.Rows.Clear()
            
            # Convert folder icon to bitmap if available
            $folderBitmap = $null
            if ($folderIcon) {
                $folderBitmap = $folderIcon.ToBitmap()
            }
            
            # Add items to grid - folders only
            foreach ($item in $items | Where-Object { $null -ne $_.folder }) {
                $isChecked = $false
                $folderKey = "$DriveId|$($item.id)"
                
                # Check if this folder is already selected
                if ($script:checkedFolders.ContainsKey($folderKey)) {
                    $isChecked = $script:checkedFolders[$folderKey]
                }
                
                # Add folder to grid
                $childCount = if ($item.folder.childCount) { $item.folder.childCount } else { 0 }
                $size = if ($item.size) { Format-FileSize -Size $item.size } else { "0 B" }
                
                $row = $dataGridView.Rows.Add(
                    $isChecked,
                    $folderBitmap,
                    $item.name,
                    "Folder",
                    $size,
                    $childCount,
                    $item.id
                )
                
                # Store checkbox state
                $script:checkedFolders[$folderKey] = $isChecked
            }
            
            $dataGridView.ResumeLayout()
            
            # Update the selection count
            Update-SelectionCount
            
            # Update the select current button text after navigation
            $currentFolderKey = "$script:currentDriveId|$script:currentItemId"
            if ($script:checkedFolders.ContainsKey($currentFolderKey) -and $script:checkedFolders[$currentFolderKey] -eq $true) {
                $selectCurrentButton.Text = "Unselect Current"
            } else {
                $selectCurrentButton.Text = "Select Current"
            }

            Hide-Loading
        }
        catch {
            $errorMessage = "Error loading folders: $($_.Exception.Message)"
            Write-Error $errorMessage
            $loadingLabel.Text = $errorMessage
            Start-Sleep -Seconds 3

            # Update the select current button text based on selection state
            $currentFolderKey = "$DriveId|$script:currentItemId"
            if ($script:checkedFolders.ContainsKey($currentFolderKey) -and $script:checkedFolders[$currentFolderKey] -eq $true) {
                $selectCurrentButton.Text = "Unselect Current"
            } else {
                $selectCurrentButton.Text = "Select Current"
            }

            # Update the select current button text after navigation
            $currentFolderKey = "$script:currentDriveId|$script:currentItemId"
            if ($script:checkedFolders.ContainsKey($currentFolderKey) -and $script:checkedFolders[$currentFolderKey] -eq $true) {
                $selectCurrentButton.Text = "Unselect Current"
            } else {
                $selectCurrentButton.Text = "Select Current"
            }

            Hide-Loading
        }
    }

    # Function to navigate back
    function Switch-NavigateBack {
        if ($script:navigationHistory[$script:currentDriveId].Count -le 1) {
            return
        }
        
        # Remove current from history
        $historyCount = $script:navigationHistory[$script:currentDriveId].Count
        $script:navigationHistory[$script:currentDriveId] = $script:navigationHistory[$script:currentDriveId][0..($historyCount - 2)]
        
        # Get previous item
        $previous = $script:navigationHistory[$script:currentDriveId][-1]
        
        # Load previous folder
        if ($previous.id -eq $script:navigationHistory[$script:currentDriveId][0].id) {
            # We're going back to root
            Sync-FolderItems -DriveId $script:currentDriveId -IsRoot $true
        } else {
            # Going back to a subfolder
            $script:currentPath[$script:currentDriveId] = @{
                "path" = $previous.path
                "id" = $previous.id
                "name" = $previous.name
            }
            $pathLabel.Text = $previous.path
            Sync-FolderItems -DriveId $script:currentDriveId -ItemId $previous.id -IsRoot $false
        }
        
        # Update back button state after navigation
        $backButton.Enabled = ($script:navigationHistory[$script:currentDriveId].Count -gt 1)
        # Force a refresh of the back button's appearance
        $backButton.BackColor = if ($backButton.Enabled) {
            [System.Drawing.Color]::FromArgb(255, 255, 255)
        } else {
            [System.Drawing.Color]::FromArgb(230, 230, 230)
        }
        $backButton.ForeColor = if ($backButton.Enabled) {
            [System.Drawing.Color]::FromArgb(60, 60, 60)
        } else {
            [System.Drawing.Color]::FromArgb(150, 150, 150)
        }
        $backButton.Refresh()

        # Update the select current button text after navigation
        $currentFolderKey = "$script:currentDriveId|$script:currentItemId"
        if ($script:checkedFolders.ContainsKey($currentFolderKey) -and $script:checkedFolders[$currentFolderKey] -eq $true) {
            $selectCurrentButton.Text = "Unselect Current"
        } else {
            $selectCurrentButton.Text = "Select Current"
        }
    }

    # Function to select the current folder
    function Select-CurrentFolder {
        if ($null -ne $script:currentDriveId -and $null -ne $script:currentItemId) {
            $folderKey = "$script:currentDriveId|$script:currentItemId"
            $folderName = if ($script:currentPath[$script:currentDriveId].path -eq "/") { "Root" } else { $script:currentPath[$script:currentDriveId].name }
            
            # Check if folder is already selected
            if ($script:checkedFolders.ContainsKey($folderKey) -and $script:checkedFolders[$folderKey] -eq $true) {
                # Folder is already selected, so unselect it
                $script:checkedFolders[$folderKey] = $false
                
                # Update the button text
                $selectCurrentButton.Text = "Select Current"
                
                # Update any visible corresponding row
                for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                    if ($dataGridView.Rows[$i].Cells[6].Value -eq $script:currentItemId) {
                        $dataGridView.Rows[$i].Cells[0].Value = $false
                        break
                    }
                }
                
                # Update the counts
                Update-SelectionCount
                
                # Show success message
                [System.Windows.Forms.MessageBox]::Show(
                    "Current folder removed from selection: $folderName",
                    "Folder Unselected",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            } else {
                # Select the folder
                $script:checkedFolders[$folderKey] = $true
                
                # Update the button text
                $selectCurrentButton.Text = "Unselect Current"
                
                # Update any visible corresponding row
                for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                    if ($dataGridView.Rows[$i].Cells[6].Value -eq $script:currentItemId) {
                        $dataGridView.Rows[$i].Cells[0].Value = $true
                        break
                    }
                }
                
                # Update the counts
                Update-SelectionCount
                
                # Show success message
                [System.Windows.Forms.MessageBox]::Show(
                    "Current folder added to selection: $folderName",
                    "Folder Selected",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    }

    # Event handler for selecting a drive
    $driveDropdown.add_SelectedIndexChanged({
        if ($driveDropdown.SelectedIndex -ge 0) {
            $driveIndex = $driveDropdown.SelectedIndex
            $script:selectedDrive = $DriveList[$driveIndex].id
            
            # Load root of selected drive
            Sync-FolderItems -DriveId $script:selectedDrive -IsRoot $true
        }
    })

    # Event handler for back button
    $backButton.add_Click({
        Switch-NavigateBack
    })

    # Event handler for refresh button
    $refreshButton.add_Click({
        if ($null -ne $script:currentDriveId -and $null -ne $script:currentItemId) {
            # Just refresh the current folder - don't change the path
            $childrenUrl = "$graphEndpoint/drives/$script:currentDriveId/items/$script:currentItemId/children"
            Show-Loading -message "Refreshing folders..."
            
            
            try {
                # Get children
                $response = Send-GraphRequest -Method GET -Uri $childrenUrl -AccessToken $AccessToken
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Failed to get children: $response"
                }
                $items = ($response | ConvertFrom-Json).value
                
                # Store current items
                $script:currentItems = $items
                
                # Clear and populate the grid
                $dataGridView.SuspendLayout()
                $dataGridView.Rows.Clear()
                
                # Convert folder icon to bitmap if available
                $folderBitmap = $null
                if ($folderIcon) {
                    $folderBitmap = $folderIcon.ToBitmap()
                }
                
                # Add items to grid - folders only
                foreach ($item in $items | Where-Object { $null -ne $_.folder }) {
                    $isChecked = $false
                    $folderKey = "$script:currentDriveId|$($item.id)"
                    
                    # Check if this folder is already selected
                    if ($script:checkedFolders.ContainsKey($folderKey)) {
                        $isChecked = $script:checkedFolders[$folderKey]
                    }
                    
                    # Add folder to grid
                    $childCount = if ($item.folder.childCount) { $item.folder.childCount } else { 0 }
                    $size = if ($item.size) { Format-FileSize -Size $item.size } else { "0 B" }
                    
                    $row = $dataGridView.Rows.Add(
                        $isChecked,
                        $folderBitmap,
                        $item.name,
                        "Folder",
                        $size,
                        $childCount,
                        $item.id
                    )
                    
                    # Store checkbox state
                    $script:checkedFolders[$folderKey] = $isChecked
                }
                
                $dataGridView.ResumeLayout()
                
                # Update the selection count
                Update-SelectionCount

                # Update the select current button text based on selection state
                $folderKey = "$DriveId|$script:currentItemId"
                if ($script:checkedFolders.ContainsKey($folderKey) -and $script:checkedFolders[$folderKey] -eq $true) {
                    $selectCurrentButton.Text = "Unselect Current"
                } else {
                    $selectCurrentButton.Text = "Select Current"
                }
                
                # Update the select current button text after navigation
                $currentFolderKey = "$script:currentDriveId|$script:currentItemId"
                if ($script:checkedFolders.ContainsKey($currentFolderKey) -and $script:checkedFolders[$currentFolderKey] -eq $true) {
                    $selectCurrentButton.Text = "Unselect Current"
                } else {
                    $selectCurrentButton.Text = "Select Current"
                }

                Hide-Loading
            } catch {
                $errorMessage = "Error refreshing folders: $($_.Exception.Message)"
                Write-Error $errorMessage
                $loadingLabel.Text = $errorMessage
                Start-Sleep -Seconds 3

                # Update the select current button text based on selection state
                $folderKey = "$DriveId|$script:currentItemId"
                if ($script:checkedFolders.ContainsKey($folderKey) -and $script:checkedFolders[$folderKey] -eq $true) {
                    $selectCurrentButton.Text = "Unselect Current"
                } else {
                    $selectCurrentButton.Text = "Select Current"
                }

                # Update the select current button text after navigation
                $currentFolderKey = "$script:currentDriveId|$script:currentItemId"
                if ($script:checkedFolders.ContainsKey($currentFolderKey) -and $script:checkedFolders[$currentFolderKey] -eq $true) {
                    $selectCurrentButton.Text = "Unselect Current"
                } else {
                    $selectCurrentButton.Text = "Select Current"
                }

                Hide-Loading
            }
        } elseif ($null -ne $script:selectedDrive) {
            Sync-FolderItems -DriveId $script:selectedDrive -IsRoot $true
        }
    })

    # Event handler for the select current folder button
    $selectCurrentButton.add_Click({
        Select-CurrentFolder
    })

    # Event handlers for DataGridView
    $dataGridView.add_CellValueChanged({
        param($sender, $e)
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $row = $dataGridView.Rows[$e.RowIndex]
            $itemId = $row.Cells[6].Value
            $folderKey = "$script:currentDriveId|$itemId"
            $script:checkedFolders[$folderKey] = $row.Cells[0].Value
            Update-SelectionCount
        }
    })

    $dataGridView.add_CurrentCellDirtyStateChanged({
        if ($dataGridView.IsCurrentCellDirty -and $dataGridView.CurrentCell.ColumnIndex -eq 0) {
            $dataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    # Event handler for double-click on a row to navigate into folder
    $dataGridView.add_CellDoubleClick({
        param($sender, $e)
        # Allow double-clicking on any cell except the checkbox column (column 0)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ne 0) {
            $itemId = $dataGridView.Rows[$e.RowIndex].Cells[6].Value
            Sync-FolderItems -DriveId $script:currentDriveId -ItemId $itemId -IsRoot $false
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

    # Select All button event
    $selectAllButton.Add_Click({
        $dataGridView.SuspendLayout()
        
        # Check all visible items
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $dataGridView.Rows[$i].Cells[0].Value = $true
            $itemId = $dataGridView.Rows[$i].Cells[6].Value
            $folderKey = "$script:currentDriveId|$itemId"
            $script:checkedFolders[$folderKey] = $true
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
            $itemId = $dataGridView.Rows[$i].Cells[6].Value
            $folderKey = "$script:currentDriveId|$itemId"
            $script:checkedFolders[$folderKey] = $false
        }
        
        $dataGridView.ResumeLayout()
        Update-SelectionCount
    })

    # Set the first drive as selected if available
    if ($DriveList.Count -gt 0) {
        # We'll set the selection after the form is loaded to avoid errors
        $form.add_Shown({
            if ($driveDropdown.Items.Count -gt 0) {
                $driveDropdown.SelectedIndex = 0
            }
        })
    }

    # Initialize the form and show
    $result = $form.ShowDialog()
    
    # Process results
    $selectedFolders = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($key in $script:checkedFolders.Keys) {
            if ($script:checkedFolders[$key]) {
                $parts = $key -split '\|'
                $driveId = $parts[0]
                $folderId = $parts[1]
                
                # Get drive info
                $drive = $DriveList | Where-Object { $_.id -eq $driveId } | Select-Object -First 1

                try {
                    # Get folder info - we need to make an API call for each selected folder
                    $folderUrl = "$graphEndpoint/drives/$driveId/items/$folderId"
                    $folderResponse = Send-GraphRequest -Method GET -Uri $folderUrl -AccessToken $AccessToken
                    if ($LASTEXITCODE -ne 0) {
                        Write-Error "Error retrieving folder details: $($folderResponse)"
                    }
                    $folder = $folderResponse | ConvertFrom-Json
                } catch {
                    Write-Error "Error retrieving folder details: $($_.Exception.Message)"
                    Pause
                    exit 1
                }
                
                # Add to result
                $selectedFolders += [PSCustomObject]@{
                    DriveId = $driveId
                    DriveName = if ($drive.name) { $drive.name } else { "Documents" }
                    SiteDisplayName = if ($drive.siteDisplayName) { $drive.siteDisplayName } else { "SharePoint" }
                    FolderId = $folderId
                    FolderName = $folder.name
                    WebUrl = $folder.webUrl
                    Path = if ($folder.parentReference.path) { $folder.parentReference.path } else { "/" }
                    webId = $DriveList | Where-Object { $_.id -eq $driveId } | Select-Object -ExpandProperty webId
                    siteId = $DriveList | Where-Object { $_.id -eq $driveId } | Select-Object -ExpandProperty siteId
                    siteWebUrl = $DriveList | Where-Object { $_.id -eq $driveId } | Select-Object -ExpandProperty webUrl
                    domainSiteId = $DriveList | Where-Object { $_.id -eq $driveId } | Select-Object -ExpandProperty domainSiteId
                }
            }
        }
    }
    
    return $selectedFolders
}

function Show-FolderNameEditForm {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [array]$SelectedFolders
    )

    # Try to load icon with error handling
    $Micon = $null
    try {
        $Micon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Microsoft_logo.ico")
    } catch {
        Write-Verbose "Unable to load icon: $_"
    }
    
    # Try to load folder icon
    $folderIcon = $null
    try {
        $folderIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Folder_icon.ico")
    } catch {
        Write-Verbose "Unable to load folder icon: $_"
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Edit Folder Names"
    $form.Size = New-Object System.Drawing.Size(800, 550)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
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

    # Top instruction panel
    $instructionPanel = New-Object System.Windows.Forms.Panel
    $instructionPanel.Size = New-Object System.Drawing.Size(760, 60)
    $instructionPanel.Location = New-Object System.Drawing.Point(10, 10)
    $instructionPanel.BackColor = [System.Drawing.Color]::FromArgb(225, 240, 250)
    $instructionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $instructionLabel = New-Object System.Windows.Forms.Label
    $instructionLabel.Size = New-Object System.Drawing.Size(740, 50)
    $instructionLabel.Location = New-Object System.Drawing.Point(10, 5)
    $instructionLabel.Text = "Note: Disable prefix before editing names, then re-enable to apply consistently. Edit the displayed name for each folder as needed. These names will be used for the shortcuts created. If you want to keep the original name, simply leave it unchanged."
    $instructionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $instructionPanel.Controls.Add($instructionLabel)

    $form.Controls.Add($instructionPanel)

    # Create the DataGridView to hold folders and editable names
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Size = New-Object System.Drawing.Size(760, 370)
    $dataGridView.Location = New-Object System.Drawing.Point(10, 80)
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
    $dataGridView.RowTemplate.Height = 36
    $dataGridView.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $dataGridView.RowsDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 249, 255)
    $dataGridView.RowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    $dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)

    # Helper function to create columns
    function Add-Column {
        param (
            [string]$Name,
            [string]$HeaderText,
            [int]$FillWeight = 30,
            [bool]$ReadOnly = $true
        )
        
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $Name
        $column.HeaderText = $HeaderText
        $column.ReadOnly = $ReadOnly
        $column.FillWeight = $FillWeight
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $column.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
        return $column
    }

    # Create image column for folder icon
    $iconColumn = New-Object System.Windows.Forms.DataGridViewImageColumn
    $iconColumn.HeaderText = ""
    $iconColumn.Width = 30
    $iconColumn.Name = "Icon"
    $iconColumn.ReadOnly = $true
    $iconColumn.FillWeight = 5
    $iconColumn.ValueType = [System.Drawing.Image]
    $dataGridView.Columns.Add($iconColumn) | Out-Null

    # Original folder name column (read-only)
    $originalNameColumn = Add-Column -Name "OriginalName" -HeaderText "Original Folder Name" -FillWeight 30
    $dataGridView.Columns.Add($originalNameColumn) | Out-Null

    # Editable display name column (user can change this)
    $displayNameColumn = Add-Column -Name "DisplayName" -HeaderText "Display Name (Editable)" -FillWeight 40 -ReadOnly $false
    $dataGridView.Columns.Add($displayNameColumn) | Out-Null

    # Site name column (read-only)
    $siteColumn = Add-Column -Name "SiteName" -HeaderText "Site" -FillWeight 25
    $dataGridView.Columns.Add($siteColumn) | Out-Null

    # Hidden columns for other folder properties
    $driveIdColumn = Add-Column -Name "DriveId" -HeaderText "DriveId" -FillWeight 1
    $driveNameColumn = Add-Column -Name "DriveName" -HeaderText "DriveName" -FillWeight 1
    $folderIdColumn = Add-Column -Name "FolderId" -HeaderText "FolderId" -FillWeight 1
    $webUrlColumn = Add-Column -Name "WebUrl" -HeaderText "WebUrl" -FillWeight 1
    $pathColumn = Add-Column -Name "Path" -HeaderText "Path" -FillWeight 1
    $webIdColumn = Add-Column -Name "webId" -HeaderText "webId" -FillWeight 1
    $eTagColumn = Add-Column -Name "eTag" -HeaderText "eTag" -FillWeight 1
    $eTagListColumn = Add-Column -Name "eTagList" -HeaderText "eTagList" -FillWeight 1
    $siteWebUrlColumn = Add-Column -Name "siteWebUrl" -HeaderText "siteWebUrl" -FillWeight 1
    $domainSiteIdColumn = Add-Column -Name "domainSiteId" -HeaderText "domainSiteId" -FillWeight 1

    $dataGridView.Columns.Add($driveIdColumn) | Out-Null
    $dataGridView.Columns.Add($driveNameColumn) | Out-Null
    $dataGridView.Columns.Add($folderIdColumn) | Out-Null
    $dataGridView.Columns.Add($webUrlColumn) | Out-Null
    $dataGridView.Columns.Add($pathColumn) | Out-Null
    $dataGridView.Columns.Add($webIdColumn) | Out-Null
    $dataGridView.Columns.Add($eTagColumn) | Out-Null
    $dataGridView.Columns.Add($eTagListColumn) | Out-Null
    $dataGridView.Columns.Add($siteWebUrlColumn) | Out-Null
    $dataGridView.Columns.Add($domainSiteIdColumn) | Out-Null

    # Hide the columns we don't need to display
    $dataGridView.Columns[4].Visible = $false
    $dataGridView.Columns[5].Visible = $false
    $dataGridView.Columns[6].Visible = $false
    $dataGridView.Columns[7].Visible = $false
    $dataGridView.Columns[8].Visible = $false
    $dataGridView.Columns[9].Visible = $false
    $dataGridView.Columns[10].Visible = $false
    $dataGridView.Columns[11].Visible = $false
    $dataGridView.Columns[12].Visible = $false
    $dataGridView.Columns[13].Visible = $false

    # Style the header
    $dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $dataGridView.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dataGridView.ColumnHeadersHeight = 35
    $dataGridView.EnableHeadersVisualStyles = $false

    $form.Controls.Add($dataGridView)

    # Convert folder icon to bitmap if available
    $folderBitmap = $null
    if ($folderIcon) {
        $folderBitmap = $folderIcon.ToBitmap()
    }

    # Add all selected folders to the grid
    foreach ($folder in $SelectedFolders) {
        $dataGridView.Rows.Add(
            $folderBitmap,
            $folder.FolderName,
            $folder.FolderName,  # Initially set display name to original name
            $folder.SiteDisplayName,
            $folder.DriveId,
            $folder.DriveName,
            $folder.FolderId,
            $folder.WebUrl,
            $folder.Path,
            $folder.webId,
            $folder.eTag,
            $folder.eTagList,
            $folder.siteWebUrl,
            $folder.domainSiteId
        ) | Out-Null
    }

    # Bottom panel to hold buttons and count information
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Size = New-Object System.Drawing.Size(760, 50)
    $bottomPanel.Location = New-Object System.Drawing.Point(10, 460)
    $bottomPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

    # Button for resetting all display names to original names
    $resetAllButton = Add-StyledButton -Text "Reset All Names" -X 10 -Y 15 -Width 150
    $bottomPanel.Controls.Add($resetAllButton)

    # Create a styled checkbox panel
    $checkboxPanel = New-Object System.Windows.Forms.Panel
    $checkboxPanel.Size = New-Object System.Drawing.Size(200, 30)
    $checkboxPanel.Location = New-Object System.Drawing.Point(170, 15)
    $checkboxPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $checkboxPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Create checkbox for toggling site name prefix
    $prefixCheckbox = New-Object System.Windows.Forms.CheckBox
    $prefixCheckbox.Text = "Add site name as prefix"
    $prefixCheckbox.Size = New-Object System.Drawing.Size(170, 23)
    $prefixCheckbox.Location = New-Object System.Drawing.Point(10, 4)
    $prefixCheckbox.Checked = $true # Enable by default
    $prefixCheckbox.Font = New-Object System.Drawing.Font("Arial", 9)
    $prefixCheckbox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $prefixCheckbox.BackColor = [System.Drawing.Color]::Transparent
    $checkboxPanel.Controls.Add($prefixCheckbox)
    $bottomPanel.Controls.Add($checkboxPanel)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Location = New-Object System.Drawing.Point(380, 15)
    $countLabel.Size = New-Object System.Drawing.Size(225, 28)
    $countLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $countLabel.Text = "$($SelectedFolders.Count) folders selected for shortcut creation"
    $bottomPanel.Controls.Add($countLabel)

    # OK button
    $okButton = Add-StyledButton -Text "OK" -X 650 -Y 10 -Width 100 -Height 35 -Primary $true
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $bottomPanel.Controls.Add($okButton)

    $form.Controls.Add($bottomPanel)

    # Function to apply site name prefixes to all display names
    function Set-SitePrefixes {
        $dataGridView.SuspendLayout()
        
        # Apply site name prefix to all display names
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $originalName = $dataGridView.Rows[$i].Cells["OriginalName"].Value
            $siteName = $dataGridView.Rows[$i].Cells["SiteName"].Value
            $currentDisplayName = $dataGridView.Rows[$i].Cells["DisplayName"].Value
            
            # Check if prefix already exists to avoid duplicates
            if (-not $currentDisplayName.StartsWith("$siteName - ")) {
                $dataGridView.Rows[$i].Cells["DisplayName"].Value = "$siteName - $currentDisplayName"
            }
        }
        
        $dataGridView.ResumeLayout()
    }

    # Function to remove site name prefixes from all display names
    function Remove-SitePrefixes {
        $dataGridView.SuspendLayout()
        
        # Remove site name prefix from all display names
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $originalName = $dataGridView.Rows[$i].Cells["OriginalName"].Value
            $siteName = $dataGridView.Rows[$i].Cells["SiteName"].Value
            $currentDisplayName = $dataGridView.Rows[$i].Cells["DisplayName"].Value
            
            # If display name starts with site name prefix, remove it
            if ($currentDisplayName.StartsWith("$siteName - ")) {
                $dataGridView.Rows[$i].Cells["DisplayName"].Value = $currentDisplayName.Substring(($siteName + " - ").Length)
            }
        }
        
        $dataGridView.ResumeLayout()
    }

    # Handle checkbox state change
    $prefixCheckbox.Add_CheckedChanged({
        if ($prefixCheckbox.Checked) {
            Set-SitePrefixes
        } else {
            Remove-SitePrefixes
        }
    })

    # Apply site prefixes by default when form loads
    $form.add_Shown({
        if ($prefixCheckbox.Checked) {
            Set-SitePrefixes
        }
    })

    # Event handler for the reset all button
    $resetAllButton.Add_Click({
        $dataGridView.SuspendLayout()
        
        # Reset all display names to original names
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $originalName = $dataGridView.Rows[$i].Cells["OriginalName"].Value
            $dataGridView.Rows[$i].Cells["DisplayName"].Value = $originalName
        }
        
        $dataGridView.ResumeLayout()

        # Re-apply prefixes if checkbox is checked
        if ($prefixCheckbox.Checked) {
            Set-SitePrefixes
        }
    })

    # Event handler for double-click on a cell to reset just that row
    $dataGridView.add_CellDoubleClick({
        param($sender, $e)
        # Only allow double-click on display name column
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq $dataGridView.Columns["DisplayName"].Index) {
            $originalName = $dataGridView.Rows[$e.RowIndex].Cells["OriginalName"].Value
            $dataGridView.Rows[$e.RowIndex].Cells["DisplayName"].Value = $originalName
        }
    })

    # Show the form
    $result = $form.ShowDialog()
    
    # Process results
    [array]$editedFolders = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Verbose "Processing form results - OK button clicked"
        # Create a mapping of edited display names (only get what was changed in the grid)
        $displayNameMap = @{}
        
        Write-Verbose "DataGridView has $($dataGridView.RowCount) rows"
        
        # Loop through grid rows to get edited display names keyed by a composite of driveId + folderId
        for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
            # Skip the "new row" that's empty at the bottom
            if ($dataGridView.Rows[$i].IsNewRow) { 
                Write-Verbose "Skipping new row at index $i"
                continue 
            }
            
            $driveId = $dataGridView.Rows[$i].Cells["DriveId"].Value
            $folderId = $dataGridView.Rows[$i].Cells["FolderId"].Value
            $displayName = $dataGridView.Rows[$i].Cells["DisplayName"].Value
            $originalName = $dataGridView.Rows[$i].Cells["OriginalName"].Value
            
            Write-Verbose "Row $i - DriveId: $driveId, FolderId: $folderId"
            Write-Verbose "Row $i - Original name: $originalName, Display name: $displayName"
            
            # Create a unique key from driveId and folderId
            $key = "$driveId|$folderId"
            
            # Store the edited display name
            $displayNameMap[$key] = $displayName
            Write-Verbose "Added to map with key: $key = $displayName"
        }
        
        Write-Verbose "Display name map created with $($displayNameMap.Count) entries"
        Write-Verbose "Original SelectedFolders collection has $($SelectedFolders.Count) folders"
        
        # Now create result objects using original data plus edited display names
        foreach ($folder in $SelectedFolders) {
            # Create the same key to look up display name
            $key = "$($folder.DriveId)|$($folder.FolderId)"
            Write-Verbose "Processing folder with key: $key"
            
            # Get the edited display name or fallback to original
            $displayName = if ($displayNameMap.ContainsKey($key)) { 
                Write-Verbose "Found edited name in map: $($displayNameMap[$key])"
                $displayNameMap[$key] 
            } else { 
                Write-Verbose "No edited name found, using original: $($folder.FolderName)"
                $folder.FolderName 
            }
            
            Write-Verbose "Creating result object with display name: $displayName"
            
            # Create a new result object with all original properties plus the edited display name
            if ($folder.DriveId -and $folder.FolderId) {
                $newFolder = [PSCustomObject]@{
                    DriveId = $folder.DriveId
                    DriveName = $folder.DriveName
                    SiteDisplayName = $folder.SiteDisplayName
                    FolderId = $folder.FolderId
                    FolderName = $folder.FolderName
                    DisplayName = $displayName
                    WebUrl = $folder.WebUrl
                    Path = $folder.Path
                    webId = $folder.webId
                    siteId = $folder.siteId
                    eTag = $folder.eTag
                    eTagList = $folder.eTagList
                    siteWebUrl = $folder.siteWebUrl
                    domainSiteId = $folder.domainSiteId
                }
            $editedFolders += $newFolder
            Write-Verbose "Added folder to result collection - DriveId: $($newFolder.DriveId), FolderId: $($newFolder.FolderId)"
            }
        }
        
        Write-Verbose "Created $($editedFolders.Count) result objects"
    }
    else {
        Write-Verbose "User canceled the operation, returning empty collection"
    }

    Write-Verbose "Returning result collection with $($editedFolders.Count) items"
    return ,[array]$editedFolders
}