 # Requires PowerShell 5.0 or later

 # Function to check and install required modules
 function Ensure-Module {
     param(
         [string]$ModuleName
     )
     if (!(Get-Module -ListAvailable -Name $ModuleName)) {
         Write-Host "Module '$ModuleName' not found. Installing..."
         try {
             Install-Module -Name $ModuleName -Force -AllowClobber -ErrorAction Stop
             Write-Host "Module '$ModuleName' installed successfully."
         } catch {
             Write-Error "Failed to install module '$ModuleName'. Error: $_"
             exit
         }
     } else {
         Write-Host "Module '$ModuleName' is already installed."
     }
 }

 # Function to check and add required assemblies
 function Ensure-Assembly {
     param(
         [string]$AssemblyName
     )
     try {
         Add-Type -AssemblyName $AssemblyName -ErrorAction Stop
         Write-Host "Assembly '$AssemblyName' loaded successfully."
     } catch {
         Write-Error "Failed to load assembly '$AssemblyName'. Ensure .NET Framework is installed. Error: $_"
         exit
     }
 }

 # Check and import required modules
 $requiredModules = @('Az.Accounts', 'Az.Resources')
 foreach ($module in $requiredModules) {
     Ensure-Module -ModuleName $module
     Import-Module $module -ErrorAction Stop
 }

 # Check and add required assemblies
 $requiredAssemblies = @('System.Windows.Forms', 'System.Windows.Forms.DataVisualization')
 foreach ($assembly in $requiredAssemblies) {
     Ensure-Assembly -AssemblyName $assembly
 }

 # Global variable to store the selected analysts file path
 $Global:AnalystsFilePath = ""

 # Function to perform the permission gap analysis
 function Perform-GapAnalysis {
     param(
         [System.Windows.Forms.DataGridView]$dataGridView,
         [System.Windows.Forms.DataVisualization.Charting.Chart]$chart,
         [System.Windows.Forms.Label]$statusLabel
     )

     # Clear previous data
     $dataGridView.Rows.Clear()
     $dataGridView.Columns.Clear()
     $chart.Series["MissingPermissions"].Points.Clear()
     $statusLabel.Text = ""

     # Check if AnalystsFilePath is set
     if ([string]::IsNullOrEmpty($Global:AnalystsFilePath) -or -not (Test-Path -Path $Global:AnalystsFilePath)) {
         [System.Windows.Forms.MessageBox]::Show("Please select a valid analysts list file before starting the analysis.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
         return
     }

     # Define required roles
     $requiredRoles = @("Azure Sentinel Reader", "Azure Sentinel Responder", "Security Reader")

     # Load analysts list from the selected file
     try {
         $securityAnalysts = Get-Content -Path $Global:AnalystsFilePath -ErrorAction Stop
     } catch {
         [System.Windows.Forms.MessageBox]::Show("Failed to read analysts list file. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
         return
     }

     # Initialize data table for the DataGridView
     $dataTable = New-Object System.Data.DataTable
     $dataTable.Columns.Add("TenantId") | Out-Null
     $dataTable.Columns.Add("SubscriptionId") | Out-Null
     $dataTable.Columns.Add("Analyst") | Out-Null
     $dataTable.Columns.Add("AssignedRoles") | Out-Null
     $dataTable.Columns.Add("MissingRoles") | Out-Null

     # Initialize a hashtable to count missing permissions per analyst
     $missingPermissionsCount = @{}

     # Connect to Azure account
     try {
         $statusLabel.Text = "Connecting to Azure..."
         $form.Refresh()
         Connect-AzAccount -ErrorAction Stop
     } catch {
         [System.Windows.Forms.MessageBox]::Show("Failed to connect to Azure account. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
         $statusLabel.Text = ""
         return
     }

     # Get list of tenants accessible via Azure Lighthouse
     try {
         $statusLabel.Text = "Retrieving tenant list..."
         $form.Refresh()
         $tenants = Get-AzTenant -ErrorAction Stop
     } catch {
         [System.Windows.Forms.MessageBox]::Show("Failed to retrieve tenant list. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
         $statusLabel.Text = ""
         return
     }

     $tenantCount = $tenants.Count
     $currentTenant = 0

     foreach ($tenant in $tenants) {
         $currentTenant++
         $statusLabel.Text = "Processing tenant $currentTenant of $tenantCount..."
         $form.Refresh()

         try {
             # Set context to the tenant
             Set-AzContext -TenantId $tenant.TenantId -ErrorAction Stop

             # Get all subscriptions in the tenant
             $subscriptions = Get-AzSubscription -TenantId $tenant.TenantId -ErrorAction Stop

             foreach ($subscription in $subscriptions) {
                 $statusLabel.Text = "Processing subscription $($subscription.Name) in tenant $currentTenant of $tenantCount..."
                 $form.Refresh()

                 # Set context to the subscription
                 Set-AzContext -SubscriptionId $subscription.Id -TenantId $tenant.TenantId -ErrorAction Stop

                 # Retrieve all role assignments in the subscription
                 $allRoleAssignments = Get-AzRoleAssignment -Scope "/subscriptions/$($subscription.Id)" -ErrorAction Stop

                 # Group role assignments by analyst
                 $assignmentsByAnalyst = $allRoleAssignments | Group-Object -Property SignInName

                 foreach ($analyst in $securityAnalysts) {
                     $assignedRoles = @()

                     $assignments = $assignmentsByAnalyst | Where-Object { $_.Name -eq $analyst }

                     if ($assignments) {
                         $assignedRoles = $assignments.Group.RoleDefinitionName
                         # Identify missing roles
                         $missingRoles = $requiredRoles | Where-Object { $_ -notin $assignedRoles }
                     } else {
                         $missingRoles = $requiredRoles
                     }

                     # Add to data table
                     $row = $dataTable.NewRow()
                     $row["TenantId"] = $tenant.TenantId
                     $row["SubscriptionId"] = $subscription.Id
                     $row["Analyst"] = $analyst
                     $row["AssignedRoles"] = ($assignedRoles -join ", ")
                     $row["MissingRoles"] = ($missingRoles -join ", ")
                     $dataTable.Rows.Add($row)

                     # Update missing permissions count
                     if ($missingPermissionsCount.ContainsKey($analyst)) {
                         $missingPermissionsCount[$analyst] += $missingRoles.Count
                     } else {
                         $missingPermissionsCount[$analyst] = $missingRoles.Count
                     }
                 }
             }
         } catch {
             [System.Windows.Forms.MessageBox]::Show("Error processing tenant '$($tenant.TenantId)': $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
             continue
         }
     }

     # Bind data table to DataGridView
     $dataGridView.DataSource = $dataTable

     # Update the chart
     foreach ($entry in $missingPermissionsCount.GetEnumerator()) {
         $point = $chart.Series["MissingPermissions"].Points.Add($entry.Value)
         $point.AxisLabel = $entry.Key
         $point.LegendText = $entry.Key
     }

     $statusLabel.Text = "Analysis complete."
 }

 # Create the Form
 $form = New-Object System.Windows.Forms.Form
 $form.Text = "Permission Gap Analysis"
 $form.Size = New-Object System.Drawing.Size(800, 700)
 $form.StartPosition = "CenterScreen"

 # Add Select File Button
 $selectFileButton = New-Object System.Windows.Forms.Button
 $selectFileButton.Text = "Select Analysts File"
 $selectFileButton.Location = New-Object System.Drawing.Point(10, 10)
 $selectFileButton.Size = New-Object System.Drawing.Size(140, 30)
 $form.Controls.Add($selectFileButton)

 # Add Label to display selected file path
 $filePathLabel = New-Object System.Windows.Forms.Label
 $filePathLabel.Location = New-Object System.Drawing.Point(160, 17)
 $filePathLabel.Size = New-Object System.Drawing.Size(620, 20)
 $form.Controls.Add($filePathLabel)

 # Add Start Button
 $startButton = New-Object System.Windows.Forms.Button
 $startButton.Text = "Start Analysis"
 $startButton.Location = New-Object System.Drawing.Point(10, 50)
 $startButton.Size = New-Object System.Drawing.Size(100, 30)
 $form.Controls.Add($startButton)

 # Add Exit Button
 $exitButton = New-Object System.Windows.Forms.Button
 $exitButton.Text = "Exit"
 $exitButton.Location = New-Object System.Drawing.Point(120, 50)
 $exitButton.Size = New-Object System.Drawing.Size(100, 30)
 $form.Controls.Add($exitButton)

 # Add Status Label
 $statusLabel = New-Object System.Windows.Forms.Label
 $statusLabel.Location = New-Object System.Drawing.Point(230, 58)
 $statusLabel.Size = New-Object System.Drawing.Size(550, 20)
 $form.Controls.Add($statusLabel)

 # Add DataGridView
 $dataGridView = New-Object System.Windows.Forms.DataGridView
 $dataGridView.Location = New-Object System.Drawing.Point(10, 90)
 $dataGridView.Size = New-Object System.Drawing.Size(760, 480)
 $form.Controls.Add($dataGridView)

 # Add Chart
 $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
 $chart.Location = New-Object System.Drawing.Point(10, 580)
 $chart.Size = New-Object System.Drawing.Size(760, 100)
 $form.Controls.Add($chart)

 # Set up the chart area
 $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
 $chart.ChartAreas.Add($chartArea)

 # Set up the legend
 $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
 $chart.Legends.Add($legend)

 # Set up the series
 $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
 $series.Name = "MissingPermissions"
 $series.ChartType = "Column"
 $chart.Series.Add($series)

 # Event handler for Select File Button
 $selectFileButton.Add_Click({
     $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
     $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
     $openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
     $openFileDialog.FilterIndex = 1
     $openFileDialog.RestoreDirectory = $true

     if ($openFileDialog.ShowDialog() -eq "OK") {
         $Global:AnalystsFilePath = $openFileDialog.FileName
         $filePathLabel.Text = "Selected File: $Global:AnalystsFilePath"
     }
 })

 # Event handler for Start Button
 $startButton.Add_Click({
     # Disable the buttons to prevent multiple clicks
     $startButton.Enabled = $false
     $selectFileButton.Enabled = $false

     # Perform the analysis
     Perform-GapAnalysis -dataGridView $dataGridView -chart $chart -statusLabel $statusLabel

     # Re-enable the buttons
     $startButton.Enabled = $true
     $selectFileButton.Enabled = $true
 })

 # Event handler for Exit Button
 $exitButton.Add_Click({
     $form.Close()
 })

 # Show the Form
 [void]$form.ShowDialog()
