# --- Require Admin ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Config ---
$TemplateFolder = Join-Path $env:USERPROFILE "ServiceTemplates"
if (-not (Test-Path $TemplateFolder)) { New-Item -ItemType Directory -Path $TemplateFolder | Out-Null }

# =====================================================
# SCREEN-AWARE SIZING
# =====================================================
$screenW = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenH = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
$appW = [Math]::Max([int]($screenW * 0.5875), 1100)
$appH = [Math]::Max([int]($screenH * 0.6912), 600)

$rightPanelW = [Math]::Max([Math]::Min([int]($appW * 0.129), 320), 260)
$margins = 54; $scrollbar = 20
$dgvW = $appW - $rightPanelW - $margins - $scrollbar
$checkboxW = [Math]::Max([Math]::Min([int]($dgvW * 0.02), 40), 28)
$dataW = $dgvW - $checkboxW

$colPcts = @{
    "Type"=0.0247;"Group"=0.0495;"DisplayName"=0.1258;"Status"=0.0330;"StartupType"=0.0371
    "ErrorControl"=0.0330;"SvcName"=0.0495;"Dependencies"=0.0412;"FileDescription"=0.1340
    "Company"=0.0536;"ProductName"=0.0557;"Description"=0.2186;"Filename"=0.1031;"LastWriteTime"=0.0412
}
$colWidths = @{}
foreach ($key in $colPcts.Keys) { $colWidths[$key] = [Math]::Max([int]($dataW * $colPcts[$key]), 30) }

$rowH = [Math]::Max([int]($screenH * 0.0139), 22)
$headerH = [Math]::Max([int]($rowH * 1.30), 30)

$script:RegBase = "HKLM:\SYSTEM\CurrentControlSet\Services"

# File version info cache - avoids re-reading same exe/sys files
$script:FileInfoCache = @{}

# --- Helper Functions ---
# Registry helpers are now inlined in Get-AllItems using batch-loaded data
function Get-FileDetails {
    param([string]$RawPath)
    $result = @{ FileDescription=""; Company=""; ProductName=""; LastWriteTime="" }
    if (-not $RawPath) { return $result }

    # Check cache first
    if ($script:FileInfoCache.ContainsKey($RawPath)) { return $script:FileInfoCache[$RawPath] }

    $filePath = $RawPath.Trim()
    if ($filePath.StartsWith('"')) {
        if ($filePath -match '^"([^"]+)"') { $filePath = $Matches[1] }
    } elseif ($filePath -match '^(\S+\.(exe|sys|dll))') { $filePath = $Matches[1] }
    $filePath = $filePath -replace '^\\\?\?\\', ''
    if ($filePath -and (Test-Path $filePath -ErrorAction SilentlyContinue)) {
        try {
            $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($filePath)
            $result.FileDescription = if ($vi.FileDescription) { $vi.FileDescription } else { "" }
            $result.Company = if ($vi.CompanyName) { $vi.CompanyName } else { "" }
            $result.ProductName = if ($vi.ProductName) { $vi.ProductName } else { "" }
        } catch {}
        try {
            $fi = Get-Item $filePath -ErrorAction SilentlyContinue
            if ($fi) { $result.LastWriteTime = $fi.LastWriteTime.ToString("M/d/yyyy H:mm") }
        } catch {}
    }

    $script:FileInfoCache[$RawPath] = $result
    return $result
}

function Get-AllItems {
    $items = [System.Collections.ArrayList]::new()

    # Pre-load all registry service data in one batch (much faster than per-service reads)
    $regData = @{}
    try {
        Get-ChildItem -Path $script:RegBase -ErrorAction SilentlyContinue | ForEach-Object {
            $svcName = $_.PSChildName
            $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            $regData[$svcName] = $props
        }
    } catch {}

    # Fast registry lookups using pre-loaded data
    function Get-RegGroup([string]$n) { $p = $regData[$n]; if ($p -and $p.Group) { return $p.Group }; return "" }
    function Get-RegDeps([string]$n) { $p = $regData[$n]; if ($p -and $p.DependOnService) { return ($p.DependOnService -join ",") }; return "" }
    function Get-RegErrorCtrl([string]$n) {
        $p = $regData[$n]; if (-not $p) { return "" }
        $v = $p.ErrorControl; $r = switch ($v) { 0 {"Ignore"} 1 {"Normal"} 2 {"Severe"} 3 {"Critical"} default {""} }; return $r
    }

    $driverHash = @{}

    Get-CimInstance Win32_SystemDriver -ErrorAction SilentlyContinue | ForEach-Object {
        $sm = switch ($_.StartMode) { "Boot"{"Boot"} "System"{"System"} "Auto"{"Automatic"} "Manual"{"Manual"} "Disabled"{"Disabled"} default{$_.StartMode} }
        $st = if ($_.State -eq "Running") { "Started" } else { "Stopped" }
        $fd = Get-FileDetails $_.PathName
        $driverHash[$_.Name] = $true
        $items.Add([PSCustomObject]@{
            Type="Driver"; Group=(Get-RegGroup $_.Name); DisplayName=$_.DisplayName; Status=$st
            StartupType=$sm; ErrorControl=(Get-RegErrorCtrl $_.Name); Name=$_.Name
            Dependencies=(Get-RegDeps $_.Name); FileDescription=$fd.FileDescription
            Company=$fd.Company; ProductName=$fd.ProductName
            Description=$(if ($_.Description) { $_.Description } else { "" })
            Filename=$(if ($_.PathName) { $_.PathName } else { "" }); LastWriteTime=$fd.LastWriteTime
        }) | Out-Null
    }

    Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | ForEach-Object {
        if (-not $driverHash.ContainsKey($_.Name)) {
            $sm = switch ($_.StartMode) { "Boot"{"Boot"} "System"{"System"} "Auto"{"Automatic"} "Manual"{"Manual"} "Disabled"{"Disabled"} default{$_.StartMode} }
            $st = if ($_.State -eq "Running") { "Started" } else { "Stopped" }
            $fd = Get-FileDetails $_.PathName
            $items.Add([PSCustomObject]@{
                Type="Service"; Group=(Get-RegGroup $_.Name); DisplayName=$_.DisplayName; Status=$st
                StartupType=$sm; ErrorControl=(Get-RegErrorCtrl $_.Name); Name=$_.Name
                Dependencies=(Get-RegDeps $_.Name); FileDescription=$fd.FileDescription
                Company=$fd.Company; ProductName=$fd.ProductName
                Description=$(if ($_.Description) { $_.Description } else { "" })
                Filename=$(if ($_.PathName) { $_.PathName } else { "" }); LastWriteTime=$fd.LastWriteTime
            }) | Out-Null
        }
    }
    return $items.ToArray()
}

function Save-Template {
    param([string]$TemplateName, [array]$SelectedItems)
    $path = Join-Path $TemplateFolder "$TemplateName.json"
    @{ Name=$TemplateName; CreatedDate=(Get-Date -Format "yyyy-MM-dd HH:mm:ss"); Items=$SelectedItems } |
        ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
}
function Get-TemplateList {
    if (Test-Path $TemplateFolder) { Get-ChildItem -Path $TemplateFolder -Filter "*.json" | ForEach-Object { $_.BaseName } }
}
function Load-Template {
    param([string]$TemplateName)
    $path = Join-Path $TemplateFolder "$TemplateName.json"
    if (Test-Path $path) { return (Get-Content -Path $path -Raw | ConvertFrom-Json) }
    return $null
}
function Delete-Template {
    param([string]$TemplateName)
    $path = Join-Path $TemplateFolder "$TemplateName.json"
    if (Test-Path $path) { Remove-Item -Path $path -Force }
}

# --- Colors ---
$BgDark=[System.Drawing.Color]::FromArgb(30,30,30)
$BgPanel=[System.Drawing.Color]::FromArgb(42,42,42)
$BgInput=[System.Drawing.Color]::FromArgb(55,55,55)
$BgSortBar=[System.Drawing.Color]::FromArgb(36,36,48)
$BgDropdown=[System.Drawing.Color]::FromArgb(45,45,52)
$Accent=[System.Drawing.Color]::FromArgb(0,122,204)
$Danger=[System.Drawing.Color]::FromArgb(200,50,50)
$Success=[System.Drawing.Color]::FromArgb(50,170,80)
$FgWhite=[System.Drawing.Color]::White
$FgDim=[System.Drawing.Color]::FromArgb(160,160,160)
$FgGreen=[System.Drawing.Color]::FromArgb(80,200,120)
$FgRed=[System.Drawing.Color]::FromArgb(240,80,80)
$FgYellow=[System.Drawing.Color]::FromArgb(240,200,80)
$FgCyan=[System.Drawing.Color]::FromArgb(80,200,220)
$FgOrange=[System.Drawing.Color]::FromArgb(240,160,60)
$GridBg=[System.Drawing.Color]::FromArgb(35,35,35)
$GridAltRow=[System.Drawing.Color]::FromArgb(40,40,40)
$GridSelBg=[System.Drawing.Color]::FromArgb(0,90,158)

$FontMain  = New-Object System.Drawing.Font("Segoe UI", 9)
$FontBold  = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$FontTitle = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$FontSmall = New-Object System.Drawing.Font("Segoe UI", 8)
$FontSortTag = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$FontDesc  = New-Object System.Drawing.Font("Segoe UI", 7)

# --- Sort + Filter State ---
$script:SortLevels = [System.Collections.ArrayList]::new()
$script:MaxSortLevels = 4
$script:ColumnFilters = @{}

# One-time warning flag for driver/critical changes
$script:DriverWarningShown = $false

# Flag to suppress CellValueChanged during grid population
$script:_isPopulating = $false

# Header checkbox state
$script:_headerCheckState = $false

# Active dropdown reference (fixes closure bug)
$script:_activeDropdown = $null

$script:DgvColToProperty = @{
    "Type"="Type";"Group"="Group";"DisplayName"="DisplayName";"Status"="Status"
    "StartupType"="StartupType";"ErrorControl"="ErrorControl";"SvcName"="Name"
    "Dependencies"="Dependencies";"FileDescription"="FileDescription";"Company"="Company"
    "ProductName"="ProductName";"Description"="Description";"Filename"="Filename";"LastWriteTime"="LastWriteTime"
}
$script:PropertyToLabel = @{
    "Type"="Type";"Group"="Group";"DisplayName"="Display Name";"Status"="Status"
    "StartupType"="Startup Type";"ErrorControl"="ErrorControl";"Name"="Name"
    "Dependencies"="Dependencies";"FileDescription"="File Description";"Company"="Company"
    "ProductName"="Product Name";"Description"="Description";"Filename"="Filename";"LastWriteTime"="Last Write Time"
}
# Short labels for narrow column headers
$script:PropertyToShortLabel = @{
    "Type"="Type";"Group"="Group";"DisplayName"="Display Name";"Status"="Status"
    "StartupType"="Start Type";"ErrorControl"="Error";"Name"="Name"
    "Dependencies"="Dependencies";"FileDescription"="File Desc";"Company"="Company"
    "ProductName"="Product";"Description"="Description";"Filename"="Filename";"LastWriteTime"="Last Write"
}

# Default sort
$script:SortLevels.Add(@{Column="Name";Direction="ASC"}) | Out-Null
$script:SortLevels.Add(@{Column="StartupType";Direction="DESC"}) | Out-Null
$script:SortLevels.Add(@{Column="Status";Direction="ASC"}) | Out-Null
$script:SortLevels.Add(@{Column="Type";Direction="DESC"}) | Out-Null

# =====================================================
# LAYOUT
# =====================================================
$pad=[int]($appW*0.009); $gap=[int]($appW*0.006)
$filterY=[int]($appH*0.049); $filterH=[int]($appH*0.027)
$sortBarY=$filterY+$filterH+2; $sortBarH=[int]($appH*0.022)
$dgvY=$sortBarY+$sortBarH+4; $bottomBarH=[int]($appH*0.085)
$dgvH=$appH-$dgvY-$bottomBarH-40
$leftAreaW=$appW-$rightPanelW-$pad-$gap-$pad; $rightX=$pad+$leftAreaW+$gap

# --- Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Service Template Manager"
$form.Size = New-Object System.Drawing.Size($appW, $appH)
$form.StartPosition = "CenterScreen"
$form.BackColor = $BgDark; $form.ForeColor = $FgWhite; $form.Font = $FontMain
$form.FormBorderStyle = "Sizable"; $form.MinimumSize = New-Object System.Drawing.Size(1100, 600)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Service Template Manager"; $lblTitle.Font = $FontTitle
$lblTitle.ForeColor = $FgWhite; $lblTitle.Location = New-Object System.Drawing.Point($pad, 12); $lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Save and restore snapshots of disabled drivers and services"
$lblSubtitle.Font = $FontSmall; $lblSubtitle.ForeColor = $FgDim
$lblSubtitle.Location = New-Object System.Drawing.Point(($pad+2), 42); $lblSubtitle.AutoSize = $true
$form.Controls.Add($lblSubtitle)

# --- Filter Bar ---
$pnlFilter = New-Object System.Windows.Forms.Panel
$pnlFilter.Location = New-Object System.Drawing.Point($pad, $filterY)
$pnlFilter.Size = New-Object System.Drawing.Size($leftAreaW, $filterH)
$pnlFilter.BackColor = $BgPanel; $pnlFilter.Anchor = "Top,Left,Right"
$form.Controls.Add($pnlFilter)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text="Search:"; $lblSearch.ForeColor=$FgDim; $lblSearch.Location=New-Object System.Drawing.Point(8,9); $lblSearch.AutoSize=$true
$pnlFilter.Controls.Add($lblSearch)
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location=New-Object System.Drawing.Point(60,6); $txtSearch.Size=New-Object System.Drawing.Size(200,24)
$txtSearch.BackColor=$BgInput; $txtSearch.ForeColor=$FgWhite; $txtSearch.BorderStyle="FixedSingle"; $txtSearch.Font=$FontMain
$pnlFilter.Controls.Add($txtSearch)

$lblFilterType = New-Object System.Windows.Forms.Label
$lblFilterType.Text="Type:"; $lblFilterType.ForeColor=$FgDim; $lblFilterType.Location=New-Object System.Drawing.Point(272,9); $lblFilterType.AutoSize=$true
$pnlFilter.Controls.Add($lblFilterType)
$cmbType = New-Object System.Windows.Forms.ComboBox
$cmbType.Location=New-Object System.Drawing.Point(306,6); $cmbType.Size=New-Object System.Drawing.Size(110,24)
$cmbType.DropDownStyle="DropDownList"; $cmbType.BackColor=$BgInput; $cmbType.ForeColor=$FgWhite; $cmbType.FlatStyle="Flat"
$cmbType.Items.AddRange(@("All","Driver","Service")); $cmbType.SelectedIndex=0
$pnlFilter.Controls.Add($cmbType)

$lblFilterStatus = New-Object System.Windows.Forms.Label
$lblFilterStatus.Text="Status:"; $lblFilterStatus.ForeColor=$FgDim; $lblFilterStatus.Location=New-Object System.Drawing.Point(426,9); $lblFilterStatus.AutoSize=$true
$pnlFilter.Controls.Add($lblFilterStatus)
$cmbStatus = New-Object System.Windows.Forms.ComboBox
$cmbStatus.Location=New-Object System.Drawing.Point(472,6); $cmbStatus.Size=New-Object System.Drawing.Size(110,24)
$cmbStatus.DropDownStyle="DropDownList"; $cmbStatus.BackColor=$BgInput; $cmbStatus.ForeColor=$FgWhite; $cmbStatus.FlatStyle="Flat"
$cmbStatus.Items.AddRange(@("All","Started","Stopped")); $cmbStatus.SelectedIndex=0
$pnlFilter.Controls.Add($cmbStatus)

$lblFilterStart = New-Object System.Windows.Forms.Label
$lblFilterStart.Text="Startup:"; $lblFilterStart.ForeColor=$FgDim; $lblFilterStart.Location=New-Object System.Drawing.Point(594,9); $lblFilterStart.AutoSize=$true
$pnlFilter.Controls.Add($lblFilterStart)
$cmbStartType = New-Object System.Windows.Forms.ComboBox
$cmbStartType.Location=New-Object System.Drawing.Point(648,6); $cmbStartType.Size=New-Object System.Drawing.Size(120,24)
$cmbStartType.DropDownStyle="DropDownList"; $cmbStartType.BackColor=$BgInput; $cmbStartType.ForeColor=$FgWhite; $cmbStartType.FlatStyle="Flat"
$cmbStartType.Items.AddRange(@("All","Boot","System","Automatic","Manual","Disabled")); $cmbStartType.SelectedIndex=0
$pnlFilter.Controls.Add($cmbStartType)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text="Refresh"; $btnRefresh.Location=New-Object System.Drawing.Point(780,5); $btnRefresh.Size=New-Object System.Drawing.Size(80,26)
$btnRefresh.FlatStyle="Flat"; $btnRefresh.BackColor=$BgInput; $btnRefresh.ForeColor=$FgWhite
$btnRefresh.FlatAppearance.BorderColor=$FgDim; $btnRefresh.Cursor="Hand"
$pnlFilter.Controls.Add($btnRefresh)

$chkAutoCheck = New-Object System.Windows.Forms.CheckBox
$chkAutoCheck.Text="Auto-check disabled"; $chkAutoCheck.ForeColor=$FgDim
$chkAutoCheck.Location=New-Object System.Drawing.Point(870,9); $chkAutoCheck.AutoSize=$true; $chkAutoCheck.Checked=$true
$pnlFilter.Controls.Add($chkAutoCheck)

# --- Sort/Filter Indicator Bar ---
$pnlSortBar = New-Object System.Windows.Forms.Panel
$pnlSortBar.Location=New-Object System.Drawing.Point($pad,$sortBarY)
$pnlSortBar.Size=New-Object System.Drawing.Size($leftAreaW,$sortBarH)
$pnlSortBar.BackColor=$BgSortBar; $pnlSortBar.Anchor="Top,Left,Right"
$form.Controls.Add($pnlSortBar)

$lblSortDisplay = New-Object System.Windows.Forms.Label
$lblSortDisplay.Location=New-Object System.Drawing.Point(8,4)
$lblSortDisplay.Size=New-Object System.Drawing.Size(($leftAreaW-360),18)
$lblSortDisplay.ForeColor=$FgOrange; $lblSortDisplay.Font=$FontSortTag; $lblSortDisplay.Anchor="Top,Left,Right"
$pnlSortBar.Controls.Add($lblSortDisplay)

$lblFilterIndicator = New-Object System.Windows.Forms.Label
$lblFilterIndicator.Location=New-Object System.Drawing.Point(($leftAreaW-540),4)
$lblFilterIndicator.Size=New-Object System.Drawing.Size(200,18)
$lblFilterIndicator.ForeColor=$FgCyan; $lblFilterIndicator.Font=$FontSortTag
$lblFilterIndicator.TextAlign="MiddleRight"; $lblFilterIndicator.Anchor="Top,Right"
$pnlSortBar.Controls.Add($lblFilterIndicator)

$btnClearSort = New-Object System.Windows.Forms.Button
$btnClearSort.Text="Clear Sort"; $btnClearSort.Location=New-Object System.Drawing.Point(($leftAreaW-330),2)
$btnClearSort.Size=New-Object System.Drawing.Size(78,($sortBarH-4)); $btnClearSort.FlatStyle="Flat"
$btnClearSort.BackColor=[System.Drawing.Color]::FromArgb(120,60,20); $btnClearSort.ForeColor=$FgOrange
$btnClearSort.Font=$FontSortTag; $btnClearSort.FlatAppearance.BorderColor=$FgOrange
$btnClearSort.FlatAppearance.BorderSize=1; $btnClearSort.Cursor="Hand"; $btnClearSort.Anchor="Top,Right"
$pnlSortBar.Controls.Add($btnClearSort)

$btnClearFilters = New-Object System.Windows.Forms.Button
$btnClearFilters.Text="Clear Filters"; $btnClearFilters.Location=New-Object System.Drawing.Point(($leftAreaW-246),2)
$btnClearFilters.Size=New-Object System.Drawing.Size(82,($sortBarH-4)); $btnClearFilters.FlatStyle="Flat"
$btnClearFilters.BackColor=[System.Drawing.Color]::FromArgb(20,80,120); $btnClearFilters.ForeColor=$FgCyan
$btnClearFilters.Font=$FontSortTag; $btnClearFilters.FlatAppearance.BorderColor=$FgCyan
$btnClearFilters.FlatAppearance.BorderSize=1; $btnClearFilters.Cursor="Hand"; $btnClearFilters.Anchor="Top,Right"
$pnlSortBar.Controls.Add($btnClearFilters)

$btnClearAll = New-Object System.Windows.Forms.Button
$btnClearAll.Text="Clear All"; $btnClearAll.Location=New-Object System.Drawing.Point(($leftAreaW-158),2)
$btnClearAll.Size=New-Object System.Drawing.Size(68,($sortBarH-4)); $btnClearAll.FlatStyle="Flat"
$btnClearAll.BackColor=$Danger; $btnClearAll.ForeColor=$FgWhite
$btnClearAll.Font=$FontSortTag; $btnClearAll.FlatAppearance.BorderSize=0; $btnClearAll.Cursor="Hand"; $btnClearAll.Anchor="Top,Right"
$pnlSortBar.Controls.Add($btnClearAll)

$btnResetDefault = New-Object System.Windows.Forms.Button
$btnResetDefault.Text="Default Sort"; $btnResetDefault.Location=New-Object System.Drawing.Point(($leftAreaW-84),2)
$btnResetDefault.Size=New-Object System.Drawing.Size(80,($sortBarH-4)); $btnResetDefault.FlatStyle="Flat"
$btnResetDefault.BackColor=$BgInput; $btnResetDefault.ForeColor=$FgDim
$btnResetDefault.Font=$FontSortTag; $btnResetDefault.FlatAppearance.BorderColor=$FgDim
$btnResetDefault.FlatAppearance.BorderSize=1; $btnResetDefault.Cursor="Hand"; $btnResetDefault.Anchor="Top,Right"
$pnlSortBar.Controls.Add($btnResetDefault)

# --- DataGridView ---
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location=New-Object System.Drawing.Point($pad,$dgvY)
$dgv.Size=New-Object System.Drawing.Size($leftAreaW,$dgvH); $dgv.Anchor="Top,Left,Right,Bottom"
$dgv.BackgroundColor=$GridBg; $dgv.GridColor=[System.Drawing.Color]::FromArgb(55,55,55)
$dgv.DefaultCellStyle.BackColor=$GridBg; $dgv.DefaultCellStyle.ForeColor=$FgWhite
$dgv.DefaultCellStyle.SelectionBackColor=$GridSelBg; $dgv.DefaultCellStyle.SelectionForeColor=$FgWhite
$dgv.DefaultCellStyle.Font=$FontMain
$dgv.AlternatingRowsDefaultCellStyle.BackColor=$GridAltRow
$dgv.ColumnHeadersDefaultCellStyle.BackColor=$BgPanel
$dgv.ColumnHeadersDefaultCellStyle.ForeColor=$FgWhite
$dgv.ColumnHeadersDefaultCellStyle.Font=$FontBold
$dgv.ColumnHeadersDefaultCellStyle.Alignment="MiddleLeft"
$dgv.EnableHeadersVisualStyles=$false
$dgv.ColumnHeadersHeight=$headerH
$dgv.ColumnHeadersHeightSizeMode="DisableResizing"
$dgv.RowHeadersVisible=$false; $dgv.AllowUserToAddRows=$false
$dgv.AllowUserToDeleteRows=$false; $dgv.AllowUserToResizeRows=$false
$dgv.SelectionMode="FullRowSelect"; $dgv.MultiSelect=$false
$dgv.ReadOnly=$false; $dgv.BorderStyle="None"
$dgv.CellBorderStyle="SingleHorizontal"; $dgv.RowTemplate.Height=$rowH; $dgv.ScrollBars="Both"

# Columns - HeaderText is just the label (arrow drawn by CellPainting)
$colDefs = @(
    @("Selected","",                   $checkboxW, $false),
    @("Type",    "Type",               $colWidths["Type"], $true),
    @("Group",   "Group",              $colWidths["Group"], $true),
    @("DisplayName","Display Name",    $colWidths["DisplayName"], $true),
    @("Status",  "Status",             $colWidths["Status"], $true),
    @("StartupType","Start Type",      $colWidths["StartupType"], $true),
    @("ErrorControl","Error",          $colWidths["ErrorControl"], $true),
    @("SvcName", "Name",               $colWidths["SvcName"], $true),
    @("FileDescription","File Desc",   $colWidths["FileDescription"], $true),
    @("Description","Description",     $colWidths["Description"], $true),
    @("Company", "Company",            $colWidths["Company"], $true),
    @("ProductName","Product",         $colWidths["ProductName"], $true),
    @("Dependencies","Dependencies",   $colWidths["Dependencies"], $true),
    @("LastWriteTime","Last Write",    $colWidths["LastWriteTime"], $true),
    @("Filename","Filename",           $colWidths["Filename"], $true)
)

foreach ($def in $colDefs) {
    if ($def[0] -eq "Selected") {
        $col = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $col.Name=$def[0]; $col.HeaderText=$def[1]; $col.Width=$def[2]
        $col.ReadOnly=$false; $col.SortMode="NotSortable"
    } elseif ($def[0] -eq "Status") {
        $col = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $col.Name=$def[0]; $col.HeaderText=$def[1]; $col.Width=$def[2]
        $col.Items.AddRange(@("Started","Stopped"))
        $col.FlatStyle="Flat"; $col.ReadOnly=$false; $col.SortMode="Programmatic"
        $col.DefaultCellStyle.BackColor=$GridBg; $col.DefaultCellStyle.ForeColor=$FgWhite
    } elseif ($def[0] -eq "StartupType") {
        $col = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $col.Name=$def[0]; $col.HeaderText=$def[1]; $col.Width=$def[2]
        $col.Items.AddRange(@("Boot","System","Automatic","Manual","Disabled"))
        $col.FlatStyle="Flat"; $col.ReadOnly=$false; $col.SortMode="Programmatic"
        $col.DefaultCellStyle.BackColor=$GridBg; $col.DefaultCellStyle.ForeColor=$FgWhite
    } else {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name=$def[0]; $col.HeaderText=$def[1]; $col.Width=$def[2]; $col.ReadOnly=$true
        $col.SortMode="Programmatic"
        if ($def[0] -eq "Description") { $col.DefaultCellStyle.Font = $FontDesc }
    }
    $dgv.Columns.Add($col) | Out-Null
}

# =====================================================
# CUSTOM HEADER PAINTING (header checkbox + right-aligned dropdown arrow)
# =====================================================
$FontArrowSmall = New-Object System.Drawing.Font("Segoe UI", 6)

# --- Pre-allocated GDI objects for CellPainting (avoids per-call alloc/dispose) ---
$script:PaintBrushBg      = New-Object System.Drawing.SolidBrush($BgPanel)
$script:PaintPenBorder    = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60,60,60), 1)
$script:PaintPenCbBorder  = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(140,140,140), 1)
$script:PaintBrushCbOff   = New-Object System.Drawing.SolidBrush($BgInput)
$script:PaintBrushCbOn    = New-Object System.Drawing.SolidBrush($Accent)
$script:PaintPenCheck     = New-Object System.Drawing.Pen($FgWhite, 2)
$script:PaintBrushWhite   = New-Object System.Drawing.SolidBrush($FgWhite)
$script:PaintBrushCyan    = New-Object System.Drawing.SolidBrush($FgCyan)
$script:PaintBrushArrow   = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(120,120,120))
$script:PaintSFText       = New-Object System.Drawing.StringFormat
$script:PaintSFText.Alignment      = "Near"
$script:PaintSFText.LineAlignment  = "Center"
$script:PaintSFText.Trimming       = "EllipsisCharacter"
$script:PaintSFText.FormatFlags    = "NoWrap"
$script:PaintSFArrow      = New-Object System.Drawing.StringFormat
$script:PaintSFArrow.Alignment     = "Center"
$script:PaintSFArrow.LineAlignment = "Center"
$script:ArrowChar = [string][char]0x25BC

# --- Enable double buffering on DGV via reflection (eliminates flicker) ---
$dgvType = $dgv.GetType()
$dbProp = $dgvType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
$dbProp.SetValue($dgv, $true, $null)

$dgv.Add_CellPainting({
    param($sender, $e)
    if ($e.RowIndex -ne -1) { return }
    if ($e.ColumnIndex -lt 0) { return }
    $colName = $sender.Columns[$e.ColumnIndex].Name

    $e.Handled = $true
    $g = $e.Graphics
    $b = $e.CellBounds

    # Fill background + bottom border (shared by all headers)
    $g.FillRectangle($script:PaintBrushBg, $b)
    $g.DrawLine($script:PaintPenBorder, $b.Left, $b.Bottom - 1, $b.Right - 1, $b.Bottom - 1)

    if ($colName -eq "Selected") {
        # Draw a checkbox in the header
        $cbSize = 14
        $cbX = $b.X + [int](($b.Width - $cbSize) / 2)
        $cbY = $b.Y + [int](($b.Height - $cbSize) / 2)
        $cbRect = New-Object System.Drawing.Rectangle($cbX, $cbY, $cbSize, $cbSize)
        $g.DrawRectangle($script:PaintPenCbBorder, $cbRect)

        $fillBrush = if ($script:_headerCheckState) { $script:PaintBrushCbOn } else { $script:PaintBrushCbOff }
        $g.FillRectangle($fillBrush, ($cbX+1), ($cbY+1), ($cbSize-2), ($cbSize-2))

        if ($script:_headerCheckState) {
            $g.DrawLine($script:PaintPenCheck, ($cbX+3), ($cbY+7), ($cbX+6), ($cbY+10))
            $g.DrawLine($script:PaintPenCheck, ($cbX+6), ($cbY+10), ($cbX+11), ($cbY+4))
        }
        return
    }

    # Data columns: header text + small dropdown arrow
    $prop = $script:DgvColToProperty[$colName]
    $headerText = $sender.Columns[$e.ColumnIndex].HeaderText
    $hasFilter = $false
    if ($prop -and $script:ColumnFilters.ContainsKey($prop)) { $hasFilter = $true }
    if ($hasFilter) { $headerText = "$headerText [F]" }

    $arrowWidth = 12
    $textRect = New-Object System.Drawing.RectangleF(($b.X + 4), $b.Y, ($b.Width - $arrowWidth - 6), $b.Height)
    $textBrush = if ($hasFilter) { $script:PaintBrushCyan } else { $script:PaintBrushWhite }
    $g.DrawString($headerText, $FontBold, $textBrush, $textRect, $script:PaintSFText)

    $arrowRect = New-Object System.Drawing.RectangleF(($b.Right - $arrowWidth - 2), $b.Y, $arrowWidth, $b.Height)
    $g.DrawString($script:ArrowChar, $FontArrowSmall, $script:PaintBrushArrow, $arrowRect, $script:PaintSFArrow)
})

$form.Controls.Add($dgv)

# =====================================================
# RIGHT-CLICK CONTEXT MENU (Stop, Start, Open Reg, Open File)
# =====================================================
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ctxMenu.BackColor = $BgPanel; $ctxMenu.ForeColor = $FgWhite
$ctxMenu.Renderer = New-Object System.Windows.Forms.ToolStripProfessionalRenderer

$mnuStart = $ctxMenu.Items.Add("Start Service")
$mnuStop  = $ctxMenu.Items.Add("Stop Service")
$ctxMenu.Items.Add("-")  # separator
$mnuRegKey  = $ctxMenu.Items.Add("Open Registry Key Location")
$mnuFileLoc = $ctxMenu.Items.Add("Open File Location")

$dgv.ContextMenuStrip = $ctxMenu

# Track which row was right-clicked
$script:_ctxRowIndex = -1

$dgv.Add_CellMouseClick({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
        $script:_ctxRowIndex = $e.RowIndex
        $dgv.ClearSelection()
        $dgv.Rows[$e.RowIndex].Selected = $true
    }
})

$mnuStart.Add_Click({
    if ($script:_ctxRowIndex -lt 0 -or $script:_ctxRowIndex -ge $dgv.Rows.Count) { return }
    $row = $dgv.Rows[$script:_ctxRowIndex]
    $svcName = $row.Cells["SvcName"].Value
    try {
        Start-Service -Name $svcName -ErrorAction Stop
        $row.Cells["Status"].Value = "Started"
        $lblStatus.Text = "Started '$svcName'."; $lblStatus.ForeColor = $FgGreen
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to start '$svcName'.`n`n$($_.Exception.Message)", "Error", "OK", "Error")
    }
})

$mnuStop.Add_Click({
    if ($script:_ctxRowIndex -lt 0 -or $script:_ctxRowIndex -ge $dgv.Rows.Count) { return }
    $row = $dgv.Rows[$script:_ctxRowIndex]
    $svcName = $row.Cells["SvcName"].Value
    try {
        Stop-Service -Name $svcName -Force -ErrorAction Stop
        $row.Cells["Status"].Value = "Stopped"
        $lblStatus.Text = "Stopped '$svcName'."; $lblStatus.ForeColor = $FgYellow
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to stop '$svcName'.`n`n$($_.Exception.Message)", "Error", "OK", "Error")
    }
})

$mnuRegKey.Add_Click({
    if ($script:_ctxRowIndex -lt 0 -or $script:_ctxRowIndex -ge $dgv.Rows.Count) { return }
    $row = $dgv.Rows[$script:_ctxRowIndex]
    $svcName = $row.Cells["SvcName"].Value
    $regPath = "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\$svcName"
    # Set last-visited key so regedit opens to the right location
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit" -Name "LastKey" -Value $regPath -ErrorAction SilentlyContinue
    Start-Process regedit.exe
})

$mnuFileLoc.Add_Click({
    if ($script:_ctxRowIndex -lt 0 -or $script:_ctxRowIndex -ge $dgv.Rows.Count) { return }
    $row = $dgv.Rows[$script:_ctxRowIndex]
    $rawPath = $row.Cells["Filename"].Value
    if (-not $rawPath) {
        [System.Windows.Forms.MessageBox]::Show("No file path available.", "Info", "OK", "Information"); return
    }
    $filePath = $rawPath.Trim()
    if ($filePath.StartsWith('"')) {
        if ($filePath -match '^"([^"]+)"') { $filePath = $Matches[1] }
    } elseif ($filePath -match '^(\S+\.(exe|sys|dll))') { $filePath = $Matches[1] }
    $filePath = $filePath -replace '^\\\?\?\\', ''
    if (Test-Path $filePath -ErrorAction SilentlyContinue) {
        $folder = Split-Path $filePath -Parent
        Start-Process explorer.exe -ArgumentList "/select,`"$filePath`""
    } else {
        [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Info", "OK", "Warning")
    }
})

# --- Status / Counts ---
$statusY = $dgvY+$dgvH+6
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location=New-Object System.Drawing.Point($pad,$statusY)
$lblStatus.Size=New-Object System.Drawing.Size(600,18); $lblStatus.ForeColor=$FgDim
$lblStatus.Font=$FontSmall; $lblStatus.Text="Ready"; $lblStatus.Anchor="Bottom,Left"
$form.Controls.Add($lblStatus)

$lblCounts = New-Object System.Windows.Forms.Label
$lblCounts.Location=New-Object System.Drawing.Point($pad,($statusY+20))
$lblCounts.Size=New-Object System.Drawing.Size(600,18); $lblCounts.ForeColor=$FgDim; $lblCounts.Font=$FontSmall; $lblCounts.Anchor="Bottom,Left"
$form.Controls.Add($lblCounts)

# --- Bottom Buttons ---
$btnY = $statusY+44
foreach ($bDef in @(
    @("btnSelectAll","Check All Visible",0,120),
    @("btnSelectNone","Uncheck All",128,120),
    @("btnExportCSV","Export Checked to CSV",256,160)
)) {
    $b = New-Object System.Windows.Forms.Button
    $b.Name=$bDef[0]; $b.Text=$bDef[1]; $b.Location=New-Object System.Drawing.Point(($pad+$bDef[2]),$btnY)
    $b.Size=New-Object System.Drawing.Size($bDef[3],28); $b.FlatStyle="Flat"; $b.BackColor=$BgInput; $b.ForeColor=$FgWhite
    $b.FlatAppearance.BorderColor=$FgDim; $b.Cursor="Hand"; $b.Anchor="Bottom,Left"
    $form.Controls.Add($b); Set-Variable -Name $bDef[0] -Value $b
}

# --- Right Panel ---
$rightPanelH = $appH-$filterY-40
$pnlRight = New-Object System.Windows.Forms.Panel
$pnlRight.Location=New-Object System.Drawing.Point($rightX,$filterY)
$pnlRight.Size=New-Object System.Drawing.Size($rightPanelW,$rightPanelH)
$pnlRight.BackColor=$BgPanel; $pnlRight.Anchor="Top,Right,Bottom"; $pnlRight.AutoScroll=$true
$form.Controls.Add($pnlRight)

$innerW = $rightPanelW-24

$lblTemplates = New-Object System.Windows.Forms.Label
$lblTemplates.Text="Saved Templates"; $lblTemplates.Font=$FontBold; $lblTemplates.ForeColor=$FgWhite
$lblTemplates.Location=New-Object System.Drawing.Point(12,10); $lblTemplates.AutoSize=$true
$pnlRight.Controls.Add($lblTemplates)

$lstTemplates = New-Object System.Windows.Forms.ListBox
$lstTemplates.Location=New-Object System.Drawing.Point(12,36); $lstTemplates.Size=New-Object System.Drawing.Size($innerW,180)
$lstTemplates.BackColor=$BgInput; $lstTemplates.ForeColor=$FgWhite; $lstTemplates.BorderStyle="FixedSingle"; $lstTemplates.Font=$FontMain
$pnlRight.Controls.Add($lstTemplates)

$lblTplInfo = New-Object System.Windows.Forms.Label
$lblTplInfo.Location=New-Object System.Drawing.Point(12,222); $lblTplInfo.Size=New-Object System.Drawing.Size($innerW,50)
$lblTplInfo.ForeColor=$FgDim; $lblTplInfo.Font=$FontSmall; $lblTplInfo.Text="Select a template to preview or apply."
$pnlRight.Controls.Add($lblTplInfo)

$lblSaveAs = New-Object System.Windows.Forms.Label
$lblSaveAs.Text="Save Checked As Template:"; $lblSaveAs.Font=$FontBold; $lblSaveAs.ForeColor=$FgWhite
$lblSaveAs.Location=New-Object System.Drawing.Point(12,280); $lblSaveAs.AutoSize=$true
$pnlRight.Controls.Add($lblSaveAs)

$txtTemplateName = New-Object System.Windows.Forms.TextBox
$txtTemplateName.Location=New-Object System.Drawing.Point(12,304); $txtTemplateName.Size=New-Object System.Drawing.Size($innerW,24)
$txtTemplateName.BackColor=$BgInput; $txtTemplateName.ForeColor=$FgWhite; $txtTemplateName.BorderStyle="FixedSingle"; $txtTemplateName.Font=$FontMain
$pnlRight.Controls.Add($txtTemplateName)

function New-StyledButton { param([string]$Text,[int]$X,[int]$Y,[int]$W,[int]$H,$Color)
    $b=New-Object System.Windows.Forms.Button; $b.Text=$Text; $b.Location=New-Object System.Drawing.Point($X,$Y)
    $b.Size=New-Object System.Drawing.Size($W,$H); $b.FlatStyle="Flat"; $b.BackColor=$Color; $b.ForeColor=$FgWhite
    $b.Font=$FontBold; $b.FlatAppearance.BorderSize=0; $b.Cursor="Hand"; return $b
}

$btnSave = New-StyledButton "Save Template" 12 338 $innerW 36 $Accent; $pnlRight.Controls.Add($btnSave)
$halfW=[int](($innerW-8)/2)
$btnLoad = New-StyledButton "Preview Template" 12 388 $halfW 34 $BgInput
$btnLoad.FlatAppearance.BorderColor=$FgDim; $btnLoad.FlatAppearance.BorderSize=1; $pnlRight.Controls.Add($btnLoad)
$btnDelete = New-StyledButton "Delete" (12+$halfW+8) 388 $halfW 34 $Danger; $pnlRight.Controls.Add($btnDelete)

$lblSep = New-Object System.Windows.Forms.Label
$lblSep.Location=New-Object System.Drawing.Point(12,436); $lblSep.Size=New-Object System.Drawing.Size($innerW,1)
$lblSep.BackColor=[System.Drawing.Color]::FromArgb(70,70,70); $pnlRight.Controls.Add($lblSep)

# --- Apply In-Grid Changes ---
$lblApplyChanges = New-Object System.Windows.Forms.Label
$lblApplyChanges.Text="Apply Grid Changes"; $lblApplyChanges.Font=$FontBold; $lblApplyChanges.ForeColor=$FgOrange
$lblApplyChanges.Location=New-Object System.Drawing.Point(12,444); $lblApplyChanges.AutoSize=$true
$pnlRight.Controls.Add($lblApplyChanges)

$lblApplyInfo = New-Object System.Windows.Forms.Label
$lblApplyInfo.Text="Changes made to Status or Start Type`ndropdowns will be applied to the system."; $lblApplyInfo.Font=$FontSmall; $lblApplyInfo.ForeColor=$FgDim
$lblApplyInfo.Location=New-Object System.Drawing.Point(12,464); $lblApplyInfo.Size=New-Object System.Drawing.Size($innerW,30)
$pnlRight.Controls.Add($lblApplyInfo)

$btnApplyGridChanges = New-StyledButton "Apply Changes" 12 498 $innerW 36 $FgOrange
$btnApplyGridChanges.BackColor=[System.Drawing.Color]::FromArgb(180,110,30)
$pnlRight.Controls.Add($btnApplyGridChanges)

$lblSep2 = New-Object System.Windows.Forms.Label
$lblSep2.Location=New-Object System.Drawing.Point(12,542); $lblSep2.Size=New-Object System.Drawing.Size($innerW,1)
$lblSep2.BackColor=[System.Drawing.Color]::FromArgb(70,70,70); $pnlRight.Controls.Add($lblSep2)

# --- Apply Template Section ---
$lblApplySection = New-Object System.Windows.Forms.Label
$lblApplySection.Text="Apply Template (changes startup types)"; $lblApplySection.Font=$FontBold; $lblApplySection.ForeColor=$FgYellow
$lblApplySection.Location=New-Object System.Drawing.Point(12,550); $lblApplySection.Size=New-Object System.Drawing.Size($innerW,36)
$pnlRight.Controls.Add($lblApplySection)

$btnApplyDisable = New-StyledButton "Apply: Disable Checked" 12 586 $innerW 36 $Danger; $pnlRight.Controls.Add($btnApplyDisable)
$btnApplyRestore = New-StyledButton "Apply: Restore from Template" 12 630 $innerW 36 $Success; $pnlRight.Controls.Add($btnApplyRestore)
$btnOpenFolder = New-StyledButton "Open Templates Folder" 12 680 $innerW 30 $BgInput
$btnOpenFolder.FlatAppearance.BorderColor=$FgDim; $btnOpenFolder.FlatAppearance.BorderSize=1; $btnOpenFolder.Font=$FontSmall
$pnlRight.Controls.Add($btnOpenFolder)

# =====================================================
# DROPDOWN FILTER (anchored below header, uses script-scope)
# =====================================================
function Close-ActiveDropdown {
    if ($script:_activeDropdown -and -not $script:_activeDropdown.IsDisposed) {
        try { $script:_activeDropdown.Close() } catch {}
    }
    $script:_activeDropdown = $null
}

function Show-ColumnDropdown {
    param([int]$ColIndex)

    $colName = $dgv.Columns[$ColIndex].Name
    if ($colName -eq "Selected") { return }
    if (-not $script:DgvColToProperty.ContainsKey($colName)) { return }

    # Close any existing dropdown first
    Close-ActiveDropdown

    $prop = $script:DgvColToProperty[$colName]
    $label = $script:PropertyToLabel[$prop]

    # Position directly below the column header
    $colLeft = -$dgv.HorizontalScrollingOffset
    for ($c = 0; $c -lt $ColIndex; $c++) {
        if ($dgv.Columns[$c].Visible) { $colLeft += $dgv.Columns[$c].Width }
    }
    $ptScreen = $dgv.PointToScreen((New-Object System.Drawing.Point($colLeft, $dgv.ColumnHeadersHeight)))

    $ddW = [Math]::Max($dgv.Columns[$ColIndex].Width, 300)
    $ddW = [Math]::Min($ddW, 400)
    $ddH = 440

    $workArea = [System.Windows.Forms.Screen]::FromControl($dgv).WorkingArea
    if (($ptScreen.X + $ddW) -gt $workArea.Right) { $ptScreen = New-Object System.Drawing.Point(($workArea.Right - $ddW - 4), $ptScreen.Y) }
    if ($ptScreen.X -lt $workArea.Left) { $ptScreen = New-Object System.Drawing.Point($workArea.Left, $ptScreen.Y) }
    if (($ptScreen.Y + $ddH) -gt $workArea.Bottom) {
        $headerTop = $dgv.PointToScreen((New-Object System.Drawing.Point($colLeft, 0)))
        $ptScreen = New-Object System.Drawing.Point($ptScreen.X, ($headerTop.Y - $ddH))
    }

    # Current sort for this column
    $curSortDir = ""
    for ($i = 0; $i -lt $script:SortLevels.Count; $i++) {
        if ($script:SortLevels[$i].Column -eq $prop) { $curSortDir = $script:SortLevels[$i].Direction; break }
    }
    $hasFilter = $script:ColumnFilters.ContainsKey($prop)
    $currentFilter = if ($hasFilter) { $script:ColumnFilters[$prop] } else { $null }

    # All unique values
    $allValues = @($script:AllItems | ForEach-Object { $_.$prop } | Sort-Object -Unique)

    $script:_ddAllValues = $allValues
    $script:_ddProp = $prop

    # Build dropdown form
    $dd = New-Object System.Windows.Forms.Form
    $dd.FormBorderStyle = "None"; $dd.Size = New-Object System.Drawing.Size($ddW, $ddH)
    $dd.StartPosition = "Manual"; $dd.Location = $ptScreen
    $dd.BackColor = $BgDropdown; $dd.ForeColor = $FgWhite; $dd.Font = $FontMain
    $dd.ShowInTaskbar = $false; $dd.TopMost = $false; $dd.Owner = $form
    $dd.Tag = "dropdown"

    $dd.Add_Paint({
        param($s, $e)
        $rect = $s.ClientRectangle
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(100,100,100), 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $rect.Width-1, $rect.Height-1); $pen.Dispose()
        $apen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0,122,204), 2)
        $e.Graphics.DrawLine($apen, 1, 1, $rect.Width-2, 1); $apen.Dispose()
    })

    $dd.Add_Deactivate({
        $f = $script:_activeDropdown
        if ($f -and -not $f.IsDisposed) {
            $f.BeginInvoke([Action]{ Close-ActiveDropdown })
        }
    })

    # Store reference
    $script:_activeDropdown = $dd

    $yOff = 8
    $lblCol = New-Object System.Windows.Forms.Label
    $lblCol.Text = $label; $lblCol.Font = $FontBold; $lblCol.ForeColor = $FgWhite
    $lblCol.Location = New-Object System.Drawing.Point(10, $yOff); $lblCol.AutoSize = $true
    $dd.Controls.Add($lblCol); $yOff += 24

    # Sort buttons
    $bw3 = [int](($ddW-30)/3)
    $btnSortAZ = New-Object System.Windows.Forms.Button
    $btnSortAZ.Text = "Sort A-Z"; $btnSortAZ.Location = New-Object System.Drawing.Point(10, $yOff)
    $btnSortAZ.Size = New-Object System.Drawing.Size($bw3, 28); $btnSortAZ.FlatStyle = "Flat"
    $btnSortAZ.BackColor = if ($curSortDir -eq "ASC") { $Accent } else { $BgInput }
    $btnSortAZ.ForeColor = $FgWhite; $btnSortAZ.Font = $FontSmall
    $btnSortAZ.FlatAppearance.BorderColor = $FgDim; $btnSortAZ.FlatAppearance.BorderSize = 1; $btnSortAZ.Cursor = "Hand"
    $dd.Controls.Add($btnSortAZ)

    $btnSortZA = New-Object System.Windows.Forms.Button
    $btnSortZA.Text = "Sort Z-A"; $btnSortZA.Location = New-Object System.Drawing.Point((10+$bw3+5), $yOff)
    $btnSortZA.Size = New-Object System.Drawing.Size($bw3, 28); $btnSortZA.FlatStyle = "Flat"
    $btnSortZA.BackColor = if ($curSortDir -eq "DESC") { $Accent } else { $BgInput }
    $btnSortZA.ForeColor = $FgWhite; $btnSortZA.Font = $FontSmall
    $btnSortZA.FlatAppearance.BorderColor = $FgDim; $btnSortZA.FlatAppearance.BorderSize = 1; $btnSortZA.Cursor = "Hand"
    $dd.Controls.Add($btnSortZA)

    $btnSortClear = New-Object System.Windows.Forms.Button
    $btnSortClear.Text = "No Sort"; $btnSortClear.Location = New-Object System.Drawing.Point((10+2*$bw3+10), $yOff)
    $btnSortClear.Size = New-Object System.Drawing.Size($bw3, 28); $btnSortClear.FlatStyle = "Flat"
    $btnSortClear.BackColor = $BgInput; $btnSortClear.ForeColor = $FgDim; $btnSortClear.Font = $FontSmall
    $btnSortClear.FlatAppearance.BorderColor = $FgDim; $btnSortClear.FlatAppearance.BorderSize = 1; $btnSortClear.Cursor = "Hand"
    $dd.Controls.Add($btnSortClear); $yOff += 36

    $sep = New-Object System.Windows.Forms.Label
    $sep.Location=New-Object System.Drawing.Point(10,$yOff); $sep.Size=New-Object System.Drawing.Size(($ddW-20),1)
    $sep.BackColor=[System.Drawing.Color]::FromArgb(80,80,80)
    $dd.Controls.Add($sep); $yOff += 6

    $lblFilt = New-Object System.Windows.Forms.Label
    $lblFilt.Text="Filter values:"; $lblFilt.Font=$FontSmall; $lblFilt.ForeColor=$FgDim
    $lblFilt.Location=New-Object System.Drawing.Point(10,$yOff); $lblFilt.AutoSize=$true
    $dd.Controls.Add($lblFilt); $yOff += 18

    # Search box
    $txtFS = New-Object System.Windows.Forms.TextBox
    $txtFS.Location=New-Object System.Drawing.Point(10,$yOff); $txtFS.Size=New-Object System.Drawing.Size(($ddW-20),22)
    $txtFS.BackColor=$BgInput; $txtFS.ForeColor=$FgWhite; $txtFS.BorderStyle="FixedSingle"; $txtFS.Font=$FontSmall
    $dd.Controls.Add($txtFS); $yOff += 26

    # Select All / Deselect / Invert
    $btnFA = New-Object System.Windows.Forms.Button
    $btnFA.Text="Select All"; $btnFA.Location=New-Object System.Drawing.Point(10,$yOff); $btnFA.Size=New-Object System.Drawing.Size($bw3,24)
    $btnFA.FlatStyle="Flat"; $btnFA.BackColor=$BgInput; $btnFA.ForeColor=$FgWhite; $btnFA.Font=$FontSmall
    $btnFA.FlatAppearance.BorderColor=$FgDim; $btnFA.FlatAppearance.BorderSize=1; $btnFA.Cursor="Hand"
    $dd.Controls.Add($btnFA)

    $btnFN = New-Object System.Windows.Forms.Button
    $btnFN.Text="Deselect All"; $btnFN.Location=New-Object System.Drawing.Point((10+$bw3+5),$yOff); $btnFN.Size=New-Object System.Drawing.Size($bw3,24)
    $btnFN.FlatStyle="Flat"; $btnFN.BackColor=$BgInput; $btnFN.ForeColor=$FgWhite; $btnFN.Font=$FontSmall
    $btnFN.FlatAppearance.BorderColor=$FgDim; $btnFN.FlatAppearance.BorderSize=1; $btnFN.Cursor="Hand"
    $dd.Controls.Add($btnFN)

    $btnFI = New-Object System.Windows.Forms.Button
    $btnFI.Text="Invert"; $btnFI.Location=New-Object System.Drawing.Point((10+2*$bw3+10),$yOff); $btnFI.Size=New-Object System.Drawing.Size($bw3,24)
    $btnFI.FlatStyle="Flat"; $btnFI.BackColor=$BgInput; $btnFI.ForeColor=$FgWhite; $btnFI.Font=$FontSmall
    $btnFI.FlatAppearance.BorderColor=$FgDim; $btnFI.FlatAppearance.BorderSize=1; $btnFI.Cursor="Hand"
    $dd.Controls.Add($btnFI); $yOff += 30

    # CheckedListBox
    $clbH = $ddH - $yOff - 44
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location=New-Object System.Drawing.Point(10,$yOff)
    $clb.Size=New-Object System.Drawing.Size(($ddW-20),$clbH)
    $clb.BackColor=$BgInput; $clb.ForeColor=$FgWhite; $clb.Font=$FontSmall
    $clb.BorderStyle="FixedSingle"; $clb.CheckOnClick=$true
    $dd.Controls.Add($clb); $yOff += $clbH + 6

    # Populate CLB helper (rebuilds list based on search, preserving check states)
    # Store references for event handlers
    $script:_ddCLB = $clb
    $script:_ddSearch = $txtFS

    # Master check states keyed by display value (survives rebuilds)
    $script:_ddMasterStates = @{}
    for ($i = 0; $i -lt $allValues.Count; $i++) {
        $dv = if ([string]::IsNullOrEmpty($allValues[$i])) { "(Blank)" } else { $allValues[$i] }
        $script:_ddMasterStates[$dv] = if ($hasFilter) { $currentFilter -contains $allValues[$i] } else { $true }
    }

    # Rebuild visible list based on search, preserving hidden check states
    function script:Rebuild-DropdownList {
        $c = $script:_ddCLB
        $searchText = $script:_ddSearch.Text.Trim().ToLower()

        # Save check states of currently visible items back to master
        for ($vi = 0; $vi -lt $c.Items.Count; $vi++) {
            $script:_ddMasterStates[$c.Items[$vi].ToString()] = $c.GetItemChecked($vi)
        }

        # Rebuild list showing only items matching search
        $c.Items.Clear()
        foreach ($val in $script:_ddAllValues) {
            $dv = if ([string]::IsNullOrEmpty($val)) { "(Blank)" } else { $val }
            if ($searchText -and -not $dv.ToLower().Contains($searchText)) { continue }
            $checked = if ($script:_ddMasterStates.ContainsKey($dv)) { $script:_ddMasterStates[$dv] } else { $true }
            $c.Items.Add($dv, $checked) | Out-Null
        }
    }

    # Initial populate (no search text, shows all)
    script:Rebuild-DropdownList

    # Typing in search hides non-matching items
    $txtFS.Add_TextChanged({ script:Rebuild-DropdownList })

    # Select All - only visible items
    $btnFA.Add_Click({
        $c = $script:_ddCLB
        for ($i = 0; $i -lt $c.Items.Count; $i++) { $c.SetItemChecked($i, $true) }
    })
    # Deselect All - only visible items
    $btnFN.Add_Click({
        $c = $script:_ddCLB
        for ($i = 0; $i -lt $c.Items.Count; $i++) { $c.SetItemChecked($i, $false) }
    })
    # Invert - only visible items
    $btnFI.Add_Click({
        $c = $script:_ddCLB
        for ($i = 0; $i -lt $c.Items.Count; $i++) { $c.SetItemChecked($i, (-not $c.GetItemChecked($i))) }
    })

    # Sort buttons
    $btnSortAZ.Add_Click({
        $p = $script:_ddProp
        for ($i=$script:SortLevels.Count-1;$i -ge 0;$i--) {
            if ($script:SortLevels[$i].Column -eq $p) { $script:SortLevels.RemoveAt($i) }
        }
        if ($script:SortLevels.Count -ge $script:MaxSortLevels) { $script:SortLevels.RemoveAt(0) }
        $script:SortLevels.Add(@{Column=$p;Direction="ASC"}) | Out-Null
        Update-SortDisplay; Populate-Grid
        Close-ActiveDropdown
    })
    $btnSortZA.Add_Click({
        $p = $script:_ddProp
        for ($i=$script:SortLevels.Count-1;$i -ge 0;$i--) {
            if ($script:SortLevels[$i].Column -eq $p) { $script:SortLevels.RemoveAt($i) }
        }
        if ($script:SortLevels.Count -ge $script:MaxSortLevels) { $script:SortLevels.RemoveAt(0) }
        $script:SortLevels.Add(@{Column=$p;Direction="DESC"}) | Out-Null
        Update-SortDisplay; Populate-Grid
        Close-ActiveDropdown
    })
    $btnSortClear.Add_Click({
        $p = $script:_ddProp
        for ($i=$script:SortLevels.Count-1;$i -ge 0;$i--) {
            if ($script:SortLevels[$i].Column -eq $p) { $script:SortLevels.RemoveAt($i) }
        }
        Update-SortDisplay; Populate-Grid
        Close-ActiveDropdown
    })

    # Apply / Cancel
    $abw = [int](($ddW-30)/2)
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text="Apply Filter"; $btnApply.Location=New-Object System.Drawing.Point(10,$yOff); $btnApply.Size=New-Object System.Drawing.Size($abw,32)
    $btnApply.FlatStyle="Flat"; $btnApply.BackColor=$Accent; $btnApply.ForeColor=$FgWhite; $btnApply.Font=$FontBold
    $btnApply.FlatAppearance.BorderSize=0; $btnApply.Cursor="Hand"
    $dd.Controls.Add($btnApply)

    $btnCnl = New-Object System.Windows.Forms.Button
    $btnCnl.Text="Cancel"; $btnCnl.Location=New-Object System.Drawing.Point((10+$abw+10),$yOff); $btnCnl.Size=New-Object System.Drawing.Size($abw,32)
    $btnCnl.FlatStyle="Flat"; $btnCnl.BackColor=$BgInput; $btnCnl.ForeColor=$FgWhite; $btnCnl.Font=$FontBold
    $btnCnl.FlatAppearance.BorderColor=$FgDim; $btnCnl.FlatAppearance.BorderSize=1; $btnCnl.Cursor="Hand"
    $dd.Controls.Add($btnCnl)

    $btnApply.Add_Click({
        $p = $script:_ddProp
        $vals = $script:_ddAllValues
        $c = $script:_ddCLB

        # Sync visible check states back to master
        for ($vi = 0; $vi -lt $c.Items.Count; $vi++) {
            $script:_ddMasterStates[$c.Items[$vi].ToString()] = $c.GetItemChecked($vi)
        }

        # Build filter from full master state (includes hidden items)
        $checkedVals = @(); $allChecked = $true
        for ($i = 0; $i -lt $vals.Count; $i++) {
            $dv = if ([string]::IsNullOrEmpty($vals[$i])) { "(Blank)" } else { $vals[$i] }
            if ($script:_ddMasterStates[$dv]) { $checkedVals += $vals[$i] } else { $allChecked = $false }
        }
        if ($allChecked -or $checkedVals.Count -eq 0) { $script:ColumnFilters.Remove($p) }
        else { $script:ColumnFilters[$p] = $checkedVals }
        Update-FilterIndicator; Populate-Grid
        Close-ActiveDropdown
    })

    $btnCnl.Add_Click({ Close-ActiveDropdown })

    $dd.Show($form)
    $dd.Activate()
}

function Update-FilterIndicator {
    $count = $script:ColumnFilters.Count
    if ($count -eq 0) { $lblFilterIndicator.Text = "" }
    else {
        $names = @(); foreach ($key in $script:ColumnFilters.Keys) { $names += $script:PropertyToLabel[$key] }
        $lblFilterIndicator.Text = "Filtered: " + ($names -join ", ")
    }
}

# =====================================================
# SORT + DATA LOGIC
# =====================================================
$script:AllItems = @()

function Update-SortDisplay {
    if ($script:SortLevels.Count -eq 0) {
        $lblSortDisplay.Text = "No sort applied."
        $lblSortDisplay.ForeColor = $FgDim
    } else {
        $parts = @(); $n = 1
        foreach ($lv in $script:SortLevels) {
            $lbl = $script:PropertyToLabel[$lv.Column]
            $d = if ($lv.Direction -eq "ASC") { "A-Z" } else { "Z-A" }
            $parts += "$n. $lbl $d"; $n++
        }
        $lblSortDisplay.Text = "Sort:  " + ($parts -join "     ")
        $lblSortDisplay.ForeColor = $FgOrange
    }
}

function Get-SortedItems {
    param([array]$Items)
    if ($script:SortLevels.Count -eq 0) { return ($Items | Sort-Object -Property DisplayName) }
    $se = @(); $rev = @($script:SortLevels); [Array]::Reverse($rev)
    foreach ($lv in $rev) {
        $p=$lv.Column; $d=($lv.Direction -eq "DESC")
        $se += @{Expression=[scriptblock]::Create("`$_.$p");Descending=$d}
    }
    return ($Items | Sort-Object -Property $se)
}

function Update-Counts {
    $t=$dgv.Rows.Count
    $ck=@($dgv.Rows|Where-Object{$_.Cells["Selected"].Value -eq $true}).Count
    $dr=@($dgv.Rows|Where-Object{$_.Cells["Type"].Value -eq "Driver"}).Count
    $lblCounts.Text="Showing $t items ($dr drivers, $($t-$dr) services)  |  $ck checked"
}

function Populate-Grid {
    $script:_isPopulating = $true
    $script:_headerCheckState = $false

    $dgv.SuspendLayout()
    $dgv.Rows.Clear()

    $fType=$cmbType.SelectedItem; $fStatus=$cmbStatus.SelectedItem; $fStart=$cmbStartType.SelectedItem
    $search=$txtSearch.Text.Trim().ToLower(); $ac=$chkAutoCheck.Checked

    # Pre-build HashSets for column filters (O(1) lookups instead of O(n))
    $filterSets = @{}
    foreach ($fp in $script:ColumnFilters.Keys) {
        $hs = New-Object System.Collections.Generic.HashSet[string]
        foreach ($v in $script:ColumnFilters[$fp]) { $hs.Add([string]$v) | Out-Null }
        $filterSets[$fp] = $hs
    }

    $filtered = [System.Collections.ArrayList]::new()
    foreach ($item in $script:AllItems) {
        if ($fType -ne "All" -and $item.Type -ne $fType) { continue }
        if ($fStatus -ne "All" -and $item.Status -ne $fStatus) { continue }
        if ($fStart -ne "All" -and $item.StartupType -ne $fStart) { continue }
        if ($search) {
            if (-not ($item.Name.ToLower().Contains($search) -or $item.DisplayName.ToLower().Contains($search) -or
                $item.Description.ToLower().Contains($search) -or $item.Group.ToLower().Contains($search) -or
                $item.Company.ToLower().Contains($search) -or $item.Filename.ToLower().Contains($search) -or
                $item.FileDescription.ToLower().Contains($search))) { continue }
        }
        $skip = $false
        foreach ($fp in $filterSets.Keys) {
            if (-not $filterSets[$fp].Contains([string]$item.$fp)) { $skip = $true; break }
        }
        if ($skip) { continue }
        $filtered.Add($item) | Out-Null
    }

    $sorted = Get-SortedItems -Items $filtered

    foreach ($item in $sorted) {
        $idx=$dgv.Rows.Add(); $row=$dgv.Rows[$idx]
        $isD=($item.StartupType -eq "Disabled")
        $row.Cells["Selected"].Value=($ac -and $isD)
        $row.Cells["Type"].Value=$item.Type; $row.Cells["Group"].Value=$item.Group
        $row.Cells["DisplayName"].Value=$item.DisplayName; $row.Cells["Status"].Value=$item.Status
        $row.Cells["StartupType"].Value=$item.StartupType; $row.Cells["ErrorControl"].Value=$item.ErrorControl
        $row.Cells["SvcName"].Value=$item.Name; $row.Cells["Dependencies"].Value=$item.Dependencies
        $row.Cells["FileDescription"].Value=$item.FileDescription; $row.Cells["Company"].Value=$item.Company
        $row.Cells["ProductName"].Value=$item.ProductName; $row.Cells["Description"].Value=$item.Description
        $row.Cells["Filename"].Value=$item.Filename; $row.Cells["LastWriteTime"].Value=$item.LastWriteTime

        if ($item.Type -eq "Driver") { $row.Cells["Type"].Style.ForeColor=$FgCyan } else { $row.Cells["Type"].Style.ForeColor=$FgWhite }
        if ($item.Status -eq "Started") { $row.Cells["Status"].Style.ForeColor=$FgGreen } else { $row.Cells["Status"].Style.ForeColor=$FgDim }
        if ($isD) { $row.Cells["StartupType"].Style.ForeColor=$FgRed }
        elseif ($item.StartupType -eq "Automatic") { $row.Cells["StartupType"].Style.ForeColor=$FgGreen }
        elseif ($item.StartupType -eq "Boot" -or $item.StartupType -eq "System") { $row.Cells["StartupType"].Style.ForeColor=$FgCyan }
        else { $row.Cells["StartupType"].Style.ForeColor=$FgYellow }
    }

    $dgv.ResumeLayout()
    $script:_isPopulating = $false
    Update-Counts
}

function Refresh-AllItems {
    $lblStatus.Text="Loading drivers and services (reading file info + registry)..."
    $lblStatus.ForeColor=$FgYellow; $form.Refresh()
    $script:AllItems = @(Get-AllItems)
    Populate-Grid; Update-SortDisplay; Update-FilterIndicator
    $lblStatus.Text="Loaded $($script:AllItems.Count) items at $(Get-Date -Format 'HH:mm:ss')  |  Screen: ${screenW}x${screenH}"
    $lblStatus.ForeColor=$FgDim
}

function Refresh-TemplateList {
    $lstTemplates.Items.Clear()
    $tl = Get-TemplateList; if ($tl) { foreach ($t in $tl) { $lstTemplates.Items.Add($t) } }
}

function Get-CheckedNames {
    $n=@(); foreach ($r in $dgv.Rows) { if ($r.Cells["Selected"].Value -eq $true) { $n += $r.Cells["SvcName"].Value } }; return $n
}

# =====================================================
# EVENTS
# =====================================================

# Left-click header = dropdown
$dgv.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    if ($e.ColumnIndex -eq 0) {
        # Toggle header checkbox
        $script:_headerCheckState = -not $script:_headerCheckState
        foreach ($r in $dgv.Rows) { $r.Cells["Selected"].Value = $script:_headerCheckState }
        $dgv.InvalidateColumn(0)
        Update-Counts
    } else {
        Show-ColumnDropdown -ColIndex $e.ColumnIndex
    }
})

$btnClearSort.Add_Click({ $script:SortLevels.Clear(); Update-SortDisplay; Populate-Grid })
$btnClearFilters.Add_Click({
    $script:ColumnFilters.Clear(); Update-FilterIndicator; Populate-Grid
    $lblStatus.Text="All column filters cleared."; $lblStatus.ForeColor=$FgCyan
})
$btnClearAll.Add_Click({
    $script:SortLevels.Clear(); $script:ColumnFilters.Clear()
    Update-SortDisplay; Update-FilterIndicator; Populate-Grid
})
$btnResetDefault.Add_Click({
    $script:SortLevels.Clear()
    $script:SortLevels.Add(@{Column="Name";Direction="ASC"}) | Out-Null
    $script:SortLevels.Add(@{Column="StartupType";Direction="DESC"}) | Out-Null
    $script:SortLevels.Add(@{Column="Status";Direction="ASC"}) | Out-Null
    $script:SortLevels.Add(@{Column="Type";Direction="DESC"}) | Out-Null
    Update-SortDisplay; Populate-Grid
})

$btnRefresh.Add_Click({ Refresh-AllItems })
# Debounce timer for search - avoids re-filtering on every keystroke
$script:SearchTimer = New-Object System.Windows.Forms.Timer
$script:SearchTimer.Interval = 300
$script:SearchTimer.Add_Tick({
    $script:SearchTimer.Stop()
    Populate-Grid
})
$txtSearch.Add_TextChanged({
    $script:SearchTimer.Stop()
    $script:SearchTimer.Start()
})
$cmbType.Add_SelectedIndexChanged({ Populate-Grid })
$cmbStatus.Add_SelectedIndexChanged({ Populate-Grid })
$cmbStartType.Add_SelectedIndexChanged({ Populate-Grid })
$chkAutoCheck.Add_CheckedChanged({ Populate-Grid })

$dgv.Add_CellContentClick({
    param($s,$e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $dgv.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit); Update-Counts
    }
})

# Commit combo box edits immediately on change
$dgv.Add_CurrentCellDirtyStateChanged({
    if ($dgv.IsCurrentCellDirty) {
        $dgv.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

# One-time warning when changing Status or StartupType on Drivers or Critical items
$dgv.Add_CellValueChanged({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    if ($script:_isPopulating) { return }
    $colName = $dgv.Columns[$e.ColumnIndex].Name
    if ($colName -ne "Status" -and $colName -ne "StartupType") { return }

    if (-not $script:DriverWarningShown) {
        $row = $dgv.Rows[$e.RowIndex]
        $itemType = $row.Cells["Type"].Value
        $errCtrl = $row.Cells["ErrorControl"].Value
        if ($itemType -eq "Driver" -or $errCtrl -eq "Critical") {
            $warnMsg = "WARNING: You are modifying a "
            if ($itemType -eq "Driver") { $warnMsg += "DRIVER" }
            if ($errCtrl -eq "Critical") {
                if ($itemType -eq "Driver") { $warnMsg += " with " }
                $warnMsg += "CRITICAL error control"
            }
            $warnMsg += ".`n`nChanging the status or startup type of system-critical components can cause boot failures or system instability.`n`nMake sure you know what you are doing before applying changes."
            [System.Windows.Forms.MessageBox]::Show($warnMsg, "Critical Component Warning", "OK", "Warning")
            $script:DriverWarningShown = $true
        }
    }
})

# Style ComboBox editing controls to match dark theme
$dgv.Add_EditingControlShowing({
    param($sender, $e)
    $ctrl = $e.Control
    if ($ctrl -is [System.Windows.Forms.ComboBox]) {
        $ctrl.BackColor = $BgInput
        $ctrl.ForeColor = $FgWhite
        $ctrl.FlatStyle = "Flat"
    }
})

# Suppress ComboBox DataError (e.g. if value doesn't match dropdown list)
$dgv.Add_DataError({
    param($sender, $e)
    $e.ThrowException = $false
})

$btnSelectAll.Add_Click({
    foreach ($r in $dgv.Rows) { $r.Cells["Selected"].Value=$true }
    $script:_headerCheckState = $true; $dgv.InvalidateColumn(0); Update-Counts
})
$btnSelectNone.Add_Click({
    foreach ($r in $dgv.Rows) { $r.Cells["Selected"].Value=$false }
    $script:_headerCheckState = $false; $dgv.InvalidateColumn(0); Update-Counts
})
$btnOpenFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $TemplateFolder })

$btnSave.Add_Click({
    $name=$txtTemplateName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { [System.Windows.Forms.MessageBox]::Show("Enter a template name.","Missing Name","OK","Warning"); return }
    if ($name -match '[\\/:*?"<>|]') { [System.Windows.Forms.MessageBox]::Show("Invalid characters.","Invalid Name","OK","Warning"); return }
    $cn=Get-CheckedNames
    if ($cn.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No items checked.","Nothing Selected","OK","Warning"); return }
    $ex=Get-TemplateList
    if ($ex -contains $name) {
        $r=[System.Windows.Forms.MessageBox]::Show("Template '$name' exists. Overwrite?","Confirm","YesNo","Question")
        if ($r -ne "Yes") { return }
    }
    $id=@()
    foreach ($row in $dgv.Rows) {
        if ($row.Cells["Selected"].Value -eq $true) {
            $id += @{
                Type=$row.Cells["Type"].Value;Group=$row.Cells["Group"].Value
                DisplayName=$row.Cells["DisplayName"].Value;Status=$row.Cells["Status"].Value
                StartupType=$row.Cells["StartupType"].Value;ErrorControl=$row.Cells["ErrorControl"].Value
                Name=$row.Cells["SvcName"].Value;Dependencies=$row.Cells["Dependencies"].Value
                FileDescription=$row.Cells["FileDescription"].Value;Company=$row.Cells["Company"].Value
                ProductName=$row.Cells["ProductName"].Value;Description=$row.Cells["Description"].Value
                Filename=$row.Cells["Filename"].Value;LastWriteTime=$row.Cells["LastWriteTime"].Value
            }
        }
    }
    Save-Template -TemplateName $name -SelectedItems $id
    $lblStatus.Text="Saved template '$name' with $($id.Count) items."; $lblStatus.ForeColor=$FgGreen
    Refresh-TemplateList; $txtTemplateName.Text=""
})

$lstTemplates.Add_SelectedIndexChanged({
    $sel=$lstTemplates.SelectedItem
    if ($sel) {
        $tpl=Load-Template -TemplateName $sel
        if ($tpl) {
            $c=$tpl.Items.Count; $dc=@($tpl.Items|Where-Object{$_.Type -eq "Driver"}).Count
            $lblTplInfo.Text="Template: $sel`nCreated: $($tpl.CreatedDate)`n$c items ($dc drivers, $($c-$dc) services)"
            $lblTplInfo.ForeColor=$FgWhite
        }
    }
})

$btnLoad.Add_Click({
    $sel=$lstTemplates.SelectedItem
    if (-not $sel) { [System.Windows.Forms.MessageBox]::Show("Select a template.","No Template","OK","Warning"); return }
    $tpl=Load-Template -TemplateName $sel; if (-not $tpl) { return }
    $tn=@(); foreach ($s in $tpl.Items) { $tn += $s.Name }
    foreach ($r in $dgv.Rows) { $r.Cells["Selected"].Value=($tn -contains $r.Cells["SvcName"].Value) }
    Update-Counts; $lblStatus.Text="Previewing template '$sel'."; $lblStatus.ForeColor=$Accent
})

$btnDelete.Add_Click({
    $sel=$lstTemplates.SelectedItem
    if (-not $sel) { [System.Windows.Forms.MessageBox]::Show("Select a template.","No Template","OK","Warning"); return }
    $r=[System.Windows.Forms.MessageBox]::Show("Delete template '$sel'?","Confirm","YesNo","Question")
    if ($r -eq "Yes") {
        Delete-Template -TemplateName $sel; Refresh-TemplateList
        $lblTplInfo.Text="Select a template to preview or apply."; $lblTplInfo.ForeColor=$FgDim
        $lblStatus.Text="Deleted template '$sel'."; $lblStatus.ForeColor=$FgRed
    }
})

$btnApplyGridChanges.Add_Click({
    $changes = @()
    foreach ($row in $dgv.Rows) {
        $svcName = $row.Cells["SvcName"].Value
        $newStatus = $row.Cells["Status"].Value
        $newStartType = $row.Cells["StartupType"].Value

        # Find the original item
        $original = $script:AllItems | Where-Object { $_.Name -eq $svcName } | Select-Object -First 1
        if (-not $original) { continue }

        $statusChanged = ($newStatus -ne $original.Status)
        $startChanged = ($newStartType -ne $original.StartupType)

        if ($statusChanged -or $startChanged) {
            $changes += @{
                Name = $svcName
                DisplayName = $row.Cells["DisplayName"].Value
                Type = $row.Cells["Type"].Value
                OldStatus = $original.Status; NewStatus = $newStatus
                OldStartType = $original.StartupType; NewStartType = $newStartType
                StatusChanged = $statusChanged; StartChanged = $startChanged
            }
        }
    }

    if ($changes.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No changes detected in Status or Start Type columns.", "No Changes", "OK", "Information")
        return
    }

    $summary = "The following $($changes.Count) change(s) will be applied:`n`n"
    foreach ($ch in $changes) {
        $summary += "$($ch.DisplayName) ($($ch.Name)):`n"
        if ($ch.StartChanged) { $summary += "  Start Type: $($ch.OldStartType) -> $($ch.NewStartType)`n" }
        if ($ch.StatusChanged) { $summary += "  Status: $($ch.OldStatus) -> $($ch.NewStatus)`n" }
        $summary += "`n"
    }
    $summary += "Proceed?"

    $confirm = [System.Windows.Forms.MessageBox]::Show($summary, "Confirm Apply Changes", "YesNo", "Question")
    if ($confirm -ne "Yes") { return }

    $ok = 0; $fl = 0
    foreach ($ch in $changes) {
        # Apply startup type change first
        if ($ch.StartChanged) {
            try {
                $stType = switch ($ch.NewStartType) {
                    "Boot"      { "Automatic" }
                    "System"    { "Automatic" }
                    "Automatic" { "Automatic" }
                    "Manual"    { "Manual" }
                    "Disabled"  { "Disabled" }
                    default     { "Manual" }
                }
                Set-Service -Name $ch.Name -StartupType $stType -ErrorAction Stop
            } catch { $fl++; continue }
        }
        # Apply status change
        if ($ch.StatusChanged) {
            try {
                if ($ch.NewStatus -eq "Started") {
                    Start-Service -Name $ch.Name -ErrorAction Stop
                } else {
                    Stop-Service -Name $ch.Name -Force -ErrorAction Stop
                }
            } catch { $fl++; continue }
        }
        $ok++
    }

    $lblStatus.Text = "Applied $ok changes. $fl failed."
    $lblStatus.ForeColor = if ($fl -gt 0) { $FgYellow } else { $FgGreen }
    Refresh-AllItems
})

$btnApplyDisable.Add_Click({
    $ck=Get-CheckedNames
    if ($ck.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No items checked.","Nothing Selected","OK","Warning"); return }
    $r=[System.Windows.Forms.MessageBox]::Show("Set $($ck.Count) items to DISABLED?`n`nRequires admin.","Confirm","YesNo","Warning")
    if ($r -ne "Yes") { return }
    $ok=0;$fl=0
    foreach ($sn in $ck) { try { Set-Service -Name $sn -StartupType Disabled -ErrorAction Stop; $ok++ } catch { $fl++ } }
    $lblStatus.Text="Disabled $ok items. $fl failed."; $lblStatus.ForeColor=if($fl -gt 0){$FgYellow}else{$FgGreen}
    Refresh-AllItems
})

$btnApplyRestore.Add_Click({
    $sel=$lstTemplates.SelectedItem
    if (-not $sel) { [System.Windows.Forms.MessageBox]::Show("Select a template.","No Template","OK","Warning"); return }
    $tpl=Load-Template -TemplateName $sel; if (-not $tpl) { return }
    $r=[System.Windows.Forms.MessageBox]::Show("Restore $($tpl.Items.Count) items from '$sel'?`n`nRequires admin.","Confirm","YesNo","Warning")
    if ($r -ne "Yes") { return }
    $ok=0;$fl=0
    foreach ($it in $tpl.Items) {
        try {
            $st = switch ($it.StartupType) { "Boot"{"Automatic"} "System"{"Automatic"} "Automatic"{"Automatic"} "Manual"{"Manual"} "Disabled"{"Disabled"} default{"Manual"} }
            Set-Service -Name $it.Name -StartupType $st -ErrorAction Stop; $ok++
        } catch { $fl++ }
    }
    $lblStatus.Text="Restored $ok items. $fl failed."; $lblStatus.ForeColor=if($fl -gt 0){$FgYellow}else{$FgGreen}
    Refresh-AllItems
})

$btnExportCSV.Add_Click({
    $ck=Get-CheckedNames
    if ($ck.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No items checked.","Nothing Selected","OK","Warning"); return }
    $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter="CSV Files|*.csv"
    $dlg.FileName="ServiceSnapshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq "OK") {
        $rows=@()
        foreach ($row in $dgv.Rows) {
            if ($row.Cells["Selected"].Value -eq $true) {
                $rows += [PSCustomObject]@{
                    Type=$row.Cells["Type"].Value;Group=$row.Cells["Group"].Value
                    "Display Name"=$row.Cells["DisplayName"].Value;Status=$row.Cells["Status"].Value
                    "Startup Type"=$row.Cells["StartupType"].Value;ErrorControl=$row.Cells["ErrorControl"].Value
                    Name=$row.Cells["SvcName"].Value;Dependencies=$row.Cells["Dependencies"].Value
                    "File Description"=$row.Cells["FileDescription"].Value;Company=$row.Cells["Company"].Value
                    "Product Name"=$row.Cells["ProductName"].Value;Description=$row.Cells["Description"].Value
                    Filename=$row.Cells["Filename"].Value;"Last Write Time"=$row.Cells["LastWriteTime"].Value
                }
            }
        }
        $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        $lblStatus.Text="Exported $($rows.Count) items to CSV."; $lblStatus.ForeColor=$FgGreen
    }
})

# --- Init ---
# Data load is deferred to Form.Shown event for instant window appearance

$form.Add_Shown({
    $form.Activate()
    # Defer data load so the form appears immediately
    $form.BeginInvoke([Action]{
        Refresh-AllItems
        Refresh-TemplateList
    })
})

# Cleanup pre-allocated GDI objects
$form.Add_FormClosed({
    $script:PaintBrushBg.Dispose()
    $script:PaintPenBorder.Dispose()
    $script:PaintPenCbBorder.Dispose()
    $script:PaintBrushCbOff.Dispose()
    $script:PaintBrushCbOn.Dispose()
    $script:PaintPenCheck.Dispose()
    $script:PaintBrushWhite.Dispose()
    $script:PaintBrushCyan.Dispose()
    $script:PaintBrushArrow.Dispose()
    $script:PaintSFText.Dispose()
    $script:PaintSFArrow.Dispose()
    $script:SearchTimer.Dispose()
})

[void]$form.ShowDialog()
$form.Dispose()
