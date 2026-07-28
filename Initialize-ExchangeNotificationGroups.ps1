<#
.SYNOPSIS
    Creates and maintains standard Exchange Online notification contacts and distribution groups.

.DESCRIPTION
    Opens a Windows Forms configuration interface where the operator defines the
    Exchange Online connection, SMTP domain, sender policy, report location, and
    notification groups and members to provision.

    The script runs in a clean Windows PowerShell process for stable use from VS Code,
    installs and imports ExchangeOnlineManagement when required, connects with modern
    authentication, creates missing mail contacts and distribution groups, adds missing
    members, applies the sender policy, displays progress, and writes a timestamped CSV
    execution report.

    No naming prefix or organization-specific identifier is required.

.PARAMETER ExchangeModuleVersion
    Exact ExchangeOnlineManagement module version to install and import.
    The default value is 3.9.2, which supports Windows PowerShell 5.1.

.PARAMETER IsolatedProcess
    Internal parameter used when the script relaunches itself in a clean process.

.EXAMPLE
    .\Initialize-ExchangeNotificationGroups.ps1

.INPUTS
    None.

.OUTPUTS
    Exchange Online mail contacts and distribution groups, plus an optional timestamped
    CSV execution report.

.NOTES
    File Name : Initialize-ExchangeNotificationGroups.ps1
    Author    : Juan Ortiz
    Version   : 2.0.1
    Created   : 2026-07-27
    Updated   : 2026-07-27

    Requirements:
      - Windows PowerShell 5.1
      - Windows desktop session
      - Internet access
      - Permission to install modules under the current user profile
      - Exchange Online permissions to manage recipients and groups

.LINK
    https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell

.LINK
    https://learn.microsoft.com/powershell/module/exchangepowershell/new-mailcontact

.LINK
    https://learn.microsoft.com/powershell/module/exchangepowershell/new-distributiongroup

.LINK
    https://learn.microsoft.com/powershell/module/exchangepowershell/add-distributiongroupmember
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [version]$ExchangeModuleVersion = [version]'3.9.2',

    [Parameter(Mandatory = $false, DontShow = $true)]
    [switch]$IsolatedProcess
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Process isolation

if (-not $IsolatedProcess) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'Save the script as a .ps1 file before running it.'
    }

    $powerShellEngine = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'

    if (-not (Test-Path -LiteralPath $powerShellEngine)) {
        throw "Windows PowerShell was not found at '$powerShellEngine'."
    }

    $childArguments = @(
        '-NoLogo'
        '-NoProfile'
        '-STA'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        $PSCommandPath
        '-IsolatedProcess'
        '-ExchangeModuleVersion'
        $ExchangeModuleVersion.ToString()
    )

    Write-Host 'Starting Exchange Online notification group provisioning in a clean Windows PowerShell process...' -ForegroundColor Cyan
    Write-Host "PowerShell engine      : $powerShellEngine" -ForegroundColor DarkGray
    Write-Host "Exchange module version: $ExchangeModuleVersion" -ForegroundColor DarkGray
    Write-Host ''

    & $powerShellEngine @childArguments
    $childExitCode = $LASTEXITCODE

    if ($childExitCode -ne 0) {
        throw "The isolated process ended with exit code $childExitCode."
    }

    return
}

#endregion Process isolation

#region Default configuration

$defaultGroupDefinitions = @(
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Microsoft'; LocalPart = 'microsoft'; BaseMembers = 'microsoft@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Support'; LocalPart = 'support'; BaseMembers = 'support@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Ingram'; LocalPart = 'ingram'; BaseMembers = 'ingram.customer@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Dark Web ID'; LocalPart = 'dark.web.id'; BaseMembers = 'dark.web.id@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Datto RMM'; LocalPart = 'datto.rmm'; BaseMembers = 'datto.rmm@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Lenovo'; LocalPart = 'lenovo'; BaseMembers = 'lenovo@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Mailjet'; LocalPart = 'mailjet'; BaseMembers = 'support@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'SaaS Alerts'; LocalPart = 'saas.alerts'; BaseMembers = 'saas.alerts@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Spanning Backup Fail'; LocalPart = 'spanning.backup.fail'; BaseMembers = 'spanning.backup@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Spanning Backup Success'; LocalPart = 'spanning.backup.success'; BaseMembers = 'spanning.backup@toconnect.net'; AdditionalMembers = '' }
    [pscustomobject]@{ Enabled = $true; DisplayName = 'Ubiquiti'; LocalPart = 'ubiquiti'; BaseMembers = 'ubiquiti@toconnect.net'; AdditionalMembers = '' }
)

#endregion Default configuration

#region Script state

$script:ExchangeConnectionCreated = $false
$script:Results = New-Object System.Collections.ArrayList
$script:RecipientCache = @{}
$script:PlannedContactAddresses = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

#endregion Script state

#region General helpers

function Enable-Tls12 {
    [CmdletBinding()]
    param()

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}

function Get-SafePropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$PropertyNames
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($propertyName in $PropertyNames) {
        $property = $InputObject.PSObject.Properties[$propertyName]

        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $null
}

function Split-AddressList {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        $Value -split '[,;\r\n]+' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function Normalize-SmtpAddress {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Address
    )

    $normalizedAddress = $Address.Trim()

    if ($normalizedAddress.StartsWith('smtp:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $normalizedAddress = $normalizedAddress.Substring(5)
    }

    return $normalizedAddress.ToLowerInvariant()
}

function Assert-ValidEmailAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Address
    )

    $normalizedAddress = Normalize-SmtpAddress -Address $Address

    try {
        $parsedAddress = New-Object System.Net.Mail.MailAddress -ArgumentList $normalizedAddress
    }
    catch {
        throw "'$Address' is not a valid email address."
    }

    if (-not [string]::Equals($parsedAddress.Address, $normalizedAddress, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "'$Address' is not a valid standalone email address."
    }
}

function Escape-OPathValue {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function ConvertTo-ExchangeAlias {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    $alias = $Value.ToLowerInvariant()
    $alias = $alias -replace '@', '-at-'
    $alias = $alias -replace '[^a-z0-9._-]', '-'
    $alias = $alias -replace '-{2,}', '-'
    $alias = $alias -replace '\.{2,}', '.'
    $alias = $alias.Trim([char[]]'.-_')

    if ([string]::IsNullOrWhiteSpace($alias)) {
        throw "A valid Exchange alias could not be generated from '$Value'."
    }

    if ($alias.Length -gt 64) {
        $alias = $alias.Substring(0, 64).Trim([char[]]'.-_')
    }

    return $alias
}

function Add-OperationResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Phase,
        [Parameter(Mandatory = $true)][string]$ObjectType,
        [Parameter(Mandatory = $false)][AllowNull()][string]$Name,
        [Parameter(Mandatory = $false)][AllowNull()][string]$Address,
        [Parameter(Mandatory = $true)][string]$Action,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $false)][AllowNull()][string]$Message
    )

    [void]$script:Results.Add(
        [pscustomobject][ordered]@{
            Timestamp  = (Get-Date).ToString('o')
            Phase      = $Phase
            ObjectType = $ObjectType
            Name       = $Name
            Address    = $Address
            Action     = $Action
            Status     = $Status
            Message    = $Message
        }
    )
}

function Write-OperationStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$Message
    )

    $color = switch ($Status) {
        'Created'  { 'Green' }
        'Added'    { 'Green' }
        'Updated'  { 'Green' }
        'Existing' { 'Yellow' }
        'Skipped'  { 'DarkGray' }
        'Preview'  { 'Cyan' }
        'Conflict' { 'Red' }
        'Failed'   { 'Red' }
        default    { 'White' }
    }

    Write-Host ("[{0}] {1}" -f $Status, $Message) -ForegroundColor $color
}

function Test-ProvisioningConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Configuration
    )

    if (-not [string]::IsNullOrWhiteSpace($Configuration.UserPrincipalName)) {
        Assert-ValidEmailAddress -Address $Configuration.UserPrincipalName
    }

    if (
        -not [string]::IsNullOrWhiteSpace($Configuration.PrimaryDomain) -and
        $Configuration.PrimaryDomain -notmatch '^(?=.{1,253}$)([a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$'
    ) {
        throw "'$($Configuration.PrimaryDomain)' is not a valid DNS domain."
    }

    if ($Configuration.CreateReport -and [string]::IsNullOrWhiteSpace($Configuration.ReportDirectory)) {
        throw 'A report directory is required when report creation is enabled.'
    }

    $enabledGroups = @($Configuration.Groups | Where-Object { $_.Enabled })

    if ($enabledGroups.Count -eq 0) {
        throw 'Enable at least one distribution group.'
    }

    $duplicateLocalParts = @(
        $enabledGroups |
            Group-Object -Property LocalPart |
            Where-Object { $_.Count -gt 1 }
    )

    if ($duplicateLocalParts.Count -gt 0) {
        throw ('Duplicate SMTP local parts were found: ' + (($duplicateLocalParts.Name) -join ', '))
    }

    foreach ($group in $enabledGroups) {
        if ([string]::IsNullOrWhiteSpace($group.DisplayName)) {
            throw 'Every enabled group must have a display name.'
        }

        if ($group.LocalPart -notmatch '^[a-z0-9]+([._-][a-z0-9]+)*$') {
            throw "SMTP local part '$($group.LocalPart)' contains unsupported characters."
        }

        foreach ($memberAddress in @($group.BaseMembers + $group.AdditionalMembers)) {
            Assert-ValidEmailAddress -Address $memberAddress
        }
    }
}

#endregion General helpers

#region Graphical configuration interface

function Show-ConfigurationDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DefaultGroups,

        [Parameter(Mandatory = $true)]
        [version]$DefaultModuleVersion
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Exchange Online Notification Groups'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(1060, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(920, 650)
    $form.ShowIcon = $false
    $form.MaximizeBox = $true

    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Text = 'Configure Exchange Online notification contacts and groups'
    $headerLabel.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $headerLabel.AutoSize = $true
    $headerLabel.Location = New-Object System.Drawing.Point(18, 16)
    $form.Controls.Add($headerLabel)

    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Text = 'Enter the connection and SMTP settings. Edit the group list as needed. Separate members with commas, semicolons, or new lines.'
    $descriptionLabel.AutoSize = $true
    $descriptionLabel.Location = New-Object System.Drawing.Point(20, 48)
    $form.Controls.Add($descriptionLabel)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(18, 78)
    $tabControl.Size = New-Object System.Drawing.Size(1006, 570)
    $tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($tabControl)

    $generalTab = New-Object System.Windows.Forms.TabPage
    $generalTab.Text = 'General'
    $tabControl.TabPages.Add($generalTab)

    $groupsTab = New-Object System.Windows.Forms.TabPage
    $groupsTab.Text = 'Groups and Members'
    $tabControl.TabPages.Add($groupsTab)

    $connectionGroup = New-Object System.Windows.Forms.GroupBox
    $connectionGroup.Text = 'Exchange Online connection'
    $connectionGroup.Location = New-Object System.Drawing.Point(16, 16)
    $connectionGroup.Size = New-Object System.Drawing.Size(952, 174)
    $connectionGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $generalTab.Controls.Add($connectionGroup)

    $upnLabel = New-Object System.Windows.Forms.Label
    $upnLabel.Text = 'Administrator UPN'
    $upnLabel.AutoSize = $true
    $upnLabel.Location = New-Object System.Drawing.Point(18, 32)
    $connectionGroup.Controls.Add($upnLabel)

    $upnTextBox = New-Object System.Windows.Forms.TextBox
    $upnTextBox.Location = New-Object System.Drawing.Point(190, 28)
    $upnTextBox.Size = New-Object System.Drawing.Size(730, 24)
    $upnTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $connectionGroup.Controls.Add($upnTextBox)

    $upnHint = New-Object System.Windows.Forms.Label
    $upnHint.Text = 'Optional. Leave blank to select the account during sign-in.'
    $upnHint.AutoSize = $true
    $upnHint.ForeColor = [System.Drawing.Color]::DimGray
    $upnHint.Location = New-Object System.Drawing.Point(190, 54)
    $connectionGroup.Controls.Add($upnHint)

    $delegatedLabel = New-Object System.Windows.Forms.Label
    $delegatedLabel.Text = 'Delegated organization'
    $delegatedLabel.AutoSize = $true
    $delegatedLabel.Location = New-Object System.Drawing.Point(18, 88)
    $connectionGroup.Controls.Add($delegatedLabel)

    $delegatedTextBox = New-Object System.Windows.Forms.TextBox
    $delegatedTextBox.Location = New-Object System.Drawing.Point(190, 84)
    $delegatedTextBox.Size = New-Object System.Drawing.Size(730, 24)
    $delegatedTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $connectionGroup.Controls.Add($delegatedTextBox)

    $delegatedHint = New-Object System.Windows.Forms.Label
    $delegatedHint.Text = 'Optional. Enter a tenant domain or tenant ID only for delegated administration.'
    $delegatedHint.AutoSize = $true
    $delegatedHint.ForeColor = [System.Drawing.Color]::DimGray
    $delegatedHint.Location = New-Object System.Drawing.Point(190, 110)
    $connectionGroup.Controls.Add($delegatedHint)

    $moduleLabel = New-Object System.Windows.Forms.Label
    $moduleLabel.Text = 'Module version'
    $moduleLabel.AutoSize = $true
    $moduleLabel.Location = New-Object System.Drawing.Point(18, 142)
    $connectionGroup.Controls.Add($moduleLabel)

    $moduleTextBox = New-Object System.Windows.Forms.TextBox
    $moduleTextBox.Location = New-Object System.Drawing.Point(190, 138)
    $moduleTextBox.Size = New-Object System.Drawing.Size(150, 24)
    $moduleTextBox.Text = $DefaultModuleVersion.ToString()
    $connectionGroup.Controls.Add($moduleTextBox)

    $settingsGroup = New-Object System.Windows.Forms.GroupBox
    $settingsGroup.Text = 'Provisioning settings'
    $settingsGroup.Location = New-Object System.Drawing.Point(16, 204)
    $settingsGroup.Size = New-Object System.Drawing.Size(952, 268)
    $settingsGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $generalTab.Controls.Add($settingsGroup)

    $domainLabel = New-Object System.Windows.Forms.Label
    $domainLabel.Text = 'Primary SMTP domain'
    $domainLabel.AutoSize = $true
    $domainLabel.Location = New-Object System.Drawing.Point(18, 34)
    $settingsGroup.Controls.Add($domainLabel)

    $domainTextBox = New-Object System.Windows.Forms.TextBox
    $domainTextBox.Location = New-Object System.Drawing.Point(190, 30)
    $domainTextBox.Size = New-Object System.Drawing.Size(730, 24)
    $domainTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $settingsGroup.Controls.Add($domainTextBox)

    $domainHint = New-Object System.Windows.Forms.Label
    $domainHint.Text = 'Optional. Leave blank to detect the default accepted domain.'
    $domainHint.AutoSize = $true
    $domainHint.ForeColor = [System.Drawing.Color]::DimGray
    $domainHint.Location = New-Object System.Drawing.Point(190, 56)
    $settingsGroup.Controls.Add($domainHint)

    $allowExternalCheckBox = New-Object System.Windows.Forms.CheckBox
    $allowExternalCheckBox.Text = 'Allow unauthenticated external senders'
    $allowExternalCheckBox.AutoSize = $true
    $allowExternalCheckBox.Checked = $true
    $allowExternalCheckBox.Location = New-Object System.Drawing.Point(190, 88)
    $settingsGroup.Controls.Add($allowExternalCheckBox)

    $previewCheckBox = New-Object System.Windows.Forms.CheckBox
    $previewCheckBox.Text = 'Preview changes only'
    $previewCheckBox.AutoSize = $true
    $previewCheckBox.Checked = $false
    $previewCheckBox.Location = New-Object System.Drawing.Point(190, 120)
    $settingsGroup.Controls.Add($previewCheckBox)

    $createReportCheckBox = New-Object System.Windows.Forms.CheckBox
    $createReportCheckBox.Text = 'Create timestamped CSV execution report'
    $createReportCheckBox.AutoSize = $true
    $createReportCheckBox.Checked = $true
    $createReportCheckBox.Location = New-Object System.Drawing.Point(190, 152)
    $settingsGroup.Controls.Add($createReportCheckBox)

    $logLabel = New-Object System.Windows.Forms.Label
    $logLabel.Text = 'Report output folder'
    $logLabel.AutoSize = $true
    $logLabel.Location = New-Object System.Drawing.Point(18, 194)
    $settingsGroup.Controls.Add($logLabel)

    $logTextBox = New-Object System.Windows.Forms.TextBox
    $logTextBox.Location = New-Object System.Drawing.Point(190, 190)
    $logTextBox.Size = New-Object System.Drawing.Size(612, 24)
    $logTextBox.Text = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { [Environment]::GetFolderPath('Desktop') }
    $logTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $settingsGroup.Controls.Add($logTextBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = 'Select Folder...'
    $browseButton.Location = New-Object System.Drawing.Point(810, 188)
    $browseButton.Size = New-Object System.Drawing.Size(110, 28)
    $browseButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $settingsGroup.Controls.Add($browseButton)

    $reportHint = New-Object System.Windows.Forms.Label
    $reportHint.Text = 'The directory is required only when report creation is enabled.'
    $reportHint.AutoSize = $true
    $reportHint.ForeColor = [System.Drawing.Color]::DimGray
    $reportHint.Location = New-Object System.Drawing.Point(190, 220)
    $settingsGroup.Controls.Add($reportHint)

    $gridInstruction = New-Object System.Windows.Forms.Label
    $gridInstruction.Text = 'Each enabled row creates or updates one distribution group. Base members are created as mail contacts when no recipient exists. Additional members must already exist as mail-enabled recipients.'
    $gridInstruction.Location = New-Object System.Drawing.Point(14, 14)
    $gridInstruction.Size = New-Object System.Drawing.Size(944, 38)
    $gridInstruction.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $groupsTab.Controls.Add($gridInstruction)

    $groupGrid = New-Object System.Windows.Forms.DataGridView
    $groupGrid.Location = New-Object System.Drawing.Point(14, 58)
    $groupGrid.Size = New-Object System.Drawing.Size(944, 430)
    $groupGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $groupGrid.AllowUserToAddRows = $true
    $groupGrid.AllowUserToDeleteRows = $true
    $groupGrid.AutoGenerateColumns = $false
    $groupGrid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $groupGrid.RowHeadersVisible = $false
    $groupGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $groupGrid.MultiSelect = $false
    $groupGrid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnKeystrokeOrF2
    $groupsTab.Controls.Add($groupGrid)

    $enabledColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $enabledColumn.Name = 'Enabled'
    $enabledColumn.HeaderText = 'Enabled'
    $enabledColumn.Width = 60
    [void]$groupGrid.Columns.Add($enabledColumn)

    $displayNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $displayNameColumn.Name = 'DisplayName'
    $displayNameColumn.HeaderText = 'Display name'
    $displayNameColumn.Width = 180
    [void]$groupGrid.Columns.Add($displayNameColumn)

    $localPartColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $localPartColumn.Name = 'LocalPart'
    $localPartColumn.HeaderText = 'SMTP local part'
    $localPartColumn.Width = 160
    [void]$groupGrid.Columns.Add($localPartColumn)

    $baseMembersColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $baseMembersColumn.Name = 'BaseMembers'
    $baseMembersColumn.HeaderText = 'Base members / contacts'
    $baseMembersColumn.Width = 250
    $baseMembersColumn.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    [void]$groupGrid.Columns.Add($baseMembersColumn)

    $additionalMembersColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $additionalMembersColumn.Name = 'AdditionalMembers'
    $additionalMembersColumn.HeaderText = 'Additional members'
    $additionalMembersColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $additionalMembersColumn.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    [void]$groupGrid.Columns.Add($additionalMembersColumn)

    foreach ($defaultGroup in $DefaultGroups) {
        $rowIndex = $groupGrid.Rows.Add()
        $row = $groupGrid.Rows[$rowIndex]
        $row.Cells['Enabled'].Value = [bool]$defaultGroup.Enabled
        $row.Cells['DisplayName'].Value = [string]$defaultGroup.DisplayName
        $row.Cells['LocalPart'].Value = [string]$defaultGroup.LocalPart
        $row.Cells['BaseMembers'].Value = [string]$defaultGroup.BaseMembers
        $row.Cells['AdditionalMembers'].Value = [string]$defaultGroup.AdditionalMembers
    }

    $addRowButton = New-Object System.Windows.Forms.Button
    $addRowButton.Text = 'Add Group'
    $addRowButton.Location = New-Object System.Drawing.Point(14, 500)
    $addRowButton.Size = New-Object System.Drawing.Size(100, 30)
    $addRowButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $groupsTab.Controls.Add($addRowButton)

    $removeRowButton = New-Object System.Windows.Forms.Button
    $removeRowButton.Text = 'Remove Group'
    $removeRowButton.Location = New-Object System.Drawing.Point(122, 500)
    $removeRowButton.Size = New-Object System.Drawing.Size(110, 30)
    $removeRowButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $groupsTab.Controls.Add($removeRowButton)

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = 'Run'
    $runButton.Size = New-Object System.Drawing.Size(110, 34)
    $runButton.Location = New-Object System.Drawing.Point(790, 666)
    $runButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($runButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.Size = New-Object System.Drawing.Size(110, 34)
    $cancelButton.Location = New-Object System.Drawing.Point(910, 666)
    $cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($cancelButton)

    # Shared folder selector used by the button and by the Run validation.
    # FolderBrowserDialog opens at runtime and writes the selected path directly
    # into the report output field.
    $selectReportFolder = {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = 'Select the folder for the timestamped CSV execution report.'
        $folderDialog.ShowNewFolderButton = $true

        if (-not [string]::IsNullOrWhiteSpace($logTextBox.Text) -and
            (Test-Path -LiteralPath $logTextBox.Text -PathType Container)) {

            $folderDialog.SelectedPath = $logTextBox.Text
        }
        elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and
            (Test-Path -LiteralPath $PSScriptRoot -PathType Container)) {

            $folderDialog.SelectedPath = $PSScriptRoot
        }

        $dialogResult = $folderDialog.ShowDialog($form)

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK -and
            -not [string]::IsNullOrWhiteSpace($folderDialog.SelectedPath)) {

            $logTextBox.Text = $folderDialog.SelectedPath
            $logTextBox.SelectionStart = $logTextBox.Text.Length
            $logTextBox.ScrollToCaret()
            return $true
        }

        return $false
    }.GetNewClosure()

    $browseButton.Add_Click({
        [void](& $selectReportFolder)
    })

    $createReportCheckBox.Add_CheckedChanged({
        $logTextBox.Enabled = $createReportCheckBox.Checked
        $browseButton.Enabled = $createReportCheckBox.Checked
    })

    $addRowButton.Add_Click({
        $rowIndex = $groupGrid.Rows.Add()
        $groupGrid.Rows[$rowIndex].Cells['Enabled'].Value = $true
        $groupGrid.CurrentCell = $groupGrid.Rows[$rowIndex].Cells['DisplayName']
        $groupGrid.BeginEdit($true)
    })

    $removeRowButton.Add_Click({
        if ($null -ne $groupGrid.CurrentRow -and -not $groupGrid.CurrentRow.IsNewRow) {
            $groupGrid.Rows.Remove($groupGrid.CurrentRow)
        }
    })

    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $runButton.Add_Click({
        try {
            $groupGrid.EndEdit()
            $parsedModuleVersion = $null

            if (-not [version]::TryParse($moduleTextBox.Text.Trim(), [ref]$parsedModuleVersion)) {
                throw 'Enter a valid ExchangeOnlineManagement module version.'
            }

            if ($createReportCheckBox.Checked) {
                $reportDirectoryIsValid = (
                    -not [string]::IsNullOrWhiteSpace($logTextBox.Text) -and
                    (Test-Path -LiteralPath $logTextBox.Text -PathType Container)
                )

                # Open the folder selector automatically when the report path is
                # empty or does not point to an existing directory.
                if (-not $reportDirectoryIsValid) {
                    $folderSelected = & $selectReportFolder

                    if (-not $folderSelected) {
                        throw 'Select a report output folder or disable report creation.'
                    }
                }
            }

            $groups = New-Object System.Collections.ArrayList

            foreach ($row in $groupGrid.Rows) {
                if ($row.IsNewRow) {
                    continue
                }

                $enabled = [bool]$row.Cells['Enabled'].Value
                $displayName = [string]$row.Cells['DisplayName'].Value
                $localPart = [string]$row.Cells['LocalPart'].Value
                $baseMembers = [string]$row.Cells['BaseMembers'].Value
                $additionalMembers = [string]$row.Cells['AdditionalMembers'].Value

                if (-not $enabled -and [string]::IsNullOrWhiteSpace($displayName) -and [string]::IsNullOrWhiteSpace($localPart)) {
                    continue
                }

                [void]$groups.Add(
                    [pscustomobject]@{
                        Enabled           = $enabled
                        DisplayName       = $displayName.Trim()
                        LocalPart         = $localPart.Trim().ToLowerInvariant()
                        BaseMembers       = @(Split-AddressList -Value $baseMembers)
                        AdditionalMembers = @(Split-AddressList -Value $additionalMembers)
                    }
                )
            }

            $configuration = [pscustomobject]@{
                UserPrincipalName     = $upnTextBox.Text.Trim()
                DelegatedOrganization = $delegatedTextBox.Text.Trim()
                PrimaryDomain         = $domainTextBox.Text.Trim().ToLowerInvariant()
                ModuleVersion         = $parsedModuleVersion
                AllowExternalSenders  = [bool]$allowExternalCheckBox.Checked
                PreviewOnly           = [bool]$previewCheckBox.Checked
                CreateReport          = [bool]$createReportCheckBox.Checked
                ReportDirectory       = $logTextBox.Text.Trim()
                Groups                = @($groups)
            }

            Test-ProvisioningConfiguration -Configuration $configuration

            $form.Tag = $configuration
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                $_.Exception.Message,
                'Configuration Error',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    })

    $form.AcceptButton = $runButton
    $form.CancelButton = $cancelButton

    $dialogResult = $form.ShowDialog()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK -or $null -eq $form.Tag) {
        throw [System.OperationCanceledException]::new('The operation was canceled.')
    }

    return $form.Tag
}

#endregion Graphical configuration interface

#region Module and connection

function Import-RequiredExchangeModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [version]$RequiredVersion
    )

    Enable-Tls12

    $moduleName = 'ExchangeOnlineManagement'
    $requiredVersionText = $RequiredVersion.ToString()

    $localApplicationData = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::LocalApplicationData
    )

    if ([string]::IsNullOrWhiteSpace($localApplicationData)) {
        $localApplicationData = $env:TEMP
    }

    $moduleCacheRoot = Join-Path `
        -Path $localApplicationData `
        -ChildPath 'ToConnect\PowerShellModules'

    $exactCacheVersionPath = Join-Path `
        -Path $moduleCacheRoot `
        -ChildPath (
            '{0}\{1}' -f
            $moduleName,
            $requiredVersionText
        )

    $documentsDirectory = [Environment]::GetFolderPath(
    [Environment+SpecialFolder]::MyDocuments
    )

    $findExactModuleManifest = {
        $candidatePaths = New-Object System.Collections.ArrayList

        $availableModules = @(
            Get-Module `
                -ListAvailable `
                -Name $moduleName `
                -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Version -eq $RequiredVersion
                }
        )

        foreach ($availableModule in $availableModules) {
            if (
                -not [string]::IsNullOrWhiteSpace(
                    [string]$availableModule.Path
                )
            ) {
                [void]$candidatePaths.Add(
                    [string]$availableModule.Path
                )
            }
        }

        $moduleRoots = New-Object System.Collections.ArrayList

        foreach ($moduleRoot in @($env:PSModulePath -split ';')) {
            if (-not [string]::IsNullOrWhiteSpace($moduleRoot)) {
                [void]$moduleRoots.Add($moduleRoot)
            }
        }

        $documentsDirectory = [Environment]::GetFolderPath(
            [Environment+SpecialFolder]::MyDocuments
        )

        if (-not [string]::IsNullOrWhiteSpace($documentsDirectory)) {
            [void]$moduleRoots.Add(
                (
                    Join-Path `
                        -Path $documentsDirectory `
                        -ChildPath 'WindowsPowerShell\Modules'
                )
            )
        }

        if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
            [void]$moduleRoots.Add(
                (
                    Join-Path `
                        -Path $env:USERPROFILE `
                        -ChildPath 'Documents\WindowsPowerShell\Modules'
                )
            )

            [void]$moduleRoots.Add(
                (
                    Join-Path `
                        -Path $env:USERPROFILE `
                        -ChildPath 'OneDrive\Documents\WindowsPowerShell\Modules'
                )
            )
        }

        if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
            [void]$moduleRoots.Add(
                (
                    Join-Path `
                        -Path $env:ProgramFiles `
                        -ChildPath 'WindowsPowerShell\Modules'
                )
            )
        }

        [void]$moduleRoots.Add($moduleCacheRoot)

        foreach ($moduleRoot in @($moduleRoots | Sort-Object -Unique)) {
            if ([string]::IsNullOrWhiteSpace($moduleRoot)) {
                continue
            }

            $versionedManifestPath = Join-Path `
                -Path $moduleRoot `
                -ChildPath (
                    '{0}\{1}\{0}.psd1' -f
                    $moduleName,
                    $requiredVersionText
                )

            $flatManifestPath = Join-Path `
                -Path $moduleRoot `
                -ChildPath (
                    '{0}\{0}.psd1' -f
                    $moduleName
                )

            [void]$candidatePaths.Add($versionedManifestPath)
            [void]$candidatePaths.Add($flatManifestPath)
        }

        foreach (
            $candidatePath in @(
                $candidatePaths |
                    Where-Object {
                        -not [string]::IsNullOrWhiteSpace($_)
                    } |
                    Sort-Object -Unique
            )
        ) {
            if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
                continue
            }

            try {
                $manifest = Test-ModuleManifest `
                    -Path $candidatePath `
                    -ErrorAction Stop

                if ([version]$manifest.Version -eq $RequiredVersion) {
                    return (
                        Resolve-Path `
                            -LiteralPath $candidatePath `
                            -ErrorAction Stop
                    ).Path
                }
            }
            catch {
                continue
            }
        }

        return $null
    }.GetNewClosure()

    $manifestPath = & $findExactModuleManifest
    $installFailure = $null

    if (-not $manifestPath) {
        Write-Host (
            'ExchangeOnlineManagement {0} was not found.' -f
            $requiredVersionText
        ) -ForegroundColor Yellow

        Write-Host 'Installing it for the current user...' `
            -ForegroundColor Yellow

        $nuGetProvider = Get-PackageProvider `
            -ListAvailable `
            -Name NuGet `
            -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $nuGetProvider) {
            Write-Host 'Installing the NuGet package provider...' `
                -ForegroundColor Yellow

            Install-PackageProvider `
                -Name NuGet `
                -Scope CurrentUser `
                -Force `
                -ErrorAction Stop |
                Out-Null
        }

        $installModuleCommand = Get-Command `
            -Name Install-Module `
            -ErrorAction Stop

        $installModuleParameters = @{
            Name            = $moduleName
            RequiredVersion = $requiredVersionText
            Repository      = 'PSGallery'
            Scope           = 'CurrentUser'
            Force           = $true
            ErrorAction     = 'Stop'
        }

        if (
            $installModuleCommand.Parameters.ContainsKey(
                'AllowClobber'
            )
        ) {
            $installModuleParameters.AllowClobber = $true
        }

        if (
            $installModuleCommand.Parameters.ContainsKey(
                'AcceptLicense'
            )
        ) {
            $installModuleParameters.AcceptLicense = $true
        }

        try {
            Install-Module @installModuleParameters
        }
        catch {
            $installFailure = (
                $_.Exception.Message -replace '[\r\n]+', ' '
            ).Trim()

            Write-Warning (
                'Install-Module did not complete successfully: {0}' -f
                $installFailure
            )
        }

        $userModuleRoot = if (
            -not [string]::IsNullOrWhiteSpace($documentsDirectory)
        ) {
            Join-Path `
                -Path $documentsDirectory `
                -ChildPath 'WindowsPowerShell\Modules'
        }
        else {
            $null
        }

        if (
            -not [string]::IsNullOrWhiteSpace($userModuleRoot) -and
            @($env:PSModulePath -split ';') -notcontains $userModuleRoot
        ) {
            $env:PSModulePath = (
                '{0};{1}' -f
                $userModuleRoot,
                $env:PSModulePath
            )
        }

        $manifestPath = & $findExactModuleManifest
    }

    if (-not $manifestPath) {
        Write-Warning (
            'The module was not visible after Install-Module. ' +
            'Downloading an isolated copy to the local module cache.'
        )

        if (-not (Test-Path -LiteralPath $moduleCacheRoot)) {
            New-Item `
                -Path $moduleCacheRoot `
                -ItemType Directory `
                -Force `
                -ErrorAction Stop |
                Out-Null
        }

        if (Test-Path -LiteralPath $exactCacheVersionPath) {
            Remove-Item `
                -LiteralPath $exactCacheVersionPath `
                -Recurse `
                -Force `
                -ErrorAction Stop
        }

        $saveModuleCommand = Get-Command `
            -Name Save-Module `
            -ErrorAction Stop

        $saveModuleParameters = @{
            Name            = $moduleName
            RequiredVersion = $requiredVersionText
            Repository      = 'PSGallery'
            Path            = $moduleCacheRoot
            Force           = $true
            ErrorAction     = 'Stop'
        }

        if (
            $saveModuleCommand.Parameters.ContainsKey(
                'AcceptLicense'
            )
        ) {
            $saveModuleParameters.AcceptLicense = $true
        }

        Write-Host (
            'Downloading ExchangeOnlineManagement {0} to: {1}' -f
            $requiredVersionText,
            $moduleCacheRoot
        ) -ForegroundColor Cyan

        Save-Module @saveModuleParameters

        $manifestPath = & $findExactModuleManifest
    }

    if (-not $manifestPath) {
        $details = if (
            -not [string]::IsNullOrWhiteSpace($installFailure)
        ) {
            " Install-Module error: $installFailure"
        }
        else {
            ''
        }

        throw (
            'ExchangeOnlineManagement {0} could not be located after ' +
            'Install-Module and Save-Module.{1}'
        ) -f $requiredVersionText, $details
    }

    Write-Host "Importing module from: $manifestPath" `
        -ForegroundColor DarkGray

    Import-Module `
        -Name $manifestPath `
        -Force `
        -ErrorAction Stop

    $loadedModule = Get-Module `
        -Name $moduleName |
        Where-Object {
            $_.Version -eq $RequiredVersion
        } |
        Select-Object -First 1

    if (-not $loadedModule) {
        throw (
            'ExchangeOnlineManagement {0} was found but could not be loaded.' -f
            $requiredVersionText
        )
    }

    Write-Host (
        'Loaded ExchangeOnlineManagement version {0}.' -f
        $loadedModule.Version
    ) -ForegroundColor DarkGray
}

function Connect-ToExchangeOnline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Configuration
    )

    $connectParameters = @{
        ShowBanner  = $false
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($Configuration.UserPrincipalName)) {
        $connectParameters.UserPrincipalName = $Configuration.UserPrincipalName
    }

    if (-not [string]::IsNullOrWhiteSpace($Configuration.DelegatedOrganization)) {
        $connectParameters.DelegatedOrganization = $Configuration.DelegatedOrganization
    }

    Write-Host 'Connecting to Exchange Online...' -ForegroundColor Cyan
    Connect-ExchangeOnline @connectParameters
    $script:ExchangeConnectionCreated = $true

    $connections = @(
        Get-ConnectionInformation -ErrorAction Stop |
            Where-Object { $_.TokenStatus -eq 'Active' -and -not $_.IsEopSession }
    )

    if ($connections.Count -eq 0) {
        throw 'Exchange Online authentication did not create an active connection.'
    }

    $connection = $connections | Select-Object -First 1
    $connectedAccount = Get-SafePropertyValue -InputObject $connection -PropertyNames @('UserPrincipalName', 'Account')
    $connectedTenant = Get-SafePropertyValue -InputObject $connection -PropertyNames @('TenantID', 'TenantId')

    if ($connectedAccount) {
        Write-Host "Connected account: $connectedAccount" -ForegroundColor Green
    }

    if ($connectedTenant) {
        Write-Host "Connected tenant : $connectedTenant" -ForegroundColor Green
    }

    return $connection
}

#endregion Module and connection

#region Exchange lookup helpers

function Get-AcceptedDomainForGroups {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$RequestedDomain
    )

    Write-Host 'Reading accepted domains...' -ForegroundColor Cyan

    $acceptedDomains = @(
        Get-AcceptedDomain `
            -ResultSize Unlimited `
            -ErrorAction Stop
    )

    if ($acceptedDomains.Count -eq 0) {
        throw 'Exchange Online did not return any accepted domains.'
    }

    $normalizedDomains = New-Object System.Collections.ArrayList

    foreach ($acceptedDomain in $acceptedDomains) {
        $domainNameValue = Get-SafePropertyValue `
            -InputObject $acceptedDomain `
            -PropertyNames @(
                'DomainName'
                'Name'
                'Identity'
            )

        if ([string]::IsNullOrWhiteSpace([string]$domainNameValue)) {
            continue
        }

        $domainName = ([string]$domainNameValue).Trim().ToLowerInvariant()

        $defaultValue = Get-SafePropertyValue `
            -InputObject $acceptedDomain `
            -PropertyNames @(
                'Default'
                'IsDefault'
            )

        $isDefault = $false

        if ($defaultValue -is [bool]) {
            $isDefault = [bool]$defaultValue
        }
        elseif (
            [string]::Equals(
                [string]$defaultValue,
                'True',
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $isDefault = $true
        }

        $domainType = Get-SafePropertyValue `
            -InputObject $acceptedDomain `
            -PropertyNames @(
                'DomainType'
                'Type'
            )

        [void]$normalizedDomains.Add(
            [pscustomobject][ordered]@{
                DomainName = $domainName
                IsDefault  = $isDefault
                DomainType = [string]$domainType
            }
        )
    }

    if ($normalizedDomains.Count -eq 0) {
        throw (
            'Accepted domain objects were returned, but no valid domain ' +
            'names could be read.'
        )
    }

    if (-not [string]::IsNullOrWhiteSpace($RequestedDomain)) {
        $normalizedRequestedDomain = $RequestedDomain.
            Trim().
            ToLowerInvariant()

        $requestedDomainMatch = @(
            $normalizedDomains |
                Where-Object {
                    [string]::Equals(
                        $_.DomainName,
                        $normalizedRequestedDomain,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        ) |
            Select-Object -First 1

        if (-not $requestedDomainMatch) {
            $availableDomainNames = @(
                $normalizedDomains.DomainName |
                    Sort-Object -Unique
            ) -join ', '

            throw (
                "'$RequestedDomain' is not an accepted domain in the " +
                "connected tenant. Available domains: $availableDomainNames"
            )
        }

        Write-Host (
            "Using the domain selected in the GUI: " +
            "'$($requestedDomainMatch.DomainName)'"
        ) -ForegroundColor Green

        return $requestedDomainMatch.DomainName
    }

    $selectedDomain = @(
        $normalizedDomains |
            Where-Object {
                $_.IsDefault
            }
    ) |
        Select-Object -First 1

    if ($selectedDomain) {
        Write-Host (
            "Using the default accepted domain: " +
            "'$($selectedDomain.DomainName)'"
        ) -ForegroundColor Green

        return $selectedDomain.DomainName
    }

    Write-Warning (
        'Exchange Online did not identify a default accepted domain. ' +
        'Searching for a custom domain.'
    )

    $selectedDomain = @(
        $normalizedDomains |
            Where-Object {
                $_.DomainName -notlike '*.onmicrosoft.com'
            } |
            Sort-Object DomainName
    ) |
        Select-Object -First 1

    if ($selectedDomain) {
        Write-Host (
            "Using the first custom accepted domain: " +
            "'$($selectedDomain.DomainName)'"
        ) -ForegroundColor Green

        return $selectedDomain.DomainName
    }

    Write-Warning (
        'No custom accepted domain was found. Using the first available domain.'
    )

    $selectedDomain = @(
        $normalizedDomains |
            Sort-Object DomainName
    ) |
        Select-Object -First 1

    if (-not $selectedDomain) {
        throw 'A usable accepted domain could not be selected.'
    }

    Write-Host (
        "Using accepted domain: '$($selectedDomain.DomainName)'"
    ) -ForegroundColor Green

    return $selectedDomain.DomainName
}

function Get-RecipientAddressKeys {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Recipient
    )

    if ($null -eq $Recipient) {
        return @()
    }

    $addresses = New-Object System.Collections.ArrayList

    foreach ($propertyName in @('PrimarySmtpAddress', 'WindowsEmailAddress', 'ExternalEmailAddress')) {
        $propertyValue = Get-SafePropertyValue -InputObject $Recipient -PropertyNames @($propertyName)

        if (-not [string]::IsNullOrWhiteSpace([string]$propertyValue)) {
            $addressText = [string]$propertyValue

            if ($addressText -match '^(?i:smtp:)?[^@\s]+@[^@\s]+$') {
                [void]$addresses.Add((Normalize-SmtpAddress -Address $addressText))
            }
        }
    }

    $emailAddresses = @(Get-SafePropertyValue -InputObject $Recipient -PropertyNames @('EmailAddresses'))

    foreach ($emailAddress in $emailAddresses) {
        $emailAddressText = [string]$emailAddress

        if ($emailAddressText.StartsWith('smtp:', [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$addresses.Add((Normalize-SmtpAddress -Address $emailAddressText))
        }
    }

    return @($addresses | Sort-Object -Unique)
}

function Add-RecipientToCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Recipient
    )

    foreach ($address in @(Get-RecipientAddressKeys -Recipient $Recipient)) {
        $script:RecipientCache[$address] = $Recipient
    }
}

function Get-RecipientBySmtpAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Address,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 10)]
        [int]$RetryCount = 0,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 30)]
        [int]$RetryDelaySeconds = 2
    )

    $normalizedAddress = Normalize-SmtpAddress -Address $Address

    if ($script:RecipientCache.ContainsKey($normalizedAddress)) {
        return $script:RecipientCache[$normalizedAddress]
    }

    $escapedAddress = Escape-OPathValue -Value $normalizedAddress

    for ($attempt = 0; $attempt -le $RetryCount; $attempt++) {
        $matches = @(
            Get-Recipient -Filter "EmailAddresses -eq '$escapedAddress'" -ResultSize 2 -ErrorAction Stop
        )

        if ($matches.Count -eq 0) {
            $matches = @(
                Get-Recipient -Filter "ExternalEmailAddress -eq '$escapedAddress'" -ResultSize 2 -ErrorAction Stop
            )
        }

        if ($matches.Count -gt 1) {
            throw "More than one recipient was returned for '$normalizedAddress'."
        }

        if ($matches.Count -eq 1) {
            Add-RecipientToCache -Recipient $matches[0]
            return $matches[0]
        }

        if ($attempt -lt $RetryCount) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    return $null
}

function Get-DistributionGroupBySmtpAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Address
    )

    $normalizedAddress = Normalize-SmtpAddress -Address $Address
    $escapedAddress = Escape-OPathValue -Value $normalizedAddress

    $groups = @(
        Get-DistributionGroup -Filter "EmailAddresses -eq '$escapedAddress'" -ResultSize 2 -ErrorAction Stop
    )

    if ($groups.Count -gt 1) {
        throw "More than one distribution group was returned for '$normalizedAddress'."
    }

    if ($groups.Count -eq 1) {
        return $groups[0]
    }

    return $null
}

function Get-UniqueRecipientAlias {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$PreferredAlias
    )

    $baseAlias = ConvertTo-ExchangeAlias -Value $PreferredAlias

    for ($index = 0; $index -lt 100; $index++) {
        $suffix = if ($index -eq 0) { '' } else { "-$($index + 1)" }
        $maximumBaseLength = 64 - $suffix.Length
        $candidateBase = $baseAlias

        if ($candidateBase.Length -gt $maximumBaseLength) {
            $candidateBase = $candidateBase.Substring(0, $maximumBaseLength)
        }

        $candidate = "$candidateBase$suffix"
        $escapedCandidate = Escape-OPathValue -Value $candidate
        $existingAlias = @(Get-Recipient -Filter "Alias -eq '$escapedCandidate'" -ResultSize 1 -ErrorAction Stop)

        if ($existingAlias.Count -eq 0) {
            return $candidate
        }
    }

    throw "A unique Exchange alias could not be generated from '$PreferredAlias'."
}

#endregion Exchange lookup helpers

#region Provisioning

function Ensure-BaseMailContacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Configuration
    )

    $enabledGroups = @($Configuration.Groups | Where-Object { $_.Enabled })

    $contactAddresses = @(
        foreach ($group in $enabledGroups) {
            foreach ($memberAddress in @($group.BaseMembers)) {
                Normalize-SmtpAddress -Address $memberAddress
            }
        }
    ) | Sort-Object -Unique

    $processedContactCount = 0

    foreach ($contactAddress in @($contactAddresses)) {
        $processedContactCount++

        Write-Progress `
            -Id 1 `
            -Activity 'Ensuring external mail contacts' `
            -Status (
                'Processing contact {0} of {1}: {2}' -f
                $processedContactCount,
                $contactAddresses.Count,
                $contactAddress
            ) `
            -PercentComplete (
                ($processedContactCount / [math]::Max($contactAddresses.Count, 1)) * 100
            )

        try {
            $existingRecipient = Get-RecipientBySmtpAddress -Address $contactAddress

            if ($existingRecipient) {
                $recipientType = [string](
                    Get-SafePropertyValue `
                        -InputObject $existingRecipient `
                        -PropertyNames @('RecipientTypeDetails', 'RecipientType')
                )

                Add-OperationResult `
                    -Phase 'Contacts' `
                    -ObjectType 'Recipient' `
                    -Name ([string]$existingRecipient.DisplayName) `
                    -Address $contactAddress `
                    -Action 'Ensure recipient' `
                    -Status 'Existing' `
                    -Message "Address already belongs to $recipientType."

                Write-OperationStatus `
                    -Status 'Existing' `
                    -Message "'$contactAddress' already exists as $recipientType."

                continue
            }

            $contactAlias = Get-UniqueRecipientAlias -PreferredAlias "ext-$contactAddress"

            if ($Configuration.PreviewOnly) {
                [void]$script:PlannedContactAddresses.Add($contactAddress)

                Add-OperationResult `
                    -Phase 'Contacts' `
                    -ObjectType 'MailContact' `
                    -Name $contactAddress `
                    -Address $contactAddress `
                    -Action 'Create mail contact' `
                    -Status 'Preview' `
                    -Message "Alias: $contactAlias"

                Write-OperationStatus `
                    -Status 'Preview' `
                    -Message "Would create mail contact '$contactAddress'."

                continue
            }

            $createdContact = New-MailContact `
                -Name $contactAlias `
                -Alias $contactAlias `
                -DisplayName $contactAddress `
                -ExternalEmailAddress $contactAddress `
                -ErrorAction Stop

            Add-RecipientToCache -Recipient $createdContact

            Add-OperationResult `
                -Phase 'Contacts' `
                -ObjectType 'MailContact' `
                -Name $contactAddress `
                -Address $contactAddress `
                -Action 'Create mail contact' `
                -Status 'Created' `
                -Message "Alias: $contactAlias"

            Write-OperationStatus `
                -Status 'Created' `
                -Message "Created mail contact '$contactAddress'."
        }
        catch {
            $errorMessage = ($_.Exception.Message -replace '[\r\n]+', ' ').Trim()

            Add-OperationResult `
                -Phase 'Contacts' `
                -ObjectType 'MailContact' `
                -Name $contactAddress `
                -Address $contactAddress `
                -Action 'Ensure mail contact' `
                -Status 'Failed' `
                -Message $errorMessage

            Write-OperationStatus `
                -Status 'Failed' `
                -Message "Could not process contact '$contactAddress': $errorMessage"
        }
    }

    Write-Progress -Id 1 -Activity 'Ensuring external mail contacts' -Completed
}

function Get-DesiredGroupMembers {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$GroupDefinition
    )

    return @(
        @($GroupDefinition.BaseMembers + $GroupDefinition.AdditionalMembers) |
            ForEach-Object { Normalize-SmtpAddress -Address $_ } |
            Sort-Object -Unique
    )
}

function Ensure-DistributionGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DomainName
    )

    $enabledGroups = @($Configuration.Groups | Where-Object { $_.Enabled })
    $requireAuthenticatedSenders = -not [bool]$Configuration.AllowExternalSenders
    $processedGroupCount = 0

    foreach ($groupDefinition in $enabledGroups) {
        $processedGroupCount++
        $groupDisplayName = $groupDefinition.DisplayName
        $groupAddress = '{0}@{1}' -f $groupDefinition.LocalPart, $DomainName

        Write-Progress `
            -Id 1 `
            -Activity 'Ensuring Exchange Online distribution groups' `
            -Status (
                'Processing group {0} of {1}: {2}' -f
                $processedGroupCount,
                $enabledGroups.Count,
                $groupDisplayName
            ) `
            -PercentComplete (
                ($processedGroupCount / [math]::Max($enabledGroups.Count, 1)) * 100
            )

        Write-Host ''
        Write-Host ('-' * 72) -ForegroundColor DarkGray
        Write-Host "Processing group: '$groupDisplayName' <$groupAddress>" -ForegroundColor Cyan

        $group = $null
        $groupPlanned = $false

        try {
            $group = Get-DistributionGroupBySmtpAddress -Address $groupAddress

            if (-not $group) {
                $addressConflict = Get-RecipientBySmtpAddress -Address $groupAddress

                if ($addressConflict) {
                    $recipientType = [string](
                        Get-SafePropertyValue `
                            -InputObject $addressConflict `
                            -PropertyNames @('RecipientTypeDetails', 'RecipientType')
                    )

                    $message = (
                        "Address '$groupAddress' already belongs to " +
                        "$recipientType '$($addressConflict.DisplayName)'."
                    )

                    Add-OperationResult `
                        -Phase 'Groups' `
                        -ObjectType 'DistributionGroup' `
                        -Name $groupDisplayName `
                        -Address $groupAddress `
                        -Action 'Create group' `
                        -Status 'Conflict' `
                        -Message $message

                    Write-OperationStatus -Status 'Conflict' -Message $message
                    continue
                }

                $groupAlias = Get-UniqueRecipientAlias -PreferredAlias $groupDefinition.LocalPart

                if ($Configuration.PreviewOnly) {
                    $groupPlanned = $true

                    Add-OperationResult `
                        -Phase 'Groups' `
                        -ObjectType 'DistributionGroup' `
                        -Name $groupDisplayName `
                        -Address $groupAddress `
                        -Action 'Create group' `
                        -Status 'Preview' `
                        -Message "Alias: $groupAlias"

                    Write-OperationStatus `
                        -Status 'Preview' `
                        -Message "Would create group '$groupDisplayName'."
                }
                else {
                    $group = New-DistributionGroup `
                        -Name $groupAlias `
                        -DisplayName $groupDisplayName `
                        -Alias $groupAlias `
                        -PrimarySmtpAddress $groupAddress `
                        -Type Distribution `
                        -MemberJoinRestriction Closed `
                        -MemberDepartRestriction Closed `
                        -RequireSenderAuthenticationEnabled:$requireAuthenticatedSenders `
                        -ErrorAction Stop

                    Add-RecipientToCache -Recipient $group

                    Add-OperationResult `
                        -Phase 'Groups' `
                        -ObjectType 'DistributionGroup' `
                        -Name $groupDisplayName `
                        -Address $groupAddress `
                        -Action 'Create group' `
                        -Status 'Created' `
                        -Message "Alias: $groupAlias"

                    Write-OperationStatus `
                        -Status 'Created' `
                        -Message "Created group '$groupDisplayName'."
                }
            }
            else {
                Add-OperationResult `
                    -Phase 'Groups' `
                    -ObjectType 'DistributionGroup' `
                    -Name ([string]$group.DisplayName) `
                    -Address $groupAddress `
                    -Action 'Ensure group' `
                    -Status 'Existing' `
                    -Message "Alias: $($group.Alias)"

                Write-OperationStatus `
                    -Status 'Existing' `
                    -Message "Group '$($group.DisplayName)' already exists."
            }

            if ($groupPlanned) {
                foreach ($memberAddress in @(Get-DesiredGroupMembers -GroupDefinition $groupDefinition)) {
                    Add-OperationResult `
                        -Phase 'Membership' `
                        -ObjectType 'DistributionGroupMember' `
                        -Name $groupDisplayName `
                        -Address $memberAddress `
                        -Action 'Add member' `
                        -Status 'Preview' `
                        -Message 'Would add after the group is created.'
                }

                continue
            }

            if (-not $group) {
                throw "Group '$groupAddress' could not be resolved."
            }

            $currentSenderAuthentication = [bool](
                Get-SafePropertyValue `
                    -InputObject $group `
                    -PropertyNames @('RequireSenderAuthenticationEnabled')
            )

            if ($currentSenderAuthentication -ne $requireAuthenticatedSenders) {
                $senderAction = if ($requireAuthenticatedSenders) {
                    'Require authenticated senders'
                }
                else {
                    'Allow unauthenticated external senders'
                }

                if ($Configuration.PreviewOnly) {
                    Add-OperationResult `
                        -Phase 'Groups' `
                        -ObjectType 'DistributionGroup' `
                        -Name $groupDisplayName `
                        -Address $groupAddress `
                        -Action 'Configure sender policy' `
                        -Status 'Preview' `
                        -Message $senderAction
                }
                else {
                    Set-DistributionGroup `
                        -Identity $group.Identity `
                        -RequireSenderAuthenticationEnabled:$requireAuthenticatedSenders `
                        -BypassSecurityGroupManagerCheck `
                        -ErrorAction Stop

                    Add-OperationResult `
                        -Phase 'Groups' `
                        -ObjectType 'DistributionGroup' `
                        -Name $groupDisplayName `
                        -Address $groupAddress `
                        -Action 'Configure sender policy' `
                        -Status 'Updated' `
                        -Message $senderAction

                    Write-OperationStatus `
                        -Status 'Updated' `
                        -Message "Updated sender policy for '$groupDisplayName'."
                }
            }

            $desiredMembers = @(Get-DesiredGroupMembers -GroupDefinition $groupDefinition)
            $existingMembers = @(
                Get-DistributionGroupMember `
                    -Identity $group.Identity `
                    -ResultSize Unlimited `
                    -ErrorAction Stop
            )

            $existingMemberAddresses = New-Object `
                'System.Collections.Generic.HashSet[string]' `
                ([System.StringComparer]::OrdinalIgnoreCase)

            foreach ($existingMember in $existingMembers) {
                foreach ($existingAddress in @(Get-RecipientAddressKeys -Recipient $existingMember)) {
                    [void]$existingMemberAddresses.Add($existingAddress)
                }
            }

            $processedMemberCount = 0

            foreach ($memberAddress in $desiredMembers) {
                $processedMemberCount++

                Write-Progress `
                    -Id 2 `
                    -ParentId 1 `
                    -Activity "Adding members to '$groupDisplayName'" `
                    -Status (
                        'Processing member {0} of {1}: {2}' -f
                        $processedMemberCount,
                        $desiredMembers.Count,
                        $memberAddress
                    ) `
                    -PercentComplete (
                        ($processedMemberCount / [math]::Max($desiredMembers.Count, 1)) * 100
                    )

                if ($existingMemberAddresses.Contains($memberAddress)) {
                    Add-OperationResult `
                        -Phase 'Membership' `
                        -ObjectType 'DistributionGroupMember' `
                        -Name $groupDisplayName `
                        -Address $memberAddress `
                        -Action 'Ensure membership' `
                        -Status 'Existing' `
                        -Message 'Member already belongs to the group.'

                    continue
                }

                $memberRecipient = Get-RecipientBySmtpAddress `
                    -Address $memberAddress `
                    -RetryCount 4 `
                    -RetryDelaySeconds 2

                if (-not $memberRecipient) {
                    if ($Configuration.PreviewOnly -and $script:PlannedContactAddresses.Contains($memberAddress)) {
                        Add-OperationResult `
                            -Phase 'Membership' `
                            -ObjectType 'DistributionGroupMember' `
                            -Name $groupDisplayName `
                            -Address $memberAddress `
                            -Action 'Add member' `
                            -Status 'Preview' `
                            -Message 'Would add after the planned contact is created.'

                        continue
                    }

                    $message = "'$memberAddress' does not exist as a mail-enabled recipient."

                    Add-OperationResult `
                        -Phase 'Membership' `
                        -ObjectType 'DistributionGroupMember' `
                        -Name $groupDisplayName `
                        -Address $memberAddress `
                        -Action 'Add member' `
                        -Status 'Failed' `
                        -Message $message

                    Write-OperationStatus -Status 'Failed' -Message $message
                    continue
                }

                $recipientAddressKeys = @(Get-RecipientAddressKeys -Recipient $memberRecipient)
                $alreadyMember = $false

                foreach ($recipientAddressKey in $recipientAddressKeys) {
                    if ($existingMemberAddresses.Contains($recipientAddressKey)) {
                        $alreadyMember = $true
                        break
                    }
                }

                if ($alreadyMember) {
                    Add-OperationResult `
                        -Phase 'Membership' `
                        -ObjectType 'DistributionGroupMember' `
                        -Name $groupDisplayName `
                        -Address $memberAddress `
                        -Action 'Ensure membership' `
                        -Status 'Existing' `
                        -Message 'The resolved recipient already belongs to the group.'

                    continue
                }

                $memberIdentity = Get-SafePropertyValue `
                    -InputObject $memberRecipient `
                    -PropertyNames @('ExternalDirectoryObjectId', 'PrimarySmtpAddress', 'Identity')

                if ([string]::IsNullOrWhiteSpace([string]$memberIdentity)) {
                    $memberIdentity = $memberAddress
                }

                if ($Configuration.PreviewOnly) {
                    Add-OperationResult `
                        -Phase 'Membership' `
                        -ObjectType 'DistributionGroupMember' `
                        -Name $groupDisplayName `
                        -Address $memberAddress `
                        -Action 'Add member' `
                        -Status 'Preview' `
                        -Message 'Would add member to group.'

                    continue
                }

                Add-DistributionGroupMember `
                    -Identity $group.Identity `
                    -Member $memberIdentity `
                    -BypassSecurityGroupManagerCheck `
                    -Confirm:$false `
                    -ErrorAction Stop

                foreach ($recipientAddressKey in $recipientAddressKeys) {
                    [void]$existingMemberAddresses.Add($recipientAddressKey)
                }

                Add-OperationResult `
                    -Phase 'Membership' `
                    -ObjectType 'DistributionGroupMember' `
                    -Name $groupDisplayName `
                    -Address $memberAddress `
                    -Action 'Add member' `
                    -Status 'Added' `
                    -Message 'Member added successfully.'

                Write-OperationStatus `
                    -Status 'Added' `
                    -Message "Added '$memberAddress' to '$groupDisplayName'."
            }

            Write-Progress -Id 2 -ParentId 1 -Activity "Adding members to '$groupDisplayName'" -Completed
        }
        catch {
            $errorMessage = ($_.Exception.Message -replace '[\r\n]+', ' ').Trim()

            Add-OperationResult `
                -Phase 'Groups' `
                -ObjectType 'DistributionGroup' `
                -Name $groupDisplayName `
                -Address $groupAddress `
                -Action 'Ensure group and membership' `
                -Status 'Failed' `
                -Message $errorMessage

            Write-OperationStatus `
                -Status 'Failed' `
                -Message "Could not process '$groupDisplayName': $errorMessage"
        }
    }

    Write-Progress -Id 1 -Activity 'Ensuring Exchange Online distribution groups' -Completed
}

#endregion Provisioning

#region Reporting

function Export-ExecutionReport {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Configuration
    )

    if (-not $Configuration.CreateReport) {
        return $null
    }

    $reportDirectory = $Configuration.ReportDirectory

    if (-not (Test-Path -LiteralPath $reportDirectory)) {
        New-Item -ItemType Directory -Path $reportDirectory -Force -ErrorAction Stop | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportPath = Join-Path -Path $reportDirectory -ChildPath "Exchange_Notification_Groups_$timestamp.csv"

    $script:Results |
        Select-Object Timestamp, Phase, ObjectType, Name, Address, Action, Status, Message |
        Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8 -Force

    return $reportPath
}

function Write-ExecutionSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ReportPath
    )

    Write-Host ''
    Write-Host ('-' * 72) -ForegroundColor DarkGray
    Write-Host 'Provisioning completed.' -ForegroundColor Green
    Write-Host "Operations recorded: $($script:Results.Count)"

    $statusGroups = @(
        $script:Results |
            Group-Object -Property Status |
            Sort-Object Count -Descending
    )

    foreach ($statusGroup in $statusGroups) {
        Write-Host ('  {0,-12} {1,6}' -f $statusGroup.Name, $statusGroup.Count)
    }

    if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
        Write-Host ''
        Write-Host "Execution report: $ReportPath"
    }

    $problemCount = @(
        $script:Results |
            Where-Object { $_.Status -in @('Failed', 'Conflict') }
    ).Count

    if ($problemCount -gt 0) {
        Write-Warning "$problemCount operation(s) require review. Check the execution report for details."
    }
}

#endregion Reporting

#region Main execution

try {
    $configuration = Show-ConfigurationDialog `
        -DefaultGroups $defaultGroupDefinitions `
        -DefaultModuleVersion $ExchangeModuleVersion

    Write-Host 'Starting Exchange Online notification group provisioning...' -ForegroundColor Cyan
    Write-Host "Process ID             : $PID" -ForegroundColor DarkGray
    Write-Host ('PowerShell version     : {0} ({1})' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition) -ForegroundColor DarkGray
    Write-Host "Exchange module version: $($configuration.ModuleVersion)" -ForegroundColor DarkGray

    $operationMode = if ($configuration.PreviewOnly) { 'Preview' } else { 'Apply changes' }
    Write-Host "Operation mode         : $operationMode" -ForegroundColor DarkGray

    Import-RequiredExchangeModule -RequiredVersion $configuration.ModuleVersion
    $null = Connect-ToExchangeOnline -Configuration $configuration

    $selectedDomain = Get-AcceptedDomainForGroups -RequestedDomain $configuration.PrimaryDomain
    Write-Host "Group SMTP domain      : $selectedDomain" -ForegroundColor Green

    $externalSenderStatus = if ($configuration.AllowExternalSenders) { 'Allowed' } else { 'Blocked' }
    Write-Host "External senders       : $externalSenderStatus" -ForegroundColor Green

    Ensure-BaseMailContacts -Configuration $configuration
    Ensure-DistributionGroups -Configuration $configuration -DomainName $selectedDomain

    $reportPath = Export-ExecutionReport -Configuration $configuration
    Write-ExecutionSummary -ReportPath $reportPath
}
catch [System.OperationCanceledException] {
    Write-Host 'Operation canceled. No changes were made.' -ForegroundColor Yellow
    exit 0
}
catch {
    $failureMessage = ($_.Exception.Message -replace '[\r\n]+', ' ').Trim()
    Write-Error "Provisioning failed: $failureMessage"
    exit 1
}
finally {
    if ($script:ExchangeConnectionCreated) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}

#endregion Main execution
