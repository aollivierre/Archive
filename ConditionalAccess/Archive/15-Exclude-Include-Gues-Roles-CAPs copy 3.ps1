#Requires -Modules Microsoft.Graph.Identity.SignIns, PSWriteHTML




function Get-ConditionalAccessPoliciesDetails {
    [CmdletBinding()]
    param()
    
    try {
        Write-Verbose "Retrieving policies from Microsoft Graph..."
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        
        if (-not $policies) {
            Write-Warning "No policies retrieved from Microsoft Graph"
            return @()
        }

        Write-Verbose "Processing $($policies.Count) policies..."
        $formattedPolicies = @(foreach ($policy in $policies) {
                Write-Verbose "Processing policy: $($policy.DisplayName)"
            
                # Extract current version from display name
                $version = if ($policy.DisplayName -match '-v(\d+\.\d+)$') {
                    [decimal]$matches[1]
                }
                else {
                    1.0
                }
            
                # Build the guest status string
                $guestStatus = @()
            
                if ($policy.Conditions.Users.IncludeGuestsOrExternalUsers) {
                    $guestStatus += "Include: Guest/External Users"
                }
            
                if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers) {
                    $guestStatus += "Exclude: Guest/External Users"
                }
            
                $currentGuestStatus = if ($guestStatus.Count -gt 0) {
                    $guestStatus -join ' | '
                }
                else {
                    "No guest configuration"
                }
            
                # Create and output the object
                [PSCustomObject]@{
                    DisplayName        = $policy.DisplayName
                    Id                 = $policy.Id
                    State              = $policy.State
                    Version            = $version
                    CurrentGuestStatus = $currentGuestStatus
                }
            })

        Write-Verbose "Returning $($formattedPolicies.Count) formatted policies"
        return $formattedPolicies
    }
    catch {
        Write-Error "Failed to retrieve or process Conditional Access Policies: $_"
        return @()
    }
}



function Get-GuestUserTypes {
    [CmdletBinding()]
    param()
    
    # Define guest types directly as PSCustomObjects to avoid empty rows
    $guestTypes = @(
        [PSCustomObject]@{
            DisplayName = "B2B collaboration guest users"
            Id          = "b2bCollaborationGuest"
            Description = "Users who are invited to collaborate with your organization"
            Category    = "B2B"
        }
        [PSCustomObject]@{
            DisplayName = "B2B collaboration member users"
            Id          = "b2bCollaborationMember"
            Description = "Members from other organizations collaborating with yours"
            Category    = "B2B"
        }
        [PSCustomObject]@{
            DisplayName = "B2B direct connect users"
            Id          = "b2bDirectConnectUser"
            Description = "Users connecting directly from partner organizations"
            Category    = "B2B"
        }
        [PSCustomObject]@{
            DisplayName = "Local guest users"
            Id          = "localGuest"
            Description = "Guest users created within your organization"
            Category    = "Local"
        }
        [PSCustomObject]@{
            DisplayName = "Service provider users"
            Id          = "serviceProvider"
            Description = "Users from service provider organizations"
            Category    = "External"
        }
        [PSCustomObject]@{
            DisplayName = "Other external users"
            Id          = "otherExternalUser"
            Description = "All other types of external users"
            Category    = "External"
        }
    )
    
    # Display available guest types in console
    Write-Host "`nAvailable Guest Types:" -ForegroundColor Cyan
    $guestTypes | Format-Table DisplayName, Category
    
    return $guestTypes
}


function Export-GuestPolicyReport {
    param(
        [Parameter(Mandatory)]
        [object[]]$Results,
        
        [Parameter(Mandatory)]
        [string]$OutputDir,
        
        [Parameter()]
        [string]$ReportName = "GuestPolicy_Update"
    )
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path -Path $OutputDir)) {
        $null = New-Item -ItemType Directory -Path $OutputDir -Force
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $OutputDir "$($ReportName)_$timestamp.html"
    $csvPath = Join-Path $OutputDir "$($ReportName)_$timestamp.csv"
    
    # Export to CSV
    $Results | Export-Csv -Path $csvPath -NoTypeInformation
    
    New-HTML -TitleText "Guest Policy Update Report" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Policy Update Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Update Details</h3>
                <ul>
                    <li>Total Policies Updated: $($Results.Count)</li>
                    <li>Operation Performed: $($Results[0].Operation)</li>
                    <li>Generated On: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</li>
                </ul>
"@
            }
        }
        
        New-HTMLSection -HeaderText "Updated Policies" {
            New-HTMLTable -DataTable $Results -ScrollX {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Success' -BackgroundColor LightGreen -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor Salmon -Color Black
            } -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5', 'searchBuilder')
        }
    }

    Write-Host "`nReports generated:" -ForegroundColor Green
    Write-Host "CSV Report: $csvPath" -ForegroundColor Green
    Write-Host "HTML Report: $htmlPath" -ForegroundColor Green

    return @{
        CSVPath  = $csvPath
        HTMLPath = $htmlPath
    }
}

function Show-GuestOperationMenu {
    [CmdletBinding()]
    param()
    
    $menuText = @"

====== Conditional Access Policy Guest Operation ======

Select an operation:
[I] Include Guest Types
[E] Exclude Guest Types
[C] Cancel Operation

Enter your choice [I/E/C]: 
"@
    
    do {
        $choice = Read-Host -Prompt $menuText
        switch ($choice.ToUpper()) {
            'I' { return 'Include' }
            'E' { return 'Exclude' }
            'C' { return 'Cancel' }
            default {
                Write-Host "Invalid selection. Please try again." -ForegroundColor Yellow
            }
        }
    } while ($true)
}


function Update-ConditionalAccessPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Policy,
        
        [Parameter(Mandatory)]
        [ValidateSet('Include', 'Exclude')]
        [string]$Operation,
        
        [Parameter(Mandatory)]
        [object[]]$GuestTypes
    )
    
    try {
        Write-Verbose "Getting current policy details for $($Policy.DisplayName)..."
        $currentPolicy = Get-MgBetaIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $Policy.Id
        
        # Increment version
        $newVersion = $Policy.Version + 0.1
        $newDisplayName = if ($Policy.DisplayName -match '-v\d+\.\d+$') {
            $Policy.DisplayName -replace '-v\d+\.\d+$', "-v$newVersion"
        }
        else {
            "$($Policy.DisplayName)-v$newVersion"
        }

        # Start with the existing conditions
        $conditions = $currentPolicy.Conditions

        # Initialize users conditions if not present
        if (-not $conditions.users) {
            $conditions.users = @{
                excludeUsers  = @()
                includeUsers  = @('All')
                excludeGroups = @()
                includeGroups = @()
            }
        }

        # Handle guest configuration
        if ($Operation -eq 'Include') {
            $conditions.users['includeGuestsOrExternalUsers'] = @{
                GuestOrExternalUserTypes = @($GuestTypes.Id)
                ExternalTenants          = @{
                    MembershipKind = "all"
                }
            }
            $conditions.users.Remove('excludeGuestsOrExternalUsers')
        }
        else {
            $conditions.users['excludeGuestsOrExternalUsers'] = @{
                GuestOrExternalUserTypes = @($GuestTypes.Id)
                ExternalTenants          = @{
                    MembershipKind = "all"
                }
            }
            $conditions.users.Remove('includeGuestsOrExternalUsers')
        }

        # Build update parameters
        $bodyParams = @{
            displayName = $newDisplayName
            state       = $currentPolicy.State
            conditions  = $conditions
        }

        # Copy controls if they exist
        if ($currentPolicy.GrantControls) {
            $bodyParams['grantControls'] = $currentPolicy.GrantControls
        }
        if ($currentPolicy.SessionControls) {
            $bodyParams['sessionControls'] = $currentPolicy.SessionControls
        }

        Write-Verbose "Update payload:`n$($bodyParams | ConvertTo-Json -Depth 10)"

        # Update policy with splatting
        $updateParams = @{
            ConditionalAccessPolicyId = $Policy.Id
            BodyParameter             = $bodyParams
        }
        
        $null = Update-MgBetaIdentityConditionalAccessPolicy @updateParams
        Write-Verbose "Successfully updated policy $($Policy.DisplayName)"
        return $true
    }
    catch {
        Write-Warning "Failed to update policy $($Policy.DisplayName): $_"
        $detailedError = if ($_.ErrorDetails.Message) {
            $_.ErrorDetails.Message
        }
        else {
            $_.Exception.Message
        }
        Write-Verbose "Detailed error: $detailedError"
        return $false
    }
}

function Update-ConditionalAccessPolicyGuestTypes {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$OutputPath = ".\Reports"
    )
    
    try {
        # Show title
        Write-Host @"
Conditional Access Policy Guest Types Update Tool
=================================================
"@ -ForegroundColor Cyan
        
        # Get operation choice
        $operation = Show-GuestOperationMenu
        if ($operation -eq 'Cancel') {
            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nSelected Operation: $operation" -ForegroundColor Green

        # Get and validate policies
        Write-Host "`nRetrieving current policy information..." -ForegroundColor Cyan
        $policies = Get-ConditionalAccessPoliciesDetails -Verbose
        if (-not $policies) {
            Write-Warning "No policies found to process."
            return
        }

        # Select policies using Out-GridView
        $selectedPolicies = $policies | 
        Select-Object DisplayName, Id, State, Version, CurrentGuestStatus | 
        Out-GridView -Title "Select Conditional Access Policies to Modify" -PassThru
        
        if (-not $selectedPolicies) {
            Write-Warning "No policies selected for processing."
            return
        }

        # Get and select guest types
        Write-Host "`nRetrieving guest types..." -ForegroundColor Cyan
        $guestTypes = Get-GuestUserTypes
        
        $selectedGuestTypes = $guestTypes | 
        Select-Object DisplayName, Id, Category | 
        Out-GridView -Title "Select Guest Types to $Operation" -PassThru
        
        if (-not $selectedGuestTypes) {
            Write-Warning "No guest types selected for processing."
            return
        }

        # Show operation summary
        Write-Host @"
`nOperation Summary:
==================
Operation: $operation guest types

Selected Policies:
$(($selectedPolicies | ForEach-Object { "- $($_.DisplayName)" }) -join "`n")

Selected Guest Types:
$(($selectedGuestTypes | ForEach-Object { "- $($_.DisplayName)" }) -join "`n")
"@ -ForegroundColor White

        # Confirm operation
        $confirm = Read-Host "`nDo you want to proceed? [Y/N]"
        if ($confirm -notmatch '^[Yy]$') {
            Write-Host "`nOperation cancelled." -ForegroundColor Yellow
            return
        }

        # Process policies and collect results
        $results = [System.Collections.ArrayList]::new()
        
        foreach ($policy in $selectedPolicies) {
            try {
                Write-Verbose "Processing policy: $($policy.DisplayName)"
                $success = Update-ConditionalAccessPolicy -Policy $policy -Operation $operation -GuestTypes $selectedGuestTypes -Verbose
                
                $resultParams = @{
                    PolicyName        = $policy.DisplayName
                    PolicyId          = $policy.Id
                    Operation         = $operation
                    GuestTypesUpdated = if ($success) { ($selectedGuestTypes.DisplayName -join ', ') } else { '' }
                    TypeCount         = if ($success) { $selectedGuestTypes.Count } else { 0 }
                    Status            = if ($success) { 'Success' } else { 'Failed' }
                    ErrorMessage      = ''
                }
                
                $null = $results.Add([PSCustomObject]$resultParams)
            }
            catch {
                $errorParams = @{
                    PolicyName        = $policy.DisplayName
                    PolicyId          = $policy.Id
                    Operation         = $operation
                    GuestTypesUpdated = ''
                    TypeCount         = 0
                    Status            = 'Failed'
                    ErrorMessage      = $_.Exception.Message
                }
                
                $null = $results.Add([PSCustomObject]$errorParams)
            }
        }

        # Generate and return report
        Export-GuestPolicyReport -Results $results -OutputDir $OutputPath
        return $results
    }
    catch {
        Write-Error "Operation failed: $_"
    }
}

# Main execution
Update-ConditionalAccessPolicyGuestTypes