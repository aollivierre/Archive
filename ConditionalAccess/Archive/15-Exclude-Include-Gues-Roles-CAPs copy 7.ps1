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
        # Pre-allocate array for better performance
        $formattedPolicies = [System.Collections.Generic.List[PSObject]]::new($policies.Count)

        foreach ($policy in $policies) {
            Write-Verbose "Processing policy: $($policy.DisplayName)"
            
            # Extract version using regex match for efficiency
            $version = 1.0
            if ($policy.DisplayName -match '-v(\d+\.\d+)$') {
                $version = [decimal]$matches[1]
            }
            
            # Build guest status efficiently
            $guestStatus = [System.Text.StringBuilder]::new()
            if ($policy.Conditions.Users.IncludeGuestsOrExternalUsers) {
                $guestStatus.Append("Include: Guest/External Users")
            }
            if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers) {
                if ($guestStatus.Length -gt 0) {
                    $guestStatus.Append(" | ")
                }
                $guestStatus.Append("Exclude: Guest/External Users")
            }
            
            $currentGuestStatus = if ($guestStatus.Length -gt 0) {
                $guestStatus.ToString()
            }
            else {
                "No guest configuration"
            }
            
            # Add to list
            $formattedPolicies.Add([PSCustomObject]@{
                DisplayName = $policy.DisplayName
                Id = $policy.Id
                State = $policy.State
                Version = $version
                CurrentGuestStatus = $currentGuestStatus
            })
        }

        Write-Verbose "Returning $($formattedPolicies.Count) formatted policies"
        return $formattedPolicies
    }
    catch {
        Write-Error "Failed to retrieve or process Conditional Access Policies: $_"
        return @()
    }
}

function Export-GuestPolicyReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Results,
        
        [Parameter(Mandatory)]
        [string]$OutputDir,
        
        [Parameter()]
        [string]$ReportName = "GuestPolicy_Update"
    )
    
    # Use hashtable for better performance
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPaths = @{
        HTML = Join-Path $OutputDir "$ReportName`_$timestamp.html"
        CSV = Join-Path $OutputDir "$ReportName`_$timestamp.csv"
    }

    # Ensure output directory exists
    $null = New-Item -ItemType Directory -Force -Path $OutputDir

    # Export to CSV
    $Results | Export-Csv -Path $reportPaths.CSV -NoTypeInformation

    # Calculate counts once for efficiency
    $totalPolicies = $Results.Count
    $successCount = ($Results | Where-Object { $_.Status -eq 'Success' }).Count
    $failedCount = $totalPolicies - $successCount

    # Generate HTML report
    New-HTML -TitleText "Guest Policy Update Report" -FilePath $reportPaths.HTML -ShowHTML {
        New-HTMLSection -HeaderText "Policy Update Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Update Details</h3>
                <ul>
                    <li>Total Policies Updated: $totalPolicies</li>
                    <li>Successful Updates: $successCount</li>
                    <li>Failed Updates: $failedCount</li>
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
    Write-Host "CSV Report: $($reportPaths.CSV)" -ForegroundColor Green
    Write-Host "HTML Report: $($reportPaths.HTML)" -ForegroundColor Green

    return $reportPaths
}

function Get-GuestUserTypes {
    [CmdletBinding()]
    param()
    
    # Return pre-defined array for better performance
    return [PSCustomObject[]]@(
        @{
            DisplayName = "Internal guest users"
            Id = "internalGuest"
            Description = "Guest users within your organization"
            Category = "Internal"
        },
        @{
            DisplayName = "B2B collaboration guest users"
            Id = "b2bCollaborationGuest"
            Description = "Users invited to collaborate with your organization"
            Category = "B2B"
        },
        @{
            DisplayName = "B2B collaboration member users"
            Id = "b2bCollaborationMember"
            Description = "Members from other organizations"
            Category = "B2B"
        },
        @{
            DisplayName = "B2B direct connect users"
            Id = "b2bDirectConnectUser"
            Description = "Users connecting directly from partner organizations"
            Category = "B2B"
        },
        @{
            DisplayName = "Other external users"
            Id = "otherExternalUser"
            Description = "All other types of external users"
            Category = "External"
        },
        @{
            DisplayName = "Service provider users"
            Id = "serviceProvider"
            Description = "Users from service provider organizations"
            Category = "External"
        }
    ) | Sort-Object Category, DisplayName
}

function Update-ConditionalAccessPolicyGuestTypes {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$OutputPath = ".\Reports"
    )
    
    try {
        # Show welcome message
        Write-Host "`nConditional Access Policy Guest Types Update Tool" -ForegroundColor Cyan
        Write-Host "=================================================" -ForegroundColor Cyan
        
        # Get operation type
        $operation = Show-GuestOperationMenu
        if ($operation -eq 'Cancel') {
            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nSelected Operation: $operation" -ForegroundColor Green
        
        # Get current policies
        Write-Host "`nRetrieving current policy information..." -ForegroundColor Cyan
        $policies = Get-ConditionalAccessPoliciesDetails
        
        if (-not $policies) {
            Write-Warning "No policies found to process."
            return
        }
        
        # Select policies to update
        Write-Host "`nSelect policies to update..." -ForegroundColor Cyan
        $selectedPolicies = $policies | 
            Out-GridView -Title "Select Conditional Access Policies to Modify" -PassThru
        
        if (-not $selectedPolicies) {
            Write-Warning "No policies selected for processing."
            return
        }

        # Select guest types
        Write-Host "`nSelect guest types to $($operation.ToLower())..." -ForegroundColor Cyan
        $guestTypes = Get-GuestUserTypes
        
        $selectedGuestTypes = $guestTypes | 
            Out-GridView -Title "Select Guest Types to $operation" -PassThru
        
        if (-not $selectedGuestTypes) {
            Write-Warning "No guest types selected. Operation cancelled."
            return
        }

        # Enhanced operation summary
        Write-Host "`nOperation Summary:" -ForegroundColor Cyan
        Write-Host "==================" -ForegroundColor Cyan
        Write-Host "Operation: $operation guest/external users" -ForegroundColor White
        Write-Host "`nSelected Policies:" -ForegroundColor White
        $selectedPolicies | ForEach-Object { Write-Host "- $($_.DisplayName)" -ForegroundColor Gray }
        Write-Host "`nSelected Guest Types:" -ForegroundColor White
        $selectedGuestTypes | ForEach-Object { Write-Host "- $($_.DisplayName)" -ForegroundColor Gray }
        
        $confirm = Read-Host "`nDo you want to proceed? [Y/N]"
        if ($confirm -notmatch '^[Yy]$') {
            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
            return
        }
        
        # Pre-allocate results array for better performance
        $results = [System.Collections.Generic.List[PSObject]]::new($selectedPolicies.Count)
        
        # Create policy update template once
        $guestConfig = @{
            guestOrExternalUserTypes = $selectedGuestTypes.Id -join ','
            externalTenants = @{
                membershipKind = "all"
            }
        }

        foreach ($policy in $selectedPolicies) {
            try {
                Write-Host "`nProcessing policy: $($policy.DisplayName)" -ForegroundColor Cyan
                
                # Create minimal update body
                $bodyParams = @{
                    conditions = @{
                        users = @{
                            "$($operation.ToLower())GuestsOrExternalUsers" = $guestConfig
                        }
                    }
                }

                # Update policy using splatting for better performance
                $updateParams = @{
                    ConditionalAccessPolicyId = $policy.Id
                    BodyParameter = $bodyParams
                    ErrorAction = 'Stop'
                }
                $null = Update-MgIdentityConditionalAccessPolicy @updateParams
                
                # Add result
                $results.Add([PSCustomObject]@{
                    PolicyName = $policy.DisplayName
                    PolicyId = $policy.Id
                    Operation = $operation
                    GuestTypesUpdated = $selectedGuestTypes.DisplayName -join ', '
                    Status = 'Success'
                    ErrorMessage = ''
                })

                Write-Host "Successfully updated policy: $($policy.DisplayName)" -ForegroundColor Green
            }
            catch {
                $errorMessage = $_.Exception.Message
                
                $results.Add([PSCustomObject]@{
                    PolicyName = $policy.DisplayName
                    PolicyId = $policy.Id
                    Operation = $operation
                    GuestTypesUpdated = ''
                    Status = 'Failed'
                    ErrorMessage = $errorMessage
                })
                
                Write-Warning "Failed to update policy $($policy.DisplayName): $errorMessage"
            }
        }
        
        if ($results.Count -gt 0) {
            Export-GuestPolicyReport -Results $results -OutputDir $OutputPath
        }
        
        return $results
    }
    catch {
        Write-Error "Operation failed: $_"
    }
}

function Show-GuestOperationMenu {
    [CmdletBinding()]
    param()
    
    $menuText = @"

====== Conditional Access Policy Guest Operation ======

Select an operation:
[I] Include Selected Guest Types
[E] Exclude Selected Guest Types
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

# Main execution
Update-ConditionalAccessPolicyGuestTypes