#Requires -Modules Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Identity.SignIns, PSWriteHTML


function Get-ConditionalAccessPoliciesDetails {
    [CmdletBinding()]
    param()
    
    try {
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        
        $formattedPolicies = $policies | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                Id = $_.Id
                State = $_.State
                CurrentAdminRoles = if ($_.Conditions.Users.IncludeRoles) {
                    "Include: $($_.Conditions.Users.IncludeRoles.Count) roles"
                } elseif ($_.Conditions.Users.ExcludeRoles) {
                    "Exclude: $($_.Conditions.Users.ExcludeRoles.Count) roles"
                } else {
                    "No admin roles configured"
                }
            }
        }
        
        return $formattedPolicies
    }
    catch {
        Write-Error "Failed to retrieve Conditional Access Policies: $_"
        return $null
    }
}


function Get-EntraAdminRolesDetailed {
    [CmdletBinding()]
    param()
    
    try {
        # Using the correct cmdlet to get directory roles
        $roleDefinitions = Get-MgBetaDirectoryRole -All | ForEach-Object {
            # Get the full role definition for each role
            Get-MgBetaDirectoryRoleTemplate -DirectoryRoleTemplateId $_.RoleTemplateId
        }
        
        $formattedRoles = foreach ($role in $roleDefinitions) {
            [PSCustomObject]@{
                DisplayName  = $_.DisplayName
                Id          = $_.Id
                Description = $_.Description
                Category    = switch -Regex ($_.DisplayName) {
                    'Admin'  { 'Administrative' }
                    'Reader' { 'Reader' }
                    default  { 'Other' }
                }
                IsEnabled   = $true
                TemplateId = $_.TemplateId
            }
        }
        
        return $formattedRoles | Sort-Object Category, DisplayName
    }
    catch {
        Write-Error "Failed to retrieve admin roles: $_"
        return $null
    }
}

function Export-AdminRolesReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Results,
        
        [Parameter(Mandatory)]
        [string]$OutputDir,
        
        [Parameter()]
        [string]$ReportName = "AdminRoles_CAPolicy_Update"
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
    
    New-HTML -TitleText "Admin Roles Policy Update Report" -FilePath $htmlPath -ShowHTML {
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
        CSVPath = $csvPath
        HTMLPath = $htmlPath
    }
}

function Update-ConditionalAccessPolicyAdminRoles {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('Include', 'Exclude')]
        [string]$Operation = 'Include',
        
        [Parameter()]
        [string]$OutputPath = ".\Reports"
    )
    
    try {
        # Get and display policies for selection
        $policies = Get-ConditionalAccessPoliciesDetails
        if (-not $policies) { 
            Write-Warning "No policies found or error retrieving policies."
            return 
        }
        
        $selectedPolicies = $policies | 
            Out-GridView -Title "Select Conditional Access Policies to Modify" -PassThru
        
        if (-not $selectedPolicies) {
            Write-Warning "No policies selected. Operation cancelled."
            return
        }
        
        # Get admin roles
        $adminRoles = Get-EntraAdminRolesDetailed
        if (-not $adminRoles) { 
            Write-Warning "No admin roles found or error retrieving roles."
            return 
        }
        
        $selectedRoles = $adminRoles | 
            Out-GridView -Title "Select Admin Roles to $Operation" -PassThru
        
        if (-not $selectedRoles) {
            Write-Warning "No roles selected. Operation cancelled."
            return
        }
        
        # Process updates
        $results = [System.Collections.ArrayList]::new()
        
        foreach ($policy in $selectedPolicies) {
            try {
                $policyDetail = Get-MgBetaIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id
                
                # Prepare the update based on operation type
                $updateParams = @{
                    ConditionalAccessPolicyId = $policy.Id
                    BodyParameter = @{
                        conditions = @{
                            users = @{
                                includeRoles = if ($Operation -eq 'Include') { 
                                    $selectedRoles.Id 
                                } else { 
                                    $policyDetail.Conditions.Users.IncludeRoles
                                }
                                excludeRoles = if ($Operation -eq 'Exclude') { 
                                    $selectedRoles.Id 
                                } else { 
                                    $policyDetail.Conditions.Users.ExcludeRoles
                                }
                            }
                        }
                    }
                }
                
                # Update the policy
                Update-MgBetaIdentityConditionalAccessPolicy @updateParams
                
                $null = $results.Add([PSCustomObject]@{
                    PolicyName = $policy.DisplayName
                    PolicyId = $policy.Id
                    Operation = $Operation
                    RolesUpdated = $selectedRoles.Count
                    Status = 'Success'
                    ErrorMessage = ''
                })
            }
            catch {
                $null = $results.Add([PSCustomObject]@{
                    PolicyName = $policy.DisplayName
                    PolicyId = $policy.Id
                    Operation = $Operation
                    RolesUpdated = 0
                    Status = 'Failed'
                    ErrorMessage = $_.Exception.Message
                })
            }
        }
        
        # Generate report
        if ($results.Count -gt 0) {
            Export-AdminRolesReport -Results $results -OutputDir $OutputPath
        }
        
        return $results
    }
    catch {
        Write-Error "Operation failed: $_"
    }
}


Update-ConditionalAccessPolicyAdminRoles