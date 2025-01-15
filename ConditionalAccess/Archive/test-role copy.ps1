function Get-EntraAdminRolesDetailed {
    [CmdletBinding()]
    param()
    
    try {
        Write-Verbose "Retrieving all directory role templates..."
        $roleTemplates = Get-MgBetaDirectoryRoleTemplate -All
        
        $formattedRoles = foreach ($role in $roleTemplates) {
            # Convert GUID to proper format if needed
            $roleId = $role.Id
            if ($roleId -notmatch '^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$') {
                Write-Warning "Invalid role ID format: $roleId"
                continue
            }
            
            [PSCustomObject]@{
                DisplayName = $role.DisplayName
                Id = $roleId
                Description = $role.Description
                Category = switch -Regex ($role.DisplayName) {
                    'Admin|Administrator' { 'Administrative' }
                    'Reader' { 'Reader' }
                    'Owner' { 'Owner' }
                    'Operator' { 'Operator' }
                    default { 'Other' }
                }
                IsBuiltIn = $true
                TemplateId = $role.Id
            }
        }
        
        # Sort roles by category and display name
        $sortedRoles = $formattedRoles | Sort-Object Category, DisplayName

        Write-Verbose "Retrieved $($sortedRoles.Count) roles"
        return $sortedRoles
    }
    catch {
        Write-Error "Failed to retrieve admin roles: $_"
        return $null
    }
}

# Test function to verify role mapping with your CA policy
function Test-RoleMappings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PolicyId
    )

    try {
        # Get the CA policy
        $policy = Get-MgBetaIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyId
        
        # Get all roles
        $allRoles = Get-EntraAdminRolesDetailed
        
        # Get included roles from policy
        $includedRoles = $policy.Conditions.Users.IncludeRoles
        
        Write-Host "`nMapping Included Roles:" -ForegroundColor Yellow
        foreach ($roleId in $includedRoles) {
            $roleInfo = $allRoles | Where-Object { $_.Id -eq $roleId }
            if ($roleInfo) {
                Write-Host "Role ID: $roleId" -ForegroundColor Green
                Write-Host "Display Name: $($roleInfo.DisplayName)" -ForegroundColor Green
                Write-Host "Description: $($roleInfo.Description)" -ForegroundColor Green
                Write-Host "-----------------"
            }
            else {
                Write-Host "Role ID: $roleId - No mapping found!" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Error "Failed to test role mappings: $_"
    }
}

# Example usage:
# First test role retrieval
$roles = Get-EntraAdminRolesDetailed -Verbose
Write-Host "Total roles found: $($roles.Count)"
$roles | Select-Object DisplayName, Category, Id | Format-Table -AutoSize

# Then test mapping with your CA policy
Test-RoleMappings -PolicyId "e1e3962e-1286-422c-b503-9a8de2ee8202"