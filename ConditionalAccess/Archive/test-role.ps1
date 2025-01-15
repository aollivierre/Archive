function Test-RoleRetrieval {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "Attempting to retrieve directory roles..." -ForegroundColor Yellow
        
        $roles = Get-MgBetaDirectoryRole -All
        Write-Host "Found $($roles.Count) roles" -ForegroundColor Green
        
        $firstRole = $roles | Select-Object -First 1
        Write-Host "Details of first role:" -ForegroundColor Yellow
        $firstRole | Format-List
        
        Write-Host "Attempting to get template details..." -ForegroundColor Yellow
        if ($firstRole.RoleTemplateId) {
            $template = Get-MgBetaDirectoryRoleTemplate -DirectoryRoleTemplateId $firstRole.RoleTemplateId
            Write-Host "Template details:" -ForegroundColor Yellow
            $template | Format-List
        }
        
    } catch {
        Write-Host "Error occurred: $_" -ForegroundColor Red
    }
}

Test-RoleRetrieval