# Required modules
#Requires -Modules Microsoft.Graph.Identity.SignIns

function Connect-GraphWithScope {
    [CmdletBinding()]
    param()
    
    $requiredScopes = @(
        "AuditLog.Read.All"
        "Directory.Read.All"
        "Policy.Read.All"
    )
    
    $currentContext = Get-MgContext
    
    if ($null -eq $currentContext) {
        Write-Host "No existing Graph connection found. Connecting..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes $requiredScopes
        return
    }
    
    $missingScopes = $requiredScopes | Where-Object { $_ -notin $currentContext.Scopes }
    
    if ($missingScopes) {
        Write-Host "Missing required scopes. Reconnecting with all required scopes..." -ForegroundColor Yellow
        Disconnect-MgGraph
        Connect-MgGraph -Scopes $requiredScopes
    }
}

function Debug-RecentSignIns {
    [CmdletBinding()]
    param (
        [Parameter()]
        [int]$NumberOfSignIns = 3
    )

    # Ensure proper Graph authentication
    Connect-GraphWithScope

    Write-Host "Fetching the most recent $NumberOfSignIns sign-ins for detailed analysis..." -ForegroundColor Yellow
    
    try {
        # Get most recent sign-ins
        $signIns = Get-MgAuditLogSignIn -Top $NumberOfSignIns -Sort "createdDateTime desc"
        
        foreach ($signIn in $signIns) {
            Write-Host "`n`n========== Sign-in Details ==========" -ForegroundColor Cyan
            Write-Host "Time: $($signIn.CreatedDateTime)" -ForegroundColor Yellow
            Write-Host "User: $($signIn.UserPrincipalName)" -ForegroundColor Yellow
            Write-Host "App: $($signIn.AppDisplayName)" -ForegroundColor Yellow
            Write-Host "Status: $($signIn.Status.ErrorCode) - $($signIn.Status.FailureReason)" -ForegroundColor Yellow
            Write-Host "Conditional Access Status: $($signIn.ConditionalAccessStatus)" -ForegroundColor Yellow
            
            Write-Host "`nApplied Conditional Access Policies:" -ForegroundColor Green
            if ($null -ne $signIn.AppliedConditionalAccessPolicies -and $signIn.AppliedConditionalAccessPolicies.Count -gt 0) {
                foreach ($policy in $signIn.AppliedConditionalAccessPolicies) {
                    Write-Host "`nPolicy Id: $($policy)" -ForegroundColor Magenta
                    # Get all properties of the policy object
                    $policy | Get-Member -MemberType Properties | ForEach-Object {
                        $propertyName = $_.Name
                        $propertyValue = $policy.$propertyName
                        Write-Host "$propertyName : $propertyValue" -ForegroundColor Gray
                    }
                }
                
                # Show raw policy data
                Write-Host "`nRaw Policy Data:" -ForegroundColor Cyan
                $signIn.AppliedConditionalAccessPolicies | ConvertTo-Json -Depth 10 | Write-Host
            }
            else {
                Write-Host "No Applied Conditional Access Policies found for this sign-in!" -ForegroundColor Red
            }
            
            Write-Host "`nAuthentication Details:" -ForegroundColor Green
            Write-Host "Client App Used: $($signIn.ClientAppUsed)" -ForegroundColor Gray
            Write-Host "Is Interactive: $($signIn.IsInteractive)" -ForegroundColor Gray
            Write-Host "Resource: $($signIn.ResourceDisplayName)" -ForegroundColor Gray
            Write-Host "IP Address: $($signIn.IPAddress)" -ForegroundColor Gray
            if ($signIn.Location) {
                Write-Host "Location: $($signIn.Location.CountryOrRegion), $($signIn.Location.City)" -ForegroundColor Gray
            }

            Write-Host "`nAuthentication Processing Details:" -ForegroundColor Green
            if ($null -ne $signIn.AuthenticationProcessingDetails -and $signIn.AuthenticationProcessingDetails.Count -gt 0) {
                $signIn.AuthenticationProcessingDetails | Format-Table Key, Value -AutoSize
            }
            else {
                Write-Host "No Authentication Processing Details found!" -ForegroundColor Red
            }
        }
        
        # Calculate summary statistics
        $policiesFound = 0
        $reportOnlyPolicies = 0
        foreach ($signIn in $signIns) {
            if ($signIn.AppliedConditionalAccessPolicies.Count -gt 0) {
                $policiesFound++
                $reportOnlyPolicies += ($signIn.AppliedConditionalAccessPolicies | 
                    Where-Object { $_.EnforcementMode -eq "reportOnly" }).Count
            }
        }
        
        # Provide summary
        Write-Host "`n`n========== Analysis Summary ==========" -ForegroundColor Cyan
        Write-Host "Total sign-ins analyzed: $($signIns.Count)" -ForegroundColor Yellow
        Write-Host "Sign-ins with CA policies: $policiesFound" -ForegroundColor Yellow
        Write-Host "Total Report-only policies found: $reportOnlyPolicies" -ForegroundColor Yellow
        
        # Get list of unique policy IDs
        Write-Host "`nUnique Policy IDs found:" -ForegroundColor Cyan
        $uniquePolicies = $signIns | 
            Select-Object -ExpandProperty AppliedConditionalAccessPolicies | 
            Select-Object -Unique Id, DisplayName
        $uniquePolicies | Format-Table Id, DisplayName -AutoSize
        
    }
    catch {
        Write-Error "Error analyzing sign-in logs: $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

# Execute the analysis
Debug-RecentSignIns -NumberOfSignIns 3