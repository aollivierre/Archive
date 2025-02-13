# Required modules
#Requires -Modules Microsoft.Graph.Identity.SignIns, PSWriteHTML

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





# Modified Export function with fixed table parameters

function Get-CAReportOnlyAnalysis {
    [CmdletBinding()]
    param (
        [Parameter()]
        [int]$Hours = 7,
        
        [Parameter()]
        [switch]$UseDays,
        
        [Parameter()]
        [string]$ReportPath = ".\CA_Analysis",

        [Parameter()]
        [switch]$IncludeDetailedReport
    )

    # Ensure output directory exists
    $null = New-Item -ItemType Directory -Force -Path $ReportPath

    # Ensure proper Graph authentication
    Connect-GraphWithScope

    # Calculate time span
    $timespan = if ($UseDays) {
        $Hours * 24
        $displayUnit = "days"
        $displayValue = $Hours
    } else {
        $Hours
        $displayUnit = "hours"
        $displayValue = $Hours
    }

    Write-Host "Fetching sign-in logs for the past $displayValue $displayUnit..." -ForegroundColor Yellow
    
    try {
        # Initialize results array using ArrayList for better performance
        $results = [System.Collections.ArrayList]::new()
        
        # Calculate date range
        $startDate = (Get-Date).AddHours(-$timespan)
        
        # Get sign-in logs with minimal console output
        $signIns = Get-MgAuditLogSignIn -Filter "createdDateTime ge $($startDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" -All
        
        $totalSignIns = 0
        $signInsCount = $signIns.Count

        foreach ($signIn in $signIns) {
            $totalSignIns++
            
            # Update progress every 100 records
            if ($totalSignIns % 100 -eq 0) {
                Write-Progress -Activity "Processing sign-in logs" -Status "Processed $totalSignIns / $signInsCount sign-ins" -PercentComplete (($totalSignIns / $signInsCount) * 100)
            }
            
            # Get policies with report-only results
            $reportOnlyPolicies = $signIn.AppliedConditionalAccessPolicies | 
                Where-Object { $_.Result -like 'reportOnly*' }
            
            if ($reportOnlyPolicies) {
                foreach ($policy in $reportOnlyPolicies) {
                    $resultObject = [PSCustomObject]@{
                        SignInId = $signIn.Id
                        Time = $signIn.CreatedDateTime
                        UserPrincipalName = $signIn.UserPrincipalName
                        UserDisplayName = $signIn.UserDisplayName
                        AppDisplayName = $signIn.AppDisplayName
                        ClientApp = $signIn.ClientAppUsed
                        DevicePlatform = $signIn.DevicePlatform
                        Location = if ($signIn.Location.City) { 
                            "$($signIn.Location.City), $($signIn.Location.CountryOrRegion)" 
                        } else { 
                            $signIn.Location.CountryOrRegion 
                        }
                        RiskLevel = $signIn.RiskLevelAggregated
                        PolicyId = $policy.Id
                        PolicyName = $policy.DisplayName
                        Result = $policy.Result
                        GrantControls = ($policy.EnforcedGrantControls -join '; ')
                        SessionControls = ($policy.EnforcedSessionControls -join '; ')
                    }
                    
                    $null = $results.Add($resultObject)
                }
            }
        }

        Write-Progress -Activity "Processing sign-in logs" -Completed
        Write-Host "Analysis complete! Found $($results.Count) report-only policy evaluations" -ForegroundColor Green

        if ($results.Count -gt 0) {
            # Generate reports
            $reportPaths = Export-CAReportOnlyAnalysis -Results $results -OutputDir $ReportPath -IncludeDetailedReport:$IncludeDetailedReport
            return $reportPaths | Out-Null
        }
        else {
            Write-Host "No report-only policy evaluations found in the specified time period." -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        Write-Error "Error analyzing sign-in logs: $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

function Export-CAReportOnlyAnalysis {
    param(
        [Parameter(Mandatory)]
        $Results,
        
        [Parameter(Mandatory)]
        [string]$OutputDir,
        
        [Parameter()]
        [string]$ReportName = "CA_ReportOnly_Analysis",

        [Parameter()]
        [switch]$IncludeDetailedReport
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $OutputDir "$($ReportName)_$timestamp.html"
    $detailedHtmlPath = Join-Path $OutputDir "$($ReportName)_Detailed_$timestamp.html"
    $csvPath = Join-Path $OutputDir "$($ReportName)_$timestamp.csv"
    
    # Export to CSV
    $Results | Export-Csv -Path $csvPath -NoTypeInformation -Force

    # Calculate summary statistics
    $metadata = @{
        GeneratedBy = $env:USERNAME
        GeneratedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalSignIns = ($Results | Select-Object -Unique SignInId).Count
        UniqueUsers = ($Results | Select-Object -Unique UserPrincipalName).Count
        UniqueApps = ($Results | Select-Object -Unique AppDisplayName).Count
        FailureCount = ($Results | Where-Object Result -eq 'reportOnlyFailure').Count
        SuccessCount = ($Results | Where-Object Result -eq 'reportOnlySuccess').Count
        InterruptedCount = ($Results | Where-Object Result -eq 'reportOnlyInterrupted').Count
    }

    # Group results for different views
    $byPolicy = $Results | Group-Object PolicyName | Select-Object @{
        N='Policy';E={$_.Name}
    }, @{
        N='Total';E={$_.Count}
    }, @{
        N='Success';E={($_.Group | Where-Object Result -eq 'reportOnlySuccess').Count}
    }, @{
        N='Failure';E={($_.Group | Where-Object Result -eq 'reportOnlyFailure').Count}
    }, @{
        N='Interrupted';E={($_.Group | Where-Object Result -eq 'reportOnlyInterrupted').Count}
    }

    $byApp = $Results | Group-Object AppDisplayName | Select-Object @{
        N='Application';E={$_.Name}
    }, @{
        N='Total Sign-ins';E={$_.Count}
    }, @{
        N='Policies Applied';E={($_.Group | Select-Object PolicyName -Unique).Count}
    }, @{
        N='Would Block';E={($_.Group | Where-Object Result -eq 'reportOnlyFailure').Count}
    }

    $byUser = $Results | Group-Object UserPrincipalName | Select-Object @{
        N='User';E={$_.Name}
    }, @{
        N='Total Sign-ins';E={$_.Count}
    }, @{
        N='Unique Apps';E={($_.Group | Select-Object AppDisplayName -Unique).Count}
    }, @{
        N='Would Block';E={($_.Group | Where-Object Result -eq 'reportOnlyFailure').Count}
    }

    # Create summary report
    New-HTML -Title "Conditional Access Report-Only Analysis" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Analysis Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Report Details</h3>
                <ul>
                    <li>Generated By: $($metadata.GeneratedBy)</li>
                    <li>Generated On: $($metadata.GeneratedOn)</li>
                    <li>Total Sign-ins Analyzed: $($metadata.TotalSignIns)</li>
                    <li>Unique Users: $($metadata.UniqueUsers)</li>
                    <li>Unique Applications: $($metadata.UniqueApps)</li>
                    <li>Policy Results:</li>
                    <ul>
                        <li>Success: $($metadata.SuccessCount)</li>
                        <li>Would Block: $($metadata.FailureCount)</li>
                        <li>Requires User Action: $($metadata.InterruptedCount)</li>
                    </ul>
                </ul>
"@
            }
        }
        
        New-HTMLSection -HeaderText "Policy Analysis" {
            New-HTMLTable -DataTable $byPolicy -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') -DisableStateSave -FixedHeader -HideFooter:$false -ScrollX -SearchBuilder {
                New-TableCondition -Name 'Failure' -ComparisonType number -Operator gt -Value 0 -BackgroundColor Salmon -Color Black
                New-TableCondition -Name 'Success' -ComparisonType number -Operator gt -Value 0 -BackgroundColor LightGreen -Color Black
                New-TableCondition -Name 'Interrupted' -ComparisonType number -Operator gt -Value 0 -BackgroundColor LightBlue -Color Black
            }
        }

        New-HTMLSection -HeaderText "Application Impact" {
            New-HTMLTable -DataTable $byApp -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') -DisableStateSave -FixedHeader -HideFooter:$false -ScrollX -SearchBuilder {
                New-TableCondition -Name 'Would Block' -ComparisonType number -Operator gt -Value 0 -BackgroundColor Salmon -Color Black
            }
        }

        New-HTMLSection -HeaderText "User Impact" {
            New-HTMLTable -DataTable $byUser -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') -DisableStateSave -FixedHeader -HideFooter:$false -ScrollX -SearchBuilder {
                New-TableCondition -Name 'Would Block' -ComparisonType number -Operator gt -Value 0 -BackgroundColor Salmon -Color Black
            }
        }
    }

    # Create detailed report if requested
    if ($IncludeDetailedReport) {
        New-HTML -Title "Conditional Access Report-Only Analysis - Detailed Results" -FilePath $detailedHtmlPath -ShowHTML {
            New-HTMLSection -HeaderText "Detailed Results" {
                New-HTMLTable -DataTable $Results -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') -DisableStateSave -FixedHeader -HideFooter:$false -ScrollX -SearchBuilder {
                    New-TableCondition -Name 'Result' -ComparisonType string -Operator eq -Value 'reportOnlyFailure' -BackgroundColor Salmon -Color Black
                    New-TableCondition -Name 'Result' -ComparisonType string -Operator eq -Value 'reportOnlySuccess' -BackgroundColor LightGreen -Color Black
                    New-TableCondition -Name 'Result' -ComparisonType string -Operator eq -Value 'reportOnlyInterrupted' -BackgroundColor LightBlue -Color Black
                }
            }
        }
    }

    Write-Host "`nReports generated:" -ForegroundColor Green
    Write-Host "CSV Report: $csvPath" -ForegroundColor Green
    Write-Host "Summary HTML Report: $htmlPath" -ForegroundColor Green
    if ($IncludeDetailedReport) {
        Write-Host "Detailed HTML Report: $detailedHtmlPath" -ForegroundColor Green
    }
    
    return @{
        CSVPath = $csvPath
        HTMLPath = $htmlPath
        DetailedHTMLPath = if ($IncludeDetailedReport) { $detailedHtmlPath } else { $null }
    }
}







# Store output in variable to prevent display
# $null = Get-CAReportOnlyAnalysis -ReportPath "C:\CAReports"

# Or run directly
# Get-CAReportOnlyAnalysis -ReportPath "C:\CAReports" > $null



# Example usage:
# Get-CAReportOnlyAnalysis -DaysToAnalyze 0 -ReportPath "C:\CAReports"

# Example usage:
# For last 7 hours (default)
# Get-CAReportOnlyAnalysis -ReportPath "C:\CAReports"

# For last 7 days
# Get-CAReportOnlyAnalysis -Hours 7 -UseDays -ReportPath "C:\CAReports"

# For specific number of hours
Get-CAReportOnlyAnalysis -Hours 1 -ReportPath "C:\CAReports"





# For last 7 hours without detailed report (default)
# Get-CAReportOnlyAnalysis -ReportPath "C:\CAReports"

# For last 7 hours with detailed report
# Get-CAReportOnlyAnalysis -ReportPath "C:\CAReports" -IncludeDetailedReport

# For last 7 days with detailed report
# Get-CAReportOnlyAnalysis -Hours 1 -UseDays -ReportPath "C:\CAReports" -IncludeDetailedReport