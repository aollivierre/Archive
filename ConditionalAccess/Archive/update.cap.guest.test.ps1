function Update-ConditionalAccessPolicyExclusion {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$PolicyId,
        
        [Parameter()]
        [ValidateSet(
            'internalGuest', 
            'b2bCollaborationGuest', 
            'b2bCollaborationMember', 
            'b2bDirectConnectUser', 
            'otherExternalUser', 
            'serviceProvider'
        )]
        [string[]]$GuestTypes = @('internalGuest')
    )
    
    try {
        # Get current policy
        $policy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyId
        
        # Create update body
        $bodyParams = @{
            conditions = @{
                users = @{
                    excludeGuestsOrExternalUsers = @{
                        guestOrExternalUserTypes = $GuestTypes -join ','
                        externalTenants          = @{
                            membershipKind = "all"
                        }
                    }
                }
            }
        }

        # Update policy
        $null = Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyId -BodyParameter $bodyParams
        
        Write-Host "Successfully updated policy '$($policy.DisplayName)' to exclude guest types: $($GuestTypes -join ', ')" -ForegroundColor Green
        
        # Return updated policy
        return Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyId
    }
    catch {
        Write-Error "Failed to update policy: $_"
    }
}


$params = @{
    PolicyId = 'e1e3962e-1286-422c-b503-9a8de2ee8202'
    GuestTypes = @(
        'internalGuest',
        'b2bCollaborationGuest',
        'b2bCollaborationMember',
        'b2bDirectConnectUser',
        'otherExternalUser',
        'serviceProvider'
    )
}

Update-ConditionalAccessPolicyExclusion @params