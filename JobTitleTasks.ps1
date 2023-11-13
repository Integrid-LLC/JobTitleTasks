param (
    [Parameter(Mandatory = $true)]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [int]$AutoTaskTicketId,

    [Parameter(Mandatory = $true)]
    [string]$FirstName,

    [Parameter(Mandatory = $true)]
    [string]$LastName,

    [Parameter(Mandatory = $true)]
    [string]$UPN,

    [Parameter(Mandatory = $true)]
    [string]$JobTitle,

    [Parameter(Mandatory = $true)]
    [string[]]$Location,

    [Parameter(Mandatory = $true)]
    [string[]]$CommunityEmails,

    [Parameter(Mandatory = $false)]
    [string[]]$Equipment
)

Import-Module OnboardingUtilities
Import-Module WynnefieldCustomUtilities

Connect-AzAccount -Identity -Subscription "Integrid Development" | Out-Null
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome
$exchangeOrg = Get-MgDomain | Where-Object { $_.IsInitial -eq $true } | Select-Object -ExpandProperty Id
Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $exchangeOrg -ShowBanner:$false


$atSecretNames = @("Autotask-ApiIntegrationCode", "Autotask-UserName", "Autotask-Secret")
$atCredentials = @{}
foreach ($sn in $atSecretNames) {
    $name = $sn.Replace("Autotask-", "")
    $secret = Get-AzKeyVaultSecret -VaultName "IntegridAPIKeys" -Name $sn
    $value = ConvertFrom-SecureString $secret.SecretValue -AsPlainText
    $atCredentials.Add($name, $value)
}
if ($atCredentials.Count -eq 0) {
    throw "Autotask API credentials not found"
}

$atPicklist = Get-AutoTaskTicketPicklist $atCredentials
if ($atPicklist.Count -eq 0) {
    throw "Autotask ticket picklist not found"
}


# ===========================================================================
if ($JobTitle -eq "Manager") {
    # First Bank email
    $respFirstBankManager = Send-FirstBankEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $upn -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Rio
    if ($respFirstBankManager -eq "Success") {
        Write-Output "=> First Bank email sent"
    }
    else {
        Write-Warning "=> Failed to send First Bank email"
    }

    # Add to community emails
    $respSharedMailbox = Grant-SharedMailboxPermissions -Requestor $upn -SharedMailboxes $communityEmails
    if ($null -eq $respSharedMailbox) {
        Write-Warning "=> Failed to add user $upn to community emails"
    }
    else {
        foreach ($r in $respSharedMailbox) {
            Write-Output "=> Full Access permission to mailbox $($r.FullAccess.Identity) granted to user SID $($r.FullAccess.UserSid)"
            Write-Output "=> SendAs permission to mailbox $($r.SendAs.Identity) granted to user SID $($r.SendAs.TrusteeSidString)"
        }
    }

    # Close ticket
    $closeTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Work Completed~"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Accepted"
    }
    $respCloseTicket = Update-AutoTaskTicket -Credentials $atCredentials @closeTicketParams
    if ($null -eq $respCloseTicket) {
        Write-Warning "=> Failed to close AutoTask ticket"
    }
    else {
        Write-Output "=> AutoTask ticket closed with Id: $($respCloseTicket.ItemId)"
    }
}

elseif ($JobTitle -eq "Regional") {
    # First Bank and Credit Card email
    $respFirstBankCCRegional = Send-FirstBankAndCreditCardEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $upn -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Rio
    if ($respFirstBankCCRegional -eq "Success") {
        Write-Output "=> First Bank and Credit Card email sent"
    }
    else {
        Write-Warning "=> Failed to send First Bank and Credit Card email"
    }

    # ResMan email
    $respResMan = Send-ResManEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $upn -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Ebonie
    if ($respResMan -eq "Success") {
        Write-Output "=> ResMan email sent"
    }
    else {
        Write-Warning "Failed to send ResMan email"
    }

    # Add to Regional Managers group
    $regionalsGroupId = Get-MgGroup -Filter "DisplayName eq 'Regional Managers'" | Select-Object -ExpandProperty Id
    $userId = Get-MgUser -UserId $upn | Select-Object -ExpandProperty Id
    $regionalsStdErr = & { New-MgGroupMember -GroupId $regionalsGroupId -DirectoryObjectId $userId } 2>&1
    if ($regionalsStdErr.Count -ne 0) {
        Write-Warning "=> Failed to add user $upn to Regional Managers group"
        Write-Output $regionalsStdErr
    }
    else {
        Write-Output "=> User $upn added to Regional Managers group"
    }
    
    # Add to community emails
    $respSharedMailbox = Grant-SharedMailboxPermissions -Requestor $upn -SharedMailboxes $communityEmails
    if ($null -eq $respSharedMailbox) {
        Write-Warning "=> Failed to add user $upn to community emails"
    }
    else {
        foreach ($r in $respSharedMailbox) {
            Write-Output "=> Full Access permission to mailbox $($r.FullAccess.Identity) granted to user SID $($r.FullAccess.UserSid)"
            Write-Output "=> SendAs permission to mailbox $($r.SendAs.Identity) granted to user SID $($r.SendAs.TrusteeSidString)"
        }
    }
    
    # Equipment Approval
    if (($null -ne $Equipment) -and ($Equipment.Count -ne 0)) {
        $respGetTicket = Get-AutoTaskTicket -Credentials $atCredentials -TicketId $AutoTaskTicketId
        $equipmentParams = @{
            FirstName            = $FirstName
            LastName             = $LastName
            UserPrincipalName    = $upn
            JobTitle             = $JobTitle
            Location             = $Location
            EquipmentList        = $Equipment
            AutoTaskTicketNumber = $respGetTicket.item.ticketNumber
            MailRecipient        = "dave@integrid.net" # TODO: Replace w/ Chrystal Rhodes in production (no default)
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Warning "=> Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }
    

    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Waiting Customer"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance"
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Warning "=> Failed to update ticket"
    }
    else {
        Write-Output "=> AutoTask ticket with Id: $($respUpdateTicket.ItemId) updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Accounting") {
    # First Bank Email
    $respFirstBankAccounting = Send-FirstBankEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $upn -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Rio
    if ($respFirstBankAccounting -eq "Success") {
        Write-Output "=> First Bank email sent"
    }
    else {
        Write-Output "Failed to send First Bank email"
    }

    # Equipment Approval
    if (($null -ne $Equipment) -and ($Equipment.Count -ne 0)) {
        $respGetTicket = Get-AutoTaskTicket -Credentials $atCredentials -TicketId $AutoTaskTicketId
        $equipmentParams = @{
            FirstName            = $FirstName
            LastName             = $LastName
            UserPrincipalName    = $upn
            JobTitle             = $JobTitle
            Location             = $Location
            EquipmentList        = $Equipment
            AutoTaskTicketNumber = $respGetTicket.item.ticketNumber
            MailRecipient        = "dave@integrid.net" # TODO: Replace w/ Tres in production (no default)
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Output "=> Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }

    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Waiting Customer"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance"
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Output "=> Failed to update ticket"
    }
    else {
        Write-Output "=> AutoTask ticket with Id: $($respUpdateTicket.ItemId) updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Compliance") {
    # Equipment Approval
    if (($null -ne $Equipment) -and ($Equipment.Count -ne 0)) {
        $respGetTicket = Get-AutoTaskTicket -Credentials $atCredentials -TicketId $AutoTaskTicketId
        $equipmentParams = @{
            FirstName            = $FirstName
            LastName             = $LastName
            UserPrincipalName    = $upn
            JobTitle             = $JobTitle
            Location             = $Location
            EquipmentList        = $Equipment
            AutoTaskTicketNumber = $respGetTicket.item.ticketNumber
            MailRecipient        = "dave@integrid.net" # TODO: Replace w/ Tres in production (no default)
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Output "=> Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }

    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Waiting Customer"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance"
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Output "=> Failed to update ticket"
    }
    else {
        Write-Output "=> AutoTask ticket with Id: $($respUpdateTicket.ItemId) updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Management") {
    # ResMan Email
    $respResMan = Send-ResManEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $upn -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Ebonie
    if ($respResMan -eq "Success") {
        Write-Output "=> ResMan email sent"
    }
    else {
        Write-Output "Failed to send ResMan email"
    }

    # Equipment Approval
    if (($null -ne $Equipment) -and ($Equipment.Count -ne 0)) {
        $respGetTicket = Get-AutoTaskTicket -Credentials $atCredentials -TicketId $AutoTaskTicketId
        $equipmentParams = @{
            FirstName            = $FirstName
            LastName             = $LastName
            UserPrincipalName    = $upn
            JobTitle             = $JobTitle
            Location             = $Location
            EquipmentList        = $Equipment
            AutoTaskTicketNumber = $respGetTicket.item.ticketNumber
            MailRecipient        = "dave@integrid.net" # TODO: Replace w/ Tres in production (no default)
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Output "=> Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }
    
    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Waiting Customer"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance"
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Output "=> Failed to update ticket"
    }
    else {
        Write-Output "=> AutoTask ticket with Id: $($respUpdateTicket.ItemId) updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Maintenance") {
    # Close ticket
    $closeTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Work Completed~"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Accepted"
    }
    $respCloseTicket = Update-AutoTaskTicket -Credentials $atCredentials @closeTicketParams
    if ($null -eq $respCloseTicket) {
        Write-Output "=> Failed to close ticket"
    }
    else {
        Write-Output "=> AutoTask ticket closed with Id: $($respCloseTicket.ItemId)"
    }
}

else {
    Write-Error "Unknown job title $JobTitle" -Category InvalidArgument
}

# Delete one-time schedule
$scheduleName = "$FirstName$LastName"
$automationAccountName = "Onboarding-Wynnefield"
$resourceGroupName = "RG-Dev"
Remove-AzAutomationSchedule -Name $scheduleName -AutomationAccountName $automationAccountName `
    -ResourceGroupName $resourceGroupName -Confirm:$false -Force