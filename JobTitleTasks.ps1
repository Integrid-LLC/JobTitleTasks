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
    [string]$AutoTaskTicketNumber,

    [Parameter(Mandatory = $true)]
    [string]$FirstName,

    [Parameter(Mandatory = $true)]
    [string]$LastName,

    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [string]$JobTitle,

    [Parameter(Mandatory = $true)]
    [string[]]$Location,

    [Parameter(Mandatory = $true)]
    [string[]]$CommunityEmails,

    [Parameter(Mandatory = $false)]
    [string[]]$Equipment,

    [Parameter(Mandatory = $true)]
    [string]$CreatedBy,

    [Parameter(Mandatory = $true)]
    [string]$LicenseName,

    [Parameter(Mandatory = $true)]
    [string]$CompanyName,

    [Parameter(Mandatory = $true)]
    [bool]$AddPax8License
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


if ($AddPax8License) {
    $pax8SecretNames = @("Pax8-client-id", "Pax8-client-secret")
    $pax8Credentials = @{}
    foreach ($sn in $pax8SecretNames) {
        $name = $sn.Replace("Pax8-", "")
        $name = $name.Replace("-", "_")
        $secret = Get-AzKeyVaultSecret -VaultName "IntegridAPIKeys" -Name $sn
        $value = ConvertFrom-SecureString $secret.SecretValue -AsPlainText
        $pax8Credentials.Add($name, $value)
    }
    if ($pax8Credentials.Count -eq 0) {
        Write-Error "Pax8 API credentials not found" -ErrorAction Stop
    }

    $token = Get-Pax8Token -Credentials $pax8Credentials
    $companyId = Get-Pax8CompanyId -CompanyName $CompanyName -Token $token
    $productId = Search-Pax8ProductIds $LicenseName
    $subscription = Get-Pax8Subscription -CompanyId $companyId -ProductId $productId -Token $token
    $qtyIncremented = $subscription.quantity + 1
    $respAddLicense = Add-Pax8Subscription -SubscriptionId $subscription.id -Quantity $qtyIncremented -Token $token
    if ($null -eq $respAddLicense) {
        Write-Error "Failed to add license to PAX8 subscription" -ErrorAction Stop
    }
    else {
        Write-Output "=> License added to PAX8 subscription id: $($respAddLicense.id)"
    }
    
    # Check license quantities in M365 every 30 seconds until new one shows up.
    do {
        Start-Sleep -Seconds 30
        $licenseData = Get-LicenseData $LicenseName
    }
    while ($licenseData.ConsumedUnits -ge $licenseData.PrepaidUnits.Enabled)
    $respAssignLicense = Set-MgUserLicense -UserId $UserPrincipalName -AddLicenses @{SkuId = $LicenseData.SkuId } -RemoveLicenses @()
    if ($null -eq $respAssignLicense) {
        Write-Error "Failed to assign license `"$($LicenseName)`" to $($UserPrincipalName)" -ErrorAction Stop
    }
    else {
        Write-Output "=> License `"$($LicenseName)`" assigned to $($respAssignLicense.DisplayName)"
    }
}


# ===========================================================================
if ($JobTitle -eq "Manager") {
    # First Bank email
    $respFirstBankManager = Send-FirstBankEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Rio
    if ($respFirstBankManager -eq "Success") {
        Write-Output "=> First Bank email sent"
    }
    else {
        Write-Warning "=> Failed to send First Bank email"
    }

    # Add to community emails
    $respSharedMailbox = Grant-SharedMailboxPermissions -Requestor $UserPrincipalName -SharedMailboxes $communityEmails
    if ($null -eq $respSharedMailbox) {
        Write-Warning "=> Failed to add user $UserPrincipalName to community emails"
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
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Scheduling Required"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Needs Attention"
    }
    $respCloseTicket = Update-AutoTaskTicket -Credentials $atCredentials @closeTicketParams
    if ($null -eq $respCloseTicket) {
        Write-Warning "=> Failed to close AutoTask ticket"
    }
    else {
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber closed."
    }
}

elseif ($JobTitle -eq "Regional") {
    # First Bank and Credit Card email
    $respFirstBankCCRegional = Send-FirstBankAndCreditCardEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Rio
    if ($respFirstBankCCRegional -eq "Success") {
        Write-Output "=> First Bank and Credit Card email sent"
    }
    else {
        Write-Warning "=> Failed to send First Bank and Credit Card email"
    }

    # ResMan email
    $respResMan = Send-ResManEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "dave@integrid.net" # TODO: Remove in production; defaults to Ebonie
    if ($respResMan -eq "Success") {
        Write-Output "=> ResMan email sent"
    }
    else {
        Write-Warning "Failed to send ResMan email"
    }

    # Add to Regional Managers group
    $regionalsGroupId = Get-MgGroup -Filter "DisplayName eq 'Regional Managers'" | Select-Object -ExpandProperty Id
    $userId = Get-MgUser -UserId $UserPrincipalName | Select-Object -ExpandProperty Id
    $regionalsStdErr = & { New-MgGroupMember -GroupId $regionalsGroupId -DirectoryObjectId $userId } 2>&1
    if ($regionalsStdErr.Count -ne 0) {
        Write-Warning "=> Failed to add user $UserPrincipalName to Regional Managers group"
        Write-Output $regionalsStdErr
    }
    else {
        Write-Output "=> User $UserPrincipalName added to Regional Managers group"
    }
    
    # Add to community emails
    $respSharedMailbox = Grant-SharedMailboxPermissions -Requestor $UserPrincipalName -SharedMailboxes $communityEmails
    if ($null -eq $respSharedMailbox) {
        Write-Warning "=> Failed to add user $UserPrincipalName to community emails"
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
            UserPrincipalName    = $UserPrincipalName
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
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Accounting") {
    # First Bank Email
    $respFirstBankAccounting = Send-FirstBankEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
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
            UserPrincipalName    = $UserPrincipalName
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
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Compliance") {
    # Equipment Approval
    if (($null -ne $Equipment) -and ($Equipment.Count -ne 0)) {
        $respGetTicket = Get-AutoTaskTicket -Credentials $atCredentials -TicketId $AutoTaskTicketId
        $equipmentParams = @{
            FirstName            = $FirstName
            LastName             = $LastName
            UserPrincipalName    = $UserPrincipalName
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
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Management") {
    # ResMan Email
    $respResMan = Send-ResManEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
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
            UserPrincipalName    = $UserPrincipalName
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
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Maintenance") {
    # Close ticket
    $closeTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "status" -Label "Scheduling Required"
        QueueId  = Get-AutoTaskPicklistValue -Picklist $atPicklist -Field "queueid" -Label "Needs Attention"
    }
    $respCloseTicket = Update-AutoTaskTicket -Credentials $atCredentials @closeTicketParams
    if ($null -eq $respCloseTicket) {
        Write-Output "=> Failed to close ticket"
    }
    else {
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber closed."
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