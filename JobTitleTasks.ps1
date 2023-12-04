param (
    [Parameter(Mandatory = $true)]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [string]$CompanyName,

    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [Parameter(Mandatory = $true)]
    [string]$SpSiteName,

    [Parameter(Mandatory = $true)]
    [string]$SpListName,

    [Parameter(Mandatory = $true)]
    [string]$FirstName,

    [Parameter(Mandatory = $true)]
    [string]$LastName,

    [Parameter(Mandatory = $true)]
    [string[]]$Location,

    [Parameter(Mandatory = $true)]
    [string]$JobTitle,

    [Parameter(Mandatory = $false)]
    [string[]]$Equipment,

    [Parameter(Mandatory = $true)]
    [string]$CreatedByEmail,

    [Parameter(Mandatory = $true)]
    [string]$CreatedByDisplayName,

    [Parameter(Mandatory = $true)]
    [string]$PasswordSender,

    [Parameter(Mandatory = $true)]
    [string]$LicenseName,

    [Parameter(Mandatory = $true)]
    [string]$ApprovalWebhookUrl,

    [Parameter(Mandatory = $true)]
    [string]$FollowUpAutomationAccount,

    [Parameter(Mandatory = $true)]
    [string]$FollowUpResourceGroup,

    [Parameter(Mandatory = $true)]
    [string]$FollowUpRunbook,

    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory = $true)]
    [int]$AutoTaskCompanyId,

    [Parameter(Mandatory = $true)]
    [int]$HuduCompanyId,

    [Parameter(Mandatory = $true)]
    [string]$M365Location,

    [Parameter(Mandatory = $true)]
    [string]$AutoTaskLocationId,

    [Parameter(Mandatory = $true)]
    [string[]]$LocationEmails,

    [Parameter(Mandatory = $true)]
    [object]$UserData,

    [Parameter(Mandatory = $true)]
    [int]$AutoTaskTicketId,

    [Parameter(Mandatory = $true)]
    [string]$AutoTaskTicketNumber,

    [Parameter(Mandatory = $true)]
    [int]$AutoTaskContactId,

    [Parameter(Mandatory = $true)]
    [string]$HuduPasswordId,

    [Parameter(Mandatory = $true)]
    [bool]$OneTimeSecretSuccess,

    [Parameter(Mandatory = $true)]
    [obj]$LicenseData,

    [Parameter(Mandatory = $true)]
    [bool]$licenseApprovalRequired,

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
    Write-Error "Autotask API credentials not found" -ErrorAction Stop
}

$atPicklist = Get-AutoTaskTicketPicklist $atCredentials
if ($atPicklist.Count -eq 0) {
    Write-Error "Autotask ticket picklist not found" -ErrorAction Stop
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
        -MailRecipient "Rio.Chamberlain@wynnefieldproperties.com"
    if ($respFirstBankManager -eq "Success") {
        Write-Output "=> First Bank email sent"
    }
    else {
        Write-Warning "Failed to send First Bank email"
    }

    # Add to community emails
    $respSharedMailbox = Grant-SharedMailboxPermissions -Requestor $UserPrincipalName -SharedMailboxes $LocationEmails
    if ($null -eq $respSharedMailbox) {
        Write-Warning "Failed to add user $UserPrincipalName to community mailboxes $($LocationEmails -join ", ")"
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
        Status   = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "status" -Label "Scheduling Required").value
        QueueId  = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "queueid" -Label "Needs Attention").value
    }
    $respCloseTicket = Update-AutoTaskTicket -Credentials $atCredentials @closeTicketParams
    if ($null -eq $respCloseTicket) {
        Write-Warning "Failed to close AutoTask ticket"
    }
    else {
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber closed."
    }
}

elseif ($JobTitle -eq "Regional") {
    # First Bank and Credit Card email
    $respFirstBankCCRegional = Send-FirstBankAndCreditCardEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "Rio.Chamberlain@wynnefieldproperties.com"
    if ($respFirstBankCCRegional -eq "Success") {
        Write-Output "=> First Bank and Credit Card email sent"
    }
    else {
        Write-Warning "Failed to send First Bank and Credit Card email"
    }

    # ResMan email
    $respResMan = Send-ResManEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "Ebonie.Smith@wynnefieldproperties.com"
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
        Write-Warning "Failed to add user $UserPrincipalName to Regional Managers group"
        Write-Output $regionalsStdErr
    }
    else {
        Write-Output "=> User $UserPrincipalName added to Regional Managers group"
    }
    
    # Add to community emails
    $respSharedMailbox = Grant-SharedMailboxPermissions -Requestor $UserPrincipalName -SharedMailboxes $LocationEmails
    if ($null -eq $respSharedMailbox) {
        Write-Warning "Failed to add user $UserPrincipalName to community mailboxes $($LocationEmails -join ", ")"
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
            MailRecipient        = "Chrystal.Rhodes@wynnefieldproperties.com"
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Warning "Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }
    

    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "status" -Label "Waiting Customer").value
        QueueId  = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance").value
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Warning "Failed to update ticket"
    }
    else {
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Accounting") {
    # First Bank Email
    $respFirstBankAccounting = Send-FirstBankEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "Rio.Chamberlain@wynnefieldproperties.com"
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
            MailRecipient        = "Tres.Cobb@wynnefieldproperties.com"
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Output "Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }

    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "status" -Label "Waiting Customer").value
        QueueId  = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance").value
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Output "Failed to update ticket"
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
            MailRecipient        = "Tres.Cobb@wynnefieldproperties.com"
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Output "Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }

    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "status" -Label "Waiting Customer").value
        QueueId  = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance").value
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Output "Failed to update ticket"
    }
    else {
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Corporate - Management") {
    # ResMan Email
    $respResMan = Send-ResManEmail -FirstName $FirstName -LastName $LastName `
        -UserPrincipalName $UserPrincipalName -JobTitle $JobTitle -Location ($Location -join ", ") `
        -MailRecipient "Ebonie.Smith@wynnefieldproperties.com"
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
            MailRecipient        = "Tres.Cobb@wynnefieldproperties.com"
        }
        $respEquipment = Send-EquipmentApprovalEmail @equipmentParams
        if ($respEquipment -eq "Success") {
            Write-Output "=> Equipment approval email sent to $($equipmentParams.MailRecipient)"
        }
        else {
            Write-Output "Failed to send equipment approval email to $($equipmentParams.MailRecipient)"
            Write-Output $respEquipment
        }
    }
    else {
        Write-Output "=> Equipment approval not required. No equipment requested."
    }
    
    # Update ticket status & queue
    $updateTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "status" -Label "Waiting Customer").value
        QueueId  = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "queueid" -Label "Pending Acceptance").value
    }
    $respUpdateTicket = Update-AutoTaskTicket -Credentials $atCredentials @updateTicketParams
    if ($null -eq $respUpdateTicket) {
        Write-Output "Failed to update ticket"
    }
    else {
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber updated to status 'Waiting Customer' and queue 'Pending Acceptance'"
    }
}

elseif ($JobTitle -eq "Maintenance") {
    # Close ticket
    $closeTicketParams = @{
        TicketId = $AutoTaskTicketId
        Status   = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "status" -Label "Scheduling Required").value
        QueueId  = (Get-AutoTaskPicklistItem -Picklist $atPicklist -Field "queueid" -Label "Needs Attention").value
    }
    $respCloseTicket = Update-AutoTaskTicket -Credentials $atCredentials @closeTicketParams
    if ($null -eq $respCloseTicket) {
        Write-Output "Failed to close ticket"
    }
    else {
        Write-Output "=> AutoTask ticket number $AutoTaskTicketNumber closed."
    }
}

else {
    Write-Error "Unknown job title $JobTitle" -Category InvalidArgument -ErrorAction Stop
}

# Delete one-time schedule
$scheduleName = "$FirstName$LastName"
$automationAccountName = "Onboarding-Wynnefield"
$resourceGroupName = "RG-Dev"
Remove-AzAutomationSchedule -Name $scheduleName -AutomationAccountName $automationAccountName `
    -ResourceGroupName $resourceGroupName -Confirm:$false -Force