<#
One-time tenant setup for the DANSS daily mail-flow report. Run once, signed in
as a Global Administrator. Safe to re-run.

  1. Grants the report app admin consent for ExchangeMessageTrace.Read.All (read-only).
  2. Creates the service principal for Microsoft's Message Trace app, which the
     Graph message trace API requires.
  3. Lets the report app send mail as ONE mailbox only, using Exchange RBAC for
     Applications. The app gets no Mail.Send permission in Entra, so it can't
     send as anyone else.
#>
$ErrorActionPreference = "Stop"

$TenantId          = "e14ef6fe-5a8f-41bb-bc63-0a2b308227ca"
$AppId             = "efcdd68c-7246-4b84-833e-01cf23a456a6"  # DANSS Daily Mail-Flow Report
$SpObjectId        = "9db89b70-9d6c-4ccd-8cd5-0ed6eabcb8e6"  # its enterprise application
$Mailbox           = "tarek.kyyali@friendgroup-iq.com"
$MessageTraceAppId = "8bd644d1-64a1-4d4b-ae52-2e0cbf64e373"
$ScopeName         = "DANSS Report Sender"

foreach ($m in "Microsoft.Graph.Applications", "ExchangeOnlineManagement") {
    if (-not (Get-Module -ListAvailable $m)) {
        Write-Host "Installing $m from the PowerShell Gallery..."
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
}

function Step($name, [scriptblock]$action) {
    try { & $action | Out-Null; Write-Host "[done] $name" -ForegroundColor Green }
    catch {
        if ($_.Exception.Message -match "already exists|Permission being assigned already exists") {
            Write-Host "[skip] $name (already in place)" -ForegroundColor DarkGray
        } else { throw }
    }
}

# --- Entra ------------------------------------------------------------------
Connect-MgGraph -TenantId $TenantId -NoWelcome `
    -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$roleId  = ($graphSp.AppRoles | Where-Object Value -eq "ExchangeMessageTrace.Read.All").Id

Step "Admin consent: ExchangeMessageTrace.Read.All" {
    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SpObjectId |
        Where-Object AppRoleId -eq $roleId
    if ($existing) { throw "already exists" }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SpObjectId `
        -PrincipalId $SpObjectId -ResourceId $graphSp.Id -AppRoleId $roleId
}
Step "Message Trace service principal" {
    if (Get-MgServicePrincipal -Filter "appId eq '$MessageTraceAppId'") { throw "already exists" }
    New-MgServicePrincipal -AppId $MessageTraceAppId
}
Disconnect-MgGraph | Out-Null

# --- Exchange Online ---------------------------------------------------------
Connect-ExchangeOnline -ShowBanner:$false

Step "Exchange service principal" {
    New-ServicePrincipal -AppId $AppId -ObjectId $SpObjectId -DisplayName "DANSS Daily Mail-Flow Report"
}
Step "Management scope '$ScopeName' ($Mailbox only)" {
    New-ManagementScope -Name $ScopeName -RecipientRestrictionFilter "PrimarySmtpAddress -eq '$Mailbox'"
}
Step "Application Mail.Send limited to that scope" {
    New-ManagementRoleAssignment -App $AppId -Role "Application Mail.Send" -CustomResourceScope $ScopeName
}

Write-Host "`nCheck: the app should be InScope=True for $Mailbox only."
Test-ServicePrincipalAuthorization -Identity $AppId -Resource $Mailbox |
    Format-Table RoleName, GrantedPermissions, AllowedResourceScope, InScope -AutoSize
Disconnect-ExchangeOnline -Confirm:$false
