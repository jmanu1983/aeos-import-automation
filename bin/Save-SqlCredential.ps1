<#
.SYNOPSIS
    Save SQL Server credentials to an encrypted CliXml file.

.DESCRIPTION
    Prompts for a SQL password and stores the credential securely using
    Windows DPAPI encryption (Import/Export-Clixml).

.PARAMETER CredentialPath
    Path where the encrypted credential file will be saved.

.PARAMETER SqlUser
    SQL Server username.
#>
param(
    [string]$CredentialPath = ".\config\sql_cred.xml",
    [string]$SqlUser = (Read-Host "SQL username")
)

$secure = Read-Host "SQL password for '$SqlUser'" -AsSecureString
$cred = New-Object System.Management.Automation.PSCredential($SqlUser, $secure)
$cred | Export-Clixml -Path $CredentialPath
Write-Host "Credentials saved to $CredentialPath" -ForegroundColor Green
