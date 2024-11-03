param(
  [string]$CredentialPath = "D:\InterfaceAeosPrestatairesPJ\config\sql_cred.xml",
  [string]$SqlUser = "nedap"
)

$secure = Read-Host "M9CiRSfWsCNF7wSW7pfEL" -AsSecureString
$cred = New-Object System.Management.Automation.PSCredential($SqlUser, $secure)
$cred | Export-Clixml -Path $CredentialPath
Write-Host "✅ Identifiants SQL sauvegardés dans $CredentialPath"
