#requires -Version 5.1
param(
  [string]$ConfigPath = "D:\DemoInterfaceAccredites\config\accredites.config.json"
)

try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [Text.Encoding]::UTF8
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ======================
# LOGGING (UTF-8, no BOM)
# ======================
$Global:LogFile = "D:\data\11345\transferts\out\logs\log_import_{0}.txt" -f (Get-Date -Format "yyyyMMdd")
$Global:LogWriter = $null
function Initialize-Logger {
  if ($Global:LogWriter) { try { $Global:LogWriter.Flush(); $Global:LogWriter.Dispose() } catch {} ; $Global:LogWriter = $null }
  $enc = New-Object System.Text.UTF8Encoding($false)
  $dir = Split-Path -Path $Global:LogFile -Parent
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $Global:LogWriter = New-Object System.IO.StreamWriter($Global:LogFile, $true, $enc)
  $Global:LogWriter.AutoFlush = $true
}
function Close-Logger {
  if ($Global:LogWriter) { try { $Global:LogWriter.Flush(); $Global:LogWriter.Dispose() } catch {} ; $Global:LogWriter = $null }
}
function Write-Log([string]$Message, [string]$Level = "INFO") {
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $line = "[{0}] [{1}] {2}" -f $ts, $Level.ToUpper(), $Message
  Write-Host $line
  if (-not $Global:LogWriter) { Initialize-Logger }
  $Global:LogWriter.WriteLine($line)
}
function Write-Section([string]$Title){ Write-Log ("--- {0} ---" -f $Title) }
function Write-KV([string]$K,[string]$V){ Write-Log ("{0} : {1}" -f $K, $V) }

# ======================
# CONFIG
# ======================
if (-not (Test-Path $ConfigPath)) { throw "Config introuvable: $ConfigPath" }
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Ensure folders
foreach($p in @($config.WorkingDir,$config.LogDir,$config.FolderPathImport,$config.FolderPathOutput)){
  if ($p -and -not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}
$Global:LogFile = Join-Path $config.LogDir ("log_import_{0}.txt" -f (Get-Date -Format "yyyyMMdd"))
Initialize-Logger
Write-Section "Start"
Write-KV "Config" $ConfigPath
$vendCat = if ($config.VendorDefaultCategory) { $config.VendorDefaultCategory } else { "NULL" }
$vendPob = if ($config.VendorPlaceOfBusiness) { $config.VendorPlaceOfBusiness } else { "NULL" }
$vendSoap = if ($config.EnableVendorSoap -eq $true) { "ON" } else { "OFF" }
Write-KV "VendorDefaultCategory" $vendCat
Write-KV "VendorPlaceOfBusiness" $vendPob
Write-KV "VendorSoap" $vendSoap

# ======================
# HELPERS
# ======================
function Get-ConnString() {
  if ($config.SqlAuthMode -and ($config.SqlAuthMode -ieq 'Windows')) {
    return "Server={0};Database={1};Trusted_Connection=True;TrustServerCertificate=True;" -f $config.SqlServer, $config.SqlDatabase
  } else {
    if (-not (Test-Path $config.CredentialPath)) { throw "CredentialPath introuvable: $($config.CredentialPath)" }
    $cred = Import-Clixml -LiteralPath $config.CredentialPath
    $pwd  = $cred.GetNetworkCredential().Password
    return "Server={0};Database={1};User ID={2};Password={3};TrustServerCertificate=True;" -f $config.SqlServer, $config.SqlDatabase, $config.SqlUser, $pwd
  }
}
function Decode-FileContent([byte[]]$Bytes) {
  if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return "" }

  # 1) Detect BOM (UTF-8 / UTF-16 LE / UTF-16 BE)
  if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
    return [Text.Encoding]::UTF8.GetString($Bytes, 0, $Bytes.Length)
  }
  if ($Bytes.Length -ge 2) {
    if ($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
      return [Text.Encoding]::Unicode.GetString($Bytes, 0, $Bytes.Length)
    }
    if ($Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) {
      return [Text.Encoding]::BigEndianUnicode.GetString($Bytes, 0, $Bytes.Length)
    }
  }

  # 2) Try strict UTF-8 roundtrip
  try {
    $utf8 = [Text.Encoding]::UTF8.GetString($Bytes,0,$Bytes.Length)
    $roundtrip = [Text.Encoding]::UTF8.GetBytes($utf8)
    if ($roundtrip.Length -eq $Bytes.Length) {
      $diff = $false
      for($i=0;$i -lt $Bytes.Length;$i++){
        if ($Bytes[$i] -ne $roundtrip[$i]) { $diff = $true; break }
      }
      if (-not $diff) { return $utf8 }
    }
  } catch { }

  # 3) Fallback Windows-1252 (typical Excel export)
  return [Text.Encoding]::GetEncoding(1252).GetString($Bytes,0,$Bytes.Length)
}
function Detect-Separator([string]$HeaderLine){
  if ([string]::IsNullOrEmpty($HeaderLine)) { return ';' }
  $chars = $HeaderLine.ToCharArray()
  $comma = (@($chars | Where-Object { $_ -eq ',' })).Count
  $semi  = (@($chars | Where-Object { $_ -eq ';' })).Count
  if ($semi -ge $comma) { return ';' } else { return ',' }
}
function Parse-CsvLine([string]$Line,[char]$Sep){
  if ($null -eq $Line) { return @() }
  $res = New-Object System.Collections.Generic.List[string]
  $sb  = New-Object System.Text.StringBuilder
  $inQ = $false
  for($i=0;$i -lt $Line.Length;$i++){
    $ch = $Line[$i]
    if ($ch -eq '"'){
      if ($inQ -and $i+1 -lt $Line.Length -and $Line[$i+1] -eq '"'){ [void]$sb.Append('"'); $i++ } else { $inQ = -not $inQ }
    } elseif ($ch -eq $Sep -and -not $inQ) {
      $res.Add($sb.ToString()); $sb.Clear() | Out-Null
    } else { [void]$sb.Append($ch) }
  }
  $res.Add($sb.ToString()); return $res.ToArray()
}
function Canon([string]$s){
  if ($null -eq $s) { return "" }
  $t = $s.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach($ch in $t.ToCharArray()){
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
  }
  $t = $sb.ToString().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
  $t = ($t -replace "[^a-z0-9/ ]"," ") -replace "\s+"," "
  return $t.Trim()
}
function Fix-Mojibake([string]$s){
  if ([string]::IsNullOrEmpty($s)) { return @($s) }
  $list = New-Object System.Collections.Generic.List[string]
  $list.Add($s) | Out-Null
  try {
    $b1252 = [Text.Encoding]::GetEncoding(1252).GetBytes($s)
    $u8    = [Text.Encoding]::UTF8.GetString($b1252,0,$b1252.Length)
    if ($u8 -ne $s) { $list.Add($u8) | Out-Null }
  } catch {}
  try {
    $bU8 = [Text.Encoding]::UTF8.GetBytes($s)
    $l1  = [Text.Encoding]::GetEncoding(1252).GetString($bU8,0,$bU8.Length)
    if ($l1 -ne $s) { $list.Add($l1) | Out-Null }
  } catch {}
  $rep = $s
  $rep = $rep -replace "Ã¯Â¿Â½","e" -replace "ÃƒÂ©","Ã©" -replace "ÃƒÂ¨","Ã¨" -replace "ÃƒÂª","Ãª" -replace "ÃƒÂ«","Ã«"
  $rep = $rep -replace "ÃƒÂ¢","Ã¢" -replace "ÃƒÂ´","Ã´" -replace "ÃƒÂ»","Ã»" -replace "ÃƒÂ®","Ã®" -replace "ÃƒÂ§","Ã§"
  $rep = $rep -replace "Ãƒ ","Ã " -replace "Ã‚",""
  if ($rep -ne $s) { $list.Add($rep) | Out-Null }
  return ($list | Select-Object -Unique)
}
function HeaderMatchesAny([string]$h,[string[]]$needles){
  $variants = Fix-Mojibake $h
  foreach($v in $variants){
    $hC = (Canon $v) -replace " ",""
    foreach($needle in $needles){
      $nC = (Canon $needle) -replace " ",""
      if (-not [string]::IsNullOrEmpty($nC)){
        if ($hC -like ("*"+$nC+"*")) { return $true }
      }
    }
  }
  return $false
}
function Build-HeaderMap([string[]]$Headers){
  $map=@{}

  # Fast-path uniquement si le layout est clairement celui attendu.
  # Attention : depuis l'ajout de la colonne "Remarques", un fichier peut aussi avoir 8 colononnes.
  if ($Headers.Length -eq 8) {
    $h7 = $Headers[7]
    if (HeaderMatchesAny $h7 @("oui/non","ouinon","badge","badge oui/non","has badge")) {
      $map["matricule"]=0
      $map["noms"]=1
      $map["prenoms"]=2
      $map["societe"]=3
      $map["reponse"]=4
      $map["echeance"]=5
      $map["entreprise"]=6
      $map["badgeyn"]=7
      return $map
    }
    if (HeaderMatchesAny $h7 @("remarques","remarque","commentaire","notes","note")) {
      $map["matricule"]=0
      $map["noms"]=1
      $map["prenoms"]=2
      $map["societe"]=3
      $map["reponse"]=4
      $map["echeance"]=5
      $map["entreprise"]=6
      $map["remarques"]=7
      return $map
    }
    # sinon : fall-through vers mapping intelligent
  }

  for($i=0;$i -lt $Headers.Length;$i++){
    $h=$Headers[$i]
    if (-not $map.ContainsKey("matricule") -and (HeaderMatchesAny $h @("matricule","index unique","personnel nr","personnel number","id"))) { $map["matricule"]=$i; continue }
    if (-not $map.ContainsKey("prenoms")   -and (HeaderMatchesAny $h @("prÃ©noms","prenoms","prÃ©nom","prenom","first name","firstname","prnoms"))) { $map["prenoms"]=$i; continue }
    if (-not $map.ContainsKey("noms")      -and (HeaderMatchesAny $h @("noms","nom","last name","lastname","surname"))) { $map["noms"]=$i; continue }
    if (-not $map.ContainsKey("societe")   -and (HeaderMatchesAny $h @("sociÃ©tÃ©/entreprise","societe/entreprise","sociÃ©tÃ©","societe","company"))) { $map["societe"]=$i; continue }
    if (-not $map.ContainsKey("reponse")   -and (HeaderMatchesAny $h @("rÃ©ponse","reponse","dÃ©but","debut","start","validfrom"))) { $map["reponse"]=$i; continue }
    if (-not $map.ContainsKey("echeance")  -and (HeaderMatchesAny $h @("Ã©chÃ©ance","echeance","fin","end","validto"))) { $map["echeance"]=$i; continue }
    if (-not $map.ContainsKey("entreprise")-and (HeaderMatchesAny $h @("entreprise","vendorcode","vendor code","code sociÃ©tÃ©","code societe"))) { $map["entreprise"]=$i; continue }

    # Optionnels
    if (-not $map.ContainsKey("badgeyn")   -and (HeaderMatchesAny $h @("oui/non","ouinon","badge","badge oui/non","has badge"))) { $map["badgeyn"]=$i; continue }
    if (-not $map.ContainsKey("remarques")-and (HeaderMatchesAny $h @("remarques","remarque","commentaire","notes","note"))) { $map["remarques"]=$i; continue }
  }

  foreach($must in @("matricule","noms","prenoms","societe","reponse","echeance","entreprise")){
    if (-not $map.ContainsKey($must)) { throw "Required header not found: " + $must }
  }
  return $map
}
function Parse-Date([string]$s){
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $t = $s.Trim().Replace([char]160,[char]32).Replace([char]8239,[char]32)
  $t = $t.Replace('-', '.').Replace('/', '.')
  $formats = @(
    "dd.MM.yyyy",
    "yyyy.MM.dd",
    "dd.MM.yy",
    "yyyy.MM.dd HH:mm:ss",
    "yyyy-MM-dd",
    "yyyy-MM-dd HH:mm:ss",
    "MM/dd/yyyy",
    "M/d/yyyy"
  )
  $out = [datetime]::MinValue
  if ([datetime]::TryParseExact($t, $formats,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$out)) {
    return $out.Date
  }
  $any = [datetime]::MinValue
  if ([datetime]::TryParse($t,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$any)) {
    return $any.Date
  }
  return $null
}

function Clean-Remarques([string]$s){
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $t = $s.Trim().Replace([char]160,[char]32).Replace([char]8239,[char]32)
  # Fichiers observÃ©s : "CURTIN;;;;;" / "Remarques;;;;;" -> supprimer les ; finaux
  $t = $t -replace ";+$",""
  # ProtÃ©ger le sÃ©parateur multi-freefields AEOS
  $t = $t -replace "\|","/"
  $t = $t -replace "\r|\n"," "
  $t = $t.Trim()
  if ($t.Length -gt 30) { $t = $t.Substring(0,30) }
  if ([string]::IsNullOrWhiteSpace($t)) { return $null }
  return $t
}

function Set-Str {
  param(
    [Parameter(Mandatory=$true)] [System.Data.DataRow] $Row,
    [Parameter(Mandatory=$true)] [string] $ColumnName,
    [AllowNull()] [string] $Value,
    [Parameter(Mandatory=$true)] $ColDefs
  )
  $def = $ColDefs[$ColumnName]
  if ([string]::IsNullOrEmpty($Value)) {
    $Row[$ColumnName] = [DBNull]::Value
    return
  }
  $s = [string]$Value
  if ($s) { $s = $s.Trim() }
  $max = $def.MaxChars
  if ($max -gt 0 -and $max -lt 4000 -and $s.Length -gt $max) { $s = $s.Substring(0,$max) }
  $Row[$ColumnName] = $s
}

# ========= SOAP (AEOS) =========
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}

function Get-SoapCredentials {
  $res = @{ User=$null; Password=$null }
  if ($config.SoapCredentialPath -and (Test-Path $config.SoapCredentialPath)) {
    try {
      $c = Import-Clixml -Path $config.SoapCredentialPath
      $res.User = $c.UserName
      $res.Password = $c.GetNetworkCredential().Password
      return $res
    } catch {}
  }
  if ($false) {
    $res.User = $config.SoapUser
    $res.Password = $config.SoapPassword
  }
  return $res
}
function Get-VendorServiceUrl([string]$Base){
  if ([string]::IsNullOrWhiteSpace($Base)) { return $null }
  if ($Base -match "/services/[^/]+$") { return $Base }
  return ($Base.TrimEnd('/')) + "/services/VendorService"
}
function Enable-SoapInsecureTlsIfRequested {
  if ($config.SoapIgnoreCertErrors -eq $true) {
    try {
      [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($sender,$cert,$chain,$errors) return $true }
      Write-Log "SOAP TLS validation is DISABLED (SoapIgnoreCertErrors=true)." "WARN"
    } catch { Write-Log ("Failed to set TLS bypass: {0}" -f $_.Exception.Message) "WARN" }
  }
}
function Invoke-AeosSoap {
  param(
    [Parameter(Mandatory=$true)][string]$SvcUrl,
    [Parameter(Mandatory=$true)][string]$InnerXml,
    [string]$User,[string]$Password,
    [int]$TimeoutSec=30,[int]$MaxRetry=3
  )
  $env = @"
<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/">
  <soapenv:Header/>
  <soapenv:Body>
$InnerXml
  </soapenv:Body>
</soapenv:Envelope>
"@
  $bytes = [Text.Encoding]::UTF8.GetBytes($env)
  $headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction"   = ""
  }
  if ($User -and $Password) {
    $headers["Authorization"] = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($User + ":" + $Password)))
  }
  $attempt=0
  while ($true){
    $attempt++
    try {
      $resp = Invoke-WebRequest -UseBasicParsing -Method Post -Uri $SvcUrl -Headers $headers -Body $bytes -TimeoutSec $TimeoutSec
      if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 300) { return $resp.Content }
      throw ("HTTP " + $resp.StatusCode + " " + $resp.StatusDescription)
    } catch {
      if ($attempt -lt $MaxRetry) {
        Write-Log ("SOAP attempt {0}/{1} failed: {2} - retrying..." -f $attempt,$MaxRetry,$_.Exception.Message) "WARN"
        Start-Sleep -Seconds ([Math]::Min(5*$attempt,10))
        continue
      } else {
        Write-Log ("SOAP failed after {0} attempts: {1}" -f $MaxRetry,$_.Exception.Message) "ERROR"
        throw
      }
    }
  }
}
function Find-AeosVendorByCode {
  param(
    [Parameter(Mandatory=$true)][string]$Code,
    [Parameter(Mandatory=$true)][string]$SvcUrl,
    [string]$User,[string]$Password
  )
  $body = @"
<ns:VendorSearchInfo xmlns:ns="http://www.nedap.com/aeosws/schema">
  <ns:VendorInfo>
    <ns:Code>$([System.Security.SecurityElement]::Escape($Code))</ns:Code>
  </ns:VendorInfo>
</ns:VendorSearchInfo>
"@
  try {
    $xml = Invoke-AeosSoap -SvcUrl $SvcUrl -InnerXml $body -User $User -Password $Password
    if ([string]::IsNullOrWhiteSpace($xml)) { return @{ Exists=$false; Id=$null; Name=$null } }
    [xml]$doc = $xml
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("s","http://schemas.xmlsoap.org/soap/envelope/")
    $ns.AddNamespace("a","http://www.nedap.com/aeosws/schema")
    $vendorNode = $doc.SelectSingleNode("//a:Vendor", $ns)
    if ($vendorNode -ne $null) {
      $idNode   = $vendorNode.SelectSingleNode("a:Id", $ns)
      $nameNode = $vendorNode.SelectSingleNode("a:Name", $ns)
      $id   = if ($idNode) { $idNode.InnerText } else { $null }
      $name = if ($nameNode){ $nameNode.InnerText } else { $null }
      return @{ Exists=$true; Id=$id; Name=$name }
    }
    return @{ Exists=$false; Id=$null; Name=$null }
  } catch {
    Write-Log ("SOAP findVendor error for {0}: {1}" -f $Code,$_.Exception.Message) "WARN"
    return @{ Exists=$false; Id=$null; Name=$null }
  }
}
function AddOrChange-AeosVendor {
  param(
    [Parameter(Mandatory=$true)][string]$Code,
    [Parameter(Mandatory=$true)][string]$Name,
    [Parameter(Mandatory=$true)][string]$SvcUrl,
    [string]$Category="P",
    [string]$Place="Geneva",
    [string]$User,[string]$Password
  )
  $info = Find-AeosVendorByCode -Code $Code -SvcUrl $SvcUrl -User $User -Password $Password
  if (-not $info.Exists) {
    # ADD vendor
    $body = @"
<ns:VendorAdd xmlns:ns="http://www.nedap.com/aeosws/schema">
  <ns:Name>$([System.Security.SecurityElement]::Escape($Name))</ns:Name>
  <ns:Code>$([System.Security.SecurityElement]::Escape($Code))</ns:Code>
  <ns:Category>$([System.Security.SecurityElement]::Escape($Category))</ns:Category>
  <ns:Place>$([System.Security.SecurityElement]::Escape($Place))</ns:Place>
  <ns:CompanyName>$([System.Security.SecurityElement]::Escape($Name))</ns:CompanyName>
</ns:VendorAdd>
"@
    try {
      [void](Invoke-AeosSoap -SvcUrl $SvcUrl -InnerXml $body -User $User -Password $Password)
      Write-Log ("SOAP addVendor {0} OK (name='{1}', cat='{2}', place='{3}')" -f $Code,$Name,$Category,$Place) "INFO"
      return $true
    } catch {
      Write-Log ("SOAP addVendor failed for {0} : {1}" -f $Code,$_.Exception.Message) "ERROR"
      return $false
    }
  } else {
    # CHANGE vendor (force Name/CompanyName/Place)
    $id = $info.Id
    if ([string]::IsNullOrWhiteSpace($id)) {
      Write-Log ("SOAP changeVendor skipped for {0}: missing Id" -f $Code) "WARN"
      return $false
    }
    $body = @"
<ns:VendorChange xmlns:ns="http://www.nedap.com/aeosws/schema">
  <ns:Id>$id</ns:Id>
  <ns:Name>$([System.Security.SecurityElement]::Escape($Name))</ns:Name>
  <ns:Category>$([System.Security.SecurityElement]::Escape($Category))</ns:Category>
  <ns:Place>$([System.Security.SecurityElement]::Escape($Place))</ns:Place>
  <ns:CompanyName>$([System.Security.SecurityElement]::Escape($Name))</ns:CompanyName>
</ns:VendorChange>
"@
    try {
      [void](Invoke-AeosSoap -SvcUrl $SvcUrl -InnerXml $body -User $User -Password $Password)
      Write-Log ("SOAP changeVendor {0} OK (id={1}, company='{2}', cat='{3}', place='{4}')" -f $Code,$id,$Name,$Category,$Place) "INFO"
      return $true
    } catch {
      Write-Log ("SOAP changeVendor failed for {0} : {1}" -f $Code,$_.Exception.Message) "ERROR"
      return $false
    }
  }
}

# ======================
# READ SOURCE FILE
# ======================
$src = Join-Path $config.FolderPathImport $config.InputFileName
if (-not (Test-Path $src)) { Write-Log ("Source file not found: {0}" -f $src) "WARN"; Close-Logger; exit 0 }
Write-KV "Source file" $src

$bytes   = [System.IO.File]::ReadAllBytes($src)
$content = Decode-FileContent $bytes
$lines   = @($content -split "(`r`n|`n|`r)")
if ($lines.Length -gt 0) { $lines[0] = $lines[0].TrimStart([char]0xFEFF) }
if ($lines.Length -lt 2) { throw "Fichier vide ou sans donnÃ©es" }

$sep = Detect-Separator $lines[0]
$headers = Parse-CsvLine $lines[0] $sep
if ($headers.Length -le 1) {
  if ($sep -eq ',') { $sep = ';' } else { $sep = ',' }
  $headers = Parse-CsvLine $lines[0] $sep
  Write-KV "Separator (fallback)" ([string]$sep)
} else {
  Write-KV "Separator" ([string]$sep)
}

Write-Section "Header"
Write-KV "Columns" ($headers.Length.ToString())
Write-KV "Header line" ($headers -join " | ")

$hmap = Build-HeaderMap $headers
Write-Section "Header mapping"
foreach($k in @("matricule","noms","prenoms","societe","reponse","echeance","entreprise","badgeyn","remarques")){
  if ($hmap.ContainsKey($k)) { Write-KV $k ([string]$hmap[$k]) }
}

# ======================
# INTROSPECT TARGET TABLE SCHEMA + BUILD MATCHING DATATABLE
# ======================
$cn = New-Object System.Data.SqlClient.SqlConnection (Get-ConnString)
$cn.Open()
$batchStart = Get-Date
try {
  # ===========================
  # Nettoyage de la table staging
  # ===========================
  Write-Section "Staging cleanup"
  $cmd = $cn.CreateCommand()
  $cmd.CommandText = "TRUNCATE TABLE dbo.PJ_PRESTATAIRES_IMPORT_STAGE;"
  $cmd.CommandType  = [System.Data.CommandType]::Text
  [void]$cmd.ExecuteNonQuery()
  Write-KV "PJ_PRESTATAIRES_IMPORT_STAGE" "Truncated"

  # ===========================
  # Lecture du schÃ©ma de la staging
  # ===========================
  $schema = @()
  $cmd = $cn.CreateCommand()
  $cmd.CommandText = @"
SELECT c.name AS ColumnName,
       t.name AS SqlType,
       c.max_length AS MaxLen,
       c.is_nullable AS IsNullable
FROM sys.columns c
JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.PJ_PRESTATAIRES_IMPORT_STAGE')
ORDER BY c.column_id;
"@

  $r = $cmd.ExecuteReader()
  while ($r.Read()) {
    $row = [PSCustomObject]@{
      ColumnName = $r.GetString(0)
      SqlType    = $r.GetString(1)
      MaxLen     = $r.GetInt16(2)
      IsNullable = $r.GetBoolean(3)
    }
    $schema += $row
  }
  $r.Close()
  if ($schema.Count -eq 0) { throw "Table dbo.PJ_PRESTATAIRES_IMPORT_STAGE introuvable." }

  $dt = New-Object System.Data.DataTable "PJ_PRESTATAIRES_IMPORT_STAGE"
  $colDefs = @{}
  foreach($s in $schema){
    $dotnetType = [type]"System.String"
    switch -Regex ($s.SqlType.ToLower()) {
      'int'              { $dotnetType = [type]"System.Int32" ; break }
      'bigint'           { $dotnetType = [type]"System.Int64" ; break }
      'smallint'         { $dotnetType = [type]"System.Int16" ; break }
      'bit'              { $dotnetType = [type]"System.Boolean" ; break }
      'datetime2|datetime|smalldatetime' { $dotnetType = [type]"System.DateTime" ; break }
      'uniqueidentifier' { $dotnetType = [type]"System.Guid" ; break }
      default            { $dotnetType = [type]"System.String" ; break }
    }
    [void]$dt.Columns.Add($s.ColumnName, $dotnetType)
    $maxChars = if ($s.SqlType -match 'nchar|nvarchar') { [int]([math]::Floor($s.MaxLen/2)) } else { $s.MaxLen }
    if ($maxChars -le 0) { $maxChars = 4000 }
    $colDefs[$s.ColumnName] = [PSCustomObject]@{ SqlType=$s.SqlType.ToLower(); MaxChars=$maxChars; IsNullable=$s.IsNullable }
  }

  $batchId = [guid]::NewGuid()
  $expected = $headers.Length
  $logged=0
  $vendorsFromAccredit = @{}

  for($i=1; $i -lt $lines.Length; $i++){
    $raw = $lines[$i]; if ([string]::IsNullOrWhiteSpace($raw)) { continue }
    $cols = Parse-CsvLine $raw $sep
    if ($cols.Length -lt $expected) { $cols = $cols + (New-Object string[] ($expected - $cols.Length)) }

    $pernr      = $cols[$hmap["matricule"]]; if ([string]::IsNullOrWhiteSpace($pernr)) { continue }
    $noms       = $cols[$hmap["noms"]]
    $prenoms    = $cols[$hmap["prenoms"]]
    $socLibRaw  = $cols[$hmap["societe"]]
    $dfrom      = $cols[$hmap["reponse"]]
    $dto        = $cols[$hmap["echeance"]]
    $vendorCode = $cols[$hmap["entreprise"]]
    $badgeYN    = if ($hmap.ContainsKey("badgeyn")) { $cols[$hmap["badgeyn"]] } else { $null }

    # Nouvelle colonne : Remarques (-> champ libre AEOS "Huissiers / Nom Etude" = freefield "Nom Etude" (description AEOS))
    $remRaw     = if ($hmap.ContainsKey("remarques")) { $cols[$hmap["remarques"]] } else { $null }
    $nomEtude   = Clean-Remarques $remRaw

    if (-not [string]::IsNullOrWhiteSpace($nomEtude) -and $nomEtude.Length -gt 30) { $nomEtude = $nomEtude.Substring(0,30) }
    # Construction des freefields Ã  passer Ã  AEOS via dbo.import (multi-valeurs sÃ©parÃ©es par '|')
    # IMPORTANT : pour AEOS, freefieldid doit contenir l'ID "fonctionnel" (description/nom dans l'application),
    # pas l'objectid SQL. Ici : "Nom Etude".
    $ffId = $null
    $ffData = $null
    $ids = New-Object System.Collections.Generic.List[string]
    $datas = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($nomEtude))  { $ids.Add('Nom Etude') | Out-Null; $datas.Add($nomEtude) | Out-Null }

    if ($ids.Count -gt 0) {
      $ffId   = ($ids.ToArray() -join '|')
      $ffData = ($datas.ToArray() -join '|')
    }

$pdFrom = Parse-Date $dfrom
    $pdTo   = Parse-Date $dto
    if ($pdTo -and -not $pdFrom) { $pdFrom = $pdTo.AddYears(-1) }
    if ($logged -lt 2 -and ($pdFrom -or $pdTo)) {
      $sFrom = if ($pdFrom) { $pdFrom.ToString("yyyy-MM-dd") } else { "NULL" }
      $sTo   = if ($pdTo)   { $pdTo.ToString("yyyy-MM-dd") }   else { "NULL" }
      Write-Log ("Line {0} -> Start: {1} ; End: {2}" -f ($i+1), $sFrom, $sTo)
      $logged++
    }

    $socLibText    = [string]$socLibRaw
    $socLib        = if ([string]::IsNullOrWhiteSpace($socLibText)) { $null } else { $socLibText.Trim() }
    $vendorCodeText = [string]$vendorCode
    $vendorCode    = if ([string]::IsNullOrWhiteSpace($vendorCodeText)) { $null } else { $vendorCodeText.Trim() }
    if ($vendorCode -and ($vendorCode -match '^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$')) {
      Write-Log ("Line {0}: vendor code '{1}' ressemble Ã  une date -> ignorÃ©." -f ($i+1), $vendorCode) "WARN"
      $vendorCode = $null
    }
    $vendorName = $socLib
    if ([string]::IsNullOrWhiteSpace($vendorName) -and $vendorCode) { $vendorName = $vendorCode }
    if ($vendorCode) {
      if (-not $vendorsFromAccredit.ContainsKey($vendorCode)) { $vendorsFromAccredit[$vendorCode] = $vendorName }
      else {
        if (-not [string]::IsNullOrWhiteSpace($vendorsFromAccredit[$vendorCode])) {
          $vendorName = $vendorsFromAccredit[$vendorCode]
        }
      }
    }
    if (-not $vendorCode) {
      Write-Log ("Line {0}: vendor code missing -> row skipped." -f ($i+1)) "WARN"
      continue
    }

    $row = $dt.NewRow()
    foreach($cName in $colDefs.Keys){
      switch ($cName) {
        'batch_id'      { $row[$cName] = $batchId ; break }
        'load_ts'       { $row[$cName] = [DateTime]::UtcNow ; break }
        'source_file'   { $row[$cName] = [System.IO.Path]::GetFileName($src) ; break }
        'personnelnr'   { Set-Str -Row $row -ColumnName $cName -Value $pernr -ColDefs $colDefs ; break }
        'lastname'      { Set-Str -Row $row -ColumnName $cName -Value $noms -ColDefs $colDefs ; break }
        'firstname'     { Set-Str -Row $row -ColumnName $cName -Value $prenoms -ColDefs $colDefs ; break }
        'initials'      { Set-Str -Row $row -ColumnName $cName -Value $prenoms -ColDefs $colDefs ; break }

        # Dates provenant du fichier prestataires :
        #   RÃ©ponse  -> $pdFrom
        #   Ã‰chÃ©ance -> $pdTo
        'arrivaldate'   { if ($pdFrom) { $row[$cName] = $pdFrom } else { $row[$cName] = [DBNull]::Value } ; break }
        'leavedate'     { if ($pdTo)   { $row[$cName] = $pdTo   } else { $row[$cName] = [DBNull]::Value } ; break }
        'validfrom'     { $row[$cName] = [DBNull]::Value ; break }
        'validto'       { $row[$cName] = [DBNull]::Value ; break }

        'company'       { Set-Str -Row $row -ColumnName $cName -Value $vendorName -ColDefs $colDefs ; break }
        'company_code'  { Set-Str -Row $row -ColumnName $cName -Value $vendorCode -ColDefs $colDefs ; break }
        'vendor_code'   { Set-Str -Row $row -ColumnName $cName -Value $vendorCode -ColDefs $colDefs ; break }
        'vendor'        { Set-Str -Row $row -ColumnName $cName -Value $vendorName -ColDefs $colDefs ; break }
        'societe_lbl'   { Set-Str -Row $row -ColumnName $cName -Value $vendorName -ColDefs $colDefs ; break }
        'badgeyn'       { Set-Str -Row $row -ColumnName $cName -Value $badgeYN -ColDefs $colDefs ; break }
        'remarques'    { Set-Str -Row $row -ColumnName $cName -Value $nomEtude -ColDefs $colDefs ; break }
        'freefieldid'   { Set-Str -Row $row -ColumnName $cName -Value $ffId   -ColDefs $colDefs ; break }
        'freefielddata' { Set-Str -Row $row -ColumnName $cName -Value $ffData -ColDefs $colDefs ; break }
        'disabled'      { $row[$cName] = 0 ; break }
        'carriertype'   { $row[$cName] = 5 ; break }
        default         { $row[$cName] = [DBNull]::Value ; break }
      }
    }
    $dt.Rows.Add($row)
  }

  Write-Section "Parsed rows"
  Write-KV "Valid rows" ($dt.Rows.Count.ToString())
  if ($dt.Rows.Count -eq 0) { throw "Aucune ligne exploitable" }

  # Persons imported summary (from staging)
  Write-Section "Persons imported (function 9)"
  Write-KV "Persons total" ($dt.Rows.Count.ToString())
  foreach($row in $dt.Rows){
    $pernr      = $null
    $lastname   = $null
    $firstname  = $null
    $company    = $null
    $vendorCode = $null

    if ($dt.Columns.Contains("personnelnr")) { $pernr = $row["personnelnr"] }
    elseif ($dt.Columns.Contains("PERNR"))   { $pernr = $row["PERNR"] }

    if ($dt.Columns.Contains("lastname"))  { $lastname  = $row["lastname"] }
    if ($dt.Columns.Contains("firstname")) { $firstname = $row["firstname"] }

    if ($dt.Columns.Contains("company"))      { $company    = $row["company"] }
    if ($dt.Columns.Contains("vendor_code"))  { $vendorCode = $row["vendor_code"] }
    elseif ($dt.Columns.Contains("company_code")) { $vendorCode = $row["company_code"] }

    Write-Log ("IMPORT person {0} ({1} {2}, company='{3}', vendor='{4}')" -f `
      $pernr, $lastname, $firstname, $company, $vendorCode) "INFO"
  }

  # BULKCOPY
  Write-Section "Staging load"
  $bulk = New-Object System.Data.SqlClient.SqlBulkCopy($cn, [System.Data.SqlClient.SqlBulkCopyOptions]::KeepNulls, $null)
  $bulk.DestinationTableName = "dbo.PJ_PRESTATAIRES_IMPORT_STAGE"
  $bulk.BatchSize = 1000
  $bulk.BulkCopyTimeout = 600
  foreach($s in $schema){ [void]$bulk.ColumnMappings.Add($s.ColumnName, $s.ColumnName) }
  try {
    $bulk.WriteToServer($dt)
    Write-KV "PJ_PRESTATAIRES_IMPORT_STAGE" "BulkCopy OK"
  } catch {
    Write-Log ("BulkCopy ERROR: " + $_.Exception.Message) "ERROR"
    foreach($col in $dt.Columns) { Write-Log ("DT {0}: {1}" -f $col.ColumnName, $col.DataType.FullName) "WARN" }
    foreach($s in $schema){ Write-Log ("DB {0}: {1}({2})" -f $s.ColumnName, $s.SqlType, $s.MaxLen) "WARN" }
    throw
  }

  # ======================
  # VENDORS : SOAP ADD/CHANGE + DB UNBLOCK
  # ======================
  $soapEnabled = $false
  if ($config.EnableVendorSoap -eq $true) {
    $svc = Get-VendorServiceUrl -Base $config.SoapBaseUrl
    $sc  = Get-SoapCredentials
    if ($svc -and $sc.User -and $sc.Password) {
      $soapEnabled = $true
      Write-Log "SOAP Vendor mode ON" "INFO"
      Enable-SoapInsecureTlsIfRequested
    } else {
      Write-Log "SOAP Vendor mode requested but incomplete settings; fallback to DB-only checks." "WARN"
    }
  } else {
    Write-Log "SOAP Vendor mode OFF" "INFO"
  }

  # 1) Existing vendors + blocked flags
  $existing = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
  $blocked  = @{}
  $cmd = $cn.CreateCommand()
  $cmd.CommandText = @"
SELECT [code],
       CASE
         WHEN [blocked] IS NULL THEN 0
         WHEN UPPER(LTRIM(RTRIM(CONVERT(nvarchar(50),[blocked])))) IN ('1','Y','YES','TRUE') THEN 1
         WHEN UPPER(LTRIM(RTRIM(CONVERT(nvarchar(50),[blocked])))) IN ('0','N','NO','FALSE') THEN 0
         ELSE 0
       END AS blocked
FROM dbo.vendor WITH (NOLOCK)
"@
  $r = $cmd.ExecuteReader()
  while ($r.Read()) {
    $code = $r.GetString(0)
    [void]$existing.Add($code)
    $val = 0
    try { $val = $r.GetInt32(1) } catch { try { $val = [int]$r.GetValue(1) } catch { $val = 0 } }
    $blocked[$code] = $val
  }
  $r.Close()

  $codesAll = @()
  foreach($k in $vendorsFromAccredit.Keys){ if ($k) { $codesAll += $k } }
  $codesAll = @($codesAll | Sort-Object -Unique)

  $missing = @()
  foreach($c2 in $codesAll){ if (-not $existing.Contains($c2)) { $missing += $c2 } }

  Write-Section "Vendors (file vs DB)"
  Write-KV "Codes unique" ($codesAll.Count.ToString())
  Write-KV "Missing in DB" ($missing.Count.ToString())
  if ($missing.Count -gt 0) {
    $sample = @($missing | Select-Object -First 10)
    $sampleText = $sample -join ", "
    if ($missing.Count -gt 10) { $sampleText = $sampleText + (" ... (+{0} more)" -f ($missing.Count - 10)) }
    Write-KV "Missing sample" $sampleText
  }

  # 2) SOAP add only missing vendors (based on DB)
  $soapAdded = 0
  $soapFailed = 0
  $soapSkipped = 0
  if ($soapEnabled -and $missing.Count -gt 0) {
    $place = if ($config.VendorPlaceOfBusiness) { $config.VendorPlaceOfBusiness } else { "Geneva" }
    $category = if ($config.VendorDefaultCategory) { $config.VendorDefaultCategory } else { "P" }
    foreach($code in $missing){
      $name = $vendorsFromAccredit[$code]
      $ok = AddOrChange-AeosVendor -Code $code -Name $name -SvcUrl $svc -Category $category -Place $place -User $sc.User -Password $sc.Password
      if ($ok) { $soapAdded++ } else { $soapFailed++ }
    }
  } elseif ($soapEnabled -and $missing.Count -eq 0 -and $codesAll.Count -gt 0) {
    $soapSkipped = $codesAll.Count
    Write-Log "All vendors already in DB; SOAP add skipped." "INFO"
  }
  if ($soapEnabled -and $codesAll.Count -gt 0) {
    if ($soapSkipped -eq 0) { $soapSkipped = $codesAll.Count - $missing.Count }
    Write-Section "Vendor SOAP summary"
    Write-KV "Skipped (already in DB)" ($soapSkipped.ToString())
    Write-KV "Added" ($soapAdded.ToString())
    Write-KV "Failed" ($soapFailed.ToString())
  }

  # 3) UNBLOCK vendors in DB if blocked=1
  if ($codesAll.Count -gt 0) {
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "UPDATE v SET v.blocked = 0, v.removaldate = NULL FROM dbo.vendor v WHERE v.[code] = @code"
    $p = $cmd.Parameters.Add("@code",[System.Data.SqlDbType]::NVarChar,50)
    $done = $false
    try {
      foreach($c3 in $codesAll){
        $p.Value = $c3
        [void]$cmd.ExecuteNonQuery()
      }
      $done = $true
    } catch {
      Write-Log ("Vendor DB unblock (bit mode) failed, retrying as char(1): {0}" -f $_.Exception.Message) "WARN"
      try {
        $cmd.CommandText = "UPDATE v SET v.blocked = N'N', v.removaldate = NULL FROM dbo.vendor v WHERE v.[code] = @code"
        foreach($c3 in $codesAll){
          $p.Value = $c3
          [void]$cmd.ExecuteNonQuery()
        }
        $done = $true
      } catch {
        Write-Log ("Vendor DB unblock skipped (update failed): {0}" -f $_.Exception.Message) "WARN"
      }
    }
    if ($done) { Write-Log "Vendor DB unblock applied." "INFO" }
  }

  # 4) Poll until all vendors appear
  if ($codesAll.Count -gt 0) {
    $maxWait = 180
    if ($config.VendorWaitTimeoutSeconds) { $maxWait = [int]$config.VendorWaitTimeoutSeconds }
    $every   = 5
    if ($config.VendorPollEverySeconds)   { $every   = [int]$config.VendorPollEverySeconds }

    $start = Get-Date
    while ((Get-Date) -lt $start.AddSeconds($maxWait)) {
      $allOk = $true
      $cmd = $cn.CreateCommand()
      $cmd.CommandText = "SELECT [blocked] FROM dbo.vendor WITH (NOLOCK) WHERE [code] = @code"
      $p2 = $cmd.Parameters.Add("@code",[System.Data.SqlDbType]::NVarChar,50)
      foreach($c4 in $codesAll){
        $p2.Value = $c4
        $val = $cmd.ExecuteScalar()
        if ($null -eq $val) { $allOk = $false; break }
        $isBlocked = 0
        try { $isBlocked = [int]$val } catch { $isBlocked = 0 }
        if ($isBlocked -ne 0) { $allOk = $false; break }
      }
      if ($allOk) { break }
      Start-Sleep -Seconds $every
    }
  }

 # ======================
# APPLY ACCREDITATIONS + READ RunId/RunTs
# ======================
Write-Section "Apply accreditations"
$spName = "dbo.sp_LoadAccreditesToImport"
if ($config.StoredProcedureToApply) { $spName = $config.StoredProcedureToApply }

$cmd = $cn.CreateCommand()
$cmd.CommandText = $spName
$cmd.CommandType = [System.Data.CommandType]::StoredProcedure

# On lit le SELECT final de la SP (RunId/RunTs)
$ds = New-Object System.Data.DataSet
$da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)

$runId = $null
$runTs = $null

try {
  [void]$da.Fill($ds)
  Write-KV "SP result" "Executed"

  if ($ds.Tables.Count -gt 0 -and $ds.Tables[0].Rows.Count -gt 0) {
    $t0 = $ds.Tables[0]
    $r0 = $t0.Rows[0]

    if ($t0.Columns.Contains("RunId")) { $runId = [string]$r0["RunId"] }
    if ($t0.Columns.Contains("RunTs")) { $runTs = [string]$r0["RunTs"] }

    Write-Section "Import summary"
    if ($runId) { Write-KV "RunId" $runId } else { Write-KV "RunId" "NULL" }
    if ($runTs) { Write-KV "RunTs" $runTs } else { Write-KV "RunTs" "NULL" }
  } else {
    Write-Log "SP executed but returned no row (no RunId/RunTs)." "WARN"
  }
} catch {
  Write-Log ("SP execution failed: " + $_.Exception.Message) "ERROR"
  throw
}
$histBlocks = $null
# ======================
# READ HIST COUNTERS (RowsInStage / PersonsQueued / BlocksQueued)
# ======================
if ($runId) {
  try {
    $cmdH = $cn.CreateCommand()
    $cmdH.CommandType = [System.Data.CommandType]::Text
    $cmdH.CommandText = @"
SELECT TOP 1
  RowsInStage, PersonsQueued, BlocksQueued, SourceFile, BatchId, Notes
FROM dbo.PJ_PRESTATAIRES_IMPORT_HIST WITH (NOLOCK)
WHERE RunId = @runid
ORDER BY RunTs DESC;
"@
    $pH = $cmdH.Parameters.Add("@runid",[System.Data.SqlDbType]::UniqueIdentifier)
    $pH.Value = [Guid]$runId

    $dtH = New-Object System.Data.DataTable
    $daH = New-Object System.Data.SqlClient.SqlDataAdapter($cmdH)
    [void]$daH.Fill($dtH)

    if ($dtH.Rows.Count -gt 0) {
      $h = $dtH.Rows[0]
      Write-KV "RowsInStage"   ([string]$h["RowsInStage"])
      Write-KV "PersonsQueued" ([string]$h["PersonsQueued"])
      Write-KV "BlocksQueued"  ([string]$h["BlocksQueued"])
      try { $histBlocks = [int]$h["BlocksQueued"] } catch { $histBlocks = $null }
      Write-KV "SourceFile"    ([string]$h["SourceFile"])
      Write-KV "BatchId"       ([string]$h["BatchId"])
      if ($h.Table.Columns.Contains("Notes")) {
        $notes = $h["Notes"]
        if ($notes -ne $null -and $notes -isnot [DBNull]) { Write-KV "Notes" ([string]$notes) }
      }
    } else {
      Write-Log ("No HIST row found for RunId={0}" -f $runId) "WARN"
    }
  } catch {
    Write-Log ("HIST read failed: " + $_.Exception.Message) "WARN"
  }
}
# ======================
# READ AUDIT DETAILS (UPSERT / BLOCK) + LOG
# ======================
if ($runId) {
  try {
    Write-Section "Audit details"
    $cmdA = $cn.CreateCommand()
    $cmdA.CommandType = [System.Data.CommandType]::Text
    $cmdA.CommandText = @"
SELECT ActionType, PersonnelNr, LastName, FirstName, VendorCode
FROM dbo.PJ_PRESTATAIRES_IMPORT_AUDIT WITH (NOLOCK)
WHERE RunId = @runid
ORDER BY AuditId;
"@
    $pA = $cmdA.Parameters.Add("@runid",[System.Data.SqlDbType]::UniqueIdentifier)
    $pA.Value = [Guid]$runId

    $rdr = $cmdA.ExecuteReader()
    $cntUpsert = 0
    $cntBlock  = 0

    while ($rdr.Read()) {
      $action = $rdr.GetString(0)
      $pernr  = $rdr.GetString(1)
      $ln     = if ($rdr.IsDBNull(2)) { "" } else { $rdr.GetString(2) }
      $fn     = if ($rdr.IsDBNull(3)) { "" } else { $rdr.GetString(3) }
      $vend   = if ($rdr.IsDBNull(4)) { "" } else { $rdr.GetString(4) }

      if ($action -eq "UPSERT" -or $action -eq "CREATE" -or $action -eq "UPDATE") { $cntUpsert++ }
      elseif ($action -eq "BLOCK") { $cntBlock++ }

      Write-Log ("{0} | {1} | {2} {3} | vendor={4}" -f $action,$pernr,$ln,$fn,$vend) "INFO"
    }
    $rdr.Close()

    if ($cntBlock -eq 0 -and $histBlocks -ne $null) { $cntBlock = $histBlocks }
    Write-Section "Audit totals"
    Write-KV "UPSERT" ([string]$cntUpsert)
    Write-KV "BLOCK"  ([string]$cntBlock)

  } catch {
    try { if ($rdr) { $rdr.Close() } } catch {}
    Write-Log ("AUDIT read failed: " + $_.Exception.Message) "WARN"
  }
}

  # ARCHIVE
  Write-Section "Archive"
  $delay = 0
  if ($config.DelayBeforeMoveSeconds) { $delay = [int]$config.DelayBeforeMoveSeconds }
  if ($delay -gt 0) { Write-KV "Delay (s)" ([string]$delay); Start-Sleep -Seconds $delay }
  $stamp = Get-Date -Format "HH'h'mm-ddMMyyyy"
  $dest  = Join-Path $config.FolderPathOutput ("prestataires_{0}.csv" -f $stamp)
  try {
    Move-Item -LiteralPath $src -Destination $dest -Force
    Write-KV "Archived file" $dest
  } catch {
    Write-Log ("Archivage impossible: {0}" -f $_.Exception.Message) "WARN"
  }
}
finally {
  if ($cn.State -eq "Open") { $cn.Close() }
  Write-Log "Done." "INFO"
  Close-Logger
}
exit 0

