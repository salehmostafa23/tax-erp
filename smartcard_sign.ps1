param([string]$DataB64)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Security

function Find-Eseal {
  foreach ($store in @("Cert:\CurrentUser\My","Cert:\LocalMachine\My")) {
    if (Test-Path $store) {
      $c = Get-ChildItem $store -ErrorAction SilentlyContinue |
           Where-Object { $_.HasPrivateKey -and $_.Issuer -like "*Egypt Trust*" -and $_.NotAfter -gt (Get-Date) } |
           Select-Object -First 1
      if ($c) { return $c }
    }
  }
  return $null
}

$cert = Find-Eseal
if (-not $cert) {
  Write-Output 'NO_CERT'
  exit 2
}

$res = @{
  cert   = [Convert]::ToBase64String($cert.RawData)
  issuer = [Convert]::ToBase64String($cert.IssuerName.RawData)
  serial = $cert.SerialNumber
}

if ($DataB64) {
  $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
  $bytes = [Convert]::FromBase64String($DataB64)
  $sig = $rsa.SignData($bytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                       [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
  $res.signature = [Convert]::ToBase64String($sig)
}

Write-Output ($res | ConvertTo-Json -Compress)