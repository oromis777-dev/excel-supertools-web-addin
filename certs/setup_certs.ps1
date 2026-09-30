$certsDir = "c:\Users\Morion\Projects\excel-supertools-web-addin - 1\certs"
if (-not (Test-Path $certsDir)) {
    New-Item -ItemType Directory -Path $certsDir -Force | Out-Null
}

$pfxPath = "$certsDir\localhost.pfx"
$pwd = ConvertTo-SecureString -String "exceldev" -Force -AsPlainText

$cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(5)
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pwd | Out-Null

Write-Host "Certificate exported to $pfxPath"
