$certsDir = "c:\Users\Morion\Projects\excel-supertools-web-addin - 1\certs"
$pfxPath = "$certsDir\localhost.pfx"
$cerPath = "$certsDir\localhost.cer"

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath, "exceldev")
[System.IO.File]::WriteAllBytes($cerPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
Write-Host "Exported CER to $cerPath with Thumbprint: $($cert.Thumbprint)"
