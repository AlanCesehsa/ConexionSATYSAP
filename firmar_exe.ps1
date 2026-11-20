$cert = Get-Item "Cert:\CurrentUser\My\F2561E970D87181A51F413E112CC9E06B8148D45"
Write-Host "Certificado encontrado: $($cert.Subject)"

$exePath = "C:\Users\Stock CESEHSA\Documents\Antigravity\ConexionSATYSAP\Conciliacion_SAP_vs_SAT.exe"
$distPath = "C:\Users\Stock CESEHSA\Documents\Antigravity\ConexionSATYSAP\dist\Conciliacion_SAP_vs_SAT.exe"

Write-Host "Firmando $exePath..."
$sig1 = Set-AuthenticodeSignature -Certificate $cert -FilePath $exePath -TimestampServer "http://timestamp.digicert.com"
Write-Host "Resultado: $($sig1.Status) - $($sig1.StatusMessage)"

Write-Host "Firmando $distPath..."
$sig2 = Set-AuthenticodeSignature -Certificate $cert -FilePath $distPath -TimestampServer "http://timestamp.digicert.com"
Write-Host "Resultado: $($sig2.Status) - $($sig2.StatusMessage)"

Write-Host "`nVerificando firma final:"
Get-AuthenticodeSignature $exePath | Format-List

