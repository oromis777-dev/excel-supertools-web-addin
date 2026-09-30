$wefDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
$files = Get-ChildItem $wefDir -Recurse -File
Write-Host "Total files in WEF: $($files.Count)"
foreach ($f in $files) {
    Write-Host "`n=== FILE: $($f.FullName) (Size: $($f.Length)) ==="
    if ($f.Length -lt 20000 -and ($f.Extension -in @(".json", ".xml", ".txt", ".log", "") -or $f.Name -match "meta")) {
        try {
            Get-Content $f.FullName -Raw | Write-Host
        } catch {
            Write-Host "Could not read text: $_"
        }
    }
}
