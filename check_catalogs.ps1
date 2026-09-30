$catalogs = Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs' -ErrorAction SilentlyContinue
Write-Host "FOUND CATALOGS: $($catalogs.Count)"
foreach ($cat in $catalogs) {
    $props = Get-ItemProperty $cat.PSPath
    Write-Host "Catalog: $($cat.PSChildName)"
    Write-Host "  Id:    $($props.Id)"
    Write-Host "  Url:   $($props.Url)"
    Write-Host "  Flags: $($props.Flags)"
    if ($props.Url) {
        $folderPath = $props.Url
        Write-Host "  Checking path: $folderPath"
        if (Test-Path $folderPath) {
            Write-Host "  Path exists: YES"
            $files = Get-ChildItem -Path $folderPath -Filter "*.xml" -ErrorAction SilentlyContinue
            Write-Host "  XML files in folder: $($files.Count)"
            foreach ($f in $files) {
                Write-Host "    - $($f.Name) (Size: $($f.Length) bytes)"
            }
        } else {
            Write-Host "  Path exists: NO (Folder not accessible or not found)"
        }
    }
}

Write-Host "`nChecking Developer Key:"
$devKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (Test-Path $devKey) {
    Get-ChildItem $devKey | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath
        Write-Host "  Developer Addin: $($_.PSChildName)"
        Write-Host "    ManifestPath: $($p.ManifestPath)"
    }
} else {
    Write-Host "  No Developer key found."
}

Write-Host "`nChecking WEF Cache:"
$wefCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefCache) {
    Write-Host "  Wef Cache exists at $wefCache"
    $cacheItems = Get-ChildItem $wefCache
    Write-Host "  Items in Wef cache: $($cacheItems.Count)"
}
