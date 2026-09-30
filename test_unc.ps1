$paths = @(
    "\\MORION_MD\Users\Morion\Projects\excel-supertools-web-addin - 1",
    "\\localhost\Users\Morion\Projects\excel-supertools-web-addin - 1",
    "\\127.0.0.1\Users\Morion\Projects\excel-supertools-web-addin - 1"
)

foreach ($p in $paths) {
    Write-Host "Testing $p ..."
    if (Test-Path $p) {
        Write-Host "  -> EXISTS! Accessible!"
        $xmls = Get-ChildItem -Path $p -Filter "*.xml"
        foreach ($x in $xmls) {
            Write-Host "     Found: $($x.Name)"
        }
    } else {
        Write-Host "  -> NOT ACCESSIBLE"
    }
}
