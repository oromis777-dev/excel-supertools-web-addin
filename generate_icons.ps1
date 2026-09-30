Add-Type -AssemblyName System.Drawing

$assetsDir = "c:\Users\Morion\Projects\excel-supertools-web-addin - 1\assets"
if (-not (Test-Path $assetsDir)) {
    New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null
}

$sizes = @(16, 32, 80)

function Draw-Ai([System.Drawing.Graphics]$g, [int]$s) {
    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#107C41"))
    $accent = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#2B88D8"))
    $center = $s / 2.0
    $r = $s * 0.42
    
    $p1 = New-Object System.Drawing.PointF ($center, ($center - $r))
    $p2 = New-Object System.Drawing.PointF (($center + $r * 0.25), ($center - $r * 0.25))
    $p3 = New-Object System.Drawing.PointF (($center + $r), $center)
    $p4 = New-Object System.Drawing.PointF (($center + $r * 0.25), ($center + $r * 0.25))
    $p5 = New-Object System.Drawing.PointF ($center, ($center + $r))
    $p6 = New-Object System.Drawing.PointF (($center - $r * 0.25), ($center + $r * 0.25))
    $p7 = New-Object System.Drawing.PointF (($center - $r), $center)
    $p8 = New-Object System.Drawing.PointF (($center - $r * 0.25), ($center - $r * 0.25))
    
    $pts = [System.Drawing.PointF[]]@($p1, $p2, $p3, $p4, $p5, $p6, $p7, $p8)
    $g.FillPolygon($brush, $pts)
    $g.FillEllipse($accent, ($center - $r * 0.25), ($center - $r * 0.25), ($r * 0.5), ($r * 0.5))
    
    $brush.Dispose()
    $accent.Dispose()
}

function Draw-Grid([System.Drawing.Graphics]$g, [int]$s) {
    $pen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#107C41")), ([Math]::Max(1.0, $s * 0.08))
    $hlBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#E83E8C"))
    $margin = $s * 0.12
    $rectW = $s - 2 * $margin
    $step = $rectW / 3.0
    
    # Highlight crosshair
    $g.FillRectangle($hlBrush, $margin, ($margin + $step), $rectW, $step)
    $g.FillRectangle($hlBrush, ($margin + $step), $margin, $step, $rectW)
    
    # Outer frame & lines
    $g.DrawRectangle($pen, $margin, $margin, $rectW, $rectW)
    $g.DrawLine($pen, $margin, ($margin + $step), ($margin + $rectW), ($margin + $step))
    $g.DrawLine($pen, $margin, ($margin + 2 * $step), ($margin + $rectW), ($margin + 2 * $step))
    $g.DrawLine($pen, ($margin + $step), $margin, ($margin + $step), ($margin + $rectW))
    $g.DrawLine($pen, ($margin + 2 * $step), $margin, ($margin + 2 * $step), ($margin + $rectW))
    
    $pen.Dispose()
    $hlBrush.Dispose()
}

function Draw-Copy([System.Drawing.Graphics]$g, [int]$s) {
    $pen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#107C41")), ([Math]::Max(1.0, $s * 0.08))
    $fillBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#E1DFDD"))
    $fgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
    $margin = $s * 0.12
    $sheetW = $s * 0.55
    $sheetH = $s * 0.65
    
    $backX = $margin + $s * 0.2
    $backY = $margin
    $g.FillRectangle($fillBrush, $backX, $backY, $sheetW, $sheetH)
    $g.DrawRectangle($pen, $backX, $backY, $sheetW, $sheetH)
    
    $frontX = $margin
    $frontY = $margin + $s * 0.2
    $g.FillRectangle($fgBrush, $frontX, $frontY, $sheetW, $sheetH)
    $g.DrawRectangle($pen, $frontX, $frontY, $sheetW, $sheetH)
    
    $linePen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#107C41")), ([Math]::Max(1.0, $s * 0.05))
    $g.DrawLine($linePen, ($frontX + $sheetW * 0.2), ($frontY + $sheetH * 0.3), ($frontX + $sheetW * 0.8), ($frontY + $sheetH * 0.3))
    $g.DrawLine($linePen, ($frontX + $sheetW * 0.2), ($frontY + $sheetH * 0.5), ($frontX + $sheetW * 0.8), ($frontY + $sheetH * 0.5))
    $g.DrawLine($linePen, ($frontX + $sheetW * 0.2), ($frontY + $sheetH * 0.7), ($frontX + $sheetW * 0.6), ($frontY + $sheetH * 0.7))
    
    $pen.Dispose()
    $linePen.Dispose()
    $fillBrush.Dispose()
    $fgBrush.Dispose()
}

foreach ($s in $sizes) {
    # 1. AI
    $bmp = New-Object System.Drawing.Bitmap $s, $s
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    Draw-Ai $g $s
    $g.Dispose()
    $bmp.Save("$assetsDir\icon-$s.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    
    # 2. Grid
    $bmp = New-Object System.Drawing.Bitmap $s, $s
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    Draw-Grid $g $s
    $g.Dispose()
    $bmp.Save("$assetsDir\grid-$s.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    
    # 3. Copy
    $bmp = New-Object System.Drawing.Bitmap $s, $s
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    Draw-Copy $g $s
    $g.Dispose()
    $bmp.Save("$assetsDir\copy-$s.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

Write-Host "Done generating 9 icons in $assetsDir"
