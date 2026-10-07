param(
    [Parameter(Mandatory = $true)]
    [string]$Manifest,
    [Parameter(Mandatory = $true)]
    [string]$OutputDir,
    [int]$TargetWidth = 2600
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$items = Get-Content -Raw -LiteralPath $Manifest | ConvertFrom-Json
foreach ($item in $items) {
    $source = [string]$item.source
    $output = Join-Path $OutputDir ([string]$item.output)
    $metafile = [System.Drawing.Imaging.Metafile]::new($source)
    try {
        $width = [Math]::Max(1, $TargetWidth)
        $height = [Math]::Max(
            1,
            [int][Math]::Round($TargetWidth * [double]$metafile.Height / [double]$metafile.Width)
        )
        $bitmap = [System.Drawing.Bitmap]::new(
            $width,
            $height,
            [System.Drawing.Imaging.PixelFormat]::Format32bppArgb
        )
        try {
            $bitmap.SetResolution(300, 300)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.Clear([System.Drawing.Color]::White)
                $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
                $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
                $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $graphics.DrawImage($metafile, 0, 0, $width, $height)
            }
            finally {
                $graphics.Dispose()
            }
            $bitmap.Save($output, [System.Drawing.Imaging.ImageFormat]::Png)
            [PSCustomObject]@{
                Source = $source
                Output = $output
                Width = $width
                Height = $height
            }
        }
        finally {
            $bitmap.Dispose()
        }
    }
    finally {
        $metafile.Dispose()
    }
}
