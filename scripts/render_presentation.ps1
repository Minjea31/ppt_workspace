param(
    [Parameter(Mandatory=$true)][string]$PresentationPath,
    [string]$BuildDirectory,
    [string]$PdfPath,
    [switch]$Overwrite
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $PresentationPath).Path
if ([IO.Path]::GetExtension($source) -ne '.pptx') { throw 'A .pptx file is required.' }
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $BuildDirectory) {
    $job = Split-Path -Leaf (Split-Path -Parent $source)
    $BuildDirectory = Join-Path (Join-Path $repoRoot 'ppt_info') $job
}
if (-not $PdfPath) { $PdfPath = [IO.Path]::ChangeExtension($source, '.pdf') }
$BuildDirectory = [IO.Path]::GetFullPath($BuildDirectory)
$PdfPath = [IO.Path]::GetFullPath($PdfPath)
if ((Test-Path -LiteralPath $PdfPath) -and -not $Overwrite) { throw 'PDF already exists. Use -Overwrite or choose a new PdfPath.' }
New-Item -ItemType Directory -Path $BuildDirectory -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $PdfPath) -Force | Out-Null
$renderDir = Join-Path $BuildDirectory 'rendered'
New-Item -ItemType Directory -Path $renderDir -Force | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$doc = $null
$script:audit = @()

function Read-Text($shape, [int]$slideNumber, [string]$name) {
    if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
        $frame = $shape.TextFrame
        $script:audit += [pscustomobject]@{
            slide=$slideNumber; name=$name; text=$frame.TextRange.Text
            x=$shape.Left; y=$shape.Top; width=$shape.Width; height=$shape.Height
            availableWidth=($shape.Width-$frame.MarginLeft-$frame.MarginRight)
            availableHeight=($shape.Height-$frame.MarginTop-$frame.MarginBottom)
            boundWidth=$frame.TextRange.BoundWidth; boundHeight=$frame.TextRange.BoundHeight
        }
    }
    if ($shape.Type -eq 6) {
        foreach ($child in $shape.GroupItems) { Read-Text $child $slideNumber ($name+'/'+$child.Name) }
    }
    if ($shape.HasTable -eq -1) {
        for ($r=1; $r -le $shape.Table.Rows.Count; $r++) {
            for ($c=1; $c -le $shape.Table.Columns.Count; $c++) {
                Read-Text $shape.Table.Cell($r,$c).Shape $slideNumber ($name+"/cell_${r}_${c}")
            }
        }
    }
}
try {
    $doc = $app.Presentations.Open($source, -1, 0, 0)
    $doc.SaveAs($PdfPath, 32)
    $rendered = @()
    foreach ($slide in $doc.Slides) {
        $png = Join-Path $renderDir ('slide_{0:D2}.png' -f $slide.SlideIndex)
        $slide.Export($png, 'PNG', 1920, 1080)
        $rendered += $png
        foreach ($shape in $slide.Shapes) { Read-Text $shape $slide.SlideIndex $shape.Name }
    }
    ConvertTo-Json -InputObject @($script:audit) -Depth 5 | Set-Content -LiteralPath (Join-Path $BuildDirectory 'text_geometry.json') -Encoding UTF8
    [pscustomobject]@{presentation=$source; pdf=$PdfPath; slideCount=$doc.Slides.Count; rendered=$rendered} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $BuildDirectory 'render_manifest.json') -Encoding UTF8
    Write-Output "Rendered $($doc.Slides.Count) slides. PDF: $PdfPath"
    Write-Output "Images and audit: $BuildDirectory"
} finally {
    if ($null -ne $doc) { $doc.Close() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
