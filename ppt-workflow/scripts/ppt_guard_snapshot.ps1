param(
    [Parameter(Mandatory=$true)][string]$InputPptx,
    [Parameter(Mandatory=$true)][string]$OutputJson
)

$ErrorActionPreference = 'Stop'
$msoTrue = -1
$msoGroup = 6
$msoPicture = 13
$msoLinkedPicture = 11

function Get-ShapeSnapshot($shape, [string]$path) {
    $text = ''
    try {
        if ($shape.HasTextFrame -eq $msoTrue -and $shape.TextFrame.HasText -eq $msoTrue) {
            $text = $shape.TextFrame.TextRange.Text
        }
    } catch {}
    $crop = $null
    if ($shape.Type -eq $msoPicture -or $shape.Type -eq $msoLinkedPicture) {
        try {
            $crop = @(
                [double]$shape.PictureFormat.CropLeft,
                [double]$shape.PictureFormat.CropTop,
                [double]$shape.PictureFormat.CropRight,
                [double]$shape.PictureFormat.CropBottom
            )
        } catch {}
    }
    $items = @([ordered]@{
        Path=$path; Name=$shape.Name; Type=[int]$shape.Type; Id=[int]$shape.Id
        Left=[math]::Round([double]$shape.Left,3); Top=[math]::Round([double]$shape.Top,3)
        Width=[math]::Round([double]$shape.Width,3); Height=[math]::Round([double]$shape.Height,3)
        Rotation=[math]::Round([double]$shape.Rotation,3); Text=$text; Crop=$crop
    })
    if ($shape.Type -eq $msoGroup) {
        for ($i=1; $i -le $shape.GroupItems.Count; $i++) {
            $items += Get-ShapeSnapshot $shape.GroupItems.Item($i) "$path/$i"
        }
    }
    return $items
}

$ppt = $null
$presentation = $null
try {
    $resolvedInput = (Resolve-Path -LiteralPath $InputPptx).Path
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($resolvedInput,$true,$true,$false)
    $slides = @()
    foreach ($slide in $presentation.Slides) {
        $shapes = @()
        for ($i=1; $i -le $slide.Shapes.Count; $i++) {
            $shapes += Get-ShapeSnapshot $slide.Shapes.Item($i) "$($slide.SlideIndex)/$i"
        }
        $effects = @()
        try {
            for ($i=1; $i -le $slide.TimeLine.MainSequence.Count; $i++) {
                $effect = $slide.TimeLine.MainSequence.Item($i)
                $effects += [ordered]@{
                    EffectType=[int]$effect.EffectType
                    TriggerType=[int]$effect.Timing.TriggerType
                    ShapeId=[int]$effect.Shape.Id
                }
            }
        } catch {}
        $slides += [ordered]@{Index=[int]$slide.SlideIndex;Shapes=$shapes;Effects=$effects}
    }
    $snapshot = [ordered]@{
        File=$resolvedInput
        SlideCount=[int]$presentation.Slides.Count
        Width=[double]$presentation.PageSetup.SlideWidth
        Height=[double]$presentation.PageSetup.SlideHeight
        Slides=$slides
    }
    $parent = Split-Path -Parent $OutputJson
    if ($parent) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    $snapshot | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
}
finally {
    if ($presentation) { $presentation.Close(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null }
    if ($ppt) { $ppt.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}

Write-Output $OutputJson
