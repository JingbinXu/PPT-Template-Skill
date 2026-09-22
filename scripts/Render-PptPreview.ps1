param(
  [Parameter(Mandatory = $true)][string]$PptxPath,
  [Parameter(Mandatory = $true)][string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) { throw "PPTX 不存在：$PptxPath" }
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null

function New-PresentationApp {
  foreach ($progId in @('KWPP.Application', 'PowerPoint.Application')) {
    try { return New-Object -ComObject $progId } catch {}
  }
  throw '无法启动 WPS 演示或 Microsoft PowerPoint，不能导出预览图。'
}

$app = $null; $presentation = $null
try {
  $app = New-PresentationApp
  try { $app.Visible = 1 } catch {}
  $presentation = $app.Presentations.Open((Resolve-Path -LiteralPath $PptxPath).Path, $false, $true, $false)
  for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
    $png = Join-Path $OutputDirectory ("slide-{0:D2}.png" -f $i)
    $presentation.Slides.Item($i).Export($png, 'PNG', 1600, 900)
  }
  try {
    Add-Type -AssemblyName System.Drawing
    $paths = Get-ChildItem -LiteralPath $OutputDirectory -Filter 'slide-*.png' | Sort-Object Name
    if ($paths.Count -gt 0) {
      $columns = 4; $thumbW = 320; $thumbH = 180; $rows = [math]::Ceiling($paths.Count / $columns)
      $sheet = New-Object System.Drawing.Bitmap ($columns * $thumbW), ($rows * $thumbH)
      $graphics = [System.Drawing.Graphics]::FromImage($sheet)
      $graphics.Clear([System.Drawing.Color]::White)
      for ($j = 0; $j -lt $paths.Count; $j++) {
        $image = [System.Drawing.Image]::FromFile($paths[$j].FullName)
        $x = ($j % $columns) * $thumbW; $y = [math]::Floor($j / $columns) * $thumbH
        $graphics.DrawImage($image, $x, $y, $thumbW, $thumbH)
        $image.Dispose()
      }
      $contact = Join-Path $OutputDirectory 'contact-sheet.png'
      $sheet.Save($contact, [System.Drawing.Imaging.ImageFormat]::Png)
      $graphics.Dispose(); $sheet.Dispose()
      Write-Output "CONTACT_SHEET=$contact"
    }
  } catch { Write-Warning "已导出单页预览，但未生成联系表：$($_.Exception.Message)" }
  Write-Output "RENDERS=$OutputDirectory"
} finally {
  if ($presentation) { try { $presentation.Close() } catch {} }
  if ($app) { try { $app.Quit() } catch {} }
}
