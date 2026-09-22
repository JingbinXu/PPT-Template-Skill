param(
  [Parameter(Mandatory = $true)][string]$TemplatePath,
  [Parameter(Mandatory = $true)][string]$PlanPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Pt([double]$Inches) { return $Inches * 72.0 }
function Get-Value($Object, [string]$Name, $Default = $null) {
  $property = $Object.PSObject.Properties[$Name]
  if ($null -eq $property -or $null -eq $property.Value) { return $Default }
  return $property.Value
}
function New-PresentationApp {
  foreach ($progId in @('KWPP.Application', 'PowerPoint.Application')) {
    try { return New-Object -ComObject $progId } catch {}
  }
  throw '无法启动 WPS 演示或 Microsoft PowerPoint 自动化。请确认本机已安装可打开模板的演示软件。'
}
function Get-SlideFont($Slide) {
  for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
    try {
      $shape = $Slide.Shapes.Item($i)
      $value = [string]$shape.TextFrame.TextRange.Text
      if (-not [string]::IsNullOrWhiteSpace($value)) {
        $font = $shape.TextFrame.TextRange.Font
        if (-not [string]::IsNullOrWhiteSpace($font.Name)) { return @{ Name = $font.Name; Color = $font.Color.RGB } }
      }
    } catch {}
  }
  return @{ Name = 'Arial'; Color = 0x333333 }
}
function Keep-Text([string]$Text, $PreserveText) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
  foreach ($fragment in $PreserveText) {
    if (-not [string]::IsNullOrWhiteSpace([string]$fragment) -and $Text.Contains([string]$fragment)) { return $true }
  }
  return $false
}
function Clear-GroupText($Group, $PreserveText) {
  for ($i = 1; $i -le $Group.GroupItems.Count; $i++) {
    $item = $Group.GroupItems.Item($i)
    try {
      if ($item.Type -eq 6) { Clear-GroupText $item $PreserveText; continue }
      $value = [string]$item.TextFrame.TextRange.Text
      if (-not [string]::IsNullOrWhiteSpace($value) -and -not (Keep-Text $value $PreserveText)) { $item.TextFrame.TextRange.Text = '' }
    } catch {}
  }
}
function Clear-TemplateText($Slide, $PreserveText) {
  for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
    $shape = $Slide.Shapes.Item($i)
    try {
      if ($shape.Type -eq 6) { Clear-GroupText $shape $PreserveText; continue }
      $value = [string]$shape.TextFrame.TextRange.Text
      if (-not [string]::IsNullOrWhiteSpace($value) -and -not (Keep-Text $value $PreserveText)) { $shape.Delete() }
    } catch {}
  }
}
function Set-Text($Shape, $Block, $Font) {
  $text = [string](Get-Value $Block 'text' '')
  $fontName = [string](Get-Value $Block 'font_name' $Font.Name)
  $fontSize = [double](Get-Value $Block 'font_size' 18)
  $alignName = [string](Get-Value $Block 'align' 'left')
  $align = switch ($alignName.ToLowerInvariant()) { 'center' { 2 } 'right' { 3 } default { 1 } }
  $bold = if ([bool](Get-Value $Block 'bold' $false)) { -1 } else { 0 }
  $Shape.TextFrame.TextRange.Text = $text
  $Shape.TextFrame.TextRange.Font.Name = $fontName
  try { $Shape.TextFrame.TextRange.Font.NameFarEast = $fontName } catch {}
  $Shape.TextFrame.TextRange.Font.Size = $fontSize
  $Shape.TextFrame.TextRange.Font.Bold = $bold
  try { $Shape.TextFrame.TextRange.Font.Color.RGB = $Font.Color } catch {}
  $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $align
  $Shape.TextFrame.WordWrap = -1
  try { $Shape.TextFrame.AutoSize = 0 } catch {}
  try {
    $Shape.TextFrame2.MarginLeft = Pt 0.06
    $Shape.TextFrame2.MarginRight = Pt 0.06
    $Shape.TextFrame2.MarginTop = Pt 0.04
    $Shape.TextFrame2.MarginBottom = Pt 0.04
    $Shape.TextFrame2.AutoSize = 0
  } catch {}
}
function Add-TextBlock($Slide, $Block, $Font) {
  $shape = $Slide.Shapes.AddTextbox(1, (Pt ([double]$Block.x)), (Pt ([double]$Block.y)), (Pt ([double]$Block.w)), (Pt ([double]$Block.h)))
  Set-Text $shape $Block $Font
}
function Add-TableBlock($Slide, $Block, $Font) {
  $headers = @($Block.headers)
  $rows = @($Block.rows)
  if ($headers.Count -lt 1) { throw '表格块必须至少包含一个表头。' }
  $shape = $Slide.Shapes.AddTable($rows.Count + 1, $headers.Count, (Pt ([double]$Block.x)), (Pt ([double]$Block.y)), (Pt ([double]$Block.w)), (Pt ([double]$Block.h)))
  for ($column = 1; $column -le $headers.Count; $column++) {
    $cell = $shape.Table.Cell(1, $column).Shape
    $headerBlock = [pscustomobject]@{ text = [string]$headers[$column - 1]; font_size = (Get-Value $Block 'header_font_size' 14); bold = $true; align = 'center' }
    Set-Text $cell $headerBlock $Font
  }
  for ($row = 1; $row -le $rows.Count; $row++) {
    $values = @($rows[$row - 1])
    for ($column = 1; $column -le $headers.Count; $column++) {
      $value = if ($column -le $values.Count) { [string]$values[$column - 1] } else { '' }
      $cell = $shape.Table.Cell($row + 1, $column).Shape
      $bodyBlock = [pscustomobject]@{ text = $value; font_size = (Get-Value $Block 'font_size' 12); bold = $false; align = 'left' }
      Set-Text $cell $bodyBlock $Font
    }
  }
}

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) { throw "视觉模板不存在：$TemplatePath" }
if (-not (Test-Path -LiteralPath $PlanPath -PathType Leaf)) { throw "页面计划不存在：$PlanPath" }
if (Test-Path -LiteralPath $OutputPath) { throw "输出文件已存在，为避免覆盖已停止：$OutputPath" }
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding utf8 | ConvertFrom-Json
if (@($plan.slides).Count -eq 0) { throw '页面计划中没有 slides。' }
New-Item -ItemType Directory -Force (Split-Path -Parent $OutputPath) | Out-Null

$app = $null; $presentation = $null
try {
  $app = New-PresentationApp
  try { $app.Visible = 1 } catch {}
  $presentation = $app.Presentations.Add()
  foreach ($spec in $plan.slides) {
    $index = [int](Get-Value $spec 'template_slide' 0)
    if ($index -lt 1) { throw '每一页必须指定有效的 template_slide。' }
    $null = $presentation.Slides.InsertFromFile($TemplatePath, $presentation.Slides.Count, $index, $index)
  }
  if ($presentation.Slides.Count -ne @($plan.slides).Count) {
    throw "模板页插入数量异常：预期 $(@($plan.slides).Count)，实际 $($presentation.Slides.Count)。"
  }
  for ($slideIndex = 0; $slideIndex -lt @($plan.slides).Count; $slideIndex++) {
    $slide = $presentation.Slides.Item($slideIndex + 1)
    $spec = $plan.slides[$slideIndex]
    $font = Get-SlideFont $slide
    if ([bool](Get-Value $spec 'clear_template_text' $true)) { Clear-TemplateText $slide @((Get-Value $spec 'preserve_text' @())) }
    foreach ($block in @($spec.blocks)) {
      switch ([string](Get-Value $block 'kind' 'text')) {
        'text' { Add-TextBlock $slide $block $font }
        'table' { Add-TableBlock $slide $block $font }
        default { throw "不支持的内容块类型：$([string](Get-Value $block 'kind' ''))" }
      }
    }
  }
  $presentation.SaveAs($OutputPath, 24)
  Write-Output "OUTPUT=$OutputPath"
  Write-Output "SLIDES=$($presentation.Slides.Count)"
} finally {
  if ($presentation) { try { $presentation.Close() } catch {} }
  if ($app) { try { $app.Quit() } catch {} }
}
