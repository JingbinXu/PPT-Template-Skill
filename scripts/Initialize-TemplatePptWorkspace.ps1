param(
  [Parameter(Mandatory = $true)][string]$OutlinePath,
  [Parameter(Mandatory = $true)][string]$TemplatePath,
  [string]$StructureReferencePath,
  [string]$OutputRoot
)

$ErrorActionPreference = 'Stop'

function Resolve-RequiredFile([string]$Path, [string]$Label) {
  if ([string]::IsNullOrWhiteSpace($Path)) { throw "缺少$Label。" }
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label不存在：$Path" }
  return (Resolve-Path -LiteralPath $Path).Path
}

$outline = Resolve-RequiredFile $OutlinePath '内容大纲'
$template = Resolve-RequiredFile $TemplatePath '视觉模板'
if ([IO.Path]::GetExtension($template).ToLowerInvariant() -ne '.pptx') {
  throw "视觉模板必须是 .pptx：$template"
}
$reference = if ($StructureReferencePath) { Resolve-RequiredFile $StructureReferencePath '结构参考' } else { $null }
if ($reference -and [IO.Path]::GetExtension($reference).ToLowerInvariant() -ne '.pptx') {
  throw "结构参考必须是 .pptx：$reference"
}

$outlineDir = Split-Path -Parent $outline
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $outlineDir 'PPT输出' }
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$stem = [IO.Path]::GetFileNameWithoutExtension($outline)
$workspace = Join-Path $OutputRoot $stem
if (Test-Path -LiteralPath $workspace) {
  throw "工作区已存在，为避免覆盖已停止：$workspace"
}

New-Item -ItemType Directory -Path $workspace, (Join-Path $workspace 'source'), (Join-Path $workspace 'qa'), (Join-Path $workspace 'renders'), (Join-Path $workspace 'build') | Out-Null
$sourceDir = Join-Path $workspace 'source'
$outlineCopy = Join-Path $sourceDir ("content_outline" + [IO.Path]::GetExtension($outline))
$templateCopy = Join-Path $sourceDir 'visual_template.pptx'
Copy-Item -LiteralPath $outline -Destination $outlineCopy
Copy-Item -LiteralPath $template -Destination $templateCopy
$referenceCopy = $null
if ($reference) {
  $referenceCopy = Join-Path $sourceDir 'structure_reference.pptx'
  Copy-Item -LiteralPath $reference -Destination $referenceCopy
}

$manifest = [ordered]@{
  created_at = (Get-Date).ToString('o')
  original_outline = $outline
  original_template = $template
  original_structure_reference = $reference
  outline_copy = $outlineCopy
  template_copy = $templateCopy
  structure_reference_copy = $referenceCopy
  workspace = $workspace
  output_pptx = (Join-Path $workspace ("$stem-正式汇报.pptx"))
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $workspace 'manifest.json') -Encoding utf8
$manifest | ConvertTo-Json -Depth 4
