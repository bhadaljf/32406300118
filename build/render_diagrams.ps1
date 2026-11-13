param(
    [string]$PlantUmlJar = "C:\Users\fy\.vscode\extensions\jebbs.plantuml-2.18.1\plantuml.jar"
)

# 由 diagrams/*.puml 生成同名 PNG，供报告内嵌。
$ErrorActionPreference = 'Stop'
$diagramDir = Split-Path -Parent $PSScriptRoot
$diagramDir = Join-Path $diagramDir 'diagrams'

if (-not (Test-Path -LiteralPath $PlantUmlJar)) { throw "找不到 plantuml.jar：$PlantUmlJar" }
$files = Get-ChildItem -LiteralPath $diagramDir -Filter '*.puml' | ForEach-Object { $_.FullName }
foreach ($f in $files) {
    & java "-Dfile.encoding=UTF-8" -jar $PlantUmlJar -charset UTF-8 -tpng $f
    if ($LASTEXITCODE -ne 0) { throw "渲染失败：$f" }
    Write-Output "rendered $f"
}
