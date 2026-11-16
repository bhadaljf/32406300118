param(
    [string]$Docx,
    [string]$Pdf
)

$root = Split-Path -Parent $PSScriptRoot
if (-not $Docx) { $Docx = Join-Path $root 'out\面向对象分析与设计实验2-软件243_封媛.docx' }
if (-not $Pdf)  { $Pdf  = Join-Path $root 'out\面向对象分析与设计实验2-软件243_封媛.pdf' }

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    # 第一遍：更新域与目录，保存
    $doc = $word.Documents.Open($Docx, $false, $false)
    $doc.Repaginate()
    $doc.Fields.Update() | Out-Null
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }
    $doc.Save()
    $doc.Close(0)

    # 第二遍：重新打开后分页已稳定，再更新一次目录并导出
    $doc = $word.Documents.Open($Docx, $false, $false)
    $doc.Repaginate()
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }
    $doc.Repaginate()
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }
    $doc.Save()
    $doc.ExportAsFixedFormat($Pdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $doc.Close(0)
    "OK pages=$pages pdf=$Pdf"
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

