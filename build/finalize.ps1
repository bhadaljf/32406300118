param(
    [string]$Docx = "D:\1学习\面向对象\32406300118\out\32406300118.docx",
    [string]$Pdf  = "D:\1学习\面向对象\32406300118\out\32406300118.pdf"
)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx, $false, $false)
    $doc.Fields.Update() | Out-Null
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }
    $doc.Repaginate()
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
