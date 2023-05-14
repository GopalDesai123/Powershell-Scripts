$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.DisplayAlerts = $false

$latest = Get-ChildItem -Path "<<SorceFolderPath>>" | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$pathSource = "<<SorceFolderPath>>\$latest"
#$pathSource = "<<SorceFolderPath>>\SourceFile.xlsx"
$pathtarget = "<<TargetFolderPath>>\TargetFile.txt"
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.DisplayAlerts = $false
$WorkBook = $Excel.Workbooks.Open($pathSource)
$NewFilePath = [System.IO.Path]::ChangeExtension("<<TargetFolderPath>>\$latest",".txt")
$Workbook.SaveAs($NewFilepath, 6)

Get-Content $NewFilePath | select -skip 1 | select -skipLast 1 | out-file -FilePath $pathtarget -Append -Encoding ascii

# cleanup
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()