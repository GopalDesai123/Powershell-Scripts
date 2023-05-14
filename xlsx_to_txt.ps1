# Create a new Excel application object and set its visibility and display alerts properties to false
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.DisplayAlerts = $false

# Get the latest file in the source directory and set the pathSource variable to its full path
$latest = Get-ChildItem -Path "<<SorceFolderPath>>" | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$pathSource = "<<SorceFolderPath>>\$latest"
#$pathSource = "<<SorceFolderPath>>\SourceFile.xlsx"

# Set the pathtarget variable to the full path of the target file
$pathtarget = "<<TargetFolderPath>>\TargetFile.txt"

# Open the latest file using the Excel application object created earlier and assign the resulting workbook object to the $Workbook variable
$WorkBook = $Excel.Workbooks.Open($pathSource)

# Change the extension of the latest file to .txt and save it to the target directory
$NewFilePath = [System.IO.Path]::ChangeExtension("<<TargetFolderPath>>\$latest",".txt")
$Workbook.SaveAs($NewFilepath, 6)

# Read the contents of the newly created .txt file, skip the first and last lines, and append the remaining content to the target file
Get-Content $NewFilePath | select -skip 1 | select -skipLast 1 | out-file -FilePath $pathtarget -Append -Encoding ascii

# Cleanup
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
