# start timer to measure elapsed time
$elapsed = [System.Diagnostics.Stopwatch]::StartNew()

# load required assemblies
[void][Reflection.Assembly]::LoadWithPartialName("System.Data")
[void][Reflection.Assembly]::LoadWithPartialName("System.Data.SqlClient")

# remove files that are older than 10 days from the specified source folder
Get-ChildItem –Path "<<sourcefolderpath>>" -Recurse | Where-Object {($_.LastWriteTime -lt (Get-Date).AddDays(-10))} | Remove-Item

# import CSV file into $Data variable and filter records based on time range
$Data = Import-Csv "<<sourcefolderpath>>\sourcefile.csv" -Delimiter ';' | where { ([datetime]::ParseExact($_.TimeString, "dd.MM.yyyy HH:mm:ss", $null)) -ge (get-date).AddHours(-1).AddMinutes(- (get-date).minute).AddSeconds(- (get-date).second) -and (([datetime]::ParseExact($_.TimeString, "dd.MM.yyyy HH:mm:ss", $null)) -lt (get-date).AddMinutes(- (get-date).minute).AddSeconds(- (get-date).second))}

# export filtered data to CSV file with timestamp in the filename
$Data | Export-Csv "<<targetfolderpath>>\targetfile_$($(get-Date -format 'ddMMyyyy_HH')).csv" -notypeinformation -Delimiter ';'

# write message to console
Write-Host "Daten kopiert"

# define variables for SQL server connection
$sqlserver = "<<Servername>>"
$database = "<<databasename>>"
$table = "<<tablename>>"

# get the latest CSV file in the specified source folder
$dir = "<<sourcefolderpath>>"
$latest = Get-ChildItem -Path $dir | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$csvfile = "<<sourcefolderpath>>\$latest"

# define delimiter variables
$oldDelimiter = ";" 
$newDelimiter = "|"

# specify whether the CSV file has column headers
$firstRowColumnNames = $true

# specify the batch size for bulk copying data to SQL Server
$batchsize = 50000

# define connection string for SQL Server
$connectionstring = "Data Source=$sqlserver;User id=user;password=pw;Initial Catalog=$database;"

# create SQLBulkCopy object and set properties
$bulkcopy = New-Object Data.SqlClient.SqlBulkCopy($connectionstring, [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock)
$bulkcopy.DestinationTableName = $table
$bulkcopy.bulkcopyTimeout = 0
$bulkcopy.batchsize = $batchsize

# create DataTable object
$datatable = New-Object System.Data.DataTable

# create StreamReader object to read CSV file
$reader = New-Object System.IO.StreamReader($csvfile)

# get column names from the first row of the CSV file
$columns = (Get-Content $csvfile -First 1).Split($oldDelimiter)

# if the CSV file has column headers, skip the first row
if ($firstRowColumnNames -eq $true) { $null = $reader.readLine() }

# add columns to the DataTable object
foreach ($column in $columns) { 
	$null = $datatable.Columns.Add()
}

# read through CSV file and add rows to DataTable object
while (($line = $reader.ReadLine()) -ne $null)  { 

    # initialize variables for counting quotes and string positions
    $sp = 0
    $point = 0
    
    # convert the line to an array of characters
    [char[]]$larr = $line
    
    # loop through each character in the line
    foreach ($item in $larr){ 
        # increment quote count if a quote is found
        if($item -eq """"){$sp++} 
        
        # check if the character is the old delimiter
        if($item -eq $oldDelimiter){ 
            # if the quote count is even, replace the delimiter with the new delimiter
            if ($sp%2 -eq 0) { 
                $line = $line.Remove($point,1).Insert($point,$newDelimiter)
            } 
        }
        # increment position counter
        $point++
    }
    # remove any remaining quotes from the line and add it to the DataTable
    $Line = $line.Replace("""","") 
    $null = $datatable.Rows.Add($line.Split($newDelimiter))  
	
    # check if the batch size has been reached, and if so, insert the batch into the database and clear the DataTable
    $i++; 
    if (($i % $batchsize) -eq 0) { 
        $bulkcopy.WriteToServer($datatable) 
        Write-Host "$i rows have been inserted in $($elapsed.Elapsed.ToString())."
        $datatable.Clear() 
    } 
} 

# insert any remaining rows in the DataTable into the database
if($datatable.Rows.Count -gt 0) {
    $bulkcopy.WriteToServer($datatable)
    $datatable.Clear()
}

# clean up resources
$reader.Close(); 
$reader.Dispose()
$bulkcopy.Close(); 
$bulkcopy.Dispose()
$datatable.Dispose()

# print the total elapsed time and perform garbage collection
Write-Host "Total Elapsed Time: $($elapsed.Elapsed.ToString())"
[System.GC]::Collect()
