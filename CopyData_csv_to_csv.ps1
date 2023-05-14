# Get all CSV files in the source folder
$files = Get-ChildItem -Path "<<sourcefolderpath>>" -Filter *.csv

# Create a new target CSV file and add header row
Add-Content "<<targetfolderpath>>\targetfile_$($(get-Date -format 'ddMMyyyy')).csv" -Value '"VarName","TimeString","VarValue","Validity","Time_ms"'

# Loop through each CSV file in the source folder
foreach ($file in $files)
{   
    # Read each line of the CSV file and filter based on date condition
    $data = foreach($line in get-content $file.FullName)
	{
	$Arr = $line.Split(',')
	if($Arr[1] -match (get-date).adddays(-1).ToString("dd.MM.yyyy"))
	{
	     $line
	}
	else
	{
		# If the line doesn't match the date condition, skip it
	}
	}

	# Add the filtered data to the target CSV file
	$data | Add-Content "<<targetfolderpath>>\targetfile_$($(get-Date -format 'ddMMyyyy')).csv"
}






