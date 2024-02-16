# Define the parent folder path
$parentFolderPath = "D:\Transfer\Outbox\" ################ amend to real path

# Define the output CSV file path
$outputCsvPath = Join-Path -Path $parentFolderPath -ChildPath "extracted_files_info.csv"

# Get all subdirectories of the parent folder
$subdirectories = Get-ChildItem -Path $parentFolderPath -Directory

# Initialize an array to hold files from each "Sent" folder
$allFilesFromSent = @()

foreach ($subdir in $subdirectories) {
    # Define the "Sent" folder path within each subdirectory
    $sentFolderPath = Join-Path -Path $subdir.FullName -ChildPath "Sent"
    
    # Check if the "Sent" folder exists
    if (Test-Path $sentFolderPath) {
        # Get all files from the "Sent" folder
        $files = Get-ChildItem -Path $sentFolderPath -File
        
        # Add files to the array
        $allFilesFromSent += $files
    }
}

# Sort the collected files by name and prepare data with name (until the first "-"), creation time, and size in KB
$sortedFilesInfo = $allFilesFromSent | Sort-Object Name | ForEach-Object {
    [PSCustomObject]@{
        Name = $_.Name.Split('-')[0].Trim()
        CreationTime = $_.CreationTime
        SizeKB = [math]::Round($_.Length / 1KB, 2)
    }
}

# Export the data to a CSV file
$sortedFilesInfo | Export-Csv -Path $outputCsvPath -NoTypeInformation

Write-Output "Files info exported to $outputCsvPath"
