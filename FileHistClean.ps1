# PowerShell Script to Delete Files Older Than 60 Days

# Define the path to the folder
$folderPath = "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\LicenseHist" # Replace with your folder path

# Get the current date
$currentDate = Get-Date

# Define the age limit (60 days)
$daysOld = 60

# Get all files in the specified folder
$files = Get-ChildItem -Path $folderPath -File

# Loop through each file
foreach ($file in $files) {
    # Calculate the age of the file
    $fileAge = $currentDate - $file.LastWriteTime

    # Check if the file is older than the specified age limit
    if ($fileAge.Days -gt $daysOld) {
        # Delete the file
        Remove-Item $file.FullName -Force
        Write-Host "Deleted file: $($file.FullName)"
    }
}

Write-Host "File deletion process completed."
