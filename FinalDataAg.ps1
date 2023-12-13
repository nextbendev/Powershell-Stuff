# Define the path to the input CSV file and the output CSV file
$inputPath = "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Ref\LicData.csv"
$outputPath = "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Ref\LicDataAgg.csv"
$currentDate = Get-Date -Format "MMddyy"

$histPath = $currentDate + "_AggedData"
$histCloudOutputPath = "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\LicenseHist\$histPath.csv"

# Read the CSV file into a variable
$licenseData = Import-Csv -Path $inputPath

# Create a hashtable to hold the count of licenses by OU
$licenseCountByOU = @{}

# Define the license types
$licenseTypes = @('M365_G5_GCC', 'VISIOCLIENT_GOV', 'EXCHANGESTANDARD_GOV')

# Initialize the hashtable with OUs and license types
foreach ($entry in $licenseData) {
    $ouPath = $entry.OU -split ',' # Split the OU path
    $ou = if ($ouPath.Count -gt 1) { $ouPath[1] } else { $ouPath[0] } # Use the second segment if available

    # Initialize the OU in the hashtable if not present
    if (-not $licenseCountByOU.ContainsKey($ou)) {
        $licenseCountByOU[$ou] = @{}
        foreach ($type in $licenseTypes) {
            $licenseCountByOU[$ou][$type] = 0
        }
    }

    # Count the licenses
    foreach ($type in $licenseTypes) {
        if ($entry.Licenses -match $type) {
            $licenseCountByOU[$ou][$type] += 1
        }
    }
}

# Convert the hashtable to a custom object for export to CSV
$results = foreach ($ou in $licenseCountByOU.Keys) {
    [PSCustomObject]@{
        OU = $ou
        G5 = $licenseCountByOU[$ou]['M365_G5_GCC']
        VISIO = $licenseCountByOU[$ou]['VISIOCLIENT_GOV']
        EXCHANGE = $licenseCountByOU[$ou]['EXCHANGESTANDARD_GOV']
    }
}

# Export the results to the CSV file
$results | Export-Csv -Path $outputPath -NoTypeInformation
$results | Export-Csv -Path $histCloudOutputPath -NoTypeInformation
 
# Output the path to the console
Write-Host "Data aggregated and saved to $outputPath"


# Define the delay time in seconds
$delayInSeconds = 5  # for a 5 second delay

# Start the sleep timer
Start-Sleep -Seconds $delayInSeconds

# Run the script after the delay
& "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Scripts\GenHtml.ps1"
