# Define the path to the HTML file
$htmlFilePath = "\\hc22itsup\zCache\LicMonDisp.html"

# Define the finalOu hash table
$finalOu =@{
    'ADM' =	'Admin'

}

# Read the HTML file into a variable
$htmlContent = Get-Content -Path $htmlFilePath -Raw

# Replace each short OU code with its full description
foreach ($ouCode in $finalOu.Keys) {
    $htmlContent = $htmlContent -replace "<td>$ouCode</td>", "<td>$($finalOu[$ouCode])</td>"
}

# Save the modified content back to the HTML file
$htmlContent | Set-Content -Path $htmlFilePath

Write-Host "OU's Cleaned"

