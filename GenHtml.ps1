# Define the paths to the files
$maxLicPath = "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Ref\cellPull.csv"
$currentLicPath = "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Ref\LicDataAgg.csv"
$outputPath = "\\hc22itsup\zCache\LicMonDisp.html"

# Read the CSV files into variables
$maxLicData = Import-Csv -Path $maxLicPath
$currentLicData = Import-Csv -Path $currentLicPath

# Create an empty array to store the results
$results = @()

# Initialize total variables
$totalG5 = 0
$totalVISIO = 0
$totalEXCHANGE = 0

$buttonTitle

# Calculate the difference
foreach ($maxLic in $maxLicData) {
    $currentLic = $currentLicData | Where-Object { $_.OU -eq $maxLic.OU }
    
    # If there is no matching OU in the current licenses, assume zero for the licenses
    if ($currentLic -eq $null) {
        $currentLic = New-Object PSObject -Property @{
            OU = $maxLic.OU
            G5 = 0
            VISIO = 0
            EXCHANGE = 0
        }
    }
    
    $diff = New-Object PSObject -Property @{
        OU = $maxLic.OU
        G5 = [int]$maxLic.G5 - [int]$currentLic.G5
        VISIO = [int]$maxLic.VISIO - [int]$currentLic.VISIO
        EXCHANGE = [int]$maxLic.EXCHANGE - [int]$currentLic.EXCHANGE
    }
    $results += $diff

     # Update totals
    $totalG5 += $diff.G5
    $totalVISIO += $diff.VISIO
    $totalEXCHANGE += $diff.EXCHANGE
}



# Define the CSS styles for the HTML output
$css = @"
<style>
    body {
        font-family: 'Arial', sans-serif;
        margin: 0;
        padding: 0;
        background-color: #f4f4f4;
    }
    h1 {
        color: #333;
        text-align: center;
    } 
    h2 {
        color: #333;
        text-align: center;
    }
    table {
        width: 80%;
        margin: 20px auto;
        border-collapse: collapse;
    }
    th, td {
        padding: 8px;
        text-align: left;
        border: 1px solid #ddd;
    }
    th {
        background-color: #4CAF50;
        color: white;
    }
 
  
    .footer {
        text-align: center;
        padding: 20px;
        font-size: 0.8em;
        color: #666;
    }
       /* New style for the Billable Dept column */
    th:first-child, td:first-child {
        max-width: 200px; /* Adjust the max-width as needed */
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
</style>
"@

$cssDark = @"
<style>
    body {
        font-family: 'Arial', sans-serif;
        margin: 0;
        padding: 0;
        background-color: #121212; /* Darker shade for better contrast */
        color: #121212; /* Light grey for readability */
    }
    h1, h2 {
        color: #FFFFFF; /* Bright white for headers */
    } 
    table {
        width: 80%;
        margin: 20px auto;
        border-collapse: collapse;
    }
    th, td {
        padding: 8px;
        text-align: left;
        border: 1px solid #2C2C2C;
    }
    th {
        background-color: #333333;
        color: #FFFFFF;
    }
    
    .footer {
        text-align: center;
        padding: 20px;
        font-size: 0.8em;
        color: #AAAAAA; /* Light grey for footer to be less prominent */
    }
    button {
        background-color: #252525; /* Dark buttons */
        color: #E0E0E0; /* Text color for buttons */
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        transition-duration: 0.4s;
        cursor: pointer;
    }
    button:hover {
        background-color: #555555; /* Lighten the button when hovered */
    }
    a {
        color: #4A90E2; /* Use a color that stands out but is not too bright */
        text-decoration: none;
    }
    a:hover {
        text-decoration: underline;
    }
    /* Style for the Billable Dept column */
    th:first-child, td:first-child {
        max-width: 200px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
</style>
"@


# Start building the HTML string
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>License Summary</title>
    $css
    <meta charset='UTF-8'>
</head>
<body>
    <h1>M365 LICENSE ALLOCATION BALANCE SHEET</h1>
    <h2>This report shows how many license's a department has to assign user's.</h2>
    <table>
        <tr>
            <th>Billable Dept</th>
            <th>G5</th>
            <th>VISIO</th>
            <th>EXCHANGE</th>
        </tr>
"@

# Append each result as a row in the HTML table
foreach ($result in $results) {
    # Function to determine the color based on the value
    function Get-Color($value) {
    if ($value -lt 0) {
        return 'color:red;'
    } elseif ($value -gt 0) {
        return 'color:green;'
    } else {
        return ''
    }
}


    $g5Color = Get-Color $result.G5
    $visioColor = Get-Color $result.VISIO
    $exchangeColor = Get-Color $result.EXCHANGE

    $html += @"
        <tr>
            <td>$($result.OU)</td>
            <td><span style='$g5Color'>$($result.G5)</span></td>
            <td><span style='$visioColor'>$($result.VISIO)</span></td>
            <td><span style='$exchangeColor'>$($result.EXCHANGE)</span></td>
        </tr>


"@
}
# Add a row for the totals
$html += @"
    <tr>
        <th>Total Availible</th>
        <th>$totalG5</th>
        <th>$totalVISIO</th>
        <th>$totalEXCHANGE</th>
    </tr>
"@

$html += @"
    </table>
    <div class='footer'>
    <p>Generated on $(Get-Date) by $($env:USERNAME).</p>
    <p>Report created by IT Department.</p>
    <button id='darkModeButton' onclick='toggleDarkMode()'>Toggle Dark Mode</button>
</div>
<script>
function toggleDarkMode() {
    var body = document.body;
    var headers = document.querySelectorAll('h1, h2'); // Select all h1 and h2 elements
    var button = document.getElementById('darkModeButton'); // Get the button by ID
    if (body.style.backgroundColor === '' || body.style.backgroundColor === 'white') {
        body.style.backgroundColor = '#333';
        body.style.color = 'white';
        button.textContent = 'Toggle Light Mode'; 
        // Iterate over each header and update the color
        headers.forEach(function(header) {
            header.style.color = 'white';
        });
    } else {
        body.style.backgroundColor = 'white';
        body.style.color = 'black';
        button.textContent = 'Toggle Dark Mode'; 
        // Iterate over each header and update the color
        headers.forEach(function(header) {
            header.style.color = '#333';
        });
    }
}
</script>
"@


# Save the HTML to a file
$html | Out-File -FilePath $outputPath


Write-Host "HTML Gen'd"

#Clean Ou's
& "\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Scripts\ouSwap.ps1"
