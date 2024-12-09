# Get all installed programs excluding Microsoft products
$installedPrograms = Get-WmiObject -Class Win32_Product | 
    Where-Object { $_.Vendor -notlike "*Microsoft*" } | 
    Select-Object Name, Version, Vendor, InstallDate

# Create HTML header with styling
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Installed Programs Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #2c3e50; text-align: center; }
        table { 
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }
        th, td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #34495e;
            color: white;
        }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #e9e9e9; }
    </style>
</head>
<body>
    <h1>Installed Programs Report</h1>
"@

# Convert data to HTML table
$htmlBody = $installedPrograms | ConvertTo-Html -Fragment

# Create HTML footer
$htmlFooter = @"
    <p style='text-align: center; color: #7f8c8d; margin-top: 20px;'>
        Report generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    </p>
</body>
</html>
"@

# Combine all HTML parts
$htmlReport = $htmlHeader + $htmlBody + $htmlFooter

# Save the HTML report to C: drive
$reportPath = "C:\InstalledPrograms.html"
$htmlReport | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "Report has been generated and saved to: $reportPath"
