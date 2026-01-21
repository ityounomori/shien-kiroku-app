$serverPath = "$PSScriptRoot\server.js"
$part2Path = "$PSScriptRoot\server_part2_fixed.js"
$outputPath = "$PSScriptRoot\server_fixed.js"

# Read ORIGINAL server.js (Assuming lines 0-1185 are SAFE - which they are based on check)
# Need to select valid lines.
# We checked that lines up to 1184 (closing brace of getPendingIncidentsByOffice) were intact.
# The corrupted part started at 1187 (// approve...).

$lines = [System.IO.File]::ReadAllLines($serverPath)

# Locate the end of getPendingIncidentsByOffice
$cutIndex = -1
for ($i = 0; $i -lt $lines.Count; $i++) {
    # Look for the last line of the new pending function
    if ($lines[$i] -match "debugRowDump: debugRowDump // \[Debug\] Return row dump") {
        # This is inside the function.
        # Find the closing brace.
    }
    if ($lines[$i] -match "return \{ data: \[\], hasMore: false, error: e.message, debugLogs: debugLogs \};") {
        # catch block end
    }
    if ($lines[$i] -match "^\}") {
        # Check if next lines are the corrupted ones
        if ($i + 2 -lt $lines.Count -and $lines[$i + 2] -match "// 謇ｿ隱榊・逅・") {
            $cutIndex = $i
            break
        }
        # Or checking for lines 1185
        if ($i -eq 1185) {
            # This aligns with our view_file
            $cutIndex = $i
            break
        }
    }
}

if ($cutIndex -eq -1) {
    # Fallback: We know roughly where it is.
    # Lines 0 to 1185 are the good part.
    $cutIndex = 1185
}

Write-Host "Cutting server.js at line $cutIndex"

$part1Lines = $lines[0..$cutIndex]

$part2Text = [System.IO.File]::ReadAllText($part2Path)

# Merge
$sb = [System.Text.StringBuilder]::new()
$part1Lines | ForEach-Object { $sb.AppendLine($_) }
$sb.AppendLine($part2Text)

# Write with UTF-8 BOM or NO BOM (VSCode likes No BOM usually, or UTF8).
# .NET UTF8Encoding(true) is BOM. (false) is No BOM.
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($outputPath, $sb.ToString(), $utf8NoBom)

Write-Host "Created server_fixed.js. Replacing..."
Remove-Item $serverPath -Force
Rename-Item $outputPath $serverPath
Write-Host "Done."
