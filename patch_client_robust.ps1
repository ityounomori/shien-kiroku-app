$clientPath = "$PSScriptRoot\record_form.html"
$newFuncPath = "$PSScriptRoot\client_fix_robust.js"

$clientLines = [System.IO.File]::ReadAllLines($clientPath)
$newFuncText = [System.IO.File]::ReadAllText($newFuncPath)

# Locate start and end lines
$startIndex = -1
$endIndex = -1

for ($i = 0; $i -lt $clientLines.Count; $i++) {
    if ($clientLines[$i] -match 'function loadIncidentPendingList\(') {
        $startIndex = $i
    }
    # Look for the NEXT function definition as the end marker
    if ($startIndex -ne -1 -and $clientLines[$i] -match 'function createIncidentPendingCard\(') {
        $endIndex = $i
        break
    }
}

if ($startIndex -ne -1 -and $endIndex -ne -1) {
    Write-Host "Found function block: Lines $($startIndex+1) to $($endIndex+1)"
    
    $part1 = $clientLines[0..($startIndex - 1)]
    $part3 = $clientLines[$endIndex..($clientLines.Count - 1)]
    
    $sb = [System.Text.StringBuilder]::new()
    $part1 | ForEach-Object { $sb.AppendLine($_) }
    $sb.AppendLine($newFuncText)
    $part3 | ForEach-Object { $sb.AppendLine($_) }
    
    [System.IO.File]::WriteAllText($clientPath, $sb.ToString())
    Write-Host "Success: Client function robustly updated."
}
else {
    Write-Host "Error: Could not locate markers."
}
