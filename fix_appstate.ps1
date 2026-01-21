$filePath = "$PSScriptRoot\record_form.html"
$patchContent = [System.IO.File]::ReadAllText("$PSScriptRoot\patch_appstate.js")

$content = [System.IO.File]::ReadAllText($filePath)

# Insert the patch right after the opening <script> tag that defines APP_STATE or near the top of the main script block.
# We look for "const APP_STATE =" or "var APP_STATE ="
# Or simply prepend it to the first <script> block that isn't a CDN link?
# Safer: Find "const APP_STATE = {" and ensure pending is there.

# Or just find "const APP_STATE = {" and replace it with:
# const APP_STATE = { pending: { offset: 0, limit: 50, hasMore: true, loading: false },

if ($content -match "const APP_STATE = \{") {
    $newContent = $content -replace "const APP_STATE = \{", "const APP_STATE = { pending: { offset: 0, limit: 50, hasMore: true, loading: false },"
    [System.IO.File]::WriteAllText($filePath, $newContent)
    Write-Host "Patched APP_STATE definition."
}
else {
    Write-Host "APP_STATE definition not found."
}
