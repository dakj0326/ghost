$src  = "$env:LOCALAPPDATA\Ghost\init.vbs"
$dest = "$env:LOCALAPPDATA\Ghost\test.vbs"

# Read original content
$content = Get-Content -Raw -Encoding UTF8 $src

# Convert LF to CRLF
$content = $content -replace "`n", "`r`n"

# Save as UTF8 without BOM
$utf8NoBom = New-Object System.Text.UTF8Encoding $False
[System.IO.File]::WriteAllText($dest, $content, $utf8NoBom)