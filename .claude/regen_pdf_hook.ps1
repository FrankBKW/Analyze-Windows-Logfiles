# PostToolUse Hook – regeneriert README.pdf wenn README.md geschrieben wurde
$raw  = [Console]::In.ReadToEnd()
try   { $json = $raw | ConvertFrom-Json } catch { exit 0 }

$fp = $json.tool_input.file_path
if (-not $fp) { $fp = $json.tool_response.filePath }
if (-not $fp) { exit 0 }

if ($fp -match 'README\.md$') {
    $dir = "C:\Claude\Analyze-Windows-Logfiles"
    Push-Location $dir
    python make_readme_pdf.py 2>&1 | Out-Null
    Pop-Location
}
