# vybe-test: powershell/convert_path_cmdlet/convert_path_pipeline_streaming
$paths = @(".", "tests")

# Piping paths into Convert-Path converts each item to an absolute path string
$converted = @($paths | Convert-Path)

if ($converted.Count -ne 2) {
    Write-Host "FAIL: expected 2 converted paths, got $($converted.Count)"
    exit 1
}

if ($converted[0] -ne $PWD.Path -or ($converted[1] -notmatch "tests$")) {
    Write-Host "FAIL: pipeline converted paths mismatch: @($($converted -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
