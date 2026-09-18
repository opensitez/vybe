# vybe-test: powershell/ast_extent_source_mapping/ast_extent_file_path_populated_on_parsed_file
# When parsed via Parser.ParseFile, the Extent.File property reflects the absolute script path
$tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
$scriptContent = "`$value = 'from temp file'"
Set-Content -Path $tempFile -Value $scriptContent

try {
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($tempFile, [ref]$tokens, [ref]$errors)

    $extentFile = $ast.Extent.File

    if ([string]::IsNullOrWhiteSpace($extentFile)) {
        Write-Host "FAIL: Extent.File was empty on parsed file"
        exit 1
    }

    if ($extentFile -ne $tempFile) {
        Write-Host "FAIL: Extent.File mismatch, expected '$tempFile', got '$extentFile'"
        exit 1
    }
} finally {
    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force
    }
}

Write-Host "PASS"
exit 0
