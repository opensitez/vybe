# vybe-test: powershell/clixml_serialization_engine/clixml_noclobber_parameter_prevents_overwrite
# -NoClobber switch prevents Export-Clixml from overwriting an existing file, throwing an error
$tmp = [System.IO.Path]::GetTempFileName()
"initial_protected_content" | Export-Clixml -Path $tmp

$threwError = $false
try {
    "attempted_overwrite" | Export-Clixml -Path $tmp -NoClobber -ErrorAction Stop
} catch {
    $threwError = $true
}

$current = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if (-not $threwError) {
    Write-Host "FAIL: Export-Clixml -NoClobber failed to throw when destination existed"
    exit 1
}

if ($current -ne "initial_protected_content") {
    Write-Host "FAIL: file content was modified despite -NoClobber"
    exit 1
}

Write-Host "PASS"
exit 0
