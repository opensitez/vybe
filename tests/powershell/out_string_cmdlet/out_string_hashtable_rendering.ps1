# vybe-test: powershell/out_string_cmdlet/out_string_hashtable_rendering
# Hashtables convert into key-value pair string formatting
$hash = @{ Environment = "Production"; ClusterId = "us-east-1" }
$output = $hash | Out-String

if ($output -notmatch "Environment" -or $output -notmatch "Production") {
    Write-Host "FAIL: Environment key-value pair missing in Out-String output"
    exit 1
}

if ($output -notmatch "ClusterId" -or $output -notmatch "us-east-1") {
    Write-Host "FAIL: ClusterId key-value pair missing in Out-String output"
    exit 1
}

Write-Host "PASS"
exit 0
