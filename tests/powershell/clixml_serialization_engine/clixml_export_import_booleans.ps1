# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_booleans
# Boolean values serialize as <B>true</B> and <B>false</B> and deserialize to [bool]
$tmp = [System.IO.Path]::GetTempFileName()
$bools = @($true, $false, $true)

$bools | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored[0] -ne $true -or $restored[1] -ne $false -or $restored[2] -ne $true) {
    Write-Host "FAIL: boolean sequence mismatch: @($($restored -join ', '))"
    exit 1
}

if (-not ($restored[0] -is [bool])) {
    Write-Host "FAIL: restored type is not [bool], got: $($restored[0].GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
