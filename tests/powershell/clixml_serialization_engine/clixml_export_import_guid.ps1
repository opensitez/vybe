# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_guid
# System.Guid instances serialize with <G> tags and round-trip preserving Guid value
$tmp = [System.IO.Path]::GetTempFileName()
$orig = [Guid]::Parse("abcdef01-2345-6789-abcd-ef0123456789")

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored -ne $orig) {
    Write-Host "FAIL: Guid value mismatch, got: $restored"
    exit 1
}

if (-not ($restored -is [Guid])) {
    Write-Host "FAIL: expected [Guid] type, got: $($restored.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
