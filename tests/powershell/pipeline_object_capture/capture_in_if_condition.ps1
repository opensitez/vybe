# vybe-test: powershell/pipeline_object_capture/capture_in_if_condition
if ($match = "hello world" | Select-String "world") {
    if ($match.Matches[0].Value -ne "world") {
        Write-Host "FAIL: pipeline assignment inside if condition expected 'world'"
        exit 1
    }
} else {
    Write-Host "FAIL: pipeline assignment in if condition evaluated false"
    exit 1
}
Write-Host "PASS"
exit 0
