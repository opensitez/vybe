# vybe-test: powershell/eventing/remove_event
New-Event -SourceIdentifier RemoveEvent
Remove-Event -SourceIdentifier RemoveEvent
$ev = Get-Event -SourceIdentifier RemoveEvent -ErrorAction SilentlyContinue
if ($ev) {
    Write-Host "FAIL: expected event removed"
    exit 1
}
Write-Host "PASS"
exit 0
