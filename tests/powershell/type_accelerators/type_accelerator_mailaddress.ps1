# vybe-test: powershell/type_accelerators/type_accelerator_mailaddress
$mail = [mailaddress]"user@example.com"
if ($mail.User -ne "user") {
    Write-Host "FAIL: User expected user, got $($mail.User)"
    exit 1
}
if ($mail.Host -ne "example.com") {
    Write-Host "FAIL: Host expected example.com, got $($mail.Host)"
    exit 1
}
Write-Host "PASS"
exit 0
