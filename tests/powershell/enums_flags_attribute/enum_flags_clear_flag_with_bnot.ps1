# vybe-test: powershell/enums_flags_attribute/enum_flags_clear_flag_with_bnot
[System.FlagsAttribute()]
enum States {
    Active = 1
    Verified = 2
    Admin = 4
}
$state = [States]::Active -bor [States]::Verified -bor [States]::Admin
# Remove Verified: state & ~Verified
$state = $state -band (-bnot [States]::Verified)
if ($state.HasFlag([States]::Verified) -or -not $state.HasFlag([States]::Active) -or -not $state.HasFlag([States]::Admin)) {
    Write-Host "FAIL: Clear flag with -bnot failed"
    exit 1
}
Write-Host "PASS"
exit 0
