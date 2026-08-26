# vybe-test: powershell/enums_flags_attribute/enum_flags_hashtable_key
[System.FlagsAttribute()]
enum HashKeyFlags {
    OptionA = 1
    OptionB = 2
}
$key = [HashKeyFlags]::OptionA -bor [HashKeyFlags]::OptionB
$ht = @{ $key = "dual_mode" }
if ($ht[$key] -ne "dual_mode") {
    Write-Host "FAIL: Flags enum as hashtable key failed"
    exit 1
}
Write-Host "PASS"
exit 0
