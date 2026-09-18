# vybe-test: powershell/extended_type_system/ets_scriptproperty_returning_collection
# ScriptProperties can compute and return collections dynamically
Update-TypeData -TypeName System.Int32 -MemberType ScriptProperty -MemberName Divisors -Value {
    $n = $this
    @(1..$n | Where-Object { ($n % $_) -eq 0 })
} -Force

$divs = (12).Divisors

if ($divs.Count -ne 6) {
    Write-Host "FAIL: expected 6 divisors for 12, got $($divs.Count)"
    exit 1
}

$expectedDivs = @(1, 2, 3, 4, 6, 12)
for ($i = 0; $i -lt 6; $i++) {
    if ($divs[$i] -ne $expectedDivs[$i]) {
        Write-Host "FAIL: divisor index $i mismatch, expected $($expectedDivs[$i]), got $($divs[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
