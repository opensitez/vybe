# vybe-test: powershell/escape_sequences/control_char_escapes
# The backtick escapes that produce control characters, checked by CODE POINT
# so the assertion does not depend on how a terminal renders them.
$cases = @(
    @("`0", 0),
    @("`a", 7),
    @("`b", 8),
    @("`f", 12),
    @("`v", 11),
    @("`e", 27)
)
foreach ($c in $cases) {
    $code = [int][char]$c[0]
    if ($code -ne $c[1]) {
        Write-Host "FAIL: expected code $($c[1]), got $code"
        exit 1
    }
}
Write-Host 'PASS'
exit 0
