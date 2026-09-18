# vybe-test: powershell/cmdlets/measure_command_cmdlet_timing
# Measure-Command measures the execution duration of a scriptblock and returns a TimeSpan
$elapsed = Measure-Command {
    $sum = 0
    for ($i = 0; $i -lt 1000; $i++) {
        $sum += $i
    }
}

if ($elapsed.GetType().Name -ne "TimeSpan") {
    Write-Host "FAIL: expected TimeSpan object, got $($elapsed.GetType().Name)"
    exit 1
}

# The duration must be non-negative
if ($elapsed.TotalMilliseconds -lt 0) {
    Write-Host "FAIL: measured duration is negative"
    exit 1
}

Write-Host "PASS"
exit 0
