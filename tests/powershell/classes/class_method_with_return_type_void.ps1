# vybe-test: powershell/classes/class_method_with_return_type_void
# A class method defined with [void] return type suppresses pipeline output and executes mutations
class MutationAccumulator {
    [int]$Total = 0

    [void] AddToTotal([int]$amount) {
        $this.Total += $amount
    }
}

$acc = [MutationAccumulator]::new()

# Calling [void] method should not emit any pipeline object
$emitted = $acc.AddToTotal(15)

if ($emitted -ne $null) {
    Write-Host "FAIL: void method emitted pipeline output: '$emitted'"
    exit 1
}

$acc.AddToTotal(25)

if ($acc.Total -ne 40) {
    Write-Host "FAIL: expected total 40, got $($acc.Total)"
    exit 1
}

Write-Host "PASS"
exit 0
