# vybe-test: powershell/classes_base_method_calls/base_method_mutating_base_field
class BaseAccount {
    [int]$Balance = 100
    [void]Deposit([int]$amount) { $this.Balance += $amount }
}
class PremiumAccount : BaseAccount {
    [void]BonusDeposit([int]$amount) {
        ([BaseAccount]$this).Deposit($amount + 10)
    }
}
$pa = [PremiumAccount]::new()
$pa.BonusDeposit(50)
if ($pa.Balance -ne 160) {
    Write-Host "FAIL: Base method mutating base field failed, got $($pa.Balance)"
    exit 1
}
Write-Host "PASS"
exit 0
