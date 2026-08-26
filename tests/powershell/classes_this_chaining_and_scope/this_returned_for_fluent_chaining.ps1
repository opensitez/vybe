# vybe-test: powershell/classes_this_chaining_and_scope/this_returned_for_fluent_chaining
class FluentBuilder {
    [string]$Content = ""
    [FluentBuilder]Add([string]$s) {
        $this.Content += "$s;"
        return $this
    }
}
$fb = [FluentBuilder]::new()
$null = $fb.Add("A").Add("B").Add("C")
if ($fb.Content -ne "A;B;C;") {
    Write-Host "FAIL: Fluent chaining returning `$this failed, got '$($fb.Content)'"
    exit 1
}
Write-Host "PASS"
exit 0
