# vybe-test: powershell/collections_generic_stack/reverse_polish_notation_evaluator
$s = [System.Collections.Generic.Stack[int]]::new()
# Evaluate: 3 4 + 2 * => (3 + 4) * 2 = 14
$s.Push(3)
$s.Push(4)
$s.Push($s.Pop() + $s.Pop())
$s.Push(2)
$s.Push($s.Pop() * $s.Pop())
$res = $s.Pop()
if ($res -ne 14 -or $s.Count -ne 0) {
    Write-Host "FAIL: RPN evaluation with Stack failed, expected 14, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
