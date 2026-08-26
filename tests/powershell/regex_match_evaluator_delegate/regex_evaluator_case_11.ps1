# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_11
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_11", "word_11", $eval)
if ($res -ne "[word_11]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
