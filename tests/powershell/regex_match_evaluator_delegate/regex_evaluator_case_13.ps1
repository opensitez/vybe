# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_13
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_13", "word_13", $eval)
if ($res -ne "[word_13]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
