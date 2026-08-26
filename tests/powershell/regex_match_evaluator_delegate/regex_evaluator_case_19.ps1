# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_19
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_19", "word_19", $eval)
if ($res -ne "[word_19]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
