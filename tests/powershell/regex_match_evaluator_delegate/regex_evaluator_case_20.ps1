# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_20
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_20", "word_20", $eval)
if ($res -ne "[word_20]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
