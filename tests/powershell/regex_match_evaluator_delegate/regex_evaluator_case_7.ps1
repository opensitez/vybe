# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_7
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_7", "word_7", $eval)
if ($res -ne "[word_7]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
