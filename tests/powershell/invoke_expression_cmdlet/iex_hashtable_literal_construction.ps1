# vybe-test: powershell/invoke_expression_cmdlet/iex_hashtable_literal_construction
# Invoke-Expression evaluates hashtable literal syntax and returns a Hashtable
$result = Invoke-Expression "@{Key='Val'; Num=99}"

if ($result.Key -ne "Val" -or $result.Num -ne 99) {
    Write-Host "FAIL: hashtable literal mismatch, Key='$($result.Key)', Num=$($result.Num)"
    exit 1
}

Write-Host "PASS"
exit 0
