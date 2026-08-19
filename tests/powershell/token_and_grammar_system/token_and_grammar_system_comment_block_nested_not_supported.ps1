# vybe-test: powershell/token_and_grammar_system/comment_block_nested_not_supported
$script = '<# outer <# nested #> tail #> write-host "x"'
$errors = @()
[System.Management.Automation.PSParser]::Tokenize($script, [ref]$errors) | Out-Null
if ($errors.Count -eq 0) {
    Write-Host 'FAIL: expected nested block comment syntax to be unsupported by parser'
    exit 1
}

Write-Host 'PASS'
exit 0
