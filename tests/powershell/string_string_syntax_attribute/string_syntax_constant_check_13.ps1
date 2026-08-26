# vybe-test: powershell/string_string_syntax_attribute/string_syntax_constant_check_13
$jsonSyntax = [System.Diagnostics.CodeAnalysis.StringSyntaxAttribute]::Json
$regexSyntax = [System.Diagnostics.CodeAnalysis.StringSyntaxAttribute]::Regex
if ($jsonSyntax -ne "Json" -or $regexSyntax -ne "Regex") { Write-Host "FAIL: StringSyntax constants failed"; exit 1 }
Write-Host "PASS"; exit 0
