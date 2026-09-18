# vybe-test: powershell/convertto_html_cmdlet/convertto_html_multiple_precontent_array
# The -PreContent parameter accepts an array of strings to emit sequential prefix blocks
$preList = @("<p>Header paragraph 1</p>", "<p>Header paragraph 2</p>")
$html = [pscustomobject]@{ Val = 99 } | ConvertTo-Html -PreContent $preList -Fragment
$text = $html -join " "

if ($text -notmatch "<p>Header paragraph 1</p>" -or $text -notmatch "<p>Header paragraph 2</p>") {
    Write-Host "FAIL: multiple PreContent entries missing from output: $text"
    exit 1
}

Write-Host "PASS"
exit 0
