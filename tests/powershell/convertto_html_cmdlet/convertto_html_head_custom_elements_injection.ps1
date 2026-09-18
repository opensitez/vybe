# vybe-test: powershell/convertto_html_cmdlet/convertto_html_head_custom_elements_injection
# The -Head parameter injects custom styles and metadata into the HTML <head> section
$customCss = "<style>th { background-color: #f0f0f0; }</style>"
$html = [pscustomobject]@{ Metric = "CPU" } | ConvertTo-Html -Head $customCss
$text = $html -join " "

if ($text -notmatch [regex]::Escape($customCss)) {
    Write-Host "FAIL: custom head styles not found in HTML document head"
    exit 1
}

Write-Host "PASS"
exit 0
