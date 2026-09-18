# vybe-test: powershell/convertto_html_cmdlet/convertto_html_multiple_pipeline_objects
# Streaming multiple objects through the pipeline renders distinct table rows for each item
$objects = @(
    [pscustomobject]@{ Rank = 1; Label = "First" },
    [pscustomobject]@{ Rank = 2; Label = "Second" },
    [pscustomobject]@{ Rank = 3; Label = "Third" }
)

$html = $objects | ConvertTo-Html -Fragment
$text = $html -join " "

if ($text -notmatch "<td>1</td><td>First</td>") {
    Write-Host "FAIL: first object row missing"
    exit 1
}

if ($text -notmatch "<td>2</td><td>Second</td>") {
    Write-Host "FAIL: second object row missing"
    exit 1
}

if ($text -notmatch "<td>3</td><td>Third</td>") {
    Write-Host "FAIL: third object row missing"
    exit 1
}

Write-Host "PASS"
exit 0
