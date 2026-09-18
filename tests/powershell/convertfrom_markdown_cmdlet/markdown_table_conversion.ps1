# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_table_conversion
# ConvertFrom-Markdown transforms GFM markdown tables into <table>, <thead>, and <tbody>
$pipe = [char]124
$tableText = "$pipe ColA $pipe ColB $pipe" + [char]10 + "$pipe---$pipe---$pipe" + [char]10 + "$pipe Val1 $pipe Val2 $pipe"

$info = ConvertFrom-Markdown -InputObject $tableText

if ($info.Html -notmatch "<table>") {
    Write-Host "FAIL: <table> element missing from converted table HTML"
    exit 1
}

if ($info.Html -notmatch "<th>ColA</th>" -or $info.Html -notmatch "<td>Val1</td>") {
    Write-Host "FAIL: table headers or data cells missing in HTML: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
