# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_as_vt100_encoded_string
# -AsVT100EncodedString populates the .VT100EncodedString property with ANSI formatted text
$info = ConvertFrom-Markdown -InputObject "# Header Title`n`n**Bold Content**" -AsVT100EncodedString

if ([string]::IsNullOrEmpty($info.VT100EncodedString)) {
    Write-Host "FAIL: VT100EncodedString was null or empty with -AsVT100EncodedString"
    exit 1
}

if ($info.VT100EncodedString -notmatch "Header Title") {
    Write-Host "FAIL: VT100EncodedString did not contain source text: '$($info.VT100EncodedString)'"
    exit 1
}

Write-Host "PASS"
exit 0
