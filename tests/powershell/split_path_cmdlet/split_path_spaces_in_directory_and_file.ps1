# vybe-test: powershell/split_path_cmdlet/split_path_spaces_in_directory_and_file
$pathWithSpaces = "/my personal documents/family vacation/photo 001.jpg"

$parent = Split-Path $pathWithSpaces -Parent
$leaf   = Split-Path $pathWithSpaces -Leaf

if ($parent -ne "/my personal documents/family vacation") {
    Write-Host "FAIL: parent with spaces mismatch, got: '$parent'"
    exit 1
}

if ($leaf -ne "photo 001.jpg") {
    Write-Host "FAIL: leaf with spaces mismatch, got: '$leaf'"
    exit 1
}

Write-Host "PASS"
exit 0
