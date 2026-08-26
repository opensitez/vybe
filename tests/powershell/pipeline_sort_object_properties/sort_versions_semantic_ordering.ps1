# vybe-test: powershell/pipeline_sort_object_properties/sort_versions_semantic_ordering
$v1 = [version]"10.0.0"
$v2 = [version]"2.1.0"
$v3 = [version]"2.0.5"
$sorted = @($v1, $v2, $v3 | Sort-Object)
if ($sorted[0] -ne $v3 -or $sorted[1] -ne $v2 -or $sorted[2] -ne $v1) {
    Write-Host "FAIL: Sort-Object Version semantic ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
