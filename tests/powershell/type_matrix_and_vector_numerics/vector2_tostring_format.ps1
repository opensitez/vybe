# vybe-test: powershell/type_matrix_and_vector_numerics/vector2_tostring_format
$v = [System.Numerics.Vector2]::new(1.5, 2.5)
$str = $v.ToString()
if (-not $str.Contains("1.5") -or -not $str.Contains("2.5")) { Write-Host "FAIL: Vector2 ToString failed, got $str"; exit 1 }
Write-Host "PASS"; exit 0
