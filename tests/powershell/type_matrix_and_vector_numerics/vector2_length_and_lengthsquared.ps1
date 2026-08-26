# vybe-test: powershell/type_matrix_and_vector_numerics/vector2_length_and_lengthsquared
$v = [System.Numerics.Vector2]::new(3.0, 4.0)
if ($v.Length() -ne 5.0 -or $v.LengthSquared() -ne 25.0) { Write-Host "FAIL: Vector2 Length failed"; exit 1 }
Write-Host "PASS"; exit 0
