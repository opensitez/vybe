# vybe-test: powershell/type_matrix_and_vector_numerics/vector2_constructor_and_properties
$v = [System.Numerics.Vector2]::new(3.0, 4.0)
if ($v.X -ne 3.0 -or $v.Y -ne 4.0) { Write-Host "FAIL: Vector2 constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
