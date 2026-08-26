# vybe-test: powershell/type_matrix_and_vector_numerics/vector3_constructor_and_properties
$v = [System.Numerics.Vector3]::new(1.0, 2.0, 3.0)
if ($v.X -ne 1.0 -or $v.Y -ne 2.0 -or $v.Z -ne 3.0) { Write-Host "FAIL: Vector3 constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
