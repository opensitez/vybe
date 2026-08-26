# vybe-test: powershell/type_matrix_and_vector_numerics/vector4_constructor_and_properties
$v = [System.Numerics.Vector4]::new(1.0, 2.0, 3.0, 4.0)
if ($v.X -ne 1.0 -or $v.W -ne 4.0) { Write-Host "FAIL: Vector4 constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
