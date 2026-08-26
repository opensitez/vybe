# vybe-test: powershell/type_matrix_and_vector_numerics/plane_constructor_normal_and_distance
$norm = [System.Numerics.Vector3]::UnitZ
$plane = [System.Numerics.Plane]::new($norm, 10.0)
if ($plane.Normal.Z -ne 1.0 -or $plane.D -ne 10.0) { Write-Host "FAIL: Plane constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
