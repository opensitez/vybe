# vybe-test: powershell/type_matrix_and_vector_numerics/vector2_normalize_static_method
$v = [System.Numerics.Vector2]::new(0.0, 5.0)
$norm = [System.Numerics.Vector2]::Normalize($v)
if ($norm.X -ne 0.0 -or $norm.Y -ne 1.0) { Write-Host "FAIL: Vector2 Normalize failed"; exit 1 }
Write-Host "PASS"; exit 0
