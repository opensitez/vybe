# vybe-test: powershell/type_matrix_and_vector_numerics/vector2_dot_product
$v1 = [System.Numerics.Vector2]::new(1.0, 2.0)
$v2 = [System.Numerics.Vector2]::new(3.0, 4.0)
$dot = [System.Numerics.Vector2]::Dot($v1, $v2)
if ($dot -ne 11.0) { Write-Host "FAIL: Vector2 Dot product expected 11, got $dot"; exit 1 }
Write-Host "PASS"; exit 0
