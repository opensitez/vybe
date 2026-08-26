# vybe-test: powershell/type_matrix_and_vector_numerics/vector2_distance_between_points
$p1 = [System.Numerics.Vector2]::new(1.0, 1.0)
$p2 = [System.Numerics.Vector2]::new(4.0, 5.0)
$dist = [System.Numerics.Vector2]::Distance($p1, $p2)
if ($dist -ne 5.0) { Write-Host "FAIL: Vector2 Distance expected 5, got $dist"; exit 1 }
Write-Host "PASS"; exit 0
