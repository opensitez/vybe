# vybe-test: powershell/type_matrix_and_vector_numerics/vector2_in_generic_list
$list = [System.Collections.Generic.List[System.Numerics.Vector2]]::new()
$list.Add([System.Numerics.Vector2]::One)
if ($list.Count -ne 1 -or $list[0].X -ne 1.0) { Write-Host "FAIL: Vector2 in List failed"; exit 1 }
Write-Host "PASS"; exit 0
