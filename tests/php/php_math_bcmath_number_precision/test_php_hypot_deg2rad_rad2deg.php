<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_hypot_deg2rad_rad2deg
// origin: languages/php/tests/php/test_php_math_bcmath_number_precision.rs
// vybe-test-mode: compile

$h = hypot(3, 4);
$rad = deg2rad(180);
$deg = rad2deg(M_PI);
echo "hypot=$h rad=" . round($rad, 2) . " deg=" . round($deg, 0);
