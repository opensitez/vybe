<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_is_nan_is_infinite_checks
// origin: languages/php/tests/php/test_php_math_bcmath_number_precision.rs
// vybe-test-mode: compile

$nan = acos(8);
$inf = log(0);
echo is_nan($nan) ? "NAN" : "NOT_NAN";
echo is_infinite($inf) ? " INF" : " NOT_INF";
