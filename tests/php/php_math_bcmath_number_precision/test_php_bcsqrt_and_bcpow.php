<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_bcsqrt_and_bcpow
// origin: languages/php/tests/php/test_php_math_bcmath_number_precision.rs
// vybe-test-mode: compile

$sqrt = bcsqrt("2", 6);
$pow = bcpow("2", "10", 0);
echo "sqrt2=$sqrt pow=$pow";
