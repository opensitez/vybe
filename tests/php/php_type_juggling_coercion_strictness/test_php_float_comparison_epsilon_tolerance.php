<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_float_comparison_epsilon_tolerance
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs
// vybe-test-mode: compile

$a = 0.1 + 0.2;
$b = 0.3;
$epsilon = 0.00001;
echo abs($a - $b) < $epsilon ? "FLOAT_EQUAL" : "FLOAT_NOT_EQUAL";
