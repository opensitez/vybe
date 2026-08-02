<?php
// vybe-test: php/type_functions_extended/php_float_epsilon_comparison
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

$a = 0.1 + 0.2;
$b = 0.3;
$close = abs($a - $b) < PHP_FLOAT_EPSILON;
echo $close ? 'close' : 'far';
