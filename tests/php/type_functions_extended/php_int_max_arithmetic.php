<?php
// vybe-test: php/type_functions_extended/php_int_max_arithmetic
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

$big = PHP_INT_MAX;
$overflow = $big + 1;
echo is_float($overflow) ? 'float' : 'int';
