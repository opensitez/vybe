<?php
// vybe-test: php/bcmath/bcadd_large_numbers
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

$a = '99999999999999999999';
$b = '1';
echo bcadd($a, $b);
