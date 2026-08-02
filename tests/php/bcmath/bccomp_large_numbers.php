<?php
// vybe-test: php/bcmath/bccomp_large_numbers
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

$big = '99999999999999999999';
$bigger = '100000000000000000000';
echo bccomp($big, $bigger);
echo bccomp($bigger, $big);
