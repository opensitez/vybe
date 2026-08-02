<?php
// vybe-test: php/bcmath/bcmath_pi_digits
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

// Leibniz formula approximation using bc (illustrative)
$pi = '0';
$sign = '1';
for ($k = 0; $k < 100; $k++) {
    $term = bcdiv($sign, bcadd(bcmul('2', (string)$k), '1'), 20);
    $pi = bcadd($pi, $term, 20);
    $sign = bcmul($sign, '-1');
}
$pi = bcmul($pi, '4', 10);
echo bccomp($pi, '3.1') > 0 && bccomp($pi, '3.2') < 0 ? 'pi in range' : 'out of range';
