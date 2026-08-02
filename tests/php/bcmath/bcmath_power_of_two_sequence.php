<?php
// vybe-test: php/bcmath/bcmath_power_of_two_sequence
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

$result = [];
for ($i = 0; $i <= 10; $i++) {
    $result[] = bcpow('2', (string)$i);
}
echo implode(',', $result);
