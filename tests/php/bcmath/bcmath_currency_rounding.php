<?php
// vybe-test: php/bcmath/bcmath_currency_rounding
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

// Avoid float rounding errors in currency
$prices = ['10.99', '5.49', '3.25', '0.01'];
$total = '0.00';
foreach ($prices as $p) {
    $total = bcadd($total, $p, 2);
}
echo $total;
