<?php
// vybe-test: php/bcmath/bcmath_financial_calculation
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

// Interest calculation without float precision loss
$principal = '1000.00';
$rate      = '0.05';
$periods   = '12';
$interest  = bcmul(bcmul($principal, $rate, 4), bcdiv($periods, '12', 4), 2);
$total     = bcadd($principal, $interest, 2);
echo $total;
