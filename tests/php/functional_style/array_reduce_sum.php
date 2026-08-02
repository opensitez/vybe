<?php
// vybe-test: php/functional_style/array_reduce_sum
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

$nums = [1, 2, 3, 4, 5];
$sum = array_reduce($nums, fn($carry, $n) => $carry + $n, 0);
echo $sum;
