<?php
// vybe-test: php/spl/spl_fixed_array_iterate
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$arr = SplFixedArray::fromArray([1, 4, 9, 16, 25]);
$sum = 0;
foreach ($arr as $v) { $sum += $v; }
echo $sum;
