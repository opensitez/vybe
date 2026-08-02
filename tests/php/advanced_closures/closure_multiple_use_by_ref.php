<?php
// vybe-test: php/advanced_closures/closure_multiple_use_by_ref
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$sum = 0;
$count = 0;
$record = function(int $n) use (&$sum, &$count): void {
    $sum += $n;
    $count++;
};
$record(10);
$record(20);
$record(30);
echo $sum;
echo $count;
