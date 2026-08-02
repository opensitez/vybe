<?php
// vybe-test: php/spl/array_iterator_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$it = new ArrayIterator([10, 20, 30]);
$sum = 0;
foreach ($it as $v) { $sum += $v; }
echo $sum;
