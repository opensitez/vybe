<?php
// vybe-test: php/spl/array_object_iterate
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$ao = new ArrayObject(['a' => 1, 'b' => 2, 'c' => 3]);
$sum = 0;
foreach ($ao as $k => $v) { $sum += $v; }
echo $sum;
