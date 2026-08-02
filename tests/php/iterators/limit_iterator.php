<?php
// vybe-test: php/iterators/limit_iterator
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

$it = new ArrayIterator(range(0, 99));
$slice = new LimitIterator($it, 5, 5);
$result = [];
foreach ($slice as $v) { $result[] = $v; }
echo implode(',', $result);
