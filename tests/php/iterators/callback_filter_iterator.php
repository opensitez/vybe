<?php
// vybe-test: php/iterators/callback_filter_iterator
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

$it = new ArrayIterator(range(1, 10));
$evens = new CallbackFilterIterator($it, fn($v) => $v % 2 === 0);
$result = [];
foreach ($evens as $v) { $result[] = $v; }
echo implode(',', $result);
