<?php
// vybe-test: php/iterators/append_iterator
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

$it1 = new ArrayIterator([1, 2, 3]);
$it2 = new ArrayIterator([4, 5, 6]);
$combined = new AppendIterator();
$combined->append($it1);
$combined->append($it2);
$result = [];
foreach ($combined as $v) { $result[] = $v; }
echo implode(',', $result);
