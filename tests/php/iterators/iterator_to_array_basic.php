<?php
// vybe-test: php/iterators/iterator_to_array_basic
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

$it = new ArrayIterator([3, 1, 4, 1, 5, 9]);
$arr = iterator_to_array($it, false);
sort($arr);
echo implode(',', $arr);
