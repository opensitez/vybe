<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_infinite_iterator_looping
// origin: languages/php/tests/php/test_php_iterators_spl_array_iterator.rs
// vybe-test-mode: compile

$it = new InfiniteIterator(new ArrayIterator([1, 2]));
$limit = new LimitIterator($it, 0, 5);
echo implode(",", iterator_to_array($limit, false));
