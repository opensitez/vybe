<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_empty_inner_iterator_halts
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs
// vybe-test-mode: compile

$arr = new ArrayIterator([]);
$inf = new InfiniteIterator($arr);
$inf->rewind();
echo !$inf->valid() ? "EMPTY_INFINITE_HALTS" : "FAIL";
