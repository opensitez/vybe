<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_instanceof_iterator
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs
// vybe-test-mode: compile

$arr = new ArrayIterator([1]);
$inf = new InfiniteIterator($arr);
echo ($inf instanceof Iterator) ? "INSTANCEOF_ITERATOR" : "FAIL";
