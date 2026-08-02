<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_multiple_rewinds
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs
// vybe-test-mode: compile

$arr = new ArrayIterator(["a", "b"]);
$inf = new InfiniteIterator($arr);
$inf->next();
$inf->rewind();
echo $inf->current() === "a" ? "MULTIPLE_REWIND_OK" : "FAIL";
