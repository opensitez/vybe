<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_next_wraps_around
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs
// vybe-test-mode: compile

$arr = new ArrayIterator(["x"]);
$inf = new InfiniteIterator($arr);
$inf->rewind();
$inf->next();
echo $inf->valid() && $inf->current() === "x" ? "WRAP_OK" : "FAIL";
