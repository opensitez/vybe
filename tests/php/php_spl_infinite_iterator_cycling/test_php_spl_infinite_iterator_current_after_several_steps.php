<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_current_after_several_steps
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs
// vybe-test-mode: compile

$arr = new ArrayIterator(["one", "two"]);
$inf = new InfiniteIterator($arr);
$inf->rewind();
$inf->next(); // two
$inf->next(); // one
$inf->next(); // two
echo $inf->current() === "two" ? "STEP_POSITION_OK" : "FAIL";
