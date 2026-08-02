<?php
// vybe-test: php/spl/array_iterator_compatibility
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$it = new ArrayIterator([3, 1, 2]);
$it->asort(SORT_NUMERIC);
echo $it->count();
echo $it->current();
