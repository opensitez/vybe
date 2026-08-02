<?php
// vybe-test: php/iterators/iterator_count
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

$it = new ArrayIterator([10, 20, 30, 40, 50]);
echo iterator_count($it);
