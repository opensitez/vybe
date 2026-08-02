<?php
// vybe-test: php/spl/array_iterator_sort_flags
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$it = new ArrayIterator(['banana', 'apple', 'cherry']);
$it->asort();
echo implode(',', iterator_to_array($it));
