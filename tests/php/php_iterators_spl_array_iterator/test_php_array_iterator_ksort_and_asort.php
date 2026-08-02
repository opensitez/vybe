<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_array_iterator_ksort_and_asort
// origin: languages/php/tests/php/test_php_iterators_spl_array_iterator.rs
// vybe-test-mode: compile

$it = new ArrayIterator(["b" => 2, "a" => 1, "c" => 3]);
$it->ksort();
print_r(iterator_to_array($it));
