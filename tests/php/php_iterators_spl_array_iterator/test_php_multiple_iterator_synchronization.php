<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_multiple_iterator_synchronization
// origin: languages/php/tests/php/test_php_iterators_spl_array_iterator.rs
// vybe-test-mode: compile

$mit = new MultipleIterator(MultipleIterator::MIT_NEED_ALL);
$mit->attachIterator(new ArrayIterator([1, 2]));
$mit->attachIterator(new ArrayIterator(["a", "b"]));

foreach ($mit as $pair) {
    echo $pair[0] . "-" . $pair[1] . "\n";
}
