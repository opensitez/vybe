<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_recursive_array_iterator_flattening
// origin: languages/php/tests/php/test_php_iterators_spl_array_iterator.rs
// vybe-test-mode: compile

$data = [1, [2, 3], [4, [5, 6]]];
$it = new RecursiveIteratorIterator(new RecursiveArrayIterator($data));
$flat = [];
foreach ($it as $v) {
    $flat[] = $v;
}
echo implode(",", $flat);
