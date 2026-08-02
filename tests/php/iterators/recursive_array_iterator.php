<?php
// vybe-test: php/iterators/recursive_array_iterator
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

$tree = ['a', ['b', 'c'], ['d', ['e', 'f']]];
$it = new RecursiveIteratorIterator(
    new RecursiveArrayIterator($tree)
);
$leaves = [];
foreach ($it as $leaf) { $leaves[] = $leaf; }
echo implode(',', $leaves);
