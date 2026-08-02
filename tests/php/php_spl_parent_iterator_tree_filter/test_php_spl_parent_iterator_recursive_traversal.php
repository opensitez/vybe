<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_recursive_traversal
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs
// vybe-test-mode: compile

$tree = [
    "level1" => [
        "level2" => ["leaf" => 1]
    ]
];
$rit = new RecursiveArrayIterator($tree);
$pit = new ParentIterator($rit);
echo count(iterator_to_array($pit)) === 1 ? "PARENT_RECURSIVE_OK" : "FAIL";
