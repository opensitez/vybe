<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_empty_array_no_parents
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs
// vybe-test-mode: compile

$rit = new RecursiveArrayIterator(["scalar1" => 1, "scalar2" => 2]);
$pit = new ParentIterator($rit);
echo count(iterator_to_array($pit)) === 0 ? "NO_PARENTS_OK" : "FAIL";
