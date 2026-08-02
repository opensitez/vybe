<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_instanceof_recursive_filter_iterator
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs
// vybe-test-mode: compile

$rit = new RecursiveArrayIterator([]);
$pit = new ParentIterator($rit);
echo ($pit instanceof RecursiveFilterIterator) ? "INSTANCEOF_RECURSIVE_FILTER" : "FAIL";
