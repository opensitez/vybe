<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_get_inner_iterator
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs
// vybe-test-mode: compile

$rit = new RecursiveArrayIterator([]);
$pit = new ParentIterator($rit);
echo $pit->getInnerIterator() === $rit ? "INNER_RIT_OK" : "FAIL";
