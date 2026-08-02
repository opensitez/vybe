<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_accept_filter_logic
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs
// vybe-test-mode: compile

$data = ["parent" => [1, 2], "scalar" => 123];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();
echo $pit->accept() ? "ACCEPT_PARENT_TRUE" : "FAIL";
