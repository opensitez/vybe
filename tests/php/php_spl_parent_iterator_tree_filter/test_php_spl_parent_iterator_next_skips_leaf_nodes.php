<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_next_skips_leaf_nodes
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs
// vybe-test-mode: compile

$data = ["leaf1" => 10, "parent1" => [20], "leaf2" => 30];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();
echo $pit->key() === "parent1" ? "FIRST_PARENT_KEY_OK" : "FAIL";
