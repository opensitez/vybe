<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_nested_sub_parent
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs
// vybe-test-mode: compile

$data = ["p1" => ["p2" => ["leaf" => "data"]]];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();
$sub = $pit->getChildren();
$subPit = new ParentIterator($sub);
$subPit->rewind();
echo $subPit->key() === "p2" ? "NESTED_PARENT_KEY_OK" : "FAIL";
