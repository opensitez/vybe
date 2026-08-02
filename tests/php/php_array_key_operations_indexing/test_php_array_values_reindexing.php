<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_values_reindexing
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs
// vybe-test-mode: compile

$arr = [10 => "a", 20 => "b", 30 => "c"];
$reindexed = array_values($arr);
echo $reindexed[0] . "-" . $reindexed[1];
