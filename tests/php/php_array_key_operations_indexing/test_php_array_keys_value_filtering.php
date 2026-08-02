<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_keys_value_filtering
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs
// vybe-test-mode: compile

$array = ["blue", "red", "green", "blue", "blue"];
$blueKeys = array_keys($array, "blue", strict: true);
echo implode(",", $blueKeys);
