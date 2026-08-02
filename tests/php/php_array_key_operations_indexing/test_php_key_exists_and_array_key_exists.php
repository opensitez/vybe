<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_key_exists_and_array_key_exists
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs
// vybe-test-mode: compile

$arr = ["name" => "Alice", "null_val" => null];
echo array_key_exists("null_val", $arr) ? "KEY_EXISTS" : "NO";
