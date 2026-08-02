<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_fill_keys_custom_keys
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs
// vybe-test-mode: compile

$keys = ["foo", 5, 10, "bar"];
$a = array_fill_keys($keys, "default");
echo $a["foo"] === "default" && $a[5] === "default" ? "FILL_KEYS_OK" : "FAIL";
