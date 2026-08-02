<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_keys_empty_array_returns_empty
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs
// vybe-test-mode: compile

$k = array_keys([]);
echo count($k) === 0 ? "EMPTY_KEYS_OK" : "FAIL";
