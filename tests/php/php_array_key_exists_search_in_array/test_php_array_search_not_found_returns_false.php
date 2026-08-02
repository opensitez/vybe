<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_search_not_found_returns_false
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs
// vybe-test-mode: compile

$arr = [1, 2, 3];
$res = array_search(99, $arr, true);
echo $res === false ? "NOT_FOUND_FALSE" : "FAIL";
