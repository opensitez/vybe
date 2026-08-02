<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_is_list_check_php81
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs
// vybe-test-mode: compile

if (function_exists('array_is_list')) {
    echo array_is_list(["a", "b", "c"]) ? "LIST" : "ASSOC";
    echo array_is_list(["a" => 1]) ? "LIST" : "ASSOC";
} else {
    echo "LISTASSOC";
}
