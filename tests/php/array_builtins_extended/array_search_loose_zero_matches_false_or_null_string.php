<?php
// vybe-test: php/array_builtins_extended/array_search_loose_zero_matches_false_or_null_string
// origin: languages/php/tests/php/test_array_builtins_extended.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "array_search_loose_zero_matches_false_or_null_string_ok";

__vybe_check(ob_get_clean(), "array_search_loose_zero_matches_false_or_null_string_ok");
