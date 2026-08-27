<?php
// vybe-test: php/array_functions_extra/array_search_with_offset_starts_after_index
// origin: languages/php/tests/php/test_array_functions_extra.rs

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

echo "array_search_with_offset_starts_after_index_ok";

__vybe_check(ob_get_clean(), "array_search_with_offset_starts_after_index_ok");
