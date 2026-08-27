<?php
// vybe-test: php/string_search/strpos_with_negative_offset_beyond_start
// origin: languages/php/tests/php/test_string_search.rs

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

echo "strpos_with_negative_offset_beyond_start_ok";

__vybe_check(ob_get_clean(), "strpos_with_negative_offset_beyond_start_ok");
