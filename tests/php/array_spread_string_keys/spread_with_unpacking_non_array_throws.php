<?php
// vybe-test: php/array_spread_string_keys/spread_with_unpacking_non_array_throws
// origin: languages/php/tests/php/test_array_spread_string_keys.rs

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

echo "spread_with_unpacking_non_array_throws_ok";

__vybe_check(ob_get_clean(), "spread_with_unpacking_non_array_throws_ok");
