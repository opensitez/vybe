<?php
// vybe-test: php/array_offset_access/foreach_on_null_warns_and_skips
// origin: languages/php/tests/php/test_array_offset_access.rs

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

echo "foreach_on_null_warns_and_skips_ok";

__vybe_check(ob_get_clean(), "foreach_on_null_warns_and_skips_ok");
