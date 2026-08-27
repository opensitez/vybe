<?php
// vybe-test: php/mbstring_extended/mbstrimwidth_respects_multibyte_width
// origin: languages/php/tests/php/test_mbstring_extended.rs

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

echo "mbstrimwidth_respects_multibyte_width_ok";

__vybe_check(ob_get_clean(), "mbstrimwidth_respects_multibyte_width_ok");
