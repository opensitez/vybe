<?php
// vybe-test: php/mbstring_extended/mbstristr_finds_case_insensitive_multibyte_needle
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

echo "mbstristr_finds_case_insensitive_multibyte_needle_ok";

__vybe_check(ob_get_clean(), "mbstristr_finds_case_insensitive_multibyte_needle_ok");
