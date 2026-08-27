<?php
// vybe-test: php/cross_lang/dictionary
// origin: languages/php/tests/php/test_cross_lang.rs

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

echo "dictionary_ok";

__vybe_check(ob_get_clean(), "dictionary_ok");
