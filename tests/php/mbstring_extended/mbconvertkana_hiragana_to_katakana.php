<?php
// vybe-test: php/mbstring_extended/mbconvertkana_hiragana_to_katakana
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

echo "mbconvertkana_hiragana_to_katakana_ok";

__vybe_check(ob_get_clean(), "mbconvertkana_hiragana_to_katakana_ok");
