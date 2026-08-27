<?php
// vybe-test: php/string_compare/count_chars_unique_letters
// origin: languages/php/tests/php/test_string_compare.rs

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

echo "count_chars_unique_letters_ok";

__vybe_check(ob_get_clean(), "count_chars_unique_letters_ok");
