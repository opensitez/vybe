<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_strtolower_preserves_spaces_and_punctuation
// origin: languages/php/tests/php/test_php_string_case_folding_mb_case.rs

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

echo "test_php_strtolower_preserves_spaces_and_punctuation_ok";

__vybe_check(ob_get_clean(), "test_php_strtolower_preserves_spaces_and_punctuation_ok");
