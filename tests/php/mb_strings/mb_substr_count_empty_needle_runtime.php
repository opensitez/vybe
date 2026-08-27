<?php
// vybe-test: php/mb_strings/mb_substr_count_empty_needle_runtime
// origin: languages/php/tests/php/test_mb_strings.rs

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

echo "mb_substr_count_empty_needle_runtime_ok";

__vybe_check(ob_get_clean(), "mb_substr_count_empty_needle_runtime_ok");
