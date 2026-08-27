<?php
// vybe-test: php/string_advanced/str_word_count_empty_subject_runtime
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo "str_word_count_empty_subject_runtime_ok";

__vybe_check(ob_get_clean(), "str_word_count_empty_subject_runtime_ok");
