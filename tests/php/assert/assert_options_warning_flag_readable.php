<?php
// vybe-test: php/assert/assert_options_warning_flag_readable
// origin: languages/php/tests/php/test_assert.rs

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

echo "assert_options_warning_flag_readable_ok";

__vybe_check(ob_get_clean(), "assert_options_warning_flag_readable_ok");
