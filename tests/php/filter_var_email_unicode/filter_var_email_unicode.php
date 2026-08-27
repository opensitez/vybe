<?php
// vybe-test: php/filter_var_email_unicode/filter_var_email_unicode
// origin: languages/php/tests/php/test_filter_var_email_unicode.rs

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

echo "filter_var_email_unicode_ok";

__vybe_check(ob_get_clean(), "filter_var_email_unicode_ok");
