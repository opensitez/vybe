<?php
// vybe-test: php/numeric_casts/octdec_stops_at_invalid_octal_digit
// origin: languages/php/tests/php/test_numeric_casts.rs

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

echo "octdec_stops_at_invalid_octal_digit_ok";

__vybe_check(ob_get_clean(), "octdec_stops_at_invalid_octal_digit_ok");
