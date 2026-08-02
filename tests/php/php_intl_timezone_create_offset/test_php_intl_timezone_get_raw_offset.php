<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_get_raw_offset
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("UTC");
    echo "Offset: " . $tz->getRawOffset();
} else {
    echo "Offset: 0";
}

__vybe_check(ob_get_clean(), "Offset: 0");
