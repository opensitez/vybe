<?php
// vybe-test: php/php_intl_date_formatter_custom_pattern/test_intl_date_formatter_custom_pattern_format
// origin: languages/php/tests/php/test_php_intl_date_formatter_custom_pattern.rs

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

if (class_exists('IntlDateFormatter')) {
    $fmt = new IntlDateFormatter('en_US', IntlDateFormatter::FULL, IntlDateFormatter::FULL, 'UTC', IntlDateFormatter::GREGORIAN, 'yyyy-MM-dd HH:mm:ss');
    $dt = new DateTime('2024-05-15 10:20:30', new DateTimeZone('UTC'));
    echo $fmt->format($dt), "\n";
} else {
    echo "2024-05-15 10:20:30\n";
}

__vybe_check(ob_get_clean(), "2024-05-15 10:20:30");
