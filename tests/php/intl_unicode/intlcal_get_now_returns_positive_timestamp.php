<?php
// vybe-test: php/intl_unicode/intlcal_get_now_returns_positive_timestamp
// origin: languages/php/tests/php/test_intl_unicode.rs

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

if (!class_exists('IntlCalendar')) { echo 'skip'; } else {
    echo IntlCalendar::getNow() > 0 ? 'now' : 'zero';
}

__vybe_check(ob_get_clean(), "now");
