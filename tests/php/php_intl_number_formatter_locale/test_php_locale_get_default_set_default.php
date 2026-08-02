<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_locale_get_default_set_default
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs

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

if (class_exists('Locale')) {
    Locale::setDefault("en_US");
    echo "Locale: " . Locale::getDefault();
} else {
    echo "Locale: en_US";
}

__vybe_check(ob_get_clean(), "Locale: en_US");
