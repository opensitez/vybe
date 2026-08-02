<?php
// vybe-test: php/php_intl_number_formatter_currency/test_number_formatter_currency_format
// origin: languages/php/tests/php/test_php_intl_number_formatter_currency.rs

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

if (class_exists('NumberFormatter')) {
    $fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
    $out = $fmt->formatCurrency(1234.56, 'USD');
    echo is_string($out) && str_contains($out, '1,234.56') ? 'currency_ok' : 'err', "\n";
} else {
    echo "currency_ok\n";
}

__vybe_check(ob_get_clean(), "currency_ok");
