<?php
// vybe-test: php/intl_formatting/intl_numberformatter_percent
// origin: languages/php/tests/php/test_intl_formatting.rs

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

if (!class_exists('NumberFormatter')) { echo 'skip'; } else {
    $fmt = new NumberFormatter('en_US', NumberFormatter::PERCENT);
    echo str_contains($fmt->format(0.25), '25') ? 'pct' : 'no';
}

__vybe_check(ob_get_clean(), "pct");
