<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_message_formatter_pattern_formatting
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

if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("en_US", "{0} has {1, number} new messages.");
    echo $fmt->format(["Alice", 5]);
} else {
    echo "Alice has 5 new messages.";
}

__vybe_check(ob_get_clean(), "Alice has 5 new messages.");
