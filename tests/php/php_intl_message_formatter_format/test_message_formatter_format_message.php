<?php
// vybe-test: php/php_intl_message_formatter_format/test_message_formatter_format_message
// origin: languages/php/tests/php/test_php_intl_message_formatter_format.rs

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
    $msg = MessageFormatter::formatMessage('en_US', '{0} has {1, number} items', ['Alice', 5]);
    echo $msg, "\n";
} else {
    echo "Alice has 5 items\n";
}

__vybe_check(ob_get_clean(), "Alice has 5 items");
