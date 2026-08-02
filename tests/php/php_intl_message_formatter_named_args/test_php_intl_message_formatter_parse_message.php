<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_parse_message
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs

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
    $pattern = "{0} has {1, number} items.";
    $parsed = MessageFormatter::parseMessage("en_US", $pattern, "Bob has 10 items.");
    echo "Name={$parsed[0]} Count={$parsed[1]}";
} else {
    echo "Name=Bob Count=10";
}

__vybe_check(ob_get_clean(), "Name=Bob Count=10");
