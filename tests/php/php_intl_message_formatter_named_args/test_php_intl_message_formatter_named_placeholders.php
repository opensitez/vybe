<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_named_placeholders
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
    $fmt = new MessageFormatter("en_US", "Hello {name}, you have {count} unread notifications.");
    echo $fmt->format(["name" => "Alice", "count" => 3]);
} else {
    echo "Hello Alice, you have 3 unread notifications.";
}

__vybe_check(ob_get_clean(), "Hello Alice, you have 3 unread notifications.");
