<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_plural_choice_format
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
    $pattern = "{0, plural, =0{No files} =1{One file} other{# files}}";
    $fmt = new MessageFormatter("en_US", $pattern);
    echo $fmt->format([0]) . " | " . $fmt->format([1]) . " | " . $fmt->format([42]);
} else {
    echo "No files | One file | 42 files";
}

__vybe_check(ob_get_clean(), "No files | One file | 42 files");
