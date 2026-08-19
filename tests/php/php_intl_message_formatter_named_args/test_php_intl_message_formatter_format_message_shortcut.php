<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_format_message_shortcut
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
    $res = MessageFormatter::formatMessage("en_US", "Result: {0}", [100]);
    echo $res === "Result: 100" ? "FORMAT_MSG_SHORTCUT_OK" : "FAIL";
} else {
    echo "FORMAT_MSG_SHORTCUT_OK";
}


__vybe_check(ob_get_clean(), "FORMAT_MSG_SHORTCUT_OK");
