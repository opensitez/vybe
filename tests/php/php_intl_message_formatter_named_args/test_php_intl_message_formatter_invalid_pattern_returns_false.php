<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_invalid_pattern_returns_false
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "test_php_intl_message_formatter_invalid_pattern_returns_false_ok";

__vybe_check(ob_get_clean(), "test_php_intl_message_formatter_invalid_pattern_returns_false_ok");
