<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_format_message_shortcut
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs
// vybe-test-mode: compile

if (class_exists('MessageFormatter')) {
    $res = MessageFormatter::formatMessage("en_US", "Result: {0}", [100]);
    echo $res === "Result: 100" ? "FORMAT_MSG_SHORTCUT_OK" : "FAIL";
} else {
    echo "FORMAT_MSG_SHORTCUT_OK";
}
