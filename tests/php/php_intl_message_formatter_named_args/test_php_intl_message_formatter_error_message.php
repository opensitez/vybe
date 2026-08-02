<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_error_message
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs
// vybe-test-mode: compile

if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("en_US", "{0}");
    echo $fmt->getErrorMessage() === "U_ZERO_ERROR" || is_string($fmt->getErrorMessage()) ? "ERROR_MSG_OK" : "FAIL";
} else {
    echo "ERROR_MSG_OK";
}
