<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_get_pattern
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs
// vybe-test-mode: compile

if (class_exists('MessageFormatter')) {
    $pattern = "Welcome {0}!";
    $fmt = new MessageFormatter("en_US", $pattern);
    echo $fmt->getPattern() === $pattern ? "GET_PATTERN_OK" : "FAIL";
} else {
    echo "GET_PATTERN_OK";
}
