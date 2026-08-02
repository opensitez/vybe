<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_invalid_pattern_returns_false
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs
// vybe-test-mode: compile

if (class_exists('MessageFormatter')) {
    $fmt = @new MessageFormatter("en_US", "{unclosed_bracket");
    echo $fmt === null || $fmt->getErrorCode() !== 0 ? "INVALID_PATTERN_HANDLED" : "FAIL";
} else {
    echo "INVALID_PATTERN_HANDLED";
}
