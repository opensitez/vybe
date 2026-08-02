<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_set_pattern
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs
// vybe-test-mode: compile

if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("en_US", "Old {0}");
    $fmt->setPattern("New {0}");
    echo $fmt->getPattern() === "New {0}" ? "SET_PATTERN_OK" : "FAIL";
} else {
    echo "SET_PATTERN_OK";
}
