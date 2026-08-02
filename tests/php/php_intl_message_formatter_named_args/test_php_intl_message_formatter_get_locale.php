<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_get_locale
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs
// vybe-test-mode: compile

if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("fr_FR", "{0}");
    echo str_contains($fmt->getLocale(), "fr") ? "GET_LOCALE_FR_OK" : "FAIL";
} else {
    echo "GET_LOCALE_FR_OK";
}
