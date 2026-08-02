<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_locale_get_display_name
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs
// vybe-test-mode: compile

if (class_exists('Locale')) {
    echo Locale::getDisplayName("fr_FR", "en_US");
}
