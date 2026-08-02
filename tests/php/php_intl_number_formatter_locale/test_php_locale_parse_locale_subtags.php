<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_locale_parse_locale_subtags
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs
// vybe-test-mode: compile

if (class_exists('Locale')) {
    $subtags = Locale::parseLocale("zh_Hans_CN");
    echo "Language=" . $subtags["language"] . " Script=" . $subtags["script"];
}
