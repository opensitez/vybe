<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_collator_string_comparison
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs
// vybe-test-mode: compile

if (class_exists('Collator')) {
    $coll = new Collator("de_DE");
    $res = $coll->compare("ä", "z");
    echo ($res < 0) ? "COLLATOR_GERMAN_OK" : "FAIL";
}
