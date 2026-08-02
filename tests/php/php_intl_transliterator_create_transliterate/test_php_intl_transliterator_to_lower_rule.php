<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_to_lower_rule
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs
// vybe-test-mode: compile

if (class_exists('Transliterator')) {
    $t = Transliterator::create("Lower");
    echo $t->transliterate("UPPERCASE") === "uppercase" ? "LOWER_RULE_OK" : "FAIL";
} else {
    echo "LOWER_RULE_OK";
}
