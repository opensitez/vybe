<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_to_upper_rule
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs
// vybe-test-mode: compile

if (class_exists('Transliterator')) {
    $t = Transliterator::create("Upper");
    echo $t->transliterate("lowercase") === "LOWERCASE" ? "UPPER_RULE_OK" : "FAIL";
} else {
    echo "UPPER_RULE_OK";
}
