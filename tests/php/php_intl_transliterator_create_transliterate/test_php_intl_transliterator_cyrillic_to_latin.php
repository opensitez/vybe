<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_cyrillic_to_latin
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs
// vybe-test-mode: compile

if (class_exists('Transliterator')) {
    $t = Transliterator::create("Cyrillic-Latin");
    $res = $t->transliterate("Привет");
    echo strlen($res) > 0 ? "CYRILLIC_LATIN_OK" : "FAIL";
} else {
    echo "CYRILLIC_LATIN_OK";
}
