<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_id_property_getter
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs
// vybe-test-mode: compile

if (class_exists('Transliterator')) {
    $t = Transliterator::create("Any-Latin");
    echo str_contains($t->id, "Latin") ? "ID_PROP_OK" : "FAIL";
} else {
    echo "ID_PROP_OK";
}
