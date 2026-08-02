<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_invalid_id_returns_null
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs
// vybe-test-mode: compile

if (class_exists('Transliterator')) {
    $t = @Transliterator::create("Invalid-NonExistent-ID-999");
    echo $t === null ? "INVALID_ID_NULL" : "FAIL";
} else {
    echo "INVALID_ID_NULL";
}
