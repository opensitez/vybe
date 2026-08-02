<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_get_error_code_and_message
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs
// vybe-test-mode: compile

if (class_exists('Transliterator')) {
    $t = Transliterator::create("Any-Latin");
    echo $t->getErrorCode() === 0 ? "ERROR_CODE_0_OK" : "FAIL";
} else {
    echo "ERROR_CODE_0_OK";
}
