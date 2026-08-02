<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_forward_direction_constant
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs
// vybe-test-mode: compile

if (defined('Transliterator::FORWARD')) {
    echo Transliterator::FORWARD === 0 ? "FORWARD_0_OK" : "FAIL";
} else {
    echo "FORWARD_0_OK";
}
