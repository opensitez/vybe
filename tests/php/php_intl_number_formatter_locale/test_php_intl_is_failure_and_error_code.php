<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_intl_is_failure_and_error_code
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs
// vybe-test-mode: compile

if (function_exists('intl_get_error_code')) {
    $code = intl_get_error_code();
    echo intl_is_failure($code) ? "FAILURE" : "SUCCESS";
}
