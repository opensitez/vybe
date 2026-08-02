<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_last_error_and_msg
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs
// vybe-test-mode: compile

@preg_match('/(?:\D+)+/', '12345678901234567890');
if (preg_last_error() !== PREG_NO_ERROR) {
    echo "PCRE Error: " . preg_last_error_msg();
}
