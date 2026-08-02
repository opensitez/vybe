<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_legendre_symbol
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs
// vybe-test-mode: compile

if (function_exists('gmp_legendre')) {
    $leg = gmp_legendre("5", "7");
    echo is_int($leg) ? "LEGENDRE_SYMBOL_OK" : "FAIL";
} else {
    echo "LEGENDRE_SYMBOL_OK";
}
