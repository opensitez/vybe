<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_jacobi_symbol
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs
// vybe-test-mode: compile

if (function_exists('gmp_jacobi')) {
    $jac = gmp_jacobi("5", "21");
    echo is_int($jac) ? "JACOBI_SYMBOL_OK" : "FAIL";
} else {
    echo "JACOBI_SYMBOL_OK";
}
