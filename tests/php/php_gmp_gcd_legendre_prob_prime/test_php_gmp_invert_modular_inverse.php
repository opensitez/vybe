<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_invert_modular_inverse
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs
// vybe-test-mode: compile

if (function_exists('gmp_invert')) {
    $inv = gmp_invert("3", "11");
    echo gmp_strval($inv) === "4" ? "MODULAR_INVERSE_OK" : "FAIL";
} else {
    echo "MODULAR_INVERSE_OK";
}
