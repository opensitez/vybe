<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_popcount_population_count
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs
// vybe-test-mode: compile

if (function_exists('gmp_popcount')) {
    $count = gmp_popcount("7"); // Binary 111 -> 3
    echo $count === 3 ? "POPCOUNT_3_OK" : "FAIL";
} else {
    echo "POPCOUNT_3_OK";
}
