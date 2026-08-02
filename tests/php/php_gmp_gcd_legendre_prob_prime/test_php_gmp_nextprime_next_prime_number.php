<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_nextprime_next_prime_number
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs
// vybe-test-mode: compile

if (function_exists('gmp_nextprime')) {
    $next = gmp_nextprime("14");
    echo gmp_strval($next) === "17" ? "NEXTPRIME_17_OK" : "FAIL";
} else {
    echo "NEXTPRIME_17_OK";
}
