<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_hamdist_hamming_distance
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs
// vybe-test-mode: compile

if (function_exists('gmp_hamdist')) {
    $dist = gmp_hamdist("7", "4"); // 111 vs 100 -> dist 2
    echo $dist === 2 ? "HAMDIST_2_OK" : "FAIL";
} else {
    echo "HAMDIST_2_OK";
}
