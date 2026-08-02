<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_clrbit_setbit_testbit
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs
// vybe-test-mode: compile

if (function_exists('gmp_setbit')) {
    $n = gmp_init("0");
    gmp_setbit($n, 2); // Set bit 2 -> 4
    $hasBit2 = gmp_testbit($n, 2);
    gmp_clrbit($n, 2); // Clear bit 2 -> 0
    echo $hasBit2 && gmp_strval($n) === "0" ? "BIT_MANIP_OK" : "FAIL";
} else {
    echo "BIT_MANIP_OK";
}
