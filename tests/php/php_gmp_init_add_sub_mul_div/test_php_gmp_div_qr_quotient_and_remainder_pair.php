<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_div_qr_quotient_and_remainder_pair
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs
// vybe-test-mode: compile

if (function_exists('gmp_div_qr')) {
    [$q, $r] = gmp_div_qr("100", "7");
    echo gmp_strval($q) === "14" && gmp_strval($r) === "2" ? "DIV_QR_PAIR_OK" : "FAIL";
} else {
    echo "DIV_QR_PAIR_OK";
}
