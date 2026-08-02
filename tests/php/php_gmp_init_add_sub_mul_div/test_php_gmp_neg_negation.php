<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_neg_negation
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs
// vybe-test-mode: compile

if (function_exists('gmp_neg')) {
    $pos = gmp_init("100");
    $neg = gmp_neg($pos);
    echo gmp_strval($neg) === "-100" ? "NEG_VAL_OK" : "FAIL";
} else {
    echo "NEG_VAL_OK";
}
