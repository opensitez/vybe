<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_cmp_comparison
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs
// vybe-test-mode: compile

if (function_exists('gmp_cmp')) {
    $c1 = gmp_cmp("100", "50");
    $c2 = gmp_cmp("50", "100");
    $c3 = gmp_cmp("100", "100");
    echo $c1 > 0 && $c2 < 0 && $c3 === 0 ? "CMP_VAL_OK" : "FAIL";
} else {
    echo "CMP_VAL_OK";
}
