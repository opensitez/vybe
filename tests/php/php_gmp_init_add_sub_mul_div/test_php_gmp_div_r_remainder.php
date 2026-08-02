<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_div_r_remainder
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs
// vybe-test-mode: compile

if (function_exists('gmp_div_r')) {
    $n1 = gmp_init("1000");
    $n2 = gmp_init("3");
    $r = gmp_div_r($n1, $n2);
    echo gmp_strval($r) === "1" ? "DIV_R_OK" : "FAIL";
} else {
    echo "DIV_R_OK";
}
