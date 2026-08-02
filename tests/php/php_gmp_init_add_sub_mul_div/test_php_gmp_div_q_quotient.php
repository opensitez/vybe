<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_div_q_quotient
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs
// vybe-test-mode: compile

if (function_exists('gmp_div_q')) {
    $n1 = gmp_init("1000");
    $n2 = gmp_init("3");
    $q = gmp_div_q($n1, $n2);
    echo gmp_strval($q) === "333" ? "DIV_Q_OK" : "FAIL";
} else {
    echo "DIV_Q_OK";
}
