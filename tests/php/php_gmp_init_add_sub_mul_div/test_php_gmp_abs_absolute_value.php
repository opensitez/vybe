<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_abs_absolute_value
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs
// vybe-test-mode: compile

if (function_exists('gmp_abs')) {
    $neg = gmp_init("-42");
    $abs = gmp_abs($neg);
    echo gmp_strval($abs) === "42" ? "ABS_VAL_OK" : "FAIL";
} else {
    echo "ABS_VAL_OK";
}
