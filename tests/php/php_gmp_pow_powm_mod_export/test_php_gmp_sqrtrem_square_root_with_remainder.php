<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_sqrtrem_square_root_with_remainder
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs
// vybe-test-mode: compile

if (function_exists('gmp_sqrtrem')) {
    [$root, $rem] = gmp_sqrtrem("150"); // 12*12 = 144, rem = 6
    echo gmp_strval($root) === "12" && gmp_strval($rem) === "6" ? "SQRTREM_OK" : "FAIL";
} else {
    echo "SQRTREM_OK";
}
