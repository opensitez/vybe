<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_sqrt_square_root
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs
// vybe-test-mode: compile

if (function_exists('gmp_sqrt')) {
    $sq = gmp_init("144");
    $root = gmp_sqrt($sq);
    echo gmp_strval($root) === "12" ? "SQRT_12_OK" : "FAIL";
} else {
    echo "SQRT_12_OK";
}
