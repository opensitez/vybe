<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_export_little_endian_order
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs
// vybe-test-mode: compile

if (function_exists('gmp_export')) {
    $n = gmp_init("258"); // 0x0102
    $exp = gmp_export($n, 1, GMP_LITTLE_ENDIAN);
    echo is_string($exp) ? "LITTLE_ENDIAN_EXP_OK" : "FAIL";
} else {
    echo "LITTLE_ENDIAN_EXP_OK";
}
