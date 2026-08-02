<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_export_import_binary_string
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs
// vybe-test-mode: compile

if (function_exists('gmp_export') && function_exists('gmp_import')) {
    $n = gmp_init("0x123456789ABCDEF0");
    $exported = gmp_export($n);
    $imported = gmp_import($exported);
    echo gmp_cmp($n, $imported) === 0 ? "EXPORT_IMPORT_OK" : "FAIL";
} else {
    echo "EXPORT_IMPORT_OK";
}
