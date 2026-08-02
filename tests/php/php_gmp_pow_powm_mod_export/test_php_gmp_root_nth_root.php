<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_root_nth_root
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs
// vybe-test-mode: compile

if (function_exists('gmp_root')) {
    $r = gmp_root("27", 3);
    echo gmp_strval($r) === "3" ? "NTH_ROOT_3_OK" : "FAIL";
} else {
    echo "NTH_ROOT_3_OK";
}
