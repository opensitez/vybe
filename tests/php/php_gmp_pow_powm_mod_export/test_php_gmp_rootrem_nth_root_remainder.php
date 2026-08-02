<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_rootrem_nth_root_remainder
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs
// vybe-test-mode: compile

if (function_exists('gmp_rootrem')) {
    [$root, $rem] = gmp_rootrem("30", 3); // 3^3 = 27, rem = 3
    echo gmp_strval($root) === "3" && gmp_strval($rem) === "3" ? "ROOTREM_OK" : "FAIL";
} else {
    echo "ROOTREM_OK";
}
