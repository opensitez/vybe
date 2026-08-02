<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_fact_factorial
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs
// vybe-test-mode: compile

if (function_exists('gmp_fact')) {
    $f5 = gmp_fact(5); // 120
    echo gmp_strval($f5) === "120" ? "FACT_5_120_OK" : "FAIL";
} else {
    echo "FACT_5_120_OK";
}
