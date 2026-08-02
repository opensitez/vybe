<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_init_base16_hex
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs
// vybe-test-mode: compile

if (function_exists('gmp_init')) {
    $hexNum = gmp_init("0x0F", 16);
    echo gmp_strval($hexNum, 10) === "15" ? "HEX_INIT_OK" : "FAIL";
} else {
    echo "HEX_INIT_OK";
}
