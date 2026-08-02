<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_base_convert_hex_dec_bin
// origin: languages/php/tests/php/test_php_math_bcmath_number_precision.rs
// vybe-test-mode: compile

$hex = "FF";
$dec = base_convert($hex, 16, 10);
$bin = base_convert($dec, 10, 2);
echo "Dec=$dec Bin=$bin";
