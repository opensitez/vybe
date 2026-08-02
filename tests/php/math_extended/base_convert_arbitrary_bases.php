<?php
// vybe-test: php/math_extended/base_convert_arbitrary_bases
// origin: languages/php/tests/php/test_math_extended.rs
// vybe-test-mode: compile

$dec = base_convert('ff', 16, 10);
$hex = base_convert('255', 10, 16);
echo $dec;
echo $hex;
