<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_not_complement
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs
// vybe-test-mode: compile

$x = 0;
$inv = ~$x;
echo is_int($inv) ? "INT_NOT_OK" : "FAIL";
