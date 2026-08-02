<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_shift_overflow_behavior
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs
// vybe-test-mode: compile

$val = 1 << 30;
echo is_int($val) ? "INT_SHIFT" : "FLOAT_SHIFT";
