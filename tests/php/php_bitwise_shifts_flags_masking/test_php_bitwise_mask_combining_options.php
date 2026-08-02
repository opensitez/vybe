<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_mask_combining_options
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs
// vybe-test-mode: compile

$options = JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR;
echo is_int($options) ? "MASK_INT" : "FAIL";
