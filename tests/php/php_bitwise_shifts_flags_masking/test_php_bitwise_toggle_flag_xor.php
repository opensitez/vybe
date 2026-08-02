<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_toggle_flag_xor
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs
// vybe-test-mode: compile

$flag = 0b0010;
$flag ^= 0b0010; // toggle off -> 0
echo $flag === 0 ? "TOGGLE_OFF" : "FAIL";
