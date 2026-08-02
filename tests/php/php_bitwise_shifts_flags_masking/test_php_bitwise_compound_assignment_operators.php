<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_compound_assignment_operators
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs
// vybe-test-mode: compile

$flags = 0;
$flags |= (1 << 0);
$flags |= (1 << 1);
$flags &= ~(1 << 0);
echo $flags === 2 ? "COMPOUND_OK" : "FAIL";
