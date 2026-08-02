<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_string_xor_encryption
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs
// vybe-test-mode: compile

$text = "Hello";
$key = "K";
$encrypted = $text ^ str_repeat($key, strlen($text));
$decrypted = $encrypted ^ str_repeat($key, strlen($text));
echo $decrypted === $text ? "XOR_ROUNDTRIP_OK" : "FAIL";
