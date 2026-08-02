<?php
// vybe-test: php/binary_data/bin2hex_random_bytes
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$bytes = random_bytes(16);
$hex = bin2hex($bytes);
echo strlen($hex) === 32 ? '32 hex chars' : 'wrong length';
echo ctype_xdigit($hex) ? ':valid hex' : ':invalid hex';
