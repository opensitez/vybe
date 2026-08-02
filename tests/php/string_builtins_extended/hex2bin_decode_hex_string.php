<?php
// vybe-test: php/string_builtins_extended/hex2bin_decode_hex_string
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$binary = hex2bin("48656c6c6f");
echo $binary;
