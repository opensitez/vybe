<?php
// vybe-test: php/string_builtins_extended/bin2hex_encode_bytes
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$hex = bin2hex("Hi");
echo strtolower($hex);
echo strlen($hex);
