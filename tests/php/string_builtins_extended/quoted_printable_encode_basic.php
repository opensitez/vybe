<?php
// vybe-test: php/string_builtins_extended/quoted_printable_encode_basic
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$encoded = quoted_printable_encode("Subject: =?UTF-8?");
echo is_string($encoded) ? "ok" : "fail";
$decoded = quoted_printable_decode($encoded);
echo is_string($decoded) ? "ok" : "fail";
