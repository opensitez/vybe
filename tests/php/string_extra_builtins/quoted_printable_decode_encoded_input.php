<?php
// vybe-test: php/string_extra_builtins/quoted_printable_decode_encoded_input
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$encoded = "Subject: =?UTF-8?Q?Hello=20World?=";
$decoded = quoted_printable_decode($encoded);
echo is_string($decoded) ? "ok" : "fail";
