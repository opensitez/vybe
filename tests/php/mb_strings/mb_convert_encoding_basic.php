<?php
// vybe-test: php/mb_strings/mb_convert_encoding_basic
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$utf8 = "Hello World";
$converted = mb_convert_encoding($utf8, 'UTF-8', 'UTF-8');
echo $converted === $utf8 ? 'same' : 'different';
