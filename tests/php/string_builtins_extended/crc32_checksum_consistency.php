<?php
// vybe-test: php/string_builtins_extended/crc32_checksum_consistency
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$a = crc32("consistent input");
$b = crc32("consistent input");
echo ($a === $b) ? "stable" : "unstable";
echo is_int($a) ? "integer" : "not integer";
