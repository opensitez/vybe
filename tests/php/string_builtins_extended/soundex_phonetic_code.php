<?php
// vybe-test: php/string_builtins_extended/soundex_phonetic_code
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$code = soundex("Smith");
echo is_string($code) ? "ok" : "fail";
echo strlen($code) === 4 ? "four" : "other";
echo soundex("Smythe") === soundex("Smith") ? "match" : "no";
