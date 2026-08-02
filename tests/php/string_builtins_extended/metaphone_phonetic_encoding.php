<?php
// vybe-test: php/string_builtins_extended/metaphone_phonetic_encoding
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$m = metaphone("Thompson");
echo is_string($m) ? "ok" : "fail";
echo strlen($m) > 0 ? "nonempty" : "empty";
echo metaphone("Thomson") === metaphone("Tomson") ? "match" : "no";
