<?php
// vybe-test: php/string_builtins_extended/addcslashes_and_stripcslashes_roundtrip
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$escaped = addcslashes("a b", " a");
$original = stripcslashes($escaped);
echo strlen($escaped);
echo $original;
