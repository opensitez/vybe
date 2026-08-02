<?php
// vybe-test: php/string_extra_builtins/addcslashes_c_style_escapes
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$s = "Hello\tWorld\n";
$escaped = addcslashes($s, "\t\n");
echo is_string($escaped) ? "ok" : "fail";
