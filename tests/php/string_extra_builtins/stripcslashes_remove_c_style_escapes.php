<?php
// vybe-test: php/string_extra_builtins/stripcslashes_remove_c_style_escapes
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$escaped = 'He said \\"hello\\"';
$s = stripcslashes($escaped);
echo is_string($s) ? "ok" : "fail";
