<?php
// vybe-test: php/string_extra_builtins/preg_last_error_after_match
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

preg_match('/\d+/', 'abc123');
$err = preg_last_error();
echo is_int($err) ? "int" : "not-int";
echo $err === PREG_NO_ERROR ? "no-error" : "error";
