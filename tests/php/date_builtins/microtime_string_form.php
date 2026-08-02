<?php
// vybe-test: php/date_builtins/microtime_string_form
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$mt = microtime();
echo is_string($mt) ? 'string' : 'not string';
