<?php
// vybe-test: php/error_handling/throw_in_arrow
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

$fn = fn($x) => $x ?? throw new Exception('null');
