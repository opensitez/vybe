<?php
// vybe-test: php/error_handling/throw_in_coalesce
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

$x = $val ?? throw new Exception('missing');
