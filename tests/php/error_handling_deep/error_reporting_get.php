<?php
// vybe-test: php/error_handling_deep/error_reporting_get
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$current = error_reporting();
echo is_int($current) ? 'is int' : 'not int';
