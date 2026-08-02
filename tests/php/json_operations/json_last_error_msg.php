<?php
// vybe-test: php/json_operations/json_last_error_msg
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

json_decode('{bad json}');
$msg = json_last_error_msg();
echo is_string($msg) ? 'string' : 'not-string';
