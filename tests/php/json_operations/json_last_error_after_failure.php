<?php
// vybe-test: php/json_operations/json_last_error_after_failure
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

json_decode('{bad json}');
$err = json_last_error();
echo $err !== JSON_ERROR_NONE ? 'error' : 'ok';
