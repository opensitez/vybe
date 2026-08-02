<?php
// vybe-test: php/json_operations/json_decode_invalid_returns_null
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$result = json_decode('not valid json', true);
var_dump($result);
