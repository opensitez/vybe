<?php
// vybe-test: php/json_operations/json_encode_integer_keys
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$map = [0 => 'zero', 1 => 'one', 2 => 'two'];
echo json_encode($map);
