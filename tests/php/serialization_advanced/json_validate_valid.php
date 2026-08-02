<?php
// vybe-test: php/serialization_advanced/json_validate_valid
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

echo json_validate('{"key": "value"}') ? 'valid' : 'invalid';
echo json_validate('[1, 2, 3]') ? 'valid' : 'invalid';
