<?php
// vybe-test: php/serialization_advanced/json_validate_invalid
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

echo json_validate('not json') ? 'valid' : 'invalid';
echo json_validate('{bad: "json"}') ? 'valid' : 'invalid';
