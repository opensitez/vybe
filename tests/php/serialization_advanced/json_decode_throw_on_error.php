<?php
// vybe-test: php/serialization_advanced/json_decode_throw_on_error
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

try {
    $v = json_decode('invalid json', true, 512, JSON_THROW_ON_ERROR);
} catch (\JsonException $e) {
    echo 'caught: ' . $e->getMessage();
}
