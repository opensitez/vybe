<?php
// vybe-test: php/serialization_advanced/json_encode_throw_on_error
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$valid = ['key' => 'value'];
try {
    $json = json_encode($valid, JSON_THROW_ON_ERROR);
    echo 'ok';
} catch (\JsonException $e) {
    echo 'error: ' . $e->getMessage();
}
