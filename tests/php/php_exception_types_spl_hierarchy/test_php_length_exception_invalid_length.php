<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_length_exception_invalid_length
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs
// vybe-test-mode: compile

function validatePin(string $pin) {
    if (strlen($pin) !== 4) {
        throw new LengthException("PIN must be exactly 4 digits");
    }
}

try {
    validatePin("12");
} catch (LengthException $e) {
    echo $e->getMessage();
}
