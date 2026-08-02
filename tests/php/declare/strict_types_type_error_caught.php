<?php
// vybe-test: php/declare/strict_types_type_error_caught
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function strictInt(int $n): int { return $n * 2; }
try {
    $result = strictInt(3);
    echo $result;
} catch (TypeError $e) {
    echo 'type error: ' . $e->getMessage();
}
