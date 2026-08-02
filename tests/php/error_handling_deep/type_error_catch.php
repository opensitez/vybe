<?php
// vybe-test: php/error_handling_deep/type_error_catch
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

declare(strict_types=1);
function mustBeInt(int $n): int { return $n; }
try {
    $r = mustBeInt(3); // valid
    echo "ok: $r";
} catch (\TypeError $e) {
    echo 'type error: ' . $e->getMessage();
}
