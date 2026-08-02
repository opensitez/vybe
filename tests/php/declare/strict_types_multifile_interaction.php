<?php
// vybe-test: php/declare/strict_types_multifile_interaction
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function divide(float $a, float $b): float {
    if ($b == 0.0) throw new \DivisionByZeroError("Division by zero");
    return $a / $b;
}
echo divide(10.0, 4.0);
