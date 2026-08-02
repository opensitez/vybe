<?php
// vybe-test: php/declare/strict_types_return_type_enforcement
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function clamp(int $v, int $lo, int $hi): int {
    return max($lo, min($hi, $v));
}
echo clamp(15, 0, 10) . ',' . clamp(-5, 0, 10);
