<?php
// vybe-test: php/declare/strict_types_variadic
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function sumInts(int ...$nums): int { return array_sum($nums); }
echo sumInts(1, 2, 3, 4, 5);
