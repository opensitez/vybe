<?php
// vybe-test: php/declare/strict_types_basic
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function add(int $a, int $b): int { return $a + $b; }
echo add(2, 3);
