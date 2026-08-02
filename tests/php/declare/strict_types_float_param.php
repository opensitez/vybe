<?php
// vybe-test: php/declare/strict_types_float_param
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function area(float $r): float { return M_PI * $r * $r; }
echo round(area(2.0), 4);
