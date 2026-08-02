<?php
// vybe-test: php/closures_advanced/closure_type_hint
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

function apply(Closure $fn, int $v): int { return $fn($v); }
echo apply(fn($x) => $x * 3, 14);
