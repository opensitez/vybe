<?php
// vybe-test: php/scope_patterns/static_vars_independent_per_function
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

function counterA(): int { static $n = 0; return ++$n; }
function counterB(): int { static $n = 0; return ++$n; }
counterA(); counterA();
counterB();
echo counterA();
echo counterB();
