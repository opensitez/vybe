<?php
// vybe-test: php/scoping/nested_function_scope
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

function outer() { $x = 1; function inner() { return 42; } return inner() + $x; } echo outer();
