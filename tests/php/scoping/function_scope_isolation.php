<?php
// vybe-test: php/scoping/function_scope_isolation
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$x = 1; function foo() { $x = 2; return $x; } echo foo(); echo $x;
