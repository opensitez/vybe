<?php
// vybe-test: php/scoping/global_keyword
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$x = 10; function foo() { global $x; return $x + 1; } echo foo();
