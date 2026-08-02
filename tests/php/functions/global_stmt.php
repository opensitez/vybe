<?php
// vybe-test: php/functions/global_stmt
// origin: languages/php/tests/php/test_functions.rs
// vybe-test-mode: compile

$g = 10; function foo() { global $g; echo $g; } foo();
